const { test } = require('node:test');
const assert = require('node:assert/strict');
const express = require('express');
const jwt = require('jsonwebtoken');
const { once } = require('node:events');

test('el reporte de Google Sheets requiere admin y reemplaza un reporte sin datos personales', async (t) => {
  const previous = {
    JWT_SECRET: process.env.JWT_SECRET,
    GOOGLE_SHEETS_API: process.env.GOOGLE_SHEETS_API,
    GOOGLE_SHEETS_SPREADSHEET_ID: process.env.GOOGLE_SHEETS_SPREADSHEET_ID,
  };
  const testCredentials = JSON.stringify({ client_email: 'reports@example.test', private_key: 'test-key' });
  process.env.JWT_SECRET = 'google-sheets-test-secret';
  process.env.GOOGLE_SHEETS_API = testCredentials;
  process.env.GOOGLE_SHEETS_SPREADSHEET_ID = 'spreadsheet_test_id';

  const writes = [];
  const queries = [];
  const googleapisPath = require.resolve('googleapis');
  const databasePath = require.resolve('../src/config/database');
  const originalGoogleapis = require.cache[googleapisPath];
  const originalDatabase = require.cache[databasePath];
  require.cache[googleapisPath] = {
    id: googleapisPath,
    filename: googleapisPath,
    loaded: true,
    exports: {
      google: {
        auth: { GoogleAuth: class GoogleAuth { constructor(options) { this.options = options; } } },
        sheets: () => ({
          spreadsheets: {
            get: async () => ({ data: { sheets: [{ properties: { title: 'Reporte_20260301_20260307', sheetId: 42 } }] } }),
            batchUpdate: async (request) => { writes.push(request); return { data: { replies: [] } }; },
            values: { update: async (request) => { writes.push(request); return { data: {} }; } },
          },
        }),
      },
    },
  };
  require.cache[databasePath] = {
    id: databasePath,
    filename: databasePath,
    loaded: true,
    exports: {
      query: async (sql, params = []) => {
        queries.push(sql);
        if (sql.includes('SELECT id, email, nombre, rol, estado_cuenta FROM usuarios')) {
          const users = {
            1: { id: 1, email: 'admin@example.test', nombre: 'Admin', rol: 'admin', estado_cuenta: 'activo' },
            2: { id: 2, email: 'staff@example.test', nombre: 'Staff', rol: 'registrador', estado_cuenta: 'activo' },
            3: { id: 3, email: 'editor@example.test', nombre: 'Editor', rol: 'editor', estado_cuenta: 'activo' },
          };
          return { rows: users[params[0]] ? [users[params[0]]] : [] };
        }
        if (sql.includes('AS today')) return { rows: [{ today: '2026-03-07' }] };
        if (sql.includes('AS reservations')) {
          return { rows: [{ reservations: 1, confirmed: 1, cancelled: 0, travelers: 2, paid_revenue: '150.00' }] };
        }
        if (sql.includes('generate_series')) return { rows: [{ date: '2026-03-01', count: 1, paid: '150.00' }] };
        if (sql.includes('SELECT metodo_pago AS method')) return { rows: [{ method: 'tarjeta', count: 1 }] };
        if (sql.includes('SELECT COALESCE(d.nombre')) return { rows: [{ name: 'Paracas', count: 1, paid: '150.00' }] };
        if (sql.includes('AS count FROM reservas')) return { rows: [{ count: 1 }] };
        if (sql.includes('to_char(created_at')) {
          return { rows: [{
            id: 7, created_at: '2026-03-02 10:00:00', fecha_tour: '2026-03-10',
            promocion_nombre: 'Tour de prueba', asistentes: 2, metodo_pago: 'tarjeta',
            estado: 'confirmada', estado_pago: 'pagado', total: '150.00',
            nombre: 'Dato privado', email: 'privado@example.test',
          }] };
        }
        throw new Error(`Consulta no esperada en prueba: ${sql}`);
      },
    },
  };

  const app = express();
  app.use(express.json());
  app.use('/api/admin', require('../src/routes/admin.routes'));
  const server = app.listen(0, '127.0.0.1');
  await once(server, 'listening');
  t.after(async () => {
    await new Promise((resolve) => server.close(resolve));
    if (originalGoogleapis) require.cache[googleapisPath] = originalGoogleapis;
    else delete require.cache[googleapisPath];
    if (originalDatabase) require.cache[databasePath] = originalDatabase;
    else delete require.cache[databasePath];
    for (const [key, value] of Object.entries(previous)) {
      if (value === undefined) delete process.env[key];
      else process.env[key] = value;
    }
  });

  async function request(userId, body) {
    const response = await fetch(`http://127.0.0.1:${server.address().port}/api/admin/reports/google-sheets`, {
      method: 'POST',
      headers: {
        ...(userId ? { Authorization: `Bearer ${jwt.sign({ id: userId }, process.env.JWT_SECRET)}` } : {}),
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(body),
    });
    return { status: response.status, data: await response.json() };
  }

  await t.test('restringe el acceso, valida fechas y comprueba la configuración', async () => {
    assert.equal((await request(null, { from: '2026-03-01', to: '2026-03-07' })).status, 401);
    assert.equal((await request(2, { from: '2026-03-01', to: '2026-03-07' })).status, 403);
    assert.equal((await request(3, { from: '2026-03-01', to: '2026-03-07' })).status, 403);
    assert.equal((await request(1, { from: '2026-02-30', to: '2026-03-07' })).status, 400);
    assert.equal((await request(1, { from: '2025-12-01', to: '2026-03-07' })).status, 400);

    delete process.env.GOOGLE_SHEETS_API;
    assert.equal((await request(1, { from: '2026-03-01', to: '2026-03-07' })).status, 503);
    process.env.GOOGLE_SHEETS_API = testCredentials;
  });

  await t.test('escribe el rango como RAW y omite nombre y correo', async () => {
    const result = await request(1, { from: '2026-03-01', to: '2026-03-07' });
    assert.equal(result.status, 200);
    assert.equal(result.data.title, 'Reporte_20260301_20260307');
    assert.match(result.data.url, /#gid=42$/);
    assert.equal(queries.filter((sql) => sql.includes('AT TIME ZONE')).length >= 1, true);

    const update = writes.find((item) => item.range === "'Reporte_20260301_20260307'!A1");
    assert.ok(update);
    assert.equal(update.valueInputOption, 'RAW');
    const content = JSON.stringify(update.requestBody.values);
    assert.match(content, /Tour de prueba/);
    assert.doesNotMatch(content, /Dato privado|privado@example\.test/);
  });
});

const { google } = require('googleapis');
const { query } = require('../config/database');
const { CmsError, endpoint } = require('../utils/cms');

const MAX_RESERVATIONS = 10000;
const SHEETS_SCOPE = 'https://www.googleapis.com/auth/spreadsheets';

function isDate(value) {
  if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
  const parsed = new Date(`${value}T00:00:00.000Z`);
  return !Number.isNaN(parsed.getTime()) && parsed.toISOString().slice(0, 10) === value;
}

function createSheetsClient() {
  const rawCredentials = process.env.GOOGLE_SHEETS_API;
  const spreadsheetId = process.env.GOOGLE_SHEETS_SPREADSHEET_ID;
  if (!rawCredentials || typeof spreadsheetId !== 'string' || !/^[A-Za-z0-9_-]+$/.test(spreadsheetId)) {
    throw new CmsError('Faltan GOOGLE_SHEETS_API o GOOGLE_SHEETS_SPREADSHEET_ID en el servidor.', 503);
  }

  let credentials;
  try {
    credentials = JSON.parse(rawCredentials);
  } catch {
    throw new CmsError('GOOGLE_SHEETS_API debe contener el JSON de una cuenta de servicio de Google.', 503);
  }
  if (!credentials || typeof credentials !== 'object' || Array.isArray(credentials)
    || typeof credentials.client_email !== 'string' || typeof credentials.private_key !== 'string') {
    throw new CmsError('GOOGLE_SHEETS_API no contiene credenciales válidas de una cuenta de servicio.', 503);
  }

  const auth = new google.auth.GoogleAuth({ credentials, scopes: [SHEETS_SCOPE] });
  return { sheets: google.sheets({ version: 'v4', auth }), spreadsheetId };
}

function reportRows({ from, to, summary, daily, methods, destinations, reservations }) {
  return [
    ['Reporte administrativo de reservas', from, to],
    ['Zona horaria', 'America/Lima'],
    [],
    ['Resumen'],
    ['Indicador', 'Valor'],
    ['Reservas', summary.reservations],
    ['Confirmadas', summary.confirmed],
    ['Canceladas', summary.cancelled],
    ['Viajeros (excluye canceladas)', summary.travelers],
    ['Ingresos pagados (PEN)', Number(summary.paid_revenue)],
    [],
    ['Actividad diaria'],
    ['Fecha (Perú)', 'Reservas', 'Ingresos pagados (PEN)'],
    ...daily.map((row) => [row.date, row.count, Number(row.paid)]),
    [],
    ['Métodos de pago'],
    ['Método', 'Reservas'],
    ...methods.map((row) => [row.method, row.count]),
    [],
    ['Destinos'],
    ['Destino', 'Reservas', 'Ingresos pagados (PEN)'],
    ...destinations.map((row) => [row.name, row.count, Number(row.paid)]),
    [],
    ['Detalle de reservas (sin datos personales)'],
    ['ID', 'Creada (Perú)', 'Fecha del tour', 'Experiencia', 'Asistentes', 'Método de pago', 'Estado', 'Estado de pago', 'Total (PEN)'],
    ...reservations.map((row) => [
      row.id, row.created_at, row.fecha_tour, row.promocion_nombre, row.asistentes,
      row.metodo_pago, row.estado, row.estado_pago, Number(row.total),
    ]),
  ];
}

const exportGoogleSheetsReport = endpoint(async (req, res) => {
  const { from, to } = req.body || {};
  if (!isDate(from) || !isDate(to) || from > to) {
    throw new CmsError('Indica un rango de fechas válido en formato AAAA-MM-DD.');
  }
  const days = (Date.parse(`${to}T00:00:00.000Z`) - Date.parse(`${from}T00:00:00.000Z`)) / 86400000 + 1;
  if (days > 90) throw new CmsError('El reporte puede abarcar como máximo 90 días.');

  const today = (await query(`SELECT (NOW() AT TIME ZONE 'America/Lima')::date::text AS today`)).rows[0].today;
  if (to > today) throw new CmsError('La fecha final no puede ser posterior a hoy en Perú.');

  const { sheets, spreadsheetId } = createSheetsClient();
  const [summaryResult, dailyResult, methodsResult, destinationsResult, countResult] = await Promise.all([
    query(`SELECT COUNT(*)::int AS reservations,
        COUNT(*) FILTER (WHERE estado = 'confirmada')::int AS confirmed,
        COUNT(*) FILTER (WHERE estado = 'cancelada')::int AS cancelled,
        COALESCE(SUM(asistentes) FILTER (WHERE estado <> 'cancelada'), 0)::int AS travelers,
        COALESCE(SUM(total) FILTER (WHERE estado_pago = 'pagado'), 0)::numeric AS paid_revenue
      FROM reservas
      WHERE (created_at AT TIME ZONE 'America/Lima')::date BETWEEN $1::date AND $2::date`, [from, to]),
    query(`WITH dates AS (
        SELECT generate_series($1::date, $2::date, '1 day'::interval)::date AS day
      )
      SELECT dates.day::text AS date, COUNT(r.id)::int AS count,
        COALESCE(SUM(r.total) FILTER (WHERE r.estado_pago = 'pagado'), 0)::numeric AS paid
      FROM dates LEFT JOIN reservas r
        ON (r.created_at AT TIME ZONE 'America/Lima')::date = dates.day
      GROUP BY dates.day ORDER BY dates.day`, [from, to]),
    query(`SELECT metodo_pago AS method, COUNT(*)::int AS count FROM reservas
      WHERE (created_at AT TIME ZONE 'America/Lima')::date BETWEEN $1::date AND $2::date
      GROUP BY metodo_pago ORDER BY metodo_pago`, [from, to]),
    query(`SELECT COALESCE(d.nombre, 'Experiencias archivadas') AS name, COUNT(r.id)::int AS count,
        COALESCE(SUM(r.total) FILTER (WHERE r.estado_pago = 'pagado'), 0)::numeric AS paid
      FROM reservas r LEFT JOIN promociones p ON p.id = r.promocion_id
      LEFT JOIN destinos d ON d.id = p.destino_id
      WHERE (r.created_at AT TIME ZONE 'America/Lima')::date BETWEEN $1::date AND $2::date
        AND r.estado <> 'cancelada'
      GROUP BY d.id, d.nombre ORDER BY count DESC, name`, [from, to]),
    query(`SELECT COUNT(*)::int AS count FROM reservas
      WHERE (created_at AT TIME ZONE 'America/Lima')::date BETWEEN $1::date AND $2::date`, [from, to]),
  ]);

  if (countResult.rows[0].count > MAX_RESERVATIONS) {
    throw new CmsError(`El período contiene más de ${MAX_RESERVATIONS} reservas. Reduce el rango para exportarlo.`, 413);
  }

  const reservationResult = await query(`SELECT id,
      to_char(created_at AT TIME ZONE 'America/Lima', 'YYYY-MM-DD HH24:MI:SS') AS created_at,
      fecha_tour::text AS fecha_tour, COALESCE(promocion_nombre, 'Experiencia archivada') AS promocion_nombre,
      asistentes, metodo_pago, estado, estado_pago, total::numeric AS total
    FROM reservas
    WHERE (created_at AT TIME ZONE 'America/Lima')::date BETWEEN $1::date AND $2::date
    ORDER BY created_at, id`, [from, to]);

  const sheetTitle = `Reporte_${from.replace(/-/g, '')}_${to.replace(/-/g, '')}`;
  let sheetId;
  try {
    const workbook = await sheets.spreadsheets.get({
      spreadsheetId,
      fields: 'sheets.properties(sheetId,title)',
    });
    const existing = workbook.data.sheets?.find((sheet) => sheet.properties.title === sheetTitle);
    if (existing) {
      sheetId = existing.properties.sheetId;
    } else {
      try {
        const created = await sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          requestBody: { requests: [{ addSheet: { properties: { title: sheetTitle } } }] },
        });
        sheetId = created.data.replies[0].addSheet.properties.sheetId;
      } catch (error) {
        const latest = await sheets.spreadsheets.get({
          spreadsheetId,
          fields: 'sheets.properties(sheetId,title)',
        });
        const concurrentSheet = latest.data.sheets?.find((sheet) => sheet.properties.title === sheetTitle);
        if (!concurrentSheet) throw error;
        sheetId = concurrentSheet.properties.sheetId;
      }
    }

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: [{ updateCells: { range: { sheetId }, fields: 'userEnteredValue' } }] },
    });
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `'${sheetTitle}'!A1`,
      valueInputOption: 'RAW',
      requestBody: {
        majorDimension: 'ROWS',
        values: reportRows({
          from, to, summary: summaryResult.rows[0], daily: dailyResult.rows,
          methods: methodsResult.rows, destinations: destinationsResult.rows,
          reservations: reservationResult.rows,
        }),
      },
    });
  } catch (error) {
    console.error('exportGoogleSheetsReport:', error.message);
    throw new CmsError('No se pudo escribir en Google Sheets. Verifica el ID de la hoja y el acceso de la cuenta de servicio.', 502);
  }

  return res.json({
    ok: true,
    title: sheetTitle,
    url: `https://docs.google.com/spreadsheets/d/${encodeURIComponent(spreadsheetId)}/edit#gid=${sheetId}`,
  });
});

module.exports = { exportGoogleSheetsReport };

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
  const rows = [
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
  ];
  const layout = {
    summaryHeader: 4,
    dailyHeader: rows.length,
    dailyStart: rows.length + 1,
  };
  rows.push(...daily.map((row) => [row.date, row.count, Number(row.paid)]));
  rows.push([], ['Métodos de pago']);
  layout.methodsTitle = rows.length - 1;
  layout.methodsHeader = rows.length;
  rows.push(['Método', 'Reservas']);
  layout.methodsStart = rows.length;
  rows.push(...methods.map((row) => [row.method, row.count]));
  rows.push([], ['Destinos']);
  layout.destinationsTitle = rows.length - 1;
  layout.destinationsHeader = rows.length;
  rows.push(['Destino', 'Reservas', 'Ingresos pagados (PEN)']);
  layout.destinationsStart = rows.length;
  rows.push(...destinations.map((row) => [row.name, row.count, Number(row.paid)]));
  rows.push([]);

  // Keep the reservation table below the four charts, even for short periods.
  while (rows.length < 57) rows.push([]);
  layout.reservationsTitle = rows.length;
  rows.push(['Detalle de reservas (sin datos personales)']);
  layout.reservationsHeader = rows.length;
  rows.push(['ID', 'Creada (Perú)', 'Fecha del tour', 'Experiencia', 'Asistentes', 'Método de pago', 'Estado', 'Estado de pago', 'Total (PEN)']);
  layout.reservationsStart = rows.length;
  rows.push(...reservations.map((row) => [
    row.id, row.created_at, row.fecha_tour, row.promocion_nombre, row.asistentes,
    row.metodo_pago, row.estado, row.estado_pago, Number(row.total),
  ]));
  return { rows, layout };
}

function gridRange(sheetId, startRowIndex, endRowIndex, startColumnIndex, endColumnIndex) {
  return { sheetId, startRowIndex, endRowIndex, startColumnIndex, endColumnIndex };
}

function reportFormattingRequests(sheetId, layout, counts) {
  const tableStyles = [
    { header: layout.dailyHeader, start: layout.dailyStart, count: counts.daily, columns: 3 },
    { header: layout.methodsHeader, start: layout.methodsStart, count: counts.methods, columns: 2 },
    { header: layout.destinationsHeader, start: layout.destinationsStart, count: counts.destinations, columns: 3 },
    { header: layout.reservationsHeader, start: layout.reservationsStart, count: counts.reservations, columns: 9 },
  ];
  const requests = [
    { repeatCell: {
      range: gridRange(sheetId, 0, layout.reservationsStart + counts.reservations, 0, 9),
      cell: { userEnteredFormat: { textFormat: { fontFamily: 'Arial', fontSize: 10 }, verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(textFormat,verticalAlignment)',
    } },
    { repeatCell: {
      range: gridRange(sheetId, 0, 1, 0, 3),
      cell: { userEnteredFormat: { backgroundColor: { red: 0.04, green: 0.29, blue: 0.35 }, textFormat: { foregroundColor: { red: 1, green: 1, blue: 1 }, bold: true, fontSize: 12 }, verticalAlignment: 'MIDDLE' } },
      fields: 'userEnteredFormat(backgroundColor,textFormat,verticalAlignment)',
    } },
    { repeatCell: {
      range: gridRange(sheetId, 1, 2, 0, 2),
      cell: { userEnteredFormat: { textFormat: { foregroundColor: { red: 0.24, green: 0.32, blue: 0.36 }, fontSize: 9 } } },
      fields: 'userEnteredFormat.textFormat',
    } },
    { repeatCell: {
      range: gridRange(sheetId, 3, 4, 0, 3),
      cell: { userEnteredFormat: { textFormat: { foregroundColor: { red: 0.04, green: 0.29, blue: 0.35 }, bold: true, fontSize: 11 } } },
      fields: 'userEnteredFormat.textFormat',
    } },
    ...[layout.dailyHeader - 1, layout.methodsTitle, layout.destinationsTitle, layout.reservationsTitle].map((rowIndex) => ({
      repeatCell: {
        range: gridRange(sheetId, rowIndex, rowIndex + 1, 0, rowIndex === layout.reservationsTitle ? 9 : 3),
        cell: { userEnteredFormat: { textFormat: { foregroundColor: { red: 0.04, green: 0.29, blue: 0.35 }, bold: true, fontSize: 11 } } },
        fields: 'userEnteredFormat.textFormat',
      },
    })),
    { repeatCell: {
      range: gridRange(sheetId, 4, 5, 0, 2),
      cell: { userEnteredFormat: { backgroundColor: { red: 0.04, green: 0.29, blue: 0.35 }, textFormat: { foregroundColor: { red: 1, green: 1, blue: 1 }, bold: true } } },
      fields: 'userEnteredFormat(backgroundColor,textFormat)',
    } },
    { repeatCell: {
      range: gridRange(sheetId, 9, 10, 1, 2),
      cell: { userEnteredFormat: { numberFormat: { type: 'CURRENCY', pattern: '"S/" #,##0.00' }, textFormat: { bold: true } } },
      fields: 'userEnteredFormat(numberFormat,textFormat)',
    } },
    { repeatCell: {
      range: gridRange(sheetId, 0, layout.reservationsStart + counts.reservations, 3, 4),
      cell: { userEnteredFormat: { backgroundColor: { red: 1, green: 1, blue: 1 } } },
      fields: 'userEnteredFormat.backgroundColor',
    } },
    { updateDimensionProperties: {
      range: { sheetId, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 },
      properties: { pixelSize: 145 }, fields: 'pixelSize',
    } },
    { updateDimensionProperties: {
      range: { sheetId, dimension: 'COLUMNS', startIndex: 1, endIndex: 3 },
      properties: { pixelSize: 145 }, fields: 'pixelSize',
    } },
    { updateDimensionProperties: {
      range: { sheetId, dimension: 'COLUMNS', startIndex: 3, endIndex: 4 },
      properties: { pixelSize: 30 }, fields: 'pixelSize',
    } },
    { updateDimensionProperties: {
      range: { sheetId, dimension: 'COLUMNS', startIndex: 4, endIndex: 9 },
      properties: { pixelSize: 120 }, fields: 'pixelSize',
    } },
    { mergeCells: { range: gridRange(sheetId, 3, 4, 0, 3), mergeType: 'MERGE_ALL' } },
    { mergeCells: { range: gridRange(sheetId, layout.methodsTitle, layout.methodsTitle + 1, 0, 2), mergeType: 'MERGE_ALL' } },
    { mergeCells: { range: gridRange(sheetId, layout.destinationsTitle, layout.destinationsTitle + 1, 0, 3), mergeType: 'MERGE_ALL' } },
    { mergeCells: { range: gridRange(sheetId, layout.reservationsTitle, layout.reservationsTitle + 1, 0, 9), mergeType: 'MERGE_ALL' } },
  ];

  for (const { header, start, count, columns } of tableStyles) {
    requests.push({
      repeatCell: {
        range: gridRange(sheetId, header, header + 1, 0, columns),
        cell: { userEnteredFormat: { backgroundColor: { red: 0.04, green: 0.29, blue: 0.35 }, textFormat: { foregroundColor: { red: 1, green: 1, blue: 1 }, bold: true } } },
        fields: 'userEnteredFormat(backgroundColor,textFormat)',
      },
    });
    if (count) requests.push({
      addBanding: {
        bandedRange: {
          range: gridRange(sheetId, header, start + count, 0, columns),
          rowProperties: {
            headerColor: { red: 0.04, green: 0.29, blue: 0.35 },
            firstBandColor: { red: 1, green: 1, blue: 1 },
            secondBandColor: { red: 0.90, green: 0.94, blue: 0.95 },
          },
        },
      },
    });
  }

  const currencyFormat = { type: 'CURRENCY', pattern: '"S/" #,##0.00' };
  for (const range of [
    gridRange(sheetId, layout.dailyStart, layout.dailyStart + counts.daily, 2, 3),
    gridRange(sheetId, layout.destinationsStart, layout.destinationsStart + counts.destinations, 2, 3),
    gridRange(sheetId, layout.reservationsStart, layout.reservationsStart + counts.reservations, 8, 9),
  ]) {
    if (range.startRowIndex < range.endRowIndex) requests.push({
      repeatCell: {
        range,
        cell: { userEnteredFormat: { numberFormat: currencyFormat } },
        fields: 'userEnteredFormat.numberFormat',
      },
    });
  }

  const statusRange = gridRange(sheetId, layout.reservationsStart, layout.reservationsStart + counts.reservations, 6, 8);
  if (counts.reservations) {
    for (const [column, value, color] of [
      [6, 'confirmada', { red: 0.82, green: 0.92, blue: 0.83 }],
      [6, 'pendiente', { red: 1, green: 0.95, blue: 0.78 }],
      [6, 'cancelada', { red: 0.98, green: 0.84, blue: 0.82 }],
      [7, 'pagado', { red: 0.82, green: 0.92, blue: 0.83 }],
      [7, 'pendiente', { red: 1, green: 0.95, blue: 0.78 }],
      [7, 'reembolsado', { red: 0.84, green: 0.88, blue: 0.96 }],
    ]) {
      requests.push({
        addConditionalFormatRule: {
          index: 0,
          rule: {
            ranges: [gridRange(sheetId, statusRange.startRowIndex, statusRange.endRowIndex, column, column + 1)],
            booleanRule: {
              condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: value }] },
              format: { backgroundColor: color },
            },
          },
        },
      });
    }
  }

  return requests;
}

function reportCharts(sheetId, layout, counts) {
  const dataRange = (startRow, endRow, startColumn, endColumn) => ({
    sourceRange: { sources: [gridRange(sheetId, startRow, endRow, startColumn, endColumn)] },
  });
  const chartPosition = (rowIndex) => ({
    overlayPosition: {
      anchorCell: { sheetId, rowIndex, columnIndex: 4 },
      widthPixels: 600,
      heightPixels: 235,
    },
  });
  const charts = [
    {
      title: 'Reservas por día',
      position: chartPosition(1),
      spec: {
        backgroundColor: { red: 1, green: 1, blue: 1 },
        basicChart: {
          chartType: 'COLUMN', legendPosition: 'NO_LEGEND', headerCount: 1,
          axis: [{ position: 'BOTTOM_AXIS', title: 'Fecha' }, { position: 'LEFT_AXIS', title: 'Reservas' }],
          domains: [{ domain: dataRange(layout.dailyHeader, layout.dailyStart + counts.daily, 0, 1) }],
          series: [{ series: dataRange(layout.dailyHeader, layout.dailyStart + counts.daily, 1, 2), targetAxis: 'LEFT_AXIS', color: { red: 0.04, green: 0.29, blue: 0.35 } }],
        },
      },
    },
    {
      title: 'Ingresos pagados por día (PEN)',
      position: chartPosition(15),
      spec: {
        backgroundColor: { red: 1, green: 1, blue: 1 },
        basicChart: {
          chartType: 'COLUMN', legendPosition: 'NO_LEGEND', headerCount: 1,
          axis: [{ position: 'BOTTOM_AXIS', title: 'Fecha' }, { position: 'LEFT_AXIS', title: 'S/' }],
          domains: [{ domain: dataRange(layout.dailyHeader, layout.dailyStart + counts.daily, 0, 1) }],
          series: [{ series: dataRange(layout.dailyHeader, layout.dailyStart + counts.daily, 2, 3), targetAxis: 'LEFT_AXIS', color: { red: 0.92, green: 0.43, blue: 0.12 } }],
        },
      },
    },
    ...(counts.methods ? [{
      title: 'Reservas por método de pago',
      position: chartPosition(29),
      spec: {
        backgroundColor: { red: 1, green: 1, blue: 1 },
        pieChart: {
          legendPosition: 'RIGHT_LEGEND', pieHole: 0.58,
          domain: dataRange(layout.methodsHeader, layout.methodsStart + counts.methods, 0, 1),
          series: dataRange(layout.methodsHeader, layout.methodsStart + counts.methods, 1, 2),
        },
      },
    }] : []),
    ...(counts.destinations ? [{
      title: 'Ingresos pagados por destino (PEN)',
      position: chartPosition(43),
      spec: {
        backgroundColor: { red: 1, green: 1, blue: 1 },
        basicChart: {
          chartType: 'BAR', legendPosition: 'NO_LEGEND', headerCount: 1,
          axis: [{ position: 'BOTTOM_AXIS', title: 'S/' }, { position: 'LEFT_AXIS', title: 'Destino' }],
          domains: [{ domain: dataRange(layout.destinationsHeader, layout.destinationsStart + counts.destinations, 0, 1) }],
          series: [{ series: dataRange(layout.destinationsHeader, layout.destinationsStart + counts.destinations, 2, 3), color: { red: 0.15, green: 0.61, blue: 0.54 } }],
        },
      },
    }] : []),
  ];
  return charts.map((chart) => ({ addChart: { chart } }));
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
      fields: 'sheets(properties(sheetId,title),charts(chartId),bandedRanges(bandedRangeId),conditionalFormats,merges)',
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

    const selectedSheet = workbook.data.sheets?.find((sheet) => sheet.properties.sheetId === sheetId);
    const clearRequests = [
      ...((selectedSheet?.charts || []).map((chart) => ({ deleteEmbeddedObject: { objectId: chart.chartId } }))),
      ...((selectedSheet?.bandedRanges || []).map((banding) => ({ deleteBanding: { bandedRangeId: banding.bandedRangeId } }))),
      ...((selectedSheet?.conditionalFormats || []).map((_, index, rules) => ({
        deleteConditionalFormatRule: { sheetId, index: rules.length - index - 1 },
      }))),
      ...((selectedSheet?.merges || []).map((range) => ({ unmergeCells: { range } }))),
      { updateCells: { range: { sheetId }, fields: 'userEnteredValue,userEnteredFormat' } },
    ];
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: clearRequests } });
    const { rows, layout } = reportRows({
      from, to, summary: summaryResult.rows[0], daily: dailyResult.rows,
      methods: methodsResult.rows, destinations: destinationsResult.rows,
      reservations: reservationResult.rows,
    });
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `'${sheetTitle}'!A1`,
      valueInputOption: 'RAW',
      requestBody: {
        majorDimension: 'ROWS',
        values: rows,
      },
    });
    const counts = {
      daily: dailyResult.rows.length, methods: methodsResult.rows.length,
      destinations: destinationsResult.rows.length, reservations: reservationResult.rows.length,
    };
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: reportFormattingRequests(sheetId, layout, counts) },
    });
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: reportCharts(sheetId, layout, counts) },
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

const DEFAULT_REPORT_RECIPIENT = 'ale21rock@gmail.com';
// If this script is bound to the target Google Sheet, leave SPREADSHEET_ID empty.
// If it is a standalone Apps Script project, paste the Google Sheet ID here.
const SPREADSHEET_ID = '';
const COBROS_SHEET_NAME = 'Cobros';
const APP_STATE_SHEET_NAME = 'AppState';
const APP_STATE_CHUNKS_SHEET_NAME = 'AppStateChunks';
const GASTOS_SHEET_NAME = 'Gastos';
const DEFAULT_STATE_KEY = 'gastos_planilla_oficial_v2';
const STATE_CHUNK_SIZE = 45000;
const DRIVE_BACKUP_FOLDER_ID = '1Bz9SO0_yQqKUA6F3XnSjxfjy74A9RmlP';
const DRIVE_BACKUP_CURRENT_FILE_NAME = 'gastos_backup_actual.json';
const DRIVE_BACKUP_DAILY_PREFIX = 'gastos_backup_';

function doGet(e) {
  try {
    const params = e && e.parameter ? e.parameter : {};
    const action = String(params.action || '').toLowerCase();

    if (action === 'get') {
      return jsonResponse_(getCobros_(params.year, params.month));
    }
    if (action === 'set') {
      return jsonResponse_(setCobro_(params.year, params.month, params.nombre, params.monto));
    }
    if (action === 'delete') {
      return jsonResponse_(deleteCobro_(params.year, params.month, params.nombre));
    }
    if (action === 'getstate') {
      return jsonResponse_(getAppState_(params.key || DEFAULT_STATE_KEY));
    }

    return jsonResponse_({ ok: false, error: 'Unsupported action.' });
  } catch (error) {
    return jsonResponse_({ ok: false, error: error.message || String(error) });
  }
}

function doPost(e) {
  try {
    const payload = parseRequestPayload_(e);
    const action = String(payload.action || '').toLowerCase();

    if (action === 'sendreport') {
      return jsonResponse_(sendCobrosReport_(payload));
    }
    if (action === 'savestate') {
      return jsonResponse_(saveAppState_(payload));
    }

    return jsonResponse_({ ok: false, error: 'Unsupported action.' });
  } catch (error) {
    return jsonResponse_({ ok: false, error: error.message || String(error) });
  }
}

function parseRequestPayload_(e) {
  if (e && e.postData && e.postData.contents) {
    return JSON.parse(e.postData.contents);
  }
  if (e && e.parameter && e.parameter.payload) {
    return JSON.parse(e.parameter.payload);
  }
  return e && e.parameter ? e.parameter : {};
}

function getSpreadsheet_() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('No spreadsheet available. Set SPREADSHEET_ID or bind this script to a Google Sheet.');
  return spreadsheet;
}

function getSheet_(name, headers) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  if (headers && headers.length) {
    const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    const mustWrite = headers.some((header, index) => current[index] !== header);
    if (mustWrite) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function getCobrosSheet_() {
  return getSheet_(COBROS_SHEET_NAME, ['year', 'month', 'nombre', 'monto', 'updatedAt']);
}

function getCobros_(year, month) {
  const sheet = getCobrosSheet_();
  const values = sheet.getDataRange().getValues();
  const result = {};
  const yearText = String(year);
  const monthText = String(month);
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (String(row[0]) === yearText && String(row[1]) === monthText && row[2]) {
      result[String(row[2])] = Number(row[3] || 0);
    }
  }
  return result;
}

function setCobro_(year, month, nombre, monto) {
  if (!nombre) throw new Error('Missing nombre.');
  const sheet = getCobrosSheet_();
  const values = sheet.getDataRange().getValues();
  const yearText = String(year);
  const monthText = String(month);
  const nombreText = String(nombre);
  const rowValue = [yearText, monthText, nombreText, Number(monto || 0), new Date().toISOString()];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (String(row[0]) === yearText && String(row[1]) === monthText && String(row[2]) === nombreText) {
      sheet.getRange(i + 1, 1, 1, rowValue.length).setValues([rowValue]);
      return { ok: true };
    }
  }
  sheet.appendRow(rowValue);
  return { ok: true };
}

function deleteCobro_(year, month, nombre) {
  const sheet = getCobrosSheet_();
  const values = sheet.getDataRange().getValues();
  const yearText = String(year);
  const monthText = String(month);
  const nombreText = String(nombre);
  for (let i = values.length - 1; i >= 1; i--) {
    const row = values[i];
    if (String(row[0]) === yearText && String(row[1]) === monthText && String(row[2]) === nombreText) {
      sheet.deleteRow(i + 1);
    }
  }
  return { ok: true };
}

function saveAppState_(payload) {
  const key = String(payload.key || DEFAULT_STATE_KEY);
  const updatedAt = String(payload.updatedAt || new Date().toISOString());
  const statePayload = {
    app: payload.app || 'gastos_planilla_oficial_v2',
    key,
    updatedAt,
    currentYear: payload.currentYear || '',
    state: payload.state || {},
    bcuMonthlyRates: payload.bcuMonthlyRates || {},
    tarjetasState: payload.tarjetasState || {},
    mesesCerrados: payload.mesesCerrados || {},
    cobrosPagos: payload.cobrosPagos || {}
  };
  const json = JSON.stringify(statePayload);
  const chunks = [];
  for (let index = 0; index < json.length; index += STATE_CHUNK_SIZE) {
    chunks.push(json.slice(index, index + STATE_CHUNK_SIZE));
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const metaSheet = getSheet_(APP_STATE_SHEET_NAME, ['key', 'updatedAt', 'chunkCount', 'bytes']);
    const chunksSheet = getSheet_(APP_STATE_CHUNKS_SHEET_NAME, ['key', 'chunkIndex', 'payload']);

    deleteRowsByKey_(chunksSheet, key, 1);
    upsertRowByKey_(metaSheet, key, [key, updatedAt, chunks.length, json.length]);
    if (chunks.length) {
      chunksSheet.getRange(chunksSheet.getLastRow() + 1, 1, chunks.length, 3)
        .setValues(chunks.map((chunk, index) => [key, index, chunk]));
    }
    syncGastosSheet_(statePayload.state, updatedAt);
  } finally {
    lock.releaseLock();
  }

  const driveBackup = saveDriveBackupSafely_(json, updatedAt);
  return { ok: true, key, updatedAt, chunks: chunks.length, driveBackup };
}

function saveDriveBackupSafely_(json, updatedAt) {
  if (!DRIVE_BACKUP_FOLDER_ID) return { ok: false, skipped: true, error: 'Missing Drive backup folder ID.' };
  try {
    const folder = DriveApp.getFolderById(DRIVE_BACKUP_FOLDER_ID);
    const current = upsertDriveTextFile_(folder, DRIVE_BACKUP_CURRENT_FILE_NAME, json, updatedAt);
    const dailyName = `${DRIVE_BACKUP_DAILY_PREFIX}${driveBackupDateKey_(updatedAt)}.json`;
    const daily = upsertDriveTextFile_(folder, dailyName, json, updatedAt);
    return {
      ok: true,
      folderId: DRIVE_BACKUP_FOLDER_ID,
      currentFileId: current.id,
      currentFileName: current.name,
      currentFileUrl: current.url,
      dailyFileId: daily.id,
      dailyFileName: daily.name,
      dailyFileUrl: daily.url
    };
  } catch (error) {
    return { ok: false, folderId: DRIVE_BACKUP_FOLDER_ID, error: error.message || String(error) };
  }
}

function upsertDriveTextFile_(folder, name, content, updatedAt) {
  const files = folder.getFilesByName(name);
  const description = `Backup generado por Control de Gastos. Actualizado: ${updatedAt || new Date().toISOString()}`;
  if (files.hasNext()) {
    const file = files.next();
    file.setContent(content);
    file.setDescription(description);
    return { id: file.getId(), name: file.getName(), url: file.getUrl() };
  }
  const file = folder.createFile(name, content, 'application/json');
  file.setDescription(description);
  return { id: file.getId(), name: file.getName(), url: file.getUrl() };
}

function driveBackupDateKey_(updatedAt) {
  const date = new Date(updatedAt || new Date().toISOString());
  const safeDate = Number.isNaN(date.getTime()) ? new Date() : date;
  const timezone = Session.getScriptTimeZone() || 'Etc/GMT';
  return Utilities.formatDate(safeDate, timezone, 'yyyy-MM-dd');
}

function probarBackupDrive() {
  const updatedAt = new Date().toISOString();
  const payload = {
    app: 'gastos_planilla_oficial_v2',
    test: true,
    updatedAt,
    message: 'Backup de prueba generado desde Apps Script.'
  };
  const folder = DriveApp.getFolderById(DRIVE_BACKUP_FOLDER_ID);
  return upsertDriveTextFile_(folder, 'gastos_backup_prueba.json', JSON.stringify(payload, null, 2), updatedAt);
}

function getAppState_(key) {
  const metaSheet = getSheet_(APP_STATE_SHEET_NAME, ['key', 'updatedAt', 'chunkCount', 'bytes']);
  const chunksSheet = getSheet_(APP_STATE_CHUNKS_SHEET_NAME, ['key', 'chunkIndex', 'payload']);
  const keyText = String(key || DEFAULT_STATE_KEY);
  const metaValues = metaSheet.getDataRange().getValues();
  let updatedAt = '';
  let chunkCount = 0;
  for (let i = 1; i < metaValues.length; i++) {
    if (String(metaValues[i][0]) === keyText) {
      updatedAt = String(metaValues[i][1] || '');
      chunkCount = Number(metaValues[i][2] || 0);
      break;
    }
  }
  if (!updatedAt || !chunkCount) return { ok: true, key: keyText, updatedAt: '', payload: null };

  const chunkValues = chunksSheet.getDataRange().getValues()
    .slice(1)
    .filter(row => String(row[0]) === keyText)
    .sort((a, b) => Number(a[1]) - Number(b[1]));
  const json = chunkValues.map(row => String(row[2] || '')).join('');
  return { ok: true, key: keyText, updatedAt, payload: JSON.parse(json) };
}

function syncGastosSheet_(state, updatedAt) {
  const headers = [
    'Año', 'Mes', 'Mes índice', 'UTE', 'OSE', 'ANTEL', 'Netflix', 'Spotify',
    'Total', 'Base UTE', 'Divisor UTE', 'Divisor OSE', 'Divisor ANTEL',
    'Divisor Netflix', 'Divisor Spotify', 'Actualizado'
  ];
  const sheet = getSheet_(GASTOS_SHEET_NAME, headers);
  const rows = [];
  Object.keys(state || {}).sort((a, b) => Number(a) - Number(b)).forEach(year => {
    const yearData = state[year] || {};
    const divisors = yearData.divisors || {};
    (yearData.months || []).forEach((month, index) => {
      rows.push([
        year,
        month.mes || '',
        index,
        Number(month.ute || 0),
        Number(month.ose || 0),
        Number(month.antel || 0),
        Number(month.netflix || 0),
        Number(month.spotify || 0),
        Number(month.total || 0),
        Number(yearData.baseUte || 0),
        Number(divisors.ute || 0),
        Number(divisors.ose || 0),
        Number(divisors.antel || 0),
        Number(divisors.netflix || 0),
        Number(divisors.spotify || 0),
        updatedAt
      ]);
    });
  });
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function deleteRowsByKey_(sheet, key, keyColumn) {
  const values = sheet.getDataRange().getValues();
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][keyColumn - 1]) === key) sheet.deleteRow(i + 1);
  }
}

function upsertRowByKey_(sheet, key, rowValues) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === key) {
      sheet.getRange(i + 1, 1, 1, rowValues.length).setValues([rowValues]);
      return;
    }
  }
  sheet.appendRow(rowValues);
}

function sendCobrosReport_(payload) {
  const report = payload.report || {};
  const rows = Array.isArray(report.rows) ? report.rows : [];
  const recipient = String(payload.recipient || DEFAULT_REPORT_RECIPIENT).trim();
  const subject = String(payload.subject || `Cobros ${report.monthName || ''} ${report.year || ''}`).trim();

  if (!recipient) throw new Error('Missing recipient email.');

  const html = buildCobrosReportHtml_(report, rows, payload.generatedAt);
  const pdfBlob = HtmlService.createHtmlOutput(html)
    .getBlob()
    .getAs(MimeType.PDF)
    .setName(`${subject.replace(/[^a-z0-9]+/gi, '_').replace(/_+/g, '_').replace(/^_|_$/g, '') || 'cobros_report'}.pdf`);

  GmailApp.sendEmail(
    recipient,
    subject,
    buildCobrosPlainText_(report, rows),
    {
      attachments: [pdfBlob],
      name: 'Control de Gastos'
    }
  );

  return { ok: true, recipient, subject };
}

function buildCobrosPlainText_(report, rows) {
  const lines = [
    `Cobros ${report.monthName || ''} ${report.year || ''}`.trim(),
    `Total a cobrar: ${formatCurrency_(report.totalToCollect)}`,
    `Total pagado: ${formatCurrency_(report.totalPaid)}`,
    `Total descontado: ${formatCurrency_(report.totalDiscounted)}`,
    ''
  ];
  rows.forEach(row => {
    lines.push(`${row.nombre}: total ${formatCurrency_(row.total)} | pagado ${formatCurrency_(row.pagado)} | desc. ${formatCurrency_(row.descuento)} | restante ${formatCurrency_(row.restante)}`);
  });
  return lines.join('\n');
}

function buildCobrosReportHtml_(report, rows, generatedAt) {
  const totalToCollect = formatCurrency_(report.totalToCollect);
  const totalPaid = formatCurrency_(report.totalPaid);
  const totalDiscounted = formatCurrency_(report.totalDiscounted);
  const tableRows = rows.map(row => `
    <tr>
      <td>${escapeHtml_(row.nombre)}</td>
      <td>${formatCurrency_(row.total)}</td>
      <td>${formatCurrency_(row.pagado)}</td>
      <td>${formatCurrency_(row.descuento)}</td>
      <td>${formatCurrency_(row.restante)}</td>
    </tr>
  `).join('');

  return `
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8" />
  <style>
    body{font-family:Arial,Helvetica,sans-serif;margin:0;padding:28px;color:#111827;background:#fff}
    h1{margin:0 0 6px;font-size:28px}
    .sub{color:#4b5563;margin:0 0 16px;font-size:14px}
    .meta{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px;margin:18px 0 22px}
    .card{border:1px solid #e5e7eb;border-radius:10px;padding:14px 16px;background:#f9fafb}
    .card span{display:block;color:#6b7280;font-size:11px;text-transform:uppercase;letter-spacing:.06em;font-weight:700;margin-bottom:6px}
    .card strong{font-size:22px}
    table{width:100%;border-collapse:collapse;margin-top:8px}
    th,td{padding:10px 8px;border-bottom:1px solid #e5e7eb;text-align:left;font-size:13px}
    th{background:#f3f4f6;color:#374151;text-transform:uppercase;letter-spacing:.05em;font-size:11px}
    td:last-child,th:last-child{text-align:right}
    .note{margin-top:14px;color:#6b7280;font-size:12px}
    @media print{body{padding:16px}.note{position:fixed;bottom:12px;left:16px;right:16px}}
  </style>
</head>
<body>
  <h1>Cobros a cobrar - ${escapeHtml_(String(report.monthName || ''))} ${escapeHtml_(String(report.year || ''))}</h1>
  <p class="sub">Solo muestra lo que paga cada persona, lo pagado, los descuentos y lo pendiente.</p>
  <div class="meta">
    <div class="card"><span>Total a cobrar</span><strong>${totalToCollect}</strong></div>
    <div class="card"><span>Total pagado</span><strong>${totalPaid}</strong></div>
    <div class="card"><span>Total descontado</span><strong>${totalDiscounted}</strong></div>
  </div>
  <table>
    <thead>
      <tr><th>Persona</th><th>Total</th><th>Pagado</th><th>Descuento</th><th>Restante</th></tr>
    </thead>
    <tbody>${tableRows}</tbody>
  </table>
  <p class="note">Generado ${escapeHtml_(String(generatedAt || new Date().toISOString()))} desde Control de Gastos.</p>
</body>
</html>`;
}

function escapeHtml_(value) {
  return String(value).replace(/[&<>"]/g, char => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;'
  }[char]));
}

function formatCurrency_(value) {
  const number = Number(value || 0);
  return `$${number.toLocaleString('es-UY', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

function jsonResponse_(payload) {
  const output = ContentService.createTextOutput(JSON.stringify(payload));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

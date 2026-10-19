/**
 * Google Apps Script backend for "Administrasi Persuratan SMPN 13 Tasikmalaya".
 * Serves the dashboard at /exec, JSON API at /exec?api=1, and accepts table
 * reads/writes through ?api=1&table=Name and doPost JSON bodies.
 * Reads/writes five tables: Tasks, Incoming, Outgoing, Masters, Settings.
 *
 * Target existing Apps Script project: 1o7jlw-sVmKfBLcgWXdS1hfMyTrL4zwJNtoq-IXuCMQ5-CeMF_5tKBO7Z
 * Target spreadsheet: 1C5Cp6ikblC_8GrJeLbRjYWtUEtX7i6InXRygjLO8_c8
 *
 * Setup in the EXISTING project (do not create a new project):
 * 1. Open the project by its Apps Script ID above and replace Code.gs with this file.
 * 2. Settings > Script properties:
 *      apiKey   = arbitrary shared secret (client sends it in the request body)
 * 3. Save, then update the existing Web app deployment to a new version in
 *    Deploy > Manage deployments. Do not create another deployment unless needed.
 * 4. Copy the Web app URL shown for that deployment into the frontend settings.
 * 5. Authorize Spreadsheet access when prompted; the script opens the spreadsheet
 *    ID above. Ensure the executing account can access that spreadsheet.
 */

/** Table definitions (1:1 with sheets). */
const CONFIG = {
  spreadsheetId: '1C5Cp6ikblC_8GrJeLbRjYWtUEtX7i6InXRygjLO8_c8',
  tables: [
    {
      name: 'Tasks',
      title: 'Surat Tugas',
      fields: [
        { key: 'id', label: 'ID', type: 'string', required: true },
        { key: 'number', label: 'Nomor Surat', type: 'string', required: true },
        { key: 'basis', label: 'Dasar Surat', type: 'string' },
        { key: 'basisNumber', label: 'Nomor Dasar', type: 'string' },
        { key: 'basisDate', label: 'Tanggal Dasar', type: 'date' },
        { key: 'recipient', label: 'Kepada', type: 'string' },
        { key: 'subject', label: 'Perihal', type: 'string' },
        { key: 'eventDate', label: 'Tanggal Kegiatan', type: 'date' },
        { key: 'eventName', label: 'Nama Kegiatan', type: 'string' },
        { key: 'eventDay', label: 'Hari', type: 'string' },
        { key: 'eventStartDate', label: 'Mulai', type: 'date' },
        { key: 'eventEndDate', label: 'Selesai', type: 'date' },
        { key: 'eventTime', label: 'Waktu', type: 'string' },
        { key: 'place', label: 'Tempat', type: 'string' },
        { key: 'origin', label: 'Asal', type: 'string' },
        { key: 'destination', label: 'Tujuan', type: 'string' },
        { key: 'travelPurpose', label: 'Maksud', type: 'string' },
        { key: 'departureDate', label: 'Keberangkatan', type: 'date' },
        { key: 'returnDate', label: 'Pulang', type: 'date' },
        { key: 'transportMode', label: 'Alat Angkutan', type: 'string' },
        { key: 'transportDays', label: 'Lama Perjalanan (hari)', type: 'number' },
        { key: 'budgetSource', label: 'Sumber Anggaran', type: 'string' },
        { key: 'itinerary', label: 'Itinerary', type: 'text' },
        { key: 'reportText', label: 'Laporan', type: 'text' },
        { key: 'reportTo', label: 'Kepada Laporan', type: 'string' },
        { key: 'recommendations', label: 'Rekomendasi', type: 'text' },
        { key: 'attachment', label: 'Lampiran', type: 'string' },
        { key: 'issuedDate', label: 'Tanggal Keluarkan', type: 'date' },
        { key: 'year', label: 'Tahun', type: 'number' },
        { key: 'createdAt', label: 'Dibuat', type: 'datetime' },
        { key: 'updatedAt', label: 'Diperbarui', type: 'datetime' },
        { key: 'assignees', label: 'Petugas (JSON)', type: 'text' },
        { key: 'students', label: 'Siswa (JSON)', type: 'text' },
        { key: 'sppdPerson', label: 'Petugas SPPD', type: 'string' },
      ],
    },
    {
      name: 'Incoming',
      title: 'Surat Masuk',
      fields: [
        { key: 'id', label: 'ID', type: 'string', required: true },
        { key: 'receivedDate', label: 'Tanggal Diterima', type: 'date', required: true },
        { key: 'letterDate', label: 'Tanggal Surat', type: 'date' },
        { key: 'letterNumber', label: 'Nomor Surat', type: 'string' },
        { key: 'sender', label: 'Pengirim', type: 'string', required: true },
        { key: 'subject', label: 'Perihal', type: 'string', required: true },
        { key: 'disposition', label: 'Disposisi', type: 'text' },
        { key: 'attachment', label: 'Lampiran', type: 'string' },
        { key: 'createdAt', label: 'Dibuat', type: 'datetime' },
      ],
    },
    {
      name: 'Outgoing',
      title: 'Surat Keluar',
      fields: [
        { key: 'id', label: 'ID', type: 'string', required: true },
        { key: 'issuedDate', label: 'Tanggal Keluarkan', type: 'date', required: true },
        { key: 'number', label: 'Nomor Surat', type: 'string', required: true },
        { key: 'letterType', label: 'Jenis Surat', type: 'string' },
        { key: 'recipient', label: 'Tujuan / Penerima', type: 'string', required: true },
        { key: 'subject', label: 'Perihal', type: 'string', required: true },
        { key: 'notes', label: 'Keterangan', type: 'text' },
        { key: 'attachment', label: 'Lampiran', type: 'string' },
        { key: 'createdAt', label: 'Dibuat', type: 'datetime' },
        { key: 'updatedAt', label: 'Diperbarui', type: 'datetime' },
      ],
    },
    {
      name: 'Masters',
      title: 'Master Guru & Siswa',
      fields: [
        { key: 'id', label: 'ID', type: 'string', required: true },
        { key: 'kind', label: 'Tipe', type: 'string', required: true },
        { key: 'name', label: 'Nama', type: 'string', required: true },
        { key: 'grade', label: 'Golongan', type: 'string' },
        { key: 'rank', label: 'Pangkat/Gol', type: 'string' },
        { key: 'position', label: 'Jabatan', type: 'string' },
        { key: 'nip', label: 'NIP', type: 'string' },
        { key: 'nuptk', label: 'NUPTK', type: 'string' },
        { key: 'nis', label: 'NIS', type: 'string' },
        { key: 'class', label: 'Kelas', type: 'string' },
        { key: 'rate', label: 'Tarif', type: 'number' },
        { key: 'createdAt', label: 'Dibuat', type: 'datetime' },
      ],
    },
    {
      name: 'Settings',
      title: 'Pengaturan aplikasi',
      fields: [
        { key: 'id', label: 'ID', type: 'string', required: true },
        { key: 'officials', label: 'Pejabat aktif (JSON)', type: 'text' },
        { key: 'officialHistory', label: 'Riwayat pejabat (JSON)', type: 'text' },
        { key: 'officialsMaster', label: 'Master petugas (JSON)', type: 'text' },
        { key: 'updatedAt', label: 'Diperbarui', type: 'datetime' },
      ],
    },
  ],
};

/** Add a header without replacing existing cells or data rows. */
function ensureTableHeaders(sheet, table) {
  const requiredHeaders = table.fields.map((field) => field.key);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow === 0 && lastColumn === 0) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    return { action: 'header dibuat', added: requiredHeaders.length };
  }

  const existingHeaders = lastColumn
    ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map((value) => String(value || '').trim())
    : [];

  // The client schema always starts with id. If row 1 is data or an unrelated
  // header, insert a new row so existing content moves intact below the header.
  if (existingHeaders[0] !== 'id') {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    return { action: 'header disisipkan; data lama dipertahankan', added: requiredHeaders.length };
  }

  const missingHeaders = requiredHeaders.filter((header) => !existingHeaders.includes(header));
  if (missingHeaders.length) {
    sheet.getRange(1, lastColumn + 1, 1, missingHeaders.length).setValues([missingHeaders]);
    return { action: 'kolom header dilengkapi', added: missingHeaders.length };
  }
  return { action: 'header sudah lengkap', added: 0 };
}

/**
 * Prepare the application tables in the configured spreadsheet.
 * Run manually from the Apps Script editor; this never clears or rewrites rows.
 */
function siapkanDatabase() {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  const report = CONFIG.tables.map((table) => {
    let sheet = spreadsheet.getSheetByName(table.name);
    const created = !sheet;
    if (!sheet) sheet = spreadsheet.insertSheet(table.name);
    const result = ensureTableHeaders(sheet, table);
    const line = `${table.name}: ${created ? 'sheet dibuat' : 'sheet diperiksa'}; ${result.action}; ${table.fields.length} kolom: ${table.fields.map((field) => field.key).join(', ')}`;
    Logger.log(line);
    return line;
  });
  Logger.log(`Inisialisasi database selesai untuk Spreadsheet ${CONFIG.spreadsheetId}.`);
  return report;
}

/** Read a sheet into an array of objects. The first column is always "id". */
function sheetToRows(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return [];
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const header = values[0] || [];
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r] || [];
    const obj = {};
    header.forEach((h, i) => {
      const raw = row[i];
      const field = tableFieldByKey_(h);
      if (raw instanceof Date && field?.type === 'date') {
        obj[h] = Utilities.formatDate(raw, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else if (raw instanceof Date && field?.type === 'datetime') {
        obj[h] = raw.toISOString();
      } else {
        obj[h] = raw === undefined ? '' : String(raw);
      }
    });
    if (Object.keys(obj).length === 0) continue;
    rows.push(obj);
  }
  return rows;
}

function tableFieldByKey_(key) {
  for (const table of CONFIG.tables) {
    const field = table.fields.find((item) => item.key === key);
    if (field) return field;
  }
  return null;
}

/** Write rows to a sheet. Expects an array of plain objects (id + fields). */
function rowsToSheet(sheet, rows) {
  if (!Array.isArray(rows)) rows = [];
  if (rows.length === 0) {
    // Clear only data rows, keep the header.
    const lastRow = sheet.getLastRow();
    sheet.getRange(2, 1, Math.max(1, lastRow - 1), sheet.getLastColumn()).clearContent();
    return;
  }
  // Build a wide 2D array: [headerRow, ...dataRows]
  const header = [];
  const fields = [];
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].forEach((h, i) => {
    header.push(h);
    if (h !== 'id') fields.push(h);
  });
  const data = rows.map((row) => {
    const arr = [];
    header.forEach((h) => {
      arr.push(row[h] === undefined || row[h] === null ? '' : row[h]);
    });
    return arr;
  });
  const maxCols = Math.max(header.length, data[0].length);
  sheet.getRange(1, 1, data.length + 1, maxCols).setValues([header, ...data]);
}

/** ============================================================
 *  API WEBHOOK
 * ============================================================ */

/** Dashboard is served at /exec; JSON API requests use ?api=1 and optional ?table=Name. */
function jsonOutput(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}

function authorised(request) {
  const expected = PropertiesService.getScriptProperties().getProperty('apiKey') || '';
  if (!expected) return true;
  const supplied = request?.parameter?.apiKey || '';
  return supplied === expected;
}

/**
 * Same-project transport for the HTML Service dashboard. ContentService
 * responses redirect to script.googleusercontent.com, which browsers can
 * reject as a cross-origin fetch. google.script.run avoids that redirect.
 */
function spreadsheetRpc(request) {
  const method = String(request?.method || 'GET').toUpperCase();
  const apiRequest = method === 'GET'
    ? { parameter: { api: '1', table: request?.table || '', apiKey: request?.apiKey || '' }, pathInfo: '' }
    : { parameter: {}, pathInfo: '', postData: { contents: JSON.stringify(request?.payload || {}) } };
  const output = method === 'GET' ? doGet(apiRequest) : doPost(apiRequest);
  return output.getContent();
}

function doGet(request) {
  const params = request?.parameter || {};
  if (String(params.api || '') !== '1') {
    return HtmlService.createHtmlOutputFromFile('Dashboard')
      .setTitle('Administrasi Persuratan | SMP Negeri 13 Tasikmalaya');
  }
  if (!authorised(request)) return jsonOutput({ ok: false, error: 'API key tidak valid.' });
  // Apps Script deployments may expose a route in pathInfo, while callers can
  // also use ?table=Tasks. Accept both and normalize optional surrounding slashes.
  const path = String(request?.pathInfo || params.table || '')
    .replace(/^\/+|\/+$/g, '').trim().toLowerCase();
  if (!path) return jsonOutput({
    spreadsheetId: CONFIG.spreadsheetId,
    capabilities: { safeIntegrationTest: true },
    tables: CONFIG.tables.map((table) => ({ name: table.name, title: table.title })),
  });
  const table = CONFIG.tables.find((item) => item.name.toLowerCase() === path);
  if (!table) return jsonOutput({ ok: false, error: 'Table not found' });
  return jsonOutput(sheetToRows(replaceSheet(table)));
}

/** Post to a table; body must include either { rows: [...] } or { table: { rows: [...] } }. */
function doPost(request) {
  let payload;
  try { payload = JSON.parse(request?.postData?.contents || '{}'); }
  catch (_) { return jsonOutput({ ok: false, error: 'JSON request tidak valid.' }); }
  if (!authorised({ parameter: { apiKey: payload.apiKey || '' } })) {
    return jsonOutput({ ok: false, error: 'API key tidak valid.' });
  }
  const tableName = String(payload.table || request?.pathInfo || request?.parameter?.table || '')
    .replace(/^\/+|\/+$/g, '').trim().toLowerCase();
  if (payload.action === 'config' && !tableName) return jsonOutput({
    spreadsheetId: CONFIG.spreadsheetId,
    capabilities: { safeIntegrationTest: true },
    tables: CONFIG.tables.map((table) => ({ name: table.name, title: table.title })),
  });
  if (!tableName) {
    return jsonOutput({ ok: false, error: 'Nama tabel wajib diisi.' });
  }
  const table = CONFIG.tables.find((t) => t.name.toLowerCase() === tableName);
  if (!table) {
    return jsonOutput({ ok: false, error: 'Table not found' });
  }
  if (payload.action === 'writeTestRow') {
    const row = payload.row || {};
    const id = String(row.id || '');
    const expectedSubject = `[UJI INTEGRASI AGENDA] ${id}`;
    if (table.name !== 'Incoming' || !/^agenda-integration-test-[\w-]+$/.test(id)
      || row.sender !== 'Agenda Surat Integration Test' || row.subject !== expectedSubject
      || row.disposition !== 'AUTO DELETE TEST ROW' || row.letterNumber !== id) {
      return jsonOutput({ ok: false, error: 'Penulisan hanya diizinkan untuk record uji Incoming yang ditandai.' });
    }
    const sheet = replaceSheet(table);
    if (sheetToRows(sheet).some((existing) => String(existing.id) === id)) {
      return jsonOutput({ ok: false, error: 'ID baris uji sudah ada; tidak ada baris yang ditambahkan.' });
    }
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowValues = headers.map((header) => row[header] === undefined || row[header] === null ? '' : row[header]);
    sheet.appendRow(rowValues);
    return jsonOutput({ ok: true, id });
  }
  if (payload.action === 'deleteTestRow') {
    const id = String(payload.id || '');
    if (table.name !== 'Incoming' || !/^agenda-integration-test-[\w-]+$/.test(id)) {
      return jsonOutput({ ok: false, error: 'Penghapusan hanya diizinkan untuk ID baris uji Incoming.' });
    }
    const sheet = replaceSheet(table);
    const target = sheetToRows(sheet).find((row) => String(row.id) === id);
    if (!target) return jsonOutput({ ok: true, deleted: false });
    if (target.sender !== 'Agenda Surat Integration Test' || target.subject !== `[UJI INTEGRASI AGENDA] ${id}` || target.disposition !== 'AUTO DELETE TEST ROW') {
      return jsonOutput({ ok: false, error: 'Baris cocok ID tetapi penanda uji tidak cocok; tidak ada data yang dihapus.' });
    }
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
    const idColumn = values[0].findIndex((header) => String(header) === 'id');
    const rowIndex = values.findIndex((row, index) => index > 0 && String(row[idColumn]) === id);
    if (idColumn < 0 || rowIndex < 1) return jsonOutput({ ok: true, deleted: false });
    sheet.deleteRow(rowIndex + 1);
    return jsonOutput({ ok: true, deleted: true, id });
  }
  if (payload.action === 'read') return jsonOutput(sheetToRows(replaceSheet(table)));
  const rows = payload.rows || payload.table?.rows || [];
  if (!Array.isArray(rows)) {
    return jsonOutput({ ok: false, error: 'Payload must contain an array of rows under rows or table.rows' });
  }
  try {
    // Normalise rows, keep only known keys + id.
    const accepted = new Set(table.fields.map((f) => f.key));
    accepted.add('id');
    const normalised = rows.map((row) => {
      if (row == null) return {};
      if (Array.isArray(row)) return {}; // malformed row
      const obj = {};
      for (const [k, v] of Object.entries(row)) {
        if (accepted.has(k)) obj[k] = v;
      }
      return obj;
    });
    // Only keep rows that have a real id (client guarantees one).
    const usable = normalised.filter((row) => row.id != null && row.id !== '');
    if (usable.length === 0) {
      return jsonOutput({ ok: true, affected: [], note: 'No rows with an id were provided' });
    }
    // Merge by id and preserve every existing row not included in this write.
    const sheet = replaceSheet(table);
    const existing = sheetToRows(sheet);
    const byId = new Map(existing.map((row) => [String(row.id), row]));
    for (const row of usable) byId.set(String(row.id), { ...byId.get(String(row.id)), ...row });
    rowsToSheet(sheet, Array.from(byId.values()));
    return jsonOutput({ ok: true, affected: usable.map((row) => row.id) });
  } catch (error) {
    return jsonOutput({ ok: false, error: error.message });
  }
}

/** Replace the sheet for a table; create it if missing (auto-scaling). */
function replaceSheet(table) {
  const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  let sheet = ss.getSheetByName(table.name);
  if (!sheet) {
    sheet = ss.insertSheet(table.name);
    sheet.setColumnWidth(1, 160);
    table.fields.forEach((f, i) => {
      sheet.setColumnWidth(i + 2, 220);
    });
  }
  ensureTableHeaders(sheet, table);
  return sheet;
}

/** ============================================================
 *  LOCK THROTTLING (per-table, per-session)
 * ============================================================ */

/** Key under which per-table locks are tracked for the current session. */
const LOCK_SESSION_KEY = 'lockState';

/** Acquire a per-table lock with a 30 second timeout; throws if unavailable. */
function acquireTableLock(tableName, timeoutMs) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(timeoutMs || 30000)) {
    throw new Error('Tabel sedang dikunci oleh proses lain. Coba lagi dalam sebentar.');
  }
  return lock;
}

/** Release a lock if it belongs to us. */
function releaseTableLock(lock) {
  if (lock && lock.hasLock()) {
    lock.releaseLock();
  }
}

/** Return the current UTC time as an ISO-8601 string (server-side source of truth). */
function getCurrentTime() {
  return new Date().toISOString();
}

/** Send a low-urgency notice when a table changes. */
function notifyTableChange(tableName, affectedCount) {
  // Optional: fire an email or chat message. Leave empty when the script is
  // activated. If you want a notification, uncomment the lines below and set
  // the recipient under Settings > Environment variables (scriptEmail).
  // const email = PropertiesService.getScriptProperties().getProperty('scriptEmail');
  // if (email) MailApp.sendEmail(email, `Data ${tableName} diperbarui`, `${affectedCount} baris terbarui di spreadsheet.`);
}

/** Export the normalized table config for legacy callers; never return the API key. */
function getConfig() {
  const clean = {
    spreadsheetId: CONFIG.spreadsheetId || '',
    capabilities: { safeTestRowDelete: true },
    tables: CONFIG.tables.map((t) => ({
      name: t.name,
      title: t.title,
      fields: t.fields.map((f) => ({ key: f.key, label: f.label, required: f.required })),
    })),
    storageVersion: 'v2',
    migrated: true,
  };
  return clean;
}

/** No deployment URL is embedded in source; copy it from the Apps Script deployment screen. */
function getUrlBase() {
  return '';
}

/** Ensure the configured spreadsheet exists; throws a clear error otherwise. */
function findSpreadsheet() {
  const id = CONFIG.spreadsheetId || PropertiesService.getScriptProperties().getProperty('spreadsheetId') || '';
  if (!id) {
    throw new Error('Tidak ada ID Spreadsheet yang dikonfigurasi. Set spreadsheetId (atau variabel spreadsheetId) di Project Settings > Environment variables, lalu jalankan sekali ulang.');
  }
  const ss = SpreadsheetApp.openById(id);
  return ss;
}



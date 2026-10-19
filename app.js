import Papa from 'papaparse';

/**
 * Google Spreadsheet backend.
 *
 * Configure the actual deployed Apps Script Web app URL and API key in Settings.
 * Spreadsheet ID and Apps Script project ID are not Web app URLs.
 *
 * The legacy localStorage blob is left untouched when remote storage is active.
 */

let API_BASE_URL = 'https://script.google.com/macros/s/AKfycbxiYMlYcacndRODXBax6kYf6BHGdBblYmCkYy9NwApT6fbGB7qu6Zg_vMBdnYdeWeUjQQ/exec';
let API_KEY = '';
const API_TIMEOUT_MS = 8000;

const STORAGE_KEY = 'smpn13.agenda-surat.v1';
const CURRENT_YEAR = new Date().getFullYear();
const PAGE_LABELS = { overview: 'Ringkasan', archive: 'Arsip surat tugas', agenda: 'Agenda surat', settings: 'Pengaturan & data' };
const defaultState = {
  tasks: [],
  incoming: [],
  outgoing: [],
  teachers: [],
  students: [],
  officialsMaster: [],
  officials: {
    principalName: '',
    principalNip: '',
    principalStartDate: '',
    treasurerName: '',
    treasurerNip: '',
    treasurerStartDate: '',
  },
  officialHistory: {
    principal: [{ name: '', nip: '', startDate: '1900-01-01' }],
    treasurer: [{ name: '', nip: '', startDate: '1900-01-01' }],
  },
  lastIssuedByYear: { 2025: 437 },
};

/**
 * Spreadsheet column helpers.
 * Sheets read/write the header as the first row of the values array.
 * Empty cells return '' (never null) so the client can tell "empty" from
 * "blank cell". Rows are objects with an explicit `id` so updates overwrite
 * a record instead of duplicating it.
 */

function nullGuardValues(values) {
  if (!Array.isArray(values)) return [];
  return values.map((row) => (row == null ? [] : row.map((v) => (v === null || v === undefined ? '' : v))));
}

function valueRowsToMatrix(rows) {
  if (!Array.isArray(rows)) return [];
  const header = [];
  const fields = [];
  if (rows.length) {
    const keys = Object.keys(rows[0]);
    header.push('id', ...keys.filter((k) => k !== 'id'));
    fields.push('id', ...keys.filter((k) => k !== 'id'));
  }
  const data = rows.map((row) => {
    const arr = [];
    header.forEach((h) => {
      arr.push(row[h] === undefined || row[h] === null ? '' : row[h]);
    });
    return arr;
  });
  return [header, ...data];
}

function matrixToRows(values) {
  if (!Array.isArray(values) || values.length < 1) return [];
  const header = Array.isArray(values[0]) ? values[0] : [];
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r] || [];
    const obj = {};
    header.forEach((h, i) => {
      const raw = row[i];
      obj[h] = raw === undefined ? '' : String(raw);
    });
    if (Object.keys(obj).length === 0) continue;
    rows.push(obj);
  }
  return rows;
}

/**
 * Which Google Sheet table holds which app list. `id` is always the first
 * column.
 */
const STORAGE_TABLES = {
  tasks: { label: 'Tasks', fields: ['number', 'basis', 'basisNumber', 'basisDate', 'recipient', 'subject', 'eventDate', 'eventName', 'eventDay', 'eventStartDate', 'eventEndDate', 'eventTime', 'place', 'origin', 'destination', 'travelPurpose', 'departureDate', 'returnDate', 'transportMode', 'transportDays', 'budgetSource', 'itinerary', 'reportText', 'reportTo', 'recommendations', 'attachment', 'issuedDate', 'year', 'createdAt', 'updatedAt'] },
  incoming: { label: 'Incoming', fields: ['receivedDate', 'letterDate', 'letterNumber', 'sender', 'subject', 'disposition', 'attachment', 'createdAt'] },
  outgoing: { label: 'Outgoing', fields: ['issuedDate', 'number', 'letterType', 'recipient', 'subject', 'notes', 'attachment', 'createdAt', 'updatedAt'] },
  teachers: { label: 'Masters', kind: 'teacher', fields: ['name', 'grade', 'rank', 'position', 'nip', 'nuptk', 'createdAt'] },
  students: { label: 'Masters', kind: 'student', fields: ['name', 'class', 'nis', 'createdAt'] },
  settings: { label: 'Settings', fields: ['officials', 'officialHistory', 'officialsMaster', 'updatedAt'] },
};

/**
 * Storage transport abstraction.
 * The rest of the app never touches localStorage directly; it goes through
 * loadStorage/saveStorage/readTable/writeTable/readAll/saveAllData. The default
 * transport is the in-memory local blob. Switching to the spreadsheet transport
 * is done via reconnectStorage(transport).
 */
let activeTransport = null;

async function initStorage() {
  const saved = await loadStorage();
  const remote = await readAll();
  if (remote && typeof remote === 'object') {
    const merged = syncFromRemote(remote);
    // Number counters remain local; signing officials are loaded from Settings.
    if (saved?.lastIssuedByYear) merged.lastIssuedByYear = { ...merged.lastIssuedByYear, ...saved.lastIssuedByYear };
    if (!merged.officialsMaster.length && saved?.officialsMaster?.length) merged.officialsMaster = saved.officialsMaster;
    return merged;
  }
  return saved;
}

function syncFromRemote(payload) {
  const next = structuredClone(defaultState);
  if (payload && typeof payload === 'object') {
    next.tasks = Array.isArray(payload.tasks) ? payload.tasks : [];
    next.incoming = Array.isArray(payload.incoming) ? payload.incoming : [];
    next.outgoing = Array.isArray(payload.outgoing) ? payload.outgoing : [];
    next.teachers = Array.isArray(payload.teachers) ? payload.teachers : [];
    next.students = Array.isArray(payload.students) ? payload.students : [];
    if (payload.officials && typeof payload.officials === 'object') next.officials = { ...defaultState.officials, ...payload.officials };
    if (Array.isArray(payload.officialsMaster)) next.officialsMaster = payload.officialsMaster;
    else if (payload.officialsMaster && typeof payload.officialsMaster === 'string') { try { next.officialsMaster = JSON.parse(payload.officialsMaster); } catch (_) {} }
    if (payload.officialHistory && typeof payload.officialHistory === 'object') {
      next.officialHistory = {
        principal: Array.isArray(payload.officialHistory.principal) ? payload.officialHistory.principal : next.officialHistory.principal,
        treasurer: Array.isArray(payload.officialHistory.treasurer) ? payload.officialHistory.treasurer : next.officialHistory.treasurer,
      };
    }
    if (payload.lastIssuedByYear && typeof payload.lastIssuedByYear === 'object') next.lastIssuedByYear = { ...defaultState.lastIssuedByYear, ...payload.lastIssuedByYear };
  }
  return next;
}

/** Load the full state from the active transport. */
async function loadStorage() {
  if (activeTransport && typeof activeTransport.loadStorage === 'function') {
    return activeTransport.loadStorage();
  }
  return loadState();
}

/** Persist the whole state atomically on the remote sheet. */
async function saveStorage(next) {
  if (activeTransport && typeof activeTransport.saveStorage === 'function') {
    return activeTransport.saveStorage(next);
  }
  saveState();
}

/** Read one spreadsheet table into a plain array. */
async function readTable(table) {
  if (activeTransport && typeof activeTransport.readTable === 'function') {
    return activeTransport.readTable(table);
  }
  return [];
}

/** Write one spreadsheet table (merge by id). */
async function writeTable(table, rows) {
  if (activeTransport && typeof activeTransport.writeTable === 'function') {
    return activeTransport.writeTable(table, rows);
  }
}

/** Read all tables at once. */
async function readAll() {
  if (activeTransport && typeof activeTransport.readAll === 'function') {
    return activeTransport.readAll();
  }
  return null;
}

/** Write all tables at once (used by migration). */
async function saveAllData(rowsByTable) {
  if (activeTransport && typeof activeTransport.saveAllData === 'function') {
    return activeTransport.saveAllData(rowsByTable);
  }
}

/** Flush the current localStorage content into the spreadsheet (never deletes local data). */
async function migrateFromLocalStorage(exported) {
  if (activeTransport && typeof activeTransport.migrateFromLocalStorage === 'function') {
    return activeTransport.migrateFromLocalStorage(exported);
  }
  const tasks = (exported.tasks || []).map((t) => ({ id: String(t.id || crypto.randomUUID()), ...t }));
  const incoming = (exported.incoming || []).map((t) => ({ id: String(t.id || crypto.randomUUID()), ...t }));
  const outgoing = (exported.outgoing || []).map((t) => ({ id: String(t.id || crypto.randomUUID()), ...t }));
  return { tasks, incoming, outgoing };
}

/** Swap the active transport at runtime (e.g., after the user selects "Spreadsheet"). */
function reconnectStorage(transport) {
  activeTransport = transport;
}

/** Read one row by id from a table. */
async function readRow(table, id) {
  const rows = await readTable(table);
  if (!Array.isArray(rows)) return null;
  return rows.find((r) => String(r.id) === String(id)) || null;
}

/** Write one row (upsert) into a table. */
async function writeRow(table, row) {
  const rows = await readTable(table);
  const idx = rows.findIndex((r) => String(r.id) === String(row.id));
  if (idx >= 0) rows[idx] = row;
  else rows.push(row);
  await writeTable(table, rows);
}

/** ---- HTTP helpers --------------------------------------------------- */

async function fetchJson(url, options) {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), API_TIMEOUT_MS || 8000);
  try {
    let text;
    let status = 200;
    // Apps Script ContentService redirects can fail browser CORS checks.
    // HTML Service pages can call their own project through google.script.run.
    if (typeof google !== 'undefined' && google.script?.run) {
      const requestUrl = new URL(url, window.location.href);
      const method = String(options?.method || 'GET').toUpperCase();
      let payload = {};
      if (method !== 'GET') {
        try { payload = JSON.parse(options?.body || '{}'); }
        catch (_) { throw new Error('Isi permintaan Spreadsheet tidak valid.'); }
      }
      text = await new Promise((resolve, reject) => {
        const runner = google.script.run
          .withSuccessHandler(resolve)
          .withFailureHandler((error) => reject(new Error(error?.message || 'Apps Script gagal memproses permintaan.')));
        runner.spreadsheetRpc({
          method,
          table: requestUrl.searchParams.get('table') || payload.table || '',
          apiKey: requestUrl.searchParams.get('apiKey') || payload.apiKey || '',
          payload,
        });
      });
    } else {
      const res = await fetch(url, { ...options, signal: ctrl.signal });
      status = res.status;
      text = await res.text();
      if (!res.ok) status = res.status;
    }
    let data = null;
    try { data = JSON.parse(text); } catch (_) { data = { raw: text }; }
    if (typeof data.raw === 'string' && /<(!doctype html|html[\s>])/i.test(data.raw)) {
      const err = new Error(`Web App mengembalikan halaman HTML (HTTP ${status}), bukan JSON. Deployment mungkin meminta login Google atau URL/izin Web App belum benar.`);
      err.status = status;
      throw err;
    }
    if (status < 200 || status >= 300) {
      const isHtml = typeof data.raw === 'string' && /<(!doctype html|html[\s>])/i.test(data.raw);
      const message = isHtml
        ? `Web App mengembalikan halaman HTML (HTTP ${status}), bukan JSON. Deployment mungkin meminta login Google atau URL/izin Web App belum benar.`
        : data.error || `HTTP ${status}`;
      const err = new Error(message);
      err.status = status;
      err.data = data;
      throw err;
    }
    return data;
  } catch (err) {
    if (err.name === 'AbortError') {
      const e = new Error('Permintaan ke penyedia data dibatalkan (waktu habis).');
      e.status = 504;
      throw e;
    }
    if (err instanceof TypeError && /fetch|network/i.test(err.message)) {
      const e = new Error('Browser gagal membaca Web App. Periksa izin akses deployment; redirect login Google lintas situs atau CORS dapat memblokir request.');
      e.cause = err;
      throw e;
    }
    throw err;
  } finally {
    clearTimeout(timer);
  }
}

function toSheetRows(table, rows) {
  const fields = STORAGE_TABLES[table] ? STORAGE_TABLES[table].fields : [];
  return valueRowsToMatrix(rows);
}


let state = loadState();
let activePage = 'overview';
let activeAgendaTab = 'incoming';
let activePrintTaskId = null;
let activeSppdPrintType = '';
let toastTimer;

function loadState() {
  try {
    const stored = JSON.parse(localStorage.getItem(STORAGE_KEY));
    if (!stored || typeof stored !== 'object') return structuredClone(defaultState);
    return {
      tasks: Array.isArray(stored.tasks) ? stored.tasks : [],
      incoming: Array.isArray(stored.incoming) ? stored.incoming : [],
      outgoing: Array.isArray(stored.outgoing) ? stored.outgoing : [],
      teachers: Array.isArray(stored.teachers) ? stored.teachers : [],
      students: Array.isArray(stored.students) ? stored.students : [],
      officialsMaster: Array.isArray(stored.officialsMaster) ? stored.officialsMaster : [],
      officials: { ...defaultState.officials, ...(stored.officials || {}) },
      officialHistory: {
        principal: Array.isArray(stored.officialHistory?.principal) ? stored.officialHistory.principal : [{ name: stored.officials?.principalName || defaultState.officials.principalName, nip: stored.officials?.principalNip || defaultState.officials.principalNip, startDate: stored.officials?.principalStartDate || '1900-01-01' }],
        treasurer: Array.isArray(stored.officialHistory?.treasurer) ? stored.officialHistory.treasurer : [{ name: stored.officials?.treasurerName || defaultState.officials.treasurerName, nip: stored.officials?.treasurerNip || defaultState.officials.treasurerNip, startDate: stored.officials?.treasurerStartDate || '1900-01-01' }],
      },
      lastIssuedByYear: { ...defaultState.lastIssuedByYear, ...(stored.lastIssuedByYear || {}) },
    };
  } catch {
    return structuredClone(defaultState);
  }
}

function saveState() {
  renderAll();
  if (activeTransport && typeof activeTransport.saveStorage === 'function') {
    // Mirror the complete state locally so fields not represented by remote
    // tables (official history and numbering counters) survive refreshes.
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); } catch (error) {
      showToast(`Penyimpanan lokal penuh/tidak dapat ditulis: ${error.message}`);
    }
    activeTransport.saveStorage(state).catch((error) => {
      showToast(`Penyimpanan spreadsheet gagal: ${error.message}`);
    });
  } else {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    } catch (error) {
      showToast(`Penyimpanan lokal penuh/tidak dapat ditulis: ${error.message}`);
    }
  }
}

/**
 * LocalStorage transport (legacy). Used by export/import backup flows;
 * the primary write path is the spreadsheet transport below.
 */
const localStore = {
  async loadStorage() {
    try {
      const stored = JSON.parse(localStorage.getItem(STORAGE_KEY));
      if (!stored || typeof stored !== 'object') return null;
      return {
        tasks: Array.isArray(stored.tasks) ? stored.tasks : [],
        incoming: Array.isArray(stored.incoming) ? stored.incoming : [],
        outgoing: Array.isArray(stored.outgoing) ? stored.outgoing : [],
        teachers: Array.isArray(stored.teachers) ? stored.teachers : [],
        students: Array.isArray(stored.students) ? stored.students : [],
        officialsMaster: Array.isArray(stored.officialsMaster) ? stored.officialsMaster : [],
        officials: stored.officials || null,
        officialHistory: stored.officialHistory || null,
        lastIssuedByYear: stored.lastIssuedByYear || null,
      };
    } catch {
      return null;
    }
  },
  async saveStorage(next) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
    renderAll();
  },
};


async function readAllFromSheet() {
  if (!API_BASE_URL) return null;
  const data = await fetchConfig();
  if (!data || !Array.isArray(data.tables)) return null;
  const rows = { tasks: [], incoming: [], outgoing: [], teachers: [], students: [], officials: null, officialHistory: null, officialsMaster: [] };
  for (const table of data.tables) {
    const list = await fetchJson(apiUrl(table.name, true), { headers: { Accept: 'application/json' }, method: 'GET' });
    if (!Array.isArray(list)) throw new Error(list?.error || `Respons tabel ${table.name} tidak valid.`);
    const values = list;
    if (table.name === 'Tasks') rows.tasks = values.map((row) => ({ ...row, assignees: parseJsonArray(row.assignees), students: parseJsonArray(row.students) }));
    else if (table.name === 'Incoming') rows.incoming = values;
    else if (table.name === 'Outgoing') rows.outgoing = values;
    else if (table.name === 'Masters') {
      rows.teachers = values.filter((row) => row.kind === 'teacher');
      rows.students = values.filter((row) => row.kind === 'student');
    } else if (table.name === 'Settings') {
      const settings = values.find((row) => row.id === 'application');
      if (settings) {
        rows.officials = parseJsonObject(settings.officials);
        rows.officialHistory = parseJsonObject(settings.officialHistory);
        rows.officialsMaster = parseJsonArray(settings.officialsMaster);
      }
    }
  }
  return rows;
}

function parseJsonArray(value) {
  if (Array.isArray(value)) return value;
  if (!value) return [];
  try { const parsed = JSON.parse(value); return Array.isArray(parsed) ? parsed : []; } catch { return []; }
}

function parseJsonObject(value) {
  if (value && typeof value === 'object') return value;
  if (!value) return null;
  try { const parsed = JSON.parse(value); return parsed && typeof parsed === 'object' && !Array.isArray(parsed) ? parsed : null; } catch { return null; }
}

function apiUrl(path = '', includeApiKey = false) {
  const url = new URL(API_BASE_URL, window.location.href);
  url.searchParams.set('api', '1');
  if (path) url.searchParams.set('table', path);
  else url.searchParams.delete('table');
  if (includeApiKey && API_KEY) url.searchParams.set('apiKey', API_KEY);
  else url.searchParams.delete('apiKey');
  return url.toString();
}

async function fetchConfig() {
  if (!API_BASE_URL) return null;
  return fetchJson(apiUrl('', true), { headers: { Accept: 'application/json' }, method: 'GET' });
}

async function writeTableRows(table, rows) {
  const response = await fetchJson(apiUrl(), {
    method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ table, rows, apiKey: API_KEY }),
  });
  if (!response?.ok) throw new Error(response?.error || `Gagal menulis tabel ${table}.`);
  return response;
}

async function deleteIntegrationTestRow(id) {
  return fetchJson(apiUrl(), {
    method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ table: 'Incoming', action: 'deleteTestRow', id, apiKey: API_KEY }),
  });
}

async function appendIntegrationTestRow(row) {
  const response = await fetchJson(apiUrl(), {
    method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ table: 'Incoming', action: 'writeTestRow', row, apiKey: API_KEY }),
  });
  if (!response?.ok || response.id !== row.id) throw new Error(response?.error || 'Backend tidak mengonfirmasi baris uji.');
  return response;
}

async function writeAllToSheet(rowsByTable) {
  if (!API_BASE_URL || API_BASE_URL.includes('{WEB_APP_ID}')) throw new Error('Isi URL deployment Web App Apps Script yang sebenarnya.');
  const names = { tasks: 'Tasks', incoming: 'Incoming', outgoing: 'Outgoing', settings: 'Settings' };
  const payload = new Map();
  for (const [table, rows] of Object.entries(rowsByTable)) {
    if (!Array.isArray(rows)) continue;
    const withIds = rows.map((row) => ({ ...row, id: String(row?.id || crypto.randomUUID()) }));
    if (table === 'teachers' || table === 'students') {
      payload.set('Masters', [...(payload.get('Masters') || []), ...withIds.map((row) => ({ ...row, kind: table === 'teachers' ? 'teacher' : 'student' }))]);
    } else if (names[table]) {
      payload.set(names[table], withIds);
    }
  }
  const results = await Promise.all([...payload].map(([table, rows]) => writeTableRows(table, rows.map((row) => ({ ...row,
      ...(Array.isArray(row.assignees) ? { assignees: JSON.stringify(row.assignees) } : {}),
      ...(Array.isArray(row.students) ? { students: JSON.stringify(row.students) } : {}),
      ...(row.officials && typeof row.officials === 'object' ? { officials: JSON.stringify(row.officials) } : {}),
      ...(row.officialHistory && typeof row.officialHistory === 'object' ? { officialHistory: JSON.stringify(row.officialHistory) } : {}),
      ...(Array.isArray(row.officialsMaster) ? { officialsMaster: JSON.stringify(row.officialsMaster) } : {}),
    })))));
  return { ok: true, written: results.length };
}

function mergeRows(existing, incoming) {
  const map = new Map();
  (existing || []).forEach((r) => map.set(String(r.id), r));
  (incoming || []).forEach((r) => {
    if (r && r.id != null) map.set(String(r.id), r);
  });
  return Array.from(map.values());
}

async function spreadsheetTransportInit() {
  const cfg = await fetchConfig();
  if (!cfg || !Array.isArray(cfg.tables)) {
    throw new Error('Respons Apps Script tidak berisi konfigurasi tabel yang valid.');
  }
  return {
    async loadStorage() {
      const data = await readAllFromSheet();
      return data || null;
    },
    async saveStorage(next) {
      await writeAllToSheet({
        ...next,
        settings: [{
          id: 'application',
          officials: next.officials || {},
          officialHistory: next.officialHistory || {},
          officialsMaster: next.officialsMaster || [],
          updatedAt: new Date().toISOString(),
        }],
      });
    },
    async readTable(table) {
      const data = await readAllFromSheet();
      if (!data || !Array.isArray(data[table])) return [];
      return data[table];
    },
    async writeTable(table, rows) {
      await writeAllToSheet({ [table]: rows });
    },
    async readAll() {
      return readAllFromSheet();
    },
    async saveAllData(rowsByTable) {
      return writeAllToSheet(rowsByTable);
    },
    async migrateFromLocalStorage(exported) {
      const tasks = (exported.tasks || []).map((t) => ({ id: String(t.id || crypto.randomUUID()), ...t }));
      const incoming = (exported.incoming || []).map((t) => ({ id: String(t.id || crypto.randomUUID()), ...t }));
      const outgoing = (exported.outgoing || []).map((t) => ({ id: String(t.id || crypto.randomUUID()), ...t }));
      const teachers = (exported.teachers || []).map((t) => ({ id: String(t.id || crypto.randomUUID()), ...t }));
      const students = (exported.students || []).map((t) => ({ id: String(t.id || crypto.randomUUID()), ...t }));
      const settings = [{ id: 'application', officials: exported.officials || {}, officialHistory: exported.officialHistory || {}, officialsMaster: exported.officialsMaster || [], updatedAt: new Date().toISOString() }];
      return { tasks, incoming, outgoing, teachers, students, settings };
    },
  };
}

/** Activate the spreadsheet transport after the user has configured the API URL + key. */
async function connectStorageToSpreadsheet() {
  if (!API_BASE_URL || API_BASE_URL.includes('{WEB_APP_ID}')) {
    showToast('Masukkan URL deployment Web App Apps Script yang sebenarnya. ID Spreadsheet/Apps Script bukan URL deployment.');
    return false;
  }
  try {
    const cfg = await fetchConfig();
    if (!cfg || !Array.isArray(cfg.tables)) throw new Error('URL tidak mengembalikan konfigurasi Apps Script yang valid.');
    const transport = await spreadsheetTransportInit();
    const remote = await transport.readAll();
    if (!remote) throw new Error('Tidak dapat membaca tabel dari Spreadsheet.');
    reconnectStorage(transport);
    const localOfficialsForConnect = state.officials;
    const localHistoryForConnect = state.officialHistory;
    const localOfficialsMasterForConnect = state.officialsMaster;
    const localIssuedNumbers = state.lastIssuedByYear;
    state = syncFromRemote(remote);
    state.lastIssuedByYear = localIssuedNumbers;
    // Seed the shared Settings row from this browser only when the sheet has none.
    if (!remote.officials) {
      state.officials = localOfficialsForConnect;
      state.officialHistory = localHistoryForConnect;
    }
    if (!state.officialsMaster.length && localOfficialsMasterForConnect?.length) state.officialsMaster = localOfficialsMasterForConnect;
    if (!remote.officials || !(remote.officialsMaster || []).length) {
      await writeAllToSheet({ settings: [{ id: 'application', officials: state.officials, officialHistory: state.officialHistory, officialsMaster: state.officialsMaster || [], updatedAt: new Date().toISOString() }] });
    }
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); } catch (error) {
      showToast(`Salinan pengaturan lokal tidak dapat ditulis: ${error.message}`);
    }
    renderAll();
    document.getElementById('spreadsheet-status').textContent = 'Koneksi berhasil. Data aktif dibaca dari Spreadsheet; salinan lokal tetap dipertahankan.';
    showToast('Koneksi ke Spreadsheet berhasil dan tabel berhasil dibaca.');
    return true;
  } catch (error) {
    document.getElementById('spreadsheet-status').textContent = `Koneksi gagal: ${error.message}`;
    showToast(`Koneksi gagal: ${error.message}`);
    return false;
  }
}

/** Export the full in-memory state as a JSON blob for backup/migration. */
function exportStateBlob() {
  return JSON.stringify({ app: 'Administrasi Persuratan SMPN 13 Tasikmalaya', version: 1, exportedAt: new Date().toISOString(), ...state }, null, 2);
}

/**
 * Import a JSON backup blob into state (without touching the spreadsheet).
 * Merges with existing data for tasks/incoming/outgoing/teachers/students, then
 * updates `state` and renders. Safe to call while the spreadsheet transport is
 * active (local-only import).
 */
function importStateBlob(parsed) {
  if (!parsed || typeof parsed !== 'object') return;
  if (Array.isArray(parsed) || parsed instanceof Blob) {
    showToast('Format cadangan tidak valid. Unggah file JSON, bukan teks mentah.');
    return;
  }
  const imported = parsed;
  const applyTable = (key) => {
    if (Array.isArray(imported[key])) {
      state[key] = imported[key].map((row) => (row && row.id != null ? row : { id: crypto.randomUUID(), ...row }));
    } else {
      state[key] = [];
    }
  };
  applyTable('tasks');
  applyTable('incoming');
  applyTable('outgoing');
  applyTable('teachers');
  applyTable('students');
  if (imported.officials && typeof imported.officials === 'object') state.officials = { ...defaultState.officials, ...imported.officials };
  if (imported.officialHistory && typeof imported.officialHistory === 'object') {
    state.officialHistory = {
      principal: Array.isArray(imported.officialHistory.principal) ? imported.officialHistory.principal : state.officialHistory.principal,
      treasurer: Array.isArray(imported.officialHistory.treasurer) ? imported.officialHistory.treasurer : state.officialHistory.treasurer,
    };
  }
  if (imported.lastIssuedByYear && typeof imported.lastIssuedByYear === 'object') state.lastIssuedByYear = { ...defaultState.lastIssuedByYear, ...imported.lastIssuedByYear };
}

function downloadStateBlob(filename, content) {
  downloadFile(filename, content, 'application/json;charset=utf-8');
}

function escapeHtml(value = '') {
  return String(value).replace(/[&<>"']/g, (character) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[character]);
}

function localDateValue(date = new Date()) {
  const local = new Date(date.getTime() - date.getTimezoneOffset() * 60000);
  return local.toISOString().slice(0, 10);
}

function formatDate(value) {
  if (!value) return '—';
  const date = new Date(`${value}T00:00:00`);
  if (Number.isNaN(date.getTime())) return value;
  return new Intl.DateTimeFormat('id-ID', { day: 'numeric', month: 'short', year: 'numeric' }).format(date);
}

function showToast(message) {
  const toast = document.getElementById('toast');
  toast.textContent = message;
  toast.classList.add('is-visible');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.remove('is-visible'), 2800);
}

function navigate(page) {
  if (!PAGE_LABELS[page]) return;
  activePage = page;
  document.querySelectorAll('.page').forEach((section) => section.classList.toggle('is-visible', section.id === `page-${page}`));
  document.querySelectorAll('.nav-link[data-page]').forEach((button) => button.classList.toggle('is-active', button.dataset.page === page));
  document.getElementById('page-breadcrumb').textContent = PAGE_LABELS[page];
  // Apps Script HtmlService runs this page inside its sandboxed iframe. Avoid
  // changing that frame's URL during navigation; the dashboard routes locally.
  if (window.top === window.self) window.location.hash = page === 'overview' ? 'ringkasan' : page;
  renderAll();
}

// Bind page navigation directly as well as through the delegated click handler.
// This keeps Apps Script HtmlService iframe navigation independent of hash routing.
document.querySelectorAll('[data-page]').forEach((button) => {
  button.addEventListener('click', (event) => {
    event.preventDefault();
    event.stopPropagation();
    navigate(button.dataset.page);
  });
});

function taskSearchText(task) {
  return [task.number, task.subject, task.basis, task.basisNumber, assigneeNames(task.assignees), assigneeNames(task.students), task.eventName, task.place, task.recipient].join(' ').toLowerCase();
}

function parseRoster(value, fields) {
  if (Array.isArray(value)) return value.map((record) => fields.map((field) => String(record?.[field] ?? '').trim())).filter((parts) => parts[0]);
  return String(value || '').split(/\r?\n/).map((line) => line.split('|').map((part) => part.trim())).filter((parts) => parts[0]);
}

function parseAssignees(value = '') {
  return parseRoster(value, ['name', 'nip', 'nuptk', 'grade', 'rank', 'position', 'rate', 'duk']);
}

function parseStudents(value = '') {
  return parseRoster(value, ['name', 'nis', 'class', 'rate']);
}

function assigneeNames(value = '') {
  return parseAssignees(value).map((parts) => parts[0]).join(', ');
}

function renderOverview() {
  const yearTasks = state.tasks.filter((task) => Number(task.year) === CURRENT_YEAR);
  const recentTasks = [...state.tasks].sort((first, second) => (second.issuedDate || '').localeCompare(first.issuedDate || '')).slice(0, 5);
  const recentRows = document.getElementById('recent-task-rows');
  recentRows.innerHTML = recentTasks.map((task) => `<tr>
    <td>${escapeHtml(task.number)}</td>
    <td><span class="row-main">${escapeHtml(task.subject || 'Tanpa perihal')}</span><span class="row-sub">${escapeHtml(task.eventName || 'Surat tugas')}</span></td>
    <td>${escapeHtml(formatDate(task.issuedDate))}</td>
    <td>${escapeHtml(assigneeNames(task.assignees) || '—')}</td>
    <td><button class="icon-button" data-action="print-options" data-id="${escapeHtml(task.id)}" aria-label="Pilih dokumen cetak ${escapeHtml(task.number)}" title="Pilih dokumen cetak">▤</button></td>
  </tr>`).join('');
  document.getElementById('recent-empty').classList.toggle('is-visible', recentTasks.length === 0);
  document.querySelector('.recent-section .table-wrap').style.display = recentTasks.length ? '' : 'none';
  document.getElementById('metric-tasks').textContent = String(state.tasks.length);
  document.getElementById('metric-incoming').textContent = String(state.incoming.length);
  document.getElementById('metric-outgoing').textContent = String(state.tasks.length + state.outgoing.length);
  document.getElementById('overview-year').textContent = CURRENT_YEAR;
  document.getElementById('metric-teachers').textContent = String(state.teachers.length);
  document.getElementById('overview-teacher-count').textContent = `${state.teachers.length} guru`;
  document.getElementById('overview-student-count').textContent = `${state.students.length} siswa`;
  const agendaRows = [
    ...state.incoming.map((entry) => ({ kind: 'Masuk', title: entry.subject, detail: `${entry.sender || 'Pengirim belum diisi'} · ${formatDate(entry.receivedDate)}`, date: entry.receivedDate })),
    ...state.tasks.map((entry) => ({ kind: 'Keluar', title: entry.subject, detail: `${entry.number} · ${formatDate(entry.issuedDate)}`, date: entry.issuedDate })),
    ...state.outgoing.map((entry) => ({ kind: 'Keluar', title: entry.subject, detail: `${entry.number} · ${formatDate(entry.issuedDate)}`, date: entry.issuedDate })),
  ].sort((first, second) => (second.date || '').localeCompare(first.date || '')).slice(0, 3);
  document.getElementById('recent-agenda').innerHTML = agendaRows.length ? agendaRows.map((entry) => `<div class="agenda-item"><span class="agenda-item-type">${escapeHtml(entry.kind)}</span><strong>${escapeHtml(entry.title || 'Tanpa perihal')}</strong><small>${escapeHtml(entry.detail)}</small></div>`).join('') : '<div class="agenda-item"><span class="agenda-item-type">Belum ada aktivitas</span><strong>Agenda surat akan muncul di sini.</strong></div>';
}

function renderArchive() {
  const search = document.getElementById('archive-search').value.trim().toLowerCase();
  const selectedYear = document.getElementById('archive-year').value;
  const years = [...new Set(state.tasks.map((task) => String(task.year || task.issuedDate?.slice(0, 4))).filter(Boolean))].sort((first, second) => second.localeCompare(first));
  const yearSelect = document.getElementById('archive-year');
  const previousYear = yearSelect.value;
  yearSelect.innerHTML = '<option value="all">Semua tahun</option>' + years.map((year) => `<option value="${escapeHtml(year)}">${escapeHtml(year)}</option>`).join('');
  yearSelect.value = years.includes(previousYear) ? previousYear : 'all';
  const yearFilter = yearSelect.value;
  const matches = [...state.tasks].filter((task) => (!search || taskSearchText(task).includes(search)) && (yearFilter === 'all' || String(task.year || task.issuedDate?.slice(0, 4)) === yearFilter)).sort((first, second) => (second.issuedDate || '').localeCompare(first.issuedDate || ''));
  document.getElementById('archive-count').textContent = `${matches.length} arsip`;
  document.getElementById('archive-rows').innerHTML = matches.map((task) => `<tr>
    <td>${escapeHtml(task.number)}</td><td>${escapeHtml(formatDate(task.issuedDate))}</td>
    <td><span class="row-main">${escapeHtml(task.subject || 'Tanpa perihal')}</span><span class="row-sub">Dasar: ${escapeHtml(task.basis || '—')}</span></td>
    <td>${escapeHtml(assigneeNames(task.assignees) || '—')}</td><td><span class="row-main">${escapeHtml(task.eventName || '—')}</span><span class="row-sub">${escapeHtml(task.place || task.recipient || '—')}</span></td>
    <td class="action-cell"><span class="archive-actions"><button class="icon-button" data-action="edit-task" data-id="${escapeHtml(task.id)}" aria-label="Edit arsip ${escapeHtml(task.number)}" title="Edit arsip">✎</button><button class="icon-button" data-action="download-task" data-id="${escapeHtml(task.id)}" aria-label="Unduh data arsip ${escapeHtml(task.number)}" title="Unduh data arsip">↓</button><button class="icon-button" data-action="print-options" data-id="${escapeHtml(task.id)}" aria-label="Pilih dokumen cetak ${escapeHtml(task.number)}" title="Pilih dokumen cetak">▤</button></span></td>
  </tr>`).join('');
  document.getElementById('archive-empty').classList.toggle('is-visible', matches.length === 0);
  document.querySelector('#page-archive .data-table').style.display = matches.length ? '' : 'none';
}

function renderAgenda() {
  const isIncoming = activeAgendaTab === 'incoming';
  document.querySelectorAll('[data-agenda-tab]').forEach((button) => {
    const selected = button.dataset.agendaTab === activeAgendaTab;
    button.classList.toggle('is-selected', selected);
    button.setAttribute('aria-selected', String(selected));
  });
  document.getElementById('print-incoming-button').hidden = !isIncoming;
  document.getElementById('print-outgoing-button').hidden = isIncoming;
  document.getElementById('print-task-button').hidden = isIncoming;
  const addButton = document.getElementById('agenda-add-button');
  addButton.dataset.action = isIncoming ? 'new-incoming' : 'new-outgoing';
  addButton.innerHTML = `<span class="button-plus">+</span> ${isIncoming ? 'Catat surat masuk' : 'Catat surat keluar'}`;
  document.getElementById('incoming-tab-count').textContent = state.incoming.length;
  document.getElementById('outgoing-tab-count').textContent = state.tasks.length + state.outgoing.length;
  const query = document.getElementById('agenda-search').value.trim().toLowerCase();
  const outgoingRows = [
    ...state.tasks.map((task) => ({ ...task, recordType: 'task', letterNumber: task.number, issuedDate: task.issuedDate, recipient: task.recipient, activity: task.eventName || 'Surat Tugas' })),
    ...state.outgoing.map((entry) => ({ ...entry, recordType: 'outgoing', letterNumber: entry.number, activity: entry.letterType || 'Surat keluar' })),
  ];
  const rows = isIncoming ? state.incoming : outgoingRows;
  const filtered = rows.filter((entry) => [entry.letterNumber, entry.subject, entry.sender, entry.recipient, entry.disposition, entry.number, entry.letterType].join(' ').toLowerCase().includes(query))
    .sort((first, second) => ((isIncoming ? second.receivedDate : second.issuedDate) || '').localeCompare((isIncoming ? first.receivedDate : first.issuedDate) || ''));
  const head = isIncoming
    ? '<tr><th>Tanggal terima</th><th>Nomor surat</th><th>Tanggal surat</th><th>Pengirim</th><th>Perihal</th><th>Disposisi</th><th>Berkas</th></tr>'
    : '<tr><th>Tanggal keluar</th><th>Nomor surat</th><th>Jenis surat</th><th>Tujuan</th><th>Perihal</th><th class="action-column">Aksi</th></tr>';
  document.getElementById('agenda-head').innerHTML = head;
  document.getElementById('agenda-rows').innerHTML = filtered.map((entry) => isIncoming
    ? `<tr><td>${escapeHtml(formatDate(entry.receivedDate))}</td><td>${escapeHtml(entry.letterNumber || '—')}</td><td>${escapeHtml(formatDate(entry.letterDate))}</td><td>${escapeHtml(entry.sender || '—')}</td><td>${escapeHtml(entry.subject || '—')}</td><td>${escapeHtml(entry.disposition || '—')}</td><td>${entry.attachment?.startsWith('{') ? `<button class="icon-button" data-action="open-incoming-file" data-id="${escapeHtml(entry.id)}" title="Buka berkas">↗</button>` : '—'}</td></tr>`
    : `<tr><td>${escapeHtml(formatDate(entry.issuedDate))}</td><td>${escapeHtml(entry.number || '—')}</td><td>${escapeHtml(entry.activity || entry.letterType || 'Surat Tugas')}</td><td>${escapeHtml(entry.recipient || '—')}</td><td>${escapeHtml(entry.subject || '—')}</td><td class="action-cell">${entry.recordType === 'task' ? `<button class="icon-button" data-action="print-options" data-id="${escapeHtml(entry.id)}" aria-label="Pilih dokumen cetak" title="Pilih dokumen cetak">▤</button>` : `<button class="icon-button" data-action="edit-outgoing" data-id="${escapeHtml(entry.id)}" aria-label="Edit surat keluar" title="Edit surat keluar">✎</button>`}</td></tr>`).join('');
  document.getElementById('agenda-empty').classList.toggle('is-visible', filtered.length === 0);
  document.querySelector('#page-agenda .data-table').style.display = filtered.length ? '' : 'none';
  document.getElementById('agenda-empty-title').textContent = isIncoming ? 'Belum ada surat masuk' : 'Belum ada surat keluar';
  document.getElementById('agenda-empty-copy').textContent = isIncoming ? 'Catat surat yang diterima untuk mulai mengisi agenda.' : 'Catat surat keluar apa pun, termasuk yang bukan Surat Tugas.';
  document.querySelector('#agenda-empty [data-action]').dataset.action = isIncoming ? 'new-incoming' : 'new-outgoing';
  document.querySelector('#agenda-empty .text-button').innerHTML = `${isIncoming ? 'Catat surat masuk' : 'Catat surat keluar'} <span aria-hidden="true">→</span>`;
}

function renderSettings() {
  const officialsForm = document.getElementById('officials-form');
  for (const [name, value] of Object.entries(state.officials)) officialsForm.elements[name].value = value;
  const validityLabel = (date) => date === '1900-01-01' ? 'Sebelum riwayat' : formatDate(date);
  document.getElementById('principal-history').innerHTML = state.officialHistory.principal.map((official) => `<tr><td>${escapeHtml(validityLabel(official.startDate))}</td><td>${escapeHtml(official.name)}</td><td>${escapeHtml(official.nip)}</td></tr>`).join('');
  document.getElementById('treasurer-history').innerHTML = state.officialHistory.treasurer.map((official) => `<tr><td>${escapeHtml(validityLabel(official.startDate))}</td><td>${escapeHtml(official.name)}</td><td>${escapeHtml(official.nip)}</td></tr>`).join('');
  document.getElementById('teacher-master-count').textContent = `${state.teachers.length} guru`;
  document.getElementById('student-master-count').textContent = `${state.students.length} siswa`;
  document.getElementById('officials-master-list').innerHTML = state.officialsMaster.map((p, i) => `<tr><td>${escapeHtml(p.name)}</td><td>${escapeHtml(p.role)}</td><td>${escapeHtml(p.nip || '')}</td><td><button type="button" class="icon-button" data-action="delete-official" data-id="${i}" aria-label="Hapus petugas">×</button></td></tr>`).join('');
  document.getElementById('storage-record-count').textContent = `${state.tasks.length + state.incoming.length + state.outgoing.length} catatan tersimpan`;
}

function renderAll() {
  renderOverview();
  renderArchive();
  renderAgenda();
  renderSettings();
}

function fieldMarkup(label, name, options = {}) {
  const { type = 'text', placeholder = '', value = '', required = false, full = false, hint = '', readonly = false } = options;
  const classes = `field${full ? ' field-full' : ''}`;
  const input = type === 'textarea'
    ? `<textarea name="${name}" placeholder="${escapeHtml(placeholder)}" ${required ? 'required' : ''}>${escapeHtml(value)}</textarea>`
    : `<input name="${name}" type="${type}" value="${escapeHtml(value)}" placeholder="${escapeHtml(placeholder)}" ${required ? 'required' : ''} ${readonly ? 'readonly' : ''}>`;
  return `<label class="${classes}"><span>${label}${required ? ' *' : ''}</span>${input}${hint ? `<small>${hint}</small>` : ''}</label>`;
}

function rosterRows(type) {
  const isTeacher = type === 'teachers';
  const records = isTeacher ? state.teachers : state.students;
  return records.map((record) => `<option value="${escapeHtml(record.id)}" data-name="${escapeHtml(record.name)}" data-nip="${escapeHtml(record.nip || '')}" data-nuptk="${escapeHtml(record.nuptk || '')}" data-grade="${escapeHtml(record.grade || '')}" data-rank="${escapeHtml(record.rank || '')}" data-position="${escapeHtml(record.position || '')}" data-nis="${escapeHtml(record.nis || '')}" data-class="${escapeHtml(record.class || '')}">${escapeHtml(record.name)}</option>`).join('');
}

function renderAssignmentRows(type, assignments = [{}]) {
  const isTeacher = type === 'teachers';
  const containerId = isTeacher ? 'teacher-assignment-rows' : 'student-assignment-rows';
  const records = isTeacher ? state.teachers : state.students;
  const options = rosterRows(type);
  const extraOption = (assignment) => {
    const personId = assignment.personId || '';
    if (!personId || records.some((record) => record.id === personId) || !assignment.name) return '';
    return `<option value="${escapeHtml(personId)}" data-name="${escapeHtml(assignment.name)}" data-nip="${escapeHtml(assignment.nip || '')}" data-nuptk="${escapeHtml(assignment.nuptk || '')}" data-grade="${escapeHtml(assignment.grade || '')}" data-rank="${escapeHtml(assignment.rank || '')}" data-position="${escapeHtml(assignment.position || '')}" data-nis="${escapeHtml(assignment.nis || '')}" data-class="${escapeHtml(assignment.class || '')}" selected>${escapeHtml(assignment.name)}</option>`;
  };
  document.getElementById(containerId).innerHTML = assignments.map((assignment, index) => isTeacher
    ? `<div class="assignment-row teacher-row" data-assignment-row="teachers"><span class="assignment-index">${index + 1}</span><select data-person-select required aria-label="Nama petugas"><option value="">${records.length ? 'Pilih petugas' : 'Unggah data guru dahulu'}</option>${options}${extraOption(assignment)}</select><output data-rank-value>${escapeHtml(assignment.rank || '—')}</output><output data-position-value>${escapeHtml(assignment.position || '—')}</output><input data-rate-value type="number" min="0" step="1" value="${escapeHtml(assignment.rate ?? '')}" placeholder="Tarif" required aria-label="Tarif petugas per hari"><button type="button" class="icon-button assignment-remove" data-remove-assignment="teachers" aria-label="Hapus petugas ${index + 1}" title="Hapus petugas" ${assignments.length === 1 ? 'disabled' : ''}>×</button></div>`
    : `<div class="assignment-row student-row" data-assignment-row="students"><span class="assignment-index">${index + 1}</span><select data-person-select aria-label="Nama siswa"><option value="">${records.length ? 'Pilih siswa' : 'Unggah data siswa dahulu'}</option>${options}${extraOption(assignment)}</select><output data-class-value>${escapeHtml(assignment.class || '—')}</output><input data-rate-value type="number" min="0" step="1" value="${escapeHtml(assignment.rate ?? '')}" placeholder="Tarif" aria-label="Tarif siswa per hari"><button type="button" class="icon-button assignment-remove" data-remove-assignment="students" aria-label="Hapus siswa ${index + 1}" title="Hapus siswa" ${assignments.length === 1 ? 'disabled' : ''}>×</button></div>`).join('');
  assignments.forEach((assignment, index) => {
    if (!assignment.personId) return;
    const row = document.querySelectorAll(`[data-assignment-row="${type}"]`)[index];
    const select = row?.querySelector('[data-person-select]');
    if (select) {
      select.value = assignment.personId;
      updateAssignmentDetails(select);
    }
  });
  if (isTeacher) syncSppdPersonOptions();
}

function updateAssignmentDetails(select) {
  const row = select.closest('[data-assignment-row]');
  const option = select.selectedOptions[0];
  if (!row || !option) return;
  const rate = row.querySelector('[data-rate-value]');
  if (rate && row.dataset.assignmentRow === 'students') rate.required = Boolean(select.value);
  const rank = row.querySelector('[data-rank-value]');
  const position = row.querySelector('[data-position-value]');
  const studentClass = row.querySelector('[data-class-value]');
  if (rank) rank.textContent = [option.dataset.grade, option.dataset.rank].filter(Boolean).join(' · ') || '—';
  if (position) position.textContent = option.dataset.position || '—';
  if (studentClass) studentClass.textContent = option.dataset.class || '—';
  if (row.dataset.assignmentRow === 'teachers') syncSppdPersonOptions();
}

function collectAssignmentRows(type) {
  const isTeacher = type === 'teachers';
  return [...document.querySelectorAll(`[data-assignment-row="${type}"]`)].map((row) => {
    const select = row.querySelector('[data-person-select]');
    if (!select.value) return null;
    const option = select.selectedOptions[0];
    const rate = row.querySelector('[data-rate-value]').value;
    if (isTeacher) return { teacherId: select.value, name: option.dataset.name, nip: option.dataset.nip, nuptk: option.dataset.nuptk, grade: option.dataset.grade, rank: option.dataset.rank, position: option.dataset.position, rate };
    return { studentId: select.value, name: option.dataset.name, nis: option.dataset.nis, class: option.dataset.class, rate };
  }).filter(Boolean);
}

function currentAssignmentRows(type) {
  return [...document.querySelectorAll(`[data-assignment-row="${type}"]`)].map((row) => {
    const select = row.querySelector('[data-person-select]');
    return {
      personId: select.value,
      rate: row.querySelector('[data-rate-value]').value,
      rank: select.selectedOptions[0]?.dataset.rank || '',
      position: select.selectedOptions[0]?.dataset.position || '',
      class: select.selectedOptions[0]?.dataset.class || '',
    };
  });
}

function syncSppdPersonOptions() {
  const select = document.getElementById('sppd-person-select');
  if (!select) return;
  const currentValue = select.value;
  const teachers = collectAssignmentRows('teachers');
  select.innerHTML = '<option value="">Pilih petugas SPPD</option>' + teachers.map((teacher) => `<option value="${escapeHtml(teacher.teacherId)}">${escapeHtml(teacher.name)}</option>`).join('');
  select.value = teachers.some((teacher) => teacher.teacherId === currentValue) ? currentValue : teachers[0]?.teacherId || '';
}

function updateTripDerivedFields(changedRangeField = '') {
  const field = (name) => document.querySelector(`#form-fields [name="${name}"]`);
  const issuedDate = field('issuedDate')?.value || '';
  const startField = field('eventStartDate');
  const endField = field('eventEndDate');
  if (!startField || !endField) return;
  if (changedRangeField === 'issuedDate' && !startField.dataset.touched && !endField.dataset.touched) {
    startField.value = issuedDate;
    endField.value = issuedDate;
  }
  if (startField.value && endField.value && endField.value < startField.value) {
    if (changedRangeField === 'eventEndDate') startField.value = endField.value;
    else endField.value = startField.value;
  }
  const origin = field('origin');
  const destination = field('destination');
  const travelPurpose = field('travelPurpose');
  const departureDate = field('departureDate');
  const returnDate = field('returnDate');
  const transportDays = field('transportDays');
  if (origin) origin.value = 'SMP Negeri 13 Tasikmalaya';
  if (destination) destination.value = field('recipient')?.value || '';
  if (travelPurpose) travelPurpose.value = field('eventName')?.value || '';
  if (departureDate) departureDate.value = startField.value;
  if (returnDate) returnDate.value = endField.value;
  if (transportDays) {
    const start = startField.value ? Date.parse(`${startField.value}T00:00:00Z`) : NaN;
    const end = endField.value ? Date.parse(`${endField.value}T00:00:00Z`) : NaN;
    transportDays.value = Number.isFinite(start) && Number.isFinite(end) ? String(Math.floor((end - start) / 86400000) + 1) : '';
  }
}

function openTaskForm(taskId = '') {
  const dialog = document.getElementById('record-dialog');
  const existingTask = state.tasks.find((task) => task.id === taskId);
  const today = localDateValue();
  document.getElementById('dialog-eyebrow').textContent = existingTask ? 'KOREKSI ARSIP' : 'ARSIP BARU';
  document.getElementById('dialog-title').textContent = existingTask ? 'Edit surat tugas' : 'Buat surat tugas';
  document.getElementById('dialog-footnote').textContent = existingTask ? 'Perubahan akan memperbarui arsip ini tanpa membuat nomor arsip baru.' : 'Nomor diisi manual; arsip dan agenda diurutkan menurut tanggal surat.';
  const form = document.getElementById('record-form');
  form.dataset.kind = 'task';
  form.dataset.editId = existingTask?.id || '';
  document.getElementById('form-fields').innerHTML = `<div class="form-grid">
    <div class="form-section-title">Identitas surat</div>
    ${fieldMarkup('Nomor surat tugas', 'number', { placeholder: 'Isi nomor surat secara manual', required: true, full: true, hint: 'Nomor dapat diisi sesuai tanggal atau nomor surat terdahulu.' })}
    ${fieldMarkup('Tanggal surat', 'issuedDate', { type: 'date', value: today, required: true })}
    ${fieldMarkup('Tujuan surat', 'recipient', { placeholder: 'Nama instansi / penerima', required: true })}
    ${fieldMarkup('Dasar surat', 'basis', { placeholder: 'Contoh: Undangan dari...', required: true, full: true })}
    ${fieldMarkup('Nomor surat dasar', 'basisNumber', { placeholder: 'Nomor surat rujukan' })}
    ${fieldMarkup('Tanggal surat dasar', 'basisDate', { type: 'date' })}
    ${fieldMarkup('Perihal', 'subject', { placeholder: 'Perihal surat tugas', required: true, full: true })}
    <div class="form-section-title">Rincian penugasan</div>
    <section class="assignment-group field-full"><div class="assignment-group-heading"><div><strong>Petugas</strong><small>Nama, pangkat/golongan, dan jabatan dari master guru.</small></div><button type="button" class="button button-outline button-add-row" data-add-assignment="teachers"><span class="button-plus">+</span> Tambah petugas</button></div><div class="assignment-labels teacher-labels"><span>No</span><span>Nama</span><span>Pangkat/Gol.</span><span>Jabatan</span><span>Tarif (Rp)</span><span></span></div><div id="teacher-assignment-rows" class="assignment-rows"></div></section>
    <section class="assignment-group field-full"><div class="assignment-group-heading"><div><strong>Siswa</strong><small>Nama dan kelas dari master siswa.</small></div><button type="button" class="button button-outline button-add-row" data-add-assignment="students"><span class="button-plus">+</span> Tambah siswa</button></div><div class="assignment-labels student-labels"><span>No</span><span>Nama siswa</span><span>Kelas</span><span>Tarif (Rp)</span><span></span></div><div id="student-assignment-rows" class="assignment-rows"></div></section>
    ${fieldMarkup('Nama kegiatan', 'eventName', { placeholder: 'Nama kegiatan', full: true })}
    ${fieldMarkup('Hari', 'eventDay', { placeholder: 'Contoh: Sabtu' })}
    ${fieldMarkup('Tanggal mulai kegiatan', 'eventStartDate', { type: 'date', value: today, required: true })}
    ${fieldMarkup('Tanggal sampai kegiatan', 'eventEndDate', { type: 'date', value: today, required: true })}
    ${fieldMarkup('Waktu', 'eventTime', { placeholder: 'Contoh: 07.30 WIB' })}
    ${fieldMarkup('Tempat', 'place', { placeholder: 'Lokasi kegiatan' })}
    ${fieldMarkup('Lampiran', 'attachment', { placeholder: 'Contoh: 1 berkas', full: true })}
    <div class="form-section-title">Perjalanan & laporan</div>
    ${fieldMarkup('Tempat berangkat', 'origin', { value: 'SMP Negeri 13 Tasikmalaya', readonly: true })}
    ${fieldMarkup('Tujuan perjalanan', 'destination', { placeholder: 'Otomatis mengikuti tujuan surat', readonly: true })}
    ${fieldMarkup('Tanggal berangkat', 'departureDate', { type: 'date', value: today, readonly: true })}
    ${fieldMarkup('Tanggal kembali', 'returnDate', { type: 'date', value: today, readonly: true })}
    <label class="field"><span>Petugas untuk SPPD *</span><select name="sppdPerson" id="sppd-person-select" required><option value="">Pilih petugas SPPD</option></select></label>
    ${fieldMarkup('Lama perjalanan (hari)', 'transportDays', { type: 'number', value: '1', readonly: true })}
    ${fieldMarkup('Alat angkutan', 'transportMode', { placeholder: 'Contoh: kendaraan umum' })}
    ${fieldMarkup('Sumber anggaran', 'budgetSource', { placeholder: 'Kode / sumber pembebanan anggaran' })}
    ${fieldMarkup('Maksud perjalanan dinas', 'travelPurpose', { placeholder: 'Otomatis mengikuti nama kegiatan', full: true, readonly: true })}
    ${fieldMarkup('Catatan perjalanan (SPPD lembar 2)', 'itinerary', { type: 'textarea', placeholder: 'Tanggal | Dari | Ke | Keterangan', full: true, hint: 'Satu baris per perpindahan, dipisahkan tanda |.' })}
    ${fieldMarkup('Isi laporan Nota Dinas', 'reportText', { type: 'textarea', placeholder: 'Ringkasan hasil pelaksanaan kegiatan', full: true })}
    ${fieldMarkup('Rekomendasi / tindak lanjut', 'recommendations', { type: 'textarea', placeholder: 'Usul, masukan, atau tindak lanjut', full: true })}
    ${fieldMarkup('Kepada (Nota Dinas)', 'reportTo', { placeholder: 'Kepala Sekolah' })}
  </div>`;
  const storedAssignments = (value, type) => {
    if (!existingTask) return [{}];
    const parsed = type === 'teachers' ? parseAssignees(value) : parseStudents(value);
    const entries = Array.isArray(value) ? value : parsed.map((parts) => type === 'teachers'
      ? { name: parts[0], nip: parts[1], nuptk: parts[2], grade: parts[3], rank: parts[4], position: parts[5], rate: parts[6] }
      : { name: parts[0], nis: parts[1], class: parts[2], rate: parts[3] });
    const records = type === 'teachers' ? state.teachers : state.students;
    const idKey = type === 'teachers' ? 'teacherId' : 'studentId';
    const mappedEntries = entries.map((entry, index) => ({
      ...entry,
      personId: entry[idKey] || records.find((record) => record.name === entry.name)?.id || (entry.name ? `archived-${type}-${index}-${existingTask.id}` : ''),
    }));
    return mappedEntries.length ? mappedEntries : [{}];
  };
  renderAssignmentRows('teachers', storedAssignments(existingTask?.assignees, 'teachers'));
  renderAssignmentRows('students', storedAssignments(existingTask?.students, 'students'));
  if (existingTask) {
    for (const [name, value] of Object.entries(existingTask)) {
      const field = form.elements.namedItem(name);
      if (field && !Array.isArray(value) && (value === null || typeof value !== 'object')) field.value = String(value ?? '');
    }
  }
  const recipient = document.querySelector('#form-fields [name="recipient"]');
  const eventName = document.querySelector('#form-fields [name="eventName"]');
  const issuedDate = document.querySelector('#form-fields [name="issuedDate"]');
  const startDate = document.querySelector('#form-fields [name="eventStartDate"]');
  const endDate = document.querySelector('#form-fields [name="eventEndDate"]');
  recipient.addEventListener('input', () => updateTripDerivedFields());
  eventName.addEventListener('input', () => updateTripDerivedFields());
  issuedDate.addEventListener('change', () => updateTripDerivedFields('issuedDate'));
  startDate.addEventListener('input', () => { startDate.dataset.touched = 'true'; updateTripDerivedFields('eventStartDate'); });
  endDate.addEventListener('input', () => { endDate.dataset.touched = 'true'; updateTripDerivedFields('eventEndDate'); });
  updateTripDerivedFields();
  if (existingTask?.sppdPerson) document.getElementById('sppd-person-select').value = existingTask.sppdPerson;
  dialog.showModal();
}

function openIncomingForm() {
  const dialog = document.getElementById('record-dialog');
  document.getElementById('dialog-eyebrow').textContent = 'AGENDA MASUK';
  document.getElementById('dialog-title').textContent = 'Catat surat masuk';
  document.getElementById('dialog-footnote').textContent = 'Catatan akan masuk ke rekap agenda surat masuk.';
  document.getElementById('record-form').dataset.kind = 'incoming';
  document.getElementById('form-fields').innerHTML = `<div class="form-grid">
    <div class="form-section-title">Data surat diterima</div>
    ${fieldMarkup('Tanggal diterima', 'receivedDate', { type: 'date', value: localDateValue(), required: true })}
    ${fieldMarkup('Tanggal surat', 'letterDate', { type: 'date' })}
    ${fieldMarkup('Nomor surat', 'letterNumber', { placeholder: 'Nomor pada surat' })}
    ${fieldMarkup('Pengirim', 'sender', { placeholder: 'Instansi / pengirim', required: true })}
    ${fieldMarkup('Perihal', 'subject', { placeholder: 'Ringkasan perihal surat', required: true, full: true })}
    ${fieldMarkup('Disposisi / tindak lanjut', 'disposition', { type: 'textarea', placeholder: 'Arahan atau tindak lanjut', full: true })}
    ${fieldMarkup('Lampiran', 'attachment', { placeholder: 'Contoh: 2 berkas' })}
    <label class="field field-full"><span>Unggah berkas surat</span><input name="incomingFile" type="file" accept=".pdf,.jpg,.jpeg,.png,.doc,.docx"><small>PDF, gambar, atau dokumen Word (maksimal 30 KB)</small></label>
  </div>`;
  dialog.showModal();
}

function collectFormData(form) {
  return Object.fromEntries(new FormData(form).entries());
}

function saveTask(form) {
  updateTripDerivedFields();
  const values = collectFormData(form);
  values.assignees = collectAssignmentRows('teachers');
  values.students = collectAssignmentRows('students');
  values.eventDate = values.eventStartDate;
  const year = Number(values.issuedDate.slice(0, 4));
  const existingIndex = state.tasks.findIndex((task) => task.id === form.dataset.editId);
  const existingTask = existingIndex >= 0 ? state.tasks[existingIndex] : null;
  const task = {
    ...existingTask,
    ...values,
    id: existingTask?.id || crypto.randomUUID(),
    year,
    number: values.number.trim(),
    createdAt: existingTask?.createdAt || new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };
  if (existingIndex >= 0) state.tasks[existingIndex] = task;
  else state.tasks.push(task);
  saveState();
  document.getElementById('record-dialog').close();
  showToast(existingTask ? `Perubahan surat tugas ${task.number} berhasil disimpan.` : `Surat tugas ${task.number} tersimpan di arsip dan agenda keluar.`);
}

function saveIncoming(form) {
  const file = form.elements.incomingFile.files[0];
  if (file && file.size > 30 * 1024) { showToast('Ukuran berkas maksimal 30 KB agar dapat disimpan dengan aman di Spreadsheet.'); return; }
  const values = collectFormData(form); delete values.incomingFile;
  const commit = (attachment) => {
    const entry = { ...values, attachment: attachment || '', id: crypto.randomUUID(), createdAt: new Date().toISOString() };
    state.incoming.push(entry); saveState(); document.getElementById('record-dialog').close(); showToast('Surat masuk tersimpan di agenda.');
  };
  if (!file) return commit('');
  const reader = new FileReader(); reader.onload = () => commit(JSON.stringify({ name: file.name, type: file.type, data: reader.result })); reader.onerror = () => showToast('Berkas gagal dibaca.'); reader.readAsDataURL(file);
}

function openOutgoingForm(outgoingId = '') {
  const dialog = document.getElementById('record-dialog');
  const form = document.getElementById('record-form');
  const existing = state.outgoing.find((entry) => entry.id === outgoingId);
  document.getElementById('dialog-eyebrow').textContent = existing ? 'KOREKSI AGENDA KELUAR' : 'AGENDA KELUAR';
  document.getElementById('dialog-title').textContent = existing ? 'Edit surat keluar' : 'Catat surat keluar';
  document.getElementById('dialog-footnote').textContent = 'Catatan ini terpisah dari Surat Tugas dan masuk ke rekap agenda surat keluar.';
  form.dataset.kind = 'outgoing';
  form.dataset.editId = existing?.id || '';
  document.getElementById('form-fields').innerHTML = `<div class="form-grid">
    <div class="form-section-title">Data surat keluar</div>
    ${fieldMarkup('Tanggal keluar', 'issuedDate', { type: 'date', value: localDateValue(), required: true })}
    ${fieldMarkup('Nomor surat', 'number', { placeholder: 'Nomor surat keluar', required: true })}
    ${fieldMarkup('Jenis surat', 'letterType', { placeholder: 'Contoh: Surat undangan, surat keterangan', required: true })}
    ${fieldMarkup('Tujuan / penerima', 'recipient', { placeholder: 'Instansi atau penerima surat', required: true })}
    ${fieldMarkup('Perihal', 'subject', { placeholder: 'Perihal surat', required: true, full: true })}
    ${fieldMarkup('Keterangan', 'notes', { type: 'textarea', placeholder: 'Catatan tambahan', full: true })}
    ${fieldMarkup('Lampiran', 'attachment', { placeholder: 'Contoh: 1 berkas' })}
  </div>`;
  if (existing) {
    for (const [name, value] of Object.entries(existing)) {
      const field = form.elements.namedItem(name);
      if (field && (value === null || typeof value !== 'object')) field.value = String(value ?? '');
    }
  }
  dialog.showModal();
}

function saveOutgoing(form) {
  const values = collectFormData(form);
  const existingIndex = state.outgoing.findIndex((entry) => entry.id === form.dataset.editId);
  const existing = existingIndex >= 0 ? state.outgoing[existingIndex] : null;
  const entry = { ...existing, ...values, id: existing?.id || crypto.randomUUID(), createdAt: existing?.createdAt || new Date().toISOString(), updatedAt: new Date().toISOString() };
  if (existingIndex >= 0) state.outgoing[existingIndex] = entry;
  else state.outgoing.push(entry);
  saveState();
  document.getElementById('record-dialog').close();
  showToast(existing ? `Surat keluar ${entry.number} diperbarui.` : `Surat keluar ${entry.number} dicatat di agenda.`);
}

function csvValue(value) {
  let text = String(value ?? '');
  if (/^[=+@\-]/.test(text)) text = `'${text}`;
  return `"${text.replace(/"/g, '""')}"`;
}

function downloadFile(filename, content, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  link.click();
  setTimeout(() => URL.revokeObjectURL(link.href), 1000);
}

function downloadTask(taskId) {
  const task = state.tasks.find((entry) => entry.id === taskId);
  if (!task) return;
  const safeNumber = String(task.number || 'arsip').replace(/[^a-z0-9_-]/gi, '-');
  downloadFile(`surat-tugas-${safeNumber}.json`, JSON.stringify(task, null, 2), 'application/json;charset=utf-8');
  showToast(`Data arsip ${task.number} diunduh.`);
}

function normalizedHeader(value) {
  return String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function csvColumn(row, aliases) {
  const headers = Object.keys(row).reduce((result, header) => ({ ...result, [normalizedHeader(header)]: row[header] }), {});
  for (const alias of aliases) {
    const value = headers[normalizedHeader(alias)];
    if (value !== undefined && String(value).trim()) return String(value).trim();
  }
  return '';
}

function importMasterCsv(file, type) {
  importMasterFile(file, type);
}

async function readExcelRows(file) {
  const module = await import('exceljs');
  const ExcelJS = module.default || module;
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());
  const sheet = workbook.worksheets[0];
  if (!sheet) throw new Error('Workbook tidak memiliki lembar data.');
  const headers = sheet.getRow(1).values.slice(1).map((value) => String(value || '').trim());
  const rows = [];
  sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;
    const values = row.values.slice(1);
    rows.push(Object.fromEntries(headers.map((header, index) => [header, values[index] ?? ''])));
  });
  return rows;
}

function readCsvRows(file) {
  return new Promise((resolve, reject) => Papa.parse(file, {
    header: true,
    skipEmptyLines: 'greedy',
    complete: ({ data, errors }) => errors.length ? reject(new Error(errors[0].message)) : resolve(data),
    error: reject,
  }));
}

async function importMasterFile(file, type) {
  try {
    const rows = file.name.toLowerCase().endsWith('.xlsx') ? await readExcelRows(file) : await readCsvRows(file);
    const teachers = type === 'teachers';
    const imported = rows.map((row) => teachers ? {
      id: crypto.randomUUID(),
      name: csvColumn(row, ['nama', 'name', 'nama guru']),
      grade: csvColumn(row, ['gol', 'golongan', 'grade']),
      rank: csvColumn(row, ['pangkat', 'pangkatGol', 'pangkat/gol', 'rank']),
      position: csvColumn(row, ['jabatan', 'position']),
      nip: csvColumn(row, ['nip']),
      nuptk: csvColumn(row, ['nuptk']),
    } : {
      id: crypto.randomUUID(),
      name: csvColumn(row, ['nama siswa', 'nama', 'name']),
      class: csvColumn(row, ['kelas', 'class']),
      nis: csvColumn(row, ['nis', 'nisn']),
    }).filter((record) => record.name);
    if (!imported.length) throw new Error('Tidak ada nama yang terbaca. Periksa judul kolom file.');
    state[type] = imported;
    saveState();
    showToast(`${imported.length} ${teachers ? 'data guru' : 'data siswa'} berhasil diimpor.`);
  } catch (error) {
    showToast(`File tidak dapat dibaca: ${error.message}`);
  }
}

async function downloadRosterTemplate(type) {
  try {
    const module = await import('exceljs');
    const ExcelJS = module.default || module;
    const isTeacher = type === 'teachers';
    const headers = isTeacher ? ['No', 'Nama', 'NIP', 'Gol', 'PANGKAT', 'JABATAN'] : ['No', 'Nama Siswa', 'Kelas'];
    const widths = isTeacher ? [8, 36, 23, 11, 27, 24] : [8, 36, 16];
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Administrasi Persuratan SMPN 13 Tasikmalaya';
    const worksheet = workbook.addWorksheet(isTeacher ? 'Data Guru' : 'Data Siswa');
    worksheet.addTable({
      name: isTeacher ? 'DataGuru' : 'DataSiswa',
      ref: 'A1',
      headerRow: true,
      style: { theme: 'TableStyleMedium2', showRowStripes: true },
      columns: headers.map((name) => ({ name })),
      rows: Array.from({ length: 50 }, (_, index) => [index + 1, ...Array(headers.length - 1).fill('')]),
    });
    worksheet.views = [{ state: 'frozen', ySplit: 1, activeCell: 'A2' }];
    worksheet.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 9 };
    widths.forEach((width, index) => { worksheet.getColumn(index + 1).width = width; });
    if (isTeacher) worksheet.getColumn(3).numFmt = '@';
    worksheet.getRow(1).height = 24;
    worksheet.getRow(1).eachCell((cell) => {
      cell.font = { bold: true, color: { argb: 'FF173F35' }, size: 11 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCEAF5' } };
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      cell.border = { top: { style: 'thin', color: { argb: 'FF26352F' } }, left: { style: 'thin', color: { argb: 'FF26352F' } }, bottom: { style: 'thin', color: { argb: 'FF26352F' } }, right: { style: 'thin', color: { argb: 'FF26352F' } } };
    });
    const buffer = await workbook.xlsx.writeBuffer();
    downloadFile(`format-data-${isTeacher ? 'guru' : 'siswa'}.xlsx`, buffer, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  } catch (error) {
    showToast(`Template Excel gagal dibuat: ${error.message}`);
  }
}

const SPREADSHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const OFFICE_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

async function readXmlPart(zip, path) {
  const file = zip.file(path);
  if (!file) throw new Error(`Bagian workbook ${path} tidak ditemukan.`);
  const xml = new DOMParser().parseFromString(await file.async('string'), 'application/xml');
  if (xml.getElementsByTagName('parsererror').length) throw new Error(`XML workbook ${path} tidak dapat dibaca.`);
  return xml;
}

function writeXmlPart(zip, path, xml) {
  zip.file(path, new XMLSerializer().serializeToString(xml));
}

function worksheetCell(sheetXml, address) {
  const existing = [...sheetXml.getElementsByTagNameNS(SPREADSHEET_NS, 'c')].find((cell) => cell.getAttribute('r') === address);
  if (existing) return existing;
  const match = address.match(/^([A-Z]+)(\d+)$/);
  if (!match) throw new Error(`Alamat sel ${address} tidak valid.`);
  const column = [...match[1]].reduce((value, letter) => value * 26 + letter.charCodeAt(0) - 64, 0);
  const rowNumber = Number(match[2]);
  const sheetData = sheetXml.getElementsByTagNameNS(SPREADSHEET_NS, 'sheetData')[0];
  let row = [...sheetData.getElementsByTagNameNS(SPREADSHEET_NS, 'row')].find((item) => Number(item.getAttribute('r')) === rowNumber);
  if (!row) {
    row = sheetXml.createElementNS(SPREADSHEET_NS, 'row');
    row.setAttribute('r', String(rowNumber));
    const nextRow = [...sheetData.getElementsByTagNameNS(SPREADSHEET_NS, 'row')].find((item) => Number(item.getAttribute('r')) > rowNumber);
    sheetData.insertBefore(row, nextRow || null);
  }
  const cell = sheetXml.createElementNS(SPREADSHEET_NS, 'c');
  cell.setAttribute('r', address);
  const nextCell = [...row.getElementsByTagNameNS(SPREADSHEET_NS, 'c')].find((item) => {
    const ref = item.getAttribute('r').match(/^([A-Z]+)/)?.[1] || '';
    const index = [...ref].reduce((value, letter) => value * 26 + letter.charCodeAt(0) - 64, 0);
    return index > column;
  });
  row.insertBefore(cell, nextCell || null);
  return cell;
}

function clearCellContent(cell) {
  [...cell.childNodes].forEach((node) => {
    if (node.nodeType === 1 && ['f', 'v', 'is'].includes(node.localName)) cell.removeChild(node);
  });
  cell.removeAttribute('t');
}

function setWorkbookCell(sheetXml, address, value, type = 'text') {
  const cell = worksheetCell(sheetXml, address);
  clearCellContent(cell);
  if (value === null || value === undefined || value === '') return;
  if (type === 'number') {
    const numeric = sheetXml.createElementNS(SPREADSHEET_NS, 'v');
    numeric.textContent = String(value);
    cell.appendChild(numeric);
    return;
  }
  cell.setAttribute('t', 'inlineStr');
  const inline = sheetXml.createElementNS(SPREADSHEET_NS, 'is');
  const text = sheetXml.createElementNS(SPREADSHEET_NS, 't');
  text.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
  text.textContent = String(value);
  inline.appendChild(text);
  cell.appendChild(inline);
}

function setWorkbookFormula(sheetXml, address, formula) {
  const cell = worksheetCell(sheetXml, address);
  [...cell.childNodes].forEach((node) => {
    if (node.nodeType === 1 && ['f', 'v', 'is'].includes(node.localName)) cell.removeChild(node);
  });
  cell.setAttribute('t', 'str');
  const formulaNode = sheetXml.createElementNS(SPREADSHEET_NS, 'f');
  formulaNode.textContent = formula;
  cell.appendChild(formulaNode);
}

function excelSerial(dateValue) {
  const [year, month, day] = String(dateValue).split('-').map(Number);
  const timestamp = Date.UTC(year, month - 1, day);
  return Math.floor((timestamp - Date.UTC(1899, 11, 30)) / 86400000);
}

function formatWorkbookDate(dateValue) {
  if (!dateValue) return '';
  const [year, month, day] = String(dateValue).split('-').map(Number);
  return new Intl.DateTimeFormat('id-ID', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'UTC' }).format(new Date(Date.UTC(year, month - 1, day)));
}

function taskNumberSequence(number) {
  const value = String(number || '').trim();
  const fullNumber = value.match(/400(?:\.\d+)+\/(\d+)\/SMPN\.13\/\d{4}/i);
  if (fullNumber) return fullNumber[1];
  if (/^\d+$/.test(value)) return value;
  throw new Error('Nomor Surat Tugas harus berupa nomor urut atau nomor lengkap 400.3.5/{nomor}/SMPN.13/{tahun}.');
}

function workbookRelationshipPaths(workbookXml, relationshipsXml) {
  const relationships = new Map([...relationshipsXml.getElementsByTagNameNS(PACKAGE_REL_NS, 'Relationship')].map((item) => [item.getAttribute('Id'), item.getAttribute('Target')]));
  return [...workbookXml.getElementsByTagNameNS(SPREADSHEET_NS, 'sheet')].map((sheet, index) => {
    const target = relationships.get(sheet.getAttributeNS(OFFICE_REL_NS, 'id'));
    const path = target?.startsWith('/') ? target.slice(1) : `xl/${target}`;
    return { name: sheet.getAttribute('name'), index, path };
  });
}

function asTeacherRecord(entry, index) {
  if (Array.isArray(entry)) return { id: entry.teacherId || `teacher-${index}`, name: entry.name || entry[0] || '', nip: entry.nip || entry[1] || '', nuptk: entry.nuptk || '', grade: entry.grade || '', rank: entry.rank || entry[2] || '', position: entry.position || entry[3] || '', rate: entry.rate || entry[4] || '' };
  return { ...entry, id: entry.teacherId || entry.id || `teacher-${index}` };
}

function asStudentRecord(entry, index) {
  if (Array.isArray(entry)) return { id: entry.studentId || `student-${index}`, name: entry.name || entry[0] || '', nis: entry.nis || entry[1] || '', class: entry.class || entry[2] || '', rate: entry.rate || entry[3] || '' };
  return { ...entry, id: entry.studentId || entry.id || `student-${index}` };
}

async function downloadOriginalWorkbook(taskId, printType) {
  try {
    const task = state.tasks.find((entry) => entry.id === taskId);
    if (!task) throw new Error('Arsip surat tugas tidak ditemukan.');
    const teachers = state.teachers.map(asTeacherRecord);
    const students = state.students.map(asStudentRecord);
    const assignedTeachers = (Array.isArray(task.assignees) ? task.assignees : parseAssignees(task.assignees).map((person) => asTeacherRecord(person))).map(asTeacherRecord);
    const assignedStudents = (Array.isArray(task.students) ? task.students : parseStudents(task.students).map((person) => asStudentRecord(person))).map(asStudentRecord);
    if (!assignedTeachers.length) throw new Error('Pilih minimal satu petugas sebelum menyiapkan dokumen cetak.');
    if (assignedTeachers.length > 15) throw new Error('Format workbook memuat maksimal 15 petugas per surat tugas.');

    const ensureRosterEntry = (roster, record) => {
      let index = roster.findIndex((item) => item.id === record.id);
      if (index < 0) {
        index = roster.length;
        roster.push({ ...record, id: record.id || `${roster === teachers ? 'teacher' : 'student'}-${index}` });
      }
      return index + 1;
    };
    const teacherIds = assignedTeachers.map((teacher) => ensureRosterEntry(teachers, teacher));
    const studentIds = assignedStudents.map((student) => ensureRosterEntry(students, student));
    if (teachers.length > 64) throw new Error('Workbook asli menyediakan 64 baris master guru.');
    if (students.length > 1034) throw new Error('Workbook asli menyediakan 1.034 baris master siswa.');
    if (assignedStudents.length > 1034) throw new Error('Workbook asli menyediakan maksimal 1.034 siswa per kegiatan.');

    const sequence = taskNumberSequence(task.number);
    const issuedDate = task.issuedDate || localDateValue();
    const startDate = task.eventStartDate || task.eventDate || issuedDate;
    const endDate = task.eventEndDate || task.eventDate || startDate;
    if (endDate < startDate) throw new Error('Tanggal sampai kegiatan tidak boleh sebelum tanggal mulai.');
    const days = Math.floor((excelSerial(endDate) - excelSerial(startDate)) / 1) + 1;
    const zipModule = await import('jszip');
    const JSZip = zipModule.default || zipModule;
    const response = await fetch(printWorkbookUrl);
    if (!response.ok) throw new Error('Template workbook asli tidak dapat dimuat.');
    const zip = await JSZip.loadAsync(await response.arrayBuffer());
    const workbookXml = await readXmlPart(zip, 'xl/workbook.xml');
    const relationshipsXml = await readXmlPart(zip, 'xl/_rels/workbook.xml.rels');
    const sheetPaths = workbookRelationshipPaths(workbookXml, relationshipsXml);
    const pathFor = (name) => sheetPaths.find((sheet) => sheet.name === name)?.path;
    const inputPath = pathFor('Isian');
    const inputXml = await readXmlPart(zip, inputPath);

    setWorkbookCell(inputXml, 'B2', sequence, 'number');
    setWorkbookCell(inputXml, 'B3', task.basis || '');
    setWorkbookCell(inputXml, 'B4', task.basisNumber || '');
    setWorkbookCell(inputXml, 'B5', formatWorkbookDate(issuedDate));
    setWorkbookCell(inputXml, 'B6', task.subject || '');
    setWorkbookCell(inputXml, 'B23', task.eventName || '');
    setWorkbookCell(inputXml, 'B25', excelSerial(startDate), 'number');
    setWorkbookCell(inputXml, 'E25', endDate === startDate ? '' : excelSerial(endDate), 'number');
    setWorkbookCell(inputXml, 'H25', days, 'number');
    setWorkbookCell(inputXml, 'B26', task.eventTime || '');
    setWorkbookCell(inputXml, 'B27', task.recipient || task.destination || '');

    for (let row = 8; row <= 22; row += 1) {
      setWorkbookCell(inputXml, `B${row}`, '');
      setWorkbookCell(inputXml, `G${row}`, '');
    }
    for (let row = 31; row <= 1081; row += 1) {
      setWorkbookCell(inputXml, `B${row}`, '');
      setWorkbookCell(inputXml, `G${row}`, '');
    }
    assignedTeachers.forEach((teacher, index) => {
      const row = 8 + index;
      setWorkbookCell(inputXml, `B${row}`, teacherIds[index], 'number');
      setWorkbookCell(inputXml, `G${row}`, teacher.rate || 0, 'number');
    });
    assignedStudents.forEach((student, index) => {
      const row = 31 + index;
      setWorkbookCell(inputXml, `B${row}`, studentIds[index], 'number');
      setWorkbookCell(inputXml, `G${row}`, student.rate || 0, 'number');
    });
    teachers.forEach((teacher, index) => {
      const row = index + 3;
      for (const [column, value] of [['Q', index + 1], ['R', teacher.name], ['S', teacher.nip], ['T', teacher.nuptk], ['U', teacher.grade], ['V', teacher.rank], ['W', teacher.position], ['X', teacher.subject]]) {
        setWorkbookCell(inputXml, `${column}${row}`, value, column === 'Q' ? 'number' : 'text');
      }
    });
    for (let row = teachers.length + 3; row <= 66; row += 1) {
      for (const column of ['Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X']) setWorkbookCell(inputXml, `${column}${row}`, '');
    }
    students.forEach((student, index) => {
      const row = index + 3;
      for (const [column, value] of [['Z', index + 1], ['AA', student.schoolId], ['AB', student.nis], ['AC', student.name], ['AD', student.gender], ['AE', student.class]]) {
        setWorkbookCell(inputXml, `${column}${row}`, value, column === 'Z' ? 'number' : 'text');
      }
    });
    for (let row = students.length + 3; row <= 1036; row += 1) {
      for (const column of ['Z', 'AA', 'AB', 'AC', 'AD', 'AE']) setWorkbookCell(inputXml, `${column}${row}`, '');
    }
    writeXmlPart(zip, inputPath, inputXml);

    const yearFormula = 'YEAR(Isian!B25)';
    const yearFormulas = {
      'Surat Tugas': ['A7', `"Nomor : 400.3.5/"&Isian!B2&"/SMPN.13/"&${yearFormula}`],
      'Surat Tugas Siswa': ['A8', `"Nomor : 400.3.5/"&Isian!B2&"/SMPN.13/"&${yearFormula}`],
      'SPPD Print': ['D8', `" 400.3.5/"&Isian!B2&"/SMPN.13/"&${yearFormula}`],
      'SPPD Lembar ke 2': ['B2', `"SPPD No.                         : 400.3.5/"&Isian!B2&"/SMPN.13/"&${yearFormula}`],
    };
    const activeSheets = { task: 'Surat Tugas', student: 'Surat Tugas Siswa', sppd: 'SPPD Print', 'sppd-continuation': 'SPPD Lembar ke 2', nota: 'Nota Dinas', transport: 'Daftar Penyerahan', 'transport-students': 'Daftar Penyerahan Siswa' };
    for (const [sheetName, [cell, formula]] of Object.entries(yearFormulas)) {
      const path = pathFor(sheetName);
      const sheetXml = await readXmlPart(zip, path);
      setWorkbookFormula(sheetXml, cell, formula);
      if (sheetName === 'SPPD Print') {
        const selectedId = String(task.sppdPerson || '');
        const selectedIndex = assignedTeachers.findIndex((teacher) => teacher.teacherId === selectedId || teacher.id === selectedId);
        setWorkbookCell(sheetXml, 'H17', Math.max(1, selectedIndex + 1), 'number');
      }
      writeXmlPart(zip, path, sheetXml);
    }
    const studentTaskPath = pathFor('Surat Tugas Siswa');
    const studentTaskXml = await readXmlPart(zip, studentTaskPath);
    for (const [row, sourceRow] of [[20, 31], [21, 32], [22, 33]]) {
      setWorkbookFormula(studentTaskXml, `B${row}`, `_xlfn.IFNA(IF(ISBLANK(A${row}),"",Isian!C${sourceRow}),"")`);
    }
    writeXmlPart(zip, studentTaskPath, studentTaskXml);
    const selectedSheet = activeSheets[printType];
    const selectedSheetInfo = sheetPaths.find((sheet) => sheet.name === selectedSheet);
    if (!selectedSheetInfo) throw new Error(`Sheet cetak ${selectedSheet} tidak ditemukan.`);
    const workbookView = workbookXml.getElementsByTagNameNS(SPREADSHEET_NS, 'workbookView')[0];
    if (workbookView) workbookView.setAttribute('activeTab', String(selectedSheetInfo.index));
    const inputViewXml = await readXmlPart(zip, inputPath);
    [...inputViewXml.getElementsByTagNameNS(SPREADSHEET_NS, 'sheetView')].forEach((view) => view.removeAttribute('tabSelected'));
    writeXmlPart(zip, inputPath, inputViewXml);
    const activeViewXml = await readXmlPart(zip, selectedSheetInfo.path);
    const activeView = activeViewXml.getElementsByTagNameNS(SPREADSHEET_NS, 'sheetView')[0];
    if (activeView) activeView.setAttribute('tabSelected', '1');
    writeXmlPart(zip, selectedSheetInfo.path, activeViewXml);
    const calculation = workbookXml.getElementsByTagNameNS(SPREADSHEET_NS, 'calcPr')[0];
    if (calculation) {
      calculation.setAttribute('calcMode', 'auto');
      calculation.setAttribute('fullCalcOnLoad', '1');
      calculation.setAttribute('forceFullCalc', '1');
    }
    writeXmlPart(zip, 'xl/workbook.xml', workbookXml);

    const output = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });
    const safeNumber = task.number.replace(/[<>:"/\\|?*]/g, '-');
    downloadFile(`surat-tugas-${safeNumber}-format-asli.xlsm`, output, 'application/vnd.ms-excel.sheet.macroEnabled.12');
    document.getElementById('print-dialog').close();
    showToast(`Workbook format asli diunduh. Buka di Excel, lalu cetak sheet ${selectedSheet}.`);
  } catch (error) {
    showToast(error.message || 'Workbook untuk cetak gagal disiapkan.');
  }
}

function exportAgenda() {
  const isIncoming = activeAgendaTab === 'incoming';
  const headers = isIncoming
    ? ['Tanggal diterima', 'Nomor surat', 'Tanggal surat', 'Pengirim', 'Perihal', 'Disposisi', 'Lampiran']
    : ['Tanggal keluar', 'Nomor surat', 'Jenis surat', 'Tujuan', 'Perihal', 'Keterangan / kegiatan', 'Lampiran'];
  const outgoingRecords = [
    ...state.tasks.map((task) => [task.issuedDate, task.number, 'Surat Tugas', task.recipient, task.subject, task.eventName, task.attachment]),
    ...state.outgoing.map((entry) => [entry.issuedDate, entry.number, entry.letterType, entry.recipient, entry.subject, entry.notes, entry.attachment]),
  ];
  const records = isIncoming ? state.incoming.map((entry) => [entry.receivedDate, entry.letterNumber, entry.letterDate, entry.sender, entry.subject, entry.disposition, entry.attachment]) : outgoingRecords;
  const csv = [headers, ...records].map((row) => row.map(csvValue).join(',')).join('\r\n');
  downloadFile(`agenda-${isIncoming ? 'masuk' : 'keluar'}-${CURRENT_YEAR}.csv`, `\ufeff${csv}`, 'text/csv;charset=utf-8');
  showToast(`Agenda surat ${isIncoming ? 'masuk' : 'keluar'} diekspor.`);
}

function exportBackup() {
  const backup = { app: 'Administrasi Persuratan SMPN 13 Tasikmalaya', version: 1, exportedAt: new Date().toISOString(), ...state };
  downloadFile(`cadangan-agenda-surat-${localDateValue()}.json`, JSON.stringify(backup, null, 2), 'application/json');
  showToast('Cadangan data berhasil diunduh.');
}

async function importBackupLegacy(file) {
  try {
    const imported = JSON.parse(await file.text());
    if (!imported || !Array.isArray(imported.tasks) || !Array.isArray(imported.incoming)) throw new Error('Format cadangan tidak sesuai.');
    const taskIds = new Set(state.tasks.map((task) => task.id));
    const incomingIds = new Set(state.incoming.map((entry) => entry.id));
    const outgoingIds = new Set(state.outgoing.map((entry) => entry.id));
    state.tasks.push(...imported.tasks.filter((task) => task.id && !taskIds.has(task.id)));
    state.incoming.push(...imported.incoming.filter((entry) => entry.id && !incomingIds.has(entry.id)));
    state.outgoing.push(...(Array.isArray(imported.outgoing) ? imported.outgoing.filter((entry) => entry.id && !outgoingIds.has(entry.id)) : []));
    if (Array.isArray(imported.teachers)) state.teachers = imported.teachers;
    if (Array.isArray(imported.students)) state.students = imported.students;
    if (imported.officials && typeof imported.officials === 'object') state.officials = { ...defaultState.officials, ...imported.officials };
    if (imported.officialHistory && typeof imported.officialHistory === 'object') {
      state.officialHistory = {
        principal: Array.isArray(imported.officialHistory.principal) ? imported.officialHistory.principal : state.officialHistory.principal,
        treasurer: Array.isArray(imported.officialHistory.treasurer) ? imported.officialHistory.treasurer : state.officialHistory.treasurer,
      };
    }
    state.lastIssuedByYear = { ...state.lastIssuedByYear, ...(imported.lastIssuedByYear || {}) };
    saveState();
    showToast('Cadangan berhasil digabungkan ke data lokal.');
  } catch (error) {
    showToast(error.message || 'Cadangan tidak dapat dibaca.');
  }
}

function openPrintOptions(taskId) {
  if (!state.tasks.some((task) => task.id === taskId)) return;
  activePrintTaskId = taskId;
  const printOptions = document.getElementById('print-options');
  if (!printOptions.dataset.defaultHtml) printOptions.dataset.defaultHtml = printOptions.innerHTML;
  printOptions.innerHTML = printOptions.dataset.defaultHtml;
  document.getElementById('print-dialog-title').textContent = 'Pilih lembar';
  const dialog = document.getElementById('print-dialog');
  if (!dialog.open) dialog.showModal();
}

function openSppdPersonOptions(type) {
  const task = state.tasks.find((entry) => entry.id === activePrintTaskId);
  if (!task) return;
  const assignees = Array.isArray(task.assignees) ? task.assignees : [];
  if (!assignees.length) {
    showToast('Tambahkan petugas terlebih dahulu sebelum mencetak SPPD.');
    return;
  }
  activeSppdPrintType = type;
  const options = document.getElementById('print-options');
  if (!options.dataset.defaultHtml) options.dataset.defaultHtml = options.innerHTML;
  options.innerHTML = `${assignees.map((person, index) => `<button class="print-option" data-sppd-person="${escapeHtml(person.teacherId || '')}"><span class="print-option-number">${String(index + 1).padStart(2, '0')}</span><span><strong>${escapeHtml(person.name || '')}</strong><small>${escapeHtml(person.nip || '')}</small></span><span class="print-option-arrow">↗</span></button>`).join('')}<button class="print-option print-back-option" data-back-print-options><span class="print-option-number">←</span><span><strong>Kembali</strong><small>Pilih lembar cetak lain</small></span></button>`;
  document.getElementById('print-dialog-title').textContent = type === 'sppd' ? 'Pilih petugas SPPD' : 'Pilih petugas lembar ke-2';
}

function money(value) {
  return new Intl.NumberFormat('id-ID').format(Number(value) || 0);
}

function printSignature(date) {
  const official = state.officials;
  return `<div class="sign"><p>Dikeluarkan di&nbsp;&nbsp; : Tasikmalaya</p><p>Pada Tanggal&nbsp;&nbsp;&nbsp;&nbsp; : ${escapeHtml(printLongDate(date))}</p><p>Kepala Sekolah</p><div class="signature"></div><p><strong>${escapeHtml(official.principalName)}</strong></p><p>NIP.&nbsp; ${escapeHtml(official.principalNip)}</p></div>`;
}

function printDocumentLegacy(taskId, type) {
  const task = state.tasks.find((entry) => entry.id === taskId);
  if (!task) return;
  const people = parseAssignees(task.assignees);
  const students = parseStudents(task.students);
  const days = Math.max(1, Number(task.transportDays) || 1);
  const rateFor = (entry, index, defaultRate) => entry[index] !== undefined && entry[index] !== '' ? Number(entry[index]) || 0 : Number(defaultRate) || 0;
  const employeeRows = people.map((person, index) => {
    const extended = person.length >= 8;
    const rankAndGrade = extended ? [person[4], person[3]].filter(Boolean).join(' / ') : person[2] || '';
    return `<tr><td>${index + 1}</td><td>${escapeHtml(person[0] || '')}</td><td>${escapeHtml(person[1] || '')}</td><td>${escapeHtml(rankAndGrade)}</td><td>${escapeHtml(extended ? person[5] || '' : person[3] || '')}</td></tr>`;
  }).join('');
  const studentRows = students.map((student, index) => `<tr><td>${index + 1}</td><td>${escapeHtml(student[0] || '')}</td><td>${escapeHtml(student[1] || '')}</td><td>${escapeHtml(student[2] || '')}</td></tr>`).join('');
  const itinerary = parseAssignees(task.itinerary);
  const requestedSppdPerson = String(task.sppdPerson || '').trim();
  const selectedTeacherName = state.teachers.find((teacher) => teacher.id === requestedSppdPerson)?.name || requestedSppdPerson;
  const selectedSppdPerson = people.find((person) => person[0].toLowerCase() === selectedTeacherName.toLowerCase()) || (requestedSppdPerson.includes('|') ? parseAssignees(requestedSppdPerson)[0] : null) || people[0] || [];
  const extendedSppdPerson = selectedSppdPerson.length >= 8;
  const routeRows = (itinerary.length ? itinerary : Array.from({ length: 5 }, () => [])).map((route, index) => `<tr><td>${index + 1}</td><td>${escapeHtml(route[0] || '')}</td><td>${escapeHtml(route[1] || task.origin || '')}</td><td>${escapeHtml(route[2] || task.destination || task.place || '')}</td><td>${escapeHtml(route[3] || '')}</td></tr>`).join('');
  const header = '<header class="letterhead"><strong>PEMERINTAH KOTA TASIKMALAYA</strong><strong>SMP NEGERI 13 TASIKMALAYA</strong><p>Jl. Letjen H. Ibrahim Adjie Km.2 Indihiang Tasikmalaya 46151 Telp. (0265) 335695</p><p>e-mail: smp13tasikmalaya@yahoo.co.id</p></header>';
  const meta = `<table class="meta"><tr><td>Dasar Surat</td><td>:</td><td>${escapeHtml(task.basis || '')}</td></tr><tr><td>Nomor</td><td>:</td><td>${escapeHtml(task.basisNumber || '')}</td></tr><tr><td>Tanggal</td><td>:</td><td>${escapeHtml(formatDate(task.basisDate))}</td></tr><tr><td>Perihal</td><td>:</td><td>${escapeHtml(task.subject || '')}</td></tr></table>`;
  const eventStart = task.eventStartDate || task.eventDate;
  const eventEnd = task.eventEndDate || task.eventDate;
  const eventDateText = eventStart && eventEnd && eventStart !== eventEnd ? `${formatDate(eventStart)} s.d. ${formatDate(eventEnd)}` : formatDate(eventStart || eventEnd);
  const eventDetail = `<p>Untuk melaksanakan kegiatan <strong>${escapeHtml(task.eventName || '—')}</strong>${task.eventDay || eventStart ? ` pada ${escapeHtml(task.eventDay || '')} ${escapeHtml(eventDateText)}` : ''}${task.eventTime ? `, pukul ${escapeHtml(task.eventTime)}` : ''}${task.place ? ` bertempat di ${escapeHtml(task.place)}` : ''}.</p>`;
  const employeeTable = `<table class="grid"><thead><tr><th>No</th><th>Nama</th><th>NIP</th><th>Pangkat / Gol.</th><th>Jabatan</th></tr></thead><tbody>${employeeRows}</tbody></table>`;
  const studentTable = `<table class="grid"><thead><tr><th>No</th><th>Nama</th><th>NIS</th><th>Kelas</th></tr></thead><tbody>${studentRows}</tbody></table>`;
  const textBlock = (text) => escapeHtml(text || '').replace(/\r?\n/g, '<br>');
  let title = '';
  let content = '';
  let landscape = false;

  if (type === 'task') {
    title = 'SURAT TUGAS';
    content = `<p class="number">Nomor: ${escapeHtml(task.number)}</p>${meta}<p>Kepala SMP Negeri 13 Tasikmalaya menugaskan:</p>${employeeTable}${eventDetail}<p>Demikian surat tugas ini dibuat untuk dilaksanakan dengan penuh tanggung jawab.</p>${task.attachment ? `<p>Lampiran: ${escapeHtml(task.attachment)}</p>` : ''}${printSignature(task.issuedDate)}`;
  } else if (type === 'student') {
    title = 'SURAT TUGAS';
    content = `<p class="subtitle">PENUGASAN PESERTA DIDIK</p><p class="number">Nomor: ${escapeHtml(task.number)}</p>${meta}<p>Kepala SMP Negeri 13 Tasikmalaya menugaskan peserta didik berikut:</p>${studentTable}${eventDetail}<p>Demikian surat tugas ini dibuat untuk dilaksanakan dengan penuh tanggung jawab.</p>${printSignature(task.issuedDate)}`;
  } else if (type === 'sppd') {
    title = 'S U R A T&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; P E R I N T A H&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; P E R J A L A N A N&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; D I N A S';
    body = `<div class="sppd-identifiers"><p>Lembaran Ke&nbsp;&nbsp;&nbsp; : I</p><p>Kode No&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : __________________</p><p>SPPD No&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : 400.3.5/<span class="number-slot">${escapeHtml(sequence)}</span>/SMPN.13/${escapeHtml(year)}</p></div><h1 class="sppd-title">SURAT PERINTAH PERJALANAN DINAS<br>(SPPD)</h1><table class="sppd-table"><tr><td>1. Pengguna Anggaran / Kuasa Pengguna Anggaran</td><td>Kepala Sekolah</td></tr><tr><td>2. Nama Pegawai yang Diperintah / NIP</td><td>${escapeHtml(selectedPerson[0] || '')}<br>${escapeHtml(selectedPerson[1] || '')}</td></tr><tr><td>3. a. Pangkat dan Golongan</td><td>${escapeHtml([selectedPerson[4], selectedPerson[3]].filter(Boolean).join(' / '))}</td></tr><tr><td>&nbsp;&nbsp;&nbsp; b. Jabatan / OPD</td><td>${escapeHtml(selectedPerson[5] || '')}</td></tr><tr><td>&nbsp;&nbsp;&nbsp; c. Tingkat Biaya Perjalanan Dinas</td><td></td></tr><tr><td>4. Maksud Perjalanan Dinas</td><td>${escapeHtml(task.eventName || '')}</td></tr><tr><td>5. Alat Angkut yang Dipergunakan</td><td>Kendaraan</td></tr><tr><td>6. a. Tempat Berangkat</td><td>SMP Negeri 13 Tasikmalaya</td></tr><tr><td>&nbsp;&nbsp;&nbsp; b. Tempat Tujuan</td><td>${escapeHtml(task.recipient || '')}</td></tr><tr><td>7. a. Lamanya Perjalanan Dinas</td><td>${days} hari</td></tr><tr><td>&nbsp;&nbsp;&nbsp; b. Tanggal Berangkat</td><td>${escapeHtml(printLongDate(start))}</td></tr><tr><td>&nbsp;&nbsp;&nbsp; c. Tanggal Harus Kembali / Tiba di Tempat Baru</td><td>${escapeHtml(printLongDate(end))}</td></tr><tr><td>8. Pengikut : Nama</td><td></td></tr><tr><td>9. Pembebanan Anggaran</td><td></td></tr><tr><td>&nbsp;&nbsp;&nbsp; a. Instansi</td><td>-</td></tr><tr><td>&nbsp;&nbsp;&nbsp; b. Mata Anggaran</td><td>-</td></tr><tr><td>10. Keterangan Lain-lain</td><td>${escapeHtml(task.subject || '')}</td></tr></table><div class="sppd-signature-grid"><div class="sppd-sign-box"><p>Dikeluarkan di&nbsp;&nbsp; : Tasikmalaya</p><p>Pada Tanggal&nbsp; : ${escapeHtml(printLongDate(task.issuedDate))}</p><p class="sppd-sign-label">Pengguna Anggaran / Kuasa Pengguna Anggaran</p><div class="signature-space"></div><p><strong>${escapeHtml(principal.name || '')}</strong></p><p>NIP. ${escapeHtml(principal.nip || '')}</p></div></div>`;
  } else if (type === 'sppd-continuation') {
    const dateBlocks = Array.from({ length: days }, (_, index) => {
      const date = new Date(`${start}T00:00:00`);
      date.setDate(date.getDate() + index);
      const dayDate = localDateValue(date);
      return `<table class="travel-day"><tr><td colspan="2"><strong>Hari ke-${index + 1} : ${escapeHtml(printLongDate(dayDate))}</strong></td></tr><tr><td>I. Tiba di : ${escapeHtml(task.recipient || '')}</td><td>Berangkat dari : ${escapeHtml(task.recipient || '')}</td></tr><tr><td>Pada Tanggal : .............................................</td><td>Ke &nbsp;: SMP Negeri 13 Tasikmalaya</td></tr><tr><td>(${escapeHtml(printLongDate(dayDate))})</td><td>Pada Tanggal : .............................................</td></tr><tr><td>Mengetahui pejabat setempat</td><td>Tanda tangan pejabat</td></tr><tr><td>(................................................)</td><td>(................................................)</td></tr></table>`;
    }).join('');
    body = `<div class="sppd-identifiers"><p>Lembaran Ke&nbsp;&nbsp;&nbsp; : II</p><p>Kode No&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : __________________</p><p>SPPD No. &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : ${escapeHtml(officialNumber)}</p><p>Petugas&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : ${escapeHtml(selectedPerson[0] || '')}</p></div><p class="sppd-route-line">Berangkat dari &nbsp;: SMP Negeri 13 Tasikmalaya</p><p class="sppd-route-line">(Tempat Kedudukan)</p><p class="sppd-route-line">Ke &nbsp;: ${escapeHtml(task.recipient || '')}</p><p class="sppd-route-line"><strong>MENGETAHUI PEJABAT SETEMPAT :</strong></p>${dateBlocks}`;
    body += `<table class="return-review"><tr><td><strong>II. Tiba kembali di : SMP Negeri 13 Tasikmalaya</strong><p>Pada Tanggal : .............................................</p><p>(${escapeHtml(printLongDate(end))})</p></td><td><strong>Telah diperiksa, dengan keterangan :</strong><p>Bahwa perjalanan dinas tersebut di atas benar dilakukan atas perintahnya dan semata-mata untuk kepentingan jabatan dalam waktu yang sesingkat-singkatnya.</p><p>Kepala Sekolah</p><div class="signature-space"></div><p><strong>${escapeHtml(principal.name || '')}</strong></p><p>NIP. ${escapeHtml(principal.nip || '')}</p></td></tr></table><p class="sppd-notes-heading">III. CATATAN LAIN-LAIN :</p><div class="sppd-notes-space"></div><p class="sppd-notes-heading">IV. PERHATIAN</p><p class="sppd-note">Pejabat yang berwenang menerbitkan SPPD, pegawai yang melakukan perjalanan dinas, para pejabat yang mengesahkan tanggal berangkat/tiba serta bendahara bertanggung jawab berdasarkan peraturan-peraturan Keuangan Negara apabila Negara mendapat rugi akibat kesalahan, kealpaan.</p>`;
    orientation = 'portrait';
  } else if (type === 'nota') {
    title = 'NOTA DINAS';
    content = `<p class="subtitle">(Laporan Hasil Pelaksanaan Kegiatan)</p><table class="details"><tr><td>1. Kepada Yth.</td><td>${escapeHtml(task.reportTo || 'Kepala Sekolah')}</td></tr><tr><td>2. Dari (yang diberi tugas)</td><td>${escapeHtml(assigneeNames(task.assignees))}</td></tr><tr><td>3. Perihal / nama kegiatan</td><td>${escapeHtml(task.eventName || task.subject || '')}</td></tr><tr><td>4. Tempat kegiatan</td><td>${escapeHtml(task.place || task.destination || '')}</td></tr><tr><td>5. Tanggal pelaksanaan</td><td>${escapeHtml(eventDateText)}</td></tr><tr><td>6. Dasar penugasan</td><td>${escapeHtml(task.basis || '')}<br>Nomor: ${escapeHtml(task.number)}<br>Tanggal: ${escapeHtml(formatDate(task.issuedDate))}</td></tr></table><p>Berdasarkan surat tugas tersebut di atas, saya laporkan hasil pelaksanaan kegiatan sebagai berikut:</p><div class="report-copy">${textBlock(task.reportText)}</div><h3>Usul, masukan, rekomendasi dan tindak lanjut:</h3><div class="report-copy">${textBlock(task.recommendations)}</div>${printSignature(task.issuedDate)}`;
  } else if (type === 'transport') {
    title = 'DAFTAR PENYERAHAN TRANSPORT';
    landscape = true;
    const rows = people.map((person, index) => { const extended = person.length >= 8; const rate = rateFor(person, extended ? 6 : 4, task.employeeRate); return `<tr><td>${index + 1}</td><td>${escapeHtml(person[0] || '')}<br>NIP. ${escapeHtml(person[1] || '')}</td><td>${escapeHtml(extended ? person[3] || '' : person[2] || '')}</td><td>${days} hari × Rp ${money(rate)}</td><td>Rp ${money(days * rate)}</td><td></td></tr>`; }).join('');
    content = `<p class="subtitle">${escapeHtml((task.eventName || '').toUpperCase())}</p><p>Nomor: ${escapeHtml(task.number)}</p><table class="grid"><thead><tr><th>No</th><th>Nama / NIP</th><th>Golongan</th><th>Volume</th><th>Jumlah diterima</th><th>Tanda tangan</th></tr></thead><tbody>${rows}</tbody></table>`;
  } else if (type === 'transport-students') {
    title = 'DAFTAR PENYERAHAN TRANSPORT SISWA';
    landscape = true;
    const rows = students.map((student, index) => { const rate = rateFor(student, 3, task.studentRate); return `<tr><td>${index + 1}</td><td>${escapeHtml(student[0] || '')}</td><td>${escapeHtml(student[2] || '')}</td><td>${days} hari × Rp ${money(rate)}</td><td>Rp ${money(days * rate)}</td><td></td></tr>`; }).join('');
    content = `<p class="subtitle">${escapeHtml((task.eventName || '').toUpperCase())}</p><p>Nomor: ${escapeHtml(task.number)}</p><table class="grid"><thead><tr><th>No</th><th>Nama</th><th>Kelas</th><th>Volume</th><th>Jumlah diterima</th><th>Tanda tangan</th></tr></thead><tbody>${rows}</tbody></table>`;
  } else {
    return;
  }

  body = body.replaceAll('AGUS ROHMAN, S.Pd., M.Si.', escapeHtml(principal.name || ''))
      if (type === 'sppd') {
        title = '';
        body = `<table class="sppd-identifiers-grid"><colgroup><col><col><col><col></colgroup><tbody><tr><td rowspan="3"></td><td>Lembaran Ke</td><td>:</td><td>I</td></tr><tr><td>Kode No</td><td>:</td><td></td></tr><tr><td>SPPD No</td><td>:</td><td>400.3.5/${escapeHtml(sequence)}/SMPN.13/${escapeHtml(year)}</td></tr></tbody></table><h1 class="sppd-title">SURAT PERINTAH PERJALANAN DINAS<br>(SPPD)</h1><table class="sppd-table"><tbody><tr><td>1. Pengguna Anggaran/Kuasa Pengguna Anggaran</td><td>Kepala Sekolah</td></tr><tr><td>2. Nama Pegawai yang diperintah / NIP</td><td>${escapeHtml(selectedPerson[0] || '')}<br>${escapeHtml(selectedPerson[1] || '')}</td></tr><tr><td>3. a. Pangkat dan Golongan</td><td>${escapeHtml([selectedPerson[4], selectedPerson[3]].filter(Boolean).join(' / '))}</td></tr><tr><td>&nbsp;&nbsp;&nbsp; b. Jabatan/OPD</td><td>${escapeHtml(selectedPerson[5] || '')}</td></tr><tr><td>&nbsp;&nbsp;&nbsp; c. Tingkat Biaya Perjalanan Dinas</td><td></td></tr><tr><td>4. Maksud Perjalanan Dinas</td><td>${escapeHtml(task.eventName || '')}</td></tr><tr><td>5. Alat angkut yang digunakan</td><td>${escapeHtml(task.transportMode || 'Kendaraan')}</td></tr><tr><td>6. a. Tempat Berangkat</td><td>a. ${escapeHtml(task.origin || 'SMP Negeri 13 Tasikmalaya')}</td></tr><tr><td>&nbsp;&nbsp;&nbsp; b. Tempat Tujuan</td><td>b. ${escapeHtml(task.recipient || '')}</td></tr><tr><td>7. a. Lamanya Perjalanan Dinas</td><td>${days} hari</td></tr><tr><td>&nbsp;&nbsp;&nbsp; b. Tanggal Berangkat</td><td>${escapeHtml(printLongDate(start))}</td></tr><tr><td>&nbsp;&nbsp;&nbsp; c. Tanggal Harus Kembali/Tiba di Tempat Baru</td><td>${escapeHtml(printLongDate(end))}</td></tr><tr><td>8. Pengikut : Nama</td><td><table class="sppd-follower-table"><thead><tr><th>Tanggal Lahir</th><th>Keterangan</th></tr></thead><tbody><tr><td>&nbsp;</td><td>&nbsp;</td></tr></tbody></table></td></tr><tr><td>9. Pembebanan Anggaran</td><td></td></tr><tr><td>&nbsp;&nbsp;&nbsp; a. Instansi</td><td>a. -</td></tr><tr><td>&nbsp;&nbsp;&nbsp; b. Mata Anggaran</td><td>b. -</td></tr><tr><td>10. Keterangan Lain-lain</td><td>${escapeHtml(task.subject || '')}</td></tr></tbody></table><div class="sppd-signature-grid"><div class="sppd-sign-box"><p>Dikeluarkan di&nbsp;&nbsp; : Tasikmalaya</p><p>Pada Tanggal&nbsp; : ${escapeHtml(printLongDate(task.issuedDate))}</p><p>Pengguna Anggaran/Kuasa Pengguna Anggaran</p><div class="signature-space"></div><p><strong>${escapeHtml(principal.name || '')}</strong></p><p>NIP. ${escapeHtml(principal.nip || '')}</p></div></div>`;
      }
      if (type === 'sppd-continuation') title = '';
      body = body.replaceAll('AGUS ROHMAN, S.Pd., M.Si.', escapeHtml(principal.name || ''))
    .replaceAll('19650927 198903 1 008', escapeHtml(principal.nip || ''));
  const printWindow = window.open('', '_blank', 'width=900,height=720');
  if (!printWindow) {
    showToast('Izinkan pop-up untuk mencetak dokumen.');
    return;
  }
  document.getElementById('print-dialog').close();
  printWindow.document.write(`<!doctype html><html lang="id"><head><meta charset="utf-8"><title>${escapeHtml(title)} - ${escapeHtml(task.number)}</title><style>
    @page{size:A4 ${landscape ? 'landscape' : 'portrait'};margin:15mm}*{box-sizing:border-box}body{margin:0;color:#111;font:12px "Times New Roman",serif;line-height:1.45}.letterhead{text-align:center;border-bottom:3px double #111;padding:0 0 9px;margin-bottom:20px}.letterhead strong{display:block;font-size:16px}.letterhead p{margin:3px 0;font-size:10px}.title{text-align:center;margin:15px 0 2px;font-size:16px;font-weight:bold;text-decoration:underline}.subtitle{text-align:center;margin:2px 0 17px;font-weight:bold}.number{text-align:center;margin:0 0 18px}.meta,.details{width:100%;border-collapse:collapse;margin:9px 0 15px}.meta td,.details td{padding:4px;vertical-align:top}.meta td:first-child{width:145px}.meta td:nth-child(2){width:15px}.details td{border:1px solid #333}.details td:first-child{width:43%}.grid{width:100%;border-collapse:collapse;margin:12px 0 17px}.grid th,.grid td{border:1px solid #333;padding:6px 5px;text-align:left;vertical-align:top}.grid th{text-align:center}.route-grid td{height:54px}.sign{width:240px;margin:25px 10px 0 auto;text-align:center}.sign p{margin:3px}.signature{height:57px}.report-copy{min-height:46px;margin:8px 0 17px}.report-copy:empty::after{content:" ";white-space:pre}.report-copy{white-space:normal}h3{margin:15px 0 6px;font-size:12px}@media screen{body{max-width:${landscape ? '1050px' : '760px'};margin:32px auto;padding:0 24px}}@media print{body{max-width:none}.letterhead{break-inside:avoid}.grid tr{break-inside:avoid}}
    </style></head><body>${header}<h1 class="title">${escapeHtml(title)}</h1>${content}<script>window.onload=()=>window.print()<\/script></body></html>`);
  printWindow.document.close();
}

async function loadAssetDataUrl(path) {
  if (String(path).startsWith('data:')) return path;
  try {
    const response = await fetch(path, { cache: 'no-store' });
    if (!response.ok) return path;
    const blob = await response.blob();
    return await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve(reader.result);
      reader.onerror = () => reject(new Error('Failed to read asset'));
      reader.readAsDataURL(blob);
    });
  } catch (error) {
    return path;
  }
}

function printLongDate(value) {
  if (!value) return '';
  const dateValue = String(value).trim();
  const date = new Date(`${dateValue}T00:00:00Z`);
  if (Number.isNaN(date.getTime())) return dateValue;
  return new Intl.DateTimeFormat('id-ID', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'UTC' }).format(date);
}

function officialFor(role, dateValue) {
  const history = [...(state.officialHistory[role] || [])].sort((first, second) => first.startDate.localeCompare(second.startDate));
  return history.filter((official) => !dateValue || official.startDate <= dateValue).at(-1) || history[0] || {};
}

function saveOfficialHistory(role, name, nip, startDate) {
  const history = state.officialHistory[role];
  const existing = history.find((official) => official.startDate === startDate);
  const official = { name, nip, startDate };
  if (existing) Object.assign(existing, official);
  else history.push(official);
  history.sort((first, second) => first.startDate.localeCompare(second.startDate));
}

let activePrintWindow = null;

async function printDocument(taskId, type, sppdPersonId = '') {
  try {
    await buildPrintDocument(taskId, type, sppdPersonId);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    showToast(`Gagal menyiapkan dokumen cetak: ${message}`);
    if (activePrintWindow && !activePrintWindow.closed) {
      activePrintWindow.document.open();
      activePrintWindow.document.write(`<!doctype html><meta charset="utf-8"><title>Gagal menyiapkan cetak</title><p style="font:16px Arial;padding:24px">Dokumen gagal disiapkan.<br><br>${escapeHtml(message)}</p>`);
      activePrintWindow.document.close();
    }
    console.error('Gagal menyiapkan dokumen cetak', error);
  }
}

async function buildPrintDocument(taskId, type, sppdPersonId = '') {
  const task = state.tasks.find((entry) => entry.id === taskId);
  if (!task) return;
  // Reserve the preview window inside the user's click gesture. Browsers block
  // window.open() if it happens only after the logo-loading await below.
  const printWindow = window.open('', '_blank', 'width=1100,height=850');
  if (!printWindow) {
    showToast('Izinkan pop-up untuk membuka pratinjau cetak.');
    return;
  }
  activePrintWindow = printWindow;
  printWindow.document.write('<!doctype html><title>Menyiapkan pratinjau cetak</title><p style="font:16px Arial;padding:24px">Menyiapkan dokumen…</p>');
  printWindow.document.close();
  const people = parseAssignees(task.assignees);
  const students = parseStudents(task.students);
  const start = task.eventStartDate || task.eventDate || task.issuedDate;
  const end = task.eventEndDate || task.eventDate || start;
  const dateRange = start && end && start !== end ? `${printLongDate(start)} s.d. ${printLongDate(end)}` : printLongDate(start || end);
  const eventDate = start ? new Date(`${start}T00:00:00Z`) : null;
  const eventDay = task.eventDay || (eventDate && !Number.isNaN(eventDate.getTime())
    ? new Intl.DateTimeFormat('id-ID', { weekday: 'long', timeZone: 'UTC' }).format(eventDate)
    : '');
  const sequence = taskNumberSequence(task.number);
  const year = (task.issuedDate || '').slice(0, 4) || String(new Date().getFullYear());
  const officialNumber = `400.3.5/${sequence}/SMPN.13/${year}`;
  const principal = officialFor('principal', task.issuedDate);
  const treasurer = officialFor('treasurer', task.issuedDate);
  const days = Math.max(1, Number(task.transportDays) || 1);
  const selectedTeacherId = String(sppdPersonId || task.sppdPerson || '');
  const selectedTeacherName = state.teachers.find((teacher) => teacher.id === selectedTeacherId)?.name || '';
  const assignedPerson = Array.isArray(task.assignees) ? task.assignees.find((person) => person.teacherId === selectedTeacherId || person.id === selectedTeacherId || (selectedTeacherName && person.name === selectedTeacherName)) : null;
  const selectedPerson = assignedPerson ? parseAssignees([assignedPerson])[0] : people[0] || [];
  const [cityLogoSrc, schoolLogoSrc] = await Promise.all([
    loadAssetDataUrl(window.GAS_ASSETS?.cityLogo || '/assets/logo-kota.jpeg'),
    loadAssetDataUrl(window.GAS_ASSETS?.schoolLogo || '/assets/logo-smpn13.png')
  ]);
  const standardHeader = `<header class="official-header"><img class="school-logo city-logo" src="${cityLogoSrc}" alt="Lambang Kota Tasikmalaya"><div><strong>PEMERINTAH KOTA TASIKMALAYA</strong><strong>SMP NEGERI 13 TASIKMALAYA</strong><p>Jl. Letjen H. Ibrahim Adjie Km.2 Indihiang Tasikmalaya 46151 Telp. (0265) 335695</p><p>e-mail : smp13tasikmalaya@yahoo.co.id</p></div><img class="school-logo school-crest" src="${schoolLogoSrc}" alt="Lambang SMP Negeri 13 Tasikmalaya"></header>`;
  const transportHeader = `<header class="official-header transport-header"><img class="school-logo city-logo" src="${cityLogoSrc}" alt="Lambang Kota Tasikmalaya"><div><strong>PEMERINTAH KOTA TASIKMALAYA</strong><strong>SMP NEGERI 13 TASIKMALAYA</strong><p>Jalan.Letjen.H.Ibrahim Adjie Km.2 Indihiang Tasikmalaya Telp.335695&nbsp;&nbsp; Pos.46151</p><p>e-mail : smp13tasikmalaya@yahoo.co.id</p></div><img class="school-logo school-crest" src="${schoolLogoSrc}" alt="Lambang SMP Negeri 13 Tasikmalaya"></header>`;
  const details = `<table class="letter-details"><tr><td>Dasar Surat</td><td>:</td><td>${escapeHtml(task.basis || '')}</td></tr><tr><td>Nomor</td><td>:</td><td>${escapeHtml(task.basisNumber || '')}</td></tr><tr><td>Tanggal</td><td>:</td><td>${escapeHtml(printLongDate(task.basisDate))}</td></tr><tr><td>Perihal</td><td>:</td><td>${escapeHtml(task.subject || '')}</td></tr></table>`;
  const eventInfo = `<p class="event-purpose">Untuk Kegiatan ${escapeHtml(task.eventName || '')} Pada :</p><table class="letter-details event-details"><tr><td>Hari</td><td>:</td><td>${escapeHtml(eventDay)}</td></tr><tr><td>Tanggal</td><td>:</td><td>${escapeHtml(dateRange)}</td></tr><tr><td>Waktu</td><td>:</td><td>${escapeHtml(task.eventTime || '')}</td></tr><tr><td>Tempat</td><td>:</td><td>${escapeHtml(task.place || task.recipient || '')}</td></tr></table>`;
  const teacherRows = people.map((person, index) => `<tr><td>${index + 1}</td><td>${escapeHtml(person[0] || '')}</td><td>${escapeHtml(person[1] || '')}</td><td>${escapeHtml(person[3] || '')}</td><td>${escapeHtml(person[5] || '')}</td></tr>`).join('');
  const studentRows = students.map((student, index) => `<tr><td>${index + 1}</td><td>${escapeHtml(student[0] || '')}</td><td>${escapeHtml(student[1] || '')}</td><td>${escapeHtml(student[2] || '')}</td><td></td></tr>`).join('');
  const assignments = `<table class="assignment-table"><thead><tr><th>No</th><th>N a m a</th><th>NIP</th><th>Pangkat/ Gol</th><th>Jabatan</th></tr></thead><tbody>${teacherRows}</tbody></table>`;
  const pupils = `<table class="assignment-table"><thead><tr><th>No</th><th>N a m a</th><th>NIS</th><th>Kelas</th><th>Ket</th></tr></thead><tbody>${studentRows}</tbody></table>`;
  const signature = `<div class="official-signature"><p>Dikeluarkan di&nbsp;&nbsp; : Tasikmalaya</p><p>Pada Tanggal&nbsp;&nbsp;&nbsp;&nbsp; : ${escapeHtml(printLongDate(task.issuedDate))}</p><p>Kepala Sekolah</p><div class="signature-space"></div><p><strong>${escapeHtml(principal.name || '')}</strong></p><p>NIP.&nbsp; ${escapeHtml(principal.nip || '')}</p></div>`;
  let title = '';
  let body = '';
  let header = standardHeader;
  let orientation = 'portrait';

  if (type === 'task' || type === 'student') {
    title = 'S U R A T&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; T U G A S';
    const list = type === 'task' ? assignments : pupils;
    body = `<p class="official-number">Nomor : ${escapeHtml(officialNumber)}</p>${details}<p class="order-heading">KEPALA SMP NEGERI 13 TASIKMALAYA</p><p class="order-heading">MENUGASKAN</p>${list}${eventInfo}<p class="closing-line">Demikian surat tugas ini kami berikan agar dilaksanakan sebagaimana mestinya.</p>${signature}`;
  } else if (type === 'sppd') {
    title = 'S U R A T&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; P E R I N T A H&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; P E R J A L A N A N&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; D I N A S';
    body = `<div class="sppd-identifiers"><p>Lembaran Ke&nbsp;&nbsp;&nbsp; : I</p><p>Kode No&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : __________________</p><p>SPPD No&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : 400.3.5/<span class="number-slot">${escapeHtml(sequence)}</span>/SMPN.13/${escapeHtml(year)}</p></div><h1 class="sppd-title">SURAT PERINTAH PERJALANAN DINAS<br>(SPPD)</h1><table class="sppd-table"><tr><td>1. Pengguna Anggaran / Kuasa Pengguna Anggaran</td><td>Kepala Sekolah</td></tr><tr><td>2. Nama Pegawai yang Diperintah / NIP</td><td>${escapeHtml(selectedPerson[0] || '')}<br>${escapeHtml(selectedPerson[1] || '')}</td></tr><tr><td>3. a. Pangkat dan Golongan</td><td>${escapeHtml([selectedPerson[4], selectedPerson[3]].filter(Boolean).join(' / '))}</td></tr><tr><td>&nbsp;&nbsp;&nbsp; b. Jabatan / OPD</td><td>${escapeHtml(selectedPerson[5] || '')}</td></tr><tr><td>&nbsp;&nbsp;&nbsp; c. Tingkat Biaya Perjalanan Dinas</td><td></td></tr><tr><td>4. Maksud Perjalanan Dinas</td><td>${escapeHtml(task.eventName || '')}</td></tr><tr><td>5. Alat Angkut yang Dipergunakan</td><td>Kendaraan</td></tr><tr><td>6. a. Tempat Berangkat</td><td>SMP Negeri 13 Tasikmalaya</td></tr><tr><td>&nbsp;&nbsp;&nbsp; b. Tempat Tujuan</td><td>${escapeHtml(task.recipient || '')}</td></tr><tr><td>7. a. Lamanya Perjalanan Dinas</td><td>${days} hari</td></tr><tr><td>&nbsp;&nbsp;&nbsp; b. Tanggal Berangkat</td><td>${escapeHtml(printLongDate(start))}</td></tr><tr><td>&nbsp;&nbsp;&nbsp; c. Tanggal Harus Kembali / Tiba di Tempat Baru</td><td>${escapeHtml(printLongDate(end))}</td></tr><tr><td>8. Pengikut : Nama</td><td></td></tr><tr><td>9. Pembebanan Anggaran</td><td></td></tr><tr><td>&nbsp;&nbsp;&nbsp; a. Instansi</td><td>-</td></tr><tr><td>&nbsp;&nbsp;&nbsp; b. Mata Anggaran</td><td>-</td></tr><tr><td>10. Keterangan Lain-lain</td><td>${escapeHtml(task.subject || '')}</td></tr></table><div class="sppd-signature-grid"><div class="sppd-sign-box"><p>Dikeluarkan di&nbsp;&nbsp; : Tasikmalaya</p><p>Pada Tanggal&nbsp; : ${escapeHtml(printLongDate(task.issuedDate))}</p><p class="sppd-sign-label">Pengguna Anggaran / Kuasa Pengguna Anggaran</p><div class="signature-space"></div><p><strong>${escapeHtml(principal.name || '')}</strong></p><p>NIP. ${escapeHtml(principal.nip || '')}</p></div></div>`;
  } else if (type === 'sppd-continuation') {
    title = '';
    const dateBlocks = Array.from({ length: days }, () => {
      return `<tr class="continuation-travel-row"><td><p>I. Tiba di : ${escapeHtml(task.recipient || '')}</p><p>Pada Tanggal : .............................................</p><div class="continuation-signature"></div><p class="continuation-signature-name">(........................................)</p></td><td><p>Berangkat dari : ${escapeHtml(task.recipient || '')}</p><p>Ke &nbsp;: SMP Negeri 13 Tasikmalaya</p><p>Pada Tanggal : .............................................</p><div class="continuation-signature"></div><p class="continuation-signature-name">(........................................)</p></td></tr>`;
    }).join('');
    body = `<table class="continuation-sheet"><tbody><tr class="continuation-top"><td></td><td><p>SPPD No. &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : ${escapeHtml(officialNumber)}</p><p>Berangkat dari &nbsp;: SMPN 13 Tasikmalaya</p><p>(Tempat Kedudukan)</p><p>Ke &nbsp;: ${escapeHtml(task.recipient || '')}</p></td></tr><tr><th class="continuation-heading" colspan="2">MENGETAHUI PEJABAT SETEMPAT :</th></tr>${dateBlocks}<tr class="continuation-return"><td><p>II. Tiba kembali di : SMP Negeri 13 Tasikmalaya</p><p class="return-date">Pada Tanggal : .............................................</p><div class="continuation-signature"></div></td><td><p>Telah diperiksa, dengan keterangan :</p><p>Bahwa perjalanan dinas tersebut di atas benar dilakukan atas perintahnya dan semata-mata untuk kepentingan jabatan dalam waktu yang sesingkat-singkatnya.</p><p>Kepala Sekolah</p><div class="continuation-signature"></div><p class="continuation-official"><strong>${escapeHtml(principal.name || '')}</strong><br>NIP. ${escapeHtml(principal.nip || '')}</p></td></tr><tr class="continuation-notes"><td>III. CATATAN LAIN-LAIN :</td><td></td></tr><tr><th class="continuation-heading attention-heading" colspan="2">IV. PERHATIAN</th></tr><tr><td class="continuation-attention" colspan="2">Pejabat yang berwenang menerbitkan SPPD, pegawai yang melakukan perjalanan dinas, para pejabat yang mengesahkan tanggal berangkat/tiba serta bendahara bertanggung jawab berdasarkan peraturan-peraturan Keuangan Negara apabila Negara mendapat rugi akibat kesalahan, kealpaan.</td></tr></tbody></table>`;
    orientation = 'portrait';
  } else if (type === 'nota') {
    title = 'NOTA DINAS';
    body = `<p class="subtitle">(Laporan Hasil Pelaksanaan Kegiatan)</p><table class="nota-table"><tr><td>1</td><td>Kepada Yth.</td><td>${escapeHtml(task.reportTo || 'Kepala Sekolah')}</td></tr><tr><td>2</td><td>Dari (yang diberi tugas)</td><td>${escapeHtml(selectedPerson[0] || '')}</td></tr><tr><td>3</td><td>Perihal/ Nama Kegiatan</td><td>${escapeHtml(task.eventName || '')}</td></tr><tr><td>4</td><td>Tempat Kegiatan</td><td>${escapeHtml(task.recipient || '')}</td></tr><tr><td>5</td><td>Tanggal Pelaksanaan</td><td>${escapeHtml(dateRange)}</td></tr><tr><td>6</td><td>Dasar Penugasan</td><td>${escapeHtml(task.basis || '')}</td></tr><tr><td></td><td>a. Surat Tugas dari</td><td>Kepala Sekolah</td></tr><tr><td></td><td>b. Nomor</td><td>${escapeHtml(sequence)}</td></tr><tr><td></td><td>c. Tanggal</td><td>${escapeHtml(printLongDate(task.issuedDate))}</td></tr></table><p class="report-intro">Berdasarkan Surat Tugas tersebut di atas, saya laporkan hasil-hasil pelaksanaan kegiatan sebagai berikut :</p><div class="report-space official-report-lines">${escapeHtml(task.reportText || '').replace(/\r?\n/g, '<br>')}</div><p class="recommendation-heading">Usul, masukan, rekomendasi dan tindak lanjut :</p><div class="report-space official-report-lines recommendation-lines">${escapeHtml(task.recommendations || '').replace(/\r?\n/g, '<br>')}</div><p class="nota-closing">Demikian Nota Dinas/ Laporan hasil pelaksanaan kegiatan ini dibuat dengan sebenarnya sebagai bentuk pertanggungjawaban atas pelaksanaan tugas yang telah dipercayakan kepada saya.</p><div class="nota-sign"><p>Tasikmalaya, .........................................</p><p>Yang Membuat Laporan,</p><div class="signature-space"></div><p>___________________________</p><p>NIP.</p></div>`;
  } else if (type === 'transport' || type === 'transport-students') {
    const isStudent = type === 'transport-students';
    header = transportHeader;
    title = isStudent ? 'DAFTAR PENYERAHAN TRANSPORT SISWA' : 'DAFTAR PENYERAHAN TRANSPORT';
    const totalAmount = (isStudent ? students : people).reduce((total, person) => total + days * (Number(isStudent ? person[3] : person[6]) || 0), 0);
    const rows = isStudent ? students.map((student, index) => {
      const rate = Number(student[3]) || 0;
      return `<tr><td>${index + 1}</td><td>${escapeHtml(student[0] || '')}</td><td>${escapeHtml(student[2] || '')}</td><td>${days} Hari x ${money(rate)}</td><td>${money(days * rate)}</td><td></td></tr>`;
    }).join('') : people.map((person, index) => {
      const rate = Number(person[6]) || 0;
      return `<tr><td>${index + 1}</td><td>${escapeHtml(person[0] || '')}<br>NIP.${escapeHtml(person[1] || '')}</td><td>${escapeHtml(person[3] || '')}</td><td>${days} Hari x ${money(rate)}</td><td>${money(days * rate)}</td><td></td></tr>`;
    }).join('');
    body = `<p class="event-banner">${escapeHtml((task.eventName || '').toUpperCase())}</p><p class="official-number">Nomor : ${escapeHtml(officialNumber)}</p><table class="transport-table${isStudent ? ' transport-students' : ''}"><thead><tr><th>NO</th><th>Nama</th><th>${isStudent ? 'Kelas' : 'Gol'}</th><th>Volume</th><th>Jumlah Yang Di Terima</th><th>Tanda Tangan</th></tr></thead><tbody>${rows}</tbody><tfoot><tr class="transport-total"><td colspan="4"><strong>JUMLAH</strong></td><td><strong>${money(totalAmount)}</strong></td><td></td></tr></tfoot></table><div class="transport-date-row"><div></div><div>Tasikmalaya, ${escapeHtml(printLongDate(task.issuedDate))}</div></div><div class="transport-approval"><div class="approval-box"><p>Setuju Dibayar :</p><p>Kepala SMPN 13 Tasikmalaya</p><div class="signature-space"></div><p><strong>${escapeHtml(principal.name || '')}</strong></p><p>NIP. ${escapeHtml(principal.nip || '')}</p></div><div class="approval-box"><p>Lunas Dibayar Tanggal :</p><p>Bendahara,</p><div class="signature-space"></div><p><strong>${escapeHtml(treasurer.name || '')}</strong></p><p>NIP. ${escapeHtml(treasurer.nip || '')}</p></div></div>`;
    orientation = 'portrait';
  } else {
    return;
  }

  if (type === 'sppd') {
    title = '';
    body = `<table class="sppd-identifiers-grid"><colgroup><col><col><col><col></colgroup><tbody><tr><td rowspan="3"></td><td>Lembaran Ke</td><td>:</td><td>I</td></tr><tr><td>Kode No</td><td>:</td><td></td></tr><tr><td>SPPD No</td><td>:</td><td>400.3.5/${escapeHtml(sequence)}/SMPN.13/${escapeHtml(year)}</td></tr></tbody></table><h1 class="sppd-title">SURAT PERINTAH PERJALANAN DINAS<br>(SPPD)</h1><table class="sppd-table"><tbody><tr><td>1. Pengguna Anggaran/Kuasa Pengguna Anggaran</td><td>Kepala Sekolah</td></tr><tr><td>2. Nama Pegawai yang diperintah / NIP</td><td>${escapeHtml(selectedPerson[0] || '')}<br>${escapeHtml(selectedPerson[1] || '')}</td></tr><tr><td>3. a. Pangkat dan Golongan</td><td>${escapeHtml([selectedPerson[4], selectedPerson[3]].filter(Boolean).join(' / '))}</td></tr><tr><td>&nbsp;&nbsp;&nbsp; b. Jabatan/OPD</td><td>${escapeHtml(selectedPerson[5] || '')}</td></tr><tr><td>&nbsp;&nbsp;&nbsp; c. Tingkat Biaya Perjalanan Dinas</td><td></td></tr><tr><td>4. Maksud Perjalanan Dinas</td><td>${escapeHtml(task.eventName || '')}</td></tr><tr><td>5. Alat angkut yang digunakan</td><td>${escapeHtml(task.transportMode || 'Kendaraan')}</td></tr><tr><td>6. a. Tempat Berangkat</td><td>a. ${escapeHtml(task.origin || 'SMP Negeri 13 Tasikmalaya')}</td></tr><tr><td>&nbsp;&nbsp;&nbsp; b. Tempat Tujuan</td><td>b. ${escapeHtml(task.recipient || '')}</td></tr><tr><td>7. a. Lamanya Perjalanan Dinas</td><td>${days} hari</td></tr><tr><td>&nbsp;&nbsp;&nbsp; b. Tanggal Berangkat</td><td>${escapeHtml(printLongDate(start))}</td></tr><tr><td>&nbsp;&nbsp;&nbsp; c. Tanggal Harus Kembali/Tiba di Tempat Baru</td><td>${escapeHtml(printLongDate(end))}</td></tr><tr><td>8. Pengikut : Nama</td><td><table class="sppd-follower-table"><thead><tr><th>Tanggal Lahir</th><th>Keterangan</th></tr></thead><tbody><tr><td>&nbsp;</td><td>&nbsp;</td></tr></tbody></table></td></tr><tr><td>9. Pembebanan Anggaran</td><td></td></tr><tr><td>&nbsp;&nbsp;&nbsp; a. Instansi</td><td>a. -</td></tr><tr><td>&nbsp;&nbsp;&nbsp; b. Mata Anggaran</td><td>b. -</td></tr><tr><td>10. Keterangan Lain-lain</td><td>${escapeHtml(task.subject || '')}</td></tr></tbody></table><div class="sppd-signature-grid"><div class="sppd-sign-box"><p>Dikeluarkan di&nbsp;&nbsp; : Tasikmalaya</p><p>Pada Tanggal&nbsp; : ${escapeHtml(printLongDate(task.issuedDate))}</p><p>Pengguna Anggaran/Kuasa Pengguna Anggaran</p><div class="signature-space"></div><p><strong>${escapeHtml(principal.name || '')}</strong></p><p>NIP. ${escapeHtml(principal.nip || '')}</p></div></div>`;
  }
  if (type === 'sppd-continuation') title = '';
  body = body.replaceAll('AGUS ROHMAN, S.Pd., M.Si.', escapeHtml(principal.name || ''))
    .replaceAll('19650927 198903 1 008', escapeHtml(principal.nip || ''));
  const documentTitle = type === 'sppd-continuation' ? officialNumber : `${title || 'Dokumen Resmi'} - ${officialNumber}`;
  document.getElementById('print-dialog').close();
  printWindow.document.write(`<!doctype html><html lang="id"><head><meta charset="utf-8"><title>${escapeHtml(documentTitle)}</title><style>
    @page{size:A4 ${orientation};margin:12.7mm}*{box-sizing:border-box}body{margin:0;color:#111;font:11pt Cambria,serif;line-height:1.3}.official-header{text-align:center;border-bottom:3px double #111;padding:0 0 7px;margin-bottom:18px}.official-header strong{display:block;font-size:14pt}.official-header p{margin:3px 0;font-size:10pt}.official-title{text-align:center;margin:13px 0 16px;font-size:12pt;font-weight:bold}.official-number{text-align:center;margin:0 0 17px;font-size:12pt}.letter-details{width:100%;border-collapse:collapse;margin:8px 0 13px}.letter-details td{padding:3px 4px;vertical-align:top;font-size:11pt}.letter-details td:first-child{width:145px}.letter-details td:nth-child(2){width:16px;text-align:center}.event-details{margin-top:18px}.event-purpose{margin:9px 0 3px;text-align:left;font-size:11pt}.order-heading{text-align:center;margin:5px 0;font-weight:bold}.assignment-table,.transport-table{width:100%;border-collapse:collapse;margin:12px 0}.assignment-table th,.assignment-table td,.transport-table th,.transport-table td{border:1px solid #222;padding:5px;text-align:center;vertical-align:top}.assignment-table td:nth-child(2),.transport-table td:nth-child(2){text-align:left}.transport-table{font-size:12pt}.transport-table.transport-students{font-size:10pt}.transport-table td:nth-child(5),.transport-table tfoot td:nth-child(2){text-align:right}.event-banner{text-align:center}.assignment-table th,.transport-table th{text-align:center}.assignment-table td:first-child,.transport-table td:first-child{text-align:center}.closing-line{margin-top:15px}.official-signature{width:255px;margin:20px 12px 0 auto}.official-signature p{margin:4px 0}.signature-space{height:54px}.subtitle{text-align:center;margin:4px 0 14px}.sppd-table,.nota-table{width:100%;border-collapse:collapse}.nota-table{font-size:10pt}.nota-table td:nth-child(3),.report-intro,.report-space,.nota-closing{text-align:justify}.sppd-table td,.nota-table td{padding:4px;vertical-align:top}.sppd-table td:first-child{width:56%}.nota-table td:first-child{width:28px}.nota-table td:nth-child(2){width:190px}.sppd-sign,.nota-sign{width:300px;margin:18px 0 0 auto}.sppd-sign p,.nota-sign p{margin:4px 0}.travel-day{width:100%;border-collapse:collapse;margin:16px 0;page-break-inside:avoid}.travel-day td{width:33.33%;padding:7px;border:1px solid #444;vertical-align:top}.report-intro{margin:14px 0 8px}.report-space{min-height:52px;white-space:normal;margin:5px 0}.recommendation-heading{margin:12px 0 5px}.sppd-signature-grid{display:flex;justify-content:flex-end;margin-top:20px}.sppd-sign-box{width:320px;text-align:left}.sppd-sign-box p{margin:4px 0}.sppd-sign-label{margin-top:12px!important}.sppd-route-line{margin:4px 0}.sppd-note{margin:8px 0;line-height:1.5}.transport-header strong{font-size:16pt}.transport-header p{font-size:12pt}.print-type-sppd .official-header strong:first-child{font-size:12pt}.print-type-nota .official-header strong:nth-child(2){font-size:16pt}.print-type-nota .official-title{font-size:14pt}.transport-date-row{display:grid;grid-template-columns:1fr auto;gap:12px;margin:18px 0 10px;font-weight:bold}.transport-approval{display:grid;grid-template-columns:1fr 1fr;gap:30px;margin-top:18px;page-break-inside:avoid}.transport-approval>div{min-width:0}.transport-approval p{margin:4px 0}.transport-approval .signature-space{height:60px}.return-review{width:100%;border-collapse:collapse;margin-top:14px;page-break-inside:avoid}.return-review td{width:50%;padding:7px;border-top:1px solid #333;border-bottom:1px solid #333;vertical-align:top}.return-review td+td{border-left:1px solid #333}.return-review p{margin:8px 0}.return-review .signature-space{height:48px}.sppd-notes-heading{margin:8px 0;font-weight:bold}.sppd-notes-space{height:56px;border-bottom:1px solid #777}.approval-box p{margin:4px 0}
    </style><style>.official-header{display:grid;grid-template-columns:72px minmax(0,1fr) 72px;align-items:center;gap:10px;text-align:center;border-bottom:3px double #111;padding:0 0 8px;margin:0 0 18px}.official-header>div{min-width:0}.official-header .school-logo{display:block;width:66px;height:72px;object-fit:contain}.official-header .city-logo{justify-self:start}.official-header .school-crest{justify-self:end}.official-header strong{display:block;font-size:14pt}.official-header p{margin:3px 0;font-size:10pt}.transport-header strong{font-size:16pt}.transport-header p{font-size:12pt}.print-type-sppd .official-header strong:first-child{font-size:12pt}.print-type-nota .official-header strong:nth-child(2){font-size:16pt}.sppd-identifiers{font-size:11px}.sppd-identifiers p{margin:2px 0}.number-slot{display:inline-block;min-width:52px;border-bottom:1px solid #111;text-align:center}.sppd-title{text-align:center;margin:9px 0 14px;font-size:14pt}.travel-day{table-layout:fixed;margin:12px 0 16px;border:1px solid #333}.travel-day td{width:50%;height:auto;padding:7px 9px;border:1px solid #444;vertical-align:top}.travel-day tr:nth-child(n+2) td{height:38px}.travel-day tr:last-child td{height:46px}.official-report-lines{height:360px;min-height:360px;overflow:hidden;padding:2px 4px;line-height:24px;background-image:linear-gradient(to bottom,transparent 23px,#888 23px,#888 24px,transparent 24px);background-size:100% 24px;white-space:normal;page-break-inside:avoid}.official-report-lines.recommendation-lines{height:120px;min-height:120px}.nota-closing{margin-top:14px}.sppd-table tr{min-height:25px}</style></head><body><main class="print-sheet print-type-${type}" id="print-sheet">${header}${title ? `<h1 class="official-title">${title}</h1>` : ''}${body}</main><script>window.onload=()=>window.print()<\/script></body></html>`);
  const autoPrintScript = printWindow.document.querySelector('script');
  if (autoPrintScript?.textContent.includes('window.print')) autoPrintScript.remove();
  printWindow.onload = null;
  const defaultPrintFontSize = type === 'nota' || type === 'transport-students' ? '10' : type.startsWith('transport') ? '12' : '11';
  const printToolbar = printWindow.document.createElement('div');
  printToolbar.className = 'print-toolbar';
  printToolbar.innerHTML = `<div class="print-toolbar-actions"><button type="button" class="print-now-button" id="print-now">Cetak</button><button type="button" id="edit-layout">Edit Layout</button><button type="button" id="save-layout">Jadikan Default</button><button type="button" id="reset-layout">Reset Layout</button><button type="button" id="close-print-preview">Tutup pratinjau</button><details><summary>Pengaturan halaman</summary><div class="print-settings"><label>Ukuran kertas<select id="print-page-size"><option value="F4">F4 / Folio (215 × 330 mm)</option><option value="A4" selected>A4 (210 × 297 mm)</option><option value="Letter">Letter (215,9 × 279,4 mm)</option><option value="Legal">Legal (215,9 × 355,6 mm)</option><option value="custom">Custom…</option></select></label><label class="custom-size" hidden>Lebar custom (mm)<input id="print-custom-width" type="number" min="50" max="600" step="0.1" value="215"></label><label class="custom-size" hidden>Tinggi custom (mm)<input id="print-custom-height" type="number" min="50" max="600" step="0.1" value="330"></label><label>Orientasi<select id="print-orientation"><option value="portrait" ${orientation === 'portrait' ? 'selected' : ''}>Portrait</option><option value="landscape" ${orientation === 'landscape' ? 'selected' : ''}>Landscape</option></select></label><label>Margin atas (mm)<input id="print-margin-top" type="number" min="0" max="60" step="0.5" value="12.7"></label><label>Margin kanan (mm)<input id="print-margin-right" type="number" min="0" max="60" step="0.5" value="12.7"></label><label>Margin bawah (mm)<input id="print-margin-bottom" type="number" min="0" max="60" step="0.5" value="12.7"></label><label>Margin kiri (mm)<input id="print-margin-left" type="number" min="0" max="60" step="0.5" value="12.7"></label><label>Jenis huruf<select id="print-font"><option value="Cambria, serif" selected>Cambria</option><option value="'Times New Roman', serif">Times New Roman</option><option value="Georgia, serif">Georgia</option><option value="Arial, sans-serif">Arial</option></select></label><label>Ukuran huruf <output id="print-font-size-value">${defaultPrintFontSize} pt</output><input id="print-font-size" type="range" min="9" max="16" value="${defaultPrintFontSize}"></label><label>Jarak baris<select id="print-line-height"><option value="1.2">Rapat</option><option value="1.3" selected>Normal</option><option value="1.5">Renggang</option></select></label></div></details><details id="layout-editor-details"><summary>Editor layout</summary><div class="layout-inspector"><p class="layout-hint">Klik/seret elemen. Tab geser 10 px ke kanan; Shift+Tab geser 10 px ke kiri. Panah menggeser 1 px.</p><div class="layout-selected" id="layout-selected">Elemen terpilih: <strong>—</strong></div><button type="button" id="layout-select-table" hidden>Pilih seluruh tabel</button><label id="layout-text-control" hidden>Edit teks <textarea id="layout-text" rows="3" placeholder="Pilih teks pada pratinjau"></textarea></label><div class="layout-field"><span>Perataan teks</span><div class="align-buttons" id="layout-align"><button type="button" data-align="left">Kiri</button><button type="button" data-align="center">Tengah</button><button type="button" data-align="right">Kanan</button><button type="button" data-align="justify">Rata</button></div></div><label>Ukuran huruf (pt) <input id="layout-font-size" type="number" min="6" max="48" step="0.5" placeholder="default"></label><label>Jarak baris <input id="layout-line-height" type="number" min="0.8" max="3" step="0.05" placeholder="default"></label><div class="layout-subhead">Ukuran elemen</div><div class="layout-two"><label>Lebar (px) <input id="layout-width" type="number" min="40" step="1" placeholder="auto"></label><label>Tinggi (px) <input id="layout-height" type="number" min="24" step="1" placeholder="auto"></label></div><div class="layout-two"><label>Geser X (px) <input id="layout-x" type="number" step="1" value="0"></label><label>Geser Y (px) <input id="layout-y" type="number" step="1" value="0"></label></div><div id="layout-table-controls" hidden><div class="layout-subhead">Pengaturan tabel</div><div class="layout-two"><label>Tinggi baris (px) <input id="layout-row-height" type="number" min="0" step="1" placeholder="auto"></label><label>Padding sel (px) <input id="layout-cell-padding" type="number" min="0" step="1" placeholder="default"></label></div><div class="layout-columns" id="layout-columns"></div></div></div></details></div>`;
  printWindow.document.body.insertBefore(printToolbar, printWindow.document.body.firstChild);
  const printSettingsStyle = printWindow.document.createElement('style');
  printSettingsStyle.textContent = '.print-toolbar{position:sticky;top:0;z-index:5;margin:0 0 18px;padding:10px 12px;border:1px solid #cbd5cd;background:#f5f8f5;color:#26352b;font:12px Arial,sans-serif}.print-toolbar-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap}.print-toolbar button,.print-toolbar summary{min-height:32px;border:1px solid #bcc9be;border-radius:3px;background:#fff;padding:0 11px;color:#26352b;font:600 12px Arial,sans-serif;cursor:pointer}.print-toolbar .print-now-button{border-color:#315f46;background:#315f46;color:#fff}.print-toolbar details{position:relative}.print-toolbar summary{display:flex;align-items:center}.print-settings{position:absolute;top:38px;left:0;display:grid;grid-template-columns:repeat(3,minmax(130px,1fr));gap:10px;width:min(620px,calc(100vw - 32px));border:1px solid #cbd5cd;background:#fff;padding:14px;box-shadow:0 8px 24px #17281b24}.print-settings label{display:grid;gap:5px;color:#506055;font-size:11px}.print-settings select{min-height:32px;border:1px solid #d3dcd4;border-radius:3px;background:#fff;padding:4px 7px;color:#26352b;font:12px Arial,sans-serif}.print-settings input[type=range]{width:100%;accent-color:#315f46}.print-settings output{color:#26352b;font-weight:600}.print-settings label:has(input){grid-template-columns:1fr auto}.print-settings input{grid-column:1/-1}.print-toolbar button:hover,.print-toolbar summary:hover{background:#eef3ee}.print-settings input[type=number]{min-height:32px;border:1px solid #d3dcd4;border-radius:3px;background:#fff;padding:4px 7px;color:#26352b;font:12px Arial,sans-serif}.print-settings label.custom-size{grid-column:1/-1}.print-toolbar button.is-active{border-color:#315f46;background:#315f46;color:#fff}.layout-inspector{display:grid;gap:10px}.layout-inspector .layout-hint{margin:0;color:#506055;font-size:11px;line-height:1.5}.layout-inspector .layout-selected{color:#26352b;font-size:11px}.layout-inspector .layout-field{display:grid;gap:5px}.layout-inspector .layout-field>span,.layout-inspector label{display:grid;gap:4px;color:#506055;font-size:11px}.layout-inspector .align-buttons{display:grid;grid-template-columns:repeat(4,1fr);gap:5px}.layout-inspector .align-buttons button{min-height:30px;border:1px solid #bcc9be;border-radius:3px;background:#fff;color:#26352b;font:600 11px Arial,sans-serif;cursor:pointer;padding:0 4px}.layout-inspector .align-buttons button.is-active{border-color:#315f46;background:#315f46;color:#fff}.layout-inspector .layout-two{display:grid;grid-template-columns:1fr 1fr;gap:8px}.layout-inspector textarea{width:100%;min-height:70px;resize:vertical;border:1px solid #d3dcd4;border-radius:3px;background:#fff;padding:6px 7px;color:#26352b;font:12px Arial,sans-serif}.layout-inspector input[type=number]{min-height:30px;border:1px solid #d3dcd4;border-radius:3px;background:#fff;padding:4px 7px;color:#26352b;font:12px Arial,sans-serif}.layout-inspector .layout-subhead{margin-top:2px;color:#315f46;font-weight:700;font-size:11px}.layout-inspector .layout-columns{display:grid;gap:6px}.layout-inspector .layout-column{display:grid;grid-template-columns:64px minmax(54px,1fr) 72px 72px;align-items:center;gap:6px;color:#506055;font-size:11px}.layout-inspector .layout-column input,.layout-inspector .layout-column select{min-height:28px;border:1px solid #d3dcd4;border-radius:3px;background:#fff;padding:2px 6px;color:#26352b;font:11px Arial,sans-serif}.print-settings{position:absolute;top:38px;left:0;z-index:6;display:grid;grid-template-columns:repeat(3,minmax(130px,1fr));gap:10px;width:min(680px,calc(100vw - 32px));max-height:calc(100vh - 90px);overflow:auto;border:1px solid #cbd5cd;background:#fff;padding:14px;box-shadow:0 8px 24px #17281b24}.print-toolbar details[open] .layout-inspector{position:fixed;top:88px;right:16px;z-index:6;width:min(380px,calc(100vw - 32px));max-height:calc(100vh - 104px);overflow:auto;border:1px solid #cbd5cd;border-radius:4px;background:#fff;padding:14px;box-shadow:0 8px 24px #17281b24}.layout-inspector [hidden]{display:none!important}.print-toolbar details[open] .layout-inspector{grid-template-columns:1fr}.print-toolbar details[open] .layout-inspector .layout-two{grid-template-columns:1fr 1fr}.print-toolbar details[open] .layout-inspector .layout-column{grid-template-columns:64px minmax(54px,1fr) 72px 72px}body{background:#eef1ee}.print-sheet{background:#fff;box-shadow:0 2px 14px #0b1a121f;margin:0 auto 28px}body.layout-editing .print-sheet [data-layout-block]{outline:1px dashed #7fae8f;outline-offset:2px;cursor:move}body.layout-editing .print-sheet [data-layout-block]:hover{outline-color:#315f46;background:#f3f8f4aa}[data-layout-block].is-selected{outline:2px solid #d98324!important;outline-offset:2px}body.layout-panel-open #print-sheet{margin-left:16px;margin-right:410px}@media(max-width:1100px){.print-toolbar details[open] .layout-inspector{position:static;top:auto;right:auto;width:auto;max-height:45vh;box-shadow:none}body.layout-panel-open #print-sheet{margin:0 auto 28px}}@media print{body{background:#fff}.print-sheet{box-shadow:none;margin:0;width:auto!important;min-height:0!important;padding:0!important}.print-toolbar{display:none!important}body.layout-editing .print-sheet [data-layout-block]{outline:none!important;background:none!important}}';
  printWindow.document.head.appendChild(printSettingsStyle);
  const layoutResizeStyle = printWindow.document.createElement('style');
  layoutResizeStyle.textContent = '.selection-resizer{position:fixed;z-index:8;pointer-events:none;border:2px solid #d98324;box-sizing:border-box}.selection-resizer[hidden]{display:none!important}.selection-resizer span{position:absolute;width:10px;height:10px;margin:-5px;border:1px solid #d98324;border-radius:2px;background:#fff;pointer-events:auto}.selection-resizer [data-resize-edge=n]{left:50%;top:0;cursor:ns-resize}.selection-resizer [data-resize-edge=s]{left:50%;bottom:0;cursor:ns-resize}.selection-resizer [data-resize-edge=e]{right:0;top:50%;cursor:ew-resize}.selection-resizer [data-resize-edge=w]{left:0;top:50%;cursor:ew-resize}.selection-resizer [data-resize-edge=ne]{right:0;top:0;cursor:nesw-resize}.selection-resizer [data-resize-edge=nw]{left:0;top:0;cursor:nwse-resize}.selection-resizer [data-resize-edge=se]{right:0;bottom:0;cursor:nwse-resize}.selection-resizer [data-resize-edge=sw]{left:0;bottom:0;cursor:nesw-resize}@media print{.selection-resizer{display:none!important}}body.layout-editing.col-resize-hover .print-sheet table,body.layout-editing.col-resize-hover .print-sheet td,body.layout-editing.col-resize-hover .print-sheet th{cursor:col-resize}body.col-resizing,body.col-resizing *{cursor:col-resize!important;user-select:none!important}';
  printWindow.document.head.appendChild(layoutResizeStyle);
  const printSettings = printWindow.document.createElement('style');
  printSettings.id = 'active-print-settings';
  printWindow.document.head.appendChild(printSettings);
  const PAPER_SIZES = { F4: [215, 330], A4: [210, 297], Letter: [215.9, 279.4], Legal: [215.9, 355.6] };
  const MM_TO_PX = 96 / 25.4;
  const LAYOUT_STORAGE_KEY = 'smpn13.agenda-surat.layout.v2';
  const readLayoutStore = () => {
    try { return JSON.parse(localStorage.getItem(LAYOUT_STORAGE_KEY)) || {}; } catch { return {}; }
  };
  const writeLayoutStore = (store) => {
    try { localStorage.setItem(LAYOUT_STORAGE_KEY, JSON.stringify(store)); } catch { /* penyimpanan penuh/ditolak */ }
  };
  const clampNumber = (value, min, max, fallback) => {
    const parsed = Number(value);
    if (!Number.isFinite(parsed)) return fallback;
    return Math.min(max, Math.max(min, parsed));
  };
  const getPaperDimensions = () => {
    const doc = printWindow.document;
    const sizeKey = doc.getElementById('print-page-size').value;
    const isCustom = sizeKey === 'custom';
    doc.querySelectorAll('.custom-size').forEach((field) => { field.hidden = !isCustom; });
    let width;
    let height;
    if (isCustom) {
      width = clampNumber(doc.getElementById('print-custom-width').value, 50, 600, 215);
      height = clampNumber(doc.getElementById('print-custom-height').value, 50, 600, 330);
    } else {
      [width, height] = PAPER_SIZES[sizeKey] || PAPER_SIZES.A4;
    }
    const pageOrientation = doc.getElementById('print-orientation').value;
    if (pageOrientation === 'landscape') { const swap = width; width = height; height = swap; }
    return { width, height };
  };
  const updatePrintSettings = () => {
    const doc = printWindow.document;
    const { width, height } = getPaperDimensions();
    const marginTop = clampNumber(doc.getElementById('print-margin-top').value, 0, 60, 15);
    const marginRight = clampNumber(doc.getElementById('print-margin-right').value, 0, 60, 15);
    const marginBottom = clampNumber(doc.getElementById('print-margin-bottom').value, 0, 60, 15);
    const marginLeft = clampNumber(doc.getElementById('print-margin-left').value, 0, 60, 15);
    const font = doc.getElementById('print-font').value;
    const fontSize = doc.getElementById('print-font-size').value;
    const lineHeight = doc.getElementById('print-line-height').value;
    doc.getElementById('print-font-size-value').textContent = `${fontSize} pt`;
    printSettings.textContent = `@page{size:${width}mm ${height}mm;margin:${marginTop}mm ${marginRight}mm ${marginBottom}mm ${marginLeft}mm}body{font-family:${font};font-size:${fontSize}pt;line-height:${lineHeight}}@media screen{#print-sheet{width:${(width * MM_TO_PX).toFixed(1)}px;min-height:${(height * MM_TO_PX).toFixed(1)}px;padding:${(marginTop * MM_TO_PX).toFixed(1)}px ${(marginRight * MM_TO_PX).toFixed(1)}px ${(marginBottom * MM_TO_PX).toFixed(1)}px ${(marginLeft * MM_TO_PX).toFixed(1)}px}}`;
  };
  let selectedBlock = null;
  let dragState = null;
  let resizeState = null;
  const COLUMN_RESIZE_HANDLE = 6;
  const blockLabel = (el) => `${el.tagName.toLowerCase()}${typeof el.className === 'string' && el.className ? `.${el.className.split(' ')[0]}` : ''}`;
  const readTranslate = (el) => {
    const match = /translate\(\s*(-?[\d.]+)px\s*,\s*(-?[\d.]+)px\s*\)/.exec(el.style.transform || '');
    return match ? { x: Number(match[1]), y: Number(match[2]) } : { x: 0, y: 0 };
  };
  const applyTranslate = (el, x, y) => { el.style.transform = `translate(${x}px, ${y}px)`; };
  const resizeOverlay = printWindow.document.createElement('div');
  resizeOverlay.className = 'selection-resizer';
  resizeOverlay.innerHTML = ['n','e','s','w','ne','nw','se','sw'].map((edge) => `<span data-resize-edge="${edge}"></span>`).join('');
  resizeOverlay.hidden = true;
  printWindow.document.body.appendChild(resizeOverlay);
  const syncResizeOverlay = () => {
    if (!selectedBlock || !printWindow.document.body.classList.contains('layout-editing')) { resizeOverlay.hidden = true; return; }
    const rect = selectedBlock.getBoundingClientRect();
    resizeOverlay.hidden = false;
    Object.assign(resizeOverlay.style, { left: `${rect.left}px`, top: `${rect.top}px`, width: `${rect.width}px`, height: `${rect.height}px` });
  };
  const selectedTable = () => (selectedBlock ? (selectedBlock.tagName === 'TABLE' ? selectedBlock : selectedBlock.closest('table') || selectedBlock.querySelector('table')) : null);
  const refreshSelectedInfo = () => {
    const strong = printWindow.document.getElementById('layout-selected')?.querySelector('strong');
    if (strong) strong.textContent = selectedBlock ? blockLabel(selectedBlock) : '—';
  };
  const setColumnWidth = (table, index, value) => {
    const percent = clampNumber(value, 0, 100, '');
    Array.from(table.rows).forEach((tr) => { if (tr.cells.length > index) tr.cells[index].style.width = percent === '' ? '' : `${percent}%`; });
  };
  const setColumnAlign = (table, index, value) => {
    Array.from(table.rows).forEach((tr) => { if (tr.cells.length > index) tr.cells[index].style.textAlign = value; });
  };
  const columnBorderAt = (table, event) => {
    const headerRow = table.rows[0];
    if (!headerRow) return null;
    const cells = Array.from(headerRow.cells);
    for (let index = 0; index < cells.length - 1; index += 1) {
      const rect = cells[index].getBoundingClientRect();
      if (Math.abs(event.clientX - rect.right) <= COLUMN_RESIZE_HANDLE) return index;
    }
    return null;
  };
  const syncColumnWidthInput = (index) => {
    const table = selectedTable();
    if (!table) return;
    const wrap = printWindow.document.getElementById('layout-columns');
    const row = wrap?.children[index];
    const input = row?.querySelector('input[type=number]');
    if (!input) return;
    const cell = table.rows[0]?.cells[index];
    const width = cell?.style.width;
    input.value = width ? (Math.round(parseFloat(width)) || '') : '';
  };
  const startColumnResize = (table, borderIndex, event) => {
    const headerRow = table.rows[0];
    const cells = Array.from(headerRow.cells);
    const totalWidth = table.getBoundingClientRect().width || 1;
    const startWidths = cells.map((cell) => cell.getBoundingClientRect().width);
    table.style.tableLayout = 'fixed';
    cells.forEach((cell, index) => setColumnWidth(table, index, (startWidths[index] / totalWidth) * 100));
    resizeState = {
      table,
      borderIndex,
      pointerX: event.clientX,
      leftIndex: borderIndex,
      rightIndex: borderIndex + 1,
      leftStart: startWidths[borderIndex],
      rightStart: startWidths[borderIndex + 1] || 0,
      totalWidth,
    };
  };
  const applyColumnResize = (event) => {
    if (!resizeState) return;
    const min = 24;
    const delta = event.clientX - resizeState.pointerX;
    let left = resizeState.leftStart + delta;
    let right = resizeState.rightStart - delta;
    if (left < min) { right += left - min; left = min; }
    if (right < min) { left += right - min; right = min; }
    setColumnWidth(resizeState.table, resizeState.leftIndex, (left / resizeState.totalWidth) * 100);
    setColumnWidth(resizeState.table, resizeState.rightIndex, (right / resizeState.totalWidth) * 100);
    syncColumnWidthInput(resizeState.leftIndex);
    syncColumnWidthInput(resizeState.rightIndex);
  };
  const buildTableControls = () => {
    const doc = printWindow.document;
    const wrap = doc.getElementById('layout-table-controls');
    const table = selectedTable();
    if (!table) { wrap.hidden = true; doc.getElementById('layout-columns').innerHTML = ''; return; }
    wrap.hidden = false;
    const headerRow = table.rows[0];
    const columnCount = headerRow ? headerRow.cells.length : 0;
    doc.getElementById('layout-row-height').value = headerRow?.cells[0]?.style.height ? parseInt(headerRow.cells[0].style.height) || '' : '';
    const padCell = table.querySelector('td,th');
    doc.getElementById('layout-cell-padding').value = padCell?.style.padding ? parseInt(padCell.style.padding) || '' : '';
    const colsWrap = doc.getElementById('layout-columns');
    colsWrap.innerHTML = '';
    for (let index = 0; index < columnCount; index += 1) {
      const cell = headerRow.cells[index];
      const row = doc.createElement('div');
      row.className = 'layout-column';
      const label = doc.createElement('span');
      label.textContent = `Kolom ${index + 1}`;
      const widthInput = doc.createElement('input');
      widthInput.type = 'number'; widthInput.min = '0'; widthInput.step = '1'; widthInput.placeholder = 'lebar %';
      widthInput.value = cell.style.width ? parseInt(cell.style.width) || '' : '';
      const makeAlignSelect = (caption, current, onChange) => {
        const select = doc.createElement('select');
        select.setAttribute('aria-label', `${caption} kolom ${index + 1}`);
        [['', caption], ['left', 'Kiri'], ['center', 'Tengah'], ['right', 'Kanan']].forEach(([value, text]) => {
          const option = doc.createElement('option');
          option.value = value; option.textContent = text;
          if (current === value) option.selected = true;
          select.appendChild(option);
        });
        select.addEventListener('change', () => onChange(select.value));
        return select;
      };
      const effectiveAlign = (target) => {
        if (!target) return '';
        const value = target.style.textAlign || doc.defaultView.getComputedStyle(target).textAlign;
        return value === 'start' ? 'left' : value === 'end' ? 'right' : value;
      };
      const headerAlign = makeAlignSelect('Header', effectiveAlign(cell), (value) => { cell.style.textAlign = value; });
      const bodyCell = table.rows[1]?.cells[index];
      const bodyAlign = makeAlignSelect('Isi', effectiveAlign(bodyCell), (value) => {
        Array.from(table.rows).slice(1).forEach((tr) => { if (tr.cells[index]) tr.cells[index].style.textAlign = value; });
      });
      widthInput.addEventListener('input', () => setColumnWidth(table, index, widthInput.value));
      row.append(label, widthInput, headerAlign, bodyAlign);
      colsWrap.appendChild(row);
    }
  };
  const syncInspectorFromBlock = () => {
    const doc = printWindow.document;
    if (!selectedBlock) return;
    const { x, y } = readTranslate(selectedBlock);
    doc.getElementById('layout-x').value = x;
    doc.getElementById('layout-y').value = y;
    doc.getElementById('layout-width').value = selectedBlock.style.width ? Math.round(parseFloat(selectedBlock.style.width)) || '' : '';
    doc.getElementById('layout-height').value = selectedBlock.style.height ? Math.round(parseFloat(selectedBlock.style.height)) || '' : '';
    doc.getElementById('layout-font-size').value = selectedBlock.style.fontSize ? parseFloat(selectedBlock.style.fontSize) || '' : '';
    doc.getElementById('layout-line-height').value = selectedBlock.style.lineHeight || '';
    const effectiveTextAlign = selectedBlock.style.textAlign || doc.defaultView.getComputedStyle(selectedBlock).textAlign;
    const normalizedTextAlign = effectiveTextAlign === 'start' ? 'left' : effectiveTextAlign === 'end' ? 'right' : effectiveTextAlign;
    doc.querySelectorAll('#layout-align button').forEach((btn) => btn.classList.toggle('is-active', btn.dataset.align === normalizedTextAlign));
    buildTableControls();
  };
  const selectBlock = (el) => {
    if (selectedBlock) selectedBlock.classList.remove('is-selected');
    selectedBlock = el || null;
    if (selectedBlock) selectedBlock.classList.add('is-selected');
    refreshSelectedInfo();
    syncInspectorFromBlock();
    const tableButton = printWindow.document.getElementById('layout-select-table');
    tableButton.hidden = !selectedBlock || !selectedBlock.closest('table') || selectedBlock.tagName === 'TABLE';
    const textControl = printWindow.document.getElementById('layout-text-control');
    textControl.hidden = !isPlainTextBlock(selectedBlock);
    if (!textControl.hidden) printWindow.document.getElementById('layout-text').value = readPlainText(selectedBlock);
    syncResizeOverlay();
  };
  const initLayoutBlocks = () => {
    const sheet = printWindow.document.getElementById('print-sheet');
    if (!sheet) return;
    const topLevel = Array.from(sheet.children).filter((child) => child.nodeType === 1);
    topLevel.forEach((child, index) => {
      child.setAttribute('data-layout-block', '');
      child.setAttribute('data-lid', String(index));
      child.setAttribute('tabindex', '0');
    });
    let nextId = topLevel.length;
    sheet.querySelectorAll('p, h1, h2, h3, td, th').forEach((child) => {
      if (child.hasAttribute('data-layout-block')) return;
      child.setAttribute('data-layout-block', '');
      child.setAttribute('data-lid', String(nextId++));
      child.setAttribute('tabindex', '0');
    });
  };
  const isPlainTextBlock = (el) => el && !el.querySelector('table') && Array.from(el.children).every((child) => child.tagName === 'BR');
  const readPlainText = (el) => (el.innerText || '').replace(/\n$/, '');
  const setPlainText = (el, value) => {
    el.replaceChildren();
    String(value).split(/\r?\n/).forEach((line, index) => {
      if (index) el.appendChild(printWindow.document.createElement('br'));
      el.appendChild(printWindow.document.createTextNode(line));
    });
  };
  initLayoutBlocks();
  const defaultBlockText = new Map(Array.from(printWindow.document.querySelectorAll('#print-sheet [data-layout-block]'))
    .filter(isPlainTextBlock).map((el) => [el.getAttribute('data-lid'), readPlainText(el)]));
  const captureLayout = () => {
    const doc = printWindow.document;
    const blocks = {};
    doc.querySelectorAll('#print-sheet [data-layout-block]').forEach((el) => {
      const { x, y } = readTranslate(el);
      const entry = {};
      if (el.style.textAlign) entry.textAlign = el.style.textAlign;
      if (el.style.fontSize) entry.fontSize = el.style.fontSize;
      if (el.style.lineHeight) entry.lineHeight = el.style.lineHeight;
      if (el.style.width) entry.width = el.style.width;
      if (el.style.height) entry.height = el.style.height;
      if (isPlainTextBlock(el)) entry.text = readPlainText(el);
      if (x || y) { entry.x = x; entry.y = y; }
      const table = el.tagName === 'TABLE' ? el : el.querySelector('table');
      if (table) {
        const headerRow = table.rows[0];
        const columns = [];
        if (headerRow) Array.from(headerRow.cells).forEach((cell, index) => { columns[index] = { width: cell.style.width || '', headerAlign: cell.style.textAlign || '', bodyAlign: table.rows[1]?.cells[index]?.style.textAlign || '' }; });
        const padCell = table.querySelector('td,th');
        entry.table = { columns, rowHeight: headerRow?.cells[0]?.style.height || '', cellPadding: padCell?.style.padding || '', tableLayout: table.style.tableLayout || '' };
      }
      if (Object.keys(entry).length) blocks[el.getAttribute('data-lid')] = entry;
    });
    const page = {
      size: doc.getElementById('print-page-size').value,
      customWidth: doc.getElementById('print-custom-width').value,
      customHeight: doc.getElementById('print-custom-height').value,
      orientation: doc.getElementById('print-orientation').value,
      marginTop: doc.getElementById('print-margin-top').value,
      marginRight: doc.getElementById('print-margin-right').value,
      marginBottom: doc.getElementById('print-margin-bottom').value,
      marginLeft: doc.getElementById('print-margin-left').value,
      font: doc.getElementById('print-font').value,
      fontSize: doc.getElementById('print-font-size').value,
      lineHeight: doc.getElementById('print-line-height').value,
    };
    return { page, blocks };
  };
  const layoutOnly = (layout) => {
    if (!layout || typeof layout !== 'object') return null;
    const sanitized = structuredClone(layout);
    Object.values(sanitized.blocks || {}).forEach((entry) => { if (entry && typeof entry === 'object') delete entry.text; });
    return sanitized;
  };
  const applyPageSettings = (page) => {
    const doc = printWindow.document;
    if (!page) return;
    const set = (id, value) => { if (value != null && value !== '') doc.getElementById(id).value = value; };
    set('print-page-size', page.size);
    set('print-custom-width', page.customWidth);
    set('print-custom-height', page.customHeight);
    set('print-orientation', page.orientation);
    set('print-margin-top', page.marginTop);
    set('print-margin-right', page.marginRight);
    set('print-margin-bottom', page.marginBottom);
    set('print-margin-left', page.marginLeft);
    set('print-font', page.font);
    set('print-font-size', page.fontSize);
    set('print-line-height', page.lineHeight);
  };
  const applyBlockLayout = (blocks) => {
    printWindow.document.querySelectorAll('#print-sheet [data-layout-block]').forEach((el) => {
      const entry = blocks[el.getAttribute('data-lid')];
      if (!entry) return;
      if (entry.textAlign) el.style.textAlign = entry.textAlign;
      if (entry.fontSize) el.style.fontSize = entry.fontSize;
      if (entry.lineHeight) el.style.lineHeight = entry.lineHeight;
      if (entry.width) el.style.width = entry.width;
      if (entry.height) el.style.height = entry.height;
      if (entry.text !== undefined && isPlainTextBlock(el)) setPlainText(el, entry.text);
      if (entry.x || entry.y) applyTranslate(el, entry.x || 0, entry.y || 0);
      if (entry.table) {
        const table = el.tagName === 'TABLE' ? el : el.querySelector('table');
        if (table) {
          if (entry.table.tableLayout) table.style.tableLayout = entry.table.tableLayout;
          if (entry.table.rowHeight) Array.from(table.rows).forEach((tr) => Array.from(tr.cells).forEach((cell) => { cell.style.height = entry.table.rowHeight; }));
          if (entry.table.cellPadding) table.querySelectorAll('td,th').forEach((cell) => { cell.style.padding = entry.table.cellPadding; });
          (entry.table.columns || []).forEach((column, index) => {
            if (!column) return;
            Array.from(table.rows).forEach((tr) => {
              const cell = tr.cells[index];
              if (!cell || tr.cells.length <= index) return;
              if (column.width) cell.style.width = column.width;
              if (column.headerAlign && tr === table.rows[0]) cell.style.textAlign = column.headerAlign;
              else if (column.bodyAlign && tr !== table.rows[0]) cell.style.textAlign = column.bodyAlign;
              else if (column.align) cell.style.textAlign = column.align;
            });
          });
        }
      }
    });
  };
  printWindow.document.getElementById('save-layout').addEventListener('click', () => {
    const store = readLayoutStore();
    store.defaults = store.defaults || {};
    store.defaults[type] = layoutOnly(captureLayout());
    writeLayoutStore(store);
    showToast('Layout ini dijadikan default untuk semua dokumen sejenis.');
  });
  printWindow.document.getElementById('reset-layout').addEventListener('click', () => {
    const store = readLayoutStore();
    if (store.defaults) delete store.defaults[type];
    delete store[type];
    writeLayoutStore(store);
    printWindow.document.querySelectorAll('#print-sheet [data-layout-block]').forEach((el) => {
      el.style.textAlign = '';
      el.style.fontSize = '';
      el.style.lineHeight = '';
      el.style.width = '';
      el.style.height = '';
      el.style.transform = '';
      const defaultText = defaultBlockText.get(el.getAttribute('data-lid'));
      if (defaultText !== undefined && isPlainTextBlock(el)) setPlainText(el, defaultText);
      const table = el.tagName === 'TABLE' ? el : el.querySelector('table');
      if (table) { table.style.tableLayout = ''; table.querySelectorAll('td,th').forEach((cell) => { cell.style.width = ''; cell.style.textAlign = ''; cell.style.height = ''; cell.style.padding = ''; }); }
    });
    selectBlock(null);
    const doc = printWindow.document;
    doc.getElementById('print-page-size').value = 'A4';
    doc.getElementById('print-orientation').value = orientation;
    ['print-margin-top', 'print-margin-right', 'print-margin-bottom', 'print-margin-left'].forEach((id) => { doc.getElementById(id).value = '12.7'; });
    doc.getElementById('print-font').value = 'Cambria, serif';
    doc.getElementById('print-font-size').value = defaultPrintFontSize;
    doc.getElementById('print-line-height').value = '1.3';
    updatePrintSettings();
    showToast('Layout dikembalikan ke pengaturan awal.');
  });
  const storedLayouts = readLayoutStore();
  const savedLayout = layoutOnly(storedLayouts.defaults?.[type] || storedLayouts[type]);
  if (savedLayout) {
    applyPageSettings(savedLayout.page);
    applyBlockLayout(savedLayout.blocks || {});
  }

  const editButton = printWindow.document.getElementById('edit-layout');
  editButton.addEventListener('click', () => {
    const body = printWindow.document.body;
    const active = body.classList.toggle('layout-editing');
    editButton.classList.toggle('is-active', active);
    editButton.textContent = active ? 'Selesai Edit' : 'Edit Layout';
    if (!active) { selectBlock(null); resizeState = null; printWindow.document.body.classList.remove('col-resize-hover', 'col-resizing'); }
  });
  printWindow.document.addEventListener('keydown', (event) => {
    if (!printWindow.document.body.classList.contains('layout-editing') || !selectedBlock) return;
    const inToolbar = Boolean(event.target.closest('.print-toolbar'));
    const editingText = event.target.id === 'layout-text';
    const { x, y } = readTranslate(selectedBlock);
    if (event.key === 'Tab' && (!inToolbar || editingText)) {
      event.preventDefault();
      const step = event.shiftKey ? -10 : 10;
      applyTranslate(selectedBlock, x + step, y);
      printWindow.document.getElementById('layout-x').value = x + step;
      syncResizeOverlay();
    } else if (!inToolbar && ['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(event.key)) {
      event.preventDefault();
      const step = event.shiftKey ? 10 : 1;
      const dx = event.key === 'ArrowLeft' ? -step : event.key === 'ArrowRight' ? step : 0;
      const dy = event.key === 'ArrowUp' ? -step : event.key === 'ArrowDown' ? step : 0;
      applyTranslate(selectedBlock, x + dx, y + dy);
      printWindow.document.getElementById('layout-x').value = x + dx;
      printWindow.document.getElementById('layout-y').value = y + dy;
      syncResizeOverlay();
    }
  });
  printWindow.document.getElementById('layout-editor-details').addEventListener('toggle', (event) => {
    printWindow.document.body.classList.toggle('layout-panel-open', event.target.open);
  });
  printWindow.document.getElementById('layout-select-table').addEventListener('click', () => {
    const table = selectedBlock?.closest('table');
    if (table) selectBlock(table);
  });
  printWindow.document.getElementById('layout-text').addEventListener('input', (event) => {
    if (isPlainTextBlock(selectedBlock)) setPlainText(selectedBlock, event.target.value);
  });
  let elementResizeState = null;
  printWindow.document.addEventListener('mousedown', (event) => {
    if (!printWindow.document.body.classList.contains('layout-editing')) return;
    const edgeHandle = event.target.closest('[data-resize-edge]');
    if (edgeHandle && selectedBlock) {
      const rect = selectedBlock.getBoundingClientRect();
      elementResizeState = { block: selectedBlock, edge: edgeHandle.dataset.resizeEdge, pointerX: event.clientX, pointerY: event.clientY, width: rect.width, height: rect.height, ...readTranslate(selectedBlock) };
      event.preventDefault();
      return;
    }
    const resizeTable = event.target.closest('#print-sheet table');
    if (resizeTable) {
      const border = columnBorderAt(resizeTable, event);
      if (border !== null) {
        startColumnResize(resizeTable, border, event);
        const block = resizeTable.closest('[data-layout-block]');
        if (block) selectBlock(block);
        printWindow.document.body.classList.add('col-resizing');
        event.preventDefault();
        return;
      }
    }
    const block = event.target.closest('[data-layout-block]');
    if (!block) return;
    selectBlock(block);
    const start = readTranslate(block);
    dragState = { block, pointerX: event.clientX, pointerY: event.clientY, startX: start.x, startY: start.y };
    event.preventDefault();
  });
  printWindow.document.addEventListener('mousemove', (event) => {
    if (elementResizeState) {
      const state = elementResizeState;
      const dx = event.clientX - state.pointerX;
      const dy = event.clientY - state.pointerY;
      const west = state.edge.includes('w');
      const north = state.edge.includes('n');
      const width = Math.max(40, state.width + (west ? -dx : state.edge.includes('e') ? dx : 0));
      const height = Math.max(24, state.height + (north ? -dy : state.edge.includes('s') ? dy : 0));
      state.block.style.width = `${width}px`;
      state.block.style.height = `${height}px`;
      if (west || north) applyTranslate(state.block, state.x + (west ? dx : 0), state.y + (north ? dy : 0));
      printWindow.document.getElementById('layout-width').value = Math.round(width);
      printWindow.document.getElementById('layout-height').value = Math.round(height);
      const moved = readTranslate(state.block);
      printWindow.document.getElementById('layout-x').value = moved.x;
      printWindow.document.getElementById('layout-y').value = moved.y;
      syncResizeOverlay();
      return;
    }
    if (resizeState) { applyColumnResize(event); return; }
    if (dragState) {
      const x = Math.round(dragState.startX + (event.clientX - dragState.pointerX));
      const y = Math.round(dragState.startY + (event.clientY - dragState.pointerY));
      applyTranslate(dragState.block, x, y);
      syncResizeOverlay();
      printWindow.document.getElementById('layout-x').value = x;
      printWindow.document.getElementById('layout-y').value = y;
      return;
    }
    if (!printWindow.document.body.classList.contains('layout-editing')) return;
    const table = event.target.closest('#print-sheet table');
    const onBorder = table ? columnBorderAt(table, event) !== null : false;
    printWindow.document.body.classList.toggle('col-resize-hover', onBorder);
  });
  printWindow.document.addEventListener('mouseup', () => { dragState = null; resizeState = null; elementResizeState = null; printWindow.document.body.classList.remove('col-resizing'); });
  printWindow.addEventListener('scroll', syncResizeOverlay, { passive: true });
  printWindow.addEventListener('resize', syncResizeOverlay);
  printWindow.document.getElementById('layout-align').addEventListener('click', (event) => {
    const button = event.target.closest('button[data-align]');
    if (!button || !selectedBlock) return;
    selectedBlock.style.textAlign = button.dataset.align;
    printWindow.document.querySelectorAll('#layout-align button').forEach((btn) => btn.classList.toggle('is-active', btn === button));
  });
  printWindow.document.getElementById('layout-font-size').addEventListener('input', (event) => {
    if (selectedBlock) selectedBlock.style.fontSize = event.target.value ? `${event.target.value}pt` : '';
  });
  printWindow.document.getElementById('layout-line-height').addEventListener('input', (event) => {
    if (selectedBlock) selectedBlock.style.lineHeight = event.target.value || '';
  });
  printWindow.document.getElementById('layout-y').addEventListener('input', (event) => {
    if (!selectedBlock) return;
    const { x } = readTranslate(selectedBlock);
    applyTranslate(selectedBlock, x, Number(event.target.value) || 0);
    syncResizeOverlay();
  });
  printWindow.document.getElementById('layout-x').addEventListener('input', (event) => {
    if (!selectedBlock) return;
    const { y } = readTranslate(selectedBlock);
    applyTranslate(selectedBlock, Number(event.target.value) || 0, y);
    syncResizeOverlay();
  });
  printWindow.document.getElementById('layout-width').addEventListener('input', (event) => {
    if (!selectedBlock) return;
    selectedBlock.style.width = event.target.value ? `${Math.max(40, Number(event.target.value))}px` : '';
    syncResizeOverlay();
  });
  printWindow.document.getElementById('layout-height').addEventListener('input', (event) => {
    if (!selectedBlock) return;
    selectedBlock.style.height = event.target.value ? `${Math.max(24, Number(event.target.value))}px` : '';
    syncResizeOverlay();
  });
  printWindow.document.getElementById('layout-row-height').addEventListener('input', (event) => {
    const table = selectedTable();
    if (!table) return;
    const value = event.target.value ? `${event.target.value}px` : '';
    Array.from(table.rows).forEach((tr) => Array.from(tr.cells).forEach((cell) => { cell.style.height = value; }));
  });
  printWindow.document.getElementById('layout-cell-padding').addEventListener('input', (event) => {
    const table = selectedTable();
    if (!table) return;
    const value = event.target.value ? `${event.target.value}px` : '';
    table.querySelectorAll('td,th').forEach((cell) => { cell.style.padding = value; });
  });

  ['print-page-size', 'print-custom-width', 'print-custom-height', 'print-orientation', 'print-margin-top', 'print-margin-right', 'print-margin-bottom', 'print-margin-left', 'print-font', 'print-font-size', 'print-line-height'].forEach((id) => {
    printWindow.document.getElementById(id).addEventListener('input', updatePrintSettings);
  });
  updatePrintSettings();
  printWindow.document.getElementById('print-now').addEventListener('click', () => {
    printWindow.focus();
    printWindow.print();
  });
  printWindow.document.getElementById('close-print-preview').addEventListener('click', () => printWindow.close());
  const printStyle = printWindow.document.createElement('style');
  printStyle.textContent = '.sppd-signature-grid{display:flex;justify-content:flex-end;margin-top:20px}.sppd-sign-box{width:320px;text-align:left}.sppd-sign-box p{margin:4px 0}.sppd-sign-label{margin-top:12px!important}.sppd-route-line{margin:4px 0}.sppd-note{margin:8px 0;line-height:1.5}.continuation-sheet{width:100%;border-collapse:collapse;table-layout:fixed;font-size:11pt;line-height:1.25}.continuation-sheet td,.continuation-sheet th{width:50%;border:1px solid #222;padding:5px 7px;vertical-align:top}.continuation-sheet p{margin:2px 0}.continuation-top td{height:100px}.continuation-heading{height:30px;text-align:center;vertical-align:middle!important;font-weight:normal}.continuation-travel-row td{height:155px}.continuation-signature{height:52px}.continuation-signature-name{text-align:center;margin-top:0!important}.continuation-return td{height:190px}.continuation-return .return-date{margin-top:26px}.continuation-return .continuation-signature{height:42px}.continuation-official{margin-top:0!important}.continuation-notes td{height:42px;vertical-align:middle}.continuation-notes td:first-child{border-right:1px solid #222}.attention-heading{height:27px}.continuation-attention{line-height:1.35}.transport-date-row{display:grid;grid-template-columns:1fr auto;gap:12px;margin:18px 0 10px;font-weight:bold}.transport-approval{display:grid;grid-template-columns:1fr 1fr;gap:30px;margin-top:18px;page-break-inside:avoid}.transport-approval>div{min-width:0}.transport-approval p{margin:4px 0}.transport-approval .signature-space{height:60px}.approval-box p{margin:4px 0}';
  printWindow.document.head.appendChild(printStyle);
  const sppdLayoutStyle = printWindow.document.createElement('style');
  sppdLayoutStyle.textContent = '.sppd-identifiers-grid,.sppd-table,.sppd-follower-table{width:100%;border-collapse:collapse;table-layout:fixed}.sppd-identifiers-grid{margin:0 0 8px}.sppd-identifiers-grid col:first-child{width:57%}.sppd-identifiers-grid col:nth-child(2){width:105px}.sppd-identifiers-grid col:nth-child(3){width:14px}.sppd-identifiers-grid td,.sppd-table>tbody>tr>td,.sppd-follower-table th,.sppd-follower-table td{border:1px solid #222;padding:3px 5px;vertical-align:top}.sppd-identifiers-grid td{height:18px}.sppd-table>tbody>tr>td:first-child{width:57%}.sppd-table>tbody>tr>td{min-height:24px}.sppd-follower-table th,.sppd-follower-table td{text-align:center}.sppd-follower-table th{font-weight:normal}.sppd-signature-grid{margin-top:6px}.sppd-sign-box{width:43%;min-width:300px}.sppd-sign-box .signature-space{height:57px}';
  printWindow.document.head.appendChild(sppdLayoutStyle);
  printWindow.document.close();
}

function printArchive(kind = 'task') {
  const cityLogoSrc = new URL('assets/logo-kota.jpeg', window.location.href).href;
  const schoolLogoSrc = new URL('assets/logo-smpn13.png', window.location.href).href;
  const tables = [
    ['ARSIP SURAT MASUK', ['No','Tanggal Terima','Nomor Surat','Tanggal Surat','Pengirim','Perihal','Disposisi'], [...state.incoming].sort((a,b)=>(a.receivedDate||'').localeCompare(b.receivedDate||'')).map((r,i)=>[i+1,formatDate(r.receivedDate),r.letterNumber,r.letterDate?formatDate(r.letterDate):'',r.sender,r.subject,r.disposition])],
    ['ARSIP SURAT KELUAR', ['No','Tanggal Keluar','Nomor Surat','Jenis','Tujuan','Perihal','Keterangan'], [...state.outgoing].sort((a,b)=>(a.issuedDate||'').localeCompare(b.issuedDate||'')).map((r,i)=>[i+1,formatDate(r.issuedDate),r.number,r.letterType,r.recipient,r.subject,r.notes])],
    ['ARSIP SURAT TUGAS', ['No','Tanggal','Nomor Surat','Perihal','Penerima','Kegiatan','Tempat'], [...state.tasks].sort((a,b)=>(a.issuedDate||'').localeCompare(b.issuedDate||'')).map((r,i)=>[i+1,formatDate(r.issuedDate),r.number,r.subject,r.recipient,r.eventName,r.place])],
  ];
  const table = tables[{ incoming: 0, outgoing: 1, task: 2 }[kind] ?? 2];
  const [title, headers, rows] = table;
  const section = `<table><thead><tr>${headers.map(h=>`<th>${h}</th>`).join('')}</tr></thead><tbody>${rows.map(row=>`<tr>${row.map(cell=>`<td>${escapeHtml(cell || '')}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
  const officials = state.officials;
  const printWindow = window.open('', '_blank', 'width=1100,height=750');
  if (!printWindow) {
    showToast('Izinkan pop-up untuk mencetak arsip.');
    return;
  }
  const archiveOfficer = state.officialsMaster.find((person) => person.role === 'Petugas') || state.officialsMaster[0] || {};
  printWindow.document.write(`<!doctype html><html lang="id"><head><meta charset="utf-8"><title>${title}</title><style>@page{size:A4 landscape;margin:12mm}body{font:10px Arial,sans-serif;color:#111}.kop{display:grid;grid-template-columns:70px 1fr 70px;align-items:center;text-align:center;border-bottom:3px double #111;padding-bottom:8px}.kop img{width:60px;height:60px;object-fit:contain}.kop strong{display:block;font-size:14px}.kop p{margin:3px}h1{text-align:center;font-size:16px}table{width:100%;border-collapse:collapse}th,td{border:1px solid #333;padding:4px;text-align:left;vertical-align:top}th{background:#eef2ee}.signatures{display:flex;justify-content:flex-end;gap:45px;margin:34px 8mm 0 0;text-align:center}.signature{width:190px}.space{height:55px}@media print{tr{break-inside:avoid}}</style></head><body><header class="kop"><img src="${cityLogoSrc}" alt=""><div><strong>PEMERINTAH KOTA TASIKMALAYA</strong><strong>SMP NEGERI 13 TASIKMALAYA</strong><p>Jl. Letjen H. Ibrahim Adjie Km.2 Indihiang Tasikmalaya 46151 Telp. (0265) 335695</p></div><img src="${schoolLogoSrc}" alt=""></header><h1>${title}</h1>${section}<div class="signatures"><div class="signature">Mengetahui,<br>Kepala Sekolah<div class="space"></div>${escapeHtml(officials.principalName || '')}<br>NIP. ${escapeHtml(officials.principalNip || '')}</div><div class="signature">${escapeHtml(archiveOfficer.role || 'Petugas')}<div class="space"></div>${escapeHtml(archiveOfficer.name || '')}<br>${archiveOfficer.nip ? `NIP. ${escapeHtml(archiveOfficer.nip)}` : ''}</div></div><script>window.onload=()=>window.print()<\\/script></body></html>`);
  printWindow.document.close();
}

function handleAction(action, element) {
  if (action === 'new-task') openTaskForm();
  if (action === 'edit-task') openTaskForm(element.dataset.id);
  if (action === 'download-task') downloadTask(element.dataset.id);
  if (action === 'new-incoming') openIncomingForm();
  if (action === 'new-outgoing') openOutgoingForm();
  if (action === 'edit-outgoing') openOutgoingForm(element.dataset.id);
  if (action === 'close-dialog') document.getElementById('record-dialog').close();
  if (action === 'close-print-dialog') document.getElementById('print-dialog').close();
  if (action === 'print-options') openPrintOptions(element.dataset.id);
  if (action === 'print-archive') printArchive(element.dataset.archiveType);
  if (action === 'open-incoming-file') {
    const record = state.incoming.find((row) => row.id === element.dataset.id);
    try { const file = JSON.parse(record.attachment); const link = document.createElement('a'); link.href = file.data; link.target = '_blank'; link.rel = 'noopener'; link.download = file.name; link.click(); } catch (_) { showToast('Berkas tidak dapat dibuka.'); }
  }
  if (action === 'delete-official') { state.officialsMaster.splice(Number(element.dataset.id), 1); saveState(); }
  if (action === 'download-teacher-template') downloadRosterTemplate('teachers');
  if (action === 'download-student-template') downloadRosterTemplate('students');
  if (action === 'export-agenda') exportAgenda();
  if (action === 'export-backup') exportBackup();
  if (action === 'connect-spreadsheet') connectStorageToSpreadsheet();
  if (action === 'migrate-to-spreadsheet') migrateToSpreadsheet();
  if (action === 'reconnect-spreadsheet') reconnectSpreadsheet();
  if (action === 'test-spreadsheet') testSpreadsheetConnection();
}

function importBackup(file) {
  (async () => {
    try {
      if (!file || !/\.json$/i.test(file.name || '') || file.size > 20 * 1024 * 1024) throw new Error('Pilih file JSON maksimal 20 MB.');
      const imported = JSON.parse(await file.text());
      const keys = ['tasks', 'incoming', 'outgoing', 'teachers', 'students'];
      if (!imported || typeof imported !== 'object' || Array.isArray(imported) || !Array.isArray(imported.tasks) || !Array.isArray(imported.incoming)
        || keys.some((key) => imported[key] !== undefined && (!Array.isArray(imported[key]) || imported[key].some((row) => !row || typeof row !== 'object' || Array.isArray(row))))) {
        throw new Error('Struktur cadangan tidak valid. Data lama tidak diubah.');
      }
      const next = { ...state };
      for (const key of keys) if (Array.isArray(imported[key])) {
        next[key] = mergeRows(state[key], imported[key].map((row) => ({ ...row, id: String(row.id || crypto.randomUUID()) })));
      }
      if (imported.officials && typeof imported.officials === 'object' && !Array.isArray(imported.officials)) next.officials = { ...state.officials, ...imported.officials };
      if (imported.officialHistory && typeof imported.officialHistory === 'object' && !Array.isArray(imported.officialHistory)) {
        next.officialHistory = {
          principal: Array.isArray(imported.officialHistory.principal) ? imported.officialHistory.principal : state.officialHistory.principal,
          treasurer: Array.isArray(imported.officialHistory.treasurer) ? imported.officialHistory.treasurer : state.officialHistory.treasurer,
        };
      }
      if (imported.lastIssuedByYear && typeof imported.lastIssuedByYear === 'object' && !Array.isArray(imported.lastIssuedByYear)) next.lastIssuedByYear = { ...state.lastIssuedByYear, ...imported.lastIssuedByYear };
      state = next;
      saveState();
      showToast(`${imported.tasks.length} arsip dari cadangan berhasil digabungkan.`);
    } catch (error) {
      showToast(`Cadangan tidak diimpor: ${error.message}`);
    }
  })();
}

function migrateToSpreadsheet() {
  if (!API_BASE_URL || API_BASE_URL.includes('{WEB_APP_ID}')) {
    showToast('Masukkan URL deployment Web App Apps Script yang sebenarnya.');
    return;
  }
  const localSnapshot = structuredClone(loadState());
  connectStorageToSpreadsheet().then(async (connected) => {
    if (!connected) return;
    try {
      const rows = Object.fromEntries(['tasks', 'incoming', 'outgoing', 'teachers', 'students'].map((key) => [key, localSnapshot[key] || []]));
      rows.settings = [{ id: 'application', officials: localSnapshot.officials || {}, officialHistory: localSnapshot.officialHistory || {}, officialsMaster: localSnapshot.officialsMaster || [], updatedAt: new Date().toISOString() }];
      await writeAllToSheet(rows);
      showToast('Data lokal berhasil digabungkan ke Spreadsheet; data yang sudah ada tetap dipertahankan.');
    } catch (error) { showToast(`Migrasi gagal: ${error.message}`); }
  });
}

function reconnectSpreadsheet() {
  if (!API_BASE_URL) {
    showToast('Isi API URL di pengaturan Google Spreadsheet terlebih dahulu.');
    return;
  }
  connectStorageToSpreadsheet().then((connected) => {
    if (!connected) return;
    showToast('Koneksi ulang ke Spreadsheet berhasil.');
  });
}

function getStorageRecordCount() {
  let total = 0;
  for (const key of ['tasks', 'incoming', 'outgoing', 'teachers', 'students']) {
    total += Array.isArray(state[key]) ? state[key].length : 0;
  }
  return total;
}

function countTable(key) {
  return Array.isArray(state[key]) ? state[key].length : 0;
}


async function testSpreadsheetConnection() {
  if (!API_BASE_URL || API_BASE_URL.includes('{WEB_APP_ID}')) {
    showToast('Masukkan URL deployment Web App Apps Script yang sebenarnya.');
    return;
  }
  try {
    const cfg = await fetchConfig();
    if (!cfg || !Array.isArray(cfg.tables)) {
      throw new Error(cfg?.error || 'Konfigurasi tabel tidak tersedia.');
    }
    const rows = await readAllFromSheet();
    const tableCounts = Object.fromEntries(Object.entries(rows).map(([key, value]) => [key, value.length]));
    if (cfg.capabilities?.safeIntegrationTest !== true) {
      const message = `Baca berhasil (${cfg.tables.length} tabel), tetapi uji tulis dilewati: deployment belum mengiklankan penghapusan baris uji yang aman. Perbarui source Apps Script lalu update deployment yang ada.`;
      document.getElementById('spreadsheet-status').textContent = message;
      showToast(message);
      return;
    }

    const id = `agenda-integration-test-${crypto.randomUUID()}`;
    const testRow = {
      id,
      receivedDate: localDateValue(),
      letterDate: '',
      letterNumber: id,
      sender: 'Agenda Surat Integration Test',
      subject: `[UJI INTEGRASI AGENDA] ${id}`,
      disposition: 'AUTO DELETE TEST ROW',
      attachment: '',
      createdAt: new Date().toISOString(),
    };
    let writeAttempted = false;
    let writeVerified = false;
    let cleanupError = null;
    let writeError = null;
    try {
      writeAttempted = true;
      await appendIntegrationTestRow(testRow);
      const afterWrite = await readAllFromSheet();
      writeVerified = afterWrite.incoming.some((row) => row.id === id);
      if (!writeVerified) throw new Error('Baris uji tidak ditemukan setelah ditulis.');
    } catch (error) {
      writeError = error;
    } finally {
      if (writeAttempted) {
        try {
          let deletionError = null;
          try {
            const deletion = await deleteIntegrationTestRow(id);
            if (!deletion?.ok) deletionError = new Error(deletion?.error || 'Backend menolak penghapusan baris uji.');
          } catch (error) { deletionError = error; }
          const afterDelete = await readAllFromSheet();
          if (afterDelete.incoming.some((row) => row.id === id)) {
            throw deletionError || new Error('Baris uji masih ditemukan setelah permintaan hapus.');
          }
        } catch (error) { cleanupError = error; }
      }
    }
    if (cleanupError) throw new Error(`Baca berhasil; tulis ${writeVerified ? 'terverifikasi' : 'dicoba'}, tetapi pembersihan baris uji ${id} gagal: ${cleanupError.message}. Hapus hanya baris dengan ID tersebut.`);
    if (writeError) throw writeError;
    const counts = `Tasks ${tableCounts.tasks}, Incoming ${tableCounts.incoming}, Outgoing ${tableCounts.outgoing}, Masters ${tableCounts.teachers + tableCounts.students}`;
    const message = `Baca ${cfg.tables.length} tabel berhasil (${counts}); tulis, baca balik, dan hapus baris uji berhasil.`;
    document.getElementById('spreadsheet-status').textContent = message;
    showToast(message);
  } catch (error) {
    document.getElementById('spreadsheet-status').textContent = `Koneksi gagal: ${error.message}`;
    showToast(`Uji koneksi gagal: ${error.message}`);
  }
}


document.addEventListener('click', (event) => {
  const addAssignment = event.target.closest('[data-add-assignment]');
  if (addAssignment) {
    const type = addAssignment.dataset.addAssignment;
    const assignments = currentAssignmentRows(type);
    assignments.push({});
    renderAssignmentRows(type, assignments);
    return;
  }
  const removeAssignment = event.target.closest('[data-remove-assignment]');
  if (removeAssignment) {
    const type = removeAssignment.dataset.removeAssignment;
    const rows = currentAssignmentRows(type);
    const row = removeAssignment.closest('[data-assignment-row]');
    const index = [...document.querySelectorAll(`[data-assignment-row="${type}"]`)].indexOf(row);
    rows.splice(index, 1);
    renderAssignmentRows(type, rows.length ? rows : [{}]);
    return;
  }
  const sppdPersonButton = event.target.closest('[data-sppd-person]');
  if (sppdPersonButton) {
    printDocument(activePrintTaskId, activeSppdPrintType, sppdPersonButton.dataset.sppdPerson);
    return;
  }
  const backToPrintOptions = event.target.closest('[data-back-print-options]');
  if (backToPrintOptions) {
    openPrintOptions(activePrintTaskId);
    return;
  }
  const printButton = event.target.closest('[data-print-type]');
  if (printButton) {
    if (printButton.dataset.printType === 'sppd' || printButton.dataset.printType === 'sppd-continuation') {
      openSppdPersonOptions(printButton.dataset.printType);
      return;
    }
    printDocument(activePrintTaskId, printButton.dataset.printType);
    return;
  }
  const pageButton = event.target.closest('[data-page]');
  if (pageButton) {
    event.preventDefault();
    navigate(pageButton.dataset.page);
    return;
  }
  const actionButton = event.target.closest('[data-action]');
  if (actionButton) handleAction(actionButton.dataset.action, actionButton);
  const agendaTab = event.target.closest('[data-agenda-tab]');
  if (agendaTab) {
    activeAgendaTab = agendaTab.dataset.agendaTab;
    renderAgenda();
  }
});

document.getElementById('record-form').addEventListener('submit', (event) => {
  event.preventDefault();
  if (event.currentTarget.dataset.kind === 'task') saveTask(event.currentTarget);
  if (event.currentTarget.dataset.kind === 'incoming') saveIncoming(event.currentTarget);
  if (event.currentTarget.dataset.kind === 'outgoing') saveOutgoing(event.currentTarget);
});

document.getElementById('officials-form').addEventListener('submit', (event) => {
  event.preventDefault();
  const values = collectFormData(event.currentTarget);
  saveOfficialHistory('principal', values.principalName.trim(), values.principalNip.trim(), values.principalStartDate);
  saveOfficialHistory('treasurer', values.treasurerName.trim(), values.treasurerNip.trim(), values.treasurerStartDate);
  state.officials = {
    principalName: values.principalName.trim(),
    principalNip: values.principalNip.trim(),
    principalStartDate: values.principalStartDate,
    treasurerName: values.treasurerName.trim(),
    treasurerNip: values.treasurerNip.trim(),
    treasurerStartDate: values.treasurerStartDate,
  };
  saveState();
  showToast('Pejabat penandatangan berhasil disimpan.');
});

document.getElementById('official-master-form').addEventListener('submit', (event) => {
  event.preventDefault();
  const values = collectFormData(event.currentTarget);
  state.officialsMaster.push({ id: crypto.randomUUID(), name: values.name.trim(), role: values.role.trim(), nip: values.nip.trim() });
  event.currentTarget.reset(); saveState(); showToast('Petugas berhasil ditambahkan ke master data.');
});

document.addEventListener('change', (event) => {
  if (event.target.matches('[data-person-select]')) updateAssignmentDetails(event.target);
});
  document.getElementById('teacher-master-input').addEventListener('change', (event) => {
  const [file] = event.target.files;
  if (file) importMasterCsv(file, 'teachers');
  event.target.value = '';
});
document.getElementById('student-master-input').addEventListener('change', (event) => {
  const [file] = event.target.files;
  if (file) importMasterCsv(file, 'students');
  event.target.value = '';
});
document.getElementById('archive-search').addEventListener('input', renderArchive);
document.getElementById('archive-year').addEventListener('change', renderArchive);
document.getElementById('agenda-search').addEventListener('input', renderAgenda);
document.getElementById('backup-input').addEventListener('change', (event) => {
  const [file] = event.target.files;
  if (file) importBackup(file);
  event.target.value = '';
});

document.getElementById('api-url-input').addEventListener('input', (event) => {
  API_BASE_URL = event.target.value.trim();
  document.getElementById('spreadsheet-status').textContent = 'URL hanya tersimpan di sesi ini. Klik "Koneksikan ke Spreadsheet" untuk memeriksa dan mengaktifkan.';
});

document.getElementById('api-key-input').addEventListener('input', (event) => {
  API_KEY = event.target.value;
});

document.getElementById('api-url-input').value = API_BASE_URL;
document.getElementById('spreadsheet-status').textContent = 'URL deployment Web App sudah diisi. Uji koneksi untuk memastikan akses baca, lalu tulis/hapus data uji yang aman.';

document.getElementById('today-label').textContent = new Intl.DateTimeFormat('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' }).format(new Date());
const initialPage = window.location.hash.slice(1);
if (initialPage === 'archive' || initialPage === 'agenda' || initialPage === 'settings') navigate(initialPage);
else renderAll();

// Load shared records on every page open. Apps Script HTML Service uses
// google.script.run; standalone hosting tries the configured Web App URL.
if (API_BASE_URL && !API_BASE_URL.includes('{WEB_APP_ID}')) {
  document.getElementById('spreadsheet-status').textContent = 'Menghubungkan dan membaca data Spreadsheet…';
  connectStorageToSpreadsheet();
}

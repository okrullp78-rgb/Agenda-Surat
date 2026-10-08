import Papa from 'papaparse';

const STORAGE_KEY = 'smpn13.agenda-surat.v1';
const CURRENT_YEAR = new Date().getFullYear();
const PAGE_LABELS = { overview: 'Ringkasan', archive: 'Arsip surat tugas', agenda: 'Agenda surat', settings: 'Pengaturan & data' };
const defaultState = {
  tasks: [],
  incoming: [],
  teachers: [],
  students: [],
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
      teachers: Array.isArray(stored.teachers) ? stored.teachers : [],
      students: Array.isArray(stored.students) ? stored.students : [],
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
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  renderAll();
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
  window.location.hash = page === 'overview' ? 'ringkasan' : page;
  renderAll();
}

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
  document.getElementById('metric-outgoing').textContent = String(state.tasks.length);
  document.getElementById('overview-year').textContent = CURRENT_YEAR;
  document.getElementById('metric-teachers').textContent = String(state.teachers.length);
  document.getElementById('overview-teacher-count').textContent = `${state.teachers.length} guru`;
  document.getElementById('overview-student-count').textContent = `${state.students.length} siswa`;
  const agendaRows = [
    ...state.incoming.map((entry) => ({ kind: 'Masuk', title: entry.subject, detail: `${entry.sender || 'Pengirim belum diisi'} · ${formatDate(entry.receivedDate)}`, date: entry.receivedDate })),
    ...state.tasks.map((entry) => ({ kind: 'Keluar', title: entry.subject, detail: `${entry.number} · ${formatDate(entry.issuedDate)}`, date: entry.issuedDate })),
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
    <td class="action-cell"><button class="icon-button" data-action="print-options" data-id="${escapeHtml(task.id)}" aria-label="Pilih dokumen cetak" title="Pilih dokumen cetak">▤</button></td>
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
  document.getElementById('agenda-add-button').dataset.action = isIncoming ? 'new-incoming' : 'new-task';
  document.getElementById('agenda-add-button').innerHTML = `<span class="button-plus">+</span> ${isIncoming ? 'Catat surat masuk' : 'Buat surat tugas'}`;
  document.getElementById('incoming-tab-count').textContent = state.incoming.length;
  document.getElementById('outgoing-tab-count').textContent = state.tasks.length;
  const query = document.getElementById('agenda-search').value.trim().toLowerCase();
  const rows = isIncoming ? state.incoming.map((entry) => ({ ...entry, category: 'Masuk' })) : state.tasks.map((task) => ({ ...task, category: 'Keluar', sender: task.recipient, letterNumber: task.number, letterDate: task.issuedDate, receivedDate: task.issuedDate }));
  const filtered = rows.filter((entry) => [entry.letterNumber, entry.subject, entry.sender, entry.recipient, entry.disposition, entry.number].join(' ').toLowerCase().includes(query)).sort((first, second) => ((isIncoming ? second.receivedDate : second.issuedDate) || '').localeCompare((isIncoming ? first.receivedDate : first.issuedDate) || ''));
  const head = isIncoming
    ? '<tr><th>Tanggal terima</th><th>Nomor surat</th><th>Tanggal surat</th><th>Pengirim</th><th>Perihal</th><th>Disposisi</th></tr>'
    : '<tr><th>Tanggal kirim</th><th>Nomor surat</th><th>Tujuan</th><th>Perihal</th><th>Kegiatan</th><th class="action-column">Aksi</th></tr>';
  document.getElementById('agenda-head').innerHTML = head;
  document.getElementById('agenda-rows').innerHTML = filtered.map((entry) => isIncoming
    ? `<tr><td>${escapeHtml(formatDate(entry.receivedDate))}</td><td>${escapeHtml(entry.letterNumber || '—')}</td><td>${escapeHtml(formatDate(entry.letterDate))}</td><td>${escapeHtml(entry.sender || '—')}</td><td>${escapeHtml(entry.subject || '—')}</td><td>${escapeHtml(entry.disposition || '—')}</td></tr>`
    : `<tr><td>${escapeHtml(formatDate(entry.issuedDate))}</td><td>${escapeHtml(entry.number)}</td><td>${escapeHtml(entry.recipient || '—')}</td><td>${escapeHtml(entry.subject || '—')}</td><td>${escapeHtml(entry.eventName || '—')}</td><td class="action-cell"><button class="icon-button" data-action="print-options" data-id="${escapeHtml(entry.id)}" aria-label="Pilih dokumen cetak" title="Pilih dokumen cetak">▤</button></td></tr>`).join('');
  document.getElementById('agenda-empty').classList.toggle('is-visible', filtered.length === 0);
  document.querySelector('#page-agenda .data-table').style.display = filtered.length ? '' : 'none';
  document.getElementById('agenda-empty-title').textContent = isIncoming ? 'Belum ada surat masuk' : 'Belum ada surat keluar';
  document.getElementById('agenda-empty-copy').textContent = isIncoming ? 'Catat surat yang diterima untuk mulai mengisi agenda.' : 'Surat tugas baru akan otomatis tercatat di agenda keluar.';
  document.querySelector('#agenda-empty [data-action]').dataset.action = isIncoming ? 'new-incoming' : 'new-task';
  document.querySelector('#agenda-empty .text-button').innerHTML = `${isIncoming ? 'Catat surat masuk' : 'Buat surat tugas'} <span aria-hidden="true">→</span>`;
}

function renderSettings() {
  const officialsForm = document.getElementById('officials-form');
  for (const [name, value] of Object.entries(state.officials)) officialsForm.elements[name].value = value;
  const validityLabel = (date) => date === '1900-01-01' ? 'Sebelum riwayat' : formatDate(date);
  document.getElementById('principal-history').innerHTML = state.officialHistory.principal.map((official) => `<tr><td>${escapeHtml(validityLabel(official.startDate))}</td><td>${escapeHtml(official.name)}</td><td>${escapeHtml(official.nip)}</td></tr>`).join('');
  document.getElementById('treasurer-history').innerHTML = state.officialHistory.treasurer.map((official) => `<tr><td>${escapeHtml(validityLabel(official.startDate))}</td><td>${escapeHtml(official.name)}</td><td>${escapeHtml(official.nip)}</td></tr>`).join('');
  document.getElementById('teacher-master-count').textContent = `${state.teachers.length} guru`;
  document.getElementById('student-master-count').textContent = `${state.students.length} siswa`;
  document.getElementById('storage-record-count').textContent = `${state.tasks.length + state.incoming.length} catatan tersimpan`;
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
  document.getElementById(containerId).innerHTML = assignments.map((assignment, index) => isTeacher
    ? `<div class="assignment-row teacher-row" data-assignment-row="teachers"><span class="assignment-index">${index + 1}</span><select data-person-select required aria-label="Nama petugas"><option value="">${records.length ? 'Pilih petugas' : 'Unggah data guru dahulu'}</option>${options}</select><output data-rank-value>${escapeHtml(assignment.rank || '—')}</output><output data-position-value>${escapeHtml(assignment.position || '—')}</output><input data-rate-value type="number" min="0" step="1" value="${escapeHtml(assignment.rate ?? '')}" placeholder="Tarif" required aria-label="Tarif petugas per hari"><button type="button" class="icon-button assignment-remove" data-remove-assignment="teachers" aria-label="Hapus petugas ${index + 1}" title="Hapus petugas" ${assignments.length === 1 ? 'disabled' : ''}>×</button></div>`
    : `<div class="assignment-row student-row" data-assignment-row="students"><span class="assignment-index">${index + 1}</span><select data-person-select aria-label="Nama siswa"><option value="">${records.length ? 'Pilih siswa' : 'Unggah data siswa dahulu'}</option>${options}</select><output data-class-value>${escapeHtml(assignment.class || '—')}</output><input data-rate-value type="number" min="0" step="1" value="${escapeHtml(assignment.rate ?? '')}" placeholder="Tarif" aria-label="Tarif siswa per hari"><button type="button" class="icon-button assignment-remove" data-remove-assignment="students" aria-label="Hapus siswa ${index + 1}" title="Hapus siswa" ${assignments.length === 1 ? 'disabled' : ''}>×</button></div>`).join('');
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

function openTaskForm() {
  const dialog = document.getElementById('record-dialog');
  const today = localDateValue();
  document.getElementById('dialog-eyebrow').textContent = 'ARSIP BARU';
  document.getElementById('dialog-title').textContent = 'Buat surat tugas';
  document.getElementById('dialog-footnote').textContent = 'Nomor diisi manual; arsip dan agenda diurutkan menurut tanggal surat.';
  document.getElementById('record-form').dataset.kind = 'task';
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
  renderAssignmentRows('teachers');
  renderAssignmentRows('students');
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
  const task = {
    ...values,
    id: crypto.randomUUID(),
    year,
    number: values.number.trim(),
    createdAt: new Date().toISOString(),
  };
  state.tasks.push(task);
  saveState();
  document.getElementById('record-dialog').close();
  showToast(`Surat tugas ${task.number} tersimpan di arsip dan agenda keluar.`);
}

function saveIncoming(form) {
  const entry = { ...collectFormData(form), id: crypto.randomUUID(), createdAt: new Date().toISOString() };
  state.incoming.push(entry);
  saveState();
  document.getElementById('record-dialog').close();
  showToast('Surat masuk tersimpan di agenda.');
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
    : ['Tanggal surat', 'Nomor surat tugas', 'Tujuan', 'Perihal', 'Nama yang diberi tugas', 'Nama kegiatan', 'Hari', 'Tanggal kegiatan', 'Waktu', 'Tempat', 'Lampiran'];
  const records = isIncoming ? state.incoming.map((entry) => [entry.receivedDate, entry.letterNumber, entry.letterDate, entry.sender, entry.subject, entry.disposition, entry.attachment]) : state.tasks.map((task) => [task.issuedDate, task.number, task.recipient, task.subject, task.assignees, task.eventName, task.eventDay, task.eventDate, task.eventTime, task.place, task.attachment]);
  const csv = [headers, ...records].map((row) => row.map(csvValue).join(',')).join('\r\n');
  downloadFile(`agenda-${isIncoming ? 'masuk' : 'keluar'}-${CURRENT_YEAR}.csv`, `\ufeff${csv}`, 'text/csv;charset=utf-8');
  showToast(`Agenda surat ${isIncoming ? 'masuk' : 'keluar'} diekspor.`);
}

function exportBackup() {
  const backup = { app: 'Administrasi Persuratan SMPN 13 Tasikmalaya', version: 1, exportedAt: new Date().toISOString(), ...state };
  downloadFile(`cadangan-agenda-surat-${localDateValue()}.json`, JSON.stringify(backup, null, 2), 'application/json');
  showToast('Cadangan data berhasil diunduh.');
}

async function importBackup(file) {
  try {
    const imported = JSON.parse(await file.text());
    if (!imported || !Array.isArray(imported.tasks) || !Array.isArray(imported.incoming)) throw new Error('Format cadangan tidak sesuai.');
    const taskIds = new Set(state.tasks.map((task) => task.id));
    const incomingIds = new Set(state.incoming.map((entry) => entry.id));
    state.tasks.push(...imported.tasks.filter((task) => task.id && !taskIds.has(task.id)));
    state.incoming.push(...imported.incoming.filter((entry) => entry.id && !incomingIds.has(entry.id)));
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
    title = 'SURAT PERINTAH PERJALANAN DINAS';
    const sppdRank = extendedSppdPerson ? [selectedSppdPerson[4], selectedSppdPerson[3]].filter(Boolean).join(' / ') : selectedSppdPerson[2] || '';
    content = `<p class="subtitle">(SPPD)</p><table class="details"><tr><td>1. Pengguna Anggaran / Kuasa Pengguna Anggaran</td><td>${escapeHtml(task.reportTo || 'Kepala Sekolah')}</td></tr><tr><td>2. Nama pegawai yang diperintah / NIP</td><td>${escapeHtml(selectedSppdPerson[0] || '')}<br>NIP. ${escapeHtml(selectedSppdPerson[1] || '')}</td></tr><tr><td>3. Pangkat dan golongan</td><td>${escapeHtml(sppdRank)}</td></tr><tr><td>4. Jabatan / OPD</td><td>${escapeHtml(extendedSppdPerson ? selectedSppdPerson[5] || '' : selectedSppdPerson[3] || '')}</td></tr><tr><td>5. Maksud perjalanan dinas</td><td>${escapeHtml(task.travelPurpose || task.eventName || '')}</td></tr><tr><td>6. Alat angkutan</td><td>${escapeHtml(task.transportMode || '')}</td></tr><tr><td>7. Tempat berangkat</td><td>${escapeHtml(task.origin || 'SMP Negeri 13 Tasikmalaya')}</td></tr><tr><td>8. Tempat tujuan</td><td>${escapeHtml(task.destination || task.place || '')}</td></tr><tr><td>9. Lama perjalanan</td><td>${days} hari</td></tr><tr><td>10. Tanggal berangkat</td><td>${escapeHtml(formatDate(task.departureDate || task.eventDate))}</td></tr><tr><td>11. Tanggal harus kembali</td><td>${escapeHtml(formatDate(task.returnDate || task.eventDate))}</td></tr><tr><td>12. Pembebanan anggaran</td><td>${escapeHtml(task.budgetSource || '')}</td></tr><tr><td>13. Keterangan lain-lain</td><td>${escapeHtml(task.subject || '')}</td></tr></table>${printSignature(task.issuedDate)}`;
  } else if (type === 'sppd-continuation') {
    title = 'SPPD LEMBAR KE-2';
    landscape = true;
    content = `<p>SPPD No. : ${escapeHtml(task.number)}</p><p>Berangkat dari: ${escapeHtml(task.origin || 'SMP Negeri 13 Tasikmalaya')} &nbsp; Ke: ${escapeHtml(task.destination || task.place || '')}</p><p>Maksud perjalanan: ${escapeHtml(task.travelPurpose || task.eventName || '')}</p><table class="grid route-grid"><thead><tr><th>No</th><th>Tanggal</th><th>Dari</th><th>Ke</th><th>Catatan / tanda tangan pejabat</th></tr></thead><tbody>${routeRows}</tbody></table>`;
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

function printLongDate(value) {
  if (!value) return '';
  const [year, month, day] = String(value).split('-').map(Number);
  return new Intl.DateTimeFormat('id-ID', { day: 'numeric', month: 'long', year: 'numeric', timeZone: 'UTC' }).format(new Date(Date.UTC(year, month - 1, day)));
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

function printDocument(taskId, type, sppdPersonId = '') {
  const task = state.tasks.find((entry) => entry.id === taskId);
  if (!task) return;
  const people = parseAssignees(task.assignees);
  const students = parseStudents(task.students);
  const start = task.eventStartDate || task.eventDate || task.issuedDate;
  const end = task.eventEndDate || task.eventDate || start;
  const dateRange = start && end && start !== end ? `${printLongDate(start)} s.d. ${printLongDate(end)}` : printLongDate(start || end);
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
  const standardHeader = `<header class="official-header"><img class="school-logo city-logo" src="/assets/logo-kota.jpeg" alt="Lambang Kota Tasikmalaya"><div><strong>PEMERINTAH KOTA TASIKMALAYA</strong><strong>SMP NEGERI 13 TASIKMALAYA</strong><p>Jl. Letjen H. Ibrahim Adjie Km.2 Indihiang Tasikmalaya 46151 Telp. (0265) 335695</p><p>e-mail : smp13tasikmalaya@yahoo.co.id</p></div><img class="school-logo school-crest" src="/assets/logo-smpn13.png" alt="Lambang SMP Negeri 13 Tasikmalaya"></header>`;
  const transportHeader = `<header class="official-header"><img class="school-logo city-logo" src="/assets/logo-kota.jpeg" alt="Lambang Kota Tasikmalaya"><div><strong>PEMERINTAH KOTA TASIKMALAYA</strong><strong>SMP NEGERI 13 TASIKMALAYA</strong><p>Jalan.Letjen.H.Ibrahim Adjie Km.2 Indihiang Tasikmalaya Telp.335695&nbsp;&nbsp; Pos.46151</p><p>e-mail : smp13tasikmalaya@yahoo.co.id</p></div><img class="school-logo school-crest" src="/assets/logo-smpn13.png" alt="Lambang SMP Negeri 13 Tasikmalaya"></header>`;
  const details = `<table class="letter-details"><tr><td>Dasar Surat</td><td>:</td><td>${escapeHtml(task.basis || '')}</td></tr><tr><td>Nomor</td><td>:</td><td>${escapeHtml(task.basisNumber || '')}</td></tr><tr><td>Tanggal</td><td>:</td><td>${escapeHtml(printLongDate(task.basisDate))}</td></tr><tr><td>Perihal</td><td>:</td><td>${escapeHtml(task.subject || '')}</td></tr></table>`;
  const eventInfo = `<table class="letter-details event-details"><tr><td>Tanggal</td><td>:</td><td>${escapeHtml(dateRange)}</td></tr><tr><td>Waktu</td><td>:</td><td>${escapeHtml(task.eventTime || '')}</td></tr><tr><td>Tempat</td><td>:</td><td>${escapeHtml(task.place || task.recipient || '')}</td></tr></table>`;
  const teacherRows = people.map((person, index) => `<tr><td>${index + 1}</td><td>${escapeHtml(person[0] || '')}</td><td>${escapeHtml(person[1] || '')}</td><td>${escapeHtml(person[3] || '')}</td><td>${escapeHtml(person[5] || '')}</td></tr>`).join('');
  const studentRows = students.map((student, index) => `<tr><td>${index + 1}</td><td>${escapeHtml(student[0] || '')}</td><td>${escapeHtml(student[1] || '')}</td><td>${escapeHtml(student[2] || '')}</td><td></td></tr>`).join('');
  const assignments = `<table class="assignment-table"><thead><tr><th>No</th><th>N a m a</th><th>NIP</th><th>Pangkat/ Gol</th><th>Jabatan</th></tr></thead><tbody>${teacherRows}</tbody></table>`;
  const pupils = `<table class="assignment-table"><thead><tr><th>No</th><th>N a m a</th><th>NIS</th><th>Kelas</th><th>Ket</th></tr></thead><tbody>${studentRows}</tbody></table>`;
  const signature = `<div class="official-signature"><p>Dikeluarkan di&nbsp;&nbsp; : Tasikmalaya</p><p>Pada Tanggal&nbsp;&nbsp;&nbsp;&nbsp; : ${escapeHtml(printLongDate(task.issuedDate))}</p><p>Kepala Sekolah</p><div class="signature-space"></div><p><strong>${escapeHtml(principal.name || '')}</strong></p><p>NIP.&nbsp; ${escapeHtml(principal.nip || '')}</p></div>`;
  if (false) {
    content = `<div class="sppd-identifiers"><p>Lembaran Ke&nbsp;&nbsp;&nbsp; : I</p><p>Kode No&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : __________________</p><p>SPPD No&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : 400.3.5/<span class="number-slot">${escapeHtml(sequence)}</span>/SMPN.13/${escapeHtml(year)}</p></div><h1 class="sppd-title">SURAT PERINTAH PERJALANAN DINAS<br>( SPPD)</h1><table class="sppd-table"><tr><td>1.  Pengguna Anggaran/Kuasa Pengguna Anggaran</td><td>Kepala Sekolah</td></tr><tr><td>2.    Nama Pegawai yang diperintah / NIP</td><td>${escapeHtml(selectedPerson[0] || '')}<br>${escapeHtml(selectedPerson[1] || '')}</td></tr><tr><td>3.    a. Pangkat dan Golongan</td><td>${escapeHtml([selectedPerson[4], selectedPerson[3]].filter(Boolean).join(' / '))}</td></tr><tr><td>      b. Jabatan/OPD</td><td>${escapeHtml(selectedPerson[5] || '')}</td></tr><tr><td>      c. Tingkat Biaya Perjalanan Dinas</td><td></td></tr><tr><td>4.  Maksud Perjalanan Dinas</td><td>${escapeHtml(task.eventName || '')}</td></tr><tr><td>5.  Alat angkut yang dipergunakan</td><td>Kendaraan</td></tr><tr><td>6.  a. Tempat Berangkat</td><td>SMP Negeri 13 Tasikmalaya</td></tr><tr><td>      b. Tempat Tujuan</td><td>${escapeHtml(task.recipient || '')}</td></tr><tr><td>7.  a. Lamanya Perjalanan Dinas</td><td>${days} hari</td></tr><tr><td>      b. Tanggal Berangkat</td><td>${escapeHtml(printLongDate(start))}</td></tr><tr><td>      c. Tanggal Harus Kembali/tiba d tempat baru</td><td>${escapeHtml(printLongDate(end))}</td></tr><tr><td>8.  Pengikut : Nama</td><td></td></tr><tr><td>9.  Pembebanan Anggaran</td><td></td></tr><tr><td>      a. Instansi</td><td>a.  -</td></tr><tr><td>      b. Mata Anggaran</td><td>b.  -</td></tr><tr><td>10. Keterangan Lain-lain</td><td>${escapeHtml(task.subject || '')}</td></tr></table><div class="sppd-sign"><p>Dikeluarkan di   :  Tasikmalaya</p><p>Pada Tanggal     : ${escapeHtml(printLongDate(task.issuedDate))}</p><p>Pengguna Anggaran/Kuasa Pengguna Anggaran</p><div class="signature-space"></div><p><strong>${escapeHtml(state.officials.principalName)}</strong></p><p>NIP.  ${escapeHtml(state.officials.principalNip)}</p></div>`;
  }
  let title = '';
  let body = '';
  let header = standardHeader;
  let orientation = 'portrait';

  if (type === 'task' || type === 'student') {
    title = 'S U R A T&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; T U G A S';
    const list = type === 'task' ? assignments : pupils;
    body = `<p class="official-number">Nomor : ${escapeHtml(officialNumber)}</p>${details}<p class="order-heading">KEPALA SMP NEGERI 13 TASIKMALAYA</p><p class="order-heading">MENUGASKAN</p>${list}${eventInfo}<p class="closing-line">Demikian surat tugas ini kami berikan agar dilaksanakan sebagaimana mestinya.</p>${signature}`;
  } else if (type === 'sppd') {
    title = '';
    body = `<div class="sppd-identifiers"><p>Lembaran Ke&nbsp;&nbsp;&nbsp; : I</p><p>Kode No&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : __________________</p><p>SPPD No&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : 400.3.5/<span class="number-slot">${escapeHtml(sequence)}</span>/SMPN.13/${escapeHtml(year)}</p></div><h1 class="sppd-title">SURAT PERINTAH PERJALANAN DINAS<br>( SPPD)</h1><table class="sppd-table"><tr><td>1.  Pengguna Anggaran/Kuasa Pengguna Anggaran</td><td>Kepala Sekolah</td></tr><tr><td>2.    Nama Pegawai yang diperintah / NIP</td><td>${escapeHtml(selectedPerson[0] || '')}<br>${escapeHtml(selectedPerson[1] || '')}</td></tr><tr><td>3.    a. Pangkat dan Golongan</td><td>${escapeHtml([selectedPerson[4], selectedPerson[3]].filter(Boolean).join(' / '))}</td></tr><tr><td>      b. Jabatan/OPD</td><td>${escapeHtml(selectedPerson[5] || '')}</td></tr><tr><td>      c. Tingkat Biaya Perjalanan Dinas</td><td></td></tr><tr><td>4.  Maksud Perjalanan Dinas</td><td>${escapeHtml(task.eventName || '')}</td></tr><tr><td>5.  Alat angkut yang dipergunakan</td><td>Kendaraan</td></tr><tr><td>6.  a. Tempat Berangkat</td><td>SMP Negeri 13 Tasikmalaya</td></tr><tr><td>      b. Tempat Tujuan</td><td>${escapeHtml(task.recipient || '')}</td></tr><tr><td>7.  a. Lamanya Perjalanan Dinas</td><td>${days} hari</td></tr><tr><td>      b. Tanggal Berangkat</td><td>${escapeHtml(printLongDate(start))}</td></tr><tr><td>      c. Tanggal Harus Kembali/tiba d tempat baru</td><td>${escapeHtml(printLongDate(end))}</td></tr><tr><td>8.  Pengikut : Nama</td><td></td></tr><tr><td>9.  Pembebanan Anggaran</td><td></td></tr><tr><td>      a. Instansi</td><td>a.  -</td></tr><tr><td>      b. Mata Anggaran</td><td>b.  -</td></tr><tr><td>10. Keterangan Lain-lain</td><td>${escapeHtml(task.subject || '')}</td></tr></table><div class="sppd-sign"><p>Dikeluarkan di   :  Tasikmalaya</p><p>Pada Tanggal     : ${escapeHtml(printLongDate(task.issuedDate))}</p><p>Pengguna Anggaran/Kuasa Pengguna Anggaran</p><div class="signature-space"></div><p><strong>AGUS ROHMAN, S.Pd., M.Si.</strong></p><p>NIP.  19650927 198903 1 008</p></div>`;
  } else if (type === 'sppd-continuation') {
    title = '';
    const dateBlocks = Array.from({ length: days }, (_, index) => {
      const date = new Date(`${start}T00:00:00`);
      date.setDate(date.getDate() + index);
      const dayDate = localDateValue(date);
      return `<table class="travel-day"><tr><td colspan="2"><strong>Hari ke-${index + 1} : ${escapeHtml(printLongDate(dayDate))}</strong></td></tr><tr><td>I.   Tiba di : ${escapeHtml(task.recipient || '')}</td><td>Berangkat dari : ${escapeHtml(task.recipient || '')}</td></tr><tr><td> Pada Tanggal : ...............................................</td><td>Ke   : SMP Negeri 13 Tasikmalaya</td></tr><tr><td>(${escapeHtml(printLongDate(dayDate))})</td><td>Pada Tanggal : ...................................................</td></tr><tr><td>Mengetahui pejabat setempat</td><td>Tanda tangan pejabat</td></tr><tr><td>(................................................)</td><td>(................................................)</td></tr></table>`;
    }).join('');
    body = `<div class="sppd-identifiers"><p>Lembaran Ke&nbsp;&nbsp;&nbsp; : II</p><p>Kode No&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : __________________</p><p>SPPD No.                         : ${escapeHtml(officialNumber)}</p><p>Petugas&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; : ${escapeHtml(selectedPerson[0] || '')}</p></div><p>Berangkat dari               : SMPN 13 Tasikmalaya</p><p>(Tempat Kedudukan)</p><p>Ke : ${escapeHtml(task.recipient || '')}</p><p>MENGETAHUI PEJABAT SETEMPAT :</p>${dateBlocks}`;
    body += `<table class="return-review"><tr><td><strong>II. Tiba kembali di : SMP Negeri 13 Tasikmalaya</strong><p>Pada Tanggal : ...............................................</p><p>(${escapeHtml(printLongDate(end))})</p></td><td><strong>Telah diperiksa, dengan keterangan :</strong><p>Bahwa perjalanan dinas tersebut diatas benar dilakukan atas perintahnya dan semata-mata untuk kepentingan jabatan dalam waktu yang sesingkat-singkatnya.</p><p>Kepala Sekolah</p><div class="signature-space"></div><p><strong>${escapeHtml(principal.name || '')}</strong></p><p>NIP.  ${escapeHtml(principal.nip || '')}</p></td></tr></table><p class="sppd-notes-heading">III.  CATATAN LAIN-LAIN :</p><div class="sppd-notes-space"></div><p class="sppd-notes-heading">IV. PERHATIAN</p><p>Pejabat yang berwenang menerbitkan SPPD, pegawai yang melakukan perjalanan dinas, para pejabat yang mengesahkan tanggal berangkat/tiba serta bendahara bertanggung jawab berdasarkan peraturan-peraturan Keuangan Negara apabila Negara mendapat rugi akibat kesalahan, kealpaan.</p>`;
    orientation = 'portrait';
  } else if (type === 'nota') {
    title = 'NOTA DINAS';
    body = `<p class="subtitle">(Laporan Hasil Pelaksanaan Kegiatan)</p><table class="nota-table"><tr><td>1</td><td>Kepada Yth.</td><td>${escapeHtml(task.reportTo || 'Kepala Sekolah')}</td></tr><tr><td>2</td><td>Dari (yang diberi tugas)</td><td>${escapeHtml(selectedPerson[0] || '')}</td></tr><tr><td>3</td><td>Perihal/ Nama Kegiatan</td><td>${escapeHtml(task.eventName || '')}</td></tr><tr><td>4</td><td>Tempat Kegiatan</td><td>${escapeHtml(task.recipient || '')}</td></tr><tr><td>5</td><td>Tanggal Pelaksanaan</td><td>${escapeHtml(dateRange)}</td></tr><tr><td>6</td><td>Dasar Penugasan</td><td>${escapeHtml(task.basis || '')}</td></tr><tr><td></td><td>a. Surat Tugas dari</td><td>Kepala Sekolah</td></tr><tr><td></td><td>b. Nomor</td><td>${escapeHtml(sequence)}</td></tr><tr><td></td><td>c. Tanggal</td><td>${escapeHtml(printLongDate(task.issuedDate))}</td></tr></table><p class="report-intro">Berdasarkan Surat Tugas tersebut di atas, saya laporkan hasil-hasil pelaksanaan kegiatan sebagai berikut :</p><div class="report-space official-report-lines">${escapeHtml(task.reportText || '').replace(/\r?\n/g, '<br>')}</div><p class="recommendation-heading">Usul, masukan, rekomendasi dan tindak lanjut :</p><div class="report-space official-report-lines recommendation-lines">${escapeHtml(task.recommendations || '').replace(/\r?\n/g, '<br>')}</div><p class="nota-closing">Demikian Nota Dinas/ Laporan hasil pelaksanaan kegiatan ini dibuat dengan sebenarnya sebagai bentuk pertanggungjawaban atas pelaksanaan tugas yang telah dipercayakan kepada saya.</p><div class="nota-sign"><p>Tasikmalaya, .........................................</p><p>Yang Membuat Laporan,</p><div class="signature-space"></div><p>___________________________</p><p>NIP.</p></div>`;
  } else if (type === 'transport' || type === 'transport-students') {
    const isStudent = type === 'transport-students';
    header = transportHeader;
    title = isStudent ? 'DAFTAR PENYERAHAN TRANSPORT SISWA' : 'DAFTAR PENYERAHAN TRANSPORT';
    const rows = isStudent ? students.map((student, index) => {
      const rate = Number(student[3]) || 0;
      return `<tr><td>${index + 1}</td><td>${escapeHtml(student[0] || '')}</td><td>${escapeHtml(student[2] || '')}</td><td>${days} Hari x ${money(rate)}</td><td>${money(days * rate)}</td><td></td></tr>`;
    }).join('') : people.map((person, index) => {
      const rate = Number(person[6]) || 0;
      return `<tr><td>${index + 1}</td><td>${escapeHtml(person[0] || '')}<br>NIP.${escapeHtml(person[1] || '')}</td><td>${escapeHtml(person[3] || '')}</td><td>${days} Hari x ${money(rate)}</td><td>${money(days * rate)}</td><td></td></tr>`;
    }).join('');
      body = `<p class="event-banner">${escapeHtml((task.eventName || '').toUpperCase())}</p><p>Nomor : ${escapeHtml(officialNumber)}</p><table class="transport-table"><thead><tr><th>NO</th><th>Nama</th><th>${isStudent ? 'Kelas' : 'Gol'}</th><th>Volume</th><th>Jumlah Yang Di Terima</th><th>Tanda Tangan</th></tr></thead><tbody>${rows}</tbody></table><div class="transport-approval"><div><p>Setuju Dibayar :</p><p>Kepala SMPN 13 Tasikmalaya</p><div class="signature-space"></div><p><strong>${escapeHtml(principal.name || '')}</strong></p><p>NIP. ${escapeHtml(principal.nip || '')}</p></div><div><p>Lunas Dibayar Tanggal :</p><p>Bendaharawan,</p><div class="signature-space"></div><p><strong>${escapeHtml(treasurer.name || '')}</strong></p><p>NIP. ${escapeHtml(treasurer.nip || '')}</p></div></div>`;
    orientation = 'landscape';
  } else {
    return;
  }

  body = body.replaceAll('AGUS ROHMAN, S.Pd., M.Si.', escapeHtml(principal.name || ''))
    .replaceAll('19650927 198903 1 008', escapeHtml(principal.nip || ''));
  const printWindow = window.open('', '_blank', 'width=900,height=720');
  if (!printWindow) {
    showToast('Izinkan pop-up untuk mencetak dokumen.');
    return;
  }
  document.getElementById('print-dialog').close();
  printWindow.document.write(`<!doctype html><html lang="id"><head><meta charset="utf-8"><title>${escapeHtml(title || 'SPPD Lembar ke 2')} - ${escapeHtml(officialNumber)}</title><style>
    @page{size:A4 ${orientation};margin:15mm}*{box-sizing:border-box}body{margin:0;color:#111;font:12px "Times New Roman",serif;line-height:1.35}.official-header{text-align:center;border-bottom:3px double #111;padding:0 0 7px;margin-bottom:18px}.official-header strong{display:block;font-size:15px}.official-header p{margin:3px 0;font-size:10px}.official-title{text-align:center;margin:13px 0 16px;font-size:16px;font-weight:bold}.official-number{text-align:center;margin:0 0 17px}.letter-details{width:100%;border-collapse:collapse;margin:8px 0 13px}.letter-details td{padding:3px 4px;vertical-align:top}.letter-details td:first-child{width:145px}.letter-details td:nth-child(2){width:16px}.event-details{margin-top:18px}.order-heading{text-align:center;margin:5px 0;font-weight:bold}.assignment-table,.transport-table{width:100%;border-collapse:collapse;margin:12px 0}.assignment-table th,.assignment-table td,.transport-table th,.transport-table td{border:1px solid #222;padding:5px;text-align:left;vertical-align:top}.assignment-table th,.transport-table th{text-align:center}.assignment-table td:first-child,.transport-table td:first-child{text-align:center}.closing-line{margin-top:15px}.official-signature{width:255px;margin:20px 12px 0 auto}.official-signature p{margin:4px 0}.signature-space{height:54px}.subtitle{text-align:center;margin:4px 0 14px}.sppd-table,.nota-table{width:100%;border-collapse:collapse}.sppd-table td,.nota-table td{padding:4px;vertical-align:top}.sppd-table td:first-child{width:56%}.nota-table td:first-child{width:28px}.nota-table td:nth-child(2){width:190px}.sppd-sign,.nota-sign{width:300px;margin:18px 0 0 auto}.sppd-sign p,.nota-sign p{margin:4px 0}.travel-day{width:100%;border-collapse:collapse;margin:16px 0;page-break-inside:avoid}.travel-day td{width:33.33%;padding:7px;border:1px solid #444;vertical-align:top}.report-intro{margin:14px 0 8px}.report-space{min-height:52px;white-space:normal;margin:5px 0}.recommendation-heading{margin:12px 0 5px}.nota-closing{margin:16px 0}.event-banner{text-align:center;font-weight:bold}.transport-table td{height:43px}@media screen{body{max-width:${orientation === 'landscape' ? '1050px' : '760px'};margin:26px auto;padding:0 22px}}@media print{body{max-width:none}.assignment-table tr,.transport-table tr{break-inside:avoid}}
    </style><style>.official-header{display:grid;grid-template-columns:72px minmax(0,1fr) 72px;align-items:center;gap:10px;text-align:center;border-bottom:3px double #111;padding:0 0 8px;margin:0 0 18px}.official-header>div{min-width:0}.official-header .school-logo{display:block;width:66px;height:72px;object-fit:contain}.official-header .city-logo{justify-self:start}.official-header .school-crest{justify-self:end}.official-header strong{display:block;font-size:15px}.official-header p{margin:3px 0;font-size:9px}.sppd-identifiers{font-size:11px}.sppd-identifiers p{margin:2px 0}.number-slot{display:inline-block;min-width:52px;border-bottom:1px solid #111;text-align:center}.sppd-title{text-align:center;margin:9px 0 14px;font-size:15px}.travel-day{table-layout:fixed;margin:12px 0 16px;border:1px solid #333}.travel-day td{width:50%;height:auto;padding:7px 9px;border:1px solid #444;vertical-align:top}.travel-day tr:nth-child(n+2) td{height:38px}.travel-day tr:last-child td{height:46px}.official-report-lines{height:360px;min-height:360px;overflow:hidden;padding:2px 4px;line-height:24px;background-image:linear-gradient(to bottom,transparent 23px,#888 23px,#888 24px,transparent 24px);background-size:100% 24px;white-space:normal;page-break-inside:avoid}.official-report-lines.recommendation-lines{height:120px;min-height:120px}.nota-closing{margin-top:14px}.sppd-table tr{min-height:25px}</style></head><body>${header}${title ? `<h1 class="official-title">${title}</h1>` : ''}${body}<script>window.onload=()=>window.print()<\/script></body></html>`);
  const printStyle = printWindow.document.createElement('style');
  printStyle.textContent = '.transport-approval{display:grid;grid-template-columns:1fr 1fr;gap:30px;margin-top:48px;page-break-inside:avoid}.transport-approval>div{min-width:0}.transport-approval p{margin:4px 0}.transport-approval .signature-space{height:60px}.return-review{width:100%;border-collapse:collapse;margin-top:14px;page-break-inside:avoid}.return-review td{width:50%;padding:7px;border-top:1px solid #333;border-bottom:1px solid #333;vertical-align:top}.return-review td+td{border-left:1px solid #333}.return-review p{margin:8px 0}.return-review .signature-space{height:48px}.sppd-notes-heading{margin:8px 0;font-weight:bold}.sppd-notes-space{height:56px;border-bottom:1px solid #777}';
  printWindow.document.head.appendChild(printStyle);
  printWindow.document.close();
}

function printArchive() {
  const archiveRows = [...state.tasks].sort((first, second) => (first.issuedDate || '').localeCompare(second.issuedDate || '')).map((task, index) => `<tr><td>${index + 1}</td><td>${escapeHtml(task.number)}</td><td>${escapeHtml(task.basis || '')}</td><td>${escapeHtml(task.basisNumber || '')}</td><td>${escapeHtml(formatDate(task.basisDate))}</td><td>${escapeHtml(task.subject || '')}</td><td>${escapeHtml(assigneeNames(task.assignees))}</td><td>${escapeHtml(task.eventName || '')}</td><td>${escapeHtml(task.eventDay || '')}</td><td>${escapeHtml(formatDate(task.eventStartDate || task.eventDate))}${task.eventEndDate && task.eventEndDate !== task.eventStartDate ? ` s.d. ${escapeHtml(formatDate(task.eventEndDate))}` : ''}</td><td>${escapeHtml(task.eventTime || '')}</td><td>${escapeHtml(task.place || '')}</td><td>${escapeHtml(task.attachment || '')}</td></tr>`).join('');
  const printWindow = window.open('', '_blank', 'width=1100,height=750');
  if (!printWindow) {
    showToast('Izinkan pop-up untuk mencetak arsip.');
    return;
  }
  printWindow.document.write(`<!doctype html><html lang="id"><head><meta charset="utf-8"><title>Arsip Surat Tugas</title><style>@page{size:A3 landscape;margin:12mm}body{font:10px Arial,sans-serif;color:#111}h1{text-align:center;font-size:16px}p{text-align:center}table{width:100%;border-collapse:collapse}th,td{border:1px solid #333;padding:4px;text-align:left;vertical-align:top}th{background:#eef2ee}@media print{tr{break-inside:avoid}}</style></head><body><h1>ARSIP SURAT TUGAS</h1><p>SMP NEGERI 13 TASIKMALAYA</p><table><thead><tr><th>No</th><th>No Surat Tugas</th><th>Dasar Surat</th><th>Nomor Surat</th><th>Tanggal Surat</th><th>Perihal</th><th>Nama yang Diberi Tugas</th><th>Nama Kegiatan</th><th>Hari</th><th>Tanggal</th><th>Waktu</th><th>Tempat</th><th>Lampiran</th></tr></thead><tbody>${archiveRows}</tbody></table><script>window.onload=()=>window.print()<\/script></body></html>`);
  printWindow.document.close();
}

function handleAction(action, element) {
  if (action === 'new-task') openTaskForm();
  if (action === 'new-incoming') openIncomingForm();
  if (action === 'close-dialog') document.getElementById('record-dialog').close();
  if (action === 'close-print-dialog') document.getElementById('print-dialog').close();
  if (action === 'print-options') openPrintOptions(element.dataset.id);
  if (action === 'print-archive') printArchive();
  if (action === 'download-teacher-template') downloadRosterTemplate('teachers');
  if (action === 'download-student-template') downloadRosterTemplate('students');
  if (action === 'export-agenda') exportAgenda();
  if (action === 'export-backup') exportBackup();
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

document.querySelectorAll('[data-agenda-tab]').forEach((button) => button.addEventListener('click', () => {
  activeAgendaTab = button.dataset.agendaTab;
  renderAgenda();
}));

document.getElementById('today-label').textContent = new Intl.DateTimeFormat('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' }).format(new Date());
const initialPage = window.location.hash.slice(1);
if (initialPage === 'archive' || initialPage === 'agenda' || initialPage === 'settings') navigate(initialPage);
else renderAll();
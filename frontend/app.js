const API_URL = 'https://script.google.com/macros/s/AKfycbwsRocB7bsQLfXiazKGI-O158ppsRnQPVsrtvzVaoyUUgMdanidkOJc_pg--lddbDGPhQ/exec';

const OFFICER_RANKS = [
  'Police Constable',
  'Sergeant',
  'Inspector',
  'Chief Inspector',
  'Superintendent',
  'Chief Superintendent',
  'Commander',
  'Deputy Assistant Commissioner',
  'Assistant Commissioner',
  'Deputy Commissioner',
  'Commissioner',
];

const SYSTEM_ROLES = ['Constable', 'Sergeant', 'Inspector', 'Chief Inspector', 'Command'];
const ACCESS_LEVELS = [...OFFICER_RANKS, ...SYSTEM_ROLES];
const TRAINING_STANDARDS = ['Taser', 'MOE', 'Blue Ticket', 'Motorbike', 'Basic', 'Response', 'IPP', 'Advanced', 'Advanced + TPAC'];
const OFFICER_STATUSES = ['Active', 'LOA', 'Suspended', 'Archived'];
const TRAINING_STATUSES = ['Not Started', 'In Progress', 'Passed', 'Failed'];
const DISCIPLINE_TYPES = ['Note', 'Warning', 'Suspension', 'Removal'];
const DISCIPLINE_STATUSES = ['Active', 'Expired', 'Appealed', 'Removed'];
const LOA_STATUSES = ['Pending', 'Approved', 'Denied', 'Cancelled'];

const state = {
  token: localStorage.getItem('mo8_token') || '',
  user: null,
  permissions: [],
  activeView: 'dashboard',
  officers: [],
  selectedOfficerId: '',
};

const elements = {
  loginForm: document.querySelector('#loginForm'),
  loginStatus: document.querySelector('#loginStatus'),
  loginView: document.querySelector('#loginView'),
  appView: document.querySelector('#appView'),
  nav: document.querySelector('#nav'),
  identity: document.querySelector('#identity'),
  currentUser: document.querySelector('#currentUser'),
  logoutButton: document.querySelector('#logoutButton'),
  passwordButton: document.querySelector('#passwordButton'),
  pageTitle: document.querySelector('#pageTitle'),
  pageSubtitle: document.querySelector('#pageSubtitle'),
  dashboardView: document.querySelector('#dashboardView'),
  editorDialog: document.querySelector('#editorDialog'),
  editorForm: document.querySelector('#editorForm'),
  editorTitle: document.querySelector('#editorTitle'),
  editorFields: document.querySelector('#editorFields'),
  editorStatus: document.querySelector('#editorStatus'),
};

document.addEventListener('click', handleDocumentClick);

elements.loginForm.addEventListener('submit', async (event) => {
  event.preventDefault();
  elements.loginStatus.textContent = 'Signing in...';
  const form = new FormData(elements.loginForm);
  const response = await api('login', {
    username: form.get('username'),
    password: form.get('password'),
    userAgent: navigator.userAgent,
  }, false);

  if (!response.ok) {
    elements.loginStatus.textContent = response.error || 'Login failed.';
    return;
  }

  state.token = response.token;
  state.user = response.user;
  state.permissions = response.permissions || [];
  localStorage.setItem('mo8_token', state.token);
  showApp();
  await showView('dashboard');
});

elements.logoutButton.addEventListener('click', async () => {
  await api('logout', {});
  localStorage.removeItem('mo8_token');
  state.token = '';
  state.user = null;
  state.permissions = [];
  showLogin();
});

elements.passwordButton.addEventListener('click', () => {
  openEditor('Change password', [
    field('CurrentPassword', 'Current password', 'password'),
    field('NewPassword', 'New password', 'password'),
  ], async (values) => api('changePassword', values), {
    successMessage: 'Password changed.',
  });
});

document.querySelectorAll('.nav-item').forEach((button) => {
  button.addEventListener('click', async () => {
    state.activeView = button.dataset.view;
    document.querySelectorAll('.nav-item').forEach((item) => item.classList.toggle('active', item === button));
    await showView(state.activeView);
  });
});

document.querySelector('#newOfficerButton').addEventListener('click', () => openOfficerEditor());
document.querySelector('#newDocumentButton').addEventListener('click', () => openDocumentEditor());
document.querySelector('#newUserButton').addEventListener('click', () => openUserEditor());

async function boot() {
  if (!API_URL || API_URL.includes('YOUR_APPS_SCRIPT')) {
    elements.loginStatus.textContent = 'Set API_URL in frontend/app.js before logging in.';
    return;
  }

  if (!state.token) {
    showLogin();
    return;
  }

  const response = await api('me', {});
  if (!response.ok) {
    localStorage.removeItem('mo8_token');
    showLogin();
    return;
  }

  state.user = response.user;
  state.permissions = response.permissions || [];
  showApp();
  await showView('dashboard');
}

function showLogin() {
  document.body.classList.remove('is-authenticated');
  elements.pageTitle.textContent = 'Sign in';
  elements.pageSubtitle.textContent = 'MO8 roleplay community administration';
  elements.loginView.hidden = false;
  elements.appView.hidden = true;
  elements.nav.hidden = true;
  elements.identity.hidden = true;
}

function showApp() {
  document.body.classList.add('is-authenticated');
  elements.loginView.hidden = true;
  elements.appView.hidden = false;
  elements.nav.hidden = false;
  elements.identity.hidden = false;
  elements.currentUser.innerHTML = `
    <strong>${escapeHtml(state.user.RobloxUsername)}</strong>
    <span>${escapeHtml(state.user.Rank || state.user.Role)}</span>
  `;
  applyPermissions();
}

function applyPermissions() {
  document.querySelectorAll('[data-permission]').forEach((node) => {
    node.hidden = !can(node.dataset.permission);
  });
}

async function showView(view) {
  const titles = {
    dashboard: ['Dashboard', 'Current MO8 overview'],
    officers: ['Officers', 'MO8 officer database'],
    officerProfile: ['Officer Profile', 'Individual record and linked history'],
    training: ['Training', 'Training standards and status'],
    discipline: ['Discipline', 'Internal roleplay administration records'],
    loa: ['Leave of Absence', 'Requests and reviews'],
    documents: ['Documents', 'Training guides and policy links'],
    users: ['Users', 'Sergeant+ login accounts'],
    audit: ['Audit Log', 'System activity trail'],
  };

  state.activeView = view;
  Object.keys(titles).forEach((key) => {
    const section = document.querySelector(`#${key}View`);
    if (section) section.hidden = key !== view;
  });

  elements.pageTitle.textContent = titles[view][0];
  elements.pageSubtitle.textContent = titles[view][1];

  const loaders = {
    dashboard: loadDashboard,
    officers: loadOfficers,
    officerProfile: () => loadOfficerProfile(state.selectedOfficerId),
    training: loadTraining,
    discipline: loadDiscipline,
    loa: loadLoa,
    documents: loadDocuments,
    users: loadUsers,
    audit: loadAudit,
  };

  await loaders[view]();
  applyPermissions();
}

async function showViewOnly(view) {
  document.querySelectorAll('#appView > section').forEach((section) => {
    section.hidden = section.id !== `${view}View`;
  });
}

async function loadDashboard() {
  await showViewOnly('dashboard');
  const response = await api('dashboard', {});
  if (!response.ok) return renderError(elements.dashboardView, response.error);

  const counts = response.counts || {};
  elements.dashboardView.innerHTML = [
    stat('Active Officers', counts.activeOfficers || 0),
    stat('Pending LOA', counts.loaPending || 0),
    stat('Training Records', counts.trainingRecords || 0),
    stat('Active Discipline', counts.activeDiscipline || 0),
  ].join('');
}

async function loadOfficers() {
  await showViewOnly('officers');
  const response = await api('listOfficers', {});
  if (!response.ok) return renderTable('#officersTable', [], ['Error'], response.error);
  state.officers = response.rows || [];
  renderTable('#officersTable', state.officers, ['RobloxUsername', 'Callsign', 'Rank', 'Status', 'JoinDate', 'UpdatedAt'], {
    rowAction: (row) => `data-open-officer="${escapeHtml(row.OfficerID)}"`,
  });
}

async function loadOfficerProfile(officerId) {
  await showViewOnly('officerProfile');
  if (!officerId) {
    document.querySelector('#officerProfileView').innerHTML = emptyState('No officer selected.');
    return;
  }

  const response = await api('getOfficerProfile', { OfficerID: officerId });
  if (!response.ok) {
    document.querySelector('#officerProfileView').innerHTML = emptyState(response.error || 'Officer not found.');
    return;
  }

  renderOfficerProfile(response);
}

async function loadTraining() {
  await showViewOnly('training');
  const [trainingResponse, officersResponse] = await Promise.all([
    api('listTraining', {}),
    api('listOfficers', {}),
  ]);
  const trainingRows = trainingResponse.rows || [];
  const officerRows = officersResponse.rows || [];
  renderTrainingMatrix(officerRows, trainingRows);
  renderTable('#trainingTable', trainingRows, ['OfficerID', 'Standard', 'Status', 'Assessor', 'DateCompleted', 'ExpiryDate']);
}

async function loadDiscipline() {
  await showViewOnly('discipline');
  const response = await api('listDiscipline', {});
  renderTable('#disciplineTable', response.rows || [], ['OfficerID', 'Type', 'Summary', 'IssuedBy', 'IssuedAt', 'Status']);
}

async function loadLoa() {
  await showViewOnly('loa');
  const response = await api('listLoa', {});
  renderTable('#loaTable', response.rows || [], ['OfficerID', 'StartDate', 'EndDate', 'Reason', 'Status', 'ReviewedBy'], {
    actions: (row) => can('APPROVE_LOA') && row.Status === 'Pending'
      ? `<button class="mini" data-review-loa="${escapeHtml(row.RequestID)}" data-status="Approved">Approve</button><button class="mini ghost" data-review-loa="${escapeHtml(row.RequestID)}" data-status="Denied">Deny</button>`
      : '',
  });
}

async function loadDocuments() {
  await showViewOnly('documents');
  const response = await api('listDocuments', {});
  renderTable('#documentsTable', response.rows || [], ['Title', 'Category', 'RequiredRole', 'Status', 'UpdatedAt', 'DriveURL']);
}

async function loadUsers() {
  await showViewOnly('users');
  const response = await api('listUsers', {});
  renderTable('#usersTable', response.rows || [], ['RobloxUsername', 'DiscordID', 'Rank', 'Role', 'Status'], {
    actions: (row) => `<button class="mini" data-reset-password="${escapeHtml(row.UserID)}">Reset password</button>`,
  });
}

async function loadAudit() {
  await showViewOnly('audit');
  const response = await api('auditLog', {});
  renderTable('#auditTable', response.rows || [], ['Timestamp', 'ActorUserID', 'Action', 'TargetType', 'TargetID']);
}

function renderOfficerProfile(data) {
  const officer = data.officer;
  const container = document.querySelector('#officerProfileView');
  container.innerHTML = `
    <div class="profile-head">
      <button class="ghost" data-view-link="officers">Back</button>
      <div>
        <h2>${escapeHtml(officer.RobloxUsername)}</h2>
        <p>${escapeHtml(officer.Callsign || 'No callsign')} / ${escapeHtml(officer.Rank || 'No rank')}</p>
      </div>
      <div class="profile-actions">
        <button data-edit-officer="${escapeHtml(officer.OfficerID)}" data-permission="EDIT_OFFICERS">Edit officer</button>
        <button data-add-discipline="${escapeHtml(officer.OfficerID)}" data-permission="ADD_DISCIPLINE">Add discipline</button>
        <button data-add-loa="${escapeHtml(officer.OfficerID)}" data-permission="CREATE_LOA">Add LOA</button>
      </div>
    </div>

    <section class="profile-grid">
      ${detailCard('Status', formatCell(officer.Status), true)}
      ${detailCard('Join date', officer.JoinDate || 'Not set')}
      ${detailCard('Discord ID', officer.DiscordID || 'Not set')}
      ${detailCard('Updated', officer.UpdatedAt || 'Not set')}
    </section>

    ${trainingChecklist(officer.OfficerID, data.training)}

    <section class="profile-notes">
      <h3>Notes</h3>
      <p>${escapeHtml(officer.Notes || 'No notes recorded.')}</p>
    </section>

    <section class="profile-columns">
      ${profileTable('Training History', data.training, ['Standard', 'Status', 'Assessor', 'DateCompleted', 'ExpiryDate'])}
      ${profileTable('Discipline', data.discipline, ['Type', 'Summary', 'IssuedAt', 'Status'])}
      ${profileTable('LOA', data.loa, ['StartDate', 'EndDate', 'Reason', 'Status'])}
    </section>
  `;
  applyPermissions();
}

function openOfficerEditor(officer = {}) {
  openEditor(officer.OfficerID ? 'Edit officer' : 'Add officer', [
    hiddenField('OfficerID', officer.OfficerID),
    field('RobloxUsername', 'Roblox username', 'text', false, officer.RobloxUsername),
    field('DiscordID', 'Discord ID', 'text', false, officer.DiscordID),
    field('Callsign', 'Callsign', 'text', false, officer.Callsign),
    selectField('Rank', 'Rank', OFFICER_RANKS, officer.Rank),
    selectField('Status', 'Status', OFFICER_STATUSES, officer.Status || 'Active'),
    field('JoinDate', 'Join date', 'date', false, officer.JoinDate),
    field('Notes', 'Notes', 'textarea', true, officer.Notes),
  ], async (values) => api('saveOfficer', values));
}

function openTrainingEditor(officerId) {
  openEditor('Add training record', [
    hiddenField('OfficerID', officerId),
    field('Standard', 'Training standard'),
    selectField('Status', 'Status', TRAINING_STATUSES, 'In Progress'),
    field('Assessor', 'Assessor', 'text', false, state.user.RobloxUsername),
    field('DateCompleted', 'Date completed', 'date'),
    field('ExpiryDate', 'Expiry date', 'date'),
    field('Notes', 'Notes', 'textarea', true),
  ], async (values) => api('saveTraining', values));
}

function openDisciplineEditor(officerId) {
  openEditor('Add discipline record', [
    hiddenField('OfficerID', officerId),
    selectField('Type', 'Type', DISCIPLINE_TYPES, 'Note'),
    field('Summary', 'Summary'),
    field('Details', 'Details', 'textarea', true),
    selectField('Status', 'Status', DISCIPLINE_STATUSES, 'Active'),
  ], async (values) => api('addDiscipline', values));
}

function openLoaEditor(officerId) {
  openEditor('Add LOA request', [
    hiddenField('OfficerID', officerId),
    field('StartDate', 'Start date', 'date'),
    field('EndDate', 'End date', 'date'),
    field('Reason', 'Reason', 'textarea', true),
  ], async (values) => api('createLoa', values));
}

function openDocumentEditor(document = {}) {
  openEditor(document.DocumentID ? 'Edit document' : 'Add document', [
    hiddenField('DocumentID', document.DocumentID),
    field('Title', 'Title', 'text', false, document.Title),
    selectField('Category', 'Category', ['Training', 'Policy', 'SOP', 'Form'], document.Category),
    field('DriveURL', 'Drive URL', 'url', false, document.DriveURL),
    selectField('RequiredRole', 'Minimum rank or role', ACCESS_LEVELS, document.RequiredRole || 'Police Constable'),
    selectField('Status', 'Status', ['Published', 'Draft', 'Archived'], document.Status),
  ], async (values) => api('saveDocument', values));
}

function openUserEditor(user = {}) {
  openEditor(user.UserID ? 'Edit user' : 'Add user', [
    hiddenField('UserID', user.UserID),
    field('RobloxUsername', 'Roblox username', 'text', false, user.RobloxUsername),
    field('DiscordID', 'Discord ID', 'text', false, user.DiscordID),
    selectField('Rank', 'Rank', OFFICER_RANKS, user.Rank || 'Police Constable'),
    selectField('Role', 'System role', SYSTEM_ROLES, user.Role || 'Constable'),
    selectField('Status', 'Status', ['Active', 'Suspended', 'Archived'], user.Status || 'Active'),
    field('TemporaryPassword', 'Temporary password', 'text', false),
  ], async (values) => api('saveUser', values), {
    successMessage: 'User saved. Copy the temporary password from the response if one was generated.',
  });
}

async function handleDocumentClick(event) {
  const viewLink = event.target.closest('[data-view-link]');
  if (viewLink) {
    await showView(viewLink.dataset.viewLink);
    return;
  }

  const officerLink = event.target.closest('[data-open-officer]');
  if (officerLink) {
    state.selectedOfficerId = officerLink.dataset.openOfficer;
    await showView('officerProfile');
    return;
  }

  const editOfficer = event.target.closest('[data-edit-officer]');
  if (editOfficer) {
    const officer = state.officers.find((row) => row.OfficerID === editOfficer.dataset.editOfficer);
    if (officer) openOfficerEditor(officer);
    return;
  }

  const trainingToggle = event.target.closest('[data-training-toggle]');
  if (trainingToggle) {
    await api('setOfficerTraining', {
      OfficerID: trainingToggle.dataset.officerId,
      Standard: trainingToggle.dataset.standard,
      Enabled: trainingToggle.checked,
    });
    await loadOfficerProfile(trainingToggle.dataset.officerId);
    return;
  }

  const addDiscipline = event.target.closest('[data-add-discipline]');
  if (addDiscipline) return openDisciplineEditor(addDiscipline.dataset.addDiscipline);

  const addLoa = event.target.closest('[data-add-loa]');
  if (addLoa) return openLoaEditor(addLoa.dataset.addLoa);

  const resetPassword = event.target.closest('[data-reset-password]');
  if (resetPassword) {
    const response = await api('resetUserPassword', { UserID: resetPassword.dataset.resetPassword });
    alert(response.ok ? `Temporary password: ${response.temporaryPassword}` : response.error);
    return;
  }

  const reviewLoa = event.target.closest('[data-review-loa]');
  if (reviewLoa) {
    await api('reviewLoa', { RequestID: reviewLoa.dataset.reviewLoa, Status: reviewLoa.dataset.status });
    await loadLoa();
  }
}

function stat(label, value) {
  return `
    <article class="stat">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
    </article>
  `;
}

function detailCard(label, value, allowHtml = false) {
  const content = allowHtml ? value : escapeHtml(value);
  return `<article class="detail-card"><span>${escapeHtml(label)}</span><strong>${content}</strong></article>`;
}

function profileTable(title, rows, columns) {
  const body = rows.length
    ? rows.map((row) => `<tr>${columns.map((column) => `<td>${formatCell(row[column])}</td>`).join('')}</tr>`).join('')
    : `<tr><td colspan="${columns.length}">No records found.</td></tr>`;
  return `
    <section class="profile-panel">
      <h3>${escapeHtml(title)}</h3>
      <div class="table-wrap compact">
        <table>
          <thead><tr>${columns.map((column) => `<th>${escapeHtml(column)}</th>`).join('')}</tr></thead>
          <tbody>${body}</tbody>
        </table>
      </div>
    </section>
  `;
}

function trainingChecklist(officerId, trainingRows) {
  const rows = TRAINING_STANDARDS.map((standard) => {
    const record = trainingRows.find((item) => item.Standard === standard && item.Status === 'Passed');
    const checked = record ? ' checked' : '';
    const disabled = can('MANAGE_TRAINING') ? '' : ' disabled';
    return `
      <label class="training-check">
        <input type="checkbox" data-training-toggle data-officer-id="${escapeHtml(officerId)}" data-standard="${escapeHtml(standard)}"${checked}${disabled}>
        <span>${escapeHtml(standard)}</span>
      </label>
    `;
  }).join('');

  return `
    <section class="cert-panel">
      <div>
        <h3>Training Certifications</h3>
        <p>Sergeants and above can assign or remove certifications.</p>
      </div>
      <div class="cert-grid">${rows}</div>
    </section>
  `;
}

function renderTrainingMatrix(officers, trainingRows) {
  const container = document.querySelector('#trainingMatrix');
  const standards = TRAINING_STANDARDS;
  if (!officers.length) {
    container.innerHTML = `<p class="empty">Training matrix will appear once officers and standards have records.</p>`;
    return;
  }

  const rows = officers.map((officer) => {
    const cells = standards.map((standard) => {
      const record = trainingRows.find((item) => item.OfficerID === officer.OfficerID && item.Standard === standard);
      return `<td>${formatCell(record ? record.Status : 'Not Started')}</td>`;
    }).join('');
    return `<tr><td>${escapeHtml(officer.RobloxUsername)}</td><td>${escapeHtml(officer.Callsign || '')}</td>${cells}</tr>`;
  }).join('');

  container.innerHTML = `
    <h3>Training Matrix</h3>
    <div class="table-wrap compact">
      <table>
        <thead>
          <tr><th>Officer</th><th>Callsign</th>${standards.map((standard) => `<th>${escapeHtml(standard)}</th>`).join('')}</tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;
}

function renderTable(selector, rows, columns, options = {}) {
  const table = document.querySelector(selector);
  if (!rows.length) {
    table.innerHTML = `<tbody><tr><td>${escapeHtml(options.emptyMessage || 'No records found.')}</td></tr></tbody>`;
    return;
  }

  const actionHeader = options.actions ? '<th>Actions</th>' : '';
  const head = `<thead><tr>${columns.map((column) => `<th>${escapeHtml(column)}</th>`).join('')}${actionHeader}</tr></thead>`;
  const body = rows.map((row) => {
    const attrs = options.rowAction ? options.rowAction(row) : '';
    const actionCell = options.actions ? `<td class="actions">${options.actions(row)}</td>` : '';
    return `<tr ${attrs}>${columns.map((column) => `<td>${formatCell(row[column])}</td>`).join('')}${actionCell}</tr>`;
  }).join('');
  table.innerHTML = `${head}<tbody>${body}</tbody>`;
}

function renderError(container, message) {
  container.innerHTML = `<article class="stat"><strong>!</strong><span>${escapeHtml(message || 'Something went wrong.')}</span></article>`;
}

function emptyState(message) {
  return `<section class="data-view"><p class="empty">${escapeHtml(message)}</p></section>`;
}

function formatCell(value) {
  const text = value === undefined || value === null ? '' : String(value);
  if (text.startsWith('https://')) {
    return `<a href="${escapeHtml(text)}" target="_blank" rel="noopener">Open</a>`;
  }
  if (['Active', 'Published', 'Passed', 'Approved'].includes(text)) {
    return `<span class="pill success">${escapeHtml(text)}</span>`;
  }
  if (['LOA', 'Pending', 'In Progress', 'Draft', 'Not Started'].includes(text)) {
    return `<span class="pill warning">${escapeHtml(text)}</span>`;
  }
  if (['Suspended', 'Archived', 'Failed', 'Denied', 'Expired', 'Removed'].includes(text)) {
    return `<span class="pill danger">${escapeHtml(text)}</span>`;
  }
  return escapeHtml(text);
}

function openEditor(title, fields, onSubmit, options = {}) {
  elements.editorTitle.textContent = title;
  elements.editorStatus.textContent = '';
  elements.editorFields.innerHTML = fields.map((item) => item.html).join('');
  elements.editorDialog.showModal();

  elements.editorForm.onsubmit = async (event) => {
    event.preventDefault();
    if (event.submitter && event.submitter.value === 'cancel') {
      elements.editorDialog.close();
      return;
    }

    elements.editorStatus.textContent = 'Saving...';
    const values = Object.fromEntries(new FormData(elements.editorForm).entries());
    const response = await onSubmit(values);
    if (!response.ok) {
      elements.editorStatus.textContent = response.error || 'Save failed.';
      return;
    }

    const generatedPassword = response.temporaryPassword ? ` Temporary password: ${response.temporaryPassword}` : '';
    if (options.successMessage || generatedPassword) {
      alert(`${options.successMessage || 'Saved.'}${generatedPassword}`);
    }

    elements.editorDialog.close();
    await showView(state.activeView);
  };
}

function hiddenField(name, value = '') {
  return { html: `<input type="hidden" name="${escapeHtml(name)}" value="${escapeHtml(value || '')}">` };
}

function field(name, label, type = 'text', wide = false, value = '') {
  const className = wide ? ' class="wide"' : '';
  if (type === 'textarea') {
    return { html: `<label${className}>${escapeHtml(label)}<textarea name="${escapeHtml(name)}">${escapeHtml(value || '')}</textarea></label>` };
  }
  return { html: `<label${className}>${escapeHtml(label)}<input type="${escapeHtml(type)}" name="${escapeHtml(name)}" value="${escapeHtml(value || '')}"></label>` };
}

function selectField(name, label, options, selected = '') {
  const optionHtml = options.map((option) => {
    const isSelected = option === selected ? ' selected' : '';
    return `<option value="${escapeHtml(option)}"${isSelected}>${escapeHtml(option)}</option>`;
  }).join('');
  return { html: `<label>${escapeHtml(label)}<select name="${escapeHtml(name)}">${optionHtml}</select></label>` };
}

async function api(action, data = {}, includeToken = true) {
  const payload = Object.assign({}, data, { action });
  if (includeToken) payload.token = state.token;

  try {
    const response = await fetch(API_URL, {
      method: 'POST',
      body: JSON.stringify(payload),
    });
    return await response.json();
  } catch (error) {
    return { ok: false, error: error.message };
  }
}

function can(permission) {
  return state.permissions.includes('FULL_ACCESS') || state.permissions.includes(permission);
}

function escapeHtml(value) {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

boot();

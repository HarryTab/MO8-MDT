const API_URL = 'https://script.google.com/macros/s/AKfycbwsRocB7bsQLfXiazKGI-O158ppsRnQPVsrtvzVaoyUUgMdanidkOJc_pg--lddbDGPhQ/exec';

const state = {
  token: localStorage.getItem('mo8_token') || '',
  user: null,
  permissions: [],
  activeView: 'dashboard',
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
  pageTitle: document.querySelector('#pageTitle'),
  pageSubtitle: document.querySelector('#pageSubtitle'),
  dashboardView: document.querySelector('#dashboardView'),
  editorDialog: document.querySelector('#editorDialog'),
  editorForm: document.querySelector('#editorForm'),
  editorTitle: document.querySelector('#editorTitle'),
  editorFields: document.querySelector('#editorFields'),
  editorStatus: document.querySelector('#editorStatus'),
};

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

document.querySelectorAll('.nav-item').forEach((button) => {
  button.addEventListener('click', async () => {
    state.activeView = button.dataset.view;
    document.querySelectorAll('.nav-item').forEach((item) => item.classList.toggle('active', item === button));
    await showView(state.activeView);
  });
});

document.querySelector('#newOfficerButton').addEventListener('click', () => {
  openEditor('Add officer', [
    field('RobloxUsername', 'Roblox username'),
    field('DiscordID', 'Discord ID'),
    field('Callsign', 'Callsign'),
    field('Rank', 'Rank'),
    selectField('Status', 'Status', ['Active', 'LOA', 'Suspended', 'Archived']),
    field('JoinDate', 'Join date', 'date'),
    field('Notes', 'Notes', 'textarea', true),
  ], async (values) => {
    return api('saveOfficer', values);
  });
});

document.querySelector('#newDocumentButton').addEventListener('click', () => {
  openEditor('Add document', [
    field('Title', 'Title'),
    selectField('Category', 'Category', ['Training', 'Policy', 'SOP', 'Form']),
    field('DriveURL', 'Drive URL'),
    selectField('RequiredRole', 'Required role', ['Sergeant', 'Inspector', 'Chief Inspector', 'Command']),
    selectField('Status', 'Status', ['Published', 'Draft', 'Archived']),
  ], async (values) => {
    return api('saveDocument', values);
  });
});

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
  elements.pageTitle.textContent = 'Sign in';
  elements.pageSubtitle.textContent = 'MO8 roleplay community administration';
  elements.loginView.hidden = false;
  elements.appView.hidden = true;
  elements.nav.hidden = true;
  elements.identity.hidden = true;
}

function showApp() {
  elements.loginView.hidden = true;
  elements.appView.hidden = false;
  elements.nav.hidden = false;
  elements.identity.hidden = false;
  elements.currentUser.textContent = `${state.user.RobloxUsername} - ${state.user.Role}`;
}

async function showView(view) {
  const titles = {
    dashboard: ['Dashboard', 'Current MO8 overview'],
    officers: ['Officers', 'MO8 officer database'],
    training: ['Training', 'Training standards and status'],
    discipline: ['Discipline', 'Internal roleplay administration records'],
    loa: ['Leave of Absence', 'Requests awaiting review'],
    documents: ['Documents', 'Training guides and policy links'],
    audit: ['Audit Log', 'System activity trail'],
  };

  Object.keys(titles).forEach((key) => {
    const section = document.querySelector(`#${key}View`);
    if (section) section.hidden = key !== view;
  });

  elements.pageTitle.textContent = titles[view][0];
  elements.pageSubtitle.textContent = titles[view][1];

  const loaders = {
    dashboard: loadDashboard,
    officers: loadOfficers,
    training: loadTraining,
    discipline: loadDiscipline,
    loa: loadLoa,
    documents: loadDocuments,
    audit: loadAudit,
  };

  await loaders[view]();
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
  renderTable('#officersTable', response.rows || [], ['RobloxUsername', 'Callsign', 'Rank', 'Status', 'JoinDate', 'UpdatedAt']);
}

async function loadTraining() {
  await showViewOnly('training');
  const response = await api('listTraining', {});
  renderTable('#trainingTable', response.rows || [], ['OfficerID', 'Standard', 'Status', 'Assessor', 'DateCompleted', 'ExpiryDate']);
}

async function loadDiscipline() {
  await showViewOnly('discipline');
  const response = await api('listDiscipline', {});
  renderTable('#disciplineTable', response.rows || [], ['OfficerID', 'Type', 'Summary', 'IssuedBy', 'IssuedAt', 'Status']);
}

async function loadLoa() {
  await showViewOnly('loa');
  const response = await api('listLoa', {});
  renderTable('#loaTable', response.rows || [], ['OfficerID', 'StartDate', 'EndDate', 'Reason', 'Status', 'ReviewedBy']);
}

async function loadDocuments() {
  await showViewOnly('documents');
  const response = await api('listDocuments', {});
  renderTable('#documentsTable', response.rows || [], ['Title', 'Category', 'RequiredRole', 'Status', 'UpdatedAt', 'DriveURL']);
}

async function loadAudit() {
  await showViewOnly('audit');
  const response = await api('auditLog', {});
  renderTable('#auditTable', response.rows || [], ['Timestamp', 'ActorUserID', 'Action', 'TargetType', 'TargetID']);
}

async function showViewOnly(view) {
  document.querySelectorAll('#appView > section').forEach((section) => {
    section.hidden = section.id !== `${view}View`;
  });
}

function stat(label, value) {
  return `<article class="stat"><strong>${escapeHtml(value)}</strong><span>${escapeHtml(label)}</span></article>`;
}

function renderTable(selector, rows, columns) {
  const table = document.querySelector(selector);
  if (!rows.length) {
    table.innerHTML = `<tbody><tr><td>No records found.</td></tr></tbody>`;
    return;
  }

  const head = `<thead><tr>${columns.map((column) => `<th>${escapeHtml(column)}</th>`).join('')}</tr></thead>`;
  const body = rows.map((row) => {
    return `<tr>${columns.map((column) => `<td>${formatCell(row[column])}</td>`).join('')}</tr>`;
  }).join('');
  table.innerHTML = `${head}<tbody>${body}</tbody>`;
}

function renderError(container, message) {
  container.innerHTML = `<article class="stat"><strong>!</strong><span>${escapeHtml(message || 'Something went wrong.')}</span></article>`;
}

function formatCell(value) {
  const text = value === undefined || value === null ? '' : String(value);
  if (text.startsWith('https://')) {
    return `<a href="${escapeHtml(text)}" target="_blank" rel="noopener">Open</a>`;
  }
  return escapeHtml(text);
}

function openEditor(title, fields, onSubmit) {
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

    elements.editorDialog.close();
    await showView(state.activeView);
  };
}

function field(name, label, type = 'text', wide = false) {
  const className = wide ? ' class="wide"' : '';
  if (type === 'textarea') {
    return { html: `<label${className}>${escapeHtml(label)}<textarea name="${escapeHtml(name)}"></textarea></label>` };
  }
  return { html: `<label${className}>${escapeHtml(label)}<input type="${escapeHtml(type)}" name="${escapeHtml(name)}"></label>` };
}

function selectField(name, label, options) {
  const optionHtml = options.map((option) => `<option value="${escapeHtml(option)}">${escapeHtml(option)}</option>`).join('');
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

function escapeHtml(value) {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

boot();

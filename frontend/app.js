const API_URL = 'https://script.google.com/macros/s/AKfycbwsRocB7bsQLfXiazKGI-O158ppsRnQPVsrtvzVaoyUUgMdanidkOJc_pg--lddbDGPhQ/exec';
const APP_VERSION = '2026-05-05-2';

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
const ACCESS_LEVELS = [...OFFICER_RANKS];
const OFFICER_TAGS = ['Roads Crime Team', 'MO8 Command', 'Roads and Traffic Policing Team', 'Bronze Command', 'Silver Command', 'Gold Command'];
const SPECIALIST_TRAINING = ['Taser', 'MOE', 'Blue Ticket', 'Motorbike'];
const DRIVING_STANDARDS = ['Basic', 'Response', 'IPP', 'Advanced', 'Advanced + TPAC'];
const TRAINING_STANDARDS = [...SPECIALIST_TRAINING, ...DRIVING_STANDARDS];
const OFFICER_STATUSES = ['Active', 'LOA', 'Suspended', 'Archived'];
const TRAINING_STATUSES = ['Not Started', 'In Progress', 'Passed', 'Failed'];
const DISCIPLINE_TYPES = ['Note', 'Warning', 'Suspension', 'Removal'];
const DISCIPLINE_STATUSES = ['Active', 'Expired', 'Appealed', 'Removed'];
const LOA_STATUSES = ['Pending', 'Approved', 'Denied', 'Cancelled'];
const CACHE_TTL_MS = 5 * 60 * 1000;
const CACHE_STORAGE_KEY = 'mo8_api_cache';
const BOOT_STORAGE_KEY = 'mo8_boot_ready';
const SESSION_STORAGE_KEY = 'mo8_session_auth';
const VERSION_STORAGE_KEY = 'mo8_app_version';
const USER_PERMISSION_MODES = ['Inherit', 'Allow', 'Deny'];
const ANNOUNCEMENT_STATUSES = ['Published', 'Draft', 'Archived'];

const state = {
  token: localStorage.getItem('mo8_token') || '',
  user: null,
  permissions: [],
  unreadNotifications: 0,
  activeView: 'dashboard',
  officers: [],
  training: [],
  trainingSummary: [],
  discipline: [],
  loa: [],
  tasks: [],
  profileDiscipline: [],
  profileLoa: [],
  documents: [],
  announcements: [],
  rankChanges: [],
  shifts: [],
  shiftStatus: null,
  users: [],
  permissionConfig: null,
  audit: [],
  cache: loadStoredCache(),
  selectedOfficerId: '',
};

const elements = {
  loginForm: document.querySelector('#loginForm'),
  loginStatus: document.querySelector('#loginStatus'),
  loginView: document.querySelector('#loginView'),
  bootView: document.querySelector('#bootView'),
  bootStatus: document.querySelector('#bootStatus'),
  bootSteps: document.querySelector('#bootSteps'),
  bootProgressBar: document.querySelector('#bootProgressBar'),
  appView: document.querySelector('#appView'),
  nav: document.querySelector('#nav'),
  identity: document.querySelector('#identity'),
  currentUser: document.querySelector('#currentUser'),
  logoutButton: document.querySelector('#logoutButton'),
  passwordButton: document.querySelector('#passwordButton'),
  notificationsButton: document.querySelector('#notificationsButton'),
  notificationMenu: document.querySelector('#notificationMenu'),
  infoDialog: document.querySelector('#infoDialog'),
  infoTitle: document.querySelector('#infoTitle'),
  infoContent: document.querySelector('#infoContent'),
  infoCloseButton: document.querySelector('#infoCloseButton'),
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
document.addEventListener('change', handleDocumentChange);

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
  storeSessionAuth(state.user, state.permissions);
  await initializeSession();
});

elements.logoutButton.addEventListener('click', async () => {
  await api('logout', {});
  localStorage.removeItem('mo8_token');
  sessionStorage.removeItem(CACHE_STORAGE_KEY);
  sessionStorage.removeItem(BOOT_STORAGE_KEY);
  sessionStorage.removeItem(SESSION_STORAGE_KEY);
  state.token = '';
  state.user = null;
  state.permissions = [];
  state.unreadNotifications = 0;
  updateNotificationBadge();
  invalidateCache();
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
elements.notificationsButton.addEventListener('click', toggleNotifications);
elements.infoCloseButton.addEventListener('click', () => elements.infoDialog.close());

document.querySelector('#officerSearch').addEventListener('input', () => renderOfficerTable());
document.querySelector('#documentSearch').addEventListener('input', () => renderDocumentTable());
document.querySelector('#documentCategoryFilter').addEventListener('change', () => renderDocumentTable());
document.querySelectorAll('[data-search-view]').forEach((input) => {
  input.addEventListener('input', () => renderSearchableView(input.dataset.searchView));
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
document.querySelector('#newAnnouncementButton').addEventListener('click', () => openAnnouncementEditor());
document.querySelector('#newUserButton').addEventListener('click', () => openUserEditor());
document.querySelector('#startShiftButton').addEventListener('click', startShift);
document.querySelector('#endShiftButton').addEventListener('click', openEndShiftEditor);
document.querySelector('#shiftPeriodFilter').addEventListener('change', loadShift);
document.querySelector('#shiftStartFilter').addEventListener('change', loadShift);
document.querySelector('#shiftEndFilter').addEventListener('change', loadShift);

async function boot() {
  clearCacheForNewVersion();
  if (!API_URL || API_URL.includes('YOUR_APPS_SCRIPT')) {
    elements.loginStatus.textContent = 'Set API_URL in frontend/app.js before logging in.';
    return;
  }

  if (!state.token) {
    showLogin();
    return;
  }

  const cachedAuth = loadSessionAuth();
  if (cachedAuth?.user) {
    state.user = cachedAuth.user;
    state.permissions = cachedAuth.permissions || [];
    showApp();
    await showView(defaultView());
    backgroundPreload();
    validateSessionQuietly();
    return;
  }

  const response = await api('me', {});
  if (!response.ok) {
    localStorage.removeItem('mo8_token');
    sessionStorage.removeItem(SESSION_STORAGE_KEY);
    showLogin();
    return;
  }

  state.user = response.user;
  state.permissions = response.permissions || [];
  storeSessionAuth(state.user, state.permissions);
  if (hasWarmBootCache()) {
    showApp();
    await showView(defaultView());
    backgroundPreload();
    return;
  }
  await initializeSession();
}

async function initializeSession() {
  showBoot();
  const tasks = bootTasks();
  for (let index = 0; index < tasks.length; index += 1) {
    const task = tasks[index];
    updateBootProgress(task.label, index, tasks.length, 'active');
    const response = await task.run();
    updateBootProgress(task.label, index + 1, tasks.length, response && response.ok === false ? 'warning' : 'complete');
    await wait(120);
  }

  showApp();
  sessionStorage.setItem(BOOT_STORAGE_KEY, String(Date.now()));
  await showView(defaultView());
  backgroundPreload();
}

function hasWarmBootCache() {
  return Boolean(getCachedResponse('myProfile', {}) && getCachedResponse('listNotifications', {}));
}

async function validateSessionQuietly() {
  const response = await api('me', {});
  if (!response.ok) {
    localStorage.removeItem('mo8_token');
    sessionStorage.removeItem(CACHE_STORAGE_KEY);
    sessionStorage.removeItem(BOOT_STORAGE_KEY);
    state.token = '';
    state.user = null;
    state.permissions = [];
    invalidateCache();
    showLogin();
    return;
  }
  state.user = response.user;
  state.permissions = response.permissions || [];
  storeSessionAuth(state.user, state.permissions);
  showApp();
}

function bootTasks() {
  const tasks = [
    { label: 'Loading operator profile', run: () => apiCached('myProfile', {}) },
    { label: 'Checking notifications', run: preloadNotifications },
  ];
  if (can('VIEW_DASHBOARD')) tasks.push({ label: 'Preparing dashboard widgets', run: () => apiCached('dashboard', {}) });
  if (can('VIEW_DOCUMENTS')) tasks.push({ label: 'Loading document access', run: () => apiCached('listDocuments', {}) });
  if (can('VIEW_ANNOUNCEMENTS')) tasks.push({ label: 'Syncing notice board', run: () => apiCached('listAnnouncements', {}) });
  tasks.push({ label: 'Checking shift status', run: () => apiCached('shiftStatus', {}) });
  if (can('VIEW_TASKS')) tasks.push({ label: 'Checking task queue', run: () => apiCached('tasks', {}) });
  tasks.push({ label: 'Opening MDT workspace', run: () => Promise.resolve({ ok: true }) });
  return tasks;
}

async function preloadNotifications() {
  const response = await apiCached('listNotifications', {});
  if (response.ok) {
    state.unreadNotifications = response.unread || 0;
    updateNotificationBadge();
  }
  return response;
}

function showBoot() {
  document.body.classList.add('is-booting');
  document.body.classList.remove('is-authenticated');
  elements.pageTitle.textContent = 'Initializing';
  elements.pageSubtitle.textContent = 'Preparing secure MDT workspace';
  elements.loginView.hidden = true;
  elements.bootView.hidden = false;
  elements.appView.hidden = true;
  elements.nav.hidden = true;
  elements.identity.hidden = true;
  elements.bootSteps.innerHTML = bootTasks().map((task) => `<span data-boot-step="${escapeHtml(task.label)}">${escapeHtml(task.label)}</span>`).join('');
  updateBootProgress('Preparing workspace', 0, 1, 'active');
}

function updateBootProgress(label, completed, total, status) {
  elements.bootStatus.textContent = label;
  elements.bootProgressBar.style.width = `${Math.min(100, Math.round((completed / total) * 100))}%`;
  document.querySelectorAll('[data-boot-step]').forEach((step) => {
    if (step.dataset.bootStep !== label) return;
    step.dataset.status = status;
  });
}

function backgroundPreload() {
  const actions = [
    can('VIEW_OFFICERS') ? ['listOfficers', {}] : null,
    can('VIEW_TRAINING') ? ['listTraining', {}] : null,
    can('VIEW_LOA') ? ['listLoa', {}] : null,
    can('VIEW_RANK_LOG') ? ['rankChangeLog', {}] : null,
    ['teamShifts', { Period: 'week' }],
  ].filter(Boolean);

  window.setTimeout(() => {
    Promise.allSettled(actions.map(([action, data]) => apiCached(action, data)));
  }, 500);
}

function wait(ms) {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}

function showLogin() {
  document.body.classList.remove('is-booting');
  document.body.classList.remove('is-authenticated');
  elements.pageTitle.textContent = 'Sign in';
  elements.pageSubtitle.textContent = 'MO8 roleplay community administration';
  elements.loginView.hidden = false;
  elements.bootView.hidden = true;
  elements.appView.hidden = true;
  elements.nav.hidden = true;
  elements.identity.hidden = true;
}

function showApp() {
  document.body.classList.remove('is-booting');
  document.body.classList.add('is-authenticated');
  elements.loginView.hidden = true;
  elements.bootView.hidden = true;
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

async function refreshNotificationBadge() {
  const response = await api('listNotifications', {});
  if (!response.ok) return;
  state.unreadNotifications = response.unread || 0;
  updateNotificationBadge();
}

function updateNotificationBadge() {
  elements.notificationsButton.classList.toggle('has-unread', state.unreadNotifications > 0);
  elements.notificationsButton.setAttribute('aria-label', state.unreadNotifications > 0
    ? `Notifications, ${state.unreadNotifications} unread`
    : 'Notifications');
}

function defaultView() {
  if (can('VIEW_DASHBOARD')) return 'dashboard';
  return 'myProfile';
}

async function showView(view) {
  const titles = {
    dashboard: ['Dashboard', 'Current MO8 overview'],
    myProfile: ['My Profile', 'Your officer record, training, LOA and notifications'],
    shift: ['Shift Log', 'Duty status and team activity'],
    tasks: ['Tasks', 'Outstanding approvals and command actions'],
    officers: ['Officers', 'MO8 officer database'],
    officerProfile: ['Officer Profile', 'Individual record and linked history'],
    rankChanges: ['Rank Change Log', 'Promotion and rank movement history'],
    training: ['Training', 'Training standards and status'],
    discipline: ['Discipline', 'Internal roleplay administration records'],
    loa: ['Leave of Absence', 'Requests and reviews'],
    documents: ['Documents', 'Training guides and policy links'],
    announcements: ['Notice Board', 'Operational updates and command notices'],
    users: ['Users', 'Sergeant+ login accounts'],
    permissions: ['Permissions', 'Role defaults and individual overrides'],
    audit: ['Audit Log', 'System activity trail'],
  };

  state.activeView = view;
  document.querySelectorAll('.nav-item').forEach((item) => item.classList.toggle('active', item.dataset.view === view));
  Object.keys(titles).forEach((key) => {
    const section = document.querySelector(`#${key}View`);
    if (section) section.hidden = key !== view;
  });

  elements.pageTitle.textContent = titles[view][0];
  elements.pageSubtitle.textContent = titles[view][1];
  renderViewLoading(view);

  const loaders = {
    dashboard: loadDashboard,
    myProfile: loadMyProfile,
    shift: loadShift,
    tasks: loadTasks,
    officers: loadOfficers,
    officerProfile: () => loadOfficerProfile(state.selectedOfficerId),
    rankChanges: loadRankChanges,
    training: loadTraining,
    discipline: loadDiscipline,
    loa: loadLoa,
    documents: loadDocuments,
    announcements: loadAnnouncements,
    users: loadUsers,
    permissions: loadPermissions,
    audit: loadAudit,
  };

  await loaders[view]();
  applyPermissions();
}

function renderViewLoading(view) {
  const section = document.querySelector(`#${view}View`);
  if (!section || isViewCached(view)) return;
  if (view === 'dashboard') {
    elements.dashboardView.innerHTML = loadingBlock('Loading dashboard widgets...');
    return;
  }
  if (view === 'tasks') document.querySelector('#tasksSummary').innerHTML = '';
  if (view === 'training') document.querySelector('#trainingMatrix').innerHTML = '';
  if (view === 'permissions') {
    document.querySelector('#permissionsMatrix').innerHTML = loadingBlock('Loading permissions...');
    document.querySelector('#userPermissionsMatrix').innerHTML = '';
    return;
  }
  const messages = {
    myProfile: 'Loading officer profile...',
    shift: 'Loading shift activity...',
    officerProfile: 'Loading officer profile...',
    tasks: 'Loading task queue...',
    officers: 'Loading officer database...',
    rankChanges: 'Loading rank change log...',
    training: 'Loading training matrix...',
    discipline: 'Loading disciplinary records...',
    loa: 'Loading LOA requests...',
    documents: 'Loading documents...',
    announcements: 'Loading notice board...',
    users: 'Loading users...',
    permissions: 'Loading permissions...',
    audit: 'Loading audit log...',
  };
  const message = messages[view] || 'Loading data...';
  const table = section.querySelector('table');
  if (table) {
    table.innerHTML = `<tbody><tr><td colspan="99">${loadingBlock(message)}</td></tr></tbody>`;
    return;
  }
  section.innerHTML = loadingBlock(message);
}

function isViewCached(view) {
  if (view === 'officerProfile') {
    return Boolean(state.cache[cacheKey('getOfficerProfile', { OfficerID: state.selectedOfficerId }, true)]);
  }
  if (view === 'shift') {
    return Boolean(state.cache[cacheKey('teamShifts', shiftQuery(), true)]);
  }
  return Boolean(state.cache[cacheKey(loaderActionForView(view), {}, true)]);
}

function loaderActionForView(view) {
  const actions = {
    dashboard: 'dashboard',
    myProfile: 'myProfile',
    shift: 'teamShifts',
    tasks: 'tasks',
    officers: 'listOfficers',
    rankChanges: 'rankChangeLog',
    training: 'listTraining',
    discipline: 'listDiscipline',
    loa: 'listLoa',
    documents: 'listDocuments',
    announcements: 'listAnnouncements',
    users: 'listUsers',
    permissions: 'permissionsConfig',
    audit: 'auditLog',
  };
  return actions[view] || view;
}

async function showViewOnly(view) {
  document.querySelectorAll('#appView > section').forEach((section) => {
    section.hidden = section.id !== `${view}View`;
  });
}

async function loadDashboard() {
  await showViewOnly('dashboard');
  const response = await apiCached('dashboard', {});
  if (!response.ok) return renderError(elements.dashboardView, response.error);

  const counts = response.counts || {};
  elements.dashboardView.innerHTML = `
    <div class="stat-row">
      ${[
    stat('Active Officers', counts.activeOfficers || 0),
    stat('Currently On LOA', counts.currentlyOnLoa || 0),
    stat('Pending LOA', counts.loaPending || 0),
    stat('Review Due', counts.trainingReviewsDue || 0),
    stat('Missing Core Training', counts.missingCoreTraining || 0),
    stat('Notices', counts.notices || 0),
  ].join('')}
    </div>
    <section class="dashboard-grid">
      ${dashboardPanel('Active LOA Status', response.activeLoa || [], ['Officer', 'Rank', 'EndDate', 'Status'])}
      ${dashboardPanel('Pending LOA', response.pendingLoa || [], ['Officer', 'Rank', 'StartDate', 'EndDate'])}
      ${announcementPanel('Notice Board', response.announcements || [])}
      ${dashboardPanel('Training Reviews', response.trainingReviewsDue || [], ['RobloxUsername', 'Standard', 'ReviewDate', 'UpdatedBy'])}
      ${dashboardPanel('Recent Documents', response.recentDocuments || [], ['Title', 'Category', 'RequiredRole', 'UpdatedAt'])}
      ${dashboardPanel('Recent Activity', response.recentAudit || [], ['Timestamp', 'Action', 'TargetType', 'TargetID'])}
    </section>
  `;
}

async function loadMyProfile() {
  await showViewOnly('myProfile');
  const response = await apiCached('myProfile', {});
  const container = document.querySelector('#myProfileView');
  if (!response.ok) {
    container.innerHTML = emptyState(response.error || 'Could not load profile.');
    return;
  }

  const officer = response.officer;
  const user = response.user;
  const notifications = response.notifications || [];
  state.unreadNotifications = notifications.filter((item) => !item.ReadAt).length;
  updateNotificationBadge();
  container.innerHTML = `
    <div class="profile-head">
      <div>
        <h2>${escapeHtml(user.RobloxUsername)}</h2>
        <p>${escapeHtml(user.Rank || user.Role)} / ${escapeHtml(user.Role)}</p>
      </div>
      <div class="profile-actions">
        <button data-request-loa>Request LOA</button>
        <button data-request-transfer>Request transfer</button>
      </div>
    </div>
    <section class="profile-grid">
      ${detailCard('Callsign', officer ? officer.Callsign || 'Not set' : 'No officer record')}
      ${detailCard('Status', officer ? formatCell(officer.EffectiveStatus || officer.Status, 'Status') : 'No record', true)}
      ${detailCard('LOA Status', officer ? loaStatusText(officer) : 'No record', true)}
      ${detailCard('Duty Status', response.shiftStatus?.onDuty ? formatCell('On Duty', 'Status') : formatCell('Off Duty', 'Status'), true)}
      ${detailCard('Discord ID', user.DiscordID || 'Not set')}
      ${detailCard('Unread notices', String(notifications.filter((item) => !item.ReadAt).length))}
    </section>
    ${officer ? tagList('Officer Tags', officer.Tags) : ''}
    ${officer ? trainingChecklist(officer.OfficerID, response.training || []) : ''}
    ${profileTable('My Rank History', response.rankChanges || [], ['ChangedAt', 'PreviousRank', 'NewRank', 'Reason', 'ChangedByName'])}
    ${profileTable('My Discipline', response.discipline || [], ['Type', 'Summary', 'IssuedAt', 'Status'])}
    ${profileTable('My LOA', response.loa || [], ['Officer', 'Rank', 'StartDate', 'EndDate', 'Status', 'ReviewReason'])}
    ${profileTable('My Transfer Requests', response.transfers || [], ['TargetDivision', 'TimeInMO8', 'Reason', 'HasPermission', 'Status', 'ReviewReason'])}
    ${profileTable('My Shift Activity', response.shifts || [], ['StartedAt', 'EndedAt', 'Status', 'Summary'])}
  `;
}

async function loadTasks() {
  await showViewOnly('tasks');
  const response = await apiCached('tasks', {});
  if (!response.ok) {
    document.querySelector('#tasksSummary').innerHTML = '';
    return renderTable('#tasksTable', [], ['Error'], { emptyMessage: response.error || 'Could not load tasks.' });
  }
  state.tasks = [
    ...(response.pendingLoa || []),
    ...(response.pendingTransfers || []),
  ];
  const counts = response.counts || {};
  document.querySelector('#tasksSummary').innerHTML = [
    stat('Pending LOA', counts.pendingLoa || 0),
    stat('Transfer Requests', counts.pendingTransfers || 0),
    stat('Total Tasks', counts.total || 0),
  ].join('');
  renderTable('#tasksTable', state.tasks, ['TaskType', 'Officer', 'Rank', 'StartDate', 'EndDate', 'TargetDivision', 'Reason'], {
    rowAction: (row) => row.TaskType === 'Transfer Request'
      ? `data-open-transfer-review="${escapeHtml(row.RequestID)}"`
      : `data-open-loa-review="${escapeHtml(row.RequestID)}"`,
    actions: (row) => row.TaskType === 'Transfer Request'
      ? `<button class="mini" data-open-transfer-review="${escapeHtml(row.RequestID)}">Review</button>`
      : `<button class="mini" data-open-loa-review="${escapeHtml(row.RequestID)}">Review</button>`,
  });
}

async function loadShift() {
  await showViewOnly('shift');
  const query = shiftQuery();
  const [statusResponse, teamResponse] = await Promise.all([
    apiCached('shiftStatus', {}),
    apiCached('teamShifts', query),
  ]);
  state.shiftStatus = statusResponse.ok ? statusResponse : null;
  state.shifts = teamResponse.ok ? teamResponse.recent || [] : [];

  const onDuty = Boolean(statusResponse.activeShift);
  document.querySelector('#startShiftButton').disabled = onDuty;
  document.querySelector('#endShiftButton').disabled = !onDuty;
  document.querySelector('#shiftSummary').innerHTML = [
    stat('Your Status', onDuty ? 'On Duty' : 'Off Duty'),
    stat('On Duty Now', teamResponse.active ? teamResponse.active.length : 0),
    stat('Recent Shifts', teamResponse.recent ? teamResponse.recent.length : 0),
  ].join('');

  renderTable('#activeShiftsTable', teamResponse.active || [], ['RobloxUsername', 'Callsign', 'Rank', 'LoaStatus', 'StartedAt']);
  renderTable('#recentShiftsTable', teamResponse.recent || [], ['RobloxUsername', 'Callsign', 'StartedAt', 'EndedAt', 'Duration', 'Summary']);
  renderTable('#shiftMetricsTable', teamResponse.metrics || [], ['RobloxUsername', 'Callsign', 'Rank', 'LoaStatus', 'Shifts', 'Duration', 'LastShift', 'ActivityFlag']);
  applyPermissions();
}

function shiftQuery() {
  const period = document.querySelector('#shiftPeriodFilter')?.value || 'week';
  const query = { Period: period };
  if (period === 'custom') {
    query.StartDate = document.querySelector('#shiftStartFilter')?.value || '';
    query.EndDate = document.querySelector('#shiftEndFilter')?.value || '';
  }
  return query;
}

async function loadOfficers() {
  await showViewOnly('officers');
  const response = await apiCached('listOfficers', {});
  if (!response.ok) return renderTable('#officersTable', [], ['Error'], response.error);
  state.officers = response.rows || [];
  renderOfficerTable();
}

function renderOfficerTable() {
  const query = document.querySelector('#officerSearch').value.toLowerCase();
  const rows = state.officers.filter((officer) => {
    return ['RobloxUsername', 'Callsign', 'Rank', 'Status', 'EffectiveStatus', 'LoaStatus', 'DutyStatus', 'Tags'].some((field) => String(officer[field] || '').toLowerCase().includes(query));
  });
  renderTable('#officersTable', rows, ['RobloxUsername', 'Callsign', 'Rank', 'EffectiveStatus', 'DutyStatus', 'Tags', 'JoinDate', 'UpdatedAt'], {
    rowAction: (row) => `data-open-officer="${escapeHtml(row.OfficerID)}"`,
  });
}

async function loadOfficerProfile(officerId) {
  await showViewOnly('officerProfile');
  const container = document.querySelector('#officerProfileView');
  if (!officerId) {
    container.innerHTML = emptyState('No officer selected.');
    return;
  }

  const requestOfficerId = officerId;
  container.innerHTML = loadingBlock('Loading officer profile...');
  const response = await apiCached('getOfficerProfile', { OfficerID: officerId });
  if (state.selectedOfficerId !== requestOfficerId) return;
  if (!response.ok) {
    container.innerHTML = emptyState(response.error || 'Officer not found.');
    return;
  }

  renderOfficerProfile(response);
}

async function loadTraining() {
  await showViewOnly('training');
  const [trainingResponse, officersResponse] = await Promise.all([
    apiCached('listTraining', {}),
    apiCached('listOfficers', {}),
  ]);
  const trainingRows = trainingResponse.rows || [];
  const officerRows = officersResponse.rows || [];
  state.training = trainingRows;
  state.trainingSummary = summarizeTraining(officerRows, trainingRows);
  renderTrainingOverview(state.trainingSummary);
  renderSearchableView('training');
}

async function loadRankChanges() {
  await showViewOnly('rankChanges');
  const response = await apiCached('rankChangeLog', {});
  if (!response.ok) {
    state.rankChanges = [];
    return renderTable('#rankChangesTable', [], ['Error'], { emptyMessage: response.error || 'Could not load rank changes.' });
  }
  state.rankChanges = response.rows || [];
  renderSearchableView('rankChanges');
}

async function loadDiscipline() {
  await showViewOnly('discipline');
  const response = await apiCached('listDiscipline', {});
  state.discipline = response.rows || [];
  renderSearchableView('discipline');
}

async function loadLoa() {
  await showViewOnly('loa');
  const response = await apiCached('listLoa', {});
  state.loa = response.rows || [];
  renderSearchableView('loa');
}

function renderLoaTable(rows) {
  renderTable('#loaTable', rows, ['Officer', 'Rank', 'StartDate', 'EndDate', 'Reason', 'Status', 'ReviewReason'], {
    actions: (row) => [
      can('CREATE_LOA') ? `<button class="mini" data-edit-loa="${escapeHtml(row.RequestID)}">Edit</button>` : '',
      can('APPROVE_LOA') && row.Status === 'Pending'
        ? `<button class="mini" data-open-loa-review="${escapeHtml(row.RequestID)}" data-status="Approved">Approve</button><button class="mini ghost" data-open-loa-review="${escapeHtml(row.RequestID)}" data-status="Denied">Deny</button>`
        : '',
      can('APPROVE_LOA') ? `<button class="mini ghost" data-delete-loa="${escapeHtml(row.RequestID)}">Delete</button>` : '',
    ].join(''),
  });
}

async function loadDocuments() {
  await showViewOnly('documents');
  const response = await apiCached('listDocuments', {});
  if (!response.ok) {
    state.documents = [];
    return renderTable('#documentsTable', [], ['Error'], { emptyMessage: response.error || 'Could not load documents.' });
  }
  state.documents = response.rows || [];
  renderDocumentTable();
}

async function loadAnnouncements() {
  await showViewOnly('announcements');
  const response = await apiCached('listAnnouncements', {});
  if (!response.ok) {
    state.announcements = [];
    document.querySelector('#announcementCards').innerHTML = emptyState(response.error || 'Could not load notices.');
    return renderTable('#announcementsTable', [], ['Error'], { emptyMessage: response.error || 'Could not load notices.' });
  }
  state.announcements = response.rows || [];
  renderSearchableView('announcements');
}

function renderDocumentTable() {
  const query = document.querySelector('#documentSearch').value.toLowerCase();
  const category = document.querySelector('#documentCategoryFilter').value;
  const rows = state.documents.filter((document) => {
    const matchesQuery = ['Title', 'Category', 'RequiredRole', 'RequiredTags', 'Status'].some((field) => String(document[field] || '').toLowerCase().includes(query));
    const matchesCategory = !category || document.Category === category;
    return matchesQuery && matchesCategory;
  });
  renderTable('#documentsTable', rows, ['Title', 'Category', 'RequiredRole', 'RequiredTags', 'Status', 'UpdatedAt', 'DriveURL'], {
    actions: (row) => can('MANAGE_DOCUMENTS')
      ? `<button class="mini" data-edit-document="${escapeHtml(row.DocumentID)}">Edit</button><button class="mini ghost" data-delete-document="${escapeHtml(row.DocumentID)}">Delete</button>`
      : '',
  });
}

function renderAnnouncementCards(rows) {
  const container = document.querySelector('#announcementCards');
  const published = rows.filter((row) => row.Status === 'Published');
  container.innerHTML = published.length
    ? published.slice(0, 6).map((notice) => `
      <article class="notice-card${truthy(notice.Pinned) ? ' pinned' : ''}">
        <div>
          <span>${truthy(notice.Pinned) ? 'Pinned' : escapeHtml(notice.Audience || 'All officers')}</span>
          <h3>${escapeHtml(notice.Title || 'Notice')}</h3>
          <p>${escapeHtml(notice.Body || '')}</p>
        </div>
        <small>${formatCell(notice.UpdatedAt, 'UpdatedAt')}</small>
      </article>
    `).join('')
    : `<p class="empty">No published notices.</p>`;
}

function renderAnnouncementsTable(rows) {
  renderTable('#announcementsTable', rows, ['Title', 'Audience', 'Status', 'Pinned', 'ExpiresAt', 'UpdatedAt'], {
    actions: (row) => can('MANAGE_ANNOUNCEMENTS')
      ? `<button class="mini" data-edit-announcement="${escapeHtml(row.AnnouncementID)}">Edit</button><button class="mini ghost" data-delete-announcement="${escapeHtml(row.AnnouncementID)}">Delete</button>`
      : '',
  });
}

function renderSearchableView(view) {
  const input = document.querySelector(`[data-search-view="${view}"]`);
  const query = input ? input.value.toLowerCase() : '';
  const rows = (state[view] || []).filter((row) => {
    if (!query) return true;
    return Object.values(row).some((value) => String(value || '').toLowerCase().includes(query));
  });

  if (view === 'training') {
    const summaryRows = (state.trainingSummary || []).filter((row) => {
      if (!query) return true;
      return Object.values(row).some((value) => String(value || '').toLowerCase().includes(query));
    });
    renderTable('#trainingTable', summaryRows, ['RobloxUsername', 'Callsign', 'Rank', 'DrivingStandard', 'SpecialistTickets', 'MissingTraining', 'ReviewDate']);
  }
  if (view === 'rankChanges') {
    renderTable('#rankChangesTable', rows, ['ChangedAt', 'RobloxUsername', 'PreviousRank', 'NewRank', 'Reason', 'ChangedByName']);
  }
  if (view === 'discipline') {
    renderTable('#disciplineTable', rows, ['Officer', 'Rank', 'Type', 'Summary', 'IssuedBy', 'IssuedAt', 'Status'], {
      actions: (row) => can('ADD_DISCIPLINE')
        ? `<button class="mini" data-edit-discipline="${escapeHtml(row.ActionID)}">Edit</button><button class="mini ghost" data-delete-discipline="${escapeHtml(row.ActionID)}">Delete</button>`
        : '',
    });
  }
  if (view === 'loa') {
    renderLoaTable(rows);
  }
  if (view === 'announcements') {
    renderAnnouncementCards(rows);
    renderAnnouncementsTable(rows);
  }
  if (view === 'users') {
    renderUsersTable(rows);
  }
  if (view === 'audit') {
    renderTable('#auditTable', rows, ['Timestamp', 'ActorUserID', 'Action', 'TargetType', 'TargetID']);
  }
}

async function loadUsers() {
  await showViewOnly('users');
  const response = await apiCached('listUsers', {});
  state.users = response.rows || [];
  renderSearchableView('users');
}

function renderUsersTable(rows) {
  renderTable('#usersTable', rows, ['RobloxUsername', 'DiscordID', 'Rank', 'Role', 'Status'], {
    actions: (row) => `<button class="mini" data-edit-user="${escapeHtml(row.UserID)}">Edit</button><button class="mini" data-reset-password="${escapeHtml(row.UserID)}">Reset password</button><button class="mini ghost" data-delete-user="${escapeHtml(row.UserID)}">Delete</button>`,
  });
}

async function loadPermissions() {
  await showViewOnly('permissions');
  const response = await apiCached('permissionsConfig', {});
  if (!response.ok) {
    document.querySelector('#permissionsMatrix').innerHTML = emptyState(response.error || 'Could not load permissions.');
    document.querySelector('#userPermissionsMatrix').innerHTML = '';
    return;
  }
  state.permissionConfig = response;
  renderPermissionsMatrix();
  renderUserPermissionsMatrix();
}

function renderPermissionsMatrix() {
  const config = state.permissionConfig;
  const rows = config.permissions.map((permission) => {
    const cells = config.roles.map((role) => {
      const enabled = rolePermissionEnabled(role, permission) ? ' checked' : '';
      const disabled = permission === 'FULL_ACCESS' ? ' disabled' : '';
      return `<td><input type="checkbox" data-role-permission data-role="${escapeHtml(role)}" data-permission="${escapeHtml(permission)}"${enabled}${disabled}></td>`;
    }).join('');
    return `<tr><td>${escapeHtml(permission)}</td>${cells}</tr>`;
  }).join('');
  document.querySelector('#permissionsMatrix').innerHTML = `
    <h3>Role permissions</h3>
    <div class="table-wrap compact">
      <table>
        <thead><tr><th>Permission</th>${config.roles.map((role) => `<th>${escapeHtml(role)}</th>`).join('')}</tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;
}

function renderUserPermissionsMatrix() {
  const config = state.permissionConfig;
  const rows = config.users.map((user) => {
    const cells = config.permissions.map((permission) => {
      const mode = userPermissionMode(user.UserID, permission);
      const options = USER_PERMISSION_MODES.map((item) => `<option value="${escapeHtml(item)}"${item === mode ? ' selected' : ''}>${escapeHtml(item)}</option>`).join('');
      return `<td><select data-user-permission data-user-id="${escapeHtml(user.UserID)}" data-permission="${escapeHtml(permission)}">${options}</select></td>`;
    }).join('');
    return `<tr><td>${escapeHtml(user.RobloxUsername)}</td><td>${escapeHtml(user.Role)}</td>${cells}</tr>`;
  }).join('');
  document.querySelector('#userPermissionsMatrix').innerHTML = `
    <h3>User-specific overrides</h3>
    <div class="table-wrap compact">
      <table>
        <thead><tr><th>User</th><th>Role</th>${config.permissions.map((permission) => `<th>${escapeHtml(permission)}</th>`).join('')}</tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
  `;
}

function rolePermissionEnabled(role, permission) {
  const explicit = state.permissionConfig.rolePermissions.find((row) => row.Role === role && row.Permission === permission);
  if (explicit) return String(explicit.Allowed).toUpperCase() === 'TRUE';
  return Boolean((state.permissionConfig.defaultPermissions?.[role] || []).includes(permission));
}

function userPermissionMode(userId, permission) {
  const override = state.permissionConfig.userPermissions.find((row) => row.UserID === userId && row.Permission === permission);
  if (!override) return 'Inherit';
  return String(override.Allowed).toUpperCase() === 'TRUE' ? 'Allow' : 'Deny';
}

async function loadAudit() {
  await showViewOnly('audit');
  const response = await apiCached('auditLog', {});
  state.audit = response.rows || [];
  renderSearchableView('audit');
}

function renderOfficerProfile(data) {
  const officer = data.officer;
  const container = document.querySelector('#officerProfileView');
  state.profileDiscipline = data.discipline || [];
  state.profileLoa = data.loa || [];
  container.innerHTML = `
    <div class="profile-head">
      <button class="ghost" data-view-link="officers">Back</button>
      <div>
        <h2>${escapeHtml(officer.RobloxUsername)}</h2>
        <p>${escapeHtml(officer.Callsign || 'No callsign')} / ${escapeHtml(officer.Rank || 'No rank')}</p>
      </div>
      <div class="profile-actions">
        <button data-edit-officer="${escapeHtml(officer.OfficerID)}" data-permission="EDIT_OFFICERS">Edit officer</button>
        <button class="ghost" data-delete-officer="${escapeHtml(officer.OfficerID)}" data-permission="ARCHIVE_OFFICERS">Delete officer</button>
        <button data-add-discipline="${escapeHtml(officer.OfficerID)}" data-permission="ADD_DISCIPLINE">Add discipline</button>
        <button data-add-loa="${escapeHtml(officer.OfficerID)}" data-permission="CREATE_LOA">Add LOA</button>
      </div>
    </div>

    <section class="profile-grid">
      ${detailCard('Status', formatCell(officer.Status), true)}
      ${detailCard('Duty Status', formatCell(officer.DutyStatus || 'Off Duty', 'Status'), true)}
      ${detailCard('Join date', officer.JoinDate || 'Not set')}
      ${detailCard('Discord ID', officer.DiscordID || 'Not set')}
      ${detailCard('Updated', officer.UpdatedAt || 'Not set')}
    </section>
    ${tagList('Officer Tags', officer.Tags)}

    ${trainingChecklist(officer.OfficerID, data.training)}

    <section class="profile-notes">
      <h3>Notes</h3>
      <p>${escapeHtml(officer.Notes || 'No notes recorded.')}</p>
    </section>

    <section class="profile-columns">
      ${profileTable('Training History', data.training, ['Standard', 'Status', 'Assessor', 'DateCompleted', 'ExpiryDate'])}
      ${profileTable('Rank History', data.rankChanges || [], ['ChangedAt', 'PreviousRank', 'NewRank', 'Reason', 'ChangedByName'])}
      ${profileTable('Discipline', data.discipline, ['Type', 'Summary', 'IssuedAt', 'Status'], {
        actions: (row) => can('ADD_DISCIPLINE') ? `<button class="mini" data-edit-discipline="${escapeHtml(row.ActionID)}">Edit</button><button class="mini ghost" data-delete-discipline="${escapeHtml(row.ActionID)}">Delete</button>` : '',
      })}
      ${profileTable('LOA', data.loa, ['Officer', 'Rank', 'StartDate', 'EndDate', 'Status', 'ReviewReason'], {
        actions: (row) => [
          can('CREATE_LOA') ? `<button class="mini" data-edit-loa="${escapeHtml(row.RequestID)}">Edit</button>` : '',
          can('APPROVE_LOA') ? `<button class="mini" data-open-loa-review="${escapeHtml(row.RequestID)}">Review</button>` : '',
          can('APPROVE_LOA') ? `<button class="mini ghost" data-delete-loa="${escapeHtml(row.RequestID)}">Delete</button>` : '',
        ].join(''),
      })}
      ${profileTable('Transfer Requests', data.transfers || [], ['TargetDivision', 'TimeInMO8', 'Reason', 'HasPermission', 'Status', 'ReviewReason'])}
      ${profileTable('Shift Activity', data.shifts || [], ['StartedAt', 'EndedAt', 'Status', 'Summary'])}
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
    field('RankChangeReason', 'Rank change reason', 'textarea', true),
    selectField('Status', 'Status', OFFICER_STATUSES, officer.Status || 'Active'),
    field('JoinDate', 'Join date', 'date', false, officer.JoinDate),
    checkboxGroupField('Tags', 'Officer tags', OFFICER_TAGS, officer.Tags),
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

function openDisciplineEditor(officerIdOrRecord) {
  const record = typeof officerIdOrRecord === 'object' ? officerIdOrRecord : {};
  const officerId = record.OfficerID || officerIdOrRecord || '';
  openEditor(record.ActionID ? 'Edit discipline record' : 'Add discipline record', [
    hiddenField('ActionID', record.ActionID),
    hiddenField('OfficerID', officerId),
    selectField('Type', 'Type', DISCIPLINE_TYPES, record.Type || 'Note'),
    field('Summary', 'Summary', 'text', false, record.Summary),
    field('Details', 'Details', 'textarea', true, record.Details),
    selectField('Status', 'Status', DISCIPLINE_STATUSES, record.Status || 'Active'),
  ], async (values) => api(values.ActionID ? 'saveDiscipline' : 'addDiscipline', values));
}

function openLoaEditor(officerIdOrRecord) {
  const record = typeof officerIdOrRecord === 'object' ? officerIdOrRecord : {};
  const officerId = record.OfficerID || officerIdOrRecord || '';
  openEditor(record.RequestID ? 'Edit LOA request' : 'Add LOA request', [
    hiddenField('RequestID', record.RequestID),
    hiddenField('OfficerID', officerId),
    field('StartDate', 'Start date', 'date', false, dateInputValue(record.StartDate)),
    field('EndDate', 'End date', 'date', false, dateInputValue(record.EndDate)),
    field('Reason', 'Reason', 'textarea', true, record.Reason),
    selectField('Status', 'Status', LOA_STATUSES, record.Status || 'Pending'),
  ], async (values) => api(values.RequestID ? 'saveLoa' : 'createLoa', values));
}

function openOwnLoaEditor() {
  openEditor('Request LOA', [
    field('StartDate', 'Start date', 'date'),
    field('EndDate', 'End date', 'date'),
    field('Reason', 'Reason', 'textarea', true),
  ], async (values) => api('requestOwnLoa', values), {
    successMessage: 'LOA request submitted for review.',
  });
}

function openTransferRequestEditor() {
  openEditor('Request transfer', [
    field('TimeInMO8', 'How long have you been in MO8?', 'text', false),
    field('TargetDivision', 'Division you wish to transfer to', 'text', false),
    field('Reason', 'Reason for transfer', 'textarea', true),
    selectField('HasPermission', 'Permission from receiving division OIC?', ['FALSE', 'TRUE'], 'FALSE'),
    field('Notes', 'Additional notes', 'textarea', true),
  ], async (values) => api('requestTransfer', values), {
    successMessage: 'Transfer request submitted for review.',
  });
}

function openTransferReviewEditor(record) {
  openEditor('Review transfer request', [
    hiddenField('RequestID', record.RequestID),
    field('Officer', 'Officer', 'text', false, record.Officer),
    field('TargetDivision', 'Target division', 'text', false, record.TargetDivision),
    field('TimeInMO8', 'Time in MO8', 'text', false, record.TimeInMO8),
    field('Reason', 'Request reason', 'textarea', true, record.Reason),
    field('Notes', 'Additional notes', 'textarea', true, record.Notes),
    selectField('Status', 'Decision', ['Approved', 'Denied'], 'Approved'),
    field('ReviewReason', 'Review reason', 'textarea', true, record.ReviewReason),
  ], async (values) => api('reviewTransfer', values), {
    successMessage: 'Transfer review saved.',
  });
}

async function startShift() {
  const response = await api('startShift', {});
  if (!response.ok) {
    showInfo('Shift start failed', `<p>${escapeHtml(response.error || 'Could not start shift.')}</p>`);
    return;
  }
  invalidateCache();
  await showView('shift');
}

function openEndShiftEditor() {
  const active = state.shiftStatus?.activeShift || {};
  openEditor('End shift', [
    field('EndedAt', 'End time', 'datetime-local', false, localDateTimeValue(active.EndedAt || new Date().toISOString())),
    field('Summary', 'Shift summary', 'textarea', true),
  ], async (values) => api('endShift', values), {
    successMessage: 'Shift ended.',
  });
}

function openLoaReviewEditor(record, status = '') {
  const currentDecision = ['Approved', 'Denied'].includes(status || record.Status) ? status || record.Status : 'Approved';
  openEditor('Review LOA request', [
    hiddenField('RequestID', record.RequestID),
    field('OfficerID', 'Officer ID', 'text', false, record.OfficerID),
    field('StartDate', 'Start date', 'date', false, dateInputValue(record.StartDate)),
    field('EndDate', 'End date', 'date', false, dateInputValue(record.EndDate)),
    field('Reason', 'Request reason', 'textarea', true, record.Reason),
    selectField('Status', 'Decision', ['Approved', 'Denied'], currentDecision),
    field('ReviewReason', 'Review reason', 'textarea', true, record.ReviewReason),
  ], async (values) => api('reviewLoa', values), {
    successMessage: 'LOA review saved.',
  });
}

function openDocumentEditor(document = {}) {
  openEditor(document.DocumentID ? 'Edit document' : 'Add document', [
    hiddenField('DocumentID', document.DocumentID),
    field('Title', 'Title', 'text', false, document.Title),
    selectField('Category', 'Category', ['Training', 'Policy', 'SOP', 'Form'], document.Category),
    field('DriveURL', 'Drive URL', 'url', false, document.DriveURL),
    selectField('RequiredRole', 'Minimum rank', ACCESS_LEVELS, document.RequiredRole || 'Police Constable'),
    checkboxGroupField('RequiredTags', 'Required tags', OFFICER_TAGS, document.RequiredTags),
    selectField('Status', 'Status', ['Published', 'Draft', 'Archived'], document.Status),
  ], async (values) => api('saveDocument', values));
}

function openAnnouncementEditor(announcement = {}) {
  openEditor(announcement.AnnouncementID ? 'Edit notice' : 'Add notice', [
    hiddenField('AnnouncementID', announcement.AnnouncementID),
    field('Title', 'Title', 'text', false, announcement.Title),
    field('Body', 'Notice text', 'textarea', true, announcement.Body),
    selectField('Audience', 'Minimum rank or role', ACCESS_LEVELS, announcement.Audience || 'Constable'),
    selectField('Status', 'Status', ANNOUNCEMENT_STATUSES, announcement.Status || 'Published'),
    selectField('Pinned', 'Pinned', ['FALSE', 'TRUE'], truthy(announcement.Pinned) ? 'TRUE' : 'FALSE'),
    field('ExpiresAt', 'Expires after', 'date', false, dateInputValue(announcement.ExpiresAt)),
  ], async (values) => api('saveAnnouncement', values));
}

function openUserEditor(user = {}) {
  openEditor(user.UserID ? 'Edit user' : 'Add user', [
    hiddenField('UserID', user.UserID),
    field('RobloxUsername', 'Roblox username', 'text', false, user.RobloxUsername),
    field('DiscordID', 'Discord ID', 'text', false, user.DiscordID),
    selectField('Rank', 'Rank', OFFICER_RANKS, user.Rank || 'Police Constable'),
    field('RankChangeReason', 'Rank change reason', 'textarea', true),
    selectField('Role', 'System role', SYSTEM_ROLES, user.Role || 'Constable'),
    selectField('Status', 'Status', ['Active', 'Suspended', 'Archived'], user.Status || 'Active'),
    field('TemporaryPassword', 'Temporary password', 'text', false),
  ], async (values) => api('saveUser', values), {
    successMessage: 'User saved. Copy the temporary password from the response if one was generated.',
  });
}

async function handleDocumentClick(event) {
  if (!event.target.closest('.notification-shell')) {
    closeNotificationMenu();
  }

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

  const deleteOfficer = event.target.closest('[data-delete-officer]');
  if (deleteOfficer) {
    await confirmDelete('Delete this officer and their linked user profile?', 'deleteOfficer', { OfficerID: deleteOfficer.dataset.deleteOfficer }, async () => {
      state.selectedOfficerId = '';
      await showView('officers');
    });
    return;
  }

  const addDiscipline = event.target.closest('[data-add-discipline]');
  if (addDiscipline) return openDisciplineEditor(addDiscipline.dataset.addDiscipline);

  const editDiscipline = event.target.closest('[data-edit-discipline]');
  if (editDiscipline) {
    const record = state.discipline.find((row) => row.ActionID === editDiscipline.dataset.editDiscipline)
      || state.profileDiscipline.find((row) => row.ActionID === editDiscipline.dataset.editDiscipline);
    if (record) openDisciplineEditor(record);
    return;
  }

  const deleteDiscipline = event.target.closest('[data-delete-discipline]');
  if (deleteDiscipline) {
    await confirmDelete('Delete this disciplinary record?', 'deleteDiscipline', { ActionID: deleteDiscipline.dataset.deleteDiscipline }, async () => {
      if (state.activeView === 'officerProfile') {
        await loadOfficerProfile(state.selectedOfficerId);
      } else {
        await loadDiscipline();
      }
    });
    return;
  }

  const addLoa = event.target.closest('[data-add-loa]');
  if (addLoa) return openLoaEditor(addLoa.dataset.addLoa);

  const requestLoa = event.target.closest('[data-request-loa]');
  if (requestLoa) return openOwnLoaEditor();

  const requestTransfer = event.target.closest('[data-request-transfer]');
  if (requestTransfer) return openTransferRequestEditor();

  const editLoa = event.target.closest('[data-edit-loa]');
  if (editLoa) {
    const record = state.loa.find((row) => row.RequestID === editLoa.dataset.editLoa)
      || state.profileLoa.find((row) => row.RequestID === editLoa.dataset.editLoa);
    if (record) openLoaEditor(record);
    return;
  }

  const reviewLoaOpen = event.target.closest('[data-open-loa-review]');
  if (reviewLoaOpen) {
    const record = state.tasks.find((row) => row.RequestID === reviewLoaOpen.dataset.openLoaReview)
      || state.loa.find((row) => row.RequestID === reviewLoaOpen.dataset.openLoaReview)
      || state.profileLoa.find((row) => row.RequestID === reviewLoaOpen.dataset.openLoaReview);
    if (record) openLoaReviewEditor(record, reviewLoaOpen.dataset.status || '');
    return;
  }

  const reviewTransferOpen = event.target.closest('[data-open-transfer-review]');
  if (reviewTransferOpen) {
    const record = state.tasks.find((row) => row.RequestID === reviewTransferOpen.dataset.openTransferReview);
    if (record) openTransferReviewEditor(record);
    return;
  }

  const deleteLoa = event.target.closest('[data-delete-loa]');
  if (deleteLoa) {
    await confirmDelete('Delete this LOA request?', 'deleteLoa', { RequestID: deleteLoa.dataset.deleteLoa }, async () => {
      if (state.activeView === 'tasks') {
        await loadTasks();
      } else if (state.activeView === 'officerProfile') {
        await loadOfficerProfile(state.selectedOfficerId);
      } else {
        await loadLoa();
      }
    });
    return;
  }

  const editDocument = event.target.closest('[data-edit-document]');
  if (editDocument) {
    const document = state.documents.find((row) => row.DocumentID === editDocument.dataset.editDocument);
    if (document) openDocumentEditor(document);
    return;
  }

  const deleteDocument = event.target.closest('[data-delete-document]');
  if (deleteDocument) {
    await confirmDelete('Delete this document link?', 'deleteDocument', { DocumentID: deleteDocument.dataset.deleteDocument }, loadDocuments);
    return;
  }

  const editAnnouncement = event.target.closest('[data-edit-announcement]');
  if (editAnnouncement) {
    const announcement = state.announcements.find((row) => row.AnnouncementID === editAnnouncement.dataset.editAnnouncement);
    if (announcement) openAnnouncementEditor(announcement);
    return;
  }

  const deleteAnnouncement = event.target.closest('[data-delete-announcement]');
  if (deleteAnnouncement) {
    await confirmDelete('Delete this notice?', 'deleteAnnouncement', { AnnouncementID: deleteAnnouncement.dataset.deleteAnnouncement }, loadAnnouncements);
    return;
  }

  const editUser = event.target.closest('[data-edit-user]');
  if (editUser) {
    const user = state.users.find((row) => row.UserID === editUser.dataset.editUser);
    if (user) openUserEditor(user);
    return;
  }

  const deleteUser = event.target.closest('[data-delete-user]');
  if (deleteUser) {
    await confirmDelete('Delete this user and their linked officer profile?', 'deleteUser', { UserID: deleteUser.dataset.deleteUser }, loadUsers);
    return;
  }

  const resetPassword = event.target.closest('[data-reset-password]');
  if (resetPassword) {
    const response = await api('resetUserPassword', { UserID: resetPassword.dataset.resetPassword });
    if (!response.ok) {
      showInfo('Password reset failed', `<p>${escapeHtml(response.error || 'Could not reset the password.')}</p>`);
      return;
    }
    showInfo('Temporary password', `
      <p>The account password has been reset.</p>
      <div class="temporary-password">${escapeHtml(response.temporaryPassword)}</div>
    `);
    return;
  }

  const reviewLoa = event.target.closest('[data-review-loa]');
  if (reviewLoa) return;
}

async function handleDocumentChange(event) {
  const rolePermission = event.target.closest('[data-role-permission]');
  if (rolePermission) {
    rolePermission.disabled = true;
    const response = await api('setRolePermission', {
      Role: rolePermission.dataset.role,
      Permission: rolePermission.dataset.permission,
      Allowed: rolePermission.checked,
    });
    if (!response.ok) {
      rolePermission.checked = !rolePermission.checked;
      showInfo('Permission update failed', `<p>${escapeHtml(response.error || 'Could not update role permission.')}</p>`);
    }
    invalidateCache();
    await loadPermissions();
    return;
  }

  const userPermission = event.target.closest('[data-user-permission]');
  if (userPermission) {
    userPermission.disabled = true;
    const response = await api('setUserPermission', {
      UserID: userPermission.dataset.userId,
      Permission: userPermission.dataset.permission,
      Mode: userPermission.value,
    });
    if (!response.ok) {
      showInfo('Permission update failed', `<p>${escapeHtml(response.error || 'Could not update user permission.')}</p>`);
    }
    invalidateCache();
    await loadPermissions();
    return;
  }

  const trainingToggle = event.target.closest('[data-training-toggle]');
  if (trainingToggle) {
    const originalChecked = !trainingToggle.checked;
    trainingToggle.disabled = true;
    const response = await api('setOfficerTraining', {
      OfficerID: trainingToggle.dataset.officerId,
      Standard: trainingToggle.dataset.standard,
      Enabled: trainingToggle.checked,
    });
    if (!response.ok) {
      trainingToggle.checked = originalChecked;
      trainingToggle.disabled = false;
      alert(response.error || 'Training update failed.');
      return;
    }
    invalidateCache();
    await loadOfficerProfile(trainingToggle.dataset.officerId);
    return;
  }

  const drivingSelect = event.target.closest('[data-driving-select]');
  if (drivingSelect) {
    drivingSelect.disabled = true;
    const response = await api('setDrivingStandard', {
      OfficerID: drivingSelect.dataset.officerId,
      Standard: drivingSelect.value,
    });
    if (!response.ok) {
      drivingSelect.disabled = false;
      alert(response.error || 'Driving standard update failed.');
      return;
    }
    invalidateCache();
    await loadOfficerProfile(drivingSelect.dataset.officerId);
    return;
  }

  const reviewDate = event.target.closest('[data-training-review]');
  if (reviewDate) {
    reviewDate.disabled = true;
    const response = await api('setTrainingReviewDate', {
      OfficerID: reviewDate.dataset.officerId,
      ReviewDate: reviewDate.value,
    });
    if (!response.ok) {
      reviewDate.disabled = false;
      showInfo('Review date failed', `<p>${escapeHtml(response.error || 'Training review date update failed.')}</p>`);
      return;
    }
    invalidateCache();
    await loadOfficerProfile(reviewDate.dataset.officerId);
    return;
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

function dashboardPanel(title, rows, columns) {
  const body = rows.length
    ? rows.map((row) => `
      <article class="dashboard-row">
        ${columns.map((column) => `
          <span>
            <small>${escapeHtml(column)}</small>
            <strong>${formatCell(row[column], column) || '&nbsp;'}</strong>
          </span>
        `).join('')}
      </article>
    `).join('')
    : `<p class="empty">No records found.</p>`;

  return `
    <article class="dashboard-panel">
      <h3>${escapeHtml(title)}</h3>
      <div class="dashboard-list">${body}</div>
    </article>
  `;
}

function announcementPanel(title, rows) {
  const body = rows.length
    ? rows.map((row) => `
      <article class="dashboard-row notice-dashboard">
        <span>
          <small>${truthy(row.Pinned) ? 'Pinned notice' : escapeHtml(row.Audience || 'Notice')}</small>
          <strong>${escapeHtml(row.Title || 'Notice')}</strong>
          <em>${escapeHtml(row.Body || '')}</em>
        </span>
      </article>
    `).join('')
    : `<p class="empty">No notices found.</p>`;

  return `
    <article class="dashboard-panel">
      <h3>${escapeHtml(title)}</h3>
      <div class="dashboard-list">${body}</div>
    </article>
  `;
}

function loaStatusText(officer) {
  if (!officer || officer.LoaStatus !== 'On LOA') return formatCell('Available', 'Status');
  const endDate = officer.CurrentLoaEnd ? ` until ${formatDisplayDate(officer.CurrentLoaEnd)}` : '';
  return `<span class="pill warning">On LOA${escapeHtml(endDate)}</span>`;
}

function detailCard(label, value, allowHtml = false) {
  const content = allowHtml ? value : escapeHtml(value);
  return `<article class="detail-card"><span>${escapeHtml(label)}</span><strong>${content}</strong></article>`;
}

function profileTable(title, rows, columns, options = {}) {
  const actionHeader = options.actions ? '<th>Actions</th>' : '';
  const body = rows.length
    ? rows.map((row) => {
      const actionCell = options.actions ? `<td class="actions">${options.actions(row)}</td>` : '';
      return `<tr>${columns.map((column) => `<td>${formatCell(row[column], column)}</td>`).join('')}${actionCell}</tr>`;
    }).join('')
    : `<tr><td colspan="${columns.length + (options.actions ? 1 : 0)}">No records found.</td></tr>`;
  return `
    <section class="profile-panel">
      <h3>${escapeHtml(title)}</h3>
      <div class="table-wrap compact">
        <table>
          <thead><tr>${columns.map((column) => `<th>${escapeHtml(column)}</th>`).join('')}${actionHeader}</tr></thead>
          <tbody>${body}</tbody>
        </table>
      </div>
    </section>
  `;
}

function tagList(title, value) {
  const tags = splitTags(value);
  return `
    <section class="tag-panel">
      <h3>${escapeHtml(title)}</h3>
      <div class="tag-list">${tags.length ? tags.map((tag) => `<span>${escapeHtml(tag)}</span>`).join('') : '<p class="empty">No tags assigned.</p>'}</div>
    </section>
  `;
}

function trainingChecklist(officerId, trainingRows) {
  const rows = SPECIALIST_TRAINING.map((standard) => {
    const record = trainingRows.find((item) => item.Standard === standard && String(item.Status) === 'Passed');
    const checked = record ? ' checked' : '';
    const disabled = can('MANAGE_TRAINING') ? '' : ' disabled';
    return `
      <label class="training-check">
        <input type="checkbox" data-training-toggle data-officer-id="${escapeHtml(officerId)}" data-standard="${escapeHtml(standard)}"${checked}${disabled}>
        <span>${escapeHtml(standard)}</span>
      </label>
    `;
  }).join('');
  const drivingRecord = DRIVING_STANDARDS.find((standard) => {
    return trainingRows.some((item) => item.Standard === standard && String(item.Status) === 'Passed');
  }) || '';
  const drivingOptions = [''].concat(DRIVING_STANDARDS).map((standard) => {
    const label = standard || 'No driving standard';
    const selected = standard === drivingRecord ? ' selected' : '';
    return `<option value="${escapeHtml(standard)}"${selected}>${escapeHtml(label)}</option>`;
  }).join('');
  const disabled = can('MANAGE_TRAINING') ? '' : ' disabled';
  const reviewDate = trainingRows.find((item) => item.ReviewDate)?.ReviewDate || trainingRows.find((item) => item.UpdatedAt)?.ReviewDate || '';

  return `
    <section class="cert-panel">
      <div>
        <h3>Training Certifications</h3>
        <p>Sergeants and above can assign specialist tickets and one driving standard.</p>
      </div>
      <div class="cert-grid">${rows}</div>
      <label class="driving-select">
        Driving standard
        <select data-driving-select data-officer-id="${escapeHtml(officerId)}"${disabled}>
          ${drivingOptions}
        </select>
      </label>
      <label class="driving-select">
        Training review date
        <input type="date" data-training-review data-officer-id="${escapeHtml(officerId)}" value="${escapeHtml(reviewDate)}"${disabled}>
      </label>
    </section>
  `;
}

function summarizeTraining(officers, trainingRows) {
  return officers.map((officer) => {
    const records = trainingRows.filter((item) => item.OfficerID === officer.OfficerID);
    const passed = (standard) => records.some((item) => item.Standard === standard && item.Status === 'Passed');
    const specialistTickets = SPECIALIST_TRAINING.filter(passed);
    const drivingStandard = DRIVING_STANDARDS.find(passed) || 'Not set';
    const missing = [
      ...SPECIALIST_TRAINING.filter((standard) => !passed(standard)),
      drivingStandard === 'Not set' ? 'Driving standard' : '',
    ].filter(Boolean);
    const reviewDate = records.find((item) => item.ReviewDate)?.ReviewDate || '';
    return {
      OfficerID: officer.OfficerID,
      RobloxUsername: officer.RobloxUsername,
      Callsign: officer.Callsign,
      Rank: officer.Rank,
      DrivingStandard: drivingStandard,
      SpecialistTickets: specialistTickets.length ? specialistTickets.join(', ') : 'None',
      MissingTraining: missing.length ? `${missing.length} missing` : 'Complete',
      MissingDetails: missing.join(', '),
      ReviewDate: reviewDate,
    };
  });
}

function renderTrainingOverview(rows) {
  const container = document.querySelector('#trainingMatrix');
  if (!rows.length) {
    container.innerHTML = `<p class="empty">Training overview will appear once officers have records.</p>`;
    return;
  }

  const complete = rows.filter((row) => row.MissingTraining === 'Complete').length;
  const needsReview = rows.filter((row) => row.ReviewDate).length;
  const noDriving = rows.filter((row) => row.DrivingStandard === 'Not set').length;
  const cards = [
    stat('Officers tracked', rows.length),
    stat('Training complete', complete),
    stat('No driving standard', noDriving),
    stat('Review dates set', needsReview),
  ].join('');

  container.innerHTML = `
    <h3>Training Overview</h3>
    <div class="training-overview">
      ${cards}
      <p>Select an officer profile to update specialist tickets, driving standard, and review date.</p>
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
    return `<tr ${attrs}>${columns.map((column) => `<td>${formatCell(row[column], column)}</td>`).join('')}${actionCell}</tr>`;
  }).join('');
  table.innerHTML = `${head}<tbody>${body}</tbody>`;
}

function renderError(container, message) {
  container.innerHTML = `<article class="stat"><strong>!</strong><span>${escapeHtml(message || 'Something went wrong.')}</span></article>`;
}

async function toggleNotifications(event) {
  event.stopPropagation();
  if (!elements.notificationMenu.hidden) {
    closeNotificationMenu();
    return;
  }

  elements.notificationMenu.hidden = false;
  const cached = getCachedResponse('listNotifications', {});
  if (cached) {
    renderNotificationMenu(cached);
  } else {
    elements.notificationMenu.innerHTML = loadingBlock('Loading notifications...');
  }

  const response = cached ? await api('listNotifications', {}) : await apiCached('listNotifications', {});
  if (!response.ok) {
    if (!cached) elements.notificationMenu.innerHTML = `<p class="empty">${escapeHtml(response.error || 'Could not load notifications.')}</p>`;
    return;
  }

  setCachedResponse('listNotifications', {}, response);
  renderNotificationMenu(response);
  if ((response.unread || 0) > 0) {
    await api('markNotificationsRead', {});
    const readResponse = {
      ...response,
      unread: 0,
      rows: (response.rows || []).map((notice) => ({ ...notice, ReadAt: notice.ReadAt || new Date().toISOString() })),
    };
    setCachedResponse('listNotifications', {}, readResponse);
    state.unreadNotifications = 0;
    updateNotificationBadge();
    invalidateCache('myProfile');
  }
}

function renderNotificationMenu(response) {
  const rows = response.rows || [];
  elements.notificationMenu.innerHTML = `
    <div class="notification-menu-head">
      <strong>Notifications</strong>
      <span>${escapeHtml(String(response.unread || 0))} unread</span>
    </div>
    ${rows.length
    ? `<div class="notice-list">${rows.map((notice) => `
      <article class="notice-item${notice.ReadAt ? '' : ' unread'}${importantNotice(notice) ? ' important' : ''}${positiveNotice(notice) ? ' positive' : ''}">
        <div>
          <strong>${escapeHtml(notice.Title || 'Notification')}</strong>
          <p>${escapeHtml(notice.Message || '')}</p>
        </div>
        <span>${formatCell(notice.CreatedAt || '', 'CreatedAt')}</span>
      </article>
    `).join('')}</div>`
    : `<p class="empty">No notifications yet.</p>`}
  `;
}

function closeNotificationMenu() {
  elements.notificationMenu.hidden = true;
}

function importantNotice(notice) {
  const text = `${notice.Title || ''} ${notice.Message || ''}`.toLowerCase();
  return ['disciplinary', 'discipline', 'denied', 'removed', 'suspended'].some((word) => text.includes(word));
}

function positiveNotice(notice) {
  const text = `${notice.Title || ''} ${notice.Message || ''}`.toLowerCase();
  return ['approved', 'passed', 'created', 'completed', 'published'].some((word) => text.includes(word));
}

async function confirmDelete(message, action, payload, onSuccess) {
  if (!window.confirm(message)) return;
  const response = await api(action, payload);
  if (!response.ok) {
    showInfo('Delete failed', `<p>${escapeHtml(response.error || 'The record could not be deleted.')}</p>`);
    return;
  }
  invalidateCache();
  await onSuccess();
}

function showInfo(title, content) {
  elements.infoTitle.textContent = title;
  elements.infoContent.innerHTML = content;
  elements.infoDialog.showModal();
}

function emptyState(message) {
  return `<section class="data-view"><p class="empty">${escapeHtml(message)}</p></section>`;
}

function loadingBlock(message) {
  return `
    <section class="loading-panel">
      <span></span>
      <p>${escapeHtml(message)}</p>
    </section>
  `;
}

function formatCell(value, column = '') {
  const text = value === undefined || value === null ? '' : String(value);
  if (isDateTimeColumn(column) && text) {
    return escapeHtml(formatDisplayDateTime(text));
  }
  if (isDateColumn(column) && text) {
    return escapeHtml(formatDisplayDate(text));
  }
  if (text.startsWith('https://')) {
    return `<a href="${escapeHtml(text)}" target="_blank" rel="noopener">Open</a>`;
  }
  if (['Active', 'Published', 'Passed', 'Approved'].includes(text)) {
    return `<span class="pill success">${escapeHtml(text)}</span>`;
  }
  if (['On Duty'].includes(text)) {
    return `<span class="pill success">${escapeHtml(text)}</span>`;
  }
  if (['LOA', 'On LOA', 'Pending', 'In Progress', 'Draft', 'Not Started'].includes(text)) {
    return `<span class="pill warning">${escapeHtml(text)}</span>`;
  }
  if (['Off Duty', 'Suspended', 'Archived', 'Failed', 'Denied', 'Expired', 'Removed', 'Low activity', 'No activity'].includes(text)) {
    return `<span class="pill danger">${escapeHtml(text)}</span>`;
  }
  return escapeHtml(text);
}

function isDateColumn(column) {
  return ['StartDate', 'EndDate', 'JoinDate', 'DateCompleted', 'ExpiryDate', 'ReviewDate', 'ExpiresAt', 'CurrentLoaEnd'].includes(column);
}

function isDateTimeColumn(column) {
  return ['UpdatedAt', 'CreatedAt', 'IssuedAt', 'ReviewedAt', 'ReadAt', 'Timestamp', 'LastLogin', 'ChangedAt', 'StartedAt', 'EndedAt', 'LastShift'].includes(column);
}

function formatDisplayDate(value) {
  const input = String(value || '').trim();
  const isoMatch = input.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) return `${isoMatch[3]}/${isoMatch[2]}/${isoMatch[1]}`;

  const date = new Date(input);
  if (!Number.isNaN(date.getTime())) {
    return `${String(date.getDate()).padStart(2, '0')}/${String(date.getMonth() + 1).padStart(2, '0')}/${date.getFullYear()}`;
  }
  return input;
}

function formatDisplayDateTime(value) {
  const input = String(value || '').trim();
  const date = new Date(input);
  if (!Number.isNaN(date.getTime())) {
    return `${String(date.getDate()).padStart(2, '0')}/${String(date.getMonth() + 1).padStart(2, '0')}/${date.getFullYear()} ${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`;
  }
  const isoMatch = input.match(/^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2})/);
  if (isoMatch) return `${isoMatch[3]}/${isoMatch[2]}/${isoMatch[1]} ${isoMatch[4]}:${isoMatch[5]}`;
  return formatDisplayDate(input);
}

function dateInputValue(value) {
  const input = String(value || '').trim();
  const isoMatch = input.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) return `${isoMatch[1]}-${isoMatch[2]}-${isoMatch[3]}`;

  const ukMatch = input.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (ukMatch) return `${ukMatch[3]}-${ukMatch[2]}-${ukMatch[1]}`;

  const date = new Date(input);
  if (!Number.isNaN(date.getTime())) {
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
  }
  return '';
}

function localDateTimeValue(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return '';
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}T${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`;
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
    const values = formValues(elements.editorForm);
    const response = await onSubmit(values);
    if (!response.ok) {
      elements.editorStatus.textContent = response.error || 'Save failed.';
      return;
    }

    const generatedPassword = response.temporaryPassword ? ` Temporary password: ${response.temporaryPassword}` : '';
    if (options.successMessage || generatedPassword) {
      showInfo('Saved', `<p>${escapeHtml(options.successMessage || 'Saved.')}</p>${generatedPassword ? `<div class="temporary-password">${escapeHtml(generatedPassword.replace(' Temporary password: ', ''))}</div>` : ''}`);
    }

    elements.editorDialog.close();
    invalidateCache();
    await showView(state.activeView);
  };
}

function formValues(form) {
  const data = new FormData(form);
  const values = {};
  data.forEach((value, key) => {
    if (values[key]) {
      values[key] = `${values[key]}, ${value}`;
    } else {
      values[key] = value;
    }
  });
  return values;
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

function checkboxGroupField(name, label, options, selected = '') {
  const selectedTags = splitTags(selected);
  const checkboxes = options.map((option) => {
    const checked = selectedTags.includes(option) ? ' checked' : '';
    return `
      <label class="training-check">
        <input type="checkbox" name="${escapeHtml(name)}" value="${escapeHtml(option)}"${checked}>
        <span>${escapeHtml(option)}</span>
      </label>
    `;
  }).join('');
  return { html: `<fieldset class="wide checkbox-group"><legend>${escapeHtml(label)}</legend>${checkboxes}</fieldset>` };
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

async function apiCached(action, data = {}, includeToken = true) {
  const key = cacheKey(action, data, includeToken);
  const cached = state.cache[key];
  if (cached && Date.now() - cached.time < CACHE_TTL_MS) {
    return cached.response;
  }

  const response = await api(action, data, includeToken);
  if (response.ok) {
    state.cache[key] = { time: Date.now(), response };
    storeCache();
  }
  return response;
}

function cacheKey(action, data, includeToken) {
  return JSON.stringify({ action, data, includeToken });
}

function getCachedResponse(action, data = {}, includeToken = true) {
  const cached = state.cache[cacheKey(action, data, includeToken)];
  if (!cached || Date.now() - cached.time >= CACHE_TTL_MS) return null;
  return cached.response;
}

function setCachedResponse(action, data = {}, response, includeToken = true) {
  state.cache[cacheKey(action, data, includeToken)] = { time: Date.now(), response };
  storeCache();
}

function loadStoredCache() {
  try {
    const parsed = JSON.parse(sessionStorage.getItem(CACHE_STORAGE_KEY) || '{}');
    return Object.keys(parsed).reduce((cache, key) => {
      if (parsed[key] && Date.now() - parsed[key].time < CACHE_TTL_MS) cache[key] = parsed[key];
      return cache;
    }, {});
  } catch (error) {
    sessionStorage.removeItem(CACHE_STORAGE_KEY);
    return {};
  }
}

function storeCache() {
  try {
    sessionStorage.setItem(CACHE_STORAGE_KEY, JSON.stringify(state.cache));
  } catch (error) {
    // Storage can fill up in older browsers; the app still works with memory cache only.
  }
}

function loadSessionAuth() {
  try {
    const parsed = JSON.parse(sessionStorage.getItem(SESSION_STORAGE_KEY) || '{}');
    if (!parsed.time || Date.now() - parsed.time > CACHE_TTL_MS) return null;
    return parsed;
  } catch (error) {
    sessionStorage.removeItem(SESSION_STORAGE_KEY);
    return null;
  }
}

function storeSessionAuth(user, permissions) {
  try {
    sessionStorage.setItem(SESSION_STORAGE_KEY, JSON.stringify({ time: Date.now(), user, permissions }));
  } catch (error) {
    // Non-critical; refresh simply falls back to normal startup.
  }
}

function clearCacheForNewVersion() {
  const storedVersion = sessionStorage.getItem(VERSION_STORAGE_KEY);
  if (storedVersion === APP_VERSION) return;
  sessionStorage.removeItem(CACHE_STORAGE_KEY);
  sessionStorage.removeItem(BOOT_STORAGE_KEY);
  sessionStorage.removeItem(SESSION_STORAGE_KEY);
  sessionStorage.setItem(VERSION_STORAGE_KEY, APP_VERSION);
  state.cache = {};
}

function invalidateCache(action = '') {
  if (!action) {
    state.cache = {};
    sessionStorage.removeItem(CACHE_STORAGE_KEY);
    sessionStorage.removeItem(BOOT_STORAGE_KEY);
    return;
  }
  Object.keys(state.cache).forEach((key) => {
    if (key.includes(`"action":"${action}"`)) delete state.cache[key];
  });
  storeCache();
}

function can(permission) {
  return state.permissions.includes('FULL_ACCESS') || state.permissions.includes(permission);
}

function truthy(value) {
  return value === true || String(value).toUpperCase() === 'TRUE';
}

function splitTags(value) {
  return String(value || '').split(/[,\n;]+/).map((tag) => tag.trim()).filter(Boolean);
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

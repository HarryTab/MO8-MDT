const API_URL = 'https://script.google.com/macros/s/AKfycbwsRocB7bsQLfXiazKGI-O158ppsRnQPVsrtvzVaoyUUgMdanidkOJc_pg--lddbDGPhQ/exec';
const APP_VERSION = '2026-05-06-2';

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

const SYSTEM_ROLES = ['Constable', 'Trainer', 'Sergeant', 'Inspector', 'Chief Inspector', 'Command'];
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
const DEVELOPMENT_CATEGORIES = ['Development', 'Training', 'Activity', 'Conduct', 'Career', 'Other'];
const DEVELOPMENT_STATUSES = ['Open', 'In Progress', 'Completed', 'Paused'];
const DASHBOARD_WIDGETS = [
  ['activeLoa', 'Active LOA Status'],
  ['pendingLoa', 'Pending LOA'],
  ['announcements', 'Notice Board'],
  ['trainingReviews', 'Training Reviews'],
  ['recentDocuments', 'Recent Documents'],
  ['recentActivity', 'Recent Activity'],
  ['pendingAppeals', 'Pending Appeals'],
  ['unassignedOfficers', 'Unassigned Officers'],
  ['lowActivity', 'Low Activity'],
  ['documentAcknowledgements', 'Document Acknowledgements'],
];

const state = {
  token: localStorage.getItem('mo8_token') || '',
  user: null,
  permissions: [],
  unreadNotifications: 0,
  activeView: 'dashboard',
  officers: [],
  training: [],
  trainingSummary: [],
  trainingOptions: [],
  courses: [],
  courseBookings: [],
  discipline: [],
  loa: [],
  tasks: [],
  profileAppeals: [],
  profileDiscipline: [],
  profileLoa: [],
  profileSupervisorRequests: [],
  profileCheckins: [],
  profileDevelopmentPlans: [],
  documents: [],
  documentFolder: '',
  announcements: [],
  rankChanges: [],
  shifts: [],
  shiftStatus: null,
  users: [],
  supervisorOptions: [],
  supervisorDashboard: null,
  permissionConfig: null,
  audit: [],
  cache: loadStoredCache(),
  selectedOfficerId: '',
  selectedBulkOfficerIds: [],
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
document.addEventListener('change', handleBulkOfficerSelection);

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
document.querySelector('#documentCategoryFilter').addEventListener('change', (event) => {
  state.documentFolder = event.target.value;
  renderDocumentTable();
});
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
document.querySelector('#bulkOfficerButton').addEventListener('click', () => openBulkOfficerEditor());
document.querySelector('#newDocumentButton').addEventListener('click', () => openDocumentEditor());
document.querySelector('#newTrainingOptionButton').addEventListener('click', () => openTrainingOptionEditor());
document.querySelector('#newCourseButton').addEventListener('click', () => openCourseEditor());
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
  if (can('VIEW_TASKS')) tasks.push({ label: 'Preparing supervisor dashboard', run: () => apiCached('supervisorDashboard', {}) });
  if (can('VIEW_COURSES')) tasks.push({ label: 'Loading training courses', run: () => apiCached('listTrainingCourses', {}) });
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
    can('VIEW_COURSES') ? ['listTrainingCourses', {}] : null,
    can('VIEW_LOA') ? ['listLoa', {}] : null,
    can('VIEW_RANK_LOG') ? ['rankChangeLog', {}] : null,
    ['teamShifts', { Period: 'week' }],
    can('VIEW_TASKS') ? ['supervisorDashboard', {}] : null,
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
  document.body.classList.toggle('is-officer-portal', isOfficerPortal());
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

function isOfficerPortal() {
  return !can('VIEW_DASHBOARD') && !can('VIEW_OFFICERS') && !can('VIEW_TASKS');
}

function applyPermissions() {
  document.querySelectorAll('[data-permission]').forEach((node) => {
    node.hidden = !can(node.dataset.permission);
  });
  document.querySelectorAll('.nav-group').forEach((group) => {
    group.hidden = !group.querySelector('.nav-item:not([hidden])');
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
    supervisor: ['Supervisor', 'Assigned officers, check-ins, development plans and workload'],
    officers: ['Officers', 'MO8 officer database'],
    officerProfile: ['Officer Profile', 'Individual record and linked history'],
    rankChanges: ['Rank Change Log', 'Promotion and rank movement history'],
    training: ['Training', 'Training standards and status'],
    courses: ['Training Courses', 'Course bookings, waitlists and trainer outcomes'],
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
    supervisor: loadSupervisor,
    officers: loadOfficers,
    officerProfile: () => loadOfficerProfile(state.selectedOfficerId),
    rankChanges: loadRankChanges,
    training: loadTraining,
    courses: loadCourses,
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
  if (view === 'supervisor') {
    document.querySelector('#supervisorSummary').innerHTML = '';
    ['#supervisorAssignedTable', '#supervisorRequestsTable', '#supervisorUnassignedTable', '#supervisorWorkloadTable', '#supervisorPlansTable', '#supervisorCheckinsTable'].forEach((selector) => {
      document.querySelector(selector).innerHTML = `<tbody><tr><td>${loadingBlock('Loading supervisor dashboard...')}</td></tr></tbody>`;
    });
    return;
  }
  if (view === 'training') document.querySelector('#trainingMatrix').innerHTML = '';
  if (view === 'courses') {
    document.querySelector('#coursesTable').innerHTML = `<tbody><tr><td>${loadingBlock('Loading training courses...')}</td></tr></tbody>`;
    document.querySelector('#courseBookingsTable').innerHTML = '';
    return;
  }
  if (view === 'documents') {
    document.querySelector('#documentExplorer').innerHTML = loadingBlock('Loading documents...');
    return;
  }
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
    supervisor: 'supervisorDashboard',
    officers: 'listOfficers',
    rankChanges: 'rankChangeLog',
    training: 'listTraining',
    courses: 'listTrainingCourses',
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
  const activeWidgets = response.widgets || DASHBOARD_WIDGETS.map(([key]) => key);
  const widget = (key, html) => activeWidgets.includes(key) ? html : '';
  elements.dashboardView.innerHTML = `
    <div class="section-head dashboard-config">
      <h2>Dashboard</h2>
      <button class="ghost" data-configure-dashboard>Widgets</button>
    </div>
    <div class="stat-row">
      ${[
    stat('Active Officers', counts.activeOfficers || 0),
    stat('Currently On LOA', counts.currentlyOnLoa || 0),
    stat('Pending LOA', counts.loaPending || 0),
    stat('Pending Appeals', counts.pendingAppeals || 0),
    stat('Review Due', counts.trainingReviewsDue || 0),
    stat('Docs To Ack', counts.pendingAcknowledgements || 0),
  ].join('')}
    </div>
    <section class="dashboard-grid">
      ${widget('activeLoa', dashboardPanel('Active LOA Status', response.activeLoa || [], ['Officer', 'Rank', 'EndDate', 'Status']))}
      ${widget('pendingLoa', dashboardPanel('Pending LOA', response.pendingLoa || [], ['Officer', 'Rank', 'StartDate', 'EndDate']))}
      ${widget('announcements', announcementPanel('Notice Board', response.announcements || []))}
      ${widget('trainingReviews', dashboardPanel('Training Reviews', response.trainingReviewsDue || [], ['RobloxUsername', 'Standard', 'ReviewDate', 'UpdatedBy']))}
      ${widget('recentDocuments', dashboardPanel('Recent Documents', response.recentDocuments || [], ['Title', 'Category', 'RequiredRole', 'UpdatedAt']))}
      ${widget('recentActivity', dashboardPanel('Recent Activity', response.recentAudit || [], ['Timestamp', 'Action', 'TargetType', 'TargetID']))}
      ${widget('pendingAppeals', dashboardPanel('Pending Appeals', response.pendingAppeals || [], ['Officer', 'Rank', 'SourceType', 'Reason']))}
      ${widget('unassignedOfficers', dashboardPanel('Unassigned Officers', response.unassignedOfficers || [], ['RobloxUsername', 'Rank', 'DutyStatus']))}
      ${widget('lowActivity', dashboardPanel('Low Activity', response.lowActivity || [], ['RobloxUsername', 'Rank', 'Duration', 'ActivityFlag']))}
      ${widget('documentAcknowledgements', dashboardPanel('Document Acknowledgements', response.documentAcknowledgements || [], ['Title', 'Category', 'RequiredRole']))}
    </section>
  `;
}

async function loadMyProfile() {
  await showViewOnly('myProfile');
  const [response, optionsResponse] = await Promise.all([
    apiCached('myProfile', {}),
    apiCached('listTrainingOptions', {}),
  ]);
  if (optionsResponse.ok) state.trainingOptions = optionsResponse.rows || [];
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
        <button data-request-supervisor>Contact supervisor</button>
      </div>
    </div>
    <section class="profile-grid">
      ${detailCard('Callsign', officer ? officer.Callsign || 'Not set' : 'No officer record')}
      ${detailCard('Status', officer ? formatCell(officer.EffectiveStatus || officer.Status, 'Status') : 'No record', true)}
      ${detailCard('LOA Status', officer ? loaStatusText(officer) : 'No record', true)}
      ${detailCard('Duty Status', response.shiftStatus?.onDuty ? formatCell('On Duty', 'Status') : formatCell('Off Duty', 'Status'), true)}
      ${detailCard('Supervisor', officer ? officer.Supervisor || 'Not assigned' : 'No record')}
      ${detailCard('Discord ID', user.DiscordID || 'Not set')}
      ${detailCard('Unread notices', String(notifications.filter((item) => !item.ReadAt).length))}
    </section>
    ${officer ? tagList('Officer Tags', officer.Tags) : ''}
    ${officer ? trainingChecklist(officer.OfficerID, response.training || []) : ''}
    ${profileTable('My Rank History', response.rankChanges || [], ['ChangedAt', 'PreviousRank', 'NewRank', 'Reason', 'ChangedByName'])}
    ${profileTable('My Discipline', response.discipline || [], ['Type', 'Summary', 'IssuedAt', 'Status'])}
    ${profileTable('My LOA', response.loa || [], ['Officer', 'Rank', 'StartDate', 'EndDate', 'Status', 'ReviewReason'], {
    actions: (row) => row.Status === 'Denied' ? `<button class="mini" data-request-appeal-source="LOA" data-request-appeal-id="${escapeHtml(row.RequestID)}">Appeal</button>` : '',
  })}
    ${profileTable('My Transfer Requests', response.transfers || [], ['TargetDivision', 'TimeInMO8', 'Reason', 'HasPermission', 'Status', 'ReviewReason'], {
    actions: (row) => row.Status === 'Denied' ? `<button class="mini" data-request-appeal-source="Transfer" data-request-appeal-id="${escapeHtml(row.RequestID)}">Appeal</button>` : '',
  })}
    ${profileTable('My Supervisor Requests', response.supervisorRequests || [], ['Category', 'Subject', 'Details', 'Supervisor', 'Status', 'ReviewReason'])}
    ${profileTable('My Appeals / Reviews', response.appeals || [], ['SourceType', 'SourceID', 'Reason', 'Status', 'ReviewReason'])}
    ${profileTable('My Development Plans', response.developmentPlans || [], ['Goal', 'Category', 'Status', 'DueDate', 'Supervisor', 'Notes'])}
    ${profileTable('My Supervisor Check-ins', response.checkins || [], ['CheckinDate', 'Supervisor', 'Summary', 'Concerns', 'DevelopmentGoals', 'FollowUpDate'])}
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
    ...(response.pendingSupervisorRequests || []),
    ...(response.pendingAppeals || []),
  ];
  const counts = response.counts || {};
  document.querySelector('#tasksSummary').innerHTML = [
    stat('Pending LOA', counts.pendingLoa || 0),
    stat('Transfer Requests', counts.pendingTransfers || 0),
    stat('Supervisor Requests', counts.pendingSupervisorRequests || 0),
    stat('Appeals', counts.pendingAppeals || 0),
    stat('Your Supervisees', counts.mySuperviseeTasks || 0),
    stat('Total Tasks', counts.total || 0),
  ].join('');
  renderTable('#tasksTable', state.tasks, ['TaskType', 'Officer', 'Rank', 'Supervisor', 'Subject', 'SourceType', 'StartDate', 'EndDate', 'TargetDivision', 'Reason'], {
    rowAction: (row) => `${row.MySupervisee ? 'class="supervisor-task"' : ''} ${taskOpenAttr(row)}`,
    actions: (row) => `<button class="mini" ${taskOpenAttr(row)}>Review</button>`,
  });
}

function taskOpenAttr(row) {
  if (row.TaskType === 'Transfer Request') return `data-open-transfer-review="${escapeHtml(row.RequestID)}"`;
  if (row.TaskType === 'Supervisor Request') return `data-open-supervisor-review="${escapeHtml(row.RequestID)}"`;
  if (row.TaskType === 'Appeal / Review') return `data-open-appeal-review="${escapeHtml(row.AppealID)}"`;
  return `data-open-loa-review="${escapeHtml(row.RequestID)}"`;
}

async function loadSupervisor() {
  await showViewOnly('supervisor');
  const response = await apiCached('supervisorDashboard', {});
  if (!response.ok) {
    document.querySelector('#supervisorSummary').innerHTML = '';
    return renderTable('#supervisorAssignedTable', [], ['Error'], { emptyMessage: response.error || 'Could not load supervisor dashboard.' });
  }

  state.supervisorDashboard = response;
  const counts = response.counts || {};
  document.querySelector('#supervisorSummary').innerHTML = [
    stat('Assigned Officers', counts.assigned || 0),
    stat('Unassigned Officers', counts.unassigned || 0),
    stat('Pending Requests', counts.pendingRequests || 0),
    stat('Open Goals', counts.openPlans || 0),
  ].join('');

  renderTable('#supervisorAssignedTable', response.assigned || [], ['RobloxUsername', 'Callsign', 'Rank', 'LoaStatus', 'LastShift', 'MonthlyActivity', 'TrainingGaps', 'DisciplineFlags', 'OpenPlans'], {
    rowAction: (row) => `data-open-officer="${escapeHtml(row.OfficerID)}"`,
    actions: (row) => `<button class="mini" data-add-checkin="${escapeHtml(row.OfficerID)}">Check-in</button><button class="mini" data-add-plan="${escapeHtml(row.OfficerID)}">Add goal</button>`,
  });
  renderTable('#supervisorRequestsTable', response.pendingRequests || [], ['Officer', 'Rank', 'Category', 'Subject', 'CreatedAt', 'Supervisor'], {
    actions: (row) => `<button class="mini" data-open-supervisor-review="${escapeHtml(row.RequestID)}">Review</button>`,
  });
  renderTable('#supervisorUnassignedTable', response.unassigned || [], ['RobloxUsername', 'Callsign', 'Rank', 'DutyStatus'], {
    rowAction: (row) => `data-open-officer="${escapeHtml(row.OfficerID)}"`,
    actions: (row) => `<button class="mini" data-assign-supervisor="${escapeHtml(row.OfficerID)}">Assign</button>`,
  });
  renderTable('#supervisorWorkloadTable', response.workload || [], ['Supervisor', 'Rank', 'AssignedOfficers', 'PendingRequests']);
  renderTable('#supervisorPlansTable', response.developmentPlans || [], ['Officer', 'Goal', 'Category', 'Status', 'DueDate', 'Notes'], {
    actions: (row) => `<button class="mini" data-edit-plan="${escapeHtml(row.PlanID)}">Edit</button>`,
  });
  renderTable('#supervisorCheckinsTable', response.checkins || [], ['Officer', 'CheckinDate', 'Summary', 'Concerns', 'DevelopmentGoals', 'FollowUpDate']);
  applyPermissions();
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
    return ['RobloxUsername', 'Callsign', 'Rank', 'Status', 'EffectiveStatus', 'LoaStatus', 'DutyStatus', 'Supervisor', 'Tags'].some((field) => String(officer[field] || '').toLowerCase().includes(query));
  });
  renderTable('#officersTable', rows, ['RobloxUsername', 'Callsign', 'Rank', 'Supervisor', 'EffectiveStatus', 'DutyStatus', 'Tags', 'JoinDate', 'UpdatedAt'], {
    rowAction: (row) => `data-open-officer="${escapeHtml(row.OfficerID)}"`,
    actions: (row) => `<label class="bulk-select"><input type="checkbox" data-bulk-officer="${escapeHtml(row.OfficerID)}"${state.selectedBulkOfficerIds.includes(row.OfficerID) ? ' checked' : ''}> Select</label>`,
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
  const [response, optionsResponse] = await Promise.all([
    apiCached('getOfficerProfile', { OfficerID: officerId }),
    apiCached('listTrainingOptions', {}),
  ]);
  if (optionsResponse.ok) state.trainingOptions = optionsResponse.rows || [];
  if (state.selectedOfficerId !== requestOfficerId) return;
  if (!response.ok) {
    container.innerHTML = emptyState(response.error || 'Officer not found.');
    return;
  }

  renderOfficerProfile(response);
}

async function loadTraining() {
  await showViewOnly('training');
  const [trainingResponse, officersResponse, optionsResponse] = await Promise.all([
    apiCached('listTraining', {}),
    apiCached('listOfficers', {}),
    apiCached('listTrainingOptions', {}),
  ]);
  const trainingRows = trainingResponse.rows || [];
  const officerRows = officersResponse.rows || [];
  state.training = trainingRows;
  state.trainingOptions = optionsResponse.rows || [];
  state.trainingSummary = summarizeTraining(officerRows, trainingRows);
  renderTrainingOverview(state.trainingSummary);
  renderTrainingOptionsPanel();
  renderSearchableView('training');
}

async function loadCourses() {
  await showViewOnly('courses');
  const response = await apiCached('listTrainingCourses', {});
  if (!response.ok) {
    state.courses = [];
    state.courseBookings = [];
    return renderTable('#coursesTable', [], ['Error'], { emptyMessage: response.error || 'Could not load courses.' });
  }
  state.courses = response.rows || [];
  state.courseBookings = response.bookings || [];
  renderSearchableView('courses');
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
    document.querySelector('#documentExplorer').innerHTML = emptyState(response.error || 'Could not load documents.');
    return;
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
  const category = state.documentFolder || document.querySelector('#documentCategoryFilter').value;
  syncDocumentCategoryFilter(category);
  const rows = state.documents.filter((document) => {
    const matchesQuery = ['Title', 'Category', 'RequiredRole', 'RequiredTags', 'Status'].some((field) => String(document[field] || '').toLowerCase().includes(query));
    const matchesCategory = !category || documentFolderName(document) === category;
    return matchesQuery && matchesCategory;
  });
  renderDocumentExplorer(rows, query, category);
}

function renderDocumentExplorer(rows, query, category) {
  const container = document.querySelector('#documentExplorer');
  const folders = documentFolders();
  const folderTiles = !category && !query ? folders.map((folder) => {
    const count = state.documents.filter((document) => documentFolderName(document) === folder).length;
    return `
      <button class="folder-tile" data-doc-folder="${escapeHtml(folder)}">
        <span class="folder-icon"></span>
        <strong>${escapeHtml(folder)}</strong>
        <small>${count} ${count === 1 ? 'file' : 'files'}</small>
      </button>
    `;
  }).join('') : '';

  const files = rows.map((document) => `
    <article class="file-row">
      <a class="file-main" href="${escapeHtml(document.DriveURL || '#')}" target="_blank" rel="noopener">
        <span class="file-icon">${escapeHtml(fileInitial(document))}</span>
        <span>
          <strong>${escapeHtml(document.Title || 'Untitled document')}</strong>
          <small>${escapeHtml(document.Category || 'Unfiled')} / ${escapeHtml(document.RequiredRole || 'Constable+')}</small>
        </span>
      </a>
      <span class="file-meta">${formatCell(document.UpdatedAt || '', 'UpdatedAt')}</span>
      <span class="file-meta">${document.RequiresAcknowledgement === 'TRUE' ? formatCell(document.Acknowledged === 'TRUE' ? 'Acknowledged' : 'Needs acknowledgement', 'Status') : formatCell(document.Status || '', 'Status')}</span>
      <div class="actions">
        ${document.DriveURL ? `<a class="mini" href="${escapeHtml(document.DriveURL)}" target="_blank" rel="noopener">Open</a>` : ''}
        ${document.RequiresAcknowledgement === 'TRUE' && document.Acknowledged !== 'TRUE' ? `<button class="mini" data-ack-document="${escapeHtml(document.DocumentID)}">Acknowledge</button>` : ''}
        ${can('MANAGE_DOCUMENTS') ? `<button class="mini" data-edit-document="${escapeHtml(document.DocumentID)}">Edit</button><button class="mini ghost" data-delete-document="${escapeHtml(document.DocumentID)}">Delete</button>` : ''}
      </div>
    </article>
  `).join('');

  container.innerHTML = `
    <div class="explorer-bar">
      <div class="breadcrumb">
        <button data-doc-folder="">Documents</button>
        ${category ? `<span>/</span><strong>${escapeHtml(category)}</strong>` : ''}
        ${query ? `<span>/</span><strong>Search results</strong>` : ''}
      </div>
      ${category || query ? `<button class="ghost mini" data-doc-folder="">All folders</button>` : ''}
    </div>
    ${folderTiles ? `<section class="folder-grid">${folderTiles}</section>` : ''}
    <section class="file-list">
      <div class="file-list-head">
        <strong>${category ? escapeHtml(category) : query ? 'Matching documents' : 'All documents'}</strong>
        <span>${rows.length} ${rows.length === 1 ? 'file' : 'files'}</span>
      </div>
      ${files || `<p class="empty">No documents found.</p>`}
    </section>
  `;
}

function documentFolders() {
  const folders = state.documents.map(documentFolderName).filter(Boolean);
  return [...new Set(folders)].sort((a, b) => a.localeCompare(b));
}

function documentFolderName(document) {
  return document.Category || 'Unfiled';
}

function fileInitial(document) {
  return String(document.Title || document.Category || 'D').trim().slice(0, 1).toUpperCase();
}

function syncDocumentCategoryFilter(category) {
  const filter = document.querySelector('#documentCategoryFilter');
  const folders = documentFolders();
  filter.innerHTML = [
    '<option value="">All folders</option>',
    ...folders.map((folder) => `<option value="${escapeHtml(folder)}">${escapeHtml(folder)}</option>`),
  ].join('');
  filter.value = category || '';
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

function renderCoursesTable(rows) {
  renderTable('#coursesTable', rows, ['Title', 'Standard', 'Trainer', 'CourseDate', 'Location', 'Capacity', 'BookedSeats', 'Waitlist', 'Status', 'MyBookingStatus'], {
    actions: (row) => [
      can('VIEW_COURSES') && !row.MyBookingStatus ? `<button class="mini" data-request-course="${escapeHtml(row.CourseID)}">Request seat</button>` : '',
      can('MANAGE_COURSES') ? `<button class="mini" data-edit-course="${escapeHtml(row.CourseID)}">Edit</button>` : '',
    ].join(''),
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
  if (view === 'courses') {
    renderCoursesTable(rows);
    renderTable('#courseBookingsTable', state.courseBookings || [], ['Course', 'Officer', 'Rank', 'Status', 'Outcome', 'RequestedAt'], {
      actions: (row) => can('MANAGE_COURSES') ? `<button class="mini" data-review-booking="${escapeHtml(row.BookingID)}">Review</button>` : '',
    });
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
  state.profileSupervisorRequests = data.supervisorRequests || [];
  state.profileAppeals = data.appeals || [];
  state.profileCheckins = data.checkins || [];
  state.profileDevelopmentPlans = data.developmentPlans || [];
  container.innerHTML = `
    <div class="profile-head">
      <button class="ghost" data-view-link="officers">Back</button>
      <div>
        <h2>${escapeHtml(officer.RobloxUsername)}</h2>
        <p>${escapeHtml(officer.Callsign || 'No callsign')} / ${escapeHtml(officer.Rank || 'No rank')}</p>
      </div>
      <div class="profile-actions">
        <button data-edit-officer="${escapeHtml(officer.OfficerID)}" data-permission="EDIT_OFFICERS">Edit officer</button>
        <button data-assign-supervisor="${escapeHtml(officer.OfficerID)}" data-permission="ASSIGN_SUPERVISORS">Supervisor</button>
        <button class="ghost" data-delete-officer="${escapeHtml(officer.OfficerID)}" data-permission="ARCHIVE_OFFICERS">Delete officer</button>
        <button data-add-discipline="${escapeHtml(officer.OfficerID)}" data-permission="ADD_DISCIPLINE">Add discipline</button>
        <button data-add-loa="${escapeHtml(officer.OfficerID)}" data-permission="CREATE_LOA">Add LOA</button>
        <button data-add-checkin="${escapeHtml(officer.OfficerID)}" data-permission="VIEW_TASKS">Check-in</button>
        <button data-add-plan="${escapeHtml(officer.OfficerID)}" data-permission="VIEW_TASKS">Add goal</button>
      </div>
    </div>

    <section class="profile-grid">
      ${detailCard('Status', formatCell(officer.Status), true)}
      ${detailCard('Duty Status', formatCell(officer.DutyStatus || 'Off Duty', 'Status'), true)}
      ${detailCard('Supervisor', officer.Supervisor || 'Not assigned')}
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
          row.Status === 'Denied' ? `<button class="mini" data-request-appeal-source="LOA" data-request-appeal-id="${escapeHtml(row.RequestID)}">Appeal</button>` : '',
          can('APPROVE_LOA') ? `<button class="mini ghost" data-delete-loa="${escapeHtml(row.RequestID)}">Delete</button>` : '',
        ].join(''),
      })}
      ${profileTable('Transfer Requests', data.transfers || [], ['TargetDivision', 'TimeInMO8', 'Reason', 'HasPermission', 'Status', 'ReviewReason'], {
        actions: (row) => row.Status === 'Denied' ? `<button class="mini" data-request-appeal-source="Transfer" data-request-appeal-id="${escapeHtml(row.RequestID)}">Appeal</button>` : '',
      })}
      ${profileTable('Supervisor Requests', data.supervisorRequests || [], ['Category', 'Subject', 'Details', 'Supervisor', 'Status', 'ReviewReason'], {
        actions: (row) => can('VIEW_TASKS') ? `<button class="mini" data-open-supervisor-review="${escapeHtml(row.RequestID)}">Review</button>` : '',
      })}
      ${profileTable('Appeals / Reviews', data.appeals || [], ['SourceType', 'SourceID', 'Reason', 'Status', 'ReviewReason'], {
        actions: (row) => can('VIEW_TASKS') ? `<button class="mini" data-open-appeal-review="${escapeHtml(row.AppealID)}">Review</button>` : '',
      })}
      ${profileTable('Development Plans', data.developmentPlans || [], ['Goal', 'Category', 'Status', 'DueDate', 'Supervisor', 'Notes'], {
        actions: (row) => can('VIEW_TASKS') ? `<button class="mini" data-edit-plan="${escapeHtml(row.PlanID)}">Edit</button>` : '',
      })}
      ${profileTable('Supervisor Check-ins', data.checkins || [], ['CheckinDate', 'Supervisor', 'Summary', 'Concerns', 'DevelopmentGoals', 'FollowUpDate'])}
      ${profileTimeline(data.timeline || [])}
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

async function openSupervisorEditor(officerId) {
  const options = await loadSupervisorOptions();
  const officer = state.officers.find((row) => row.OfficerID === officerId) || {};
  openEditor('Assign supervisor', [
    hiddenField('OfficerID', officerId),
    supervisorSelectField('SupervisorUserID', 'Supervisor', options, officer.SupervisorUserID || ''),
  ], async (values) => api('setOfficerSupervisor', values), {
    successMessage: 'Supervisor assignment saved.',
  });
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

function openTrainingOptionEditor(option = {}) {
  openEditor(option.OptionID ? 'Edit training option' : 'Add training option', [
    hiddenField('OptionID', option.OptionID),
    field('Name', 'Name', 'text', false, option.Name),
    selectField('Type', 'Type', ['Specialist', 'Driving'], option.Type || 'Specialist'),
    selectField('Status', 'Status', ['Active', 'Archived'], option.Status || 'Active'),
    field('SortOrder', 'Sort order', 'number', false, option.SortOrder),
  ], async (values) => api('saveTrainingOption', values), {
    successMessage: 'Training option saved.',
  });
}

async function openCourseEditor(course = {}) {
  const optionsResponse = await apiCached('listTrainingOptions', {});
  const standards = (optionsResponse.rows || []).map((option) => option.Name);
  const trainers = await loadTrainerOptions();
  openEditor(course.CourseID ? 'Edit training course' : 'Create training course', [
    hiddenField('CourseID', course.CourseID),
    field('Title', 'Title', 'text', false, course.Title),
    selectField('Standard', 'Training standard', standards, course.Standard),
    trainerSelectField('TrainerUserID', 'Trainer', trainers, course.TrainerUserID || state.user.UserID),
    field('CourseDate', 'Course date/time', 'datetime-local', false, localDateTimeValue(course.CourseDate)),
    field('Location', 'Location', 'text', false, course.Location),
    field('Capacity', 'Capacity', 'number', false, course.Capacity || '4'),
    selectField('Status', 'Status', ['Scheduled', 'Completed', 'Cancelled', 'Archived'], course.Status || 'Scheduled'),
    field('Notes', 'Notes', 'textarea', true, course.Notes),
  ], async (values) => api('saveTrainingCourse', values), {
    successMessage: 'Training course saved.',
  });
}

function openCourseBookingReviewEditor(record) {
  openEditor('Review course booking', [
    hiddenField('BookingID', record.BookingID),
    field('Course', 'Course', 'text', false, record.Course),
    field('Officer', 'Officer', 'text', false, record.Officer),
    selectField('Status', 'Booking status', ['Approved', 'Waitlist', 'Denied', 'Completed', 'Cancelled'], record.Status || 'Approved'),
    selectField('Outcome', 'Outcome', ['', 'Passed', 'Failed', 'Did Not Attend'], record.Outcome || ''),
    field('Notes', 'Notes', 'textarea', true, record.Notes),
  ], async (values) => api('reviewCourseBooking', values), {
    successMessage: 'Course booking updated.',
  });
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

function openSupervisorRequestEditor() {
  openEditor('Contact supervisor', [
    selectField('Category', 'Request type', ['General', 'Welfare', 'Training', 'Activity', 'Guidance', 'Other'], 'General'),
    field('Subject', 'Subject', 'text', false),
    field('Details', 'Details', 'textarea', true),
  ], async (values) => api('requestSupervisorSupport', values), {
    successMessage: 'Supervisor request submitted.',
  });
}

function openSupervisorReviewEditor(record) {
  openEditor('Review supervisor request', [
    hiddenField('RequestID', record.RequestID),
    field('Officer', 'Officer', 'text', false, record.Officer),
    field('Subject', 'Subject', 'text', false, record.Subject),
    field('Details', 'Request details', 'textarea', true, record.Details),
    selectField('Status', 'Status', ['Completed', 'Pending', 'Denied'], 'Completed'),
    field('ReviewReason', 'Response / notes', 'textarea', true, record.ReviewReason),
  ], async (values) => api('reviewSupervisorRequest', values), {
    successMessage: 'Supervisor request updated.',
  });
}

function openCheckinEditor(officerId) {
  openEditor('Log supervisor check-in', [
    hiddenField('OfficerID', officerId),
    field('CheckinDate', 'Check-in date', 'date', false, dateInputValue(new Date().toISOString())),
    field('Summary', 'Summary', 'textarea', true),
    field('Concerns', 'Concerns', 'textarea', true),
    field('DevelopmentGoals', 'Development goals', 'textarea', true),
    field('FollowUpDate', 'Follow-up date', 'date'),
  ], async (values) => api('saveSupervisorCheckin', values), {
    successMessage: 'Supervisor check-in logged.',
  });
}

function openDevelopmentPlanEditor(officerIdOrRecord) {
  const record = typeof officerIdOrRecord === 'object' ? officerIdOrRecord : {};
  const officerId = record.OfficerID || officerIdOrRecord || '';
  openEditor(record.PlanID ? 'Edit development goal' : 'Add development goal', [
    hiddenField('PlanID', record.PlanID),
    hiddenField('OfficerID', officerId),
    field('Goal', 'Goal', 'textarea', true, record.Goal),
    selectField('Category', 'Category', DEVELOPMENT_CATEGORIES, record.Category || 'Development'),
    selectField('Status', 'Status', DEVELOPMENT_STATUSES, record.Status || 'Open'),
    field('DueDate', 'Due date', 'date', false, dateInputValue(record.DueDate)),
    field('Notes', 'Notes', 'textarea', true, record.Notes),
  ], async (values) => api('saveDevelopmentPlan', values), {
    successMessage: 'Development plan saved.',
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
    selectField('RequiresAcknowledgement', 'Requires acknowledgement', ['FALSE', 'TRUE'], truthy(document.RequiresAcknowledgement) ? 'TRUE' : 'FALSE'),
    selectField('Status', 'Status', ['Published', 'Draft', 'Archived'], document.Status),
  ], async (values) => api('saveDocument', values));
}

async function openBulkOfficerEditor() {
  const options = await loadSupervisorOptions();
  openEditor('Bulk officer actions', [
    field('OfficerIDs', 'Selected officer IDs', 'textarea', true, state.selectedBulkOfficerIds.join(', ')),
    selectField('Status', 'Set status', ['No change', ...OFFICER_STATUSES], 'No change'),
    bulkSupervisorSelectField('SupervisorUserID', 'Set supervisor', options),
    checkboxGroupField('Tags', 'Replace tags', OFFICER_TAGS, ''),
    field('TrainingReviewDate', 'Training review date', 'date'),
  ], async (values) => {
    if (values.Status === 'No change') values.Status = '';
    if (values.SupervisorUserID === '__NO_CHANGE__') delete values.SupervisorUserID;
    return api('bulkUpdateOfficers', values);
  }, {
    successMessage: 'Bulk officer update saved.',
  });
}

function openAppealEditor(sourceType, sourceId) {
  openEditor('Request review / appeal', [
    hiddenField('SourceType', sourceType),
    hiddenField('SourceID', sourceId),
    field('Reason', 'Reason for review', 'textarea', true),
  ], async (values) => api('requestAppeal', values), {
    successMessage: 'Review request submitted.',
  });
}

function openAppealReviewEditor(record) {
  openEditor('Review appeal', [
    hiddenField('AppealID', record.AppealID),
    field('Officer', 'Officer', 'text', false, record.Officer),
    field('SourceType', 'Source type', 'text', false, record.SourceType),
    field('Reason', 'Appeal reason', 'textarea', true, record.Reason),
    selectField('Status', 'Status', ['Completed', 'Approved', 'Denied', 'Pending'], record.Status || 'Completed'),
    field('ReviewReason', 'Response / notes', 'textarea', true, record.ReviewReason),
  ], async (values) => api('reviewAppeal', values), {
    successMessage: 'Appeal review saved.',
  });
}

function openDashboardWidgetEditor() {
  const current = getCachedResponse('dashboard', {})?.widgets || DASHBOARD_WIDGETS.map(([key]) => key);
  openEditor('Dashboard widgets', [
    checkboxGroupField('Widgets', 'Visible widgets', DASHBOARD_WIDGETS.map(([key, label]) => `${key}:${label}`), current.map((key) => `${key}:${DASHBOARD_WIDGETS.find(([item]) => item === key)?.[1] || key}`).join(', ')),
  ], async (values) => {
    values.Widgets = splitTags(values.Widgets).map((item) => item.split(':')[0]).join(', ');
    return api('saveDashboardWidgets', values);
  }, {
    successMessage: 'Dashboard widgets saved.',
  });
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
  if (officerLink && (event.target === officerLink || !event.target.closest('button, a, input, select, textarea'))) {
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

  const assignSupervisor = event.target.closest('[data-assign-supervisor]');
  if (assignSupervisor) {
    await openSupervisorEditor(assignSupervisor.dataset.assignSupervisor);
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

  const requestSupervisor = event.target.closest('[data-request-supervisor]');
  if (requestSupervisor) return openSupervisorRequestEditor();

  const addCheckin = event.target.closest('[data-add-checkin]');
  if (addCheckin) return openCheckinEditor(addCheckin.dataset.addCheckin);

  const addPlan = event.target.closest('[data-add-plan]');
  if (addPlan) return openDevelopmentPlanEditor(addPlan.dataset.addPlan);

  const editPlan = event.target.closest('[data-edit-plan]');
  if (editPlan) {
    const record = (state.supervisorDashboard?.developmentPlans || []).find((row) => row.PlanID === editPlan.dataset.editPlan)
      || state.profileDevelopmentPlans.find((row) => row.PlanID === editPlan.dataset.editPlan);
    if (record) openDevelopmentPlanEditor(record);
    return;
  }

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

  const reviewSupervisorOpen = event.target.closest('[data-open-supervisor-review]');
  if (reviewSupervisorOpen) {
    const record = state.tasks.find((row) => row.RequestID === reviewSupervisorOpen.dataset.openSupervisorReview)
      || state.profileSupervisorRequests.find((row) => row.RequestID === reviewSupervisorOpen.dataset.openSupervisorReview)
      || (state.supervisorDashboard?.pendingRequests || []).find((row) => row.RequestID === reviewSupervisorOpen.dataset.openSupervisorReview);
    if (record) openSupervisorReviewEditor(record);
    return;
  }

  const reviewAppealOpen = event.target.closest('[data-open-appeal-review]');
  if (reviewAppealOpen) {
    const record = state.tasks.find((row) => row.AppealID === reviewAppealOpen.dataset.openAppealReview)
      || state.profileAppeals.find((row) => row.AppealID === reviewAppealOpen.dataset.openAppealReview);
    if (record) openAppealReviewEditor(record);
    return;
  }

  const requestAppeal = event.target.closest('[data-request-appeal-source]');
  if (requestAppeal) {
    openAppealEditor(requestAppeal.dataset.requestAppealSource, requestAppeal.dataset.requestAppealId);
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

  const documentFolder = event.target.closest('[data-doc-folder]');
  if (documentFolder) {
    state.documentFolder = documentFolder.dataset.docFolder || '';
    renderDocumentTable();
    return;
  }

  const acknowledgeDocument = event.target.closest('[data-ack-document]');
  if (acknowledgeDocument) {
    const response = await api('acknowledgeDocument', { DocumentID: acknowledgeDocument.dataset.ackDocument });
    if (!response.ok) {
      showInfo('Acknowledgement failed', `<p>${escapeHtml(response.error || 'Could not acknowledge document.')}</p>`);
      return;
    }
    invalidateCache('listDocuments');
    invalidateCache('dashboard');
    await loadDocuments();
    return;
  }

  const configureDashboard = event.target.closest('[data-configure-dashboard]');
  if (configureDashboard) {
    openDashboardWidgetEditor();
    return;
  }

  const deleteDocument = event.target.closest('[data-delete-document]');
  if (deleteDocument) {
    await confirmDelete('Delete this document link?', 'deleteDocument', { DocumentID: deleteDocument.dataset.deleteDocument }, loadDocuments);
    return;
  }

  const editTrainingOption = event.target.closest('[data-edit-training-option]');
  if (editTrainingOption) {
    const option = state.trainingOptions.find((row) => row.OptionID === editTrainingOption.dataset.editTrainingOption);
    if (option) openTrainingOptionEditor(option);
    return;
  }

  const editCourse = event.target.closest('[data-edit-course]');
  if (editCourse) {
    const course = state.courses.find((row) => row.CourseID === editCourse.dataset.editCourse);
    if (course) openCourseEditor(course);
    return;
  }

  const requestCourse = event.target.closest('[data-request-course]');
  if (requestCourse) {
    const response = await api('requestCourseSeat', { CourseID: requestCourse.dataset.requestCourse });
    if (!response.ok) {
      showInfo('Course request failed', `<p>${escapeHtml(response.error || 'Could not request this course.')}</p>`);
      return;
    }
    invalidateCache('listTrainingCourses');
    await loadCourses();
    return;
  }

  const reviewBooking = event.target.closest('[data-review-booking]');
  if (reviewBooking) {
    const booking = state.courseBookings.find((row) => row.BookingID === reviewBooking.dataset.reviewBooking);
    if (booking) openCourseBookingReviewEditor(booking);
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

function handleBulkOfficerSelection(event) {
  const checkbox = event.target.closest('[data-bulk-officer]');
  if (!checkbox) return;
  const officerId = checkbox.dataset.bulkOfficer;
  state.selectedBulkOfficerIds = checkbox.checked
    ? [...new Set([...state.selectedBulkOfficerIds, officerId])]
    : state.selectedBulkOfficerIds.filter((id) => id !== officerId);
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

function profileTimeline(rows) {
  const body = rows.length
    ? rows.map((row) => `
      <article class="timeline-item">
        <span>${formatCell(row.Date, isDateTimeColumn('CreatedAt') ? 'CreatedAt' : 'Date')}</span>
        <strong>${escapeHtml(row.Type || 'Update')} / ${escapeHtml(row.Title || '')}</strong>
        <p>${escapeHtml(row.Detail || '')}</p>
      </article>
    `).join('')
    : `<p class="empty">No timeline events yet.</p>`;
  return `
    <section class="profile-panel">
      <h3>Status Timeline</h3>
      <div class="timeline-list">${body}</div>
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
  const specialistOptions = trainingOptionNames('Specialist');
  const drivingStandards = trainingOptionNames('Driving');
  const rows = specialistOptions.map((standard) => {
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
  const drivingRecord = drivingStandards.find((standard) => {
    return trainingRows.some((item) => item.Standard === standard && String(item.Status) === 'Passed');
  }) || '';
  const drivingOptions = [''].concat(drivingStandards).map((standard) => {
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

function trainingOptionNames(type) {
  const options = state.trainingOptions.length ? state.trainingOptions : [
    ...SPECIALIST_TRAINING.map((name, index) => ({ Name: name, Type: 'Specialist', SortOrder: index + 1 })),
    ...DRIVING_STANDARDS.map((name, index) => ({ Name: name, Type: 'Driving', SortOrder: index + 1 })),
  ];
  return options
    .filter((option) => option.Type === type && option.Status !== 'Archived')
    .sort((a, b) => Number(a.SortOrder || 0) - Number(b.SortOrder || 0) || String(a.Name).localeCompare(String(b.Name)))
    .map((option) => option.Name);
}

function summarizeTraining(officers, trainingRows) {
  const specialistOptions = trainingOptionNames('Specialist');
  const drivingStandards = trainingOptionNames('Driving');
  return officers.map((officer) => {
    const records = trainingRows.filter((item) => item.OfficerID === officer.OfficerID);
    const passed = (standard) => records.some((item) => item.Standard === standard && item.Status === 'Passed');
    const specialistTickets = specialistOptions.filter(passed);
    const drivingStandard = drivingStandards.find(passed) || 'Not set';
    const missing = [
      ...specialistOptions.filter((standard) => !passed(standard)),
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

function renderTrainingOptionsPanel() {
  const container = document.querySelector('#trainingOptionsPanel');
  const rows = state.trainingOptions || [];
  if (!can('MANAGE_TRAINING_OPTIONS')) {
    container.innerHTML = '';
    return;
  }
  container.innerHTML = `
    <h3>Training Options</h3>
    <div class="table-wrap compact">
      <table id="trainingOptionsTable"></table>
    </div>
  `;
  renderTable('#trainingOptionsTable', rows, ['Name', 'Type', 'Status', 'SortOrder', 'UpdatedAt'], {
    actions: (row) => `<button class="mini" data-edit-training-option="${escapeHtml(row.OptionID)}">Edit</button>`,
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
  if (['Active', 'Published', 'Passed', 'Approved', 'Acknowledged', 'Completed'].includes(text)) {
    return `<span class="pill success">${escapeHtml(text)}</span>`;
  }
  if (['On Duty'].includes(text)) {
    return `<span class="pill success">${escapeHtml(text)}</span>`;
  }
  if (['LOA', 'On LOA', 'Pending', 'In Progress', 'Draft', 'Not Started', 'Needs acknowledgement'].includes(text)) {
    return `<span class="pill warning">${escapeHtml(text)}</span>`;
  }
  if (['Off Duty', 'Suspended', 'Archived', 'Failed', 'Denied', 'Expired', 'Removed', 'Low activity', 'No activity'].includes(text)) {
    return `<span class="pill danger">${escapeHtml(text)}</span>`;
  }
  return escapeHtml(text);
}

function isDateColumn(column) {
  return ['StartDate', 'EndDate', 'JoinDate', 'DateCompleted', 'ExpiryDate', 'ReviewDate', 'ExpiresAt', 'CurrentLoaEnd', 'CheckinDate', 'FollowUpDate', 'DueDate'].includes(column);
}

function isDateTimeColumn(column) {
  return ['UpdatedAt', 'CreatedAt', 'IssuedAt', 'ReviewedAt', 'ReadAt', 'Timestamp', 'LastLogin', 'ChangedAt', 'StartedAt', 'EndedAt', 'LastShift', 'CourseDate', 'RequestedAt'].includes(column);
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

function supervisorSelectField(name, label, options, selected = '') {
  const optionHtml = [
    `<option value="">No supervisor assigned</option>`,
    ...options.map((option) => {
      const isSelected = option.UserID === selected ? ' selected' : '';
      return `<option value="${escapeHtml(option.UserID)}"${isSelected}>${escapeHtml(option.RobloxUsername)} - ${escapeHtml(option.Rank || option.Role || '')}</option>`;
    }),
  ].join('');
  return { html: `<label>${escapeHtml(label)}<select name="${escapeHtml(name)}">${optionHtml}</select></label>` };
}

function bulkSupervisorSelectField(name, label, options) {
  const optionHtml = [
    `<option value="__NO_CHANGE__">No change</option>`,
    `<option value="">No supervisor assigned</option>`,
    ...options.map((option) => `<option value="${escapeHtml(option.UserID)}">${escapeHtml(option.RobloxUsername)} - ${escapeHtml(option.Rank || option.Role || '')}</option>`),
  ].join('');
  return { html: `<label>${escapeHtml(label)}<select name="${escapeHtml(name)}">${optionHtml}</select></label>` };
}

async function loadSupervisorOptions() {
  if (state.supervisorOptions.length) return state.supervisorOptions;
  const response = await apiCached('supervisorOptions', {});
  state.supervisorOptions = response.ok ? response.rows || [] : [];
  return state.supervisorOptions;
}

async function loadTrainerOptions() {
  const response = await apiCached('courseTrainers', {});
  return response.ok ? response.rows || [] : [state.user].filter(Boolean);
}

function trainerSelectField(name, label, options, selected = '') {
  const optionHtml = options.map((option) => {
    const isSelected = option.UserID === selected ? ' selected' : '';
    return `<option value="${escapeHtml(option.UserID)}"${isSelected}>${escapeHtml(option.RobloxUsername)} - ${escapeHtml(option.Rank || option.Role || '')}</option>`;
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

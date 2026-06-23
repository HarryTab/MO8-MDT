const API_URL = 'https://script.google.com/macros/s/AKfycbwsRocB7bsQLfXiazKGI-O158ppsRnQPVsrtvzVaoyUUgMdanidkOJc_pg--lddbDGPhQ/exec';
const APP_VERSION = '2026-06-23-3';
const SUPABASE_CONFIG = window.MO8_SUPABASE || {};
const USE_SUPABASE = Boolean(SUPABASE_CONFIG.url && SUPABASE_CONFIG.anonKey && window.supabase);
const supabaseClient = USE_SUPABASE ? window.supabase.createClient(SUPABASE_CONFIG.url, SUPABASE_CONFIG.anonKey) : null;

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
const ALL_PERMISSIONS = [
  'VIEW_DASHBOARD',
  'VIEW_TASKS',
  'VIEW_OFFICERS',
  'VIEW_RANK_LOG',
  'ADD_OFFICERS',
  'EDIT_OFFICERS',
  'ARCHIVE_OFFICERS',
  'ASSIGN_SUPERVISORS',
  'VIEW_TRAINING',
  'MANAGE_TRAINING',
  'VIEW_COURSES',
  'MANAGE_COURSES',
  'MANAGE_TRAINING_OPTIONS',
  'VIEW_DISCIPLINE',
  'ADD_DISCIPLINE',
  'VIEW_LOA',
  'CREATE_LOA',
  'APPROVE_LOA',
  'VIEW_DOCUMENTS',
  'MANAGE_DOCUMENTS',
  'VIEW_ANNOUNCEMENTS',
  'MANAGE_ANNOUNCEMENTS',
  'MANAGE_USERS',
  'REVIEW_ACCOUNT_REQUESTS',
  'RESET_PASSWORDS',
  'VIEW_AUDIT_LOG',
  'MANAGE_PERMISSIONS',
  'CHANGE_OWN_PASSWORD',
  'FULL_ACCESS',
];
const CACHE_TTL_MS = 5 * 60 * 1000;
const CACHE_STORAGE_KEY = 'mo8_api_cache';
const BOOT_STORAGE_KEY = 'mo8_boot_ready';
const SESSION_STORAGE_KEY = 'mo8_session_auth';
const VERSION_STORAGE_KEY = 'mo8_app_version';
const DASHBOARD_LAYOUT_STORAGE_KEY = 'mo8_dashboard_layout';
const USER_PERMISSION_MODES = ['Inherit', 'Allow', 'Deny'];
const ANNOUNCEMENT_STATUSES = ['Published', 'Draft', 'Archived'];
const DEVELOPMENT_CATEGORIES = ['Development', 'Training', 'Activity', 'Conduct', 'Career', 'Other'];
const DEVELOPMENT_STATUSES = ['Open', 'In Progress', 'Completed', 'Paused'];
const DASHBOARD_WIDGETS = [
  ['myActions', 'My Actions'],
  ['activeLoa', 'Active LOA Status'],
  ['pendingLoa', 'Pending LOA'],
  ['announcements', 'Notice Board'],
  ['upcomingTraining', 'Upcoming Training'],
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
  selectedCourseId: '',
  dashboardInteraction: null,
  calendarInstance: null,
  operations: { actions: [], search: [], savedViews: [], calendar: [], probation: [], reviews: [], restrictions: [], handovers: [] },
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
  mobileMenuButton: document.querySelector('#mobileMenuButton'),
  mobileMenuLabel: document.querySelector('#mobileMenuLabel'),
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
document.addEventListener('input', handleSearchableOfficerInput);
document.addEventListener('pointerdown', handleDashboardPointerDown);
document.addEventListener('pointermove', handleDashboardPointerMove);
document.addEventListener('pointerup', handleDashboardPointerUp);
document.addEventListener('pointercancel', handleDashboardPointerUp);

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
  localStorage.setItem('mo8_auth_provider', response.provider || (USE_SUPABASE ? 'supabase' : 'apps-script'));
  localStorage.setItem('mo8_token', state.token);
  storeSessionAuth(state.user, state.permissions);
  await initializeSession();
});

document.querySelector('#requestAccountButton')?.addEventListener('click', () => {
  const form = document.querySelector('#accountRequestForm');
  form.reset();
  document.querySelector('#accountRequestRank').innerHTML = OFFICER_RANKS.map((rank) => `<option>${escapeHtml(rank)}</option>`).join('');
  document.querySelector('#accountRequestStatus').textContent = '';
  form.querySelector('button[value="submit"]').disabled = false;
  document.querySelector('#accountRequestDialog').showModal();
});
document.querySelector('#accountRequestForm')?.addEventListener('submit', async (event) => {
  event.preventDefault();
  if (event.submitter?.value === 'cancel') {
    document.querySelector('#accountRequestDialog').close();
    return;
  }
  const status = document.querySelector('#accountRequestStatus');
  status.textContent = 'Sending request...';
  const values = Object.fromEntries(new FormData(event.currentTarget));
  const response = await api('requestAccount', values, false);
  if (!response.ok) {
    status.textContent = response.error || 'Could not send account request.';
    return;
  }
  status.textContent = 'Request sent. An Inspector+ will review your details.';
  event.currentTarget.querySelector('button[type="submit"]').disabled = true;
  setTimeout(() => document.querySelector('#accountRequestDialog').close(), 1400);
});

elements.logoutButton.addEventListener('click', async () => {
  await api('logout', {});
  localStorage.removeItem('mo8_token');
  localStorage.removeItem('mo8_auth_provider');
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
elements.mobileMenuButton?.addEventListener('click', () => toggleMobileNav());

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
    closeMobileNav();
  });
});

document.querySelector('#newOfficerButton').addEventListener('click', () => openOfficerEditor());
document.querySelector('#bulkOfficerButton').addEventListener('click', () => openBulkOfficerEditor());
document.querySelector('#newDocumentButton').addEventListener('click', () => openDocumentEditor());
document.querySelector('#newFolderButton').addEventListener('click', () => openFolderEditor());
document.querySelector('#newTrainingOptionButton').addEventListener('click', () => openTrainingOptionEditor());
document.querySelector('#newCourseButton').addEventListener('click', () => openCourseEditor());
document.querySelector('#newAnnouncementButton').addEventListener('click', () => openAnnouncementEditor());
document.querySelector('#newUserButton').addEventListener('click', () => openUserEditor());
document.querySelector('#startShiftButton').addEventListener('click', startShift);
document.querySelector('#endShiftButton').addEventListener('click', openEndShiftEditor);
document.querySelector('#requestRetrospectiveShiftButton')?.addEventListener('click', openRetrospectiveShiftEditor);
document.querySelector('#shiftPeriodFilter').addEventListener('change', loadShift);
document.querySelector('#shiftStartFilter').addEventListener('change', loadShift);
document.querySelector('#shiftEndFilter').addEventListener('change', loadShift);
document.querySelector('#globalSearchInput')?.addEventListener('input', debounce(runGlobalSearch, 220));
document.querySelector('#savedViewSelect')?.addEventListener('change', applySavedView);
document.querySelector('#saveSearchViewButton')?.addEventListener('click', openSavedViewEditor);
document.querySelector('#calendarDatePicker')?.addEventListener('change', (event) => state.calendarInstance?.gotoDate(event.target.value));
document.querySelector('#calendarTypeFilter')?.addEventListener('change', renderCalendar);
document.querySelector('#calendarViewMode')?.addEventListener('change', (event) => state.calendarInstance?.changeView(event.target.value));
document.querySelector('#calendarPreviousButton')?.addEventListener('click', () => state.calendarInstance?.prev());
document.querySelector('#calendarNextButton')?.addEventListener('click', () => state.calendarInstance?.next());
document.querySelector('#calendarTodayButton')?.addEventListener('click', () => state.calendarInstance?.today());
document.querySelector('#newCalendarEventButton')?.addEventListener('click', () => openCalendarEventEditor());
document.querySelector('#developmentSearch')?.addEventListener('input', renderDevelopmentTables);
document.querySelector('#newProbationButton')?.addEventListener('click', () => openProbationEditor());
document.querySelector('#newReviewButton')?.addEventListener('click', () => openPerformanceReviewEditor());
document.querySelector('#newRestrictionButton')?.addEventListener('click', () => openRestrictionEditor());
document.querySelector('#newHandoverButton')?.addEventListener('click', () => openHandoverEditor());
document.querySelector('#handoverSearch')?.addEventListener('input', renderHandoverBoard);
document.querySelector('#handoverStatusFilter')?.addEventListener('change', renderHandoverBoard);

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
    { label: 'Checking personal actions', run: () => apiCached('myActions', {}) },
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
    ['operationalCalendar', {}],
    ['savedViews', {}],
    can('VIEW_TASKS') ? ['supervisorDashboard', {}] : null,
    can('VIEW_TASKS') ? ['developmentRecords', {}] : null,
    can('VIEW_TASKS') ? ['listHandovers', {}] : null,
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
  document.body.classList.remove('is-officer-portal');
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

function toggleMobileNav(force) {
  const shouldOpen = typeof force === 'boolean' ? force : !document.body.classList.contains('mobile-nav-open');
  document.body.classList.toggle('mobile-nav-open', shouldOpen);
  elements.mobileMenuButton?.setAttribute('aria-expanded', String(shouldOpen));
}

function closeMobileNav() {
  toggleMobileNav(false);
}

function defaultView() {
  if (can('VIEW_DASHBOARD')) return 'dashboard';
  return 'myProfile';
}

async function showView(view) {
  const titles = {
    dashboard: ['Dashboard', 'Current MO8 overview'],
    globalSearch: ['Search', 'Search across the MDT and open saved views'],
    calendar: ['Calendar', 'LOA, courses, reviews and operational events'],
    myProfile: ['My Profile', 'Your officer record, training, LOA and notifications'],
    shift: ['Shift Log', 'Duty status and team activity'],
    tasks: ['Tasks', 'Outstanding approvals and command actions'],
    supervisor: ['Supervisor', 'Assigned officers, check-ins, development plans and workload'],
    officers: ['Officers', 'MO8 officer database'],
    officerProfile: ['Officer Profile', 'Individual record and linked history'],
    development: ['Officer Development', 'Probation, performance and temporary restrictions'],
    rankChanges: ['Rank Change Log', 'Promotion and rank movement history'],
    training: ['Training', 'Training standards and status'],
    courses: ['Training Courses', 'Course bookings, waitlists and trainer outcomes'],
    discipline: ['Discipline', 'Internal roleplay administration records'],
    loa: ['Leave of Absence', 'Requests and reviews'],
    documents: ['Documents', 'Training guides and policy links'],
    announcements: ['Notice Board', 'Operational updates and command notices'],
    users: ['Users', 'Sergeant+ login accounts'],
    permissions: ['Permissions', 'Role defaults and individual overrides'],
    reports: ['Reports', 'Command performance and compliance reporting'],
    handover: ['Command Handover', 'Outstanding operational matters and ownership'],
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
  if (elements.mobileMenuLabel) elements.mobileMenuLabel.textContent = titles[view][0];
  renderViewLoading(view);

  const loaders = {
    dashboard: loadDashboard,
    globalSearch: loadGlobalSearch,
    calendar: loadCalendar,
    myProfile: loadMyProfile,
    shift: loadShift,
    tasks: loadTasks,
    supervisor: loadSupervisor,
    officers: loadOfficers,
    officerProfile: () => loadOfficerProfile(state.selectedOfficerId),
    development: loadDevelopment,
    rankChanges: loadRankChanges,
    training: loadTraining,
    courses: loadCourses,
    discipline: loadDiscipline,
    loa: loadLoa,
    documents: loadDocuments,
    announcements: loadAnnouncements,
    users: loadUsers,
    permissions: loadPermissions,
    reports: loadReports,
    handover: loadHandover,
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
  if (view === 'globalSearch') {
    document.querySelector('#globalSearchResults').innerHTML = loadingBlock('Preparing global search...');
    return;
  }
  if (view === 'calendar') {
    document.querySelector('#calendarGrid').innerHTML = loadingBlock('Loading operational calendar...');
    return;
  }
  if (view === 'development') {
    document.querySelector('#developmentSummary').innerHTML = '';
    ['#probationTable', '#performanceReviewsTable', '#restrictionsTable'].forEach((selector) => {
      document.querySelector(selector).innerHTML = `<tbody><tr><td>${loadingBlock('Loading development records...')}</td></tr></tbody>`;
    });
    return;
  }
  if (view === 'handover') {
    document.querySelector('#handoverBoard').innerHTML = loadingBlock('Loading command handover...');
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
    globalSearch: 'savedViews',
    calendar: 'operationalCalendar',
    myProfile: 'myProfile',
    shift: 'teamShifts',
    tasks: 'tasks',
    supervisor: 'supervisorDashboard',
    officers: 'listOfficers',
    development: 'developmentRecords',
    rankChanges: 'rankChangeLog',
    training: 'listTraining',
    courses: 'listTrainingCourses',
    discipline: 'listDiscipline',
    loa: 'listLoa',
    documents: 'listDocuments',
    announcements: 'listAnnouncements',
    users: 'listUsers',
    permissions: 'permissionsConfig',
    reports: 'commandReports',
    handover: 'listHandovers',
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
  const layout = getDashboardLayout(activeWidgets);
  const widgetBodies = {
    myActions: ['My Actions', (response.myActions || []).length ? (response.myActions || []).slice(0, 6).map(actionCard).join('') : '<p class="empty">No outstanding actions.</p>'],
    activeLoa: ['Active LOA Status', dashboardRows(response.activeLoa || [], ['Officer', 'Rank', 'EndDate', 'Status'])],
    pendingLoa: ['Pending LOA', dashboardRows(response.pendingLoa || [], ['Officer', 'Rank', 'StartDate', 'EndDate'])],
    announcements: ['Notice Board', announcementRows(response.announcements || [])],
    upcomingTraining: ['Upcoming Training', dashboardRows(response.upcomingTraining || [], ['Title', 'Standard', 'CourseDate', 'Status', 'Location'])],
    trainingReviews: ['Training Reviews', dashboardRows(response.trainingReviewsDue || [], ['RobloxUsername', 'Standard', 'ReviewDate', 'UpdatedBy'])],
    recentDocuments: ['Recent Documents', dashboardRows(response.recentDocuments || [], ['Title', 'Category', 'RequiredRole', 'UpdatedAt'])],
    recentActivity: ['Recent Activity', dashboardRows(response.recentAudit || [], ['Timestamp', 'Action', 'TargetType', 'TargetID'])],
    pendingAppeals: ['Pending Appeals', dashboardRows(response.pendingAppeals || [], ['Officer', 'Rank', 'SourceType', 'Reason'])],
    unassignedOfficers: ['Unassigned Officers', dashboardRows(response.unassignedOfficers || [], ['RobloxUsername', 'Rank', 'DutyStatus'])],
    lowActivity: ['Low Activity', dashboardRows(response.lowActivity || [], ['RobloxUsername', 'Rank', 'Duration', 'ActivityFlag'])],
    documentAcknowledgements: ['Document Acknowledgements', dashboardRows(response.documentAcknowledgements || [], ['Title', 'Category', 'RequiredRole'])],
  };
  const renderedWidgets = layout.order
    .filter((key) => activeWidgets.includes(key) && widgetBodies[key])
    .map((key, index) => dashboardWidget(key, widgetBodies[key][0], widgetBodies[key][1], layout, index))
    .join('');
  elements.dashboardView.innerHTML = `
    <div class="section-head dashboard-config">
      <h2>Dashboard</h2>
      <div class="dashboard-config-actions">
        <button class="ghost" data-reset-dashboard-layout>Reset layout</button>
        <button class="ghost" data-configure-dashboard>Widgets</button>
      </div>
    </div>
    <div class="stat-row">
      ${[
    stat('Active Officers', counts.activeOfficers || 0),
    stat('Currently On LOA', counts.currentlyOnLoa || 0),
    stat('Pending LOA', counts.loaPending || 0),
    stat('Pending Appeals', counts.pendingAppeals || 0),
    stat('Review Due', counts.trainingReviewsDue || 0),
    stat('Docs To Ack', counts.pendingAcknowledgements || 0),
    stat('Upcoming Training', counts.upcomingTraining || 0),
  ].join('')}
    </div>
    <section class="dashboard-grid dashboard-widget-grid">
      ${renderedWidgets || emptyState('No dashboard widgets selected.')}
    </section>
  `;
}

function actionCard(row) {
  return `<article class="action-card priority-${escapeHtml(String(row.Priority || 'normal').toLowerCase())}" ${row.View ? `data-view-link="${escapeHtml(row.View)}"` : ''}>
    <span>${escapeHtml(row.Type || 'Action')}</span><strong>${escapeHtml(row.Title || '')}</strong>
    <p>${escapeHtml(row.Detail || '')}</p><small>${row.DueDate ? formatDisplayDate(row.DueDate) : ''}</small>
  </article>`;
}

async function loadGlobalSearch() {
  await showViewOnly('globalSearch');
  const response = await apiCached('savedViews', {});
  state.operations.savedViews = response.rows || [];
  renderSavedViewOptions();
  await runGlobalSearch();
}

async function runGlobalSearch() {
  const query = document.querySelector('#globalSearchInput')?.value.trim() || '';
  const container = document.querySelector('#globalSearchResults');
  if (!container) return;
  container.innerHTML = loadingBlock('Searching the MDT...');
  const response = await api('globalSearch', { Query: query });
  if (!response.ok) return renderError(container, response.error);
  state.operations.search = response.rows || [];
  const grouped = Object.groupBy ? Object.groupBy(state.operations.search, (row) => row.Type) : state.operations.search.reduce((result, row) => {
    (result[row.Type] ||= []).push(row); return result;
  }, {});
  container.innerHTML = state.operations.search.length ? Object.entries(grouped).map(([type, rows]) => `
    <section class="search-result-group"><h3>${escapeHtml(type)}</h3><div class="search-result-list">${rows.map((row) => `
      <article class="search-result" ${row.View ? `data-view-link="${escapeHtml(row.View)}"` : ''}${row.OfficerID ? ` data-open-officer="${escapeHtml(row.OfficerID)}"` : ''}>
        <strong>${escapeHtml(row.Title)}</strong><p>${escapeHtml(row.Detail || '')}</p><span>${escapeHtml(row.Meta || '')}</span>
      </article>`).join('')}</div></section>`).join('') : emptyState(query ? 'No matching records found.' : 'Enter a search term or select a saved view.');
}

function renderSavedViewOptions() {
  const select = document.querySelector('#savedViewSelect');
  if (!select) return;
  select.innerHTML = '<option value="">Saved views</option>' + state.operations.savedViews.map((row) => `<option value="${escapeHtml(row.ViewID)}">${escapeHtml(row.Name)}</option>`).join('');
}

function applySavedView(event) {
  const saved = state.operations.savedViews.find((row) => row.ViewID === event.target.value);
  if (!saved) return;
  document.querySelector('#globalSearchInput').value = saved.Query || '';
  runGlobalSearch();
}

function openSavedViewEditor() {
  openEditor('Save search view', [field('Name', 'View name'), field('Query', 'Search query', 'text', false, document.querySelector('#globalSearchInput').value)], async (values) => api('saveSavedView', values), {
    successMessage: 'Saved view created.', onSuccess: async () => { invalidateCache('savedViews'); await loadGlobalSearch(); },
  });
}

async function loadCalendar() {
  await showViewOnly('calendar');
  const response = await apiCached('operationalCalendar', {});
  if (!response.ok) return renderError(document.querySelector('#calendarGrid'), response.error);
  state.operations.calendar = response.rows || [];
  const types = [...new Set(state.operations.calendar.map((row) => row.Type).filter(Boolean))];
  document.querySelector('#calendarTypeFilter').innerHTML = '<option value="">All event types</option>' + types.map((type) => `<option>${escapeHtml(type)}</option>`).join('');
  renderCalendar();
}

function renderCalendar() {
  const container = document.querySelector('#calendarGrid');
  if (!container) return;
  if (!window.FullCalendar) return renderCalendarFallback(container);
  const events = calendarEvents();
  if (!state.calendarInstance) {
    state.calendarInstance = new FullCalendar.Calendar(container, {
      initialView: document.querySelector('#calendarViewMode')?.value || 'dayGridMonth',
      headerToolbar: { start: 'title', center: '', end: '' },
      firstDay: 1,
      allDaySlot: true,
      allDayText: 'All day',
      dayMaxEvents: true,
      nowIndicator: true,
      slotMinTime: '06:00:00',
      slotMaxTime: '24:00:00',
      slotDuration: '00:30:00',
      height: 'auto',
      events,
      datesSet(info) {
        const picker = document.querySelector('#calendarDatePicker');
        if (picker) picker.value = calendarDateKey(info.view.currentStart);
        const mode = document.querySelector('#calendarViewMode');
        if (mode) mode.value = info.view.type;
      },
      dateClick(info) { showCalendarDay(info.dateStr.slice(0, 10)); },
      eventClick(info) { showCalendarEvent(info.event.id); },
    });
    state.calendarInstance.render();
  } else {
    state.calendarInstance.removeAllEvents();
    state.calendarInstance.addEventSource(events);
    state.calendarInstance.updateSize();
  }
}

function calendarEvents() {
  const type = document.querySelector('#calendarTypeFilter')?.value || '';
  return state.operations.calendar.filter((row) => !type || row.Type === type).map((row) => {
    const allDay = calendarIsAllDay(row);
    const typeClass = String(row.Type || 'event').toLowerCase().replace(/[^a-z0-9]+/g, '-');
    return {
      id: row.ID,
      title: row.Title,
      start: row.Start,
      end: allDay && row.End ? calendarDateKey(addCalendarDays(calendarDate(row.End), 1)) : row.End || undefined,
      allDay,
      classNames: [`calendar-event-${typeClass}`],
      extendedProps: { type: row.Type, detail: row.Detail, location: row.Location },
    };
  });
}

function calendarIsAllDay(row) {
  if (row.Type === 'LOA') return true;
  return !String(row.Start || '').includes('T') && !String(row.Start || '').match(/\d{2}:\d{2}/);
}

function renderCalendarFallback(container) {
  container.innerHTML = emptyState('The calendar component could not load. Refresh the page to try again.');
}

function localMonthValue(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
}

function calendarDate(value) {
  if (value instanceof Date) return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  const match = String(value || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  const date = new Date(value);
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function addCalendarDays(date, days) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate() + days);
}

function calendarDateKey(value) {
  const date = calendarDate(value);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

function calendarDayNumber(value) {
  const date = calendarDate(value);
  return Math.floor(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()) / 86400000);
}

function calendarEventIncludes(row, date) {
  const day = calendarDate(date);
  return calendarDate(row.Start) <= day && calendarDate(row.End || row.Start) >= day;
}

function showCalendarDay(dateKey) {
  const rows = state.operations.calendar.filter((row) => calendarEventIncludes(row, dateKey));
  showInfo(formatDisplayDate(dateKey), rows.length ? `<div class="calendar-day-details">${rows.map((row) => `
    <article><span>${escapeHtml(row.Type)}</span><strong>${escapeHtml(row.Title)}</strong><p>${escapeHtml(row.Detail || '')}</p><small>${escapeHtml(row.Location || '')}${row.End && calendarDateKey(row.End) !== calendarDateKey(row.Start) ? ` / Until ${escapeHtml(formatDisplayDate(row.End))}` : ''}</small>${row.Editable ? `<button class="mini" data-edit-calendar="${escapeHtml(row.ID)}">Edit</button>` : ''}</article>`).join('')}</div>` : '<p class="empty">No entries on this day.</p>');
}

function showCalendarEvent(eventId) {
  const row = state.operations.calendar.find((item) => item.ID === eventId);
  if (!row) return;
  showInfo(row.Title, `<div class="calendar-day-details"><article><span>${escapeHtml(row.Type)}</span><strong>${escapeHtml(formatDisplayDate(row.Start))}${row.End ? ` to ${escapeHtml(formatDisplayDate(row.End))}` : ''}</strong><p>${escapeHtml(row.Detail || 'No additional details.')}</p><small>${escapeHtml(row.Location || '')}</small>${row.Editable ? `<button class="mini" data-edit-calendar="${escapeHtml(row.ID)}">Edit</button>` : ''}</article></div>`);
}

function openCalendarEventEditor(record = {}) {
  openEditor(record.ID ? 'Edit calendar event' : 'Add calendar event', [
    hiddenField('EventID', record.ID || ''), field('Title', 'Title', 'text', false, record.Title || ''),
    selectField('EventType', 'Type', ['Operational', 'Meeting', 'Deadline', 'Training', 'Other'], record.Type || 'Operational'),
    field('StartsAt', 'Starts at', 'datetime-local', false, localDateTimeValue(record.Start)), field('EndsAt', 'Ends at', 'datetime-local', false, localDateTimeValue(record.End)),
    field('Location', 'Location', 'text', false, record.Location || ''), field('Details', 'Details', 'textarea', true, record.Detail || ''),
  ], async (values) => api('saveCalendarEvent', values), { successMessage: 'Calendar event saved.', onSuccess: async () => { invalidateCache('operationalCalendar'); await loadCalendar(); } });
}

async function loadDevelopment() {
  await showViewOnly('development');
  const response = await apiCached('developmentRecords', {});
  if (!response.ok) return renderError(document.querySelector('#developmentSummary'), response.error);
  state.officers = response.officers || state.officers;
  state.operations.probation = response.probation || [];
  state.operations.reviews = response.reviews || [];
  state.operations.restrictions = response.restrictions || [];
  document.querySelector('#developmentSummary').innerHTML = [
    stat('Active Probation', state.operations.probation.filter((row) => row.Status === 'Active').length),
    stat('Reviews Due', state.operations.reviews.filter((row) => row.NextReviewDate && new Date(row.NextReviewDate) <= new Date()).length),
    stat('Active Restrictions', state.operations.restrictions.filter((row) => row.Status === 'Active').length),
  ].join('');
  renderDevelopmentTables();
}

function renderDevelopmentTables() {
  const query = document.querySelector('#developmentSearch')?.value.toLowerCase() || '';
  const filter = (rows) => rows.filter((row) => Object.values(row).some((value) => String(value || '').toLowerCase().includes(query)));
  renderTable('#probationTable', filter(state.operations.probation), ['Officer', 'Rank', 'Stage', 'Status', 'Progress', 'TargetDate', 'Reviewer'], { actions: (row) => `<button class="mini" data-edit-probation="${escapeHtml(row.ProbationID)}">Edit</button><button class="mini ghost" data-delete-probation="${escapeHtml(row.ProbationID)}">Delete</button>` });
  renderTable('#performanceReviewsTable', filter(state.operations.reviews), ['Officer', 'ReviewDate', 'Rating', 'Reviewer', 'NextReviewDate', 'Objectives'], { actions: (row) => `<button class="mini" data-edit-performance-review="${escapeHtml(row.ReviewID)}">Edit</button><button class="mini ghost" data-delete-performance-review="${escapeHtml(row.ReviewID)}">Delete</button>` });
  renderTable('#restrictionsTable', filter(state.operations.restrictions), ['Officer', 'RestrictionType', 'Details', 'StartsOn', 'EndsOn', 'Status'], { actions: (row) => `<button class="mini" data-edit-restriction="${escapeHtml(row.RestrictionID)}">Edit</button><button class="mini ghost" data-delete-restriction="${escapeHtml(row.RestrictionID)}">Delete</button>` });
}

function officerRecordField(selected = '') {
  const selectedOfficer = state.officers.find((row) => row.OfficerID === selected);
  const options = state.officers.map((row) => `<button type="button" data-officer-option="${escapeHtml(row.OfficerID)}" data-officer-label="${escapeHtml(`${row.RobloxUsername} - ${row.Rank}${row.Callsign ? ` / ${row.Callsign}` : ''}`)}"><strong>${escapeHtml(row.RobloxUsername)}</strong><span>${escapeHtml(row.Rank)}${row.Callsign ? ` / ${escapeHtml(row.Callsign)}` : ''}</span></button>`).join('');
  const label = selectedOfficer ? `${selectedOfficer.RobloxUsername} - ${selectedOfficer.Rank}${selectedOfficer.Callsign ? ` / ${selectedOfficer.Callsign}` : ''}` : '';
  return { html: `<label class="wide searchable-officer-field">Officer<input type="search" data-officer-search required autocomplete="off" placeholder="Search by username, callsign or rank" value="${escapeHtml(label)}"><input type="hidden" name="OfficerID" value="${escapeHtml(selected)}"><div class="searchable-officer-options" hidden>${options}</div></label>` };
}

async function ensureOfficerRecords() {
  if (state.officers.length) return;
  const response = await apiCached('listOfficers', {});
  if (response.ok) state.officers = response.rows || [];
}

async function openProbationEditor(record = {}) {
  await ensureOfficerRecords();
  openEditor(record.ProbationID ? 'Edit probation record' : 'Add probation record', [hiddenField('ProbationID', record.ProbationID || ''), officerRecordField(record.OfficerID),
    customSelectField('Stage', 'CustomStage', 'Stage', ['Initial', 'Foundation', 'Independent Patrol', 'Final Review'], record.Stage || 'Initial'), selectField('Status', 'Status', ['Active', 'Paused', 'Completed', 'Extended'], record.Status || 'Active'),
    field('StartDate', 'Start date', 'date', false, dateInputValue(record.StartDate)), field('TargetDate', 'Target date', 'date', false, dateInputValue(record.TargetDate)), field('Progress', 'Progress %', 'number', false, record.Progress || 0),
    field('Requirements', 'Requirements / sign-offs', 'textarea', true, record.Requirements || ''), field('Notes', 'Notes', 'textarea', true, record.Notes || ''),
  ], async (values) => { if (values.Stage === 'Custom') values.Stage = values.CustomStage; return api('saveProbation', values); }, operationsSaveOptions('Probation record saved.', loadDevelopment));
}

async function openPerformanceReviewEditor(record = {}) {
  await ensureOfficerRecords();
  openEditor(record.ReviewID ? 'Edit performance review' : 'Add performance review', [hiddenField('ReviewID', record.ReviewID || ''), officerRecordField(record.OfficerID),
    field('ReviewDate', 'Review date', 'date', false, dateInputValue(record.ReviewDate) || new Date().toISOString().slice(0, 10)), field('PeriodStart', 'Period start', 'date', false, dateInputValue(record.PeriodStart)), field('PeriodEnd', 'Period end', 'date', false, dateInputValue(record.PeriodEnd)),
    selectField('Rating', 'Rating', ['Exceeds Expectations', 'Meets Expectations', 'Development Required', 'Unsatisfactory'], record.Rating || 'Meets Expectations'),
    field('ActivitySummary', 'Activity summary', 'textarea', true, record.ActivitySummary || ''), field('Strengths', 'Strengths', 'textarea', true, record.Strengths || ''), field('Improvements', 'Improvements', 'textarea', true, record.Improvements || ''), field('Objectives', 'Objectives', 'textarea', true, record.Objectives || ''), field('NextReviewDate', 'Next review date', 'date', false, dateInputValue(record.NextReviewDate)),
  ], async (values) => api('savePerformanceReview', values), operationsSaveOptions('Performance review saved.', loadDevelopment));
}

async function openRestrictionEditor(record = {}) {
  await ensureOfficerRecords();
  openEditor(record.RestrictionID ? 'Edit restriction' : 'Add restriction', [hiddenField('RestrictionID', record.RestrictionID || ''), officerRecordField(record.OfficerID),
    customSelectField('RestrictionType', 'CustomRestrictionType', 'Restriction', ['No Driving', 'Modified Duties', 'Training Suspended', 'Operational Restriction', 'Temporary Attachment', 'Other'], record.RestrictionType || 'Modified Duties'),
    field('Details', 'Details', 'textarea', true, record.Details || ''), field('StartsOn', 'Starts on', 'date', false, dateInputValue(record.StartsOn) || new Date().toISOString().slice(0, 10)), field('EndsOn', 'Ends on', 'date', false, dateInputValue(record.EndsOn)), selectField('Status', 'Status', ['Active', 'Expired', 'Removed'], record.Status || 'Active'),
  ], async (values) => { if (values.RestrictionType === 'Custom') values.RestrictionType = values.CustomRestrictionType; return api('saveRestriction', values); }, operationsSaveOptions('Restriction saved.', loadDevelopment));
}

async function loadReports() {
  await showViewOnly('reports');
  const container = document.querySelector('#reportsView');
  const response = await apiCached('commandReports', {});
  if (!response.ok) return renderError(container, response.error);
  const metrics = response.metrics || {};
  container.innerHTML = `<div class="section-head"><h2>Command Reports</h2><button class="ghost" data-export-report>Export CSV</button></div>
    <div class="stat-row">${[stat('Active Officers', metrics.ActiveOfficers || 0), stat('On LOA', metrics.OnLoa || 0), stat('Training Expiring', metrics.TrainingExpiring || 0), stat('Reviews Due', metrics.ReviewsDue || 0), stat('Active Restrictions', metrics.ActiveRestrictions || 0), stat('Open Handovers', metrics.OpenHandovers || 0)].join('')}</div>
    <section class="report-grid">${reportPanel('Training expiry', response.trainingExpiry || [], ['Officer', 'Standard', 'ExpiryDate', 'DaysRemaining'])}${reportPanel('Officer activity', response.activity || [], ['Officer', 'Rank', 'Shifts', 'Hours', 'LastShift', 'Status'])}${reportPanel('Development compliance', response.development || [], ['Officer', 'Probation', 'LastReview', 'NextReview', 'Restriction'])}${reportPanel('Supervisor workload', response.supervisors || [], ['Supervisor', 'AssignedOfficers', 'OpenActions'])}</section>`;
  state.operations.report = response;
}

function reportPanel(title, rows, columns) {
  return `<article class="dashboard-panel report-panel"><h3>${escapeHtml(title)}</h3><div class="table-wrap compact"><table><thead><tr>${columns.map((column) => `<th>${escapeHtml(column)}</th>`).join('')}</tr></thead><tbody>${rows.length ? rows.map((row) => `<tr>${columns.map((column) => `<td data-label="${escapeHtml(column)}">${formatCell(row[column], column)}</td>`).join('')}</tr>`).join('') : '<tr><td colspan="6">No records found.</td></tr>'}</tbody></table></div></article>`;
}

function exportCommandReport() {
  const report = state.operations.report || {};
  const rows = [['Section', 'Officer', 'Item', 'Value', 'Date']];
  (report.trainingExpiry || []).forEach((row) => rows.push(['Training expiry', row.Officer, row.Standard, row.DaysRemaining, row.ExpiryDate]));
  (report.activity || []).forEach((row) => rows.push(['Activity', row.Officer, row.Rank, row.Hours, row.LastShift]));
  (report.development || []).forEach((row) => rows.push(['Development', row.Officer, row.Probation, row.Restriction, row.NextReview]));
  const csv = rows.map((row) => row.map((value) => `"${String(value || '').replaceAll('"', '""')}"`).join(',')).join('\n');
  const link = document.createElement('a');
  link.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
  link.download = `mo8-command-report-${new Date().toISOString().slice(0, 10)}.csv`;
  link.click();
  URL.revokeObjectURL(link.href);
}

async function loadHandover() {
  await showViewOnly('handover');
  const response = await apiCached('listHandovers', {});
  if (!response.ok) return renderError(document.querySelector('#handoverBoard'), response.error);
  state.operations.handovers = response.rows || [];
  renderHandoverBoard();
}

function renderHandoverBoard() {
  const container = document.querySelector('#handoverBoard');
  if (!container) return;
  const query = document.querySelector('#handoverSearch')?.value.toLowerCase() || '';
  const status = document.querySelector('#handoverStatusFilter')?.value || '';
  const rows = state.operations.handovers.filter((row) => (!status || row.Status === status) && Object.values(row).some((value) => String(value || '').toLowerCase().includes(query)));
  container.innerHTML = rows.length ? ['Critical', 'High', 'Normal', 'Low'].map((priority) => `<section class="handover-column"><h3>${priority}</h3>${rows.filter((row) => row.Priority === priority).map((row) => `<article class="handover-card status-${escapeHtml(row.Status.toLowerCase().replaceAll(' ', '-'))}"><span>${escapeHtml(row.Category)}</span><h4>${escapeHtml(row.Title)}</h4><p>${escapeHtml(row.Details)}</p><small>Owner: ${escapeHtml(row.Owner || 'Unassigned')}${row.Recipients ? ` / Sent to: ${escapeHtml(row.Recipients)}` : ''}${row.DueAt ? ` / ${formatDisplayDateTime(row.DueAt)}` : ''}</small><div class="row-actions"><button class="mini" data-edit-handover="${escapeHtml(row.HandoverID)}">Open</button><button class="mini danger" data-delete-handover="${escapeHtml(row.HandoverID)}">Delete</button></div></article>`).join('') || '<p class="empty">None.</p>'}</section>`).join('') : emptyState('No handover entries found.');
}

async function openHandoverEditor(record = {}) {
  const response = await apiCached('listUsers', {});
  const users = response.ok ? response.rows || [] : [state.user].filter(Boolean);
  openEditor(record.HandoverID ? 'Edit handover' : 'Add handover', [hiddenField('HandoverID', record.HandoverID || ''), field('Title', 'Title', 'text', false, record.Title || ''), selectField('Category', 'Category', ['Operational', 'Staffing', 'Training', 'Conduct', 'Welfare', 'System', 'Other'], record.Category || 'Operational'), selectField('Priority', 'Priority', ['Critical', 'High', 'Normal', 'Low'], record.Priority || 'Normal'), supervisorSelectField('OwnerUserID', 'Owner', users, record.OwnerUserID || state.user?.UserID || ''), userCheckboxGroupField('RecipientUserIDs', 'Send to officers', users, record.RecipientUserIDs || ''), field('Details', 'Details', 'textarea', true, record.Details || ''), field('DueAt', 'Due at', 'datetime-local', false, localDateTimeValue(record.DueAt)), selectField('Status', 'Status', ['Open', 'In Progress', 'Resolved'], record.Status || 'Open'), field('Resolution', 'Resolution', 'textarea', true, record.Resolution || '')], async (values) => api('saveHandover', values), operationsSaveOptions('Handover saved.', loadHandover));
}

function operationsSaveOptions(message, loader) {
  return { successMessage: message, onSuccess: async () => { invalidateCache(); await loader(); } };
}

async function reloadDevelopmentContext() {
  invalidateCache();
  if (state.activeView === 'officerProfile') return loadOfficerProfile(state.selectedOfficerId);
  return loadDevelopment();
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
      ${detailCard('Supervisor', officer ? supervisorProfileLink(officer) : 'No record', true)}
      ${detailCard('Discord ID', user.DiscordID || 'Not set')}
      ${detailCard('Unread notices', String(notifications.filter((item) => !item.ReadAt).length))}
    </section>
    ${officer ? tagList('Officer Tags', officer.Tags) : ''}
    ${officer ? trainingChecklist(officer.OfficerID, response.training || []) : ''}
    ${profileTable('My Upcoming Training', response.upcomingTraining || [], ['Title', 'Standard', 'CourseDate', 'Status', 'Location'])}
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
    ${profileTable('My Probation / Competency', response.probation || [], ['Stage', 'Status', 'Progress', 'StartDate', 'TargetDate', 'Requirements', 'Notes'])}
    ${profileTable('My Performance Reviews', response.performanceReviews || [], ['ReviewDate', 'Rating', 'Strengths', 'Improvements', 'Objectives', 'NextReviewDate'])}
    ${profileTable('My Temporary Restrictions', response.restrictions || [], ['RestrictionType', 'Details', 'StartsOn', 'EndsOn', 'Status'])}
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
    ...(response.pendingCourseBookings || []),
    ...(response.pendingAppeals || []),
    ...(response.probationReviews || []),
    ...(response.performanceReviews || []),
    ...(response.restrictionReviews || []),
    ...(response.accountRequests || []),
    ...(response.retrospectiveShifts || []),
  ];
  const counts = response.counts || {};
  document.querySelector('#tasksSummary').innerHTML = [
    stat('Pending LOA', counts.pendingLoa || 0),
    stat('Transfer Requests', counts.pendingTransfers || 0),
    stat('Supervisor Requests', counts.pendingSupervisorRequests || 0),
    stat('Course Requests', counts.pendingCourseBookings || 0),
    stat('Appeals', counts.pendingAppeals || 0),
    stat('Development Reviews', counts.developmentReviews || 0),
    stat('Account Requests', counts.accountRequests || 0),
    stat('Shift Requests', counts.retrospectiveShifts || 0),
    stat('Your Supervisees', counts.mySuperviseeTasks || 0),
    stat('Total Tasks', counts.total || 0),
  ].join('');
  renderTable('#tasksTable', state.tasks, ['TaskType', 'Officer', 'Rank', 'Supervisor', 'Course', 'Subject', 'SourceType', 'StartDate', 'EndDate', 'TargetDivision', 'Reason'], {
    rowAction: (row) => `${row.MySupervisee ? 'class="supervisor-task"' : ''} ${taskOpenAttr(row)}`,
    actions: (row) => `<button class="mini" ${taskOpenAttr(row)}>Review</button>`,
  });
}

function taskOpenAttr(row) {
  if (row.TaskType === 'Account Request') return `data-open-account-request="${escapeHtml(row.RequestID)}"`;
  if (row.TaskType === 'Retrospective Shift') return `data-open-retrospective-shift="${escapeHtml(row.RequestID)}"`;
  if (row.TaskType === 'Probation Review') return `data-edit-probation="${escapeHtml(row.ProbationID)}"`;
  if (row.TaskType === 'Performance Review') return `data-task-performance-review="${escapeHtml(row.OfficerID)}"${row.ReviewID ? ` data-review-id="${escapeHtml(row.ReviewID)}"` : ''}`;
  if (row.TaskType === 'Restriction Review') return `data-edit-restriction="${escapeHtml(row.RestrictionID)}"`;
  if (row.TaskType === 'Transfer Request') return `data-open-transfer-review="${escapeHtml(row.RequestID)}"`;
  if (row.TaskType === 'Supervisor Request') return `data-open-supervisor-review="${escapeHtml(row.RequestID)}"`;
  if (row.TaskType === 'Appeal / Review') return `data-open-appeal-review="${escapeHtml(row.AppealID)}"`;
  if (row.TaskType === 'Course Booking') return `data-task-course-bookings="${escapeHtml(row.CourseID)}"`;
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
    stat('Your Shifts', teamResponse.myStats?.Shifts || 0),
    stat('Your Duration', teamResponse.myStats?.Duration || '0h 0m'),
  ].join('');

  renderTable('#activeShiftsTable', teamResponse.active || [], ['RobloxUsername', 'Callsign', 'Rank', 'LoaStatus', 'StartedAt']);
  renderTable('#recentShiftsTable', teamResponse.recent || [], ['RobloxUsername', 'Callsign', 'StartedAt', 'EndedAt', 'Duration', 'Summary'], {
    actions: (row) => can('VIEW_TASKS') ? `<button class="mini" data-edit-shift="${escapeHtml(row.ShiftID)}">Edit</button>` : '',
  });
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
  if (can('MANAGE_COURSES') && !state.selectedCourseId) {
    state.selectedCourseId = state.courses.find((course) => Number(course.PendingRequests || 0) > 0)?.CourseID || '';
  }
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
  const currentPath = normalizeFolderPath(state.documentFolder || document.querySelector('#documentCategoryFilter').value);
  syncDocumentCategoryFilter(currentPath);
  const rows = state.documents.filter((document) => {
    const matchesQuery = ['Title', 'Category', 'FolderPath', 'FileName', 'RequiredRole', 'RequiredTags', 'Status'].some((field) => String(document[field] || '').toLowerCase().includes(query));
    if (query) return matchesQuery;
    return documentParentPath(documentFolderName(document)) === currentPath;
  });
  renderDocumentExplorer(rows, query, currentPath);
}

function renderDocumentExplorer(rows, query, currentPath) {
  const container = document.querySelector('#documentExplorer');
  const folders = query ? [] : childFolders(currentPath);
  const folderTiles = folders.map((folder) => {
    const path = currentPath ? `${currentPath}/${folder}` : folder;
    const count = state.documents.filter((document) => documentFolderName(document) === path || documentFolderName(document).startsWith(`${path}/`)).length;
    return `
      <button class="folder-tile" data-doc-folder="${escapeHtml(path)}">
        <span class="folder-icon"></span>
        <strong>${escapeHtml(folder)}</strong>
        <small>${count} ${count === 1 ? 'file' : 'files'}</small>
      </button>
    `;
  }).join('');

  const files = rows.map((document) => `
    <article class="file-row">
      <a class="file-main" href="${escapeHtml(document.DriveURL || '#')}" target="_blank" rel="noopener">
        <span class="file-icon">${escapeHtml(fileInitial(document))}</span>
        <span>
          <strong>${escapeHtml(document.Title || 'Untitled document')}</strong>
          <small>${escapeHtml(document.Category || 'Unfiled')} / ${escapeHtml(document.FileName || document.RequiredRole || 'Constable+')}</small>
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
        ${folderBreadcrumb(currentPath)}
        ${query ? `<span>/</span><strong>Search results</strong>` : ''}
      </div>
      <div class="explorer-actions">
        ${currentPath || query ? `<button class="ghost mini" data-doc-folder="">All folders</button>` : ''}
        ${currentPath ? `<button class="ghost mini" data-doc-folder="${escapeHtml(documentParentPath(currentPath))}">Up one level</button>` : ''}
      </div>
    </div>
    ${folderTiles ? `<section class="folder-grid">${folderTiles}</section>` : ''}
    <section class="file-list">
      <div class="file-list-head">
        <strong>${query ? 'Matching documents' : currentPath ? escapeHtml(folderBaseName(currentPath)) : 'All documents'}</strong>
        <span>${rows.length} ${rows.length === 1 ? 'file' : 'files'}</span>
      </div>
      ${files || `<p class="empty">No documents found.</p>`}
    </section>
  `;
}

function documentFolders() {
  const folders = state.documents.flatMap((document) => folderAncestors(documentFolderName(document))).filter(Boolean);
  return [...new Set(folders)].sort((a, b) => a.localeCompare(b));
}

function documentFolderName(document) {
  return normalizeFolderPath(document.FolderPath || document.Category || 'Unfiled');
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

function childFolders(parentPath = '') {
  const children = new Set();
  documentFolders().forEach((folder) => {
    if (documentParentPath(folder) === parentPath) children.add(folderBaseName(folder));
  });
  return [...children].sort((a, b) => a.localeCompare(b));
}

function folderAncestors(path = '') {
  const parts = normalizeFolderPath(path).split('/').filter(Boolean);
  return parts.map((_, index) => parts.slice(0, index + 1).join('/'));
}

function folderBreadcrumb(path = '') {
  const parts = normalizeFolderPath(path).split('/').filter(Boolean);
  return parts.map((part, index) => {
    const crumbPath = parts.slice(0, index + 1).join('/');
    const isLast = index === parts.length - 1;
    return `<span>/</span>${isLast ? `<strong>${escapeHtml(part)}</strong>` : `<button data-doc-folder="${escapeHtml(crumbPath)}">${escapeHtml(part)}</button>`}`;
  }).join('');
}

function normalizeFolderPath(path = '') {
  return String(path || '')
    .split('/')
    .map((part) => part.trim())
    .filter(Boolean)
    .join('/');
}

function documentParentPath(path = '') {
  const parts = normalizeFolderPath(path).split('/').filter(Boolean);
  parts.pop();
  return parts.join('/');
}

function folderBaseName(path = '') {
  const parts = normalizeFolderPath(path).split('/').filter(Boolean);
  return parts[parts.length - 1] || 'Documents';
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
  renderTable('#coursesTable', rows, ['Title', 'Standard', 'Trainer', 'CoTrainers', 'CourseDate', 'Location', 'Capacity', 'BookedSeats', 'PendingRequests', 'Waitlist', 'Status', 'MyBookingStatus'], {
    actions: (row) => [
      can('VIEW_COURSES') && !row.MyBookingStatus ? `<button class="mini" data-request-course="${escapeHtml(row.CourseID)}">Request seat</button>` : '',
      canManageCourse(row) ? `<button class="mini ${Number(row.PendingRequests || 0) > 0 ? 'warning-action' : 'ghost'}" data-course-bookings="${escapeHtml(row.CourseID)}">Review${Number(row.PendingRequests || 0) > 0 ? ` (${escapeHtml(row.PendingRequests)})` : ''}</button>` : '',
      can('MANAGE_COURSES') ? `<button class="mini" data-edit-course="${escapeHtml(row.CourseID)}">Edit</button>` : '',
      can('MANAGE_COURSES') ? `<button class="mini ghost" data-delete-course="${escapeHtml(row.CourseID)}">Delete</button>` : '',
    ].join(''),
  });
}

function renderCourseBookingsTable(courseId = state.selectedCourseId) {
  const selectedCourse = state.courses.find((row) => row.CourseID === courseId);
  const bookings = (state.courseBookings || []).filter((row) => !courseId || row.CourseID === courseId);
  const panelTitle = document.querySelector('#courseBookingsTitle');
  if (panelTitle) {
    panelTitle.textContent = selectedCourse
      ? `Bookings / ${selectedCourse.Title}${Number(selectedCourse.PendingRequests || 0) > 0 ? ` / ${selectedCourse.PendingRequests} waiting` : ''}`
      : 'Bookings';
  }
  renderTable('#courseBookingsTable', bookings, ['Course', 'Officer', 'Rank', 'Status', 'Outcome', 'RequestedAt'], {
    emptyMessage: selectedCourse ? 'No bookings for this course yet.' : 'Select Review on a course, or wait for bookings to appear.',
    actions: (row) => {
      if (!canManageCourse(selectedCourse)) return '';
      return [
        ['Approved', 'Approve'],
        ['Waitlist', 'Waitlist'],
        ['Denied', 'Deny'],
        ['Completed', 'Complete'],
        ['Cancelled', 'Cancel'],
      ].map(([status, label]) => `<button class="mini ${row.Status === status ? 'success-action' : ''}" data-quick-booking-status="${escapeHtml(status)}" data-booking-id="${escapeHtml(row.BookingID)}">${label}</button>`).join('')
        + `<button class="mini ghost" data-review-booking="${escapeHtml(row.BookingID)}">Full review</button>`;
    },
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
    renderTrainingSpreadsheet(summaryRows);
  }
  if (view === 'courses') {
    renderCoursesTable(rows);
    if (state.selectedCourseId && !rows.some((row) => row.CourseID === state.selectedCourseId)) state.selectedCourseId = '';
    renderCourseBookingsTable();
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

function renderTrainingSpreadsheet(rows) {
  const table = document.querySelector('#trainingTable');
  const specialistOptions = trainingOptionNames('Specialist');
  const drivingStandards = trainingOptionNames('Driving');
  const disabled = can('MANAGE_TRAINING') ? '' : ' disabled';
  table.className = 'training-matrix-table';
  if (!rows.length) {
    table.innerHTML = '<tbody><tr><td>No officers found.</td></tr></tbody>';
    return;
  }
  const rankIndex = (rank) => OFFICER_RANKS.indexOf(rank);
  const groups = [
    ['Senior Officers', rows.filter((row) => rankIndex(row.Rank) >= rankIndex('Inspector'))],
    ['Supervisory Officers', rows.filter((row) => row.Rank === 'Sergeant')],
    ['Police Constables', rows.filter((row) => rankIndex(row.Rank) < rankIndex('Sergeant') || rankIndex(row.Rank) < 0)],
  ].filter(([, groupRows]) => groupRows.length);
  const header = `<thead><tr><th>Callsign</th><th>Officer</th><th>Rank</th><th>Driving</th>${specialistOptions.map((standard) => `<th>${escapeHtml(standard)}</th>`).join('')}<th>Review Date</th></tr></thead>`;
  const body = groups.map(([groupName, groupRows]) => `<tbody><tr class="training-group-row"><th colspan="${specialistOptions.length + 5}">${escapeHtml(groupName)}</th></tr>${groupRows.map((officer) => {
    const records = state.training.filter((record) => record.OfficerID === officer.OfficerID);
    const passed = (standard) => records.some((record) => record.Standard === standard && record.Status === 'Passed');
    const driving = drivingStandards.find(passed) || '';
    const drivingOptions = [''].concat(drivingStandards).map((standard) => `<option value="${escapeHtml(standard)}"${standard === driving ? ' selected' : ''}>${escapeHtml(standard || 'Not set')}</option>`).join('');
    const reviewDate = records.find((record) => record.ReviewDate)?.ReviewDate || '';
    return `<tr data-training-officer="${escapeHtml(officer.OfficerID)}"><td data-label="Callsign">${escapeHtml(officer.Callsign || '-')}</td><td data-label="Officer"><button class="training-officer-link" data-open-officer="${escapeHtml(officer.OfficerID)}">${escapeHtml(officer.RobloxUsername)}</button></td><td data-label="Rank"><span class="rank-cell">${escapeHtml(officer.Rank)}</span></td><td data-label="Driving"><select data-driving-select data-training-context="matrix" data-officer-id="${escapeHtml(officer.OfficerID)}"${disabled}>${drivingOptions}</select></td>${specialistOptions.map((standard) => `<td class="training-toggle-cell" data-label="${escapeHtml(standard)}"><input type="checkbox" data-training-toggle data-training-context="matrix" data-officer-id="${escapeHtml(officer.OfficerID)}" data-standard="${escapeHtml(standard)}"${passed(standard) ? ' checked' : ''}${disabled} aria-label="${escapeHtml(`${officer.RobloxUsername}: ${standard}`)}"></td>`).join('')}<td data-label="Review Date"><input type="date" data-training-review data-training-context="matrix" data-officer-id="${escapeHtml(officer.OfficerID)}" value="${escapeHtml(dateInputValue(reviewDate))}"${disabled}></td></tr>`;
  }).join('')}</tbody>`).join('');
  table.innerHTML = header + body;
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
      return `<td data-label="${escapeHtml(role)}"><input type="checkbox" data-role-permission data-role="${escapeHtml(role)}" data-permission="${escapeHtml(permission)}"${enabled}${disabled}></td>`;
    }).join('');
    return `<tr><td data-label="Permission">${escapeHtml(permission)}</td>${cells}</tr>`;
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
      return `<td data-label="${escapeHtml(permission)}"><select data-user-permission data-user-id="${escapeHtml(user.UserID)}" data-permission="${escapeHtml(permission)}">${options}</select></td>`;
    }).join('');
    return `<tr><td data-label="User">${escapeHtml(user.RobloxUsername)}</td><td data-label="Role">${escapeHtml(user.Role)}</td>${cells}</tr>`;
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
  state.operations.probation = data.probation || [];
  state.operations.reviews = data.performanceReviews || [];
  state.operations.restrictions = data.restrictions || [];
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
      ${detailCard('Supervisor', supervisorProfileLink(officer), true)}
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
      ${profileTable('Probation / Competency', data.probation || [], ['Stage', 'Status', 'Progress', 'StartDate', 'TargetDate', 'Requirements', 'Notes'], { actions: (row) => can('VIEW_TASKS') ? `<button class="mini" data-edit-probation="${escapeHtml(row.ProbationID)}">Edit</button><button class="mini ghost" data-delete-probation="${escapeHtml(row.ProbationID)}">Delete</button>` : '' })}
      ${profileTable('Performance Reviews', data.performanceReviews || [], ['ReviewDate', 'Rating', 'Strengths', 'Improvements', 'Objectives', 'NextReviewDate'], { actions: (row) => can('VIEW_TASKS') ? `<button class="mini" data-edit-performance-review="${escapeHtml(row.ReviewID)}">Edit</button><button class="mini ghost" data-delete-performance-review="${escapeHtml(row.ReviewID)}">Delete</button>` : '' })}
      ${profileTable('Temporary Restrictions', data.restrictions || [], ['RestrictionType', 'Details', 'StartsOn', 'EndsOn', 'Status'], { actions: (row) => can('VIEW_TASKS') ? `<button class="mini" data-edit-restriction="${escapeHtml(row.RestrictionID)}">Edit</button><button class="mini ghost" data-delete-restriction="${escapeHtml(row.RestrictionID)}">Delete</button>` : '' })}
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
    userCheckboxGroupField('CoTrainerUserIDs', 'Co-trainers', trainers, course.CoTrainerUserIDs),
    field('CourseDate', 'Course date/time', 'datetime-local', false, localDateTimeValue(course.CourseDate)),
    field('Location', 'Location', 'text', false, course.Location),
    field('Capacity', 'Capacity', 'number', false, course.Capacity || '4'),
    selectField('Status', 'Status', ['Scheduled', 'Completed', 'Cancelled'], course.Status || 'Scheduled'),
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

function openRetrospectiveShiftEditor() {
  openEditor('Request retrospective shift', [
    field('StartedAt', 'Shift started', 'datetime-local'),
    field('EndedAt', 'Shift ended', 'datetime-local'),
    field('Summary', 'Shift summary', 'textarea', true),
    field('Reason', 'Why was the shift not logged?', 'textarea', true),
  ], async (values) => api('requestRetrospectiveShift', values), {
    successMessage: 'Retrospective shift sent to your supervisor.',
    onSuccess: async () => { invalidateCache('tasks'); invalidateCache('teamShifts'); await loadShift(); },
  });
}

function openRetrospectiveShiftReview(record) {
  openEditor('Review retrospective shift', [
    hiddenField('RequestID', record.RequestID),
    field('Officer', 'Officer', 'text', false, record.Officer),
    field('StartedAt', 'Shift started', 'datetime-local', false, localDateTimeValue(record.StartedAt)),
    field('EndedAt', 'Shift ended', 'datetime-local', false, localDateTimeValue(record.EndedAt)),
    field('Summary', 'Shift summary', 'textarea', true, record.Summary),
    field('Reason', 'Request reason', 'textarea', true, record.Reason),
    selectField('Status', 'Decision', ['Approved', 'Denied'], 'Approved'),
    field('ReviewReason', 'Decision notes', 'textarea', true, record.ReviewReason),
  ], async (values) => api('reviewRetrospectiveShift', values), {
    successMessage: 'Retrospective shift reviewed.',
    onSuccess: async () => { invalidateCache('tasks'); invalidateCache('teamShifts'); await loadTasks(); },
  });
}

function openAccountRequestReview(record) {
  openEditor('Review account request', [
    hiddenField('RequestID', record.RequestID),
    field('RobloxUsername', 'Roblox username', 'text', false, record.RobloxUsername),
    selectField('Rank', 'Rank', OFFICER_RANKS, record.Rank),
    field('DiscordID', 'Discord user ID', 'text', false, record.DiscordID),
    selectField('Status', 'Decision', ['Approved', 'Denied'], 'Approved'),
    field('ReviewNotes', 'Decision notes', 'textarea', true, record.ReviewNotes),
  ], async (values) => supabaseInvokeAdminUsers(Object.assign({ action: 'reviewAccountRequest' }, values)), {
    successMessage: 'Account request reviewed.',
    onSuccess: async (response) => {
      invalidateCache('tasks');
      await loadTasks();
      if (response?.temporaryPassword) showInfo('Account created', `<p>The account was created and its temporary credentials were sent by Discord.</p>`);
    },
  });
}

function openShiftEditor(record) {
  openEditor('Edit shift', [
    hiddenField('ShiftID', record.ShiftID),
    field('Officer', 'Officer', 'text', false, record.RobloxUsername),
    field('StartedAt', 'Started at', 'datetime-local', false, localDateTimeValue(record.StartedAt)),
    field('EndedAt', 'Ended at', 'datetime-local', false, localDateTimeValue(record.EndedAt)),
    field('Summary', 'Summary', 'textarea', true, record.Summary),
  ], async (values) => api('saveShift', values), {
    successMessage: 'Shift updated.',
    onSuccess: async () => { invalidateCache('teamShifts'); await loadShift(); },
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
    field('Category', 'Folder path', 'text', false, documentFolderName(document) || state.documentFolder || 'General'),
    fileField('DocumentFile', document.FileName ? `Replace uploaded file (${document.FileName})` : 'Upload file'),
    field('DriveURL', 'External URL', 'url', false, document.StoragePath ? '' : document.DriveURL),
    selectField('RequiredRole', 'Minimum rank', ACCESS_LEVELS, document.RequiredRole || 'Police Constable'),
    checkboxGroupField('RequiredTags', 'Required tags', OFFICER_TAGS, document.RequiredTags),
    selectField('RequiresAcknowledgement', 'Requires acknowledgement', ['FALSE', 'TRUE'], truthy(document.RequiresAcknowledgement) ? 'TRUE' : 'FALSE'),
    selectField('Status', 'Status', ['Published', 'Draft', 'Archived'], document.Status),
  ], async (values) => api('saveDocument', values));
}

function openFolderEditor() {
  openEditor('New folder', [
    field('ParentPath', 'Parent folder', 'text', false, state.documentFolder || ''),
    field('FolderName', 'Folder name', 'text', false),
  ], async (values) => {
    const folderPath = normalizeFolderPath([values.ParentPath, values.FolderName].filter(Boolean).join('/'));
    if (!folderPath) return { ok: false, error: 'Folder name is required.' };
    state.documentFolder = folderPath;
    renderDocumentTable();
    return { ok: true };
  }, {
    successMessage: 'Folder created. Add a document to store files in it.',
  });
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

function handleDashboardPointerDown(event) {
  const dragHandle = event.target.closest('[data-widget-drag]');
  const resizeHandle = event.target.closest('[data-widget-resize]');
  if (!dragHandle && !resizeHandle) return;

  const card = event.target.closest('[data-dashboard-widget]');
  if (!card) return;
  event.preventDefault();

  const mode = resizeHandle ? 'resize' : 'drag';
  const activeWidgets = getCachedResponse('dashboard', {})?.widgets || DASHBOARD_WIDGETS.map(([item]) => item);
  const layout = getDashboardLayout(activeWidgets);
  state.dashboardInteraction = {
    mode,
    card,
    key: card.dataset.dashboardWidget,
    startX: event.clientX,
    startY: event.clientY,
    startSize: layout.sizes[card.dataset.dashboardWidget] || 'normal',
  };
  card.classList.add(mode === 'resize' ? 'is-resizing' : 'is-dragging');
  document.body.classList.add('dashboard-layout-active');
}

function handleDashboardPointerMove(event) {
  const interaction = state.dashboardInteraction;
  if (!interaction) return;
  event.preventDefault();

  if (interaction.mode === 'resize') {
    resizeDashboardWidget(interaction, event.clientX);
    return;
  }

  const target = document.elementFromPoint(event.clientX, event.clientY)?.closest?.('[data-dashboard-widget]');
  if (!target || target === interaction.card) return;

  const grid = interaction.card.closest('.dashboard-widget-grid');
  if (!grid || target.closest('.dashboard-widget-grid') !== grid) return;

  const targetBox = target.getBoundingClientRect();
  const before = event.clientY < targetBox.top + targetBox.height / 2;
  grid.insertBefore(interaction.card, before ? target : target.nextSibling);
  saveDashboardDomOrder();
}

function handleDashboardPointerUp() {
  const interaction = state.dashboardInteraction;
  if (!interaction) return;
  interaction.card.classList.remove('is-dragging', 'is-resizing');
  document.body.classList.remove('dashboard-layout-active');
  state.dashboardInteraction = null;
}

function resizeDashboardWidget(interaction, clientX) {
  const sizes = ['normal', 'wide', 'large'];
  const startIndex = Math.max(0, sizes.indexOf(interaction.startSize));
  const moved = clientX - interaction.startX;
  const nextIndex = Math.max(0, Math.min(sizes.length - 1, startIndex + Math.round(moved / 140)));
  const nextSize = sizes[nextIndex];
  if (interaction.card.dataset.widgetSize === nextSize) return;

  interaction.card.dataset.widgetSize = nextSize;
  sizes.forEach((size) => interaction.card.classList.toggle(`widget-${size}`, size === nextSize));
  updateDashboardWidget(interaction.key, (layout, key) => {
    layout.sizes[key] = nextSize;
    layout.order = [...document.querySelectorAll('[data-dashboard-widget]')].map((card) => card.dataset.dashboardWidget);
  });
  const label = interaction.card.querySelector('.dashboard-widget-title small');
  if (label) label.textContent = nextSize === 'large' ? 'Large widget' : nextSize === 'wide' ? 'Wide widget' : 'Standard widget';
}

async function handleDocumentClick(event) {
  if (event.target.closest('[data-widget-drag], [data-widget-resize]')) return;
  if (!event.target.closest('.notification-shell')) {
    closeNotificationMenu();
  }
  if (document.body.classList.contains('mobile-nav-open')
    && !event.target.closest('.module-dock')
    && !event.target.closest('.mobile-menu-button')) {
    closeMobileNav();
  }

  const officerOption = event.target.closest('[data-officer-option]');
  if (officerOption) {
    const field = officerOption.closest('.searchable-officer-field');
    const input = field.querySelector('[data-officer-search]');
    input.value = officerOption.dataset.officerLabel;
    input.setCustomValidity('');
    field.querySelector('input[name="OfficerID"]').value = officerOption.dataset.officerOption;
    field.querySelector('.searchable-officer-options').hidden = true;
    return;
  }
  const officerSearch = event.target.closest('[data-officer-search]');
  if (officerSearch) {
    officerSearch.closest('.searchable-officer-field').querySelector('.searchable-officer-options').hidden = false;
  }
  if (!event.target.closest('.searchable-officer-field')) {
    document.querySelectorAll('.searchable-officer-options').forEach((options) => { options.hidden = true; });
  }

  const editCalendar = event.target.closest('[data-edit-calendar]');
  if (editCalendar) {
    const record = state.operations.calendar.find((row) => row.ID === editCalendar.dataset.editCalendar);
    if (elements.infoDialog.open) elements.infoDialog.close();
    if (record) openCalendarEventEditor(record);
    return;
  }
  const calendarDay = event.target.closest('[data-calendar-day]');
  if (calendarDay) {
    showCalendarDay(calendarDay.dataset.calendarDay);
    return;
  }
  const calendarEvent = event.target.closest('[data-calendar-event-detail]');
  if (calendarEvent) {
    showCalendarEvent(calendarEvent.dataset.calendarEventDetail);
    return;
  }
  const editProbation = event.target.closest('[data-edit-probation]');
  if (editProbation) {
    const record = state.operations.probation.find((row) => row.ProbationID === editProbation.dataset.editProbation) || state.tasks.find((row) => row.ProbationID === editProbation.dataset.editProbation);
    if (record) openProbationEditor(record);
    return;
  }
  const deleteProbation = event.target.closest('[data-delete-probation]');
  if (deleteProbation) {
    await confirmDelete('Delete this probation record?', 'deleteProbation', { ProbationID: deleteProbation.dataset.deleteProbation }, reloadDevelopmentContext);
    return;
  }
  const editReview = event.target.closest('[data-edit-performance-review]');
  if (editReview) {
    const record = state.operations.reviews.find((row) => row.ReviewID === editReview.dataset.editPerformanceReview);
    if (record) openPerformanceReviewEditor(record);
    return;
  }
  const taskPerformanceReview = event.target.closest('[data-task-performance-review]');
  if (taskPerformanceReview) {
    const record = state.operations.reviews.find((row) => row.ReviewID === taskPerformanceReview.dataset.reviewId)
      || state.tasks.find((row) => row.ReviewID === taskPerformanceReview.dataset.reviewId)
      || { OfficerID: taskPerformanceReview.dataset.taskPerformanceReview };
    openPerformanceReviewEditor(record);
    return;
  }
  const deletePerformanceReview = event.target.closest('[data-delete-performance-review]');
  if (deletePerformanceReview) {
    await confirmDelete('Delete this performance review?', 'deletePerformanceReview', { ReviewID: deletePerformanceReview.dataset.deletePerformanceReview }, reloadDevelopmentContext);
    return;
  }
  const editRestriction = event.target.closest('[data-edit-restriction]');
  if (editRestriction) {
    const record = state.operations.restrictions.find((row) => row.RestrictionID === editRestriction.dataset.editRestriction) || state.tasks.find((row) => row.RestrictionID === editRestriction.dataset.editRestriction);
    if (record) openRestrictionEditor(record);
    return;
  }
  const deleteRestriction = event.target.closest('[data-delete-restriction]');
  if (deleteRestriction) {
    await confirmDelete('Delete this restriction?', 'deleteRestriction', { RestrictionID: deleteRestriction.dataset.deleteRestriction }, reloadDevelopmentContext);
    return;
  }
  const editHandover = event.target.closest('[data-edit-handover]');
  if (editHandover) {
    const record = state.operations.handovers.find((row) => row.HandoverID === editHandover.dataset.editHandover);
    if (record) openHandoverEditor(record);
    return;
  }
  const editShift = event.target.closest('[data-edit-shift]');
  if (editShift) {
    const record = state.shifts.find((row) => row.ShiftID === editShift.dataset.editShift);
    if (record) openShiftEditor(record);
    return;
  }
  const retrospectiveShift = event.target.closest('[data-open-retrospective-shift]');
  if (retrospectiveShift) {
    const record = state.tasks.find((row) => row.RequestID === retrospectiveShift.dataset.openRetrospectiveShift);
    if (record) openRetrospectiveShiftReview(record);
    return;
  }
  const accountRequest = event.target.closest('[data-open-account-request]');
  if (accountRequest) {
    const record = state.tasks.find((row) => row.RequestID === accountRequest.dataset.openAccountRequest);
    if (record) openAccountRequestReview(record);
    return;
  }
  const deleteHandover = event.target.closest('[data-delete-handover]');
  if (deleteHandover) {
    await confirmDelete('Delete this handover?', 'deleteHandover', { HandoverID: deleteHandover.dataset.deleteHandover }, loadHandover);
    return;
  }
  if (event.target.closest('[data-export-report]')) {
    exportCommandReport();
    return;
  }

  const resetDashboardLayout = event.target.closest('[data-reset-dashboard-layout]');
  if (resetDashboardLayout) {
    localStorage.removeItem(DASHBOARD_LAYOUT_STORAGE_KEY);
    await loadDashboard();
    return;
  }

  const viewLink = event.target.closest('[data-view-link]');
  if (viewLink) {
    await showView(viewLink.dataset.viewLink);
    return;
  }

  const officerLink = event.target.closest('[data-open-officer]');
  if (officerLink && (officerLink.matches('button, a') || !event.target.closest('button, a, input, select, textarea'))) {
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
    state.documentFolder = normalizeFolderPath(documentFolder.dataset.docFolder || '');
    document.querySelector('#documentSearch').value = '';
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

  const deleteCourse = event.target.closest('[data-delete-course]');
  if (deleteCourse) {
    await confirmDelete('Delete this training course and notify affected attendees?', 'deleteTrainingCourse', { CourseID: deleteCourse.dataset.deleteCourse }, loadCourses);
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

  const courseBookings = event.target.closest('[data-course-bookings]');
  if (courseBookings) {
    state.selectedCourseId = courseBookings.dataset.courseBookings;
    renderCourseBookingsTable();
    document.querySelector('#courseBookingsTable')?.scrollIntoView({ behavior: 'smooth', block: 'center' });
    return;
  }

  const taskCourseBookings = event.target.closest('[data-task-course-bookings]');
  if (taskCourseBookings) {
    state.selectedCourseId = taskCourseBookings.dataset.taskCourseBookings;
    await showView('courses');
    document.querySelector('#courseBookingsTable')?.scrollIntoView({ behavior: 'smooth', block: 'center' });
    return;
  }

  const quickBookingStatus = event.target.closest('[data-quick-booking-status]');
  if (quickBookingStatus) {
    quickBookingStatus.disabled = true;
    quickBookingStatus.textContent = 'Saving...';
    const response = await api('reviewCourseBooking', {
      BookingID: quickBookingStatus.dataset.bookingId,
      Status: quickBookingStatus.dataset.quickBookingStatus,
    });
    if (!response.ok) {
      showInfo('Booking update failed', `<p>${escapeHtml(response.error || 'Could not update this booking.')}</p>`);
      quickBookingStatus.disabled = false;
      return;
    }
    invalidateCache('listTrainingCourses');
    invalidateCache('dashboard');
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
  const customSelect = event.target.closest('[data-custom-select]');
  if (customSelect) {
    const customField = elements.editorFields.querySelector(`[data-custom-field="${CSS.escape(customSelect.dataset.customSelect)}"]`);
    if (customField) {
      customField.hidden = customSelect.value !== 'Custom';
      const input = customField.querySelector('input');
      input.required = customSelect.value === 'Custom';
      if (!input.required) input.value = '';
    }
    return;
  }
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
    if (trainingToggle.dataset.trainingContext === 'matrix') {
      if (trainingToggle.checked) {
        state.training.push({ OfficerID: trainingToggle.dataset.officerId, Standard: trainingToggle.dataset.standard, Status: 'Passed' });
      } else {
        state.training = state.training.filter((record) => !(record.OfficerID === trainingToggle.dataset.officerId && record.Standard === trainingToggle.dataset.standard));
      }
      trainingToggle.disabled = !can('MANAGE_TRAINING');
      return;
    }
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
    if (drivingSelect.dataset.trainingContext === 'matrix') {
      const drivingStandards = trainingOptionNames('Driving');
      state.training = state.training.filter((record) => !(record.OfficerID === drivingSelect.dataset.officerId && drivingStandards.includes(record.Standard)));
      if (drivingSelect.value) state.training.push({ OfficerID: drivingSelect.dataset.officerId, Standard: drivingSelect.value, Status: 'Passed' });
      drivingSelect.disabled = !can('MANAGE_TRAINING');
      return;
    }
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
    if (reviewDate.dataset.trainingContext === 'matrix') {
      const existing = state.training.find((record) => record.OfficerID === reviewDate.dataset.officerId);
      if (existing) existing.ReviewDate = reviewDate.value;
      else state.training.push({ OfficerID: reviewDate.dataset.officerId, Standard: '', Status: '', ReviewDate: reviewDate.value });
      reviewDate.disabled = !can('MANAGE_TRAINING');
      return;
    }
    await loadOfficerProfile(reviewDate.dataset.officerId);
    return;
  }
}

function handleSearchableOfficerInput(event) {
  const input = event.target.closest('[data-officer-search]');
  if (!input) return;
  const field = input.closest('.searchable-officer-field');
  const query = input.value.trim().toLowerCase();
  field.querySelector('input[name="OfficerID"]').value = '';
  input.setCustomValidity('Select an officer from the search results.');
  const options = field.querySelector('.searchable-officer-options');
  let visible = 0;
  options.querySelectorAll('[data-officer-option]').forEach((option) => {
    const show = !query || option.textContent.toLowerCase().includes(query);
    option.hidden = !show;
    if (show) visible += 1;
  });
  options.hidden = false;
  options.classList.toggle('no-results', visible === 0);
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

function getDashboardLayout(activeWidgets = DASHBOARD_WIDGETS.map(([key]) => key)) {
  const defaults = DASHBOARD_WIDGETS.map(([key]) => key).filter((key) => activeWidgets.includes(key));
  let saved = {};
  try {
    saved = JSON.parse(localStorage.getItem(DASHBOARD_LAYOUT_STORAGE_KEY) || '{}');
  } catch {
    saved = {};
  }
  const savedOrder = Array.isArray(saved.order) ? saved.order : [];
  const order = [...savedOrder.filter((key) => activeWidgets.includes(key)), ...defaults.filter((key) => !savedOrder.includes(key))];
  const sizes = { ...(saved.sizes || {}) };
  return { order, sizes };
}

function saveDashboardLayout(layout) {
  localStorage.setItem(DASHBOARD_LAYOUT_STORAGE_KEY, JSON.stringify(layout));
}

function updateDashboardWidget(key, updater) {
  const activeWidgets = getCachedResponse('dashboard', {})?.widgets || DASHBOARD_WIDGETS.map(([item]) => item);
  const layout = getDashboardLayout(activeWidgets);
  updater(layout, key);
  saveDashboardLayout(layout);
}

function saveDashboardDomOrder() {
  const activeWidgets = getCachedResponse('dashboard', {})?.widgets || DASHBOARD_WIDGETS.map(([item]) => item);
  const layout = getDashboardLayout(activeWidgets);
  layout.order = [...document.querySelectorAll('[data-dashboard-widget]')].map((card) => card.dataset.dashboardWidget);
  saveDashboardLayout(layout);
}

function dashboardWidget(key, title, body, layout) {
  const size = layout.sizes[key] || 'normal';
  const sizeLabel = size === 'large' ? 'Large widget' : size === 'wide' ? 'Wide widget' : 'Standard widget';
  return `
    <article class="dashboard-panel dashboard-widget widget-${escapeHtml(size)}" data-dashboard-widget="${escapeHtml(key)}">
      <div class="dashboard-widget-head">
        <button class="dashboard-drag-handle" data-widget-drag type="button" aria-label="Move ${escapeHtml(title)} widget">
          <span></span>
          <span></span>
        </button>
        <div class="dashboard-widget-title">
          <h3>${escapeHtml(title)}</h3>
          <small>${escapeHtml(sizeLabel)}</small>
        </div>
      </div>
      <div class="dashboard-list">${body}</div>
      <button class="dashboard-resize-handle" data-widget-resize type="button" aria-label="Resize ${escapeHtml(title)} widget"></button>
    </article>
  `;
}

function dashboardRows(rows, columns) {
  return rows.length
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
}

function announcementRows(rows) {
  return rows.length
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

function supervisorProfileLink(officer = {}) {
  if (!officer.Supervisor || officer.Supervisor === 'Not assigned') return 'Not assigned';
  if (!officer.SupervisorOfficerID) return escapeHtml(officer.Supervisor);
  return `<button class="link-button" data-open-officer="${escapeHtml(officer.SupervisorOfficerID)}">${escapeHtml(officer.Supervisor)}</button>`;
}

function profileTable(title, rows, columns, options = {}) {
  const actionHeader = options.actions ? '<th>Actions</th>' : '';
  const body = rows.length
    ? rows.map((row) => {
      const actionCell = options.actions ? `<td class="actions" data-label="Actions">${options.actions(row)}</td>` : '';
      return `<tr>${columns.map((column) => `<td data-label="${escapeHtml(column)}">${formatCell(row[column], column)}</td>`).join('')}${actionCell}</tr>`;
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
      <p>Use the personnel matrix below to update specialist tickets, driving level, and review dates.</p>
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
    const actionCell = options.actions ? `<td class="actions" data-label="Actions">${options.actions(row)}</td>` : '';
    return `<tr ${attrs}>${columns.map((column) => `<td data-label="${escapeHtml(column)}">${formatCell(row[column], column)}</td>`).join('')}${actionCell}</tr>`;
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
  if (['LOA', 'On LOA', 'Pending', 'Requested', 'Waitlist', 'In Progress', 'Draft', 'Not Started', 'Needs acknowledgement'].includes(text)) {
    return `<span class="pill warning">${escapeHtml(text)}</span>`;
  }
  if (['Off Duty', 'Suspended', 'Archived', 'Failed', 'Denied', 'Cancelled', 'Expired', 'Removed', 'Low activity', 'No activity'].includes(text)) {
    return `<span class="pill danger">${escapeHtml(text)}</span>`;
  }
  return escapeHtml(text);
}

function isDateColumn(column) {
  return ['StartDate', 'EndDate', 'JoinDate', 'DateCompleted', 'ExpiryDate', 'ReviewDate', 'ExpiresAt', 'CurrentLoaEnd', 'CheckinDate', 'FollowUpDate', 'DueDate', 'TargetDate', 'NextReviewDate', 'PeriodStart', 'PeriodEnd', 'StartsOn', 'EndsOn', 'LastReview'].includes(column);
}

function isDateTimeColumn(column) {
  return ['UpdatedAt', 'CreatedAt', 'IssuedAt', 'ReviewedAt', 'ReadAt', 'Timestamp', 'LastLogin', 'ChangedAt', 'StartedAt', 'EndedAt', 'LastShift', 'CourseDate', 'RequestedAt', 'DueAt'].includes(column);
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
    if (options.onSuccess) await options.onSuccess(response);
    else await showView(state.activeView);
  };
}

function formValues(form) {
  const data = new FormData(form);
  const values = {};
  data.forEach((value, key) => {
    if (value instanceof File) {
      if (value.size > 0) values[key] = value;
      return;
    }
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

function fileField(name, label) {
  return { html: `<label>${escapeHtml(label)}<input type="file" name="${escapeHtml(name)}"></label>` };
}

function selectField(name, label, options, selected = '') {
  const optionHtml = options.map((option) => {
    const isSelected = option === selected ? ' selected' : '';
    return `<option value="${escapeHtml(option)}"${isSelected}>${escapeHtml(option)}</option>`;
  }).join('');
  return { html: `<label>${escapeHtml(label)}<select name="${escapeHtml(name)}">${optionHtml}</select></label>` };
}

function customSelectField(name, customName, label, options, selected = '') {
  const isCustom = Boolean(selected && !options.includes(selected));
  const selectedValue = isCustom ? 'Custom' : selected;
  const optionHtml = [...options, 'Custom'].map((option) => `<option value="${escapeHtml(option)}"${option === selectedValue ? ' selected' : ''}>${escapeHtml(option)}</option>`).join('');
  return { html: `<label>${escapeHtml(label)}<select name="${escapeHtml(name)}" data-custom-select="${escapeHtml(customName)}">${optionHtml}</select></label><label${isCustom ? '' : ' hidden'} data-custom-field="${escapeHtml(customName)}">Custom ${escapeHtml(label.toLowerCase())}<input name="${escapeHtml(customName)}" value="${escapeHtml(isCustom ? selected : '')}"></label>` };
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

function userCheckboxGroupField(name, label, options, selected = '') {
  const selectedIds = splitTags(selected);
  const checkboxes = options.map((option) => {
    const checked = selectedIds.includes(option.UserID) ? ' checked' : '';
    return `
      <label class="training-check">
        <input type="checkbox" name="${escapeHtml(name)}" value="${escapeHtml(option.UserID)}"${checked}>
        <span>${escapeHtml(option.RobloxUsername)} - ${escapeHtml(option.Rank || option.Role || '')}</span>
      </label>
    `;
  }).join('');
  return { html: `<fieldset class="wide checkbox-group"><legend>${escapeHtml(label)}</legend>${checkboxes}</fieldset>` };
}

async function api(action, data = {}, includeToken = true) {
  if (USE_SUPABASE || localStorage.getItem('mo8_auth_provider') === 'supabase') {
    return supabaseApi(action, data, includeToken);
  }

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

async function supabaseApi(action, data = {}, includeToken = true) {
  try {
    if (!supabaseClient) return { ok: false, error: 'Supabase is not configured.' };

    if (action === 'login') {
      const email = supabaseEmailForLogin(data.username);
      const { data: authData, error } = await supabaseClient.auth.signInWithPassword({
        email,
        password: data.password,
      });
      if (error) return { ok: false, error: error.message };
      const session = authData.session;
      const profileResponse = await supabaseCurrentProfile();
      if (!profileResponse.ok) return profileResponse;
      return {
        ok: true,
        provider: 'supabase',
        token: session.access_token,
        user: profileResponse.user,
        permissions: profileResponse.permissions,
        expiresAt: session.expires_at ? new Date(session.expires_at * 1000).toISOString() : '',
      };
    }

    if (action === 'logout') {
      await supabaseClient.auth.signOut();
      return { ok: true, loggedOut: true };
    }

    if (includeToken) {
      const { data: sessionData } = await supabaseClient.auth.getSession();
      if (!sessionData.session) return { ok: false, error: 'Session expired. Please sign in again.' };
      state.token = sessionData.session.access_token;
      localStorage.setItem('mo8_token', state.token);
    }

    const handlers = {
      me: supabaseCurrentProfile,
      myProfile: supabaseMyProfile,
      listNotifications: supabaseListNotifications,
      markNotificationsRead: supabaseMarkNotificationsRead,
      dashboard: supabaseDashboard,
      saveDashboardWidgets: supabaseSaveDashboardWidgets,
      myActions: supabaseMyActions,
      globalSearch: supabaseGlobalSearch,
      savedViews: supabaseSavedViews,
      saveSavedView: supabaseSaveSavedView,
      operationalCalendar: supabaseOperationalCalendar,
      saveCalendarEvent: supabaseSaveCalendarEvent,
      developmentRecords: supabaseDevelopmentRecords,
      saveProbation: supabaseSaveProbation,
      deleteProbation: (payload) => supabaseDeleteOperationsRecord('probation_records', 'probation_id', payload.ProbationID),
      savePerformanceReview: supabaseSavePerformanceReview,
      deletePerformanceReview: (payload) => supabaseDeleteOperationsRecord('performance_reviews', 'review_id', payload.ReviewID),
      saveRestriction: supabaseSaveRestriction,
      deleteRestriction: (payload) => supabaseDeleteOperationsRecord('officer_restrictions', 'restriction_id', payload.RestrictionID),
      commandReports: supabaseCommandReports,
      listHandovers: supabaseListHandovers,
      saveHandover: supabaseSaveHandover,
      deleteHandover: supabaseDeleteHandover,
      tasks: supabaseTasks,
      supervisorDashboard: supabaseSupervisorDashboard,
      listOfficers: supabaseListOfficers,
      getOfficerProfile: supabaseGetOfficerProfile,
      saveOfficer: supabaseSaveOfficer,
      setOfficerSupervisor: supabaseSetOfficerSupervisor,
      supervisorOptions: supabaseSupervisorOptions,
      deleteOfficer: supabaseDeleteOfficer,
      listTraining: supabaseListTraining,
      listTrainingOptions: supabaseListTrainingOptions,
      saveTrainingOption: supabaseSaveTrainingOption,
      listTrainingCourses: supabaseListTrainingCourses,
      courseTrainers: supabaseCourseTrainers,
      saveTrainingCourse: supabaseSaveTrainingCourse,
      deleteTrainingCourse: supabaseDeleteTrainingCourse,
      requestCourseSeat: supabaseRequestCourseSeat,
      reviewCourseBooking: supabaseReviewCourseBooking,
      saveTraining: supabaseSaveTraining,
      setOfficerTraining: supabaseSetOfficerTraining,
      setDrivingStandard: supabaseSetDrivingStandard,
      setTrainingReviewDate: supabaseSetTrainingReviewDate,
      listDiscipline: supabaseListDiscipline,
      addDiscipline: supabaseSaveDiscipline,
      saveDiscipline: supabaseSaveDiscipline,
      deleteDiscipline: supabaseDeleteDiscipline,
      listLoa: supabaseListLoa,
      requestOwnLoa: supabaseRequestOwnLoa,
      requestTransfer: supabaseRequestTransfer,
      reviewTransfer: supabaseReviewTransfer,
      requestAppeal: supabaseRequestAppeal,
      reviewAppeal: supabaseReviewAppeal,
      requestSupervisorSupport: supabaseRequestSupervisorSupport,
      reviewSupervisorRequest: supabaseReviewSupervisorRequest,
      saveSupervisorCheckin: supabaseSaveSupervisorCheckin,
      saveDevelopmentPlan: supabaseSaveDevelopmentPlan,
      createLoa: supabaseSaveLoa,
      saveLoa: supabaseSaveLoa,
      reviewLoa: supabaseReviewLoa,
      deleteLoa: supabaseDeleteLoa,
      listDocuments: supabaseListDocuments,
      acknowledgeDocument: supabaseAcknowledgeDocument,
      saveDocument: supabaseSaveDocument,
      deleteDocument: supabaseDeleteDocument,
      listAnnouncements: supabaseListAnnouncements,
      saveAnnouncement: supabaseSaveAnnouncement,
      deleteAnnouncement: supabaseDeleteAnnouncement,
      listUsers: supabaseListUsers,
      saveUser: supabaseSaveUser,
      deleteUser: supabaseDeleteUser,
      bulkUpdateOfficers: supabaseBulkUpdateOfficers,
      permissionsConfig: supabasePermissionsConfig,
      setRolePermission: supabaseSetRolePermission,
      setUserPermission: supabaseSetUserPermission,
      resetUserPassword: supabaseResetUserPassword,
      changePassword: supabaseChangePassword,
      auditLog: supabaseAuditLog,
      rankChangeLog: supabaseRankChangeLog,
      shiftStatus: supabaseShiftStatus,
      startShift: supabaseStartShift,
      endShift: supabaseEndShift,
      teamShifts: supabaseTeamShifts,
      requestRetrospectiveShift: supabaseRequestRetrospectiveShift,
      reviewRetrospectiveShift: supabaseReviewRetrospectiveShift,
      saveShift: supabaseSaveShift,
      requestAccount: supabaseRequestAccount,
    };

    if (!handlers[action]) {
      return {
        ok: false,
        error: `${action} has not been migrated to Supabase yet.`,
      };
    }
    return await handlers[action](data);
  } catch (error) {
    return { ok: false, error: error.message || String(error) };
  }
}

function supabaseEmailForLogin(username) {
  const raw = String(username || '').trim();
  if (raw.includes('@')) return raw;
  return `${raw.toLowerCase().replace(/[^a-z0-9._-]/g, '')}@mo8.local`;
}

async function supabaseCurrentProfile() {
  const { data: authData, error: authError } = await supabaseClient.auth.getUser();
  if (authError || !authData.user) return { ok: false, error: authError?.message || 'Not signed in.' };

  const { data: profile, error } = await supabaseClient
    .from('profiles')
    .select('*')
    .eq('auth_user_id', authData.user.id)
    .maybeSingle();
  if (error) return { ok: false, error: error.message };
  if (!profile) return { ok: false, error: 'This Supabase login is not linked to an MDT profile yet.' };

  const permissions = await supabasePermissions(profile);
  return { ok: true, user: supabaseUser(profile), permissions };
}

async function supabasePermissions(profile) {
  const checks = await Promise.all(ALL_PERMISSIONS.map(async (permission) => {
    const { data, error } = await supabaseClient.rpc('has_permission', { permission_name: permission });
    if (error) throw new Error(error.message);
    return data ? permission : '';
  }));
  return checks.filter(Boolean);
}

async function supabaseMyProfile() {
  const me = await supabaseCurrentProfile();
  if (!me.ok) return me;
  const profile = await supabaseProfileByUserId(me.user.UserID);
  const officer = await supabaseOfficerForMember(profile.member_id);
  const [training, matrix, discipline, loa, shifts, rankChanges, notifications, documents, announcements, profiles, officers, probation, performanceReviews, restrictions] = await Promise.all([
    officer ? supabaseRows('training_records', 'officer_id', officer.officer_id) : [],
    officer ? supabaseRows('training_matrix', 'officer_id', officer.officer_id) : [],
    officer ? supabaseRows('disciplinary_actions', 'officer_id', officer.officer_id) : [],
    officer ? supabaseRows('loa_requests', 'officer_id', officer.officer_id) : [],
    officer ? supabaseRows('shift_logs', 'officer_id', officer.officer_id, 'started_at') : [],
    supabaseRows('rank_changes', 'member_id', profile.member_id, 'changed_at'),
    supabaseNotificationRows(profile.member_id),
    supabaseVisibleDocuments(),
    supabaseVisibleAnnouncements(),
    supabaseAll('profiles'),
    supabaseAll('officers'),
    officer ? supabaseOptionalRows('probation_records', 'officer_id', officer.officer_id) : [],
    officer ? supabaseOptionalRows('performance_reviews', 'officer_id', officer.officer_id, 'review_date') : [],
    officer ? supabaseOptionalRows('officer_restrictions', 'officer_id', officer.officer_id) : [],
  ]);
  return {
    ok: true,
    user: me.user,
    officer: officer ? decorateSupabaseOfficer(officer, { loa, shifts, profiles, officers }) : null,
    training: [...training.map(supabaseTrainingRecord), ...matrix.flatMap(supabaseTrainingMatrixRecords)],
    discipline: discipline.map(supabaseDiscipline),
    loa: loa.map(supabaseLoa),
    transfers: [],
    supervisorRequests: [],
    appeals: [],
    checkins: [],
    developmentPlans: [],
    probation: probation.map((row) => ({ ProbationID: row.probation_id, OfficerID: row.officer_id, Stage: row.stage, Status: row.status, Progress: row.progress, StartDate: row.start_date, TargetDate: row.target_date, Requirements: row.requirements, Notes: row.notes })),
    performanceReviews: performanceReviews.map((row) => ({ ReviewID: row.review_id, OfficerID: row.officer_id, ReviewDate: row.review_date, PeriodStart: row.period_start, PeriodEnd: row.period_end, Rating: row.rating, ActivitySummary: row.activity_summary, Strengths: row.strengths, Improvements: row.improvements, Objectives: row.objectives, NextReviewDate: row.next_review_date })),
    restrictions: restrictions.map((row) => ({ RestrictionID: row.restriction_id, OfficerID: row.officer_id, RestrictionType: row.restriction_type, Details: row.details, StartsOn: row.starts_on, EndsOn: row.ends_on, Status: row.status })),
    shifts: shifts.map(supabaseShift),
    shiftStatus: await supabaseShiftStatus(),
    rankChanges: rankChanges.map(supabaseRankChange),
    documents: documents.rows || [],
    announcements: announcements.rows || [],
    notifications,
  };
}

async function supabaseDashboard() {
  const [officers, loa, documents, announcements, actions, widgetPreferences] = await Promise.all([
    supabaseClient.from('officers').select('*'),
    supabaseClient.from('loa_requests').select('*'),
    supabaseVisibleDocuments(),
    supabaseVisibleAnnouncements(),
    supabaseMyActions(),
    state.user?.UserID ? supabaseRows('dashboard_widgets', 'user_id', state.user.UserID) : [],
  ]);
  if (officers.error) return { ok: false, error: officers.error.message };
  if (loa.error) return { ok: false, error: loa.error.message };
  const activeLoa = (loa.data || []).filter((row) => row.status === 'Approved' && isTodayInRange(row.start_date, row.end_date));
  const pendingLoa = (loa.data || []).filter((row) => row.status === 'Pending');
  return {
    ok: true,
    widgets: widgetPreferences.length ? widgetPreferences.filter((row) => row.enabled).map((row) => row.widget_key) : DASHBOARD_WIDGETS.map(([key]) => key),
    counts: {
      activeOfficers: (officers.data || []).filter((row) => row.status === 'Active').length,
      currentlyOnLoa: activeLoa.length,
      loaPending: pendingLoa.length,
      pendingAppeals: 0,
      trainingReviewsDue: 0,
      pendingAcknowledgements: 0,
      upcomingTraining: 0,
    },
    activeLoa: activeLoa.map(supabaseLoa).slice(0, 5),
    pendingLoa: pendingLoa.map(supabaseLoa).slice(0, 5),
    announcements: announcements.rows || [],
    recentDocuments: documents.rows || [],
    trainingReviewsDue: [],
    recentAudit: [],
    pendingAppeals: [],
    unassignedOfficers: (officers.data || []).filter((row) => row.status !== 'Archived' && !row.supervisor_user_id).map(supabaseOfficer).slice(0, 5),
    lowActivity: [],
    documentAcknowledgements: [],
    upcomingTraining: [],
    myActions: actions.rows || [],
  };
}

async function supabaseSaveDashboardWidgets(data) {
  const me = await supabaseCurrentProfile();
  if (!me.ok) return me;
  const widgets = splitTags(data.Widgets || '');
  const rows = DASHBOARD_WIDGETS.map(([key]) => ({
    user_id: me.user.UserID,
    widget_key: key,
    enabled: widgets.includes(key),
    updated_at: new Date().toISOString(),
  }));
  const { error } = await supabaseClient.from('dashboard_widgets').upsert(rows, { onConflict: 'user_id,widget_key' });
  return error ? { ok: false, error: error.message } : { ok: true, widgets };
}

async function supabaseMyActions() {
  const me = await supabaseCurrentProfile();
  if (!me.ok) return me;
  const profile = await supabaseProfileByUserId(me.user.UserID);
  const officer = await supabaseOfficerForMember(profile.member_id);
  const [documents, acknowledgements, courses, bookings, training, reviews, probation, restrictions, handovers, notifications] = await Promise.all([
    supabaseVisibleDocuments(), supabaseAll('document_acknowledgements'), supabaseAll('training_courses'), supabaseAll('course_bookings'),
    officer ? supabaseRows('training_records', 'officer_id', officer.officer_id) : [], officer ? supabaseOptionalRows('performance_reviews', 'officer_id', officer.officer_id) : [],
    officer ? supabaseOptionalRows('probation_records', 'officer_id', officer.officer_id) : [], officer ? supabaseOptionalRows('officer_restrictions', 'officer_id', officer.officer_id) : [],
    can('VIEW_TASKS') ? supabaseOptionalAll('handover_entries') : [], supabaseNotificationRows(profile.member_id),
  ]);
  const now = new Date();
  const inDays = (value) => Math.ceil((new Date(value) - now) / 86400000);
  const rows = [];
  (documents.rows || []).filter((doc) => truthy(doc.RequiresAcknowledgement) && !acknowledgements.some((ack) => ack.document_id === doc.DocumentID && ack.user_id === me.user.UserID)).forEach((doc) => rows.push({ Group: 'Urgent', Priority: 'High', Type: 'Document', Title: doc.Title, Detail: 'Acknowledgement required.', View: 'documents' }));
  const expiringTraining = training.filter((row) => row.expiry_date && inDays(row.expiry_date) <= 60);
  expiringTraining.forEach((row) => rows.push({ Group: inDays(row.expiry_date) < 0 ? 'Urgent' : 'Due Soon', Priority: inDays(row.expiry_date) < 0 ? 'Critical' : 'High', Type: 'Training', Title: `${row.standard} ${inDays(row.expiry_date) < 0 ? 'expired' : 'expires soon'}`, Detail: `Expiry: ${formatDisplayDate(row.expiry_date)}`, DueDate: row.expiry_date, View: 'myProfile' }));
  await Promise.all(expiringTraining.map((row) => {
    const marker = `${row.standard} / ${row.expiry_date}`;
    if (notifications.some((notice) => String(notice.Message || '').includes(marker))) return null;
    return supabaseNotify(profile.member_id, inDays(row.expiry_date) < 0 ? 'Training qualification expired' : 'Training qualification expiring', `${marker}. Please contact a trainer or your supervisor to arrange renewal.`, me.user.UserID);
  }));
  courses.filter((course) => new Date(course.course_date) >= now && bookings.some((booking) => booking.course_id === course.course_id && booking.officer_id === officer?.officer_id && ['Approved', 'Waitlist'].includes(booking.status))).forEach((course) => rows.push({ Group: 'Upcoming', Priority: 'Normal', Type: 'Course', Title: course.title, Detail: `${course.location || 'Location TBC'} / ${formatDisplayDateTime(course.course_date)}`, DueDate: course.course_date, View: 'courses' }));
  reviews.filter((row) => row.next_review_date && inDays(row.next_review_date) <= 30).forEach((row) => rows.push({ Group: 'Due Soon', Priority: 'High', Type: 'Review', Title: 'Performance review due', Detail: `Due ${formatDisplayDate(row.next_review_date)}`, DueDate: row.next_review_date, View: 'myProfile' }));
  probation.filter((row) => row.status === 'Active').forEach((row) => rows.push({ Group: 'Information', Priority: 'Normal', Type: 'Probation', Title: `${row.stage} probation`, Detail: `${row.progress}% complete`, DueDate: row.target_date, View: 'myProfile' }));
  restrictions.filter((row) => row.status === 'Active').forEach((row) => rows.push({ Group: 'Urgent', Priority: 'High', Type: 'Restriction', Title: row.restriction_type, Detail: row.details || 'Active restriction on your officer record.', DueDate: row.ends_on, View: 'myProfile' }));
  handovers.filter((row) => row.status !== 'Resolved' && (row.owner_user_id === me.user.UserID || (row.recipient_user_ids || []).includes(me.user.UserID) || !row.owner_user_id)).forEach((row) => rows.push({ Group: row.priority === 'Critical' ? 'Urgent' : 'Due Soon', Priority: row.priority, Type: 'Handover', Title: row.title, Detail: row.details, DueDate: row.due_at, View: 'handover' }));
  return { ok: true, rows };
}

async function supabaseGlobalSearch(data) {
  const query = String(data.Query || '').trim().toLowerCase();
  if (!query) return { ok: true, rows: [] };
  const [officers, documents, courses, announcements] = await Promise.all([supabaseAll('officers'), supabaseVisibleDocuments(), supabaseAll('training_courses'), supabaseVisibleAnnouncements()]);
  const matches = (values) => values.some((value) => String(value || '').toLowerCase().includes(query));
  return { ok: true, rows: [
    ...officers.filter((row) => matches([row.roblox_username, row.callsign, row.rank, row.tags?.join(' ')])).map((row) => ({ Type: 'Officers', Title: row.roblox_username, Detail: `${row.rank} / ${row.callsign || 'No callsign'}`, Meta: row.status, OfficerID: row.officer_id })),
    ...(documents.rows || []).filter((row) => matches([row.Title, row.Category, row.FileName])).map((row) => ({ Type: 'Documents', Title: row.Title, Detail: row.Category, Meta: row.FileName, View: 'documents' })),
    ...courses.filter((row) => matches([row.title, row.standard, row.location])).map((row) => ({ Type: 'Courses', Title: row.title, Detail: `${row.standard} / ${formatDisplayDateTime(row.course_date)}`, Meta: row.status, View: 'courses' })),
    ...(announcements.rows || []).filter((row) => matches([row.Title, row.Body, row.Audience])).map((row) => ({ Type: 'Notices', Title: row.Title, Detail: row.Body, Meta: row.Audience, View: 'announcements' })),
  ].slice(0, 100) };
}

async function supabaseSavedViews() {
  const me = await supabaseCurrentProfile(); if (!me.ok) return me;
  const rows = await supabaseOptionalRows('saved_views', 'user_id', me.user.UserID, 'created_at');
  return { ok: true, rows: rows.map((row) => ({ ViewID: row.view_id, Name: row.name, Module: row.module, Query: row.query || '', Filters: row.filters })) };
}

async function supabaseSaveSavedView(data) {
  const me = await supabaseCurrentProfile(); if (!me.ok) return me;
  const { error } = await supabaseClient.from('saved_views').upsert({ user_id: me.user.UserID, name: data.Name, module: 'Search', query: data.Query || '', filters: {} }, { onConflict: 'user_id,name' });
  return error ? { ok: false, error: error.message } : { ok: true };
}

async function supabaseOperationalCalendar() {
  const [events, courses, loa, reviews, checkins, officers] = await Promise.all([supabaseOptionalAll('calendar_events'), supabaseAll('training_courses'), supabaseAll('loa_requests'), supabaseOptionalAll('performance_reviews'), supabaseAll('supervisor_checkins'), supabaseAll('officers')]);
  const officerName = (id) => officers.find((row) => row.officer_id === id)?.roblox_username || 'Officer';
  return { ok: true, rows: [
    ...events.map((row) => ({ ID: row.event_id, Type: row.event_type, Title: row.title, Start: row.starts_at, End: row.ends_at, Detail: row.details, Location: row.location, Editable: can('VIEW_TASKS') })),
    ...courses.map((row) => ({ ID: row.course_id, Type: 'Training', Title: row.title, Start: row.course_date, Detail: row.standard, Location: row.location })),
    ...loa.filter((row) => row.status === 'Approved').map((row) => ({ ID: row.request_id, Type: 'LOA', Title: `${officerName(row.officer_id)} on LOA`, Start: row.start_date, End: row.end_date, Detail: row.reason })),
    ...reviews.filter((row) => row.next_review_date).map((row) => ({ ID: row.review_id, Type: 'Review', Title: `${officerName(row.officer_id)} review`, Start: row.next_review_date, Detail: row.rating })),
    ...checkins.filter((row) => row.follow_up_date).map((row) => ({ ID: row.checkin_id, Type: 'Follow-up', Title: `${officerName(row.officer_id)} check-in`, Start: row.follow_up_date, Detail: row.development_goals })),
  ].filter((row) => row.Start) };
}

async function supabaseSaveCalendarEvent(data) {
  const me = await supabaseCurrentProfile(); if (!me.ok) return me;
  const record = { title: data.Title, event_type: data.EventType, starts_at: data.StartsAt, ends_at: data.EndsAt || null, location: data.Location || '', details: data.Details || '', created_by: me.user.UserID, updated_at: new Date().toISOString() };
  const query = data.EventID ? supabaseClient.from('calendar_events').update(record).eq('event_id', data.EventID) : supabaseClient.from('calendar_events').insert(record);
  const { error } = await query; return error ? { ok: false, error: error.message } : { ok: true };
}

async function supabaseDevelopmentRecords() {
  const [officers, profiles, probation, reviews, restrictions] = await Promise.all([supabaseAll('officers'), supabaseAll('profiles'), supabaseOptionalAll('probation_records'), supabaseOptionalAll('performance_reviews'), supabaseOptionalAll('officer_restrictions')]);
  const officer = (id) => officers.find((row) => row.officer_id === id) || {};
  const profile = (id) => profiles.find((row) => row.user_id === id) || {};
  return { ok: true, officers: officers.map(supabaseOfficer),
    probation: probation.map((row) => ({ ProbationID: row.probation_id, OfficerID: row.officer_id, Officer: officer(row.officer_id).roblox_username, Rank: officer(row.officer_id).rank, Stage: row.stage, Status: row.status, StartDate: row.start_date, TargetDate: row.target_date, Progress: row.progress, Requirements: row.requirements, Notes: row.notes, Reviewer: profile(row.reviewer_user_id).roblox_username || '', ReviewerUserID: row.reviewer_user_id })),
    reviews: reviews.map((row) => ({ ReviewID: row.review_id, OfficerID: row.officer_id, Officer: officer(row.officer_id).roblox_username, ReviewDate: row.review_date, PeriodStart: row.period_start, PeriodEnd: row.period_end, Rating: row.rating, ActivitySummary: row.activity_summary, Strengths: row.strengths, Improvements: row.improvements, Objectives: row.objectives, NextReviewDate: row.next_review_date, Reviewer: profile(row.reviewer_user_id).roblox_username || '' })),
    restrictions: restrictions.map((row) => ({ RestrictionID: row.restriction_id, OfficerID: row.officer_id, Officer: officer(row.officer_id).roblox_username, RestrictionType: row.restriction_type, Details: row.details, StartsOn: row.starts_on, EndsOn: row.ends_on, Status: row.status })),
  };
}

async function supabaseSaveProbation(data) {
  const me = await supabaseCurrentProfile(); if (!me.ok) return me;
  const record = { officer_id: data.OfficerID, stage: data.Stage, status: data.Status, start_date: data.StartDate || null, target_date: data.TargetDate || null, progress: Number(data.Progress || 0), requirements: data.Requirements || '', notes: data.Notes || '', reviewer_user_id: me.user.UserID, updated_by: me.user.UserID, updated_at: new Date().toISOString() };
  const query = data.ProbationID ? supabaseClient.from('probation_records').update(record).eq('probation_id', data.ProbationID) : supabaseClient.from('probation_records').insert(record);
  const { error } = await query; if (error) return { ok: false, error: error.message };
  const officer = await supabaseById('officers', 'officer_id', data.OfficerID); await supabaseNotify(officer?.member_id, 'Probation record updated', `${data.Stage}: ${data.Progress || 0}% complete.`, me.user.UserID); return { ok: true };
}

async function supabaseSavePerformanceReview(data) {
  const me = await supabaseCurrentProfile(); if (!me.ok) return me;
  const record = { officer_id: data.OfficerID, review_date: data.ReviewDate, period_start: data.PeriodStart || null, period_end: data.PeriodEnd || null, rating: data.Rating, activity_summary: data.ActivitySummary || '', strengths: data.Strengths || '', improvements: data.Improvements || '', objectives: data.Objectives || '', next_review_date: data.NextReviewDate || null, reviewer_user_id: me.user.UserID, updated_at: new Date().toISOString() };
  const query = data.ReviewID ? supabaseClient.from('performance_reviews').update(record).eq('review_id', data.ReviewID) : supabaseClient.from('performance_reviews').insert(record);
  const { error } = await query; if (error) return { ok: false, error: error.message };
  const officer = await supabaseById('officers', 'officer_id', data.OfficerID); await supabaseNotify(officer?.member_id, 'Performance review recorded', `Rating: ${data.Rating}. ${data.Objectives ? `Objectives: ${data.Objectives}` : ''}`, me.user.UserID); return { ok: true };
}

async function supabaseSaveRestriction(data) {
  const me = await supabaseCurrentProfile(); if (!me.ok) return me;
  const record = { officer_id: data.OfficerID, restriction_type: data.RestrictionType, details: data.Details || '', starts_on: data.StartsOn, ends_on: data.EndsOn || null, status: data.Status, imposed_by: me.user.UserID, updated_at: new Date().toISOString() };
  const query = data.RestrictionID ? supabaseClient.from('officer_restrictions').update(record).eq('restriction_id', data.RestrictionID) : supabaseClient.from('officer_restrictions').insert(record);
  const { error } = await query; if (error) return { ok: false, error: error.message };
  const officer = await supabaseById('officers', 'officer_id', data.OfficerID); await supabaseNotify(officer?.member_id, `Restriction ${data.Status.toLowerCase()}`, `${data.RestrictionType}: ${data.Details || 'See your MDT profile.'}`, me.user.UserID); return { ok: true };
}

async function supabaseDeleteOperationsRecord(table, idColumn, id) {
  if (!id) return { ok: false, error: 'Record ID is required.' };
  const { error } = await supabaseClient.from(table).delete().eq(idColumn, id);
  if (error) return { ok: false, error: error.message };
  await supabaseAudit(state.user?.UserID, 'DELETE_RECORD', table, id);
  return { ok: true };
}

async function supabaseCommandReports() {
  const [officers, loa, training, shifts, reviews, probation, restrictions, handovers, profiles] = await Promise.all([supabaseAll('officers'), supabaseAll('loa_requests'), supabaseAll('training_records'), supabaseAll('shift_logs'), supabaseOptionalAll('performance_reviews'), supabaseOptionalAll('probation_records'), supabaseOptionalAll('officer_restrictions'), supabaseOptionalAll('handover_entries'), supabaseAll('profiles')]);
  const now = new Date(); const cutoff = new Date(now); cutoff.setDate(cutoff.getDate() - 30);
  const name = (id) => officers.find((row) => row.officer_id === id)?.roblox_username || id;
  const expiry = training.filter((row) => row.expiry_date && (new Date(row.expiry_date) - now) / 86400000 <= 60).map((row) => ({ Officer: name(row.officer_id), Standard: row.standard, ExpiryDate: row.expiry_date, DaysRemaining: Math.ceil((new Date(row.expiry_date) - now) / 86400000) }));
  const activity = officers.map((officer) => { const own = shifts.filter((row) => row.officer_id === officer.officer_id && new Date(row.started_at) >= cutoff); const ms = own.reduce((sum, row) => sum + Math.max(0, new Date(row.ended_at || Date.now()) - new Date(row.started_at)), 0); const last = shifts.filter((row) => row.officer_id === officer.officer_id).sort((a, b) => String(b.started_at).localeCompare(String(a.started_at)))[0]; return { Officer: officer.roblox_username, Rank: officer.rank, Shifts: own.length, Hours: (ms / 3600000).toFixed(1), LastShift: last?.started_at || '', Status: officer.status }; });
  const development = officers.map((officer) => { const ownReviews = reviews.filter((row) => row.officer_id === officer.officer_id).sort((a, b) => String(b.review_date).localeCompare(String(a.review_date))); const prob = probation.find((row) => row.officer_id === officer.officer_id && row.status === 'Active'); const restriction = restrictions.find((row) => row.officer_id === officer.officer_id && row.status === 'Active'); return { Officer: officer.roblox_username, Probation: prob ? `${prob.stage} (${prob.progress}%)` : '', LastReview: ownReviews[0]?.review_date || '', NextReview: ownReviews[0]?.next_review_date || '', Restriction: restriction?.restriction_type || '' }; });
  const supervisors = profiles.filter((profile) => officers.some((officer) => officer.supervisor_user_id === profile.user_id)).map((profile) => ({ Supervisor: profile.roblox_username, AssignedOfficers: officers.filter((officer) => officer.supervisor_user_id === profile.user_id).length, OpenActions: handovers.filter((row) => row.owner_user_id === profile.user_id && row.status !== 'Resolved').length }));
  return { ok: true, metrics: { ActiveOfficers: officers.filter((row) => row.status === 'Active').length, OnLoa: loa.filter((row) => row.status === 'Approved' && isTodayInRange(row.start_date, row.end_date)).length, TrainingExpiring: expiry.length, ReviewsDue: reviews.filter((row) => row.next_review_date && new Date(row.next_review_date) <= now).length, ActiveRestrictions: restrictions.filter((row) => row.status === 'Active').length, OpenHandovers: handovers.filter((row) => row.status !== 'Resolved').length }, trainingExpiry: expiry, activity, development, supervisors };
}

async function supabaseListHandovers() {
  const [rows, profiles] = await Promise.all([supabaseOptionalAll('handover_entries'), supabaseAll('profiles')]);
  return { ok: true, rows: rows.sort((a, b) => String(b.created_at).localeCompare(String(a.created_at))).map((row) => ({ HandoverID: row.handover_id, Title: row.title, Category: row.category, Priority: row.priority, Details: row.details, OwnerUserID: row.owner_user_id, Owner: profiles.find((profile) => profile.user_id === row.owner_user_id)?.roblox_username || '', RecipientUserIDs: row.recipient_user_ids || [], Recipients: (row.recipient_user_ids || []).map((id) => profiles.find((profile) => profile.user_id === id)?.roblox_username).filter(Boolean).join(', '), DueAt: row.due_at, Status: row.status, Resolution: row.resolution, CreatedAt: row.created_at, UpdatedAt: row.updated_at })) };
}

async function supabaseSaveHandover(data) {
  const me = await supabaseCurrentProfile(); if (!me.ok) return me;
  const recipients = splitTags(data.RecipientUserIDs || '');
  const record = { title: data.Title, category: data.Category, priority: data.Priority, details: data.Details, due_at: data.DueAt || null, status: data.Status, resolution: data.Resolution || '', owner_user_id: data.OwnerUserID || me.user.UserID, recipient_user_ids: recipients, created_by: me.user.UserID, updated_at: new Date().toISOString() };
  const query = data.HandoverID ? supabaseClient.from('handover_entries').update(record).eq('handover_id', data.HandoverID) : supabaseClient.from('handover_entries').insert(record);
  const { error } = await query;
  if (error) return { ok: false, error: error.message };
  const owner = record.owner_user_id ? await supabaseById('profiles', 'user_id', record.owner_user_id) : null;
  if (owner && owner.user_id !== me.user.UserID) await supabaseNotify(owner.member_id, 'Command handover assigned', `${data.Priority}: ${data.Title}. Due: ${data.DueAt ? formatDisplayDateTime(data.DueAt) : 'No deadline'}.`, me.user.UserID);
  for (const userId of recipients.filter((id) => id !== me.user.UserID && id !== owner?.user_id)) {
    const recipient = await supabaseById('profiles', 'user_id', userId);
    if (recipient) await supabaseNotify(recipient.member_id, 'Handover sent to you', `${data.Priority}: ${data.Title}.`, me.user.UserID);
  }
  return { ok: true };
}

async function supabaseDeleteHandover(data) {
  return supabaseDeleteOperationsRecord('handover_entries', 'handover_id', data.HandoverID);
}

async function supabaseListOfficers() {
  const { data, error } = await supabaseClient.from('officers').select('*').order('rank').order('roblox_username');
  if (error) return { ok: false, error: error.message };
  const [loa, shifts, profiles] = await Promise.all([
    supabaseAll('loa_requests'),
    supabaseAll('shift_logs'),
    supabaseAll('profiles'),
  ]);
  return { ok: true, rows: (data || []).map((row) => decorateSupabaseOfficer(row, { loa, shifts, profiles })) };
}

async function supabaseGetOfficerProfile(data) {
  const { data: officer, error } = await supabaseClient.from('officers').select('*').eq('officer_id', data.OfficerID).maybeSingle();
  if (error) return { ok: false, error: error.message };
  if (!officer) return { ok: false, error: 'Officer not found.' };
  const [training, matrix, discipline, loa, transfers, supervisorRequests, appeals, checkins, plans, shifts, ranks, profiles, officers, probation, performanceReviews, restrictions] = await Promise.all([
    supabaseRows('training_records', 'officer_id', officer.officer_id),
    supabaseRows('training_matrix', 'officer_id', officer.officer_id),
    supabaseRows('disciplinary_actions', 'officer_id', officer.officer_id),
    supabaseRows('loa_requests', 'officer_id', officer.officer_id),
    supabaseRows('transfer_requests', 'officer_id', officer.officer_id),
    supabaseRows('supervisor_requests', 'officer_id', officer.officer_id),
    supabaseRows('appeals', 'officer_id', officer.officer_id),
    supabaseRows('supervisor_checkins', 'officer_id', officer.officer_id, 'checkin_date'),
    supabaseRows('development_plans', 'officer_id', officer.officer_id, 'created_at'),
    supabaseRows('shift_logs', 'officer_id', officer.officer_id, 'started_at'),
    supabaseRows('rank_changes', 'member_id', officer.member_id, 'changed_at'),
    supabaseAll('profiles'),
    supabaseAll('officers'),
    supabaseOptionalRows('probation_records', 'officer_id', officer.officer_id),
    supabaseOptionalRows('performance_reviews', 'officer_id', officer.officer_id, 'review_date'),
    supabaseOptionalRows('officer_restrictions', 'officer_id', officer.officer_id),
  ]);
  const payload = {
    officer: decorateSupabaseOfficer(officer, { loa, shifts, profiles, officers }),
    training: [...training.map(supabaseTrainingRecord), ...matrix.flatMap(supabaseTrainingMatrixRecords)],
    discipline: discipline.map(supabaseDiscipline),
    loa: loa.map(supabaseLoa),
    transfers: transfers.map(supabaseTransfer),
    supervisorRequests: supervisorRequests.map(supabaseSupervisorRequest),
    appeals: appeals.map(supabaseAppeal),
    checkins: checkins.map(supabaseCheckin),
    developmentPlans: plans.map(supabaseDevelopmentPlan),
    shifts: shifts.map(supabaseShift),
    rankChanges: ranks.map(supabaseRankChange),
    probation: probation.map((row) => ({ Stage: row.stage, Status: row.status, Progress: `${row.progress}%`, StartDate: row.start_date, TargetDate: row.target_date, Requirements: row.requirements, Notes: row.notes })),
    performanceReviews: performanceReviews.map((row) => ({ ReviewDate: row.review_date, Rating: row.rating, Strengths: row.strengths, Improvements: row.improvements, Objectives: row.objectives, NextReviewDate: row.next_review_date })),
    restrictions: restrictions.map((row) => ({ RestrictionType: row.restriction_type, Details: row.details, StartsOn: row.starts_on, EndsOn: row.ends_on, Status: row.status })),
  };
  payload.timeline = supabaseTimeline(payload);
  return Object.assign({ ok: true }, payload);
}

async function supabaseSaveOfficer(data) {
  const me = await supabaseCurrentProfile();
  if (!me.ok) return me;
  const existing = data.OfficerID ? await supabaseById('officers', 'officer_id', data.OfficerID) : null;
  const memberId = existing?.member_id || data.MemberID || `MBR_${crypto.randomUUID().replaceAll('-', '')}`;
  const officerId = data.OfficerID || idForSupabase('OFF');
  const record = {
    officer_id: officerId,
    member_id: memberId,
    roblox_username: data.RobloxUsername || '',
    discord_id: data.DiscordID || '',
    callsign: data.Callsign || '',
    rank: data.Rank || 'Police Constable',
    status: data.Status || 'Active',
    join_date: data.JoinDate || null,
    tags: splitTags(data.Tags || ''),
    notes: data.Notes || '',
  };
  const result = data.OfficerID
    ? await supabaseClient.from('officers').update(record).eq('officer_id', data.OfficerID)
    : await supabaseClient.from('officers').insert(record);
  if (result.error) return { ok: false, error: result.error.message };

  await supabaseUpsertProfileForOfficer(record, data.Role || rankToRole(record.rank), me.user.UserID);
  if (existing && existing.rank !== record.rank) {
    await supabaseInsert('rank_changes', {
      member_id: memberId,
      officer_id: officerId,
      roblox_username: record.roblox_username,
      previous_rank: existing.rank || '',
      new_rank: record.rank,
      reason: data.RankChangeReason || '',
      changed_by: me.user.UserID,
    });
    await supabaseNotify(memberId, 'Rank updated', `Your rank changed from ${existing.rank || 'No rank'} to ${record.rank}.`, me.user.UserID);
  }
  await supabaseAudit(me.user.UserID, data.OfficerID ? 'UPDATE_OFFICER' : 'CREATE_OFFICER', 'Officer', officerId, record);
  return { ok: true, OfficerID: officerId };
}

async function supabaseSetOfficerSupervisor(data) {
  const officer = await supabaseById('officers', 'officer_id', data.OfficerID);
  const previousSupervisorId = officer?.supervisor_user_id || '';
  const nextSupervisorId = data.SupervisorUserID || '';
  if (previousSupervisorId === nextSupervisorId) {
    return { ok: true, changed: false, message: 'Supervisor unchanged.' };
  }
  const { error } = await supabaseClient.from('officers').update({ supervisor_user_id: nextSupervisorId || null }).eq('officer_id', data.OfficerID);
  if (error) return { ok: false, error: error.message };
  const supervisor = nextSupervisorId ? await supabaseById('profiles', 'user_id', nextSupervisorId) : null;
  const previousSupervisor = previousSupervisorId ? await supabaseById('profiles', 'user_id', previousSupervisorId) : null;
  if (officer) {
    await supabaseNotify(officer.member_id, 'Supervisor updated', notificationDetails([
      detailLine('Officer', officer.roblox_username),
      detailLine('Supervisor', supervisor ? supervisor.roblox_username : 'Not assigned'),
      detailLine('Updated by', state.user?.RobloxUsername),
    ]), state.user?.UserID);
  }
  if (supervisor && officer) {
    await supabaseNotify(supervisor.member_id, 'New supervisee assigned', notificationDetails([
      detailLine('Officer', officer.roblox_username),
      detailLine('Rank', officer.rank),
      detailLine('Assigned by', state.user?.RobloxUsername),
    ]), state.user?.UserID);
  }
  if (previousSupervisor && previousSupervisor.user_id !== supervisor?.user_id && officer) {
    await supabaseNotify(previousSupervisor.member_id, 'Supervisee reassigned', `${officer.roblox_username} is no longer assigned to you.`, state.user?.UserID);
  }
  return { ok: true };
}

async function supabaseSupervisorOptions() {
  const { data, error } = await supabaseClient.from('profiles').select('*').in('status', ['Active']).order('roblox_username');
  if (error) return { ok: false, error: error.message };
  return { ok: true, rows: (data || []).filter((user) => user.role !== 'Constable').map(supabaseUser) };
}

async function supabaseDeleteOfficer(data) {
  const officer = await supabaseById('officers', 'officer_id', data.OfficerID);
  if (officer) await supabaseClient.from('profiles').delete().eq('member_id', officer.member_id);
  const { error } = await supabaseClient.from('officers').delete().eq('officer_id', data.OfficerID);
  return error ? { ok: false, error: error.message } : { ok: true };
}

async function supabaseListTraining() {
  const [records, matrix] = await Promise.all([
    supabaseClient.from('training_records').select('*').order('updated_at', { ascending: false }),
    supabaseClient.from('training_matrix').select('*').order('updated_at', { ascending: false }),
  ]);
  if (records.error) return { ok: false, error: records.error.message };
  if (matrix.error) return { ok: false, error: matrix.error.message };
  return {
    ok: true,
    rows: [
      ...(records.data || []).map(supabaseTrainingRecord),
      ...(matrix.data || []).flatMap(supabaseTrainingMatrixRecords),
    ],
  };
}

async function supabaseListTrainingOptions() {
  const { data, error } = await supabaseClient.from('training_options').select('*').order('sort_order');
  if (error) return { ok: false, error: error.message };
  return { ok: true, rows: (data || []).map(supabaseTrainingOption) };
}

async function supabaseListTrainingCourses() {
  const [courses, bookings, profiles, officers] = await Promise.all([
    supabaseAll('training_courses'),
    supabaseAll('course_bookings'),
    supabaseAll('profiles'),
    supabaseAll('officers'),
  ]);
  const meProfile = state.user ? await supabaseProfileByUserId(state.user.UserID) : null;
  const myOfficer = meProfile ? (officers || []).find((row) => row.member_id === meProfile.member_id) : null;
  return {
    ok: true,
    rows: (courses || []).filter((row) => row.status !== 'Archived').map((course) => {
      const courseBookings = (bookings || []).filter((booking) => booking.course_id === course.course_id);
      const trainer = (profiles || []).find((profile) => profile.user_id === course.trainer_user_id);
      const coTrainers = (profiles || []).filter((profile) => (course.co_trainer_user_ids || []).includes(profile.user_id));
      const myBooking = myOfficer ? courseBookings.find((booking) => booking.officer_id === myOfficer.officer_id) : null;
      return Object.assign(supabaseCourse(course, trainer, coTrainers), {
        BookedSeats: courseBookings.filter((booking) => ['Approved', 'Requested', 'Completed'].includes(booking.status)).length,
        PendingRequests: courseBookings.filter((booking) => booking.status === 'Requested').length,
        Waitlist: courseBookings.filter((booking) => booking.status === 'Waitlist').length,
        MyBookingStatus: myBooking ? myBooking.status : '',
      });
    }).sort((a, b) => String(a.CourseDate || '').localeCompare(String(b.CourseDate || ''))),
    bookings: (bookings || []).map((booking) => {
      const course = (courses || []).find((row) => row.course_id === booking.course_id) || {};
      const officer = (officers || []).find((row) => row.officer_id === booking.officer_id) || {};
      return supabaseCourseBooking(booking, course, officer);
    }),
  };
}

async function supabaseCourseTrainers() {
  const { data, error } = await supabaseClient.from('profiles').select('*').eq('status', 'Active').order('roblox_username');
  if (error) return { ok: false, error: error.message };
  return { ok: true, rows: (data || []).filter((user) => ['Trainer', 'Sergeant', 'Inspector', 'Chief Inspector', 'Command'].includes(user.role)).map(supabaseUser) };
}

async function supabaseSaveTrainingCourse(data) {
  const record = {
    title: data.Title || '',
    standard: data.Standard || '',
    trainer_user_id: data.TrainerUserID || state.user?.UserID || null,
    co_trainer_user_ids: splitTags(data.CoTrainerUserIDs || '').filter((id) => id !== (data.TrainerUserID || state.user?.UserID || '')),
    course_date: data.CourseDate || null,
    location: data.Location || '',
    capacity: Number(data.Capacity || 0),
    status: data.Status || 'Scheduled',
    notes: data.Notes || '',
  };
  if (!data.CourseID) record.created_by = state.user?.UserID || null;
  const result = data.CourseID
    ? await supabaseClient.from('training_courses').update(record).eq('course_id', data.CourseID).select().single()
    : await supabaseClient.from('training_courses').insert(record).select().single();
  return result.error ? { ok: false, error: result.error.message } : { ok: true, CourseID: result.data.course_id };
}

async function supabaseDeleteTrainingCourse(data) {
  const [course, bookings, officers] = await Promise.all([
    supabaseById('training_courses', 'course_id', data.CourseID),
    supabaseRows('course_bookings', 'course_id', data.CourseID),
    supabaseAll('officers'),
  ]);
  if (!course) return { ok: false, error: 'Course not found.' };

  const affectedBookings = (bookings || []).filter((booking) => booking.outcome !== 'Passed');
  const { error } = await supabaseClient.from('training_courses').delete().eq('course_id', data.CourseID);
  if (error) return { ok: false, error: error.message };

  await Promise.all(affectedBookings.map(async (booking) => {
    const officer = (officers || []).find((row) => row.officer_id === booking.officer_id);
    if (!officer) return null;
    return supabaseNotify(officer.member_id, 'Training course cancelled', notificationDetails([
      detailLine('Course', course.title),
      detailLine('Standard', course.standard),
      detailLine('Date', formatDisplayDateTime(course.course_date)),
      detailLine('Location', course.location),
      detailLine('Status', 'Cancelled'),
    ]), state.user?.UserID);
  }));
  return { ok: true };
}

async function supabaseRequestCourseSeat(data) {
  const profile = await supabaseProfileByUserId(state.user.UserID);
  const officer = await supabaseOfficerForMember(profile.member_id);
  if (!officer) return { ok: false, error: 'No linked officer profile.' };
  const course = await supabaseById('training_courses', 'course_id', data.CourseID);
  const bookings = await supabaseRows('course_bookings', 'course_id', data.CourseID);
  const approved = bookings.filter((booking) => ['Approved', 'Completed'].includes(booking.status)).length;
  const status = Number(course.capacity || 0) && approved >= Number(course.capacity || 0) ? 'Waitlist' : 'Requested';
  const bookingId = idForSupabase('CBK');
  const result = await supabaseClient.from('course_bookings').insert({
    booking_id: bookingId,
    course_id: data.CourseID,
    officer_id: officer.officer_id,
    status,
    notes: data.Notes || '',
  });
  if (result.error) return { ok: false, error: result.error.message };
  await supabaseNotify(officer.member_id, 'Training course requested', notificationDetails([
    detailLine('Course', course.title),
    detailLine('Standard', course.standard),
    detailLine('Date', formatDisplayDateTime(course.course_date)),
    detailLine('Location', course.location),
    detailLine('Status', status),
    detailLine('Decision', status === 'Requested' ? 'Awaiting trainer review' : 'Added to waitlist'),
  ]), state.user?.UserID);
  if (course.trainer_user_id) {
    const trainer = await supabaseById('profiles', 'user_id', course.trainer_user_id);
    if (trainer) await supabaseNotify(trainer.member_id, 'Course booking requested', notificationDetails([
      detailLine('Officer', officer.roblox_username),
      detailLine('Rank', officer.rank),
      detailLine('Course', course.title),
      detailLine('Date', formatDisplayDateTime(course.course_date)),
      detailLine('Status', status),
    ]), state.user?.UserID);
  }
  await Promise.all((course.co_trainer_user_ids || []).map(async (userId) => {
    const coTrainer = await supabaseById('profiles', 'user_id', userId);
    if (!coTrainer) return null;
    return supabaseNotify(coTrainer.member_id, 'Course booking requested', notificationDetails([
      detailLine('Officer', officer.roblox_username),
      detailLine('Rank', officer.rank),
      detailLine('Course', course.title),
      detailLine('Date', formatDisplayDateTime(course.course_date)),
      detailLine('Status', status),
    ]), state.user?.UserID);
  }));
  return { ok: true, BookingID: bookingId, Status: status };
}

async function supabaseReviewCourseBooking(data) {
  const existing = await supabaseById('course_bookings', 'booking_id', data.BookingID);
  const course = existing ? await supabaseById('training_courses', 'course_id', existing.course_id) : null;
  const officer = existing ? await supabaseById('officers', 'officer_id', existing.officer_id) : null;
  const update = {
    status: data.Status || existing.status || 'Approved',
    outcome: data.Outcome || existing.outcome || '',
    notes: data.Notes || existing.notes || '',
    reviewed_by: state.user?.UserID || null,
    reviewed_at: new Date().toISOString(),
  };
  const { error } = await supabaseClient.from('course_bookings').update(update).eq('booking_id', data.BookingID);
  if (error) return { ok: false, error: error.message };
  if (update.outcome === 'Passed') {
    await supabaseSaveTraining({ OfficerID: existing.officer_id, Standard: course.standard, Status: 'Passed', Assessor: state.user?.RobloxUsername || '', Notes: `Passed via course ${course.title}` });
  }
  if (officer && course) {
    await supabaseNotify(officer.member_id, 'Training course updated', notificationDetails([
      detailLine('Course', course.title),
      detailLine('Standard', course.standard),
      detailLine('Date', formatDisplayDateTime(course.course_date)),
      detailLine('Status', update.status),
      detailLine('Outcome', update.outcome),
      detailLine('Reviewed by', state.user?.RobloxUsername),
      detailLine('Notes', update.notes),
    ]), state.user?.UserID);
  }
  return { ok: true, BookingID: data.BookingID };
}

async function supabaseSaveTrainingOption(data) {
  const record = {
    name: data.Name || '',
    type: data.Type || 'Specialist',
    status: data.Status || 'Active',
    sort_order: Number(data.SortOrder || 0),
    updated_by: state.user?.UserID || null,
  };
  const result = data.OptionID
    ? await supabaseClient.from('training_options').update(record).eq('option_id', data.OptionID).select().single()
    : await supabaseClient.from('training_options').insert(record).select().single();
  return result.error ? { ok: false, error: result.error.message } : { ok: true, OptionID: result.data.option_id };
}

async function supabaseSaveTraining(data) {
  const record = {
    officer_id: data.OfficerID,
    standard: data.Standard || '',
    status: data.Status || 'Not Started',
    assessor: data.Assessor || state.user?.RobloxUsername || '',
    date_completed: data.DateCompleted || null,
    expiry_date: data.ExpiryDate || null,
    notes: data.Notes || '',
  };
  const result = data.TrainingID
    ? await supabaseClient.from('training_records').update(record).eq('training_id', data.TrainingID).select().single()
    : await supabaseClient.from('training_records').insert(record).select().single();
  if (result.error) return { ok: false, error: result.error.message };
  const officer = await supabaseById('officers', 'officer_id', record.officer_id);
  if (officer) await supabaseNotify(officer.member_id, 'Training record updated', notificationDetails([
    detailLine('Officer', officer.roblox_username),
    detailLine('Training', record.standard),
    detailLine('Status', record.status),
    detailLine('Assessor', record.assessor),
    detailLine('Date completed', formatDisplayDate(record.date_completed)),
  ]), state.user?.UserID);
  return { ok: true, TrainingID: result.data.training_id };
}

async function supabaseSetOfficerTraining(data) {
  const officer = await supabaseById('officers', 'officer_id', data.OfficerID);
  if (!officer) return { ok: false, error: 'Officer not found.' };
  const fieldMap = { Taser: 'taser', MOE: 'moe', 'Blue Ticket': 'blue_ticket', Motorbike: 'motorbike' };
  const fieldName = fieldMap[data.Standard];
  if (!fieldName) {
    const existing = await supabaseClient.from('training_records').select('training_id').eq('officer_id', officer.officer_id).eq('standard', data.Standard).eq('status', 'Passed');
    if (existing.error) return { ok: false, error: existing.error.message };
    const enabled = truthy(data.Enabled);
    if (enabled && !existing.data?.length) {
      const result = await supabaseClient.from('training_records').insert({ officer_id: officer.officer_id, standard: data.Standard, status: 'Passed', assessor: state.user?.RobloxUsername || '', date_completed: new Date().toISOString().slice(0, 10) });
      if (result.error) return { ok: false, error: result.error.message };
    }
    if (!enabled && existing.data?.length) {
      const result = await supabaseClient.from('training_records').delete().in('training_id', existing.data.map((row) => row.training_id));
      if (result.error) return { ok: false, error: result.error.message };
    }
    await supabaseNotify(officer.member_id, 'Training updated', notificationDetails([detailLine('Training', data.Standard), detailLine('Status', enabled ? 'Added' : 'Removed'), detailLine('Updated by', state.user?.RobloxUsername)]), state.user?.UserID);
    return { ok: true };
  }
  const row = {
    officer_id: officer.officer_id,
    member_id: officer.member_id,
    roblox_username: officer.roblox_username,
    [fieldName]: truthy(data.Enabled),
    updated_by: state.user?.RobloxUsername || '',
  };
  const { error } = await supabaseClient.from('training_matrix').upsert(row, { onConflict: 'officer_id' });
  if (error) return { ok: false, error: error.message };
  await supabaseNotify(officer.member_id, 'Training updated', notificationDetails([
    detailLine('Training', data.Standard),
    detailLine('Status', truthy(data.Enabled) ? 'Added' : 'Removed'),
    detailLine('Updated by', state.user?.RobloxUsername),
  ]), state.user?.UserID);
  return { ok: true };
}

async function supabaseSetDrivingStandard(data) {
  const officer = await supabaseById('officers', 'officer_id', data.OfficerID);
  if (!officer) return { ok: false, error: 'Officer not found.' };
  const { error } = await supabaseClient.from('training_matrix').upsert({
    officer_id: officer.officer_id,
    member_id: officer.member_id,
    roblox_username: officer.roblox_username,
    driving_standard: data.Standard || '',
    updated_by: state.user?.RobloxUsername || '',
  }, { onConflict: 'officer_id' });
  if (error) return { ok: false, error: error.message };
  await supabaseNotify(officer.member_id, 'Driving standard updated', notificationDetails([
    detailLine('Driving standard', data.Standard || 'Not set'),
    detailLine('Updated by', state.user?.RobloxUsername),
  ]), state.user?.UserID);
  return { ok: true };
}

async function supabaseSetTrainingReviewDate(data) {
  const officer = await supabaseById('officers', 'officer_id', data.OfficerID);
  if (!officer) return { ok: false, error: 'Officer not found.' };
  const { error } = await supabaseClient.from('training_matrix').upsert({
    officer_id: officer.officer_id,
    member_id: officer.member_id,
    roblox_username: officer.roblox_username,
    review_date: data.ReviewDate || null,
    updated_by: state.user?.RobloxUsername || '',
  }, { onConflict: 'officer_id' });
  return error ? { ok: false, error: error.message } : { ok: true };
}

async function supabaseListDiscipline() {
  const [rows, officers] = await Promise.all([supabaseAll('disciplinary_actions'), supabaseAll('officers')]);
  return { ok: true, rows: (rows || []).map((row) => supabaseDiscipline(row, (officers || []).find((officer) => officer.officer_id === row.officer_id))) };
}

async function supabaseSaveDiscipline(data) {
  const record = {
    officer_id: data.OfficerID,
    type: data.Type || 'Note',
    summary: data.Summary || '',
    details: data.Details || '',
    issued_by: state.user?.UserID || null,
    status: data.Status || 'Active',
  };
  const result = data.ActionID
    ? await supabaseClient.from('disciplinary_actions').update(record).eq('action_id', data.ActionID).select().single()
    : await supabaseClient.from('disciplinary_actions').insert(record).select().single();
  if (result.error) return { ok: false, error: result.error.message };
  const officer = await supabaseById('officers', 'officer_id', record.officer_id);
  if (officer) await supabaseNotify(officer.member_id, 'Discipline record updated', notificationDetails([
    detailLine('Officer', officer.roblox_username),
    detailLine('Type', record.type),
    detailLine('Summary', record.summary),
    detailLine('Status', record.status),
    detailLine('Issued by', state.user?.RobloxUsername),
  ]), state.user?.UserID);
  return { ok: true, ActionID: result.data.action_id };
}

async function supabaseDeleteDiscipline(data) {
  const { error } = await supabaseClient.from('disciplinary_actions').delete().eq('action_id', data.ActionID);
  return error ? { ok: false, error: error.message } : { ok: true };
}

async function supabaseListLoa() {
  const [rows, officers, profiles] = await Promise.all([supabaseAll('loa_requests'), supabaseAll('officers'), supabaseAll('profiles')]);
  return { ok: true, rows: (rows || []).map((row) => supabaseLoa(row, (officers || []).find((officer) => officer.officer_id === row.officer_id), profiles || [])) };
}

async function supabaseRequestOwnLoa(data) {
  const profile = await supabaseProfileByUserId(state.user.UserID);
  const officer = await supabaseOfficerForMember(profile.member_id);
  if (!officer) return { ok: false, error: 'No linked officer profile.' };
  const response = await supabaseSaveLoa(Object.assign({}, data, { OfficerID: officer.officer_id, Status: 'Pending' }));
  if (!response.ok) return response;
  const message = notificationDetails([
    detailLine('Officer', officer.roblox_username),
    detailLine('Rank', officer.rank),
    detailLine('Start date', formatDisplayDate(data.StartDate)),
    detailLine('End date', formatDisplayDate(data.EndDate)),
    detailLine('Status', 'Pending'),
    detailLine('Reason', data.Reason),
  ]);
  await supabaseNotify(officer.member_id, 'LOA request submitted', message, state.user?.UserID);
  if (officer.supervisor_user_id) {
    const supervisor = await supabaseById('profiles', 'user_id', officer.supervisor_user_id);
    if (supervisor) await supabaseNotify(supervisor.member_id, 'Supervisee LOA request submitted', message, state.user?.UserID);
  }
  return response;
}

async function supabaseSaveLoa(data) {
  const record = {
    officer_id: data.OfficerID,
    start_date: data.StartDate || null,
    end_date: data.EndDate || null,
    reason: data.Reason || '',
    status: data.Status || 'Pending',
    review_reason: data.ReviewReason || '',
  };
  const result = data.RequestID
    ? await supabaseClient.from('loa_requests').update(record).eq('request_id', data.RequestID).select().single()
    : await supabaseClient.from('loa_requests').insert(record).select().single();
  return result.error ? { ok: false, error: result.error.message } : { ok: true, RequestID: result.data.request_id };
}

async function supabaseReviewLoa(data) {
  const update = {
    status: data.Status || 'Approved',
    review_reason: data.ReviewReason || '',
    reviewed_by: state.user?.UserID || null,
    reviewed_at: new Date().toISOString(),
  };
  const { error } = await supabaseClient.from('loa_requests').update(update).eq('request_id', data.RequestID);
  if (error) return { ok: false, error: error.message };
  const loa = await supabaseById('loa_requests', 'request_id', data.RequestID);
  const officer = loa ? await supabaseById('officers', 'officer_id', loa.officer_id) : null;
  if (officer) await supabaseNotify(officer.member_id, `LOA ${update.status}`, notificationDetails([
    detailLine('Officer', officer.roblox_username),
    detailLine('Start date', formatDisplayDate(loa.start_date)),
    detailLine('End date', formatDisplayDate(loa.end_date)),
    detailLine('Status', update.status),
    detailLine('Reviewed by', state.user?.RobloxUsername),
    detailLine('Decision reason', update.review_reason),
  ]), state.user?.UserID);
  return { ok: true };
}

async function supabaseDeleteLoa(data) {
  const { error } = await supabaseClient.from('loa_requests').delete().eq('request_id', data.RequestID);
  return error ? { ok: false, error: error.message } : { ok: true };
}

async function supabaseRequestTransfer(data) {
  const profile = await supabaseProfileByUserId(state.user.UserID);
  const officer = await supabaseOfficerForMember(profile.member_id);
  if (!officer) return { ok: false, error: 'No linked officer profile.' };
  const result = await supabaseClient.from('transfer_requests').insert({
    officer_id: officer.officer_id,
    current_division: 'MO8',
    target_division: data.TargetDivision || '',
    time_in_mo8: data.TimeInMO8 || '',
    reason: data.Reason || '',
    has_permission: data.HasPermission || '',
    notes: data.Notes || '',
    status: 'Pending',
  }).select().single();
  if (result.error) return { ok: false, error: result.error.message };
  await supabaseNotify(officer.member_id, 'Transfer request submitted', notificationDetails([
    detailLine('Officer', officer.roblox_username),
    detailLine('Target division', data.TargetDivision || 'Another division'),
    detailLine('Time in MO8', data.TimeInMO8),
    detailLine('Status', 'Pending'),
    detailLine('Reason', data.Reason),
  ]), state.user?.UserID);
  if (officer.supervisor_user_id) {
    const supervisor = await supabaseById('profiles', 'user_id', officer.supervisor_user_id);
    if (supervisor) await supabaseNotify(supervisor.member_id, 'Supervisee transfer request', notificationDetails([
      detailLine('Officer', officer.roblox_username),
      detailLine('Rank', officer.rank),
      detailLine('Target division', data.TargetDivision),
      detailLine('Status', 'Pending'),
    ]), state.user?.UserID);
  }
  return { ok: true, RequestID: result.data.request_id };
}

async function supabaseReviewTransfer(data) {
  const { error } = await supabaseClient.from('transfer_requests').update({
    status: data.Status || 'Approved',
    review_reason: data.ReviewReason || '',
    reviewed_by: state.user?.UserID || null,
    reviewed_at: new Date().toISOString(),
  }).eq('request_id', data.RequestID);
  if (error) return { ok: false, error: error.message };
  const transfer = await supabaseById('transfer_requests', 'request_id', data.RequestID);
  const officer = transfer ? await supabaseById('officers', 'officer_id', transfer.officer_id) : null;
  if (officer) await supabaseNotify(officer.member_id, 'Transfer request updated', notificationDetails([
    detailLine('Officer', officer.roblox_username),
    detailLine('Target division', transfer.target_division),
    detailLine('Status', data.Status || 'Reviewed'),
    detailLine('Reviewed by', state.user?.RobloxUsername),
    detailLine('Decision reason', data.ReviewReason),
  ]), state.user?.UserID);
  return { ok: true };
}

async function supabaseRequestAppeal(data) {
  const profile = await supabaseProfileByUserId(state.user.UserID);
  const officer = await supabaseOfficerForMember(profile.member_id);
  if (!officer) return { ok: false, error: 'No linked officer profile.' };
  const result = await supabaseClient.from('appeals').insert({
    officer_id: officer.officer_id,
    source_type: data.SourceType || '',
    source_id: data.SourceID || '',
    reason: data.Reason || '',
    status: 'Pending',
  }).select().single();
  return result.error ? { ok: false, error: result.error.message } : { ok: true, AppealID: result.data.appeal_id };
}

async function supabaseReviewAppeal(data) {
  const { error } = await supabaseClient.from('appeals').update({
    status: data.Status || 'Completed',
    review_reason: data.ReviewReason || '',
    reviewed_by: state.user?.UserID || null,
    reviewed_at: new Date().toISOString(),
  }).eq('appeal_id', data.AppealID);
  return error ? { ok: false, error: error.message } : { ok: true };
}

async function supabaseTasks() {
  const [officers, loa, transfers, supervisorRequests, appeals, profiles, courses, courseBookings, probation, performance, restrictions, accountRequestRows, retrospectiveRows] = await Promise.all([
    supabaseAll('officers'),
    supabaseAll('loa_requests'),
    supabaseAll('transfer_requests'),
    supabaseAll('supervisor_requests'),
    supabaseAll('appeals'),
    supabaseAll('profiles'),
    supabaseAll('training_courses'),
    supabaseAll('course_bookings'),
    supabaseOptionalAll('probation_records'),
    supabaseOptionalAll('performance_reviews'),
    supabaseOptionalAll('officer_restrictions'),
    can('REVIEW_ACCOUNT_REQUESTS') ? supabaseOptionalAll('account_requests') : [],
    supabaseOptionalAll('retrospective_shift_requests'),
  ]);
  const decorate = (row, type) => {
    const officer = (officers || []).find((item) => item.officer_id === row.officer_id) || {};
    const supervisor = (profiles || []).find((profile) => profile.user_id === officer.supervisor_user_id) || {};
    return Object.assign({}, row, {
      TaskType: type,
      OfficerID: officer.officer_id || row.officer_id,
      Officer: officer.roblox_username || row.officer_id,
      Rank: officer.rank || '',
      Supervisor: supervisor.roblox_username || 'Not assigned',
      SupervisorUserID: supervisor.user_id || '',
      MySupervisee: supervisor.user_id === state.user?.UserID,
      RequestID: row.request_id,
      AppealID: row.appeal_id,
      Subject: row.subject || '',
      SourceType: row.source_type || '',
      SourceID: row.source_id || '',
      StartDate: row.start_date || '',
      EndDate: row.end_date || '',
      TargetDivision: row.target_division || '',
      TimeInMO8: row.time_in_mo8 || '',
      Reason: row.reason || '',
      Notes: row.notes || '',
      Category: row.category || '',
      Details: row.details || '',
      ReviewReason: row.review_reason || '',
      Status: row.status || '',
    });
  };
  const pendingLoa = (loa || []).filter((row) => row.status === 'Pending').map((row) => decorate(row, 'LOA Approval'));
  const pendingTransfers = (transfers || []).filter((row) => row.status === 'Pending').map((row) => decorate(row, 'Transfer Request'));
  const pendingSupervisorRequests = (supervisorRequests || []).filter((row) => row.status === 'Pending').map((row) => decorate(row, 'Supervisor Request'));
  const pendingAppeals = (appeals || []).filter((row) => row.status === 'Pending').map((row) => decorate(row, 'Appeal / Review'));
  const pendingCourseBookings = (courseBookings || []).filter((row) => row.status === 'Requested').map((booking) => {
    const course = (courses || []).find((row) => row.course_id === booking.course_id) || {};
    const officer = (officers || []).find((row) => row.officer_id === booking.officer_id) || {};
    const trainerIds = [course.trainer_user_id, ...(course.co_trainer_user_ids || [])].filter(Boolean);
    const assignedToMe = trainerIds.includes(state.user?.UserID);
    if (!assignedToMe && !can('MANAGE_COURSES')) return null;
    return {
      TaskType: 'Course Booking',
      BookingID: booking.booking_id,
      CourseID: booking.course_id,
      Course: course.title || booking.course_id,
      OfficerID: officer.officer_id || booking.officer_id,
      Officer: officer.roblox_username || booking.officer_id,
      Rank: officer.rank || '',
      Supervisor: '',
      Subject: course.standard || '',
      SourceType: 'Training Course',
      Reason: booking.notes || '',
      Status: booking.status || '',
      MySupervisee: assignedToMe,
      RequestedAt: booking.requested_at || '',
    };
  }).filter(Boolean);
  const assignedOfficers = (officers || []).filter((officer) => officer.supervisor_user_id === state.user?.UserID && officer.status !== 'Archived');
  const supervisorName = state.user?.RobloxUsername || '';
  const probationReviews = (probation || []).filter((row) => row.status === 'Active' && Number(row.progress || 0) < 100 && assignedOfficers.some((officer) => officer.officer_id === row.officer_id)).map((row) => {
    const officer = assignedOfficers.find((item) => item.officer_id === row.officer_id) || {};
    return { TaskType: 'Probation Review', ProbationID: row.probation_id, OfficerID: row.officer_id, Officer: officer.roblox_username, Rank: officer.rank, Supervisor: supervisorName, Subject: `${row.stage} / ${row.progress}% complete`, Reason: row.requirements || row.notes || 'Probation record requires completion.', Status: row.status, StartDate: row.start_date, EndDate: row.target_date, Stage: row.stage, Progress: row.progress, TargetDate: row.target_date, Requirements: row.requirements, Notes: row.notes, MySupervisee: true };
  });
  const performanceReviews = assignedOfficers.map((officer) => {
    const latest = (performance || []).filter((row) => row.officer_id === officer.officer_id).sort((a, b) => String(b.review_date).localeCompare(String(a.review_date)))[0];
    if (latest?.next_review_date && new Date(latest.next_review_date) > new Date()) return null;
    return { TaskType: 'Performance Review', ReviewID: latest?.review_id || '', OfficerID: officer.officer_id, Officer: officer.roblox_username, Rank: officer.rank, Supervisor: supervisorName, Subject: latest ? 'Scheduled performance review due' : 'No performance review recorded', Reason: latest?.objectives || 'A performance review should be completed.', Status: latest ? 'Due' : 'Not Started', ReviewDate: latest?.review_date || '', PeriodStart: latest?.period_start || '', PeriodEnd: latest?.period_end || '', Rating: latest?.rating || 'Meets Expectations', ActivitySummary: latest?.activity_summary || '', Strengths: latest?.strengths || '', Improvements: latest?.improvements || '', Objectives: latest?.objectives || '', NextReviewDate: latest?.next_review_date || '', MySupervisee: true };
  }).filter(Boolean);
  const restrictionReviews = (restrictions || []).filter((row) => row.status === 'Active' && assignedOfficers.some((officer) => officer.officer_id === row.officer_id)).map((row) => {
    const officer = assignedOfficers.find((item) => item.officer_id === row.officer_id) || {};
    return { TaskType: 'Restriction Review', RestrictionID: row.restriction_id, OfficerID: row.officer_id, Officer: officer.roblox_username, Rank: officer.rank, Supervisor: supervisorName, Subject: row.restriction_type, Reason: row.details || 'Active restriction requires oversight.', Status: row.status, StartDate: row.starts_on, EndDate: row.ends_on, RestrictionType: row.restriction_type, Details: row.details, StartsOn: row.starts_on, EndsOn: row.ends_on, MySupervisee: true };
  });
  const accountRequests = (accountRequestRows || []).filter((row) => row.status === 'Pending').map((row) => ({ TaskType: 'Account Request', RequestID: row.request_id, Officer: row.roblox_username, RobloxUsername: row.roblox_username, Rank: row.rank, DiscordID: row.discord_id, Subject: 'New MDT account', Reason: 'Identity and team membership require verification.', ReviewNotes: row.review_notes || '', Status: row.status, CreatedAt: row.created_at }));
  const retrospectiveShifts = (retrospectiveRows || []).filter((row) => row.status === 'Pending').map((row) => {
    const task = decorate(row, 'Retrospective Shift');
    return Object.assign(task, { StartedAt: row.started_at, EndedAt: row.ended_at, Summary: row.summary || '', Reason: row.reason || '' });
  }).filter((row) => row.MySupervisee || can('FULL_ACCESS'));
  const all = [...pendingLoa, ...pendingTransfers, ...pendingSupervisorRequests, ...pendingCourseBookings, ...pendingAppeals, ...probationReviews, ...performanceReviews, ...restrictionReviews, ...accountRequests, ...retrospectiveShifts];
  return {
    ok: true,
    pendingLoa,
    pendingTransfers,
    pendingSupervisorRequests,
    pendingCourseBookings,
    pendingAppeals,
    probationReviews,
    performanceReviews,
    restrictionReviews,
    accountRequests,
    retrospectiveShifts,
    counts: {
      pendingLoa: pendingLoa.length,
      pendingTransfers: pendingTransfers.length,
      pendingSupervisorRequests: pendingSupervisorRequests.length,
      pendingCourseBookings: pendingCourseBookings.length,
      pendingAppeals: pendingAppeals.length,
      developmentReviews: probationReviews.length + performanceReviews.length + restrictionReviews.length,
      accountRequests: accountRequests.length,
      retrospectiveShifts: retrospectiveShifts.length,
      mySuperviseeTasks: all.filter((row) => row.MySupervisee).length,
      total: all.length,
    },
  };
}

async function supabaseSupervisorDashboard() {
  const [officers, shifts, loa, discipline, plans, checkins, supervisorRequests, profiles] = await Promise.all([
    supabaseAll('officers'),
    supabaseAll('shift_logs'),
    supabaseAll('loa_requests'),
    supabaseAll('disciplinary_actions'),
    supabaseAll('development_plans'),
    supabaseAll('supervisor_checkins'),
    supabaseAll('supervisor_requests'),
    supabaseAll('profiles'),
  ]);
  const assigned = (officers || []).filter((row) => row.supervisor_user_id === state.user?.UserID).map((officer) => {
    const officerShifts = (shifts || []).filter((row) => row.officer_id === officer.officer_id);
    return Object.assign(decorateSupabaseOfficer(officer, { loa, shifts: officerShifts, profiles }), {
      LastShift: officerShifts.sort((a, b) => String(b.started_at).localeCompare(String(a.started_at)))[0]?.started_at || '',
      MonthlyActivity: durationText(officerShifts.reduce((sum, shift) => sum + shiftMs(shift), 0)),
      TrainingGaps: '',
      DisciplineFlags: (discipline || []).filter((row) => row.officer_id === officer.officer_id && row.status === 'Active').length,
      OpenPlans: (plans || []).filter((row) => row.officer_id === officer.officer_id && row.status !== 'Completed').length,
    });
  });
  const unassigned = (officers || []).filter((row) => row.status !== 'Archived' && !row.supervisor_user_id).map((row) => decorateSupabaseOfficer(row, { loa, shifts, profiles }));
  const pendingRequests = (supervisorRequests || []).filter((row) => row.status === 'Pending').map((row) => supabaseSupervisorRequest(row, (officers || []).find((officer) => officer.officer_id === row.officer_id), profiles || []));
  const workload = (profiles || []).filter((profile) => profile.role !== 'Constable').map((profile) => ({
    Supervisor: profile.roblox_username,
    Rank: profile.rank,
    AssignedOfficers: (officers || []).filter((officer) => officer.supervisor_user_id === profile.user_id).length,
  }));
  return {
    ok: true,
    assigned,
    unassigned,
    pendingRequests,
    workload,
    plans: (plans || []).map((row) => supabaseDevelopmentPlan(row, (officers || []).find((officer) => officer.officer_id === row.officer_id), profiles || [])),
    checkins: (checkins || []).map((row) => supabaseCheckin(row, (officers || []).find((officer) => officer.officer_id === row.officer_id), profiles || [])),
    counts: {
      assigned: assigned.length,
      unassigned: unassigned.length,
      pendingRequests: pendingRequests.length,
      openPlans: (plans || []).filter((row) => row.status !== 'Completed').length,
    },
  };
}

async function supabaseRequestSupervisorSupport(data) {
  const profile = await supabaseProfileByUserId(state.user.UserID);
  const officer = await supabaseOfficerForMember(profile.member_id);
  if (!officer) return { ok: false, error: 'No linked officer profile.' };
  const result = await supabaseClient.from('supervisor_requests').insert({
    officer_id: officer.officer_id,
    supervisor_user_id: officer.supervisor_user_id || null,
    category: data.Category || 'General',
    subject: data.Subject || '',
    details: data.Details || '',
    status: 'Pending',
  }).select().single();
  if (result.error) return { ok: false, error: result.error.message };
  await supabaseNotify(officer.member_id, 'Supervisor request submitted', notificationDetails([
    detailLine('Subject', data.Subject || 'Request'),
    detailLine('Category', data.Category || 'General'),
    detailLine('Status', 'Pending'),
    detailLine('Details', data.Details),
  ]), state.user?.UserID);
  if (officer.supervisor_user_id) {
    const supervisor = await supabaseById('profiles', 'user_id', officer.supervisor_user_id);
    if (supervisor) await supabaseNotify(supervisor.member_id, 'Supervisee request submitted', notificationDetails([
      detailLine('Officer', officer.roblox_username),
      detailLine('Rank', officer.rank),
      detailLine('Subject', data.Subject || 'Request'),
      detailLine('Category', data.Category || 'General'),
      detailLine('Status', 'Pending'),
    ]), state.user?.UserID);
  }
  return { ok: true, RequestID: result.data.request_id };
}

async function supabaseReviewSupervisorRequest(data) {
  const { error } = await supabaseClient.from('supervisor_requests').update({
    status: data.Status || 'Completed',
    review_reason: data.ReviewReason || '',
    reviewed_by: state.user?.UserID || null,
    reviewed_at: new Date().toISOString(),
  }).eq('request_id', data.RequestID);
  if (error) return { ok: false, error: error.message };
  const request = await supabaseById('supervisor_requests', 'request_id', data.RequestID);
  const officer = request ? await supabaseById('officers', 'officer_id', request.officer_id) : null;
  if (officer) await supabaseNotify(officer.member_id, 'Supervisor request updated', notificationDetails([
    detailLine('Subject', request.subject || 'Request'),
    detailLine('Category', request.category || 'General'),
    detailLine('Status', data.Status || 'Updated'),
    detailLine('Reviewed by', state.user?.RobloxUsername),
    detailLine('Decision reason', data.ReviewReason),
  ]), state.user?.UserID);
  return { ok: true };
}

async function supabaseSaveSupervisorCheckin(data) {
  const result = await supabaseClient.from('supervisor_checkins').insert({
    officer_id: data.OfficerID,
    supervisor_user_id: state.user?.UserID || null,
    checkin_date: data.CheckinDate || new Date().toISOString().slice(0, 10),
    summary: data.Summary || '',
    concerns: data.Concerns || '',
    development_goals: data.DevelopmentGoals || '',
    follow_up_date: data.FollowUpDate || null,
    created_by: state.user?.UserID || null,
  }).select().single();
  if (result.error) return { ok: false, error: result.error.message };
  const officer = await supabaseById('officers', 'officer_id', data.OfficerID);
  if (officer) await supabaseNotify(officer.member_id, 'Supervisor check-in logged', 'A supervisor check-in has been added to your profile.', state.user?.UserID);
  return { ok: true, CheckinID: result.data.checkin_id };
}

async function supabaseSaveDevelopmentPlan(data) {
  const record = {
    officer_id: data.OfficerID,
    supervisor_user_id: state.user?.UserID || null,
    goal: data.Goal || '',
    category: data.Category || 'Development',
    status: data.Status || 'Open',
    due_date: data.DueDate || null,
    notes: data.Notes || '',
    created_by: state.user?.UserID || null,
  };
  const result = data.PlanID
    ? await supabaseClient.from('development_plans').update(record).eq('plan_id', data.PlanID).select().single()
    : await supabaseClient.from('development_plans').insert(record).select().single();
  if (result.error) return { ok: false, error: result.error.message };
  const officer = await supabaseById('officers', 'officer_id', data.OfficerID);
  if (officer) await supabaseNotify(officer.member_id, 'Development plan updated', `Development goal updated: ${record.goal}`, state.user?.UserID);
  return { ok: true, PlanID: result.data.plan_id };
}

async function supabaseListNotifications() {
  const me = await supabaseCurrentProfile();
  if (!me.ok) return me;
  const profile = await supabaseProfileByUserId(me.user.UserID);
  const rows = await supabaseNotificationRows(profile.member_id);
  return { ok: true, rows, unread: rows.filter((row) => !row.ReadAt).length };
}

async function supabaseMarkNotificationsRead() {
  const me = await supabaseCurrentProfile();
  if (!me.ok) return me;
  const profile = await supabaseProfileByUserId(me.user.UserID);
  const { error } = await supabaseClient
    .from('notifications')
    .update({ read_at: new Date().toISOString() })
    .eq('member_id', profile.member_id)
    .is('read_at', null);
  return error ? { ok: false, error: error.message } : { ok: true, read: true };
}

async function supabaseListDocuments() {
  return supabaseVisibleDocuments();
}

async function supabaseSaveDocument(data) {
  const existing = data.DocumentID ? await supabaseById('documents', 'document_id', data.DocumentID) : null;
  const documentId = data.DocumentID || idForSupabase('DOC');
  const uploadedFile = data.DocumentFile instanceof File && data.DocumentFile.size > 0 ? data.DocumentFile : null;
  let storagePath = existing?.storage_path || null;
  let fileName = existing?.file_name || null;
  let fileSize = existing?.file_size || null;
  let fileType = existing?.file_type || null;

  if (uploadedFile) {
    const path = `${documentId}/${Date.now()}-${safeStorageFileName(uploadedFile.name)}`;
    const upload = await supabaseClient.storage
      .from('mo8-documents')
      .upload(path, uploadedFile, {
        cacheControl: '3600',
        upsert: false,
        contentType: uploadedFile.type || 'application/octet-stream',
      });
    if (upload.error) return { ok: false, error: upload.error.message };
    if (existing?.storage_path) {
      await supabaseClient.storage.from('mo8-documents').remove([existing.storage_path]);
    }
    storagePath = path;
    fileName = uploadedFile.name;
    fileSize = uploadedFile.size;
    fileType = uploadedFile.type || '';
  }

  const record = {
    document_id: documentId,
    title: data.Title || '',
    category: normalizeFolderPath(data.Category || state.documentFolder || 'General') || 'General',
    drive_url: uploadedFile ? '' : data.DriveURL || existing?.drive_url || '',
    storage_path: storagePath,
    file_name: fileName,
    file_size: fileSize,
    file_type: fileType,
    required_role: data.RequiredRole || 'Police Constable',
    required_tags: splitTags(data.RequiredTags || ''),
    requires_acknowledgement: truthy(data.RequiresAcknowledgement),
    status: data.Status || 'Published',
    updated_by: state.user?.UserID || null,
  };
  const result = data.DocumentID
    ? await supabaseClient.from('documents').update(record).eq('document_id', data.DocumentID)
    : await supabaseClient.from('documents').insert(record);
  return result.error ? { ok: false, error: result.error.message } : { ok: true, DocumentID: documentId };
}

async function supabaseDeleteDocument(data) {
  const existing = await supabaseById('documents', 'document_id', data.DocumentID);
  if (existing?.storage_path) {
    await supabaseClient.storage.from('mo8-documents').remove([existing.storage_path]);
  }
  const { error } = await supabaseClient.from('documents').delete().eq('document_id', data.DocumentID);
  return error ? { ok: false, error: error.message } : { ok: true };
}

async function supabaseAcknowledgeDocument(data) {
  const profile = await supabaseProfileByUserId(state.user.UserID);
  const officer = await supabaseOfficerForMember(profile.member_id);
  const { error } = await supabaseClient.from('document_acknowledgements').upsert({
    document_id: data.DocumentID,
    member_id: profile.member_id,
    officer_id: officer?.officer_id || null,
    user_id: state.user.UserID,
    acknowledged_at: new Date().toISOString(),
  }, { onConflict: 'document_id,user_id' });
  return error ? { ok: false, error: error.message } : { ok: true };
}

async function supabaseListAnnouncements() {
  return supabaseVisibleAnnouncements();
}

async function supabaseSaveAnnouncement(data) {
  const record = {
    title: data.Title || '',
    body: data.Body || '',
    audience: data.Audience || '',
    status: data.Status || 'Published',
    pinned: truthy(data.Pinned),
    expires_at: data.ExpiresAt || null,
    updated_by: state.user?.UserID || null,
  };
  const result = data.AnnouncementID
    ? await supabaseClient.from('announcements').update(record).eq('announcement_id', data.AnnouncementID).select().single()
    : await supabaseClient.from('announcements').insert(record).select().single();
  return result.error ? { ok: false, error: result.error.message } : { ok: true, AnnouncementID: result.data.announcement_id };
}

async function supabaseDeleteAnnouncement(data) {
  const { error } = await supabaseClient.from('announcements').delete().eq('announcement_id', data.AnnouncementID);
  return error ? { ok: false, error: error.message } : { ok: true };
}

async function supabaseListUsers() {
  const { data, error } = await supabaseClient.from('profiles').select('*').order('roblox_username');
  return error ? { ok: false, error: error.message } : { ok: true, rows: (data || []).map(supabaseUser) };
}

async function supabaseRequestAccount(data) {
  const username = String(data.RobloxUsername || '').trim();
  const discordId = String(data.DiscordID || '').trim();
  if (!username || !OFFICER_RANKS.includes(data.Rank || '') || !/^\d{15,22}$/.test(discordId)) {
    return { ok: false, error: 'Enter a valid Roblox username, rank, and Discord user ID.' };
  }
  const { error } = await supabaseClient.from('account_requests').insert({
    roblox_username: username,
    rank: data.Rank,
    discord_id: discordId,
    status: 'Pending',
  });
  if (error?.code === '23505') return { ok: false, error: 'A pending request already exists for that Roblox username.' };
  return error ? { ok: false, error: error.message } : { ok: true };
}

async function supabaseSaveUser(data) {
  const response = await supabaseInvokeAdminUsers(Object.assign({ action: 'saveUser' }, data));
  return response;
}

async function supabaseDeleteUser(data) {
  return supabaseInvokeAdminUsers({ action: 'deleteUser', UserID: data.UserID });
}

async function supabaseBulkUpdateOfficers(data) {
  const ids = splitTags(data.OfficerIDs || '');
  const updates = {};
  if (data.Status) updates.status = data.Status;
  if (data.Tags) updates.tags = splitTags(data.Tags);
  if (Object.keys(updates).length) {
    const { error } = await supabaseClient.from('officers').update(updates).in('officer_id', ids);
    if (error) return { ok: false, error: error.message };
  }
  let supervisorChanges = 0;
  if (Object.prototype.hasOwnProperty.call(data, 'SupervisorUserID')) {
    for (const OfficerID of ids) {
      const response = await supabaseSetOfficerSupervisor({ OfficerID, SupervisorUserID: data.SupervisorUserID || '' });
      if (!response.ok) return response;
      if (response.changed !== false) supervisorChanges += 1;
    }
  }
  if (data.TrainingReviewDate) {
    for (const OfficerID of ids) await supabaseSetTrainingReviewDate({ OfficerID, ReviewDate: data.TrainingReviewDate });
  }
  return { ok: true, updated: ids.length, supervisorChanges };
}

async function supabasePermissionsConfig() {
  const [users, rolePermissions, userPermissions] = await Promise.all([
    supabaseAll('profiles'),
    supabaseAll('permissions'),
    supabaseAll('user_permissions'),
  ]);
  return {
    ok: true,
    roles: SYSTEM_ROLES,
    permissions: ALL_PERMISSIONS,
    users: (users || []).map(supabaseUser),
    rolePermissions: (rolePermissions || []).map((row) => ({ Role: row.role, Permission: row.permission, Allowed: row.allowed ? 'TRUE' : 'FALSE' })),
    userPermissions: (userPermissions || []).map((row) => ({ UserID: row.user_id, Permission: row.permission, Allowed: row.allowed, UpdatedBy: row.updated_by || '', UpdatedAt: row.updated_at || '' })),
    defaultPermissions: {},
  };
}

async function supabaseSetRolePermission(data) {
  const { error } = await supabaseClient.from('permissions').upsert({
    role: data.Role,
    permission: data.Permission,
    allowed: truthy(data.Allowed),
  }, { onConflict: 'role,permission' });
  return error ? { ok: false, error: error.message } : { ok: true };
}

async function supabaseSetUserPermission(data) {
  const mode = data.Mode || 'Inherit';
  if (mode === 'Inherit') {
    const { error } = await supabaseClient
      .from('user_permissions')
      .delete()
      .eq('user_id', data.UserID)
      .eq('permission', data.Permission);
    return error ? { ok: false, error: error.message } : { ok: true };
  }
  const { error } = await supabaseClient.from('user_permissions').upsert({
    user_id: data.UserID,
    permission: data.Permission,
    allowed: mode,
    updated_by: state.user?.UserID || null,
  }, { onConflict: 'user_id,permission' });
  return error ? { ok: false, error: error.message } : { ok: true };
}

async function supabaseResetUserPassword(data) {
  return supabaseInvokeAdminUsers({ action: 'resetPassword', UserID: data.UserID || '' });
}

async function supabaseChangePassword(data) {
  return supabaseInvokeAdminUsers({ action: 'changePassword', NewPassword: data.NewPassword || '' });
}

async function supabaseInvokeAdminUsers(payload) {
  const { data: sessionData } = await supabaseClient.auth.getSession();
  const accessToken = sessionData.session?.access_token;
  if (!accessToken) return { ok: false, error: 'Session expired. Please sign in again.' };

  const response = await fetch(`${SUPABASE_CONFIG.url}/functions/v1/admin-users`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      apikey: SUPABASE_CONFIG.anonKey,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(payload),
  });

  const text = await response.text();
  let body = {};
  try {
    body = text ? JSON.parse(text) : {};
  } catch (error) {
    body = { error: text || response.statusText };
  }

  if (!response.ok) {
    return { ok: false, error: body.error || `Admin function failed with HTTP ${response.status}.` };
  }
  return body || { ok: false, error: 'No response from admin-users function.' };
}

async function supabaseAuditLog() {
  const { data, error } = await supabaseClient.from('audit_log').select('*').order('timestamp', { ascending: false }).limit(200);
  return error ? { ok: false, error: error.message } : { ok: true, rows: (data || []).map(supabaseAuditRow) };
}

async function supabaseRankChangeLog() {
  const { data, error } = await supabaseClient.from('rank_changes').select('*').order('changed_at', { ascending: false });
  return error ? { ok: false, error: error.message } : { ok: true, rows: (data || []).map(supabaseRankChange) };
}

async function supabaseShiftStatus() {
  const me = await supabaseCurrentProfile();
  if (!me.ok) return me;
  const profile = await supabaseProfileByUserId(me.user.UserID);
  const officer = await supabaseOfficerForMember(profile.member_id);
  if (!officer) return { ok: true, onDuty: false, activeShift: null };
  const { data, error } = await supabaseClient
    .from('shift_logs')
    .select('*')
    .eq('officer_id', officer.officer_id)
    .eq('status', 'On Duty')
    .is('ended_at', null)
    .order('started_at', { ascending: false })
    .limit(1);
  if (error) return { ok: false, error: error.message };
  return { ok: true, onDuty: Boolean(data && data.length), activeShift: data && data.length ? supabaseShift(data[0]) : null };
}

async function supabaseStartShift() {
  const profile = await supabaseProfileByUserId(state.user.UserID);
  const officer = await supabaseOfficerForMember(profile.member_id);
  if (!officer) return { ok: false, error: 'No linked officer profile.' };
  const status = await supabaseShiftStatus();
  if (status.activeShift) return { ok: false, error: 'You already have an active shift.' };
  const result = await supabaseClient.from('shift_logs').insert({
    officer_id: officer.officer_id,
    member_id: officer.member_id,
    roblox_username: officer.roblox_username,
    callsign: officer.callsign || '',
    rank: officer.rank || '',
    status: 'On Duty',
  }).select().single();
  return result.error ? { ok: false, error: result.error.message } : { ok: true, ShiftID: result.data.shift_id };
}

async function supabaseEndShift(data) {
  const status = await supabaseShiftStatus();
  const active = status.activeShift;
  if (!active) return { ok: false, error: 'No active shift found.' };
  const { error } = await supabaseClient.from('shift_logs').update({
    ended_at: data.EndedAt || new Date().toISOString(),
    summary: data.Summary || '',
    status: 'Completed',
    updated_at: new Date().toISOString(),
  }).eq('shift_id', active.ShiftID);
  return error ? { ok: false, error: error.message } : { ok: true };
}

async function supabaseTeamShifts(data = {}) {
  const [officers, shifts, loa] = await Promise.all([supabaseAll('officers'), supabaseAll('shift_logs'), supabaseAll('loa_requests')]);
  const active = (shifts || []).filter((row) => row.status === 'On Duty' && !row.ended_at).map(supabaseShift);
  const filtered = filterShiftsByQuery(shifts || [], data);
  const metrics = (officers || []).map((officer) => {
    const officerShifts = filtered.filter((shift) => shift.officer_id === officer.officer_id);
    const total = officerShifts.reduce((sum, shift) => sum + shiftMs(shift), 0);
    const last = officerShifts.sort((a, b) => String(b.started_at).localeCompare(String(a.started_at)))[0];
    const onLoa = (loa || []).some((request) => request.officer_id === officer.officer_id && request.status === 'Approved' && isTodayInRange(request.start_date, request.end_date));
    return {
      RobloxUsername: officer.roblox_username,
      Callsign: officer.callsign || '',
      Rank: officer.rank || '',
      LoaStatus: onLoa ? 'On LOA' : 'Available',
      Shifts: officerShifts.length,
      Duration: durationText(total),
      LastShift: last?.started_at || '',
      ActivityFlag: onLoa ? 'On LOA' : total > 0 ? 'Active' : 'No activity',
    };
  });
  const profile = await supabaseProfileByUserId(state.user.UserID);
  const ownOfficer = (officers || []).find((officer) => officer.member_id === profile?.member_id);
  const ownShifts = ownOfficer ? filtered.filter((shift) => shift.officer_id === ownOfficer.officer_id) : [];
  return {
    ok: true,
    active,
    recent: filtered.sort((a, b) => String(b.started_at).localeCompare(String(a.started_at))).slice(0, 80).map((row) => {
      const converted = supabaseShift(row);
      converted.Duration = durationText(shiftMs(row));
      return converted;
    }),
    metrics,
    myStats: { Shifts: ownShifts.length, Duration: durationText(ownShifts.reduce((sum, shift) => sum + shiftMs(shift), 0)) },
  };
}

async function supabaseRequestRetrospectiveShift(data) {
  const start = new Date(data.StartedAt);
  const end = new Date(data.EndedAt);
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime()) || end <= start) return { ok: false, error: 'The shift end must be after its start.' };
  if (end > new Date()) return { ok: false, error: 'A retrospective shift cannot end in the future.' };
  const profile = await supabaseProfileByUserId(state.user.UserID);
  const officer = await supabaseOfficerForMember(profile.member_id);
  if (!officer) return { ok: false, error: 'No linked officer profile.' };
  const { error } = await supabaseClient.from('retrospective_shift_requests').insert({ officer_id: officer.officer_id, started_at: start.toISOString(), ended_at: end.toISOString(), summary: data.Summary || '', reason: data.Reason || '', status: 'Pending' });
  if (error) return { ok: false, error: error.message };
  const supervisor = officer.supervisor_user_id ? await supabaseById('profiles', 'user_id', officer.supervisor_user_id) : null;
  if (supervisor) await supabaseNotify(supervisor.member_id, 'Retrospective shift requested', `${officer.roblox_username} submitted a shift for ${formatDisplayDateTime(start.toISOString())}.`, state.user.UserID);
  return { ok: true };
}

async function supabaseReviewRetrospectiveShift(data) {
  const request = await supabaseById('retrospective_shift_requests', 'request_id', data.RequestID);
  if (!request || request.status !== 'Pending') return { ok: false, error: 'This request is no longer pending.' };
  const start = new Date(data.StartedAt);
  const end = new Date(data.EndedAt);
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime()) || end <= start) return { ok: false, error: 'The shift end must be after its start.' };
  const officer = await supabaseById('officers', 'officer_id', request.officer_id);
  if (data.Status === 'Approved') {
    const { error: shiftError } = await supabaseClient.from('shift_logs').insert({ officer_id: officer.officer_id, member_id: officer.member_id, roblox_username: officer.roblox_username, callsign: officer.callsign || '', rank: officer.rank || '', started_at: start.toISOString(), ended_at: end.toISOString(), summary: data.Summary || '', status: 'Completed' });
    if (shiftError) return { ok: false, error: shiftError.message };
  }
  const { error } = await supabaseClient.from('retrospective_shift_requests').update({ status: data.Status, review_reason: data.ReviewReason || '', reviewed_by: state.user.UserID, reviewed_at: new Date().toISOString(), started_at: start.toISOString(), ended_at: end.toISOString(), summary: data.Summary || '' }).eq('request_id', data.RequestID);
  if (!error) await supabaseNotify(officer.member_id, `Retrospective shift ${String(data.Status).toLowerCase()}`, data.ReviewReason || 'Your retrospective shift request has been reviewed.', state.user.UserID);
  return error ? { ok: false, error: error.message } : { ok: true };
}

async function supabaseSaveShift(data) {
  const start = new Date(data.StartedAt);
  const end = new Date(data.EndedAt);
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime()) || end <= start) return { ok: false, error: 'The shift end must be after its start.' };
  const { error } = await supabaseClient.from('shift_logs').update({ started_at: start.toISOString(), ended_at: end.toISOString(), summary: data.Summary || '', status: 'Completed', updated_at: new Date().toISOString() }).eq('shift_id', data.ShiftID);
  return error ? { ok: false, error: error.message } : { ok: true };
}

async function supabaseVisibleDocuments() {
  const { data, error } = await supabaseClient
    .from('documents')
    .select('*')
    .order('category', { ascending: true })
    .order('title', { ascending: true });
  if (error) return { ok: false, error: error.message, rows: [] };
  const rows = await Promise.all((data || []).map(async (row) => {
    if (!row.storage_path) return supabaseDocument(row);
    const { data: signed } = await supabaseClient.storage
      .from('mo8-documents')
      .createSignedUrl(row.storage_path, 60 * 60);
    return supabaseDocument(Object.assign({}, row, { signed_url: signed?.signedUrl || '' }));
  }));
  return { ok: true, rows };
}

async function supabaseVisibleAnnouncements() {
  const { data, error } = await supabaseClient
    .from('announcements')
    .select('*')
    .order('pinned', { ascending: false })
    .order('updated_at', { ascending: false });
  if (error) return { ok: false, error: error.message, rows: [] };
  return { ok: true, rows: (data || []).map(supabaseAnnouncement) };
}

async function supabaseProfileByUserId(userId) {
  const { data, error } = await supabaseClient.from('profiles').select('*').eq('user_id', userId).limit(1).maybeSingle();
  if (error) throw new Error(error.message);
  return data;
}

async function supabaseAll(table) {
  const { data, error } = await supabaseClient.from(table).select('*');
  if (error) throw new Error(error.message);
  return data || [];
}

async function supabaseOptionalAll(table) {
  try {
    return await supabaseAll(table);
  } catch (error) {
    if (/does not exist|schema cache/i.test(error.message || '')) return [];
    throw error;
  }
}

async function supabaseOptionalRows(table, column, value, orderColumn = '') {
  try {
    return await supabaseRows(table, column, value, orderColumn);
  } catch (error) {
    if (/does not exist|schema cache/i.test(error.message || '')) return [];
    throw error;
  }
}

async function supabaseById(table, column, value) {
  if (!value) return null;
  const { data, error } = await supabaseClient.from(table).select('*').eq(column, value).limit(1).maybeSingle();
  if (error) throw new Error(error.message);
  return data;
}

async function supabaseInsert(table, record) {
  const { data, error } = await supabaseClient.from(table).insert(record).select().limit(1).maybeSingle();
  if (error) throw new Error(error.message);
  return data;
}

async function supabaseOfficerForMember(memberId) {
  if (!memberId) return null;
  const { data, error } = await supabaseClient.from('officers').select('*').eq('member_id', memberId).order('created_at', { ascending: false }).limit(1).maybeSingle();
  if (error) throw new Error(error.message);
  return data;
}

async function supabaseRows(table, column, value, orderColumn = '') {
  let query = supabaseClient.from(table).select('*').eq(column, value);
  if (orderColumn) query = query.order(orderColumn, { ascending: false });
  const { data, error } = await query;
  if (error) throw new Error(error.message);
  return data || [];
}

async function supabaseOfficerProfile(memberId) {
  const { data, error } = await supabaseClient.from('profiles').select('*').eq('member_id', memberId).order('created_at', { ascending: false }).limit(1).maybeSingle();
  if (error) throw new Error(error.message);
  return data;
}

async function supabaseUpsertProfileForOfficer(officer, role, actorUserId) {
  const existing = await supabaseOfficerProfile(officer.member_id);
  const record = {
    member_id: officer.member_id,
    roblox_username: officer.roblox_username,
    discord_id: officer.discord_id || '',
    rank: officer.rank || 'Police Constable',
    role: role || rankToRole(officer.rank),
    status: officer.status || 'Active',
    created_by: actorUserId || null,
  };
  if (existing) {
    await supabaseClient.from('profiles').update(record).eq('user_id', existing.user_id);
    return existing.user_id;
  }
  const inserted = await supabaseInsert('profiles', record);
  return inserted.user_id;
}

async function supabaseUpsertOfficerForProfile(profile) {
  const existing = await supabaseOfficerForMember(profile.member_id);
  const record = {
    member_id: profile.member_id,
    roblox_username: profile.roblox_username,
    discord_id: profile.discord_id || '',
    rank: profile.rank || 'Police Constable',
    status: profile.status || 'Active',
  };
  if (existing) return supabaseClient.from('officers').update(record).eq('officer_id', existing.officer_id);
  return supabaseClient.from('officers').insert(record);
}

async function supabaseAudit(actorUserId, action, targetType, targetId, details = {}) {
  try {
    await supabaseClient.from('audit_log').insert({
      actor_user_id: actorUserId || null,
      action,
      target_type: targetType,
      target_id: targetId,
      details,
    });
  } catch (error) {
    // Audit failure should not block the user action.
  }
}

async function supabaseNotify(memberId, title, message, actorUserId) {
  if (!memberId) return;
  try {
    const { data } = await supabaseClient.from('notifications').insert({
      member_id: memberId,
      title,
      message,
      actor_user_id: actorUserId || null,
    }).select('notification_id').maybeSingle();
    if (data?.notification_id) sendDiscordNotification(data.notification_id);
  } catch (error) {
    // Notification failure should not block the user action.
  }
}

async function sendDiscordNotification(notificationId) {
  if (!USE_SUPABASE || !notificationId) return;
  try {
    await supabaseClient.functions.invoke('discord-alerts', {
      body: {
        action: 'sendNotification',
        notificationId,
      },
    });
  } catch (error) {
    // Discord delivery is best-effort. In-app notifications remain the source of truth.
  }
}

function rankToRole(rank = '') {
  if (['Commissioner', 'Deputy Commissioner', 'Assistant Commissioner', 'Deputy Assistant Commissioner', 'Commander', 'Chief Superintendent', 'Superintendent'].includes(rank)) return 'Command';
  if (rank === 'Chief Inspector') return 'Chief Inspector';
  if (rank === 'Inspector') return 'Inspector';
  if (rank === 'Sergeant') return 'Sergeant';
  return 'Constable';
}

function idForSupabase(prefix) {
  return `${prefix}_${crypto.randomUUID().replaceAll('-', '')}`;
}

function safeStorageFileName(name) {
  return String(name || 'document')
    .replace(/[^a-zA-Z0-9._-]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 120) || 'document';
}

async function supabaseNotificationRows(memberId) {
  const { data, error } = await supabaseClient
    .from('notifications')
    .select('*')
    .eq('member_id', memberId)
    .order('created_at', { ascending: false })
    .limit(20);
  if (error) throw new Error(error.message);
  return (data || []).map(supabaseNotification);
}

function supabaseUser(row) {
  return {
    UserID: row.user_id,
    MemberID: row.member_id,
    RobloxUsername: row.roblox_username,
    DiscordID: row.discord_id || '',
    Rank: row.rank,
    Role: row.role,
    Status: row.status,
    LastLogin: row.last_login || '',
    CreatedAt: row.created_at || '',
    CreatedBy: row.created_by || '',
  };
}

function supabaseOfficer(row) {
  return {
    OfficerID: row.officer_id,
    MemberID: row.member_id,
    RobloxUsername: row.roblox_username,
    DiscordID: row.discord_id || '',
    Callsign: row.callsign || '',
    Rank: row.rank || '',
    Status: row.status || '',
    EffectiveStatus: row.status || '',
    JoinDate: row.join_date || '',
    SupervisorUserID: row.supervisor_user_id || '',
    Tags: (row.tags || []).join(', '),
    Notes: row.notes || '',
    CreatedAt: row.created_at || '',
    UpdatedAt: row.updated_at || '',
  };
}

function decorateSupabaseOfficer(row, context = {}) {
  const officer = supabaseOfficer(row);
  const activeLoa = (context.loa || []).find((request) => request.officer_id === row.officer_id && request.status === 'Approved' && isTodayInRange(request.start_date, request.end_date));
  const activeShift = (context.shifts || []).find((shift) => shift.officer_id === row.officer_id && shift.status === 'On Duty' && !shift.ended_at);
  const supervisor = (context.profiles || []).find((profile) => profile.user_id === row.supervisor_user_id);
  const supervisorOfficer = supervisor ? (context.officers || []).find((item) => item.member_id === supervisor.member_id) : null;
  return Object.assign(officer, {
    EffectiveStatus: activeLoa ? 'LOA' : officer.Status,
    LoaStatus: activeLoa ? 'On LOA' : 'Available',
    DutyStatus: activeShift ? 'On Duty' : 'Off Duty',
    Supervisor: supervisor ? supervisor.roblox_username : 'Not assigned',
    SupervisorOfficerID: supervisorOfficer?.officer_id || '',
  });
}

function supabaseDocument(row) {
  return {
    DocumentID: row.document_id,
    Title: row.title,
    Category: row.category,
    FolderPath: row.category,
    DriveURL: row.signed_url || row.drive_url,
    ExternalURL: row.drive_url || '',
    StoragePath: row.storage_path || '',
    FileName: row.file_name || '',
    FileSize: row.file_size || '',
    FileType: row.file_type || '',
    RequiredRole: row.required_role || '',
    RequiredTags: (row.required_tags || []).join(', '),
    RequiresAcknowledgement: row.requires_acknowledgement ? 'TRUE' : 'FALSE',
    Status: row.status,
    UpdatedBy: row.updated_by || '',
    UpdatedAt: row.updated_at || '',
  };
}

function supabaseAnnouncement(row) {
  return {
    AnnouncementID: row.announcement_id,
    Title: row.title,
    Body: row.body || '',
    Audience: row.audience || '',
    Status: row.status,
    Pinned: row.pinned ? 'TRUE' : 'FALSE',
    ExpiresAt: row.expires_at || '',
    UpdatedBy: row.updated_by || '',
    UpdatedAt: row.updated_at || '',
  };
}

function supabaseTrainingOption(row) {
  return {
    OptionID: row.option_id,
    Name: row.name,
    Type: row.type,
    Status: row.status,
    SortOrder: row.sort_order,
    UpdatedBy: row.updated_by || '',
    UpdatedAt: row.updated_at || '',
  };
}

function supabaseCourse(row, trainer = {}, coTrainers = []) {
  return {
    CourseID: row.course_id,
    Title: row.title,
    Standard: row.standard,
    TrainerUserID: row.trainer_user_id || '',
    Trainer: trainer.roblox_username || row.trainer_user_id || '',
    CoTrainerUserIDs: (row.co_trainer_user_ids || []).join(', '),
    CoTrainers: coTrainers.map((profile) => profile.roblox_username).filter(Boolean).join(', '),
    CourseDate: row.course_date || '',
    Location: row.location || '',
    Capacity: row.capacity || 0,
    Status: row.status,
    Notes: row.notes || '',
    CreatedBy: row.created_by || '',
    CreatedAt: row.created_at || '',
    UpdatedAt: row.updated_at || '',
  };
}

function supabaseCourseBooking(row, course = {}, officer = {}) {
  return {
    BookingID: row.booking_id,
    CourseID: row.course_id,
    OfficerID: row.officer_id,
    Course: course.title || row.course_id,
    Standard: course.standard || '',
    CourseDate: course.course_date || '',
    Officer: officer.roblox_username || row.officer_id,
    Rank: officer.rank || '',
    Status: row.status,
    Outcome: row.outcome || '',
    Notes: row.notes || '',
    RequestedAt: row.requested_at || '',
    ReviewedBy: row.reviewed_by || '',
    ReviewedAt: row.reviewed_at || '',
  };
}

function supabaseNotification(row) {
  return {
    NotificationID: row.notification_id,
    MemberID: row.member_id,
    Title: row.title,
    Message: row.message || '',
    CreatedAt: row.created_at || '',
    ReadAt: row.read_at || '',
    ActorUserID: row.actor_user_id || '',
  };
}

function supabaseTrainingRecord(row) {
  return {
    TrainingID: row.training_id,
    OfficerID: row.officer_id,
    Standard: row.standard,
    Status: row.status,
    Assessor: row.assessor || '',
    DateCompleted: row.date_completed || '',
    ExpiryDate: row.expiry_date || '',
    Notes: row.notes || '',
    ReviewDate: row.review_date || '',
    UpdatedAt: row.updated_at || '',
  };
}

function supabaseTrainingMatrixRecords(row) {
  const standards = [
    ['taser', 'Taser'],
    ['moe', 'MOE'],
    ['blue_ticket', 'Blue Ticket'],
    ['motorbike', 'Motorbike'],
  ];
  const rows = standards
    .filter(([field]) => row[field])
    .map(([, standard]) => ({
      TrainingID: `MATRIX_${row.officer_id}_${standard}`,
      OfficerID: row.officer_id,
      Standard: standard,
      Status: 'Passed',
      Assessor: row.updated_by || '',
      DateCompleted: '',
      ExpiryDate: '',
      Notes: 'Training matrix certification',
      ReviewDate: row.review_date || '',
      UpdatedAt: row.updated_at || '',
    }));
  if (row.driving_standard) {
    rows.push({
      TrainingID: `MATRIX_${row.officer_id}_DRIVING`,
      OfficerID: row.officer_id,
      Standard: row.driving_standard,
      Status: 'Passed',
      Assessor: row.updated_by || '',
      DateCompleted: '',
      ExpiryDate: '',
      Notes: 'Driving standard',
      ReviewDate: row.review_date || '',
      UpdatedAt: row.updated_at || '',
    });
  }
  return rows;
}

function supabaseDiscipline(row, officer = {}) {
  return {
    ActionID: row.action_id,
    OfficerID: row.officer_id,
    Officer: officer.roblox_username || row.officer_id,
    Rank: officer.rank || '',
    Type: row.type,
    Summary: row.summary,
    Details: row.details || '',
    IssuedBy: row.issued_by || '',
    IssuedAt: row.issued_at || '',
    Status: row.status,
  };
}

function supabaseLoa(row, officer = {}, profiles = []) {
  const reviewer = profiles.find((profile) => profile.user_id === row.reviewed_by) || {};
  return {
    RequestID: row.request_id,
    OfficerID: row.officer_id,
    Officer: officer.roblox_username || row.officer_id,
    Rank: officer.rank || '',
    StartDate: row.start_date || '',
    EndDate: row.end_date || '',
    Reason: row.reason || '',
    Status: row.status,
    ReviewReason: row.review_reason || '',
    ReviewedBy: reviewer.roblox_username || row.reviewed_by || '',
    ReviewedAt: row.reviewed_at || '',
    CreatedAt: row.created_at || '',
  };
}

function supabaseTransfer(row, officer = {}, profiles = []) {
  const reviewer = profiles.find((profile) => profile.user_id === row.reviewed_by) || {};
  return {
    RequestID: row.request_id,
    OfficerID: row.officer_id,
    Officer: officer.roblox_username || row.officer_id,
    Rank: officer.rank || '',
    CurrentDivision: row.current_division || '',
    TargetDivision: row.target_division || '',
    TimeInMO8: row.time_in_mo8 || '',
    Reason: row.reason || '',
    HasPermission: row.has_permission || '',
    Notes: row.notes || '',
    Status: row.status || '',
    ReviewReason: row.review_reason || '',
    ReviewedBy: reviewer.roblox_username || row.reviewed_by || '',
    ReviewedAt: row.reviewed_at || '',
    CreatedAt: row.created_at || '',
  };
}

function supabaseSupervisorRequest(row, officer = {}, profiles = []) {
  const supervisor = profiles.find((profile) => profile.user_id === (row.supervisor_user_id || officer.supervisor_user_id)) || {};
  return {
    RequestID: row.request_id,
    OfficerID: row.officer_id,
    Officer: officer.roblox_username || row.officer_id,
    Rank: officer.rank || '',
    Supervisor: supervisor.roblox_username || 'Not assigned',
    SupervisorUserID: supervisor.user_id || '',
    Category: row.category || '',
    Subject: row.subject || '',
    Details: row.details || '',
    Status: row.status || '',
    ReviewReason: row.review_reason || '',
    ReviewedBy: row.reviewed_by || '',
    ReviewedAt: row.reviewed_at || '',
    CreatedAt: row.created_at || '',
  };
}

function supabaseAppeal(row, officer = {}) {
  return {
    AppealID: row.appeal_id,
    OfficerID: row.officer_id,
    Officer: officer.roblox_username || row.officer_id,
    Rank: officer.rank || '',
    SourceType: row.source_type || '',
    SourceID: row.source_id || '',
    Reason: row.reason || '',
    Status: row.status || '',
    ReviewReason: row.review_reason || '',
    ReviewedBy: row.reviewed_by || '',
    ReviewedAt: row.reviewed_at || '',
    CreatedAt: row.created_at || '',
  };
}

function supabaseCheckin(row, officer = {}, profiles = []) {
  const supervisor = profiles.find((profile) => profile.user_id === row.supervisor_user_id) || {};
  return {
    CheckinID: row.checkin_id,
    OfficerID: row.officer_id,
    Officer: officer.roblox_username || row.officer_id,
    Supervisor: supervisor.roblox_username || row.supervisor_user_id || '',
    CheckinDate: row.checkin_date || '',
    Summary: row.summary || '',
    Concerns: row.concerns || '',
    DevelopmentGoals: row.development_goals || '',
    FollowUpDate: row.follow_up_date || '',
    CreatedBy: row.created_by || '',
    CreatedAt: row.created_at || '',
  };
}

function supabaseDevelopmentPlan(row, officer = {}, profiles = []) {
  const supervisor = profiles.find((profile) => profile.user_id === row.supervisor_user_id) || {};
  return {
    PlanID: row.plan_id,
    OfficerID: row.officer_id,
    Officer: officer.roblox_username || row.officer_id,
    Supervisor: supervisor.roblox_username || row.supervisor_user_id || '',
    Goal: row.goal || '',
    Category: row.category || '',
    Status: row.status || '',
    DueDate: row.due_date || '',
    Notes: row.notes || '',
    CreatedBy: row.created_by || '',
    CreatedAt: row.created_at || '',
    UpdatedAt: row.updated_at || '',
  };
}

function supabaseShift(row) {
  return {
    ShiftID: row.shift_id,
    OfficerID: row.officer_id,
    MemberID: row.member_id || '',
    RobloxUsername: row.roblox_username || '',
    Callsign: row.callsign || '',
    Rank: row.rank || '',
    StartedAt: row.started_at || '',
    EndedAt: row.ended_at || '',
    Summary: row.summary || '',
    Status: row.status || '',
    UpdatedAt: row.updated_at || '',
  };
}

function supabaseRankChange(row) {
  return {
    ChangeID: row.change_id,
    MemberID: row.member_id || '',
    OfficerID: row.officer_id || '',
    UserID: row.user_id || '',
    RobloxUsername: row.roblox_username || '',
    PreviousRank: row.previous_rank || '',
    NewRank: row.new_rank || '',
    Reason: row.reason || '',
    ChangedBy: row.changed_by || '',
    ChangedAt: row.changed_at || '',
  };
}

function supabaseAuditRow(row) {
  return {
    AuditID: row.audit_id,
    Timestamp: row.timestamp || '',
    ActorUserID: row.actor_user_id || '',
    Action: row.action || '',
    TargetType: row.target_type || '',
    TargetID: row.target_id || '',
    Details: JSON.stringify(row.details || {}),
  };
}

function supabaseTimeline(data) {
  const items = [];
  const add = (Date, Type, Title, Detail = '') => Date && items.push({ Date, Type, Title, Detail });
  (data.rankChanges || []).forEach((row) => add(row.ChangedAt, 'Rank', `${row.PreviousRank || 'No rank'} to ${row.NewRank}`, row.Reason));
  (data.training || []).forEach((row) => add(row.DateCompleted || row.UpdatedAt, 'Training', row.Standard, row.Status));
  (data.discipline || []).forEach((row) => add(row.IssuedAt, 'Discipline', `${row.Type}: ${row.Summary}`, row.Status));
  (data.loa || []).forEach((row) => add(row.CreatedAt || row.StartDate, 'LOA', `${row.Status} LOA`, `${row.StartDate || ''} to ${row.EndDate || ''}`));
  (data.shifts || []).forEach((row) => add(row.StartedAt, 'Shift', row.Status || 'Shift logged', row.Summary));
  return items.sort((a, b) => String(b.Date || '').localeCompare(String(a.Date || ''))).slice(0, 80);
}

function filterShiftsByQuery(shifts, query) {
  const now = new Date();
  const start = query.StartDate ? new Date(query.StartDate) : query.Period === 'month' ? new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000) : new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
  const end = query.EndDate ? new Date(`${query.EndDate}T23:59:59`) : now;
  return shifts.filter((shift) => {
    const date = new Date(shift.started_at || shift.StartedAt);
    return !Number.isNaN(date.getTime()) && date >= start && date <= end;
  });
}

function shiftMs(shift) {
  const start = new Date(shift.started_at || shift.StartedAt);
  const endedAt = shift.ended_at || shift.EndedAt;
  const end = endedAt ? new Date(endedAt) : new Date();
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) return 0;
  return Math.max(0, end.getTime() - start.getTime());
}

function durationText(ms) {
  const minutes = Math.floor(ms / 60000);
  const hours = Math.floor(minutes / 60);
  const mins = minutes % 60;
  return `${hours}h ${mins}m`;
}

function isTodayInRange(start, end) {
  const today = new Date();
  const startDate = start ? new Date(start) : null;
  const endDate = end ? new Date(end) : null;
  return Boolean(startDate && endDate && startDate <= today && endDate >= today);
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

function canManageCourse(course = {}) {
  if (can('MANAGE_COURSES')) return true;
  const coTrainerIds = splitTags(course.CoTrainerUserIDs || '');
  return course.TrainerUserID === state.user?.UserID || coTrainerIds.includes(state.user?.UserID);
}

function notificationDetails(lines) {
  return lines.filter(Boolean).join('\n');
}

function detailLine(label, value) {
  const text = value === undefined || value === null ? '' : String(value).trim();
  return text ? `${label}: ${text}` : '';
}

function truthy(value) {
  return value === true || String(value).toUpperCase() === 'TRUE';
}

function splitTags(value) {
  return String(value || '').split(/[,\n;]+/).map((tag) => tag.trim()).filter(Boolean);
}

function debounce(callback, delay = 200) {
  let timer;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => callback(...args), delay);
  };
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

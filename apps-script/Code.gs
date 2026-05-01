const CONFIG = {
  sessionHours: 12,
  roles: ['Constable', 'Sergeant', 'Inspector', 'Chief Inspector', 'Command'],
  ranks: [
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
  ],
  specialistTraining: ['Taser', 'MOE', 'Blue Ticket', 'Motorbike'],
  drivingStandards: ['Basic', 'Response', 'IPP', 'Advanced', 'Advanced + TPAC'],
  sheets: {
    users: 'Users',
    sessions: 'Sessions',
    officers: 'Officers',
    training: 'TrainingRecords',
    trainingMatrix: 'TrainingMatrix',
    discipline: 'DisciplinaryActions',
    loa: 'LOARequests',
    documents: 'Documents',
    permissions: 'Permissions',
    audit: 'AuditLog',
    notifications: 'Notifications',
  },
};

const HEADERS = {
  Users: ['UserID', 'MemberID', 'RobloxUsername', 'DiscordID', 'Rank', 'Role', 'PasswordHash', 'Salt', 'Status', 'LastLogin', 'CreatedAt', 'CreatedBy'],
  Sessions: ['SessionID', 'UserID', 'TokenHash', 'CreatedAt', 'ExpiresAt', 'RevokedAt', 'UserAgent'],
  Officers: ['OfficerID', 'MemberID', 'RobloxUsername', 'DiscordID', 'Callsign', 'Rank', 'Status', 'JoinDate', 'Notes', 'CreatedAt', 'UpdatedAt'],
  TrainingRecords: ['TrainingID', 'OfficerID', 'Standard', 'Status', 'Assessor', 'DateCompleted', 'ExpiryDate', 'Notes', 'UpdatedAt'],
  TrainingMatrix: ['OfficerID', 'MemberID', 'RobloxUsername', 'Taser', 'MOE', 'Blue Ticket', 'Motorbike', 'DrivingStandard', 'ReviewDate', 'UpdatedAt', 'UpdatedBy'],
  DisciplinaryActions: ['ActionID', 'OfficerID', 'Type', 'Summary', 'Details', 'IssuedBy', 'IssuedAt', 'Status'],
  LOARequests: ['RequestID', 'OfficerID', 'StartDate', 'EndDate', 'Reason', 'Status', 'ReviewedBy', 'ReviewedAt', 'CreatedAt'],
  Documents: ['DocumentID', 'Title', 'Category', 'DriveURL', 'RequiredRole', 'Status', 'UpdatedBy', 'UpdatedAt'],
  Permissions: ['Role', 'Permission', 'Allowed'],
  AuditLog: ['AuditID', 'Timestamp', 'ActorUserID', 'Action', 'TargetType', 'TargetID', 'Details'],
  Notifications: ['NotificationID', 'MemberID', 'Title', 'Message', 'CreatedAt', 'ReadAt', 'ActorUserID'],
};

const DEFAULT_PERMISSIONS = {
  Constable: ['VIEW_DASHBOARD', 'VIEW_DOCUMENTS', 'CHANGE_OWN_PASSWORD'],
  Sergeant: ['VIEW_DASHBOARD', 'VIEW_OFFICERS', 'VIEW_TRAINING', 'MANAGE_TRAINING', 'VIEW_LOA', 'CREATE_LOA', 'VIEW_DOCUMENTS', 'CHANGE_OWN_PASSWORD'],
  Inspector: ['VIEW_DASHBOARD', 'VIEW_OFFICERS', 'EDIT_OFFICERS', 'ADD_OFFICERS', 'VIEW_TRAINING', 'MANAGE_TRAINING', 'VIEW_DISCIPLINE', 'ADD_DISCIPLINE', 'VIEW_LOA', 'CREATE_LOA', 'APPROVE_LOA', 'VIEW_DOCUMENTS', 'MANAGE_DOCUMENTS', 'CHANGE_OWN_PASSWORD'],
  'Chief Inspector': ['VIEW_DASHBOARD', 'VIEW_OFFICERS', 'EDIT_OFFICERS', 'ADD_OFFICERS', 'ARCHIVE_OFFICERS', 'VIEW_TRAINING', 'MANAGE_TRAINING', 'VIEW_DISCIPLINE', 'ADD_DISCIPLINE', 'VIEW_LOA', 'CREATE_LOA', 'APPROVE_LOA', 'VIEW_DOCUMENTS', 'MANAGE_DOCUMENTS', 'VIEW_AUDIT_LOG', 'CHANGE_OWN_PASSWORD'],
  Command: ['FULL_ACCESS'],
};

function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(HEADERS).forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    sheet.clear();
    sheet.getRange(1, 1, 1, HEADERS[sheetName].length).setValues([HEADERS[sheetName]]);
    sheet.setFrozenRows(1);
  });

  seedPermissions_();
}

function createInitialAdmin() {
  const robloxUsername = 'YourRobloxUsername';
  const discordId = 'YourDiscordID';
  const temporaryPassword = 'ChangeMe123!';

  const users = getTable_(CONFIG.sheets.users);
  if (users.rows.some((user) => String(user.RobloxUsername).toLowerCase() === robloxUsername.toLowerCase())) {
    throw new Error('Initial admin already exists.');
  }

  const salt = randomToken_();
  const now = now_();
  appendObject_(CONFIG.sheets.users, {
    UserID: id_('USR'),
    MemberID: id_('MBR'),
    RobloxUsername: robloxUsername,
    DiscordID: discordId,
    Rank: 'Commissioner',
    Role: 'Command',
    PasswordHash: hashPassword_(temporaryPassword, salt),
    Salt: salt,
    Status: 'Active',
    LastLogin: '',
    CreatedAt: now,
    CreatedBy: 'system',
  });
}

function doGet(e) {
  return handleRequest_(e);
}

function doPost(e) {
  return handleRequest_(e);
}

function handleRequest_(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const payload = parsePayload_(e);
    const action = payload.action || '';

    const publicActions = {
      ping: () => ok_({ message: 'MO8 MDT API online', time: now_() }),
      login: () => login_(payload),
    };

    if (publicActions[action]) {
      return json_(publicActions[action]());
    }

    const auth = requireSession_(payload.token);
    const protectedActions = {
      logout: () => logout_(auth, payload),
      me: () => ok_({ user: publicUser_(auth.user), permissions: getUserPermissions_(auth.user.Role) }),
      dashboard: () => requirePermission_(auth, 'VIEW_DASHBOARD', () => dashboard_(auth)),
      myProfile: () => getMyProfile_(auth),
      listNotifications: () => listNotifications_(auth),
      markNotificationsRead: () => markNotificationsRead_(auth),
      listOfficers: () => requirePermission_(auth, 'VIEW_OFFICERS', () => listRows_(CONFIG.sheets.officers)),
      getOfficerProfile: () => requirePermission_(auth, 'VIEW_OFFICERS', () => getOfficerProfile_(payload)),
      saveOfficer: () => requirePermission_(auth, payload.OfficerID ? 'EDIT_OFFICERS' : 'ADD_OFFICERS', () => saveOfficer_(auth, payload)),
      listTraining: () => requirePermission_(auth, 'VIEW_TRAINING', () => listTrainingCertifications_()),
      saveTraining: () => requirePermission_(auth, 'MANAGE_TRAINING', () => saveTraining_(auth, payload)),
      setOfficerTraining: () => requirePermission_(auth, 'MANAGE_TRAINING', () => setOfficerTraining_(auth, payload)),
      setDrivingStandard: () => requirePermission_(auth, 'MANAGE_TRAINING', () => setDrivingStandard_(auth, payload)),
      setTrainingReviewDate: () => requirePermission_(auth, 'MANAGE_TRAINING', () => setTrainingReviewDate_(auth, payload)),
      listDiscipline: () => requirePermission_(auth, 'VIEW_DISCIPLINE', () => listRows_(CONFIG.sheets.discipline)),
      addDiscipline: () => requirePermission_(auth, 'ADD_DISCIPLINE', () => addDiscipline_(auth, payload)),
      listLoa: () => requirePermission_(auth, 'VIEW_LOA', () => listRows_(CONFIG.sheets.loa)),
      createLoa: () => requirePermission_(auth, 'CREATE_LOA', () => createLoa_(auth, payload)),
      reviewLoa: () => requirePermission_(auth, 'APPROVE_LOA', () => reviewLoa_(auth, payload)),
      listDocuments: () => requirePermission_(auth, 'VIEW_DOCUMENTS', () => listDocuments_(auth)),
      saveDocument: () => requirePermission_(auth, 'MANAGE_DOCUMENTS', () => saveDocument_(auth, payload)),
      listUsers: () => requirePermission_(auth, 'MANAGE_USERS', () => listUsers_()),
      saveUser: () => requirePermission_(auth, 'MANAGE_USERS', () => saveUser_(auth, payload)),
      resetUserPassword: () => requirePermission_(auth, 'RESET_PASSWORDS', () => resetUserPassword_(auth, payload)),
      changePassword: () => requirePermission_(auth, 'CHANGE_OWN_PASSWORD', () => changePassword_(auth, payload)),
      auditLog: () => requirePermission_(auth, 'VIEW_AUDIT_LOG', () => listRows_(CONFIG.sheets.audit)),
    };

    if (!protectedActions[action]) {
      return json_(fail_('Unknown action.'));
    }

    return json_(protectedActions[action]());
  } catch (err) {
    return json_(fail_(err.message || String(err)));
  } finally {
    try {
      lock.releaseLock();
    } catch (err) {
      // Ignore lock release errors.
    }
  }
}

function login_(payload) {
  const username = String(payload.username || '').trim();
  const password = String(payload.password || '');
  if (!username || !password) return fail_('Username and password are required.');

  const users = getTable_(CONFIG.sheets.users);
  const match = users.rows.find((user) => String(user.RobloxUsername).toLowerCase() === username.toLowerCase());
  if (!match || match.Status !== 'Active') return fail_('Invalid login.');

  const submittedHash = hashPassword_(password, match.Salt);
  if (submittedHash !== match.PasswordHash) return fail_('Invalid login.');

  const rawToken = randomToken_() + randomToken_();
  const tokenHash = hash_(rawToken);
  const now = new Date();
  const expires = new Date(now.getTime() + CONFIG.sessionHours * 60 * 60 * 1000);

  appendObject_(CONFIG.sheets.sessions, {
    SessionID: id_('SES'),
    UserID: match.UserID,
    TokenHash: tokenHash,
    CreatedAt: now.toISOString(),
    ExpiresAt: expires.toISOString(),
    RevokedAt: '',
    UserAgent: payload.userAgent || '',
  });

  updateRow_(CONFIG.sheets.users, 'UserID', match.UserID, { LastLogin: now.toISOString() });
  audit_(match.UserID, 'LOGIN', 'User', match.UserID, { username });

  return ok_({
    token: rawToken,
    user: publicUser_(match),
    permissions: getUserPermissions_(match.Role),
    expiresAt: expires.toISOString(),
  });
}

function logout_(auth, payload) {
  const sessions = getTable_(CONFIG.sheets.sessions);
  const tokenHash = hash_(payload.token || '');
  const session = sessions.rows.find((row) => row.TokenHash === tokenHash && !row.RevokedAt);
  if (session) {
    updateRow_(CONFIG.sheets.sessions, 'SessionID', session.SessionID, { RevokedAt: now_() });
  }
  audit_(auth.user.UserID, 'LOGOUT', 'User', auth.user.UserID, {});
  return ok_({ loggedOut: true });
}

function dashboard_() {
  const officers = getTable_(CONFIG.sheets.officers).rows;
  const training = listTrainingCertifications_().rows || [];
  const discipline = getTable_(CONFIG.sheets.discipline).rows;
  const loa = getTable_(CONFIG.sheets.loa).rows;
  const documents = getTable_(CONFIG.sheets.documents).rows;
  const audit = getTable_(CONFIG.sheets.audit).rows.slice(-8).reverse();
  const missingCoreTraining = officers.filter((officer) => {
    const officerTraining = getTrainingForOfficer_(officer);
    return officerTraining.some((row) => ['Taser', 'MOE', 'Blue Ticket', 'Motorbike'].includes(row.Standard) && row.Status !== 'Passed');
  }).length;
  return ok_({
    counts: {
      activeOfficers: officers.filter((officer) => officer.Status === 'Active').length,
      loaPending: loa.filter((request) => request.Status === 'Pending').length,
      trainingRecords: training.length,
      activeDiscipline: discipline.filter((action) => action.Status === 'Active').length,
      missingCoreTraining,
      documents: documents.filter((document) => document.Status === 'Published').length,
    },
    recentAudit: audit,
    recentDocuments: documents.slice(-5).reverse(),
    pendingLoa: loa.filter((request) => request.Status === 'Pending').slice(-5).reverse(),
  });
}

function getMyProfile_(auth) {
  const officer = findOfficerForUser_(auth.user);
  const documents = listDocuments_(auth).rows || [];
  const notifications = listNotifications_(auth).rows || [];
  if (!officer) return ok_({ user: publicUser_(auth.user), officer: null, training: [], loa: [], documents, notifications });
  return ok_({
    user: publicUser_(auth.user),
    officer,
    training: getTrainingForOfficer_(officer),
    loa: getTable_(CONFIG.sheets.loa).rows.filter((row) => row.OfficerID === officer.OfficerID),
    documents,
    notifications,
  });
}

function getOfficerProfile_(payload) {
  if (!payload.OfficerID) return fail_('OfficerID is required.');
  const officer = getTable_(CONFIG.sheets.officers).rows.find((row) => row.OfficerID === payload.OfficerID);
  if (!officer) return fail_('Officer not found.');

  const training = getTrainingForOfficer_(officer);
  const discipline = getTable_(CONFIG.sheets.discipline).rows.filter((row) => row.OfficerID === payload.OfficerID);
  const loa = getTable_(CONFIG.sheets.loa).rows.filter((row) => row.OfficerID === payload.OfficerID);
  return ok_({ officer, training, discipline, loa });
}

function saveOfficer_(auth, payload) {
  const now = now_();
  const officer = {
    MemberID: payload.MemberID || '',
    RobloxUsername: payload.RobloxUsername || '',
    DiscordID: payload.DiscordID || '',
    Callsign: payload.Callsign || '',
    Rank: payload.Rank || '',
    Status: payload.Status || 'Active',
    JoinDate: payload.JoinDate || '',
    Notes: payload.Notes || '',
    UpdatedAt: now,
  };

  if (payload.OfficerID) {
    const existingOfficer = getTable_(CONFIG.sheets.officers).rows.find((row) => row.OfficerID === payload.OfficerID);
    if (!existingOfficer) return fail_('Officer not found.');
    officer.MemberID = officer.MemberID || existingOfficer.MemberID || memberIdForUsername_(officer.RobloxUsername) || id_('MBR');
    updateRow_(CONFIG.sheets.officers, 'OfficerID', payload.OfficerID, officer);
    ensureTrainingMatrixRow_(Object.assign({ OfficerID: payload.OfficerID }, officer));
    const userSync = syncUserForOfficer_(auth, Object.assign({ OfficerID: payload.OfficerID }, officer));
    notifyMember_(officer.MemberID, 'Officer record updated', 'Your MO8 officer record was updated.', auth.user.UserID);
    audit_(auth.user.UserID, 'UPDATE_OFFICER', 'Officer', payload.OfficerID, officer);
    return ok_({ OfficerID: payload.OfficerID });
  }

  officer.MemberID = officer.MemberID || memberIdForUsername_(officer.RobloxUsername) || id_('MBR');
  officer.OfficerID = id_('OFF');
  officer.CreatedAt = now;
  appendObject_(CONFIG.sheets.officers, officer);
  ensureTrainingMatrixRow_(officer);
  const userSync = syncUserForOfficer_(auth, officer);
  notifyMember_(officer.MemberID, 'MO8 account created', 'Your MDT account and officer profile were created.', auth.user.UserID);
  audit_(auth.user.UserID, 'CREATE_OFFICER', 'Officer', officer.OfficerID, officer);
  return ok_({ OfficerID: officer.OfficerID, temporaryPassword: userSync.temporaryPassword || '' });
}

function saveTraining_(auth, payload) {
  const record = {
    OfficerID: payload.OfficerID || '',
    Standard: payload.Standard || '',
    Status: payload.Status || 'Not Started',
    Assessor: payload.Assessor || auth.user.RobloxUsername,
    DateCompleted: payload.DateCompleted || '',
    ExpiryDate: payload.ExpiryDate || '',
    Notes: payload.Notes || '',
    UpdatedAt: now_(),
  };

  if (payload.TrainingID) {
    updateRow_(CONFIG.sheets.training, 'TrainingID', payload.TrainingID, record);
    audit_(auth.user.UserID, 'UPDATE_TRAINING', 'Training', payload.TrainingID, record);
    return ok_({ TrainingID: payload.TrainingID });
  }

  record.TrainingID = id_('TRN');
  appendObject_(CONFIG.sheets.training, record);
  audit_(auth.user.UserID, 'CREATE_TRAINING', 'Training', record.TrainingID, record);
  return ok_({ TrainingID: record.TrainingID });
}

function setOfficerTraining_(auth, payload) {
  const officerId = payload.OfficerID || '';
  const standard = payload.Standard || '';
  const enabled = String(payload.Enabled).toLowerCase() === 'true' || payload.Enabled === true;
  if (!officerId || !standard) return fail_('OfficerID and Standard are required.');
  if (!CONFIG.specialistTraining.includes(standard)) return fail_('Unknown training standard.');

  const officer = getTable_(CONFIG.sheets.officers).rows.find((row) => row.OfficerID === officerId);
  if (!officer) return fail_('Officer not found.');
  const matrixRow = ensureTrainingMatrixRow_(officer);
  updateRow_(CONFIG.sheets.trainingMatrix, 'OfficerID', officerId, {
    MemberID: officer.MemberID || '',
    RobloxUsername: officer.RobloxUsername || '',
    [standard]: enabled ? 'TRUE' : 'FALSE',
    UpdatedAt: now_(),
    UpdatedBy: auth.user.RobloxUsername,
  });

  const details = { OfficerID: officerId, Standard: standard, Enabled: enabled, MatrixRow: matrixRow._rowNumber };
  notifyMember_(officer.MemberID, 'Training certification updated', `${standard} was ${enabled ? 'assigned' : 'removed'} on your profile.`, auth.user.UserID);
  audit_(auth.user.UserID, 'SET_TRAINING', 'TrainingMatrix', officerId, details);
  return ok_(details);
}

function setDrivingStandard_(auth, payload) {
  const officerId = payload.OfficerID || '';
  const selectedStandard = payload.Standard || '';
  if (!officerId) return fail_('OfficerID is required.');
  if (selectedStandard && !CONFIG.drivingStandards.includes(selectedStandard)) return fail_('Unknown driving standard.');

  const officer = getTable_(CONFIG.sheets.officers).rows.find((row) => row.OfficerID === officerId);
  if (!officer) return fail_('Officer not found.');
  ensureTrainingMatrixRow_(officer);
  updateRow_(CONFIG.sheets.trainingMatrix, 'OfficerID', officerId, {
    MemberID: officer.MemberID || '',
    RobloxUsername: officer.RobloxUsername || '',
    DrivingStandard: selectedStandard,
    UpdatedAt: now_(),
    UpdatedBy: auth.user.RobloxUsername,
  });

  audit_(auth.user.UserID, 'SET_DRIVING_STANDARD', 'Officer', officerId, { selectedStandard });
  notifyMember_(officer.MemberID, 'Driving standard updated', `Your driving standard is now ${selectedStandard || 'not set'}.`, auth.user.UserID);
  return ok_({ OfficerID: officerId, selectedStandard });
}

function setTrainingReviewDate_(auth, payload) {
  const officerId = payload.OfficerID || '';
  const reviewDate = payload.ReviewDate || '';
  if (!officerId) return fail_('OfficerID is required.');

  const officer = getTable_(CONFIG.sheets.officers).rows.find((row) => row.OfficerID === officerId);
  if (!officer) return fail_('Officer not found.');
  ensureTrainingMatrixRow_(officer);
  updateRow_(CONFIG.sheets.trainingMatrix, 'OfficerID', officerId, {
    MemberID: officer.MemberID || '',
    RobloxUsername: officer.RobloxUsername || '',
    ReviewDate: reviewDate,
    UpdatedAt: now_(),
    UpdatedBy: auth.user.RobloxUsername,
  });

  audit_(auth.user.UserID, 'SET_TRAINING_REVIEW_DATE', 'Officer', officerId, { reviewDate });
  notifyMember_(officer.MemberID, 'Training review updated', `Your training review date is now ${reviewDate || 'not set'}.`, auth.user.UserID);
  return ok_({ OfficerID: officerId, reviewDate });
}

function addDiscipline_(auth, payload) {
  const action = {
    ActionID: id_('DIS'),
    OfficerID: payload.OfficerID || '',
    Type: payload.Type || 'Note',
    Summary: payload.Summary || '',
    Details: payload.Details || '',
    IssuedBy: auth.user.UserID,
    IssuedAt: now_(),
    Status: payload.Status || 'Active',
  };
  appendObject_(CONFIG.sheets.discipline, action);
  audit_(auth.user.UserID, 'CREATE_DISCIPLINE', 'Discipline', action.ActionID, action);
  return ok_({ ActionID: action.ActionID });
}

function reviewLoa_(auth, payload) {
  if (!payload.RequestID) return fail_('RequestID is required.');
  const update = {
    Status: payload.Status || 'Pending',
    ReviewedBy: auth.user.UserID,
    ReviewedAt: now_(),
  };
  updateRow_(CONFIG.sheets.loa, 'RequestID', payload.RequestID, update);
  audit_(auth.user.UserID, 'REVIEW_LOA', 'LOARequest', payload.RequestID, update);
  return ok_({ RequestID: payload.RequestID });
}

function createLoa_(auth, payload) {
  const request = {
    RequestID: id_('LOA'),
    OfficerID: payload.OfficerID || '',
    StartDate: payload.StartDate || '',
    EndDate: payload.EndDate || '',
    Reason: payload.Reason || '',
    Status: payload.Status || 'Pending',
    ReviewedBy: '',
    ReviewedAt: '',
    CreatedAt: now_(),
  };
  if (!request.OfficerID) return fail_('OfficerID is required.');
  appendObject_(CONFIG.sheets.loa, request);
  audit_(auth.user.UserID, 'CREATE_LOA', 'LOARequest', request.RequestID, request);
  return ok_({ RequestID: request.RequestID });
}

function saveDocument_(auth, payload) {
  const document = {
    Title: payload.Title || '',
    Category: payload.Category || 'Training',
    DriveURL: payload.DriveURL || '',
    RequiredRole: payload.RequiredRole || 'Police Constable',
    Status: payload.Status || 'Published',
    UpdatedBy: auth.user.UserID,
    UpdatedAt: now_(),
  };

  if (payload.DocumentID) {
    updateRow_(CONFIG.sheets.documents, 'DocumentID', payload.DocumentID, document);
    audit_(auth.user.UserID, 'UPDATE_DOCUMENT', 'Document', payload.DocumentID, document);
    return ok_({ DocumentID: payload.DocumentID });
  }

  document.DocumentID = id_('DOC');
  appendObject_(CONFIG.sheets.documents, document);
  audit_(auth.user.UserID, 'CREATE_DOCUMENT', 'Document', document.DocumentID, document);
  return ok_({ DocumentID: document.DocumentID });
}

function listTrainingCertifications_() {
  const officers = getTable_(CONFIG.sheets.officers).rows;
  const rows = [];
  officers.forEach((officer) => {
    getTrainingForOfficer_(officer).forEach((training) => rows.push(training));
  });
  return ok_({ rows });
}

function getTrainingForOfficer_(officer) {
  const matrix = ensureTrainingMatrixRow_(officer);
  const rows = [];
  CONFIG.specialistTraining.forEach((standard) => {
    rows.push(trainingObject_(officer, standard, truthy_(matrix[standard]) ? 'Passed' : 'Not Started', matrix));
  });
  CONFIG.drivingStandards.forEach((standard) => {
    rows.push(trainingObject_(officer, standard, matrix.DrivingStandard === standard ? 'Passed' : 'Not Started', matrix));
  });
  return rows;
}

function trainingObject_(officer, standard, status, matrix) {
  return {
    TrainingID: `${officer.OfficerID}_${standard}`,
    OfficerID: officer.OfficerID,
    MemberID: officer.MemberID || '',
    RobloxUsername: officer.RobloxUsername,
    Standard: standard,
    Status: status,
    Assessor: matrix.UpdatedBy || '',
    DateCompleted: status === 'Passed' ? String(matrix.UpdatedAt || '').slice(0, 10) : '',
    ExpiryDate: '',
    ReviewDate: matrix.ReviewDate || '',
    Notes: 'Certification matrix',
    UpdatedAt: matrix.UpdatedAt || '',
  };
}

function ensureTrainingMatrixRow_(officer) {
  const table = getTable_(CONFIG.sheets.trainingMatrix);
  const existing = table.rows.find((row) => String(row.OfficerID) === String(officer.OfficerID));
  if (existing) return existing;

  const legacy = getTable_(CONFIG.sheets.training).rows.filter((row) => String(row.OfficerID) === String(officer.OfficerID));
  const row = {
    OfficerID: officer.OfficerID,
    MemberID: officer.MemberID || '',
    RobloxUsername: officer.RobloxUsername || '',
    Taser: legacyPassed_(legacy, 'Taser') ? 'TRUE' : 'FALSE',
    MOE: legacyPassed_(legacy, 'MOE') ? 'TRUE' : 'FALSE',
    'Blue Ticket': legacyPassed_(legacy, 'Blue Ticket') ? 'TRUE' : 'FALSE',
    Motorbike: legacyPassed_(legacy, 'Motorbike') ? 'TRUE' : 'FALSE',
    DrivingStandard: legacyDriving_(legacy),
    UpdatedAt: now_(),
    UpdatedBy: 'system',
  };
  appendObject_(CONFIG.sheets.trainingMatrix, row);
  return getTable_(CONFIG.sheets.trainingMatrix).rows.find((entry) => String(entry.OfficerID) === String(officer.OfficerID)) || row;
}

function legacyPassed_(rows, standard) {
  return rows.some((row) => row.Standard === standard && row.Status === 'Passed');
}

function legacyDriving_(rows) {
  const match = CONFIG.drivingStandards.slice().reverse().find((standard) => legacyPassed_(rows, standard));
  return match || '';
}

function truthy_(value) {
  return value === true || String(value).toUpperCase() === 'TRUE';
}

function listDocuments_(auth) {
  const rows = getTable_(CONFIG.sheets.documents).rows
    .filter((row) => row.Status === 'Published' || canManageDocuments_(auth))
    .filter((row) => canAccessDocument_(auth.user, row.RequiredRole || 'Police Constable'));
  return ok_({ rows });
}

function listUsers_() {
  const rows = getTable_(CONFIG.sheets.users).rows.map(publicUser_);
  return ok_({ rows });
}

function saveUser_(auth, payload) {
  const now = now_();
  const user = {
    MemberID: payload.MemberID || '',
    RobloxUsername: payload.RobloxUsername || '',
    DiscordID: payload.DiscordID || '',
    Rank: payload.Rank || 'Police Constable',
    Role: payload.Role || roleForRank_(payload.Rank || 'Police Constable'),
    Status: payload.Status || 'Active',
  };

  if (!user.RobloxUsername) return fail_('Roblox username is required.');

  if (payload.UserID) {
    const existingUser = getTable_(CONFIG.sheets.users).rows.find((row) => row.UserID === payload.UserID);
    if (!existingUser) return fail_('User not found.');
    user.MemberID = user.MemberID || existingUser.MemberID || memberIdForUsername_(user.RobloxUsername) || id_('MBR');
    updateRow_(CONFIG.sheets.users, 'UserID', payload.UserID, user);
    syncOfficerForUser_(auth, Object.assign({ UserID: payload.UserID }, user));
    notifyMember_(user.MemberID, 'Account updated', 'Your MDT account details were updated.', auth.user.UserID);
    audit_(auth.user.UserID, 'UPDATE_USER', 'User', payload.UserID, user);
    return ok_({ UserID: payload.UserID });
  }

  const temporaryPassword = payload.TemporaryPassword || randomPassword_();
  const salt = randomToken_();
  user.UserID = id_('USR');
  user.MemberID = user.MemberID || memberIdForUsername_(user.RobloxUsername) || id_('MBR');
  user.PasswordHash = hashPassword_(temporaryPassword, salt);
  user.Salt = salt;
  user.LastLogin = '';
  user.CreatedAt = now;
  user.CreatedBy = auth.user.UserID;
  appendObject_(CONFIG.sheets.users, user);
  syncOfficerForUser_(auth, user);
  notifyMember_(user.MemberID, 'MDT account created', 'Your MDT user account was created.', auth.user.UserID);
  audit_(auth.user.UserID, 'CREATE_USER', 'User', user.UserID, publicUser_(user));
  return ok_({ UserID: user.UserID, temporaryPassword });
}

function resetUserPassword_(auth, payload) {
  if (!payload.UserID) return fail_('UserID is required.');
  const temporaryPassword = payload.TemporaryPassword || randomPassword_();
  const salt = randomToken_();
  updateRow_(CONFIG.sheets.users, 'UserID', payload.UserID, {
    PasswordHash: hashPassword_(temporaryPassword, salt),
    Salt: salt,
  });
  const user = getTable_(CONFIG.sheets.users).rows.find((row) => row.UserID === payload.UserID);
  if (user) notifyMember_(user.MemberID, 'Password reset', 'Your MDT password was reset by command staff.', auth.user.UserID);
  audit_(auth.user.UserID, 'RESET_PASSWORD', 'User', payload.UserID, {});
  return ok_({ UserID: payload.UserID, temporaryPassword });
}

function changePassword_(auth, payload) {
  const currentPassword = String(payload.CurrentPassword || '');
  const newPassword = String(payload.NewPassword || '');
  if (!currentPassword || !newPassword) return fail_('Current and new password are required.');
  if (newPassword.length < 8) return fail_('New password must be at least 8 characters.');

  const user = getTable_(CONFIG.sheets.users).rows.find((row) => row.UserID === auth.user.UserID);
  if (!user) return fail_('User not found.');
  if (hashPassword_(currentPassword, user.Salt) !== user.PasswordHash) return fail_('Current password is incorrect.');

  const salt = randomToken_();
  updateRow_(CONFIG.sheets.users, 'UserID', user.UserID, {
    PasswordHash: hashPassword_(newPassword, salt),
    Salt: salt,
  });
  audit_(auth.user.UserID, 'CHANGE_PASSWORD', 'User', user.UserID, {});
  return ok_({ changed: true });
}

function requireSession_(token) {
  if (!token) throw new Error('Login required.');
  const sessions = getTable_(CONFIG.sheets.sessions);
  const users = getTable_(CONFIG.sheets.users);
  const tokenHash = hash_(token);
  const now = new Date();
  const session = sessions.rows.find((row) => row.TokenHash === tokenHash && !row.RevokedAt && new Date(row.ExpiresAt) > now);
  if (!session) throw new Error('Session expired. Please log in again.');

  const user = users.rows.find((row) => row.UserID === session.UserID);
  if (!user || user.Status !== 'Active') throw new Error('User is not active.');
  ensureMemberIdentity_(user);
  return { session, user };
}

function requirePermission_(auth, permission, fn) {
  const permissions = getUserPermissions_(auth.user.Role);
  if (!permissions.includes('FULL_ACCESS') && !permissions.includes(permission)) {
    throw new Error('You do not have permission to perform this action.');
  }
  return fn();
}

function getUserPermissions_(role) {
  const table = getTable_(CONFIG.sheets.permissions);
  const permissions = table.rows
    .filter((row) => row.Role === role && String(row.Allowed).toUpperCase() === 'TRUE')
    .map((row) => row.Permission);
  if (role === 'Constable' && !permissions.includes('VIEW_DASHBOARD')) permissions.push('VIEW_DASHBOARD');
  if (role === 'Constable' && !permissions.includes('VIEW_DOCUMENTS')) permissions.push('VIEW_DOCUMENTS');
  if (role === 'Sergeant' && !permissions.includes('MANAGE_TRAINING')) permissions.push('MANAGE_TRAINING');
  if (permissions.includes('APPROVE_LOA') && !permissions.includes('VIEW_LOA')) permissions.push('VIEW_LOA');
  if (permissions.includes('CREATE_LOA_REVIEW_NOTE') && !permissions.includes('CREATE_LOA')) permissions.push('CREATE_LOA');
  if (!permissions.includes('CHANGE_OWN_PASSWORD')) permissions.push('CHANGE_OWN_PASSWORD');
  return permissions;
}

function syncUserForOfficer_(auth, officer) {
  const username = String(officer.RobloxUsername || '').trim();
  if (!username) return {};
  const users = getTable_(CONFIG.sheets.users);
  const role = roleForRank_(officer.Rank || 'Police Constable');
  const memberId = officer.MemberID || memberIdForUsername_(username) || id_('MBR');
  const existing = users.rows.find((row) => row.MemberID && row.MemberID === memberId)
    || users.rows.find((row) => String(row.RobloxUsername).toLowerCase() === username.toLowerCase());
  const update = {
    MemberID: memberId,
    RobloxUsername: username,
    DiscordID: officer.DiscordID || '',
    Rank: officer.Rank || 'Police Constable',
    Role: role,
    Status: officer.Status === 'Suspended' ? 'Suspended' : officer.Status === 'Archived' ? 'Archived' : 'Active',
  };

  if (existing) {
    updateRow_(CONFIG.sheets.users, 'UserID', existing.UserID, update);
    return { UserID: existing.UserID };
  }

  const temporaryPassword = randomPassword_();
  const salt = randomToken_();
  update.UserID = id_('USR');
  update.PasswordHash = hashPassword_(temporaryPassword, salt);
  update.Salt = salt;
  update.LastLogin = '';
  update.CreatedAt = now_();
  update.CreatedBy = auth.user.UserID;
  appendObject_(CONFIG.sheets.users, update);
  audit_(auth.user.UserID, 'CREATE_LINKED_USER', 'User', update.UserID, publicUser_(update));
  return { UserID: update.UserID, temporaryPassword };
}

function syncOfficerForUser_(auth, user) {
  const username = String(user.RobloxUsername || '').trim();
  if (!username) return {};
  const officers = getTable_(CONFIG.sheets.officers);
  const memberId = user.MemberID || memberIdForUsername_(username) || id_('MBR');
  const existing = officers.rows.find((row) => row.MemberID && row.MemberID === memberId)
    || officers.rows.find((row) => String(row.RobloxUsername).toLowerCase() === username.toLowerCase());
  const update = {
    MemberID: memberId,
    RobloxUsername: username,
    DiscordID: user.DiscordID || '',
    Callsign: existing ? existing.Callsign : '',
    Rank: user.Rank || 'Police Constable',
    Status: user.Status === 'Suspended' ? 'Suspended' : user.Status === 'Archived' ? 'Archived' : 'Active',
    JoinDate: existing ? existing.JoinDate : '',
    Notes: existing ? existing.Notes : '',
    UpdatedAt: now_(),
  };

  if (existing) {
    updateRow_(CONFIG.sheets.officers, 'OfficerID', existing.OfficerID, update);
    ensureTrainingMatrixRow_(Object.assign({ OfficerID: existing.OfficerID }, update));
    return { OfficerID: existing.OfficerID };
  }

  update.OfficerID = id_('OFF');
  update.CreatedAt = now_();
  appendObject_(CONFIG.sheets.officers, update);
  ensureTrainingMatrixRow_(update);
  audit_(auth.user.UserID, 'CREATE_LINKED_OFFICER', 'Officer', update.OfficerID, update);
  return { OfficerID: update.OfficerID };
}

function roleForRank_(rank) {
  if (rank === 'Police Constable') return 'Constable';
  if (rank === 'Sergeant') return 'Sergeant';
  if (rank === 'Inspector') return 'Inspector';
  if (rank === 'Chief Inspector') return 'Chief Inspector';
  return 'Command';
}

function canManageDocuments_(auth) {
  const permissions = getUserPermissions_(auth.user.Role);
  return permissions.includes('FULL_ACCESS') || permissions.includes('MANAGE_DOCUMENTS');
}

function canAccessDocument_(user, required) {
  const rank = user.Rank || 'Police Constable';
  const role = user.Role || roleForRank_(rank);
  const rankIndex = CONFIG.ranks.indexOf(required);
  if (rankIndex !== -1) return CONFIG.ranks.indexOf(rank) >= rankIndex;

  const roleOrder = ['Constable', 'Sergeant', 'Inspector', 'Chief Inspector', 'Command'];
  const roleIndex = roleOrder.indexOf(required);
  if (roleIndex !== -1) return roleOrder.indexOf(role) >= roleIndex;
  return true;
}

function seedPermissions_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.sheets.permissions);
  const rows = [];
  Object.keys(DEFAULT_PERMISSIONS).forEach((role) => {
    DEFAULT_PERMISSIONS[role].forEach((permission) => {
      rows.push([role, permission, true]);
    });
  });
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }
}

function listRows_(sheetName) {
  return ok_({ rows: getTable_(sheetName).rows });
}

function getTable_(sheetName) {
  const sheet = ensureSheet_(sheetName);
  if (!sheet) throw new Error(`Missing sheet: ${sheetName}`);
  const values = sheet.getDataRange().getValues();
  const headers = values.shift() || [];
  const rows = values
    .filter((row) => row.some((cell) => cell !== ''))
    .map((row, index) => {
      const object = { _rowNumber: index + 2 };
      headers.forEach((header, columnIndex) => {
        object[header] = row[columnIndex];
      });
      return object;
    });
  return { sheet, headers, rows };
}

function ensureSheet_(sheetName) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(sheetName);
  const headers = HEADERS[sheetName];
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
    return sheet;
  }

  if (headers) {
    const current = sheet.getLastColumn() ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
    const missing = headers.filter((header) => !current.includes(header));
    if (missing.length) {
      sheet.getRange(1, current.length + 1, 1, missing.length).setValues([missing]);
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function appendObject_(sheetName, object) {
  const table = getTable_(sheetName);
  const values = table.headers.map((header) => object[header] !== undefined ? object[header] : '');
  table.sheet.appendRow(values);
}

function updateRow_(sheetName, key, value, updates) {
  const table = getTable_(sheetName);
  const row = table.rows.find((entry) => String(entry[key]) === String(value));
  if (!row) throw new Error(`${sheetName} row not found.`);

  Object.keys(updates).forEach((field) => {
    const columnIndex = table.headers.indexOf(field);
    if (columnIndex !== -1) {
      table.sheet.getRange(row._rowNumber, columnIndex + 1).setValue(updates[field]);
    }
  });
}

function audit_(actorUserId, action, targetType, targetId, details) {
  appendObject_(CONFIG.sheets.audit, {
    AuditID: id_('AUD'),
    Timestamp: now_(),
    ActorUserID: actorUserId,
    Action: action,
    TargetType: targetType,
    TargetID: targetId,
    Details: JSON.stringify(details || {}),
  });
}

function parsePayload_(e) {
  const params = Object.assign({}, e && e.parameter ? e.parameter : {});
  if (params.payload) {
    try {
      return Object.assign(params, JSON.parse(params.payload));
    } catch (err) {
      throw new Error('Invalid payload JSON.');
    }
  }

  if (e && e.postData && e.postData.contents) {
    try {
      return Object.assign(params, JSON.parse(e.postData.contents));
    } catch (err) {
      return params;
    }
  }

  return params;
}

function json_(object) {
  return ContentService
    .createTextOutput(JSON.stringify(object))
    .setMimeType(ContentService.MimeType.JSON);
}

function ok_(data) {
  return Object.assign({ ok: true }, data || {});
}

function fail_(message) {
  return { ok: false, error: message };
}

function publicUser_(user) {
  return {
    UserID: user.UserID,
    MemberID: user.MemberID,
    RobloxUsername: user.RobloxUsername,
    DiscordID: user.DiscordID,
    Rank: user.Rank,
    Role: user.Role,
    Status: user.Status,
  };
}

function findOfficerForUser_(user) {
  const officers = getTable_(CONFIG.sheets.officers).rows;
  return officers.find((officer) => officer.MemberID && officer.MemberID === user.MemberID)
    || officers.find((officer) => String(officer.RobloxUsername).toLowerCase() === String(user.RobloxUsername).toLowerCase())
    || null;
}

function memberIdForUsername_(username) {
  const normalized = String(username || '').toLowerCase();
  if (!normalized) return '';
  const users = getTable_(CONFIG.sheets.users).rows;
  const user = users.find((row) => String(row.RobloxUsername).toLowerCase() === normalized);
  if (user && user.MemberID) return user.MemberID;
  const officers = getTable_(CONFIG.sheets.officers).rows;
  const officer = officers.find((row) => String(row.RobloxUsername).toLowerCase() === normalized);
  return officer && officer.MemberID ? officer.MemberID : '';
}

function ensureMemberIdentity_(user) {
  if (!user || user.MemberID) return user;
  const memberId = memberIdForUsername_(user.RobloxUsername) || id_('MBR');
  user.MemberID = memberId;
  updateRow_(CONFIG.sheets.users, 'UserID', user.UserID, { MemberID: memberId });

  const officers = getTable_(CONFIG.sheets.officers).rows;
  const officer = officers.find((row) => String(row.RobloxUsername).toLowerCase() === String(user.RobloxUsername).toLowerCase());
  if (officer && !officer.MemberID) {
    updateRow_(CONFIG.sheets.officers, 'OfficerID', officer.OfficerID, { MemberID: memberId });
  }
  return user;
}

function notifyMember_(memberId, title, message, actorUserId) {
  if (!memberId) return;
  appendObject_(CONFIG.sheets.notifications, {
    NotificationID: id_('NTF'),
    MemberID: memberId,
    Title: title,
    Message: message,
    CreatedAt: now_(),
    ReadAt: '',
    ActorUserID: actorUserId || '',
  });
}

function listNotifications_(auth) {
  const memberId = auth.user.MemberID || '';
  const rows = getTable_(CONFIG.sheets.notifications).rows
    .filter((row) => row.MemberID === memberId)
    .slice(-20)
    .reverse();
  return ok_({ rows, unread: rows.filter((row) => !row.ReadAt).length });
}

function markNotificationsRead_(auth) {
  const memberId = auth.user.MemberID || '';
  const table = getTable_(CONFIG.sheets.notifications);
  table.rows.filter((row) => row.MemberID === memberId && !row.ReadAt).forEach((row) => {
    table.sheet.getRange(row._rowNumber, table.headers.indexOf('ReadAt') + 1).setValue(now_());
  });
  return ok_({ read: true });
}

function id_(prefix) {
  return `${prefix}_${Utilities.getUuid().replace(/-/g, '').slice(0, 16).toUpperCase()}`;
}

function now_() {
  return new Date().toISOString();
}

function today_() {
  return new Date().toISOString().slice(0, 10);
}

function randomToken_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

function randomPassword_() {
  return `MO8-${Utilities.getUuid().slice(0, 8)}!`;
}

function hashPassword_(password, salt) {
  return hash_(`${salt}:${password}`);
}

function hash_(value) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8);
  return bytes.map((byte) => {
    const normalized = byte < 0 ? byte + 256 : byte;
    return (`0${normalized.toString(16)}`).slice(-2);
  }).join('');
}

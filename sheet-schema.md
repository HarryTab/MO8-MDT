const CONFIG = {
  sessionHours: 12,
  roles: ['Sergeant', 'Inspector', 'Chief Inspector', 'Command'],
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
  sheets: {
    users: 'Users',
    sessions: 'Sessions',
    officers: 'Officers',
    training: 'TrainingRecords',
    discipline: 'DisciplinaryActions',
    loa: 'LOARequests',
    documents: 'Documents',
    permissions: 'Permissions',
    audit: 'AuditLog',
  },
};

const HEADERS = {
  Users: ['UserID', 'RobloxUsername', 'DiscordID', 'Rank', 'Role', 'PasswordHash', 'Salt', 'Status', 'LastLogin', 'CreatedAt', 'CreatedBy'],
  Sessions: ['SessionID', 'UserID', 'TokenHash', 'CreatedAt', 'ExpiresAt', 'RevokedAt', 'UserAgent'],
  Officers: ['OfficerID', 'RobloxUsername', 'DiscordID', 'Callsign', 'Rank', 'Status', 'JoinDate', 'Notes', 'CreatedAt', 'UpdatedAt'],
  TrainingRecords: ['TrainingID', 'OfficerID', 'Standard', 'Status', 'Assessor', 'DateCompleted', 'ExpiryDate', 'Notes', 'UpdatedAt'],
  DisciplinaryActions: ['ActionID', 'OfficerID', 'Type', 'Summary', 'Details', 'IssuedBy', 'IssuedAt', 'Status'],
  LOARequests: ['RequestID', 'OfficerID', 'StartDate', 'EndDate', 'Reason', 'Status', 'ReviewedBy', 'ReviewedAt', 'CreatedAt'],
  Documents: ['DocumentID', 'Title', 'Category', 'DriveURL', 'RequiredRole', 'Status', 'UpdatedBy', 'UpdatedAt'],
  Permissions: ['Role', 'Permission', 'Allowed'],
  AuditLog: ['AuditID', 'Timestamp', 'ActorUserID', 'Action', 'TargetType', 'TargetID', 'Details'],
};

const DEFAULT_PERMISSIONS = {
  Sergeant: ['VIEW_DASHBOARD', 'VIEW_OFFICERS', 'VIEW_TRAINING', 'VIEW_DOCUMENTS', 'CREATE_LOA_REVIEW_NOTE'],
  Inspector: ['VIEW_DASHBOARD', 'VIEW_OFFICERS', 'EDIT_OFFICERS', 'ADD_OFFICERS', 'VIEW_TRAINING', 'MANAGE_TRAINING', 'VIEW_DISCIPLINE', 'ADD_DISCIPLINE', 'APPROVE_LOA', 'VIEW_DOCUMENTS', 'MANAGE_DOCUMENTS'],
  'Chief Inspector': ['VIEW_DASHBOARD', 'VIEW_OFFICERS', 'EDIT_OFFICERS', 'ADD_OFFICERS', 'ARCHIVE_OFFICERS', 'VIEW_TRAINING', 'MANAGE_TRAINING', 'VIEW_DISCIPLINE', 'ADD_DISCIPLINE', 'APPROVE_LOA', 'VIEW_DOCUMENTS', 'MANAGE_DOCUMENTS', 'VIEW_AUDIT_LOG'],
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
      listOfficers: () => requirePermission_(auth, 'VIEW_OFFICERS', () => listRows_(CONFIG.sheets.officers)),
      saveOfficer: () => requirePermission_(auth, payload.OfficerID ? 'EDIT_OFFICERS' : 'ADD_OFFICERS', () => saveOfficer_(auth, payload)),
      listTraining: () => requirePermission_(auth, 'VIEW_TRAINING', () => listRows_(CONFIG.sheets.training)),
      saveTraining: () => requirePermission_(auth, 'MANAGE_TRAINING', () => saveTraining_(auth, payload)),
      listDiscipline: () => requirePermission_(auth, 'VIEW_DISCIPLINE', () => listRows_(CONFIG.sheets.discipline)),
      addDiscipline: () => requirePermission_(auth, 'ADD_DISCIPLINE', () => addDiscipline_(auth, payload)),
      listLoa: () => requirePermission_(auth, 'APPROVE_LOA', () => listRows_(CONFIG.sheets.loa)),
      reviewLoa: () => requirePermission_(auth, 'APPROVE_LOA', () => reviewLoa_(auth, payload)),
      listDocuments: () => requirePermission_(auth, 'VIEW_DOCUMENTS', () => listRows_(CONFIG.sheets.documents)),
      saveDocument: () => requirePermission_(auth, 'MANAGE_DOCUMENTS', () => saveDocument_(auth, payload)),
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
  const training = getTable_(CONFIG.sheets.training).rows;
  const discipline = getTable_(CONFIG.sheets.discipline).rows;
  const loa = getTable_(CONFIG.sheets.loa).rows;
  return ok_({
    counts: {
      activeOfficers: officers.filter((officer) => officer.Status === 'Active').length,
      loaPending: loa.filter((request) => request.Status === 'Pending').length,
      trainingRecords: training.length,
      activeDiscipline: discipline.filter((action) => action.Status === 'Active').length,
    },
  });
}

function saveOfficer_(auth, payload) {
  const now = now_();
  const officer = {
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
    updateRow_(CONFIG.sheets.officers, 'OfficerID', payload.OfficerID, officer);
    audit_(auth.user.UserID, 'UPDATE_OFFICER', 'Officer', payload.OfficerID, officer);
    return ok_({ OfficerID: payload.OfficerID });
  }

  officer.OfficerID = id_('OFF');
  officer.CreatedAt = now;
  appendObject_(CONFIG.sheets.officers, officer);
  audit_(auth.user.UserID, 'CREATE_OFFICER', 'Officer', officer.OfficerID, officer);
  return ok_({ OfficerID: officer.OfficerID });
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

function saveDocument_(auth, payload) {
  const document = {
    Title: payload.Title || '',
    Category: payload.Category || 'Training',
    DriveURL: payload.DriveURL || '',
    RequiredRole: payload.RequiredRole || 'Sergeant',
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
  return permissions;
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
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
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
    RobloxUsername: user.RobloxUsername,
    DiscordID: user.DiscordID,
    Rank: user.Rank,
    Role: user.Role,
    Status: user.Status,
  };
}

function id_(prefix) {
  return `${prefix}_${Utilities.getUuid().replace(/-/g, '').slice(0, 16).toUpperCase()}`;
}

function now_() {
  return new Date().toISOString();
}

function randomToken_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
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

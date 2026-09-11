const SHEET_ID = '1woWi_Cyt1bPl2l-XJ9pHQYzVNWzNd4ngO5g0W3wqjvM';
const USERS_SHEET_NAME = 'Usuarios';

function getLootSheet() {
  return SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
}

function getUsersSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(USERS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(USERS_SHEET_NAME);
    sheet.appendRow(['username', 'passwordHash', 'salt', 'role', 'pClass', 'token']);
  }
  return sheet;
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function hashPassword(password, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + ':' + salt);
  return bytes.map(b => ((b < 0 ? b + 256 : b).toString(16)).padStart(2, '0')).join('');
}

function findUserRow(sheet, username) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username) return { rowIndex: i + 1, row: data[i] };
  }
  return null;
}

function findUserByToken(sheet, token) {
  if (!token) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][5] === token) return { rowIndex: i + 1, row: data[i] };
  }
  return null;
}

function requireRole(token, allowedRoles) {
  const sheet = getUsersSheet();
  const auth = findUserByToken(sheet, token);
  if (!auth) return { ok: false, error: 'Sesión inválida, inicia sesión de nuevo.' };
  const role = auth.row[3];
  if (allowedRoles.indexOf(role) === -1) return { ok: false, error: 'No tienes permiso para esta acción.' };
  return { ok: true, auth, role };
}

function doGet(e) {
  const token = (e.parameter && e.parameter.token) || '';
  const usersSheet = getUsersSheet();
  const auth = findUserByToken(usersSheet, token);
  if (!auth) {
    return jsonOut({ ok: false, error: 'Debes iniciar sesión para ver el historial.' });
  }

  const sheet = getLootSheet();
  const rows = sheet.getDataRange().getValues();
  rows.shift(); // quita encabezados
  const tz = Session.getScriptTimeZone();
  const records = rows
    .filter(r => r[0] !== '')
    .map(r => ({
      date: (r[0] instanceof Date) ? Utilities.formatDate(r[0], tz, 'dd/MM/yyyy') : r[0],
      name: r[1],
      pClass: r[2],
      boss: r[3],
      item: r[4]
    }))
    .reverse();
  return jsonOut(records);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // espera hasta 10s a que se libere, para no procesar dos escrituras a la vez
  } catch (err) {
    return jsonOut({ ok: false, error: 'El sistema está ocupado, intenta de nuevo en unos segundos.' });
  }

  try {
    const p = e.parameter;
    const action = p.action || 'addLoot'; // compatibilidad con versiones previas del panel

    if (action === 'register') return handleRegister(p);
    if (action === 'login') return handleLogin(p);
    if (action === 'myProgress') return handleMyProgress(p);
    if (action === 'setRole') return handleSetRole(p);
    if (action === 'addLoot') return handleAddLoot(p);
    if (action === 'listPlayers') return handleListPlayers(p);
    if (action === 'resetPassword') return handleResetPassword(p);

    return jsonOut({ ok: false, error: 'Acción desconocida.' });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function handleRegister(p) {
  const username = (p.username || '').trim();
  const password = p.password || '';
  const pClass = p.pClass || '';
  if (!username || !password) return jsonOut({ ok: false, error: 'Falta usuario o contraseña.' });

  const sheet = getUsersSheet();
  if (findUserRow(sheet, username)) return jsonOut({ ok: false, error: 'Ese nombre de PJ ya está registrado.' });

  const isFirstUser = sheet.getLastRow() <= 1;
  const salt = Utilities.getUuid();
  const hash = hashPassword(password, salt);
  const role = isFirstUser ? 'admin' : 'player';
  sheet.appendRow([username, hash, salt, role, pClass, '']);
  return jsonOut({ ok: true, role });
}

function handleLogin(p) {
  const username = (p.username || '').trim();
  const password = p.password || '';
  const sheet = getUsersSheet();
  const found = findUserRow(sheet, username);
  if (!found) return jsonOut({ ok: false, error: 'Usuario no encontrado.' });

  const storedHash = found.row[1];
  const salt = found.row[2];
  const role = found.row[3];
  const pClass = found.row[4];
  const hash = hashPassword(password, salt);
  if (hash !== storedHash) return jsonOut({ ok: false, error: 'Contraseña incorrecta.' });

  const token = Utilities.getUuid() + Utilities.getUuid();
  sheet.getRange(found.rowIndex, 6).setValue(token);
  return jsonOut({ ok: true, token, role, username, pClass });
}

function handleMyProgress(p) {
  const usersSheet = getUsersSheet();
  const auth = findUserByToken(usersSheet, p.token);
  if (!auth) return jsonOut({ ok: false, error: 'Sesión inválida, inicia sesión de nuevo.' });
  const username = auth.row[0];

  const lootSheet = getLootSheet();
  const rows = lootSheet.getDataRange().getValues();
  rows.shift();
  const tz = Session.getScriptTimeZone();
  const records = rows
    .filter(r => r[1] === username)
    .map(r => ({
      date: (r[0] instanceof Date) ? Utilities.formatDate(r[0], tz, 'dd/MM/yyyy') : r[0],
      name: r[1],
      pClass: r[2],
      boss: r[3],
      item: r[4]
    }))
    .reverse();
  return jsonOut({ ok: true, records });
}

function handleAddLoot(p) {
  const check = requireRole(p.token, ['admin', 'officer']);
  if (!check.ok) return jsonOut(check);

  const lootSheet = getLootSheet();
  try {
    lootSheet.appendRow([p.date, p.name, p.pClass, p.boss, p.item]);
    return jsonOut({ ok: true });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function handleListPlayers(p) {
  const check = requireRole(p.token, ['admin', 'officer']);
  if (!check.ok) return jsonOut(check);

  const sheet = getUsersSheet();
  const data = sheet.getDataRange().getValues();
  data.shift();
  const players = data
    .filter(r => r[0] !== '')
    .map(r => ({ username: r[0], pClass: r[4] }))
    .sort((a, b) => a.username.localeCompare(b.username));
  return jsonOut({ ok: true, players });
}

function handleResetPassword(p) {
  const check = requireRole(p.token, ['admin']);
  if (!check.ok) return jsonOut(check);

  const newPassword = p.newPassword || '';
  if (newPassword.length < 4) return jsonOut({ ok: false, error: 'La contraseña debe tener al menos 4 caracteres.' });

  const sheet = getUsersSheet();
  const target = findUserRow(sheet, (p.targetUsername || '').trim());
  if (!target) return jsonOut({ ok: false, error: 'Ese usuario no existe.' });

  const salt = Utilities.getUuid();
  const hash = hashPassword(newPassword, salt);
  sheet.getRange(target.rowIndex, 2).setValue(hash); // passwordHash
  sheet.getRange(target.rowIndex, 3).setValue(salt); // salt
  sheet.getRange(target.rowIndex, 6).setValue(''); // invalida cualquier sesión anterior
  return jsonOut({ ok: true });
}

function handleSetRole(p) {
  const check = requireRole(p.token, ['admin']);
  if (!check.ok) return jsonOut(check);

  const sheet = getUsersSheet();
  const target = findUserRow(sheet, (p.targetUsername || '').trim());
  if (!target) return jsonOut({ ok: false, error: 'Ese usuario no existe.' });

  const newRole = p.newRole;
  if (['admin', 'officer', 'player'].indexOf(newRole) === -1) return jsonOut({ ok: false, error: 'Rol inválido.' });

  sheet.getRange(target.rowIndex, 4).setValue(newRole);
  return jsonOut({ ok: true });
}

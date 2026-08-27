/** ===================== CONFIG ===================== **/
const CFG = {
  INVENTORY_SHEET: 'Inventory_Master',
  STUDENTS_SHEET: 'Students',
  STUDENTS_DATA_START_ROW: 2,
  STUDENT_ID_COLUMN: 2,
  STUDENT_NAME_COLUMN: 3,
  LOG_SHEET: 'Transaction_Log',
  TIMEZONE: Session.getScriptTimeZone() || 'America/New_York',
};

/** ===================== WEB APP ===================== **/
function doGet(e) {
  const inst = (e && e.parameter && e.parameter.inst) ? String(e.parameter.inst).trim() : '';
  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.initialInst = inst;
  return tpl.evaluate()
    .setTitle('Music Inventory Checkout');
}

/** ===================== SHEET HELPERS ===================== **/
function getSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Sheet not found: ${name}`);
  return sh;
}

function ensureStudentsSheet_() {
  const ss = SpreadsheetApp.getActive();
  const names = [CFG.STUDENTS_SHEET, 'Students'];
  for (let i = 0; i < names.length; i++) {
    const sh = ss.getSheetByName(names[i]);
    if (sh) return sh;
  }
  throw new Error(`Sheet not found. Expected one of: ${names.join(', ')}`);
}

function ensureLogSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CFG.LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CFG.LOG_SHEET);
    sh.getRange(1, 1, 1, 7).setValues([['Timestamp','Action','InstrumentID','StudentID','StudentName','UserEmail','Note']]);
  }
  return sh;
}

function headerMap_(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const map = {};
  headers.forEach((h, i) => { if (h) map[h] = i + 1; }); // 1-based col index
  return map;
}

function now_() { return new Date(); }

function formatDate_(d) {
  if (!d) return '';
  return Utilities.formatDate(new Date(d), CFG.TIMEZONE, "yyyy-MM-dd HH:mm:ss");
}

function getUserEmail_() {
  try { return Session.getActiveUser().getEmail() || ''; } catch (e) { return ''; }
}

function normalizeName_(name) {
  return String(name || '').replace(/\s+/g, ' ').trim();
}

function normalizeForCompare_(s) {
  return String(s || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase();
}

/** ===================== INVENTORY API ===================== **/
function getInstrumentRecord_(instId, includePrivateFields) {
  instId = String(instId || '').trim();
  if (!instId) return { ok: false, error: 'Missing instrument id' };

  const sh = getSheet_(CFG.INVENTORY_SHEET);
  const hm = headerMap_(sh);

  const required = ['InstrumentID','Status','CurrentStudentID','CurrentStudentName','DateOut','DateIn','Notes'];
  required.forEach(k => { if (!hm[k]) throw new Error(`Missing header in ${CFG.INVENTORY_SHEET}: ${k}`); });

  const idCol = hm['InstrumentID'];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: false, error: 'No inventory rows found' };

  const ids = sh.getRange(2, idCol, lastRow - 1, 1).getValues().map(r => String(r[0]).trim());
  const idx = ids.findIndex(v => v === instId);
  if (idx === -1) return { ok: false, error: `InstrumentID not found: ${instId}` };

  const row = idx + 2;
  const get = (name) => sh.getRange(row, hm[name]).getValue();

  const optional = (names) => {
    for (let i = 0; i < names.length; i++) {
      if (hm[names[i]]) return get(names[i]);
    }
    return '';
  };
  const source = typeof getSourceInstrumentById_ === 'function' ? getSourceInstrumentById_(instId) : null;

  return {
    ok: true,
    row,
    instrument: {
      InstrumentID: instId,
      Status: String(get('Status') || '').trim(),
      RepairStatus: String(optional(['RepairStatus', 'Repair Status']) || '').trim(),
      InstrumentName: source ? source.InstrumentName : String(optional(['Instrument', 'InstrumentName', 'Instrument Name']) || '').trim(),
      Make: source ? source.Make : String(optional(['Make']) || '').trim(),
      Model: source ? source.Model : String(optional(['Model']) || '').trim(),
      Condition: source ? source.Condition : String(optional(['Condition']) || '').trim(),
      Availability: String(get('Status') || '').trim().toUpperCase() === 'OUT' ? 'OUT' : 'IN',
      CurrentStudentID: includePrivateFields ? String(get('CurrentStudentID') || '').trim() : '',
      CurrentStudentName: includePrivateFields ? String(get('CurrentStudentName') || '').trim() : '',
      DateOut: includePrivateFields && get('DateOut') ? formatDate_(get('DateOut')) : '',
      DateIn: includePrivateFields && get('DateIn') ? formatDate_(get('DateIn')) : '',
      Notes: includePrivateFields ? String(get('Notes') || '').trim() : '',
      Serial: includePrivateFields ? (source ? source.Serial : String(optional(['Serial']) || '').trim()) : '',
    }
  };
}

function getInstrumentById(instId) {
  const record = getInstrumentRecord_(instId, false);
  if (!record.ok) return record;
  return { ok: true, instrument: record.instrument };
}

function getAdminInstrumentById(instId, adminToken) {
  try {
    requireAdmin_(adminToken);
    return getInstrumentRecord_(instId, true);
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err) };
  }
}

function isRepairHold_(repairStatus) {
  const value = String(repairStatus || '').trim().toUpperCase();
  return value && ['NO', 'NONE', 'CLEAR', 'CLEARED', 'COMPLETE', 'COMPLETED', 'READY'].indexOf(value) === -1;
}

/** ===================== STUDENTS API ===================== **/
function listStudents(adminToken) {
  try {
    requireAdmin_(adminToken);
  } catch (err) {
    return { ok: false, error: 'Admin authorization required.' };
  }

  const sh = ensureStudentsSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < CFG.STUDENTS_DATA_START_ROW) return { ok: true, students: [] };
  const values = sh.getRange(
    CFG.STUDENTS_DATA_START_ROW,
    CFG.STUDENT_ID_COLUMN,
    lastRow - CFG.STUDENTS_DATA_START_ROW + 1,
    2
  ).getValues();

  const students = [];
  for (let r = 0; r < values.length; r++) {
    const id = String(values[r][0] || '').trim();
    const name = normalizeName_(values[r][1]);
    if (!/^\d+$/.test(id) || !/^[^,]+,\s*[^,]+$/.test(name)) continue;
    students.push({ StudentID: id, StudentName: name });
  }

  students.sort((a,b) => a.StudentID.localeCompare(b.StudentID));
  return { ok: true, students };
}

function getStudentById_(studentId) {
  studentId = String(studentId || '').trim();
  if (!/^\d+$/.test(studentId)) return null;

  const sh = ensureStudentsSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < CFG.STUDENTS_DATA_START_ROW) return null;
  const values = sh.getRange(
    CFG.STUDENTS_DATA_START_ROW,
    CFG.STUDENT_ID_COLUMN,
    lastRow - CFG.STUDENTS_DATA_START_ROW + 1,
    2
  ).getValues();

  for (let r = 0; r < values.length; r++) {
    const id = String(values[r][0] || '').trim();
    if (id !== studentId) continue;

    const name = normalizeName_(values[r][1]);
    if (!/^[^,]+,\s*[^,]+$/.test(name)) return null;
    return { StudentID: id, StudentName: name };
  }

  return null;
}

function addStudent(studentId, studentName, adminToken) {
  try {
    requireAdmin_(adminToken);
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err) };
  }

  studentId = String(studentId || '').trim();
  studentName = normalizeName_(String(studentName || '').trim());

  if (!studentId) return { ok: false, error: 'Student ID is required.' };
  if (!studentName) return { ok: false, error: 'Student name is required.' };

  // Student ID must be numbers only
  if (!/^\d+$/.test(studentId)) {
    return { ok: false, error: 'Student ID must be numbers only.' };
  }

  // Require "Last, First" format
  if (!/^[^,]+,\s*[^,]+$/.test(studentName)) {
    return { ok: false, error: 'Name format must be "Last, First". Example: "Student, Example".' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sh = ensureStudentsSheet_();
    const lastRow = sh.getLastRow();
    const data = lastRow >= CFG.STUDENTS_DATA_START_ROW
      ? sh.getRange(
          CFG.STUDENTS_DATA_START_ROW,
          CFG.STUDENT_ID_COLUMN,
          lastRow - CFG.STUDENTS_DATA_START_ROW + 1,
          1
        ).getValues()
      : [];

    for (let r = 0; r < data.length; r++) {
      const existingId = String(data[r][0] || '').trim();
      if (existingId && existingId === studentId) {
        return { ok: false, error: 'That Student ID already exists.' };
      }
    }

    const row = Math.max(lastRow + 1, CFG.STUDENTS_DATA_START_ROW);
    sh.getRange(row, CFG.STUDENT_ID_COLUMN).setValue(studentId);
    sh.getRange(row, CFG.STUDENT_NAME_COLUMN).setValue(studentName);

    return { ok: true, student: { StudentID: studentId, StudentName: studentName } };
  } finally {
    lock.releaseLock();
  }
}

/** ===================== CHECKOUT / CHECKIN ===================== **/
function signOutInstrument(instId, studentId, signatureName, note, override, overridePassword) {
  instId = String(instId || '').trim();
  studentId = String(studentId || '').trim();
  signatureName = normalizeName_(String(signatureName || '').trim());
  note = String(note || '').trim();
  override = (override === true || String(override).toLowerCase() === 'true');

  if (!instId) return { ok: false, error: 'Missing InstrumentID.' };
  if (!studentId) return { ok: false, error: 'Student ID is required.' };
  if (!signatureName) return { ok: false, error: 'Signature is required (Last, First).' };

  if (!/^[^,]+,\s*[^,]+$/.test(signatureName)) {
    return { ok: false, error: 'Signature must be "Last, First". Example: "Student, Example".' };
  }

  const student = getStudentById_(studentId);
  if (!student) return { ok: false, error: 'Student ID/signature could not be validated.' };

  if (normalizeForCompare_(signatureName) !== normalizeForCompare_(student.StudentName)) {
    return { ok: false, error: 'Student ID/signature could not be validated.' };
  }

  if (override) {
    try {
      if (!verifyAdminPassword_(overridePassword)) {
        return { ok: false, error: 'Incorrect admin password. Override was not completed.' };
      }
    } catch (err) {
      return { ok: false, error: err && err.message ? err.message : String(err) };
    }
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const inv = getInstrumentRecord_(instId, true);
    if (!inv.ok) return inv;
    const sh = getSheet_(CFG.INVENTORY_SHEET);
    const hm = headerMap_(sh);
    const status = String(inv.instrument.Status || '').toUpperCase();

    if (isRepairHold_(inv.instrument.RepairStatus)) {
      return { ok: false, error: `Blocked: ${instId} is on repair hold (${inv.instrument.RepairStatus}).` };
    }
    if (status === 'OUT' && !override) {
      return { ok: false, error: `Blocked: ${instId} is already OUT. Use Override with a note if necessary.` };
    }
    if (status === 'OUT' && override && !note) {
      return { ok: false, error: 'Override requires a note.' };
    }

    const row = inv.row;
    sh.getRange(row, hm['Status']).setValue('OUT');
    sh.getRange(row, hm['CurrentStudentID']).setValue(student.StudentID);
    sh.getRange(row, hm['CurrentStudentName']).setValue(student.StudentName);
    sh.getRange(row, hm['DateOut']).setValue(now_());
    sh.getRange(row, hm['DateIn']).setValue('');
    if (note) sh.getRange(row, hm['Notes']).setValue(note);
    appendLog_('SIGN_OUT' + (override ? '_OVERRIDE' : ''), instId, student.StudentID, student.StudentName, note);
    return getInstrumentById(instId);
  } finally {
    lock.releaseLock();
  }
}

function signInInstrument(instId, studentId, signatureName, note) {
  instId = String(instId || '').trim();
  studentId = String(studentId || '').trim();
  signatureName = normalizeName_(String(signatureName || '').trim());
  note = String(note || '').trim();

  if (!instId) return { ok: false, error: 'Missing InstrumentID.' };
  if (!studentId) return { ok: false, error: 'Student ID is required to sign in.' };
  if (!signatureName) return { ok: false, error: 'Signature is required (Last, First).' };

  // Require "Last, First" format for signature too
  if (!/^[^,]+,\s*[^,]+$/.test(signatureName)) {
    return { ok: false, error: 'Signature must be "Last, First". Example: "Student, Example".' };
  }

  const student = getStudentById_(studentId);
  if (!student) return { ok: false, error: 'Student ID/signature could not be validated.' };

  // Signature must match roster name for that ID
  if (normalizeForCompare_(signatureName) !== normalizeForCompare_(student.StudentName)) {
    return { ok: false, error: 'Student ID/signature could not be validated.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const inv = getInstrumentRecord_(instId, true);
    if (!inv.ok) return inv;
    const currentId = String(inv.instrument.CurrentStudentID || '').trim();
    const status = String(inv.instrument.Status || '').toUpperCase();

    if (status === 'OUT' && currentId && currentId !== studentId) {
      return { ok: false, error: 'This instrument is checked out to a different Student ID. Ask staff for help.' };
    }

    const sh = getSheet_(CFG.INVENTORY_SHEET);
    const hm = headerMap_(sh);
    const row = inv.row;
    sh.getRange(row, hm['Status']).setValue('IN');
    sh.getRange(row, hm['DateIn']).setValue(now_());
    sh.getRange(row, hm['CurrentStudentID']).setValue('');
    sh.getRange(row, hm['CurrentStudentName']).setValue('');
    if (note) sh.getRange(row, hm['Notes']).setValue(note);
    appendLog_('SIGN_IN', instId, student.StudentID, student.StudentName, note);
    return getInstrumentById(instId);
  } finally {
    lock.releaseLock();
  }
}

/** ===================== LOGGING ===================== **/
function appendLog_(action, instId, studentId, studentName, note) {
  const sh = ensureLogSheet_();
  const email = getUserEmail_();
  sh.appendRow([formatDate_(now_()), action, instId, studentId, studentName, email, note || '']);
}

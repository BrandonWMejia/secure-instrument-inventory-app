const SOURCE_INVENTORY_SHEET = 'Source_Inventory';
const SOURCE_INVENTORY_SHEET_ALIASES = [
  SOURCE_INVENTORY_SHEET,
  'Source_Inventory'
];
const SOURCE_HEADER_ROW = 2;
const SOURCE_DATA_START_ROW = SOURCE_HEADER_ROW + 1;
const SOURCE_ID_SCAN_START_ROW = SOURCE_HEADER_ROW + 1;
const AUDIT_LOG_SHEET = 'Inventory_Audit_Log';
const ADMIN_PASSWORD_PROPERTY = 'ADMIN_PASSWORD';
const ADMIN_EMAILS_PROPERTY = 'ADMIN_EMAILS';
const WEB_APP_URL_PROPERTY = 'WEB_APP_URL';
const ADMIN_SESSION_SECONDS = 1800;
const ADMIN_SESSION_PREFIX = 'admin-session:';

function adminLogin(password) {
  try {
    if (!verifyAdminPassword_(password)) {
      return { ok: false, error: 'Incorrect admin password.' };
    }
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err) };
  }

  return { ok: true, token: createAdminSession_(), method: 'password' };
}

function createAdminSession_() {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put(
    ADMIN_SESSION_PREFIX + token,
    'true',
    ADMIN_SESSION_SECONDS
  );
  return token;
}

function verifyAdminPassword_(password) {
  const expected = PropertiesService.getScriptProperties().getProperty(ADMIN_PASSWORD_PROPERTY);
  if (!expected) throw new Error('Admin password has not been configured.');
  return String(password || '') === expected;
}

function getActiveUserEmail_() {
  try {
    return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  } catch (err) {
    return '';
  }
}

function getAdminEmails_() {
  const raw = PropertiesService.getScriptProperties().getProperty(ADMIN_EMAILS_PROPERTY) || '';
  return raw
    .split(/[,;\n]+/)
    .map(email => String(email || '').trim().toLowerCase())
    .filter(Boolean);
}

function isAdminEmail_() {
  const email = getActiveUserEmail_();
  return !!email && getAdminEmails_().indexOf(email) !== -1;
}

function getAdminAccessStatus() {
  try {
    const email = getActiveUserEmail_();
    if (!email) {
      return {
        ok: true,
        authorized: false,
        email: '',
        reason: 'Google account email was not available for this deployment.'
      };
    }

    if (!isAdminEmail_()) {
      return {
        ok: true,
        authorized: false,
        email: email,
        reason: 'This Google account is not listed as an admin.'
      };
    }

    return {
      ok: true,
      authorized: true,
      email: email,
      token: createAdminSession_(),
      method: 'google_account'
    };
  } catch (err) {
    return { ok: false, authorized: false, error: err && err.message ? err.message : String(err) };
  }
}

function requireAdmin_(token) {
  token = String(token || '').trim();
  if (!token || CacheService.getScriptCache().get(ADMIN_SESSION_PREFIX + token) !== 'true') {
    throw new Error('Admin authorization required. Please unlock the Admin section again.');
  }
}

function getSourceInventorySheet_() {
  const ss = SpreadsheetApp.getActive();
  for (let i = 0; i < SOURCE_INVENTORY_SHEET_ALIASES.length; i++) {
    const sh = ss.getSheetByName(SOURCE_INVENTORY_SHEET_ALIASES[i]);
    if (sh) return sh;
  }
  throw new Error(`Sheet not found. Expected one of: ${SOURCE_INVENTORY_SHEET_ALIASES.join(', ')}`);
}

function getInstrumentPrefix_(instrument, make, model) {
  // Classify from the instrument name only so a make/model cannot accidentally
  // change an instrument's permanent ID family.
  const s = String(instrument || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
  const has = word => new RegExp(`\\b${word}\\b`).test(s);
  const hasAny = words => words.some(has);

  // Saxophone size matters; word order may be "Alto Sax" or "Saxophone - Alto".
  if (hasAny(['sax', 'saxophone'])) {
    if (has('contrabass') || (has('contra') && has('bass'))) return 'CBSAX';
    if (has('bass')) return 'BASSAX';
    if (hasAny(['baritone', 'bari'])) return 'BSAX';
    if (has('tenor')) return 'TSAX';
    if (has('alto')) return 'ASAX';
    if (has('sopranino')) return 'SNSAX';
    if (has('soprano')) return 'SSAX';
    return 'SAX';
  }

  // Clarinet keys (Bb, Eb, A, etc.) intentionally share CLAR.
  if (hasAny(['clarinet', 'clar'])) {
    if (has('contrabass') || (has('contra') && has('bass'))) return 'CBCLAR';
    if (has('bass')) return 'BCLAR';
    if (has('alto')) return 'ACLAR';
    return 'CLAR';
  }

  if (has('piccolo')) return 'PIC';
  if (has('flute')) {
    if (has('bass')) return 'BFL';
    if (has('alto')) return 'AFL';
    return 'FL';
  }
  if (has('oboe')) return 'OB';
  if (hasAny(['contrabassoon', 'contrafagotto'])) return 'CBSN';
  if (has('bassoon')) return 'BSN';

  // Brass keys intentionally do not create new prefixes.
  if (has('trumpet')) return 'TPT';
  if (has('cornet')) return 'COR';
  if (has('flugelhorn')) return 'FLUG';
  if ((has('french') && has('horn')) || (has('horn') && has('in'))) return 'HN';
  if (has('trombone')) return has('bass') ? 'BTBN' : 'TBN';
  if (hasAny(['baritone', 'euphonium'])) return 'EUPH';
  if (has('tuba')) return 'TBA';

  if (has('snare')) return 'SNR';
  if (has('bass') && has('drum')) return 'BD';
  if (hasAny(['cymbal', 'cymbals'])) return 'CYM';
  if (hasAny(['bell', 'bells', 'glockenspiel', 'glock'])) return 'GLK';
  if (has('xylophone')) return 'XYL';
  if (has('marimba')) return 'MAR';
  if (hasAny(['vibraphone', 'vibes'])) return 'VIB';
  if (has('timpani')) return 'TIM';
  if (has('percussion')) return 'PERC';
  if (hasAny(['piano', 'keyboard'])) return 'KEY';

  return 'MISC';
}

function getSourceHeaderMap_(sh) {
  const headers = sh.getRange(SOURCE_HEADER_ROW, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map(h => String(h || '').trim());
  const map = {};
  headers.forEach((header, index) => {
    if (header) map[header] = index + 1;
  });
  return map;
}

function findSourceColumn_(headerMap, names) {
  const normalizeHeader = value => String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
  const normalizedMap = {};
  Object.keys(headerMap).forEach(header => {
    normalizedMap[normalizeHeader(header)] = headerMap[header];
  });

  for (let i = 0; i < names.length; i++) {
    if (headerMap[names[i]]) return headerMap[names[i]];
    const normalized = normalizeHeader(names[i]);
    if (normalizedMap[normalized]) return normalizedMap[normalized];
  }
  return 0;
}

function sourceMakeColumn_(headerMap) {
  return findSourceColumn_(headerMap, [
    'Make', 'Instrument Make', 'Manufacturer', 'Brand', 'Brand/Make', 'Manufacturer/Make'
  ]);
}

function sourceModelColumn_(headerMap) {
  return findSourceColumn_(headerMap, [
    'Model', 'Instrument Model', 'Model Number', 'Model #', 'Model No'
  ]);
}

function findDuplicateSerials_(serial) {
  serial = String(serial || '').trim();
  if (!serial) return [];

  const sh = getSourceInventorySheet_();

  const hm = getSourceHeaderMap_(sh);
  const serialColumn = findSourceColumn_(hm, ['Serial', 'Serial Number', 'Serial #', 'Instrument Serial Number']);
  const idColumn = findSourceColumn_(hm, ['Instrument_ID']);
  const nameColumn = findSourceColumn_(hm, ['Instrument Name']);
  if (!serialColumn) throw new Error('Missing source inventory header: Serial.');
  if (!idColumn) throw new Error('Missing source inventory header: Instrument_ID.');

  const lastRow = sh.getLastRow();
  if (lastRow < SOURCE_DATA_START_ROW) return [];

  const values = sh.getRange(
    SOURCE_DATA_START_ROW,
    1,
    lastRow - SOURCE_DATA_START_ROW + 1,
    sh.getLastColumn()
  ).getValues();
  const normalized = serial.toLowerCase();

  return values
    .filter(row => String(row[serialColumn - 1] || '').trim().toLowerCase() === normalized)
    .map(row => ({
      InstrumentID: String(row[idColumn - 1] || '').trim(),
      InstrumentName: nameColumn ? String(row[nameColumn - 1] || '').trim() : ''
    }));
}

function listInstrumentSuggestions(adminToken) {
  try {
    requireAdmin_(adminToken);

    const sh = getSourceInventorySheet_();

    const hm = getSourceHeaderMap_(sh);
    const nameColumn = findSourceColumn_(hm, ['Instrument Name']);
    const makeColumn = sourceMakeColumn_(hm);
    const modelColumn = sourceModelColumn_(hm);
    if (!nameColumn) throw new Error('Missing source inventory header: Instrument Name.');

    const lastRow = sh.getLastRow();
    if (lastRow < SOURCE_DATA_START_ROW) {
      return { ok: true, names: [], makesByInstrument: {}, modelsByInstrument: {} };
    }

    const values = sh.getRange(
      SOURCE_DATA_START_ROW,
      1,
      lastRow - SOURCE_DATA_START_ROW + 1,
      sh.getLastColumn()
    ).getValues();

    const unique = {};
    const makesByInstrument = {};
    const modelsByInstrument = {};
    values.forEach(row => {
      const name = String(row[nameColumn - 1] || '').trim();
      if (!name) return;

      const key = name.toLowerCase();
      const make = makeColumn ? String(row[makeColumn - 1] || '').trim() : '';
      const model = modelColumn ? String(row[modelColumn - 1] || '').trim() : '';

      unique[key] = name;
      if (!makesByInstrument[key]) makesByInstrument[key] = {};
      if (!modelsByInstrument[key]) modelsByInstrument[key] = {};
      if (make) makesByInstrument[key][make.toLowerCase()] = make;
      if (model) modelsByInstrument[key][model.toLowerCase()] = model;
    });

    const names = Object.keys(unique)
      .map(key => unique[key])
      .sort((a, b) => a.localeCompare(b));

    const sortedMakes = {};
    const sortedModels = {};
    Object.keys(unique).forEach(key => {
      sortedMakes[key] = Object.keys(makesByInstrument[key])
        .map(valueKey => makesByInstrument[key][valueKey])
        .sort((a, b) => a.localeCompare(b));
      sortedModels[key] = Object.keys(modelsByInstrument[key])
        .map(valueKey => modelsByInstrument[key][valueKey])
        .sort((a, b) => a.localeCompare(b));
    });

    return {
      ok: true,
      names: names,
      makesByInstrument: sortedMakes,
      modelsByInstrument: sortedModels
    };
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err) };
  }
}

function getSourceInstrumentById_(instrumentId) {
  instrumentId = String(instrumentId || '').trim();
  if (!instrumentId) return null;

  const sh = getSourceInventorySheet_();

  const hm = getSourceHeaderMap_(sh);
  const idColumn = findSourceColumn_(hm, ['Instrument_ID']);
  if (!idColumn) throw new Error('Missing source inventory header: Instrument_ID.');

  const lastRow = sh.getLastRow();
  if (lastRow < SOURCE_DATA_START_ROW) return null;

  const values = sh.getRange(
    SOURCE_DATA_START_ROW,
    1,
    lastRow - SOURCE_DATA_START_ROW + 1,
    sh.getLastColumn()
  ).getValues();

  for (let r = 0; r < values.length; r++) {
    if (String(values[r][idColumn - 1] || '').trim() !== instrumentId) continue;

    const value = names => {
      const column = findSourceColumn_(hm, names);
      return column ? String(values[r][column - 1] || '').trim() : '';
    };
    return {
      InstrumentID: instrumentId,
      InstrumentUID: value(['Instrument_UID']),
      InstrumentName: value(['Instrument Name']),
      Make: value(['Make', 'Instrument Make', 'Manufacturer', 'Brand', 'Brand/Make', 'Manufacturer/Make']),
      Model: value(['Model', 'Instrument Model', 'Model Number', 'Model #', 'Model No']),
      Serial: value(['Serial', 'Serial Number', 'Serial #', 'Instrument Serial Number']),
      Condition: value(['Condition'])
    };
  }

  return null;
}

function loadAuditInstrument(instrumentId, adminToken) {
  try {
    requireAdmin_(adminToken);

    const source = getSourceInstrumentById_(instrumentId);
    if (!source) return { ok: false, error: `InstrumentID not found in source inventory: ${instrumentId}` };

    const operational = getInstrumentRecord_(source.InstrumentID, true);
    return {
      ok: true,
      instrument: Object.assign({}, source, {
        ExpectedStatus: operational.ok ? operational.instrument.Status : '',
        ExpectedStudentID: operational.ok ? operational.instrument.CurrentStudentID : '',
        ExpectedStudentName: operational.ok ? operational.instrument.CurrentStudentName : ''
      })
    };
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err) };
  }
}

function getAuditSessionProgress_(session) {
  session = String(session || '').trim();
  const sourceSheet = getSourceInventorySheet_();
  const sourceHeaders = getSourceHeaderMap_(sourceSheet);
  const sourceIdColumn = findSourceColumn_(sourceHeaders, ['Instrument_ID']);
  let total = 0;
  if (sourceIdColumn && sourceSheet.getLastRow() >= SOURCE_DATA_START_ROW) {
    total = sourceSheet.getRange(
      SOURCE_DATA_START_ROW,
      sourceIdColumn,
      sourceSheet.getLastRow() - SOURCE_DATA_START_ROW + 1,
      1
    ).getDisplayValues().filter(row => String(row[0] || '').trim()).length;
  }

  const audited = {};
  if (session) {
    const auditSheet = ensureAuditLogSheet_();
    if (auditSheet.getLastRow() >= 2) {
      const headers = auditSheet.getRange(1, 1, 1, auditSheet.getLastColumn()).getDisplayValues()[0];
      const sessionColumn = headers.indexOf('AuditSession');
      const idColumn = headers.indexOf('InstrumentID');
      if (sessionColumn !== -1 && idColumn !== -1) {
        auditSheet.getRange(2, 1, auditSheet.getLastRow() - 1, auditSheet.getLastColumn())
          .getDisplayValues().forEach(row => {
            if (String(row[sessionColumn] || '').trim() === session) {
              const id = String(row[idColumn] || '').trim();
              if (id) audited[id] = true;
            }
          });
      }
    }
  }

  const completed = Object.keys(audited).length;
  return {
    session: session,
    total: total,
    completed: completed,
    remaining: Math.max(total - completed, 0),
    percent: total ? Math.round((completed / total) * 100) : 0
  };
}

function getAdminDashboard(adminToken, auditSession) {
  try {
    requireAdmin_(adminToken);
    const sh = getSheet_(CFG.INVENTORY_SHEET);
    const hm = headerMap_(sh);
    const counts = { total: 0, in: 0, out: 0, repair: 0 };
    if (sh.getLastRow() >= 2) {
      const values = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getDisplayValues();
      const idColumn = hm['InstrumentID'];
      const statusColumn = hm['Status'];
      const repairColumn = hm['RepairStatus'] || hm['Repair Status'];
      values.forEach(row => {
        if (!idColumn || !String(row[idColumn - 1] || '').trim()) return;
        counts.total++;
        const status = statusColumn ? String(row[statusColumn - 1] || '').trim().toUpperCase() : '';
        if (status === 'OUT') counts.out++;
        else if (status === 'IN') counts.in++;
        if (repairColumn && isRepairHold_(row[repairColumn - 1])) counts.repair++;
      });
    }
    return { ok: true, counts: counts, audit: getAuditSessionProgress_(auditSession) };
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err) };
  }
}

function ensureAuditLogSheet_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(AUDIT_LOG_SHEET);
  if (!sh) throw new Error(`Sheet not found: ${AUDIT_LOG_SHEET}`);

  const expectedHeaders = [
    'Timestamp', 'AuditSession', 'InstrumentID', 'InstrumentUID', 'InstrumentName',
    'Make', 'Model', 'Serial', 'ExpectedStatus', 'ExpectedStudentID',
    'ExpectedStudentName', 'ExpectedCondition', 'Found', 'SerialMatches',
    'CasePresent', 'AccessoriesComplete', 'Condition', 'RepairNeeded',
    'Notes', 'StaffEmail'
  ];

  if (sh.getLastRow() === 0 || sh.getLastColumn() === 0) {
    sh.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    return sh;
  }

  const existing = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getValues()[0]
    .map(header => String(header || '').trim());
  const missing = expectedHeaders.filter(header => existing.indexOf(header) === -1);
  if (missing.length) {
    sh.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
  }
  return sh;
}

function saveAuditResult(audit, adminToken) {
  audit = audit || {};
  try {
    requireAdmin_(adminToken);

    const session = String(audit.session || '').trim();
    const instrumentId = String(audit.instrumentId || '').trim();
    const found = String(audit.found || '').trim().toUpperCase();
    if (!session) return { ok: false, error: 'Audit Session is required.' };
    if (!instrumentId) return { ok: false, error: 'Instrument ID is required.' };
    if (found !== 'YES' && found !== 'NO') return { ok: false, error: 'Found must be Yes or No.' };

    const source = getSourceInstrumentById_(instrumentId);
    if (!source) return { ok: false, error: `InstrumentID not found in source inventory: ${instrumentId}` };
    const operational = getInstrumentRecord_(instrumentId, true);

    const record = {
      Timestamp: formatDate_(now_()),
      AuditSession: session,
      InstrumentID: source.InstrumentID,
      InstrumentUID: source.InstrumentUID,
      InstrumentName: source.InstrumentName,
      Make: source.Make,
      Model: source.Model,
      Serial: source.Serial,
      ExpectedStatus: operational.ok ? operational.instrument.Status : '',
      ExpectedStudentID: operational.ok ? operational.instrument.CurrentStudentID : '',
      ExpectedStudentName: operational.ok ? operational.instrument.CurrentStudentName : '',
      ExpectedCondition: source.Condition,
      Found: found,
      SerialMatches: String(audit.serialMatches || '').trim().toUpperCase(),
      CasePresent: String(audit.casePresent || '').trim().toUpperCase(),
      AccessoriesComplete: String(audit.accessoriesComplete || '').trim().toUpperCase(),
      Condition: String(audit.condition || '').trim(),
      AuditCondition: String(audit.condition || '').trim(),
      RepairNeeded: String(audit.repairNeeded || '').trim().toUpperCase(),
      Notes: String(audit.notes || '').trim(),
      StaffEmail: getUserEmail_()
    };

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const sh = ensureAuditLogSheet_();
      const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
        .getValues()[0]
        .map(header => String(header || '').trim());
      sh.appendRow(headers.map(header => record[header] !== undefined ? record[header] : ''));
    } finally {
      lock.releaseLock();
    }

    return {
      ok: true,
      instrumentId: source.InstrumentID,
      audit: getAuditSessionProgress_(session)
    };
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err) };
  }
}

function createInstrumentRecord_(instrumentName, make, model, serial, condition, allowDuplicateSerial) {
  instrumentName = String(instrumentName || '').trim();
  make = String(make || '').trim();
  model = String(model || '').trim();
  serial = String(serial || '').trim();
  condition = String(condition || '').trim();
  allowDuplicateSerial = allowDuplicateSerial === true;

  if (!instrumentName) throw new Error('Instrument Name is required.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sh = getSourceInventorySheet_();

    const hm = getSourceHeaderMap_(sh);
    const columns = {
      uid: findSourceColumn_(hm, ['Instrument_UID']),
      id: findSourceColumn_(hm, ['Instrument_ID']),
      name: findSourceColumn_(hm, ['Instrument Name']),
      make: sourceMakeColumn_(hm),
      model: sourceModelColumn_(hm),
      serial: findSourceColumn_(hm, ['Serial', 'Serial Number', 'Serial #', 'Instrument Serial Number']),
      condition: findSourceColumn_(hm, ['Condition']),
    };

    ['uid', 'id', 'name', 'serial', 'condition'].forEach(key => {
      if (!columns[key]) throw new Error(`Missing source inventory header for ${key}.`);
    });
    if (make && !columns.make) {
      throw new Error('Could not find a Make/Manufacturer/Brand column in the source inventory.');
    }
    if (model && !columns.model) {
      throw new Error('Could not find a Model/Model Number column in the source inventory.');
    }

    const duplicateSerials = findDuplicateSerials_(serial);
    if (duplicateSerials.length && !allowDuplicateSerial) {
      const err = new Error('Duplicate serial number confirmation required.');
      err.code = 'DUPLICATE_SERIAL';
      err.duplicates = duplicateSerials;
      throw err;
    }

    const prefix = getInstrumentPrefix_(instrumentName, make, model);
    const lastRow = sh.getLastRow();
    let highest = 0;
    const usedIds = {};

    if (lastRow >= SOURCE_ID_SCAN_START_ROW) {
      const existingIds = sh.getRange(
        SOURCE_ID_SCAN_START_ROW,
        columns.id,
        lastRow - SOURCE_ID_SCAN_START_ROW + 1,
        1
      ).getValues();

      existingIds.forEach(row => {
        const existing = String(row[0] || '').trim();
        if (existing) usedIds[existing] = true;
        const match = existing.match(new RegExp(`^${prefix}-(\\d+)$`));
        if (match) highest = Math.max(highest, parseInt(match[1], 10));
      });
    }

    let next = highest + 1;
    let instrumentId = `${prefix}-${String(next).padStart(3, '0')}`;
    while (usedIds[instrumentId]) {
      next++;
      instrumentId = `${prefix}-${String(next).padStart(3, '0')}`;
    }
    const instrumentUid = Utilities.getUuid();
    const newRow = Math.max(lastRow + 1, SOURCE_DATA_START_ROW);

    sh.getRange(newRow, columns.uid).setNumberFormat('@').setValue(instrumentUid);
    sh.getRange(newRow, columns.id).setNumberFormat('@').setValue(instrumentId);
    sh.getRange(newRow, columns.name).setValue(instrumentName);
    if (columns.make) sh.getRange(newRow, columns.make).setValue(make);
    if (columns.model) sh.getRange(newRow, columns.model).setValue(model);
    sh.getRange(newRow, columns.serial).setNumberFormat('@').setValue(serial);
    sh.getRange(newRow, columns.condition).setValue(condition);

    return { Instrument_UID: instrumentUid, Instrument_ID: instrumentId, row: newRow };
  } finally {
    lock.releaseLock();
  }
}

function addInstrument(form, adminToken, confirmDuplicateSerial) {
  form = form || {};
  try {
    requireAdmin_(adminToken);
    const instrument = createInstrumentRecord_(
      form.instrumentName,
      form.make,
      form.model,
      form.serial,
      form.condition,
      confirmDuplicateSerial === true
    );
    syncValidQrLabelsSafely_();
    return { ok: true, instrument: instrument };
  } catch (err) {
    if (err && err.code === 'DUPLICATE_SERIAL') {
      return {
        ok: false,
        warning: 'DUPLICATE_SERIAL',
        error: 'That serial number already exists. Review the matching instrument before adding anyway.',
        duplicates: err.duplicates || []
      };
    }
    return { ok: false, error: err && err.message ? err.message : String(err) };
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Inventory Tools')
    .addItem('Add New Instrument', 'showAddInstrumentPrompt')
    .addItem('Generate Missing IDs', 'backfillInstrumentUidsAndIds')
    .addItem('Configure QR Web App URL', 'configureQrWebAppUrl')
    .addItem('Sync Valid QR Labels', 'syncValidQrLabels')
    .addItem('Install QR Automation Triggers', 'installQrAutomationTriggers')
    .addItem('Install Parent Consent Trigger', 'installParentConsentTrigger')
    .addToUi();
}

function normalizeWebAppUrl_(url) {
  url = String(url || '').trim();
  if (!url) throw new Error('Web app URL is required.');

  url = url.split('#')[0].split('?')[0].replace(/\/+$/, '');
  if (!/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec$/.test(url)) {
    throw new Error('Use the deployed web app URL ending in /exec, not the Apps Script editor URL or /dev test URL.');
  }
  return url;
}

function configureQrWebAppUrl() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Configure QR Web App URL',
    'Paste the deployed Google Apps Script web app URL ending in /exec:',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const url = normalizeWebAppUrl_(response.getResponseText());
  PropertiesService.getScriptProperties().setProperty(WEB_APP_URL_PROPERTY, url);
  const result = syncValidQrLabels();
  ui.alert(`QR web app URL saved. Synced ${result.count} valid QR labels.`);
}

function getQrWebAppUrl_() {
  const configured = PropertiesService.getScriptProperties().getProperty(WEB_APP_URL_PROPERTY);
  if (configured) return normalizeWebAppUrl_(configured);

  const detected = ScriptApp.getService().getUrl();
  if (detected && /\/exec(?:[?#]|$)/.test(detected)) return normalizeWebAppUrl_(detected);

  throw new Error('QR web app URL is not configured. Run Inventory Tools → Configure QR Web App URL.');
}

function getValidSourceInstrumentIds_() {
  const sh = getSourceInventorySheet_();
  const hm = getSourceHeaderMap_(sh);
  const idColumn = findSourceColumn_(hm, ['Instrument_ID']);
  if (!idColumn) throw new Error('Missing source inventory header: Instrument_ID.');

  const lastRow = sh.getLastRow();
  if (lastRow < SOURCE_DATA_START_ROW) return [];

  const ids = sh.getRange(
    SOURCE_DATA_START_ROW,
    idColumn,
    lastRow - SOURCE_DATA_START_ROW + 1,
    1
  ).getValues();
  const unique = {};
  ids.forEach(row => {
    const id = String(row[0] || '').trim();
    if (id) unique[id] = true;
  });
  return Object.keys(unique).sort();
}

function syncValidQrLabels() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('QR_Labels');
  if (!sh) throw new Error('Sheet not found: QR_Labels');

  const webAppUrl = getQrWebAppUrl_();

  const ids = getValidSourceInstrumentIds_();
  const rows = ids.map(id => [id, '', `${webAppUrl}?inst=${encodeURIComponent(id)}`]);
  const existingLastRow = sh.getLastRow();

  sh.getRange(1, 1, 1, 3).setValues([['InstrumentID', 'QR', 'CodeText']]);
  if (existingLastRow > 1) sh.getRange(2, 1, existingLastRow - 1, 3).clearContent();
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 3).setValues(rows);
    const formulas = ids.map((id, index) => [
      `=IMAGE("https://quickchart.io/qr?text="&ENCODEURL(C${index + 2})&"&size=200")`
    ]);
    sh.getRange(2, 2, formulas.length, 1).setFormulas(formulas);
    const links = ids.map(id => {
      const url = `${webAppUrl}?inst=${encodeURIComponent(id)}`;
      return [SpreadsheetApp.newRichTextValue().setText(url).setLinkUrl(url).build()];
    });
    sh.getRange(2, 3, links.length, 1).setRichTextValues(links);
  }

  return { ok: true, count: ids.length };
}

function syncValidQrLabelsSafely_() {
  try {
    return syncValidQrLabels();
  } catch (err) {
    console.error(`Automatic QR label sync failed: ${err && err.message ? err.message : err}`);
    return { ok: false, error: err && err.message ? err.message : String(err) };
  }
}

function installQrAutomationTriggers() {
  const handlers = ['onInventorySourceEditInstalled', 'onInventoryStructureChange', 'scheduledQrLabelSync'];
  ScriptApp.getProjectTriggers()
    .filter(trigger => handlers.indexOf(trigger.getHandlerFunction()) !== -1)
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));

  const ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('onInventorySourceEditInstalled')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  ScriptApp.newTrigger('onInventoryStructureChange')
    .forSpreadsheet(ss)
    .onChange()
    .create();
  ScriptApp.newTrigger('scheduledQrLabelSync')
    .timeBased()
    .everyHours(6)
    .create();

  syncValidQrLabels();
  SpreadsheetApp.getUi().alert('Automatic QR label sync installed and initial sync completed.');
}

function scheduledQrLabelSync() {
  syncValidQrLabelsSafely_();
}

function onInventorySourceEditInstalled(e) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  if (SOURCE_INVENTORY_SHEET_ALIASES.indexOf(sh.getName()) === -1) return;
  if (e.range.getLastRow() < SOURCE_DATA_START_ROW) return;

  syncValidQrLabelsSafely_();
}

function onInventoryStructureChange(e) {
  if (!e || ['INSERT_ROW', 'REMOVE_ROW'].indexOf(e.changeType) === -1) return;
  syncValidQrLabelsSafely_();
}

function installParentConsentTrigger() {
  const handler = 'onParentConsentFormSubmit';
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === handler)
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger(handler)
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onFormSubmit()
    .create();

  SpreadsheetApp.getUi().alert('Parent consent duplicate-update trigger installed.');
}

function onParentConsentFormSubmit(e) {
  if (!e || !e.range) throw new Error('This function must run from a spreadsheet form-submit trigger.');

  const sh = e.range.getSheet();
  const validSheetNames = [CFG.STUDENTS_SHEET, 'Students'];
  if (validSheetNames.indexOf(sh.getName()) === -1) return;

  const submittedRow = e.range.getRow();
  if (submittedRow < CFG.STUDENTS_DATA_START_ROW) return;

  const studentId = String(sh.getRange(submittedRow, CFG.STUDENT_ID_COLUMN).getDisplayValue() || '').trim();
  if (!/^\d+$/.test(studentId)) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (submittedRow === CFG.STUDENTS_DATA_START_ROW) return;

    const lastColumn = sh.getLastColumn();
    const submittedValues = sh.getRange(submittedRow, 1, 1, lastColumn).getValues();
    const existingIds = sh.getRange(
      CFG.STUDENTS_DATA_START_ROW,
      CFG.STUDENT_ID_COLUMN,
      submittedRow - CFG.STUDENTS_DATA_START_ROW,
      1
    ).getDisplayValues();

    for (let i = 0; i < existingIds.length; i++) {
      if (String(existingIds[i][0] || '').trim() !== studentId) continue;

      const existingRow = CFG.STUDENTS_DATA_START_ROW + i;
      sh.getRange(existingRow, 1, 1, lastColumn).setValues(submittedValues);
      sh.deleteRow(submittedRow);
      return;
    }
  } finally {
    lock.releaseLock();
  }
}

function showAddInstrumentPrompt() {
  const ui = SpreadsheetApp.getUi();

  const instrumentResp = ui.prompt(
    'Add New Instrument',
    'Enter Instrument Name, for example Alto Sax, Trumpet, Clarinet:',
    ui.ButtonSet.OK_CANCEL
  );
  if (instrumentResp.getSelectedButton() !== ui.Button.OK) return;
  const instrumentName = instrumentResp.getResponseText().trim();
  if (!instrumentName) return ui.alert('Instrument Name is required.');

  const makeResp = ui.prompt(
    'Add New Instrument',
    'Enter Make, for example Yamaha, Selmer, Bach:',
    ui.ButtonSet.OK_CANCEL
  );
  if (makeResp.getSelectedButton() !== ui.Button.OK) return;
  const make = makeResp.getResponseText().trim();

  const modelResp = ui.prompt(
    'Add New Instrument',
    'Enter Model:',
    ui.ButtonSet.OK_CANCEL
  );
  if (modelResp.getSelectedButton() !== ui.Button.OK) return;
  const model = modelResp.getResponseText().trim();

  const serialResp = ui.prompt(
    'Add New Instrument',
    'Enter Serial Number:',
    ui.ButtonSet.OK_CANCEL
  );
  if (serialResp.getSelectedButton() !== ui.Button.OK) return;
  const serial = serialResp.getResponseText().trim();

  const conditionResp = ui.prompt(
    'Add New Instrument',
    'Enter Condition, for example Good, Fair, Needs Repair:',
    ui.ButtonSet.OK_CANCEL
  );
  if (conditionResp.getSelectedButton() !== ui.Button.OK) return;
  const condition = conditionResp.getResponseText().trim();

  const duplicates = findDuplicateSerials_(serial);
  if (duplicates.length) {
    const matching = duplicates.map(item => `${item.InstrumentID || 'Unknown ID'} ${item.InstrumentName || ''}`.trim()).join('\n');
    const choice = ui.alert(
      'Duplicate Serial Number Warning',
      `Serial ${serial} already exists for:\n${matching}\n\nAdd the instrument anyway?`,
      ui.ButtonSet.YES_NO
    );
    if (choice !== ui.Button.YES) return;
  }

  const result = addInstrumentToInventory_(instrumentName, make, model, serial, condition, true);
  syncValidQrLabelsSafely_();
  ui.alert(`Instrument added successfully.\nInstrument ID: ${result.Instrument_ID}`);
}

function addInstrumentToInventory_(instrumentName, make, model, serial, condition, allowDuplicateSerial) {
  return createInstrumentRecord_(instrumentName, make, model, serial, condition, allowDuplicateSerial === true);
}
function onEdit(e) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  if (SOURCE_INVENTORY_SHEET_ALIASES.indexOf(sh.getName()) === -1) return;

  const row = e.range.getRow();

  // Ignore title/header rows
  if (row < SOURCE_DATA_START_ROW) return;

  generateIdsForRow_(sh, row);
}

function generateIdsForRow_(sh, row) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const hm = getSourceHeaderMap_(sh);
    const uidColumn = findSourceColumn_(hm, ['Instrument_UID']);
    const idColumn = findSourceColumn_(hm, ['Instrument_ID']);
    const nameColumn = findSourceColumn_(hm, ['Instrument Name']);
    const makeColumn = sourceMakeColumn_(hm);
    const modelColumn = sourceModelColumn_(hm);
    if (!uidColumn || !idColumn || !nameColumn) return;

    const instrumentName = String(sh.getRange(row, nameColumn).getValue() || '').trim();
    if (!instrumentName) return;

    const uidCell = sh.getRange(row, uidColumn);
    if (!String(uidCell.getValue() || '').trim()) {
      uidCell.setNumberFormat('@').setValue(Utilities.getUuid());
    }

    const idCell = sh.getRange(row, idColumn);
    if (String(idCell.getValue() || '').trim()) return;

    const make = makeColumn ? String(sh.getRange(row, makeColumn).getValue() || '').trim() : '';
    const model = modelColumn ? String(sh.getRange(row, modelColumn).getValue() || '').trim() : '';
    const prefix = getInstrumentPrefix_(instrumentName, make, model);
    const lastRow = sh.getLastRow();
    const existingIds = sh.getRange(
      SOURCE_ID_SCAN_START_ROW,
      idColumn,
      Math.max(lastRow - SOURCE_ID_SCAN_START_ROW + 1, 1),
      1
    ).getValues();

    let highest = 0;
    const used = {};
    existingIds.forEach(valueRow => {
      const existing = String(valueRow[0] || '').trim();
      if (!existing) return;
      used[existing] = true;
      const match = existing.match(new RegExp(`^${prefix}-(\\d+)$`));
      if (match) highest = Math.max(highest, parseInt(match[1], 10));
    });

    let next = highest + 1;
    let instrumentId = `${prefix}-${String(next).padStart(3, '0')}`;
    while (used[instrumentId]) {
      next++;
      instrumentId = `${prefix}-${String(next).padStart(3, '0')}`;
    }
    idCell.setNumberFormat('@').setValue(instrumentId);
  } finally {
    lock.releaseLock();
  }

  syncValidQrLabelsSafely_();
}

function backfillInstrumentUidsAndIds() {
  const sh = getSourceInventorySheet_();

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 3) return; // no data rows

  // 🔥 HEADERS ARE ON ROW 2
  const headerRow = 2;
  const dataStartRow = SOURCE_DATA_START_ROW;

  const headers = sh.getRange(headerRow, 1, 1, lastCol)
                    .getValues()[0]
                    .map(h => (h + '').trim());

  const col = (name) => headers.indexOf(name);
  const colAny = (names) => {
    for (let i = 0; i < names.length; i++) {
      const index = col(names[i]);
      if (index !== -1) return index;
    }
    return -1;
  };

  const iUID = col('Instrument_UID');
  const iID  = col('Instrument_ID');
  const iInstrument = col('Instrument Name');
  const iMake = colAny(['Instrument Make', 'Make', 'Manufacturer', 'Brand']);
  const iModel = colAny(['Instrument Model', 'Model', 'Model Number', 'Model #']);

  if (iUID === -1 || iID === -1 || iInstrument === -1) {
    throw new Error('Missing required headers: Instrument_UID, Instrument_ID, Instrument Name');
  }

  const data = sh.getRange(dataStartRow, 1, lastRow - headerRow, lastCol).getValues();

  const used = new Set();
  const counters = {};

  // Scan existing IDs
  for (let r = 0; r < data.length; r++) {
    const existing = (data[r][iID] + '').trim();
    if (!existing) continue;
    used.add(existing);

    const m = existing.match(/^([A-Z0-9]+)-(\d+)$/);
    if (m) {
      const prefix = m[1];
      const n = parseInt(m[2], 10);
      counters[prefix] = Math.max(counters[prefix] || 0, n);
    }
  }

  const uidWrites = [];
  const idWrites = [];

  for (let r = 0; r < data.length; r++) {
    const rowIndex = dataStartRow + r;

    const uid = (data[r][iUID] + '').trim();
    const instId = (data[r][iID] + '').trim();

    const instrument = (data[r][iInstrument] + '').trim();
    const make = iMake !== -1 ? (data[r][iMake] + '').trim() : '';
    const model = iModel !== -1 ? (data[r][iModel] + '').trim() : '';

    // Never assign permanent identifiers to an empty source row.
    if (!instrument) continue;

    if (!uid) {
      uidWrites.push({ row: rowIndex, col: iUID + 1, value: Utilities.getUuid() });
    }

    if (!instId) {
      const prefix = getInstrumentPrefix_(instrument, make, model);
      let next = (counters[prefix] || 0) + 1;
      let newId = `${prefix}-${String(next).padStart(3, '0')}`;

      while (used.has(newId)) {
        next++;
        newId = `${prefix}-${String(next).padStart(3, '0')}`;
      }

      counters[prefix] = next;
      used.add(newId);
      idWrites.push({ row: rowIndex, col: iID + 1, value: newId });
    }
  }

  uidWrites.forEach(w => sh.getRange(w.row, w.col).setNumberFormat('@').setValue(w.value));
  idWrites.forEach(w => sh.getRange(w.row, w.col).setNumberFormat('@').setValue(w.value));
  syncValidQrLabelsSafely_();
}

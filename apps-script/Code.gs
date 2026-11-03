/**
 * RSVP backend for the invitation page. Bound to the event's Google Sheet.
 * Contract: docs/specs/api.md. Only this file touches guest data.
 *
 * Deploy: Execute as Me, Access Anyone. After editing, use
 * Deploy > Manage deployments > Edit > New version (saving alone does not update the live URL).
 */

const TZ = 'Asia/Kuala_Lumpur';
const TAB_CONFIG = 'Config';
const TAB_GUESTS = 'Guests';
const TAB_RESPONSES = 'Responses';

const GUEST_HEADERS = ['guest_id', 'name', 'pax_allowed', 'phone', 'link', 'whatsapp_link', 'reply', 'pax_confirmed', 'responded', 'first_opened', 'status'];
const OPENED_COL = 10; // column J: written by the script the first time a guest opens their personal link
const RESPONSE_HEADERS = ['timestamp', 'guest_id', 'name', 'attending', 'pax', 'note', 'source', 'verified', 'updated', 'phone'];
const PHONE_COL = 10; // column J: added later, so the script only uses it when its header is "phone" (see responseCols_)

// 32 characters, no I or O, no 0 or 1. 8 characters = 32^8 (about 1.1 trillion) combinations.
const ID_ALPHABET = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
const ID_LENGTH = 8;
const ID_PATTERN = /^[ABCDEFGHJKLMNPQRSTUVWXYZ2-9]{8}$/;

const MAX_NAME = 120;
const MAX_NOTE = 300;
const MAX_PHONE = 20;
const PHONE_PATTERN = /^[0-9+\-() ]+$/;
const MAX_CONTACT_URL = 500;
const MAX_PAX_CEILING = 20;
const DEFAULT_SHARED_MAX_PAX = 20;

const MSG_NOT_FOUND = 'Invitation not found';
const MSG_CLOSED = 'RSVP is closed';
const MSG_BUSY = 'Busy, please try again';
const MSG_INVALID = 'Invalid request';
const MSG_ERROR = 'Something went wrong';

// Admin page: key in Config "admin_key", checked here only. Sessions and the guess lockout live in CacheService.
const MIN_ADMIN_KEY = 24;
const ADMIN_KEY_LENGTH = 32;
const ADMIN_SESSION_SECONDS = 2 * 60 * 60;
const ADMIN_MAX_FAILS = 10;
const ADMIN_FAIL_WINDOW_SECONDS = 10 * 60;
const ADMIN_LOCK_SECONDS = 15 * 60;
const MSG_ADMIN_OFF = 'Admin is not set up. In the Sheet, run RSVP tools > Generate admin key';
const MSG_ADMIN_DENIED = 'Wrong admin key';
const MSG_ADMIN_LOCKED = 'Too many wrong keys. Try again in 15 minutes';
const MSG_ADMIN_SESSION = 'Session expired. Please log in again';
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

/* ------------------------------------------------------------------ */
/* Sheet menu                                                          */
/* ------------------------------------------------------------------ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('RSVP tools')
    .addItem('Set up tabs (first time)', 'setupSheet')
    .addItem('Generate IDs', 'generateIds')
    .addItem('Check setup', 'selfTest')
    .addItem('Generate admin key', 'generateAdminKey')
    .addSeparator()
    .addItem('Clear test data (keeps guests)', 'clearTestData')
    .addItem('Reset for a new event (deletes guests too)', 'resetForNewEvent')
    .addToUi();
}

/**
 * First-time setup: creates the Config, Guests and Responses tabs with the right headers,
 * placeholder Config values and three FAKE test guests. Never touches a tab that already exists.
 */
function setupSheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const created = [];

  const ensure = function (name, headers, rows) {
    if (ss.getSheetByName(name)) return;
    const sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    if (rows && rows.length) sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    created.push(name);
  };

  ensure(TAB_CONFIG, ['key', 'value'], [
    ['event_title', 'Garden Dinner Celebration'],
    ['rsvp_deadline', '2026-12-10'],
    ['site_url', 'https://example.pages.dev'],
    ['shared_max_pax', 20],
    ['msg_vip', 'Dear {name}, you are warmly invited to {event}. Please open your personal invitation and RSVP by {deadline}: {link}'],
    ['msg_shared', 'You are warmly invited to {event}. Open the invitation and RSVP by {deadline}: {link}'],
    ['rsvp_open', true],
    ['contact_link', '']
  ]);
  ensure(TAB_GUESTS, GUEST_HEADERS, [
    ['', 'Test Guest One', 2, ''],
    ['', 'Test Guest Two', 3, ''],
    ['', 'Test Guest Three', 1, '']
  ]);
  ensure(TAB_RESPONSES, RESPONSE_HEADERS, null);

  // Remove the empty default tab Google creates, if it is still untouched.
  const blank = ss.getSheetByName('Sheet1');
  if (blank && created.length && blank.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(blank);

  ui.alert(
    created.length
      ? 'Created: ' + created.join(', ') + '.\n\nNext: change site_url in Config to your real page address, then run RSVP tools > Generate IDs.'
      : 'Nothing to do: all three tabs already exist. Use "Check setup" to verify them.'
  );
}

/** Fills blank guest_id cells (rows that have a name) and the formula columns E to H. Never changes existing IDs. */
function generateIds() {
  const ui = SpreadsheetApp.getUi();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    ui.alert(MSG_BUSY);
    return;
  }
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_GUESTS);
    if (!sheet) {
      ui.alert('Tab "' + TAB_GUESTS + '" not found. Run "Check setup".');
      return;
    }
    const last = lastDataRow_(sheet, 2); // by name column
    if (last < 2) {
      ui.alert('No guests found. Type names in column B first.');
      return;
    }
    const rows = last - 1;
    const data = sheet.getRange(2, 1, rows, 4).getValues(); // A to D
    const used = {};
    data.forEach(function (r) {
      const id = String(r[0]).trim().toUpperCase();
      if (id) used[id] = true;
    });

    // Columns I to K were added later: write their headers if an older Sheet lacks them.
    const extraHeaders = sheet.getRange(1, 9, 1, 3);
    const currentExtra = extraHeaders.getValues()[0];
    extraHeaders.setValues([currentExtra.map(function (v, k) { return String(v).trim() === '' ? GUEST_HEADERS[8 + k] : v; })]);

    const ids = [];
    let created = 0;
    const problems = [];
    const seen = {};
    data.forEach(function (r, i) {
      const name = String(r[1]).trim();
      let id = String(r[0]).trim();
      if (name && id) {
        if (!ID_PATTERN.test(id.toUpperCase())) {
          problems.push('Row ' + (i + 2) + ': guest_id "' + id + '" is not valid (it must be exactly 8 characters made by this tool). Clear the cell and run Generate IDs again.');
        } else if (seen[id.toUpperCase()]) {
          problems.push('Row ' + (i + 2) + ': guest_id "' + id + '" is already used in row ' + seen[id.toUpperCase()] + '. Clear this cell and run Generate IDs again.');
        } else {
          seen[id.toUpperCase()] = i + 2;
        }
      }
      if (name && !id) {
        id = newUniqueId_(used);
        used[id] = true;
        seen[id] = i + 2;
        created++;
      }
      if (name) {
        const pax = Number(r[2]);
        if (!Number.isInteger(pax) || pax < 1 || pax > MAX_PAX_CEILING) {
          problems.push('Row ' + (i + 2) + ': pax_allowed must be a whole number from 1 to ' + MAX_PAX_CEILING);
        }
      }
      ids.push([id]);
    });
    sheet.getRange(2, 1, rows, 1).setValues(ids);

    // Formulas for E to I and K (J is written by the script), only where the cell is still empty
    // (never overwrite a host's own formula or typed value).
    fillFormulas_(sheet, rows, data, 5, 5, function (r) { return formulasForRow_(r).slice(0, 5); });
    fillFormulas_(sheet, rows, data, 11, 1, function (r) { return [formulasForRow_(r)[5]]; });

    let msg = 'Created ' + created + ' new ID(s). Existing IDs were not changed.';
    if (problems.length) msg += '\n\nPlease fix:\n' + problems.join('\n');
    ui.alert(msg);
  } finally {
    lock.releaseLock();
  }
}

function fillFormulas_(sheet, rows, data, firstCol, nCols, freshFor) {
  const range = sheet.getRange(2, firstCol, rows, nCols);
  const existingFormulas = range.getFormulas();
  const existingValues = range.getValues();
  const merged = existingFormulas.map(function (fRow, i) {
    const hasName = String(data[i][1]).trim() !== '';
    const fresh = hasName ? freshFor(i + 2) : null;
    return fRow.map(function (f, c) {
      if (f) return f;
      if (existingValues[i][c] !== '') return existingValues[i][c];
      return fresh ? fresh[c] : '';
    });
  });
  range.setFormulas(merged);
}

function formulasForRow_(r) {
  const cfg = function (key) { return 'VLOOKUP("' + key + '",Config!A:B,2,FALSE)'; };
  const link = '=IF(A' + r + '="","",' + cfg('site_url') + '&"/?id="&A' + r + ')';
  const msg =
    'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(' + cfg('msg_vip') +
    ',"{name}",B' + r + '),"{event}",' + cfg('event_title') +
    '),"{deadline}",' + ordinalDateFormula_(cfg('rsvp_deadline')) + '),"{link}",E' + r + ')';
  const wa = '=IF(OR(A' + r + '="",D' + r + '=""),"","https://wa.me/"&D' + r + '&"?text="&ENCODEURL(' + msg + '))';
  const reply = '=IF(A' + r + '="","",IFERROR(INDEX(Responses!D:D,MATCH(A' + r + ',Responses!B:B,0)),""))';
  const paxConfirmed = '=IF(A' + r + '="","",IFERROR(INDEX(Responses!E:E,MATCH(A' + r + ',Responses!B:B,0)),""))';
  const responded = '=IF(A' + r + '="","",COUNTIF(Responses!B:B,A' + r + ')>0)';
  const status = '=IF(A' + r + '="","",IF(I' + r + '=TRUE,"Replied: "&G' + r + ',IF(J' + r + '<>"","Opened, no reply","Not opened")))';
  return [link, wa, reply, paxConfirmed, responded, status];
}

/** Sheets formula for a date cell as "6th November 2026", matching longDate_. */
function ordinalDateFormula_(cell) {
  const day = 'DAY(' + cell + ')';
  return day + '&IF(AND(' + day + '>=11,' + day + '<=13),"th",CHOOSE(MOD(' + day + ',10)+1,"th","st","nd","rd","th","th","th","th","th","th"))&TEXT(' + cell + '," mmmm yyyy")';
}

function newUniqueId_(used) {
  for (let attempt = 0; attempt < 50; attempt++) {
    const id = randomId_();
    if (!used[id]) return id;
  }
  throw new Error('E_ID_GEN');
}

/** 16 hex characters of a random UUID give 8 bytes; a byte mod 32 is uniform because 256 is a multiple of 32. */
function randomId_() {
  const hex = Utilities.getUuid().replace(/-/g, '');
  let id = '';
  for (let i = 0; i < ID_LENGTH; i++) {
    id += ID_ALPHABET.charAt(parseInt(hex.substr(i * 2, 2), 16) % ID_ALPHABET.length);
  }
  return id;
}

/* ------------------------------------------------------------------ */
/* Fresh start (soft delete: everything is archived to new tabs first) */
/* ------------------------------------------------------------------ */

/**
 * Archives, then clears, every reply in Responses and every first_opened time in Guests. Guest names, pax,
 * IDs and Config are kept, so personal links keep working. Use after testing with your own guests.
 */
function clearTestData() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    'Clear test data',
    'This moves ALL replies (Responses tab) and all "first_opened" times (Guests tab) into new tabs named "Archive ...", ' +
      'then clears them here.\n\nGuest names, IDs and the Config tab are kept. Nothing is lost: the archive tabs hold a copy.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) {
    ui.alert('Cancelled. Nothing was changed.');
    return;
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    ui.alert(MSG_BUSY);
    return;
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let archived;
    try {
      archived = archiveBeforeDelete_(ss, 'clear');
    } catch (err) {
      logError_('archive', err);
      ui.alert('The archive copy failed, so NOTHING was cleared. Please try again.');
      return;
    }
    const removed = clearResponses_(ss);
    clearOpened_(ss);
    ui.alert(
      'Done. Cleared ' + removed + ' reply row(s) and all first_opened times.\n\n' +
        (archived.length ? 'Archived copy: ' + archived.join(', ') + '.' : 'There was nothing to archive.') +
        '\n\nGuests and Config were kept.'
    );
  } finally {
    lock.releaseLock();
  }
}

/**
 * Starts a new event on this same Sheet: archives, then clears, all replies AND the whole guest list
 * (names, pax, phone, IDs, first_opened). Config is kept. Old personal links stop working at once.
 * The host must type RESET.
 */
function resetForNewEvent() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Reset for a new event',
    'This moves ALL replies AND the ENTIRE guest list (names, guest counts, phones and IDs) into new tabs named "Archive ...", ' +
      'then clears them here. Every personal link made so far will stop working.\n\n' +
      'The Config tab is kept: remember to update it for the new event. Nothing is lost: the archive tabs hold a copy.\n\n' +
      'Type RESET to continue:',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK || String(response.getResponseText()).trim() !== 'RESET') {
    ui.alert('Cancelled. Nothing was changed.');
    return;
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    ui.alert(MSG_BUSY);
    return;
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let archived;
    try {
      archived = archiveBeforeDelete_(ss, 'reset');
    } catch (err) {
      logError_('archive', err);
      ui.alert('The archive copy failed, so NOTHING was cleared. Please try again.');
      return;
    }
    const removed = clearResponses_(ss);
    clearOpened_(ss);
    const guests = ss.getSheetByName(TAB_GUESTS);
    let guestRows = 0;
    if (guests) {
      guestRows = Math.max(0, lastDataRow_(guests, 2) - 1);
      const max = guests.getMaxRows();
      if (max > 1) guests.getRange(2, 1, max - 1, 4).clearContent(); // A to D: id, name, pax, phone
    }
    ui.alert(
      'Done. Cleared ' + removed + ' reply row(s) and ' + guestRows + ' guest row(s). Config was kept.\n\n' +
        (archived.length ? 'Archived copy: ' + archived.join(', ') + '.' : 'There was nothing to archive.') +
        '\n\nNext: update the Config tab, type the new guests (columns B to D), then run RSVP tools > Generate IDs.'
    );
  } finally {
    lock.releaseLock();
  }
}

/**
 * Copies the data that is about to be cleared into NEW tabs, checks the copy, and returns the tab names.
 * Throws if a copy cannot be verified (the caller then deletes nothing). A half-made archive tab is removed.
 * mode 'clear': replies, plus guests that have a first_opened time. mode 'reset': replies plus every guest row.
 * Archive tabs keep the original column order so a copy-paste restores them.
 */
function archiveBeforeDelete_(ss, mode) {
  const stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH.mm');
  const made = [];
  try {
    const responses = ss.getSheetByName(TAB_RESPONSES);
    if (responses) {
      const last = lastDataRow_(responses, 1);
      if (last >= 2) {
        const values = responses.getRange(1, 1, last, responseCols_(responses)).getValues();
        made.push(writeArchiveTab_(ss, 'Archive ' + stamp + ' Responses', values, [3, 6]));
      }
    }

    const guests = ss.getSheetByName(TAB_GUESTS);
    if (guests) {
      const lastGuest = lastDataRow_(guests, 2);
      if (lastGuest >= 2) {
        const hasOpened = String(guests.getRange(1, OPENED_COL).getValue()).trim() === 'first_opened';
        const rows = guests.getRange(2, 1, lastGuest - 1, OPENED_COL).getValues();
        const snapshot = [['guest_id', 'name', 'pax_allowed', 'phone', 'first_opened']];
        rows.forEach(function (r) {
          const opened = hasOpened ? r[OPENED_COL - 1] : '';
          const anything = r[0] !== '' || r[1] !== '' || r[2] !== '' || r[3] !== '';
          if (mode === 'reset' ? anything : opened !== '') snapshot.push([r[0], r[1], r[2], r[3], opened]);
        });
        if (snapshot.length > 1) made.push(writeArchiveTab_(ss, 'Archive ' + stamp + ' Guests', snapshot, [2]));
      }
    }
  } catch (err) {
    made.forEach(function (name) {
      const sheet = ss.getSheetByName(name);
      if (sheet) ss.deleteSheet(sheet);
    });
    throw err;
  }
  return made;
}

/** Creates a new tab with the values (text format on the given 1-based columns), verifies it, returns its name. */
function writeArchiveTab_(ss, baseName, values, textCols) {
  let name = baseName;
  for (let n = 2; ss.getSheetByName(name); n++) name = baseName + ' (' + n + ')';
  const sheet = ss.insertSheet(name);
  try {
    // Plain-text format on name/note columns so archived text can never run as a formula.
    textCols.forEach(function (c) { sheet.getRange(1, c, values.length, 1).setNumberFormat('@'); });
    sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
    sheet.setFrozenRows(1);
    const check = sheet.getRange(1, 1, values.length, 1).getValues();
    for (let i = 0; i < values.length; i++) {
      if (String(check[i][0]) !== String(values[i][0])) throw new Error('E_ARCHIVE_VERIFY');
    }
  } catch (err) {
    ss.deleteSheet(sheet);
    throw err;
  }
  return name;
}

/** Clears Responses columns A to I (and J when it is the phone column) below the header. Returns how many reply rows there were. */
function clearResponses_(ss) {
  const sheet = ss.getSheetByName(TAB_RESPONSES);
  if (!sheet) return 0;
  const count = Math.max(0, lastDataRow_(sheet, 1) - 1);
  const max = sheet.getMaxRows();
  if (max > 1) sheet.getRange(2, 1, max - 1, responseCols_(sheet)).clearContent();
  return count;
}

/** Clears the first_opened column in Guests, only when that column really is first_opened. */
function clearOpened_(ss) {
  const sheet = ss.getSheetByName(TAB_GUESTS);
  if (!sheet) return;
  if (String(sheet.getRange(1, OPENED_COL).getValue()).trim() !== 'first_opened') return;
  const max = sheet.getMaxRows();
  if (max > 1) sheet.getRange(2, OPENED_COL, max - 1, 1).clearContent();
}

/** Shows what is wrong with the Sheet setup. */
function selfTest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const problems = [];

  const checkHeaders = function (tabName, expected) {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) {
      problems.push('Tab "' + tabName + '" is missing.');
      return null;
    }
    const header = sheet.getRange(1, 1, 1, expected.length).getValues()[0];
    expected.forEach(function (h, i) {
      if (String(header[i]).trim() !== h) {
        problems.push(tabName + ' column ' + (i + 1) + ' should be "' + h + '" but is "' + header[i] + '".');
      }
    });
    return sheet;
  };

  const cfgSheet = ss.getSheetByName(TAB_CONFIG);
  if (!cfgSheet) problems.push('Tab "' + TAB_CONFIG + '" is missing.');
  checkHeaders(TAB_GUESTS, GUEST_HEADERS);
  checkHeaders(TAB_RESPONSES, RESPONSE_HEADERS);

  if (cfgSheet) {
    const cfg = readConfig_();
    ['event_title', 'rsvp_deadline', 'site_url', 'shared_max_pax', 'msg_vip', 'msg_shared', 'rsvp_open'].forEach(function (k) {
      if (cfg[k] === undefined || cfg[k] === '') problems.push('Config "' + k + '" is missing or empty.');
    });
    if (!deadlineString_(cfg.rsvp_deadline)) problems.push('Config "rsvp_deadline" must be a date like 2026-12-10. Until fixed, RSVPs are closed.');
    const sm = Number(cfg.shared_max_pax);
    if (!Number.isInteger(sm) || sm < 1 || sm > MAX_PAX_CEILING) problems.push('Config "shared_max_pax" must be a whole number from 1 to ' + MAX_PAX_CEILING + '.');
    if (/\/$/.test(String(cfg.site_url || ''))) problems.push('Config "site_url" must not end with a slash.');
    if (parseBool_(cfg.rsvp_open) === null) problems.push('Config "rsvp_open" must be TRUE or FALSE.');
    const adminKey = String(cfg.admin_key === undefined || cfg.admin_key === null ? '' : cfg.admin_key).trim();
    if (adminKey && adminKey.length < MIN_ADMIN_KEY) problems.push('Config "admin_key" is too short to be safe. Run RSVP tools > Generate admin key.');
    const contact = String(cfg.contact_link === undefined || cfg.contact_link === null ? '' : cfg.contact_link).trim();
    if (contact && !contactUrl_(cfg)) problems.push('Config "contact_link" must start with https:// (for example https://wa.link/abc123), or be left empty.');
  }

  SpreadsheetApp.getUi().alert(problems.length ? 'Please fix:\n\n' + problems.join('\n') : 'Setup looks good.');
}

/* ------------------------------------------------------------------ */
/* Web app                                                             */
/* ------------------------------------------------------------------ */

function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    if (p.action === 'guest') return handleGuest_(p.id);
    if (p.action === 'config') return handleConfig_();
    return fail_(MSG_INVALID);
  } catch (err) {
    logError_('doGet', err);
    return fail_(MSG_ERROR);
  }
}

function doPost(e) {
  try {
    const body = parseBody_(e);
    if (body && body.action === 'admin') return handleAdmin_(body);
    return handleRsvp_(e);
  } catch (err) {
    logError_('doPost', err);
    return fail_(MSG_ERROR);
  }
}

function handleConfig_() {
  const cfg = readConfig_();
  return ok_('OK', {
    deadline: deadlineString_(cfg.rsvp_deadline) || '',
    open: isOpen_(cfg),
    sharedMaxPax: sharedMaxPax_(cfg),
    contactUrl: contactUrl_(cfg)
  });
}

function handleGuest_(rawId) {
  const id = normaliseId_(rawId);
  if (!id) return fail_(MSG_NOT_FOUND);
  const guest = findGuest_(id);
  if (!guest) return fail_(MSG_NOT_FOUND);

  recordOpen_(guest);

  const cfg = readConfig_();
  const reply = findResponse_(id);
  return ok_('OK', {
    name: guest.name,
    pax: guest.pax,
    replied: reply ? { attending: reply.attending, pax: reply.pax } : null,
    deadline: deadlineString_(cfg.rsvp_deadline) || '',
    open: isOpen_(cfg),
    contactUrl: contactUrl_(cfg)
  });
}

/**
 * Stamps first_opened (column J) the first time a guest's personal link loads. Best effort: it never
 * blocks or fails the lookup, and is skipped on an older Sheet that has no first_opened column.
 */
function recordOpen_(guest) {
  if (guest.opened !== '') return;
  try {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(3000)) return;
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_GUESTS);
      if (String(sheet.getRange(1, OPENED_COL).getValue()).trim() !== 'first_opened') return;
      const cell = sheet.getRange(guest.row, OPENED_COL);
      if (cell.getValue() === '') cell.setValue(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'));
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    logError_('recordOpen', err);
  }
}

function handleRsvp_(e) {
  // 1. Parse
  let body;
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '');
  } catch (parseErr) {
    return fail_(MSG_INVALID);
  }
  if (!body || typeof body !== 'object' || Array.isArray(body) || body.action !== 'rsvp') return fail_(MSG_INVALID);

  // 2. Closed?
  const cfg = readConfig_();
  if (!isOpen_(cfg)) return fail_(MSG_CLOSED);

  // 3. Honeypot: pretend success, write nothing
  if (body.website !== undefined && body.website !== null && String(body.website).trim() !== '') {
    const fakeAttending = body.attending === 'yes' ? 'yes' : 'no';
    return ok_('RSVP saved', { attending: fakeAttending, pax: fakeAttending === 'yes' ? 1 : 0 });
  }

  // 4. Identity and pax cap (always from the Sheet, never from the client)
  let name;
  let source;
  let guestId = '';
  let paxMax;
  if (body.id !== undefined && body.id !== null && body.id !== '') {
    const id = normaliseId_(body.id);
    if (!id) return fail_(MSG_NOT_FOUND);
    const guest = findGuest_(id);
    if (!guest) return fail_(MSG_NOT_FOUND);
    guestId = id;
    name = guest.name;
    paxMax = guest.pax;
    source = 'vip';
  } else {
    if (typeof body.name !== 'string') return fail_('Please enter your name');
    const trimmed = body.name.trim();
    if (trimmed.length < 1 || trimmed.length > MAX_NAME) return fail_('Please enter your name');
    name = trimmed;
    paxMax = sharedMaxPax_(cfg);
    source = 'shared';
  }

  // 5. Attending and pax
  if (body.attending !== 'yes' && body.attending !== 'no') return fail_('Please choose yes or no');
  let pax = 0;
  if (body.attending === 'yes') {
    if (typeof body.pax !== 'number' || !Number.isInteger(body.pax) || body.pax < 1 || body.pax > paxMax) {
      return fail_('Number of guests must be between 1 and ' + paxMax);
    }
    pax = body.pax;
  }

  // 6. Note
  let note = '';
  if (body.note !== undefined && body.note !== null) {
    if (typeof body.note !== 'string') return fail_(MSG_INVALID);
    note = body.note.trim();
    if (note.length > MAX_NOTE) return fail_('Note must be at most ' + MAX_NOTE + ' characters');
  }

  // 6b. Phone: optional, shared links only (a VIP's phone is already in Guests, so it is ignored for them)
  let phone = '';
  if (source === 'shared' && body.phone !== undefined && body.phone !== null) {
    if (typeof body.phone !== 'string') return fail_(MSG_INVALID);
    phone = body.phone.trim();
    const digits = phone.replace(/\D/g, '').length;
    if (phone && (phone.length > MAX_PHONE || !PHONE_PATTERN.test(phone) || digits < 8 || digits > 15)) {
      return fail_('Please enter a valid phone number');
    }
  }

  // 7. Sanitise, lock, write
  const safeName = sanitise_(name);
  const safeNote = sanitise_(note);
  const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return fail_(MSG_BUSY);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_RESPONSES);
    if (!sheet) throw new Error('E_NO_RESPONSES_TAB');

    const existing = source === 'vip' ? findResponse_(guestId) : null;
    if (existing) {
      // VIP changes their reply: update in place (D attending, E pax, F note, I updated)
      const textRange = sheet.getRange(existing.row, 4, 1, 3);
      sheet.getRange(existing.row, 6).setNumberFormat('@');
      textRange.setValues([[body.attending, pax, safeNote]]);
      sheet.getRange(existing.row, 9).setValue(now);
    } else {
      const row = lastDataRow_(sheet, 1) + 1;
      const values = [now, guestId, safeName, body.attending, pax, safeNote, source, source === 'vip', ''];
      // Plain-text format on name and note stops Sheets reading them as formulas.
      sheet.getRange(row, 3).setNumberFormat('@');
      sheet.getRange(row, 6).setNumberFormat('@');
      if (responseCols_(sheet) === PHONE_COL) {
        // Plain text keeps a leading + or 0. PHONE_PATTERN allows no letters, so it cannot hold a formula.
        sheet.getRange(row, PHONE_COL).setNumberFormat('@');
        values.push(phone);
      }
      sheet.getRange(row, 1, 1, values.length).setValues([values]);
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  // 8. Done
  return ok_('RSVP saved', { attending: body.attending, pax: pax });
}

/* ------------------------------------------------------------------ */
/* Admin page (POST action "admin")                                    */
/* ------------------------------------------------------------------ */

/** Menu: makes a new random admin key, saves it in Config, and shows it once. Logs out every open admin page. */
function generateAdminKey() {
  const ui = SpreadsheetApp.getUi();
  const cfg = readConfig_();
  if (String(cfg.admin_key || '').trim()) {
    const answer = ui.alert('Generate admin key', 'This replaces the current admin key. Anyone logged in to the admin page is logged out.\n\nContinue?', ui.ButtonSet.YES_NO);
    if (answer !== ui.Button.YES) {
      ui.alert('Cancelled. The key was not changed.');
      return;
    }
  }
  const key = randomKey_();
  writeConfigValue_('admin_key', key);
  ui.alert('New admin key (also saved in Config > admin_key):\n\n' + key + '\n\nKeep it private. Anyone with this key can see and change the guest list.');
}

/** 32 characters from 3 UUIDs, using only their random hex digits (not the version or variant digits). */
function randomKey_() {
  let hex = '';
  while (hex.length < ADMIN_KEY_LENGTH * 2) {
    const u = Utilities.getUuid().replace(/-/g, '');
    hex += u.substr(0, 12) + u.substr(20, 12);
  }
  let key = '';
  for (let i = 0; i < ADMIN_KEY_LENGTH; i++) key += ID_ALPHABET.charAt(parseInt(hex.substr(i * 2, 2), 16) % ID_ALPHABET.length);
  return key;
}

function parseBody_(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '');
    return body && typeof body === 'object' && !Array.isArray(body) ? body : null;
  } catch (err) {
    return null;
  }
}

function handleAdmin_(body) {
  if (body.op === 'login') return adminLogin_(body.key);
  if (!adminSessionValid_(body.token)) return fail_(MSG_ADMIN_SESSION);
  if (body.op === 'overview') return ok_('OK', adminOverview_());
  if (body.op === 'addGuest') return adminAddGuest_(body);
  if (body.op === 'updateGuest') return adminUpdateGuest_(body);
  if (body.op === 'setConfig') return adminSetConfig_(body);
  return fail_(MSG_INVALID);
}

function adminKey_() {
  const key = String(readConfig_().admin_key || '').trim();
  return key.length >= MIN_ADMIN_KEY ? key : '';
}

/** Compares every character so the time taken does not reveal how much of a guess was right. */
function sameText_(a, b) {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

function adminLogin_(rawKey) {
  const cache = CacheService.getScriptCache();
  if (cache.get('admin_locked')) return fail_(MSG_ADMIN_LOCKED);
  const key = adminKey_();
  if (!key) return fail_(MSG_ADMIN_OFF);
  if (typeof rawKey !== 'string' || !sameText_(rawKey.trim(), key)) {
    const fails = Number(cache.get('admin_fails') || 0) + 1;
    cache.put('admin_fails', String(fails), ADMIN_FAIL_WINDOW_SECONDS);
    if (fails >= ADMIN_MAX_FAILS) {
      cache.put('admin_locked', '1', ADMIN_LOCK_SECONDS);
      cache.remove('admin_fails');
      return fail_(MSG_ADMIN_LOCKED);
    }
    return fail_(MSG_ADMIN_DENIED);
  }
  cache.remove('admin_fails');
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  // The session remembers which key it was made with, so a new key logs everyone out.
  cache.put('admin_s_' + token, key, ADMIN_SESSION_SECONDS);
  return ok_('OK', { token: token, expiresInSeconds: ADMIN_SESSION_SECONDS });
}

function adminSessionValid_(token) {
  if (typeof token !== 'string' || !/^[0-9a-f]{64}$/.test(token)) return false;
  const key = adminKey_();
  const stored = CacheService.getScriptCache().get('admin_s_' + token);
  return Boolean(key && stored && sameText_(stored, key));
}

/** Everything the admin page shows, in one call. Never includes admin_key. */
function adminOverview_() {
  const cfg = readConfig_();
  const siteUrl = String(cfg.site_url || '').trim().replace(/\/+$/, '');
  const deadline = deadlineString_(cfg.rsvp_deadline) || '';
  const fill = function (template, name, link) {
    return String(template || '')
      .split('{name}').join(name)
      .split('{event}').join(String(cfg.event_title || ''))
      .split('{deadline}').join(longDate_(deadline))
      .split('{link}').join(link);
  };

  const responses = readResponses_();
  const latest = {};
  responses.forEach(function (r) { if (r.guestId) latest[r.guestId] = r; });

  const guests = readGuests_().map(function (g) {
    const link = siteUrl && g.id ? siteUrl + '/?id=' + g.id : '';
    const reply = latest[g.id] || null;
    const message = link ? fill(cfg.msg_vip, g.name, link) : '';
    return {
      id: g.id,
      name: g.name,
      pax: g.pax,
      phone: g.phone,
      link: link,
      // With a phone, WhatsApp opens that chat; without one, it lets you pick the contact.
      whatsappUrl: message ? 'https://wa.me/' + g.phone + '?text=' + encodeURIComponent(message) : '',
      firstOpened: g.opened,
      reply: reply ? { attending: reply.attending, pax: reply.pax, note: reply.note, updated: reply.updated || reply.timestamp } : null,
      status: reply ? 'replied' : (g.opened ? 'opened' : 'not_opened')
    };
  });

  const yes = responses.filter(function (r) { return r.attending === 'yes'; });
  const sharedMessage = siteUrl ? fill(cfg.msg_shared, '', siteUrl) : '';
  return {
    settings: {
      eventTitle: String(cfg.event_title || ''),
      deadline: deadline,
      rsvpOpen: parseBool_(cfg.rsvp_open) === true,
      open: isOpen_(cfg),
      sharedMaxPax: sharedMaxPax_(cfg),
      siteUrl: siteUrl,
      sharedLink: siteUrl,
      sharedWhatsappUrl: sharedMessage ? 'https://wa.me/?text=' + encodeURIComponent(sharedMessage) : ''
    },
    summary: {
      invited: guests.length,
      replied: guests.filter(function (g) { return g.status === 'replied'; }).length,
      notOpened: guests.filter(function (g) { return g.status === 'not_opened'; }).length,
      attendingReplies: yes.length,
      declinedReplies: responses.length - yes.length,
      peopleComing: yes.reduce(function (sum, r) { return sum + r.pax; }, 0),
      sharedReplies: responses.filter(function (r) { return r.source === 'shared'; }).length
    },
    guests: guests,
    responses: responses
  };
}

/** Checks name, pax and phone typed on the admin page. Returns an error message or the clean values. */
function cleanGuestInput_(body) {
  if (typeof body.name !== 'string') return { error: 'Please enter a name' };
  const name = body.name.trim();
  if (name.length < 1 || name.length > MAX_NAME) return { error: 'Name must be 1 to ' + MAX_NAME + ' characters' };
  if (!Number.isInteger(body.pax) || body.pax < 1 || body.pax > MAX_PAX_CEILING) return { error: 'Guests allowed must be a whole number from 1 to ' + MAX_PAX_CEILING };
  let phone = '';
  if (body.phone !== undefined && body.phone !== null && body.phone !== '') {
    if (typeof body.phone !== 'string') return { error: MSG_INVALID };
    phone = body.phone.replace(/[\s\-()+]/g, '');
    if (!/^\d{8,15}$/.test(phone) || phone.charAt(0) === '0') return { error: 'Phone must be digits with the country code, like 60123456789' };
  }
  return { name: sanitise_(name), pax: body.pax, phone: phone };
}

function adminAddGuest_(body) {
  const input = cleanGuestInput_(body);
  if (input.error) return fail_(input.error);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return fail_(MSG_BUSY);
  let id;
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_GUESTS);
    if (!sheet) throw new Error('E_NO_GUESTS_TAB');
    const used = {};
    readGuests_().forEach(function (g) { used[g.id] = true; });
    id = newUniqueId_(used);
    const row = lastDataRow_(sheet, 2) + 1;
    // Plain text keeps the name from being read as a formula and the phone's digits exactly as typed.
    sheet.getRange(row, 2).setNumberFormat('@');
    sheet.getRange(row, 4).setNumberFormat('@');
    sheet.getRange(row, 1, 1, 4).setValues([[id, input.name, input.pax, input.phone]]);
    const formulas = formulasForRow_(row);
    sheet.getRange(row, 5, 1, 5).setFormulas([formulas.slice(0, 5)]);
    sheet.getRange(row, 11, 1, 1).setFormulas([[formulas[5]]]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return ok_('Guest added', { id: id });
}

function adminUpdateGuest_(body) {
  const id = normaliseId_(body.id);
  if (!id) return fail_(MSG_NOT_FOUND);
  const input = cleanGuestInput_(body);
  if (input.error) return fail_(input.error);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return fail_(MSG_BUSY);
  try {
    const guest = readGuests_().filter(function (g) { return g.id === id; })[0];
    if (!guest) return fail_(MSG_NOT_FOUND);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_GUESTS);
    sheet.getRange(guest.row, 2).setNumberFormat('@');
    sheet.getRange(guest.row, 4).setNumberFormat('@');
    sheet.getRange(guest.row, 2, 1, 3).setValues([[input.name, input.pax, input.phone]]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return ok_('Guest updated', { id: id });
}

/** Changes rsvp_open and/or rsvp_deadline. Anything else in Config stays a Sheet-only setting. */
function adminSetConfig_(body) {
  const changes = [];
  if (body.rsvpOpen !== undefined) {
    if (typeof body.rsvpOpen !== 'boolean') return fail_(MSG_INVALID);
    changes.push(['rsvp_open', body.rsvpOpen]);
  }
  if (body.deadline !== undefined) {
    if (!validDate_(body.deadline)) return fail_('Deadline must be a real date like 2026-11-06');
    changes.push(['rsvp_deadline', body.deadline]);
  }
  if (!changes.length) return fail_(MSG_INVALID);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return fail_(MSG_BUSY);
  try {
    changes.forEach(function (c) { writeConfigValue_(c[0], c[1]); });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return ok_('Settings saved', {});
}

function validDate_(s) {
  if (typeof s !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
  const d = new Date(s + 'T00:00:00Z');
  return !isNaN(d.getTime()) && d.toISOString().slice(0, 10) === s;
}

/** "2026-11-06" -> "6th November 2026" (same as the WhatsApp formula in Guests). */
function longDate_(ymd) {
  if (!ymd) return '';
  const p = ymd.split('-');
  return ordinal_(Number(p[2])) + ' ' + MONTHS[Number(p[1]) - 1] + ' ' + p[0];
}

/** 1 -> "1st", 2 -> "2nd", 3 -> "3rd", 11 to 13 -> "th", 21 -> "21st". */
function ordinal_(n) {
  if (n % 100 >= 11 && n % 100 <= 13) return n + 'th';
  return n + (['th', 'st', 'nd', 'rd'][n % 10] || 'th');
}

/** Writes one Config value, adding the key row if it is missing. */
function writeConfigValue_(key, value) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_CONFIG);
  if (!sheet) throw new Error('E_NO_CONFIG_TAB');
  const last = Math.max(sheet.getLastRow(), 1);
  const keys = sheet.getRange(1, 1, last, 1).getValues();
  for (let i = 1; i < keys.length; i++) {
    if (String(keys[i][0]).trim() === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.getRange(last + 1, 1, 1, 2).setValues([[key, value]]);
}

/* ------------------------------------------------------------------ */
/* Sheet reads                                                         */
/* ------------------------------------------------------------------ */

/** Sheets may turn a written "2026-11-02 10:00:00" into a date cell; give it back as text either way. */
function cellText_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm:ss');
  return v === undefined || v === null ? '' : String(v).trim();
}

/** Every guest row that has a valid ID and a name. Rows without an ID yet are skipped (run Generate IDs). */
function readGuests_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_GUESTS);
  if (!sheet) throw new Error('E_NO_GUESTS_TAB');
  const last = lastDataRow_(sheet, 2);
  if (last < 2) return [];
  const out = [];
  sheet.getRange(2, 1, last - 1, OPENED_COL).getValues().forEach(function (r, i) {
    const id = String(r[0]).trim().toUpperCase();
    const name = String(r[1]).trim();
    if (!ID_PATTERN.test(id) || !name) return;
    const pax = Number(r[2]);
    out.push({
      row: i + 2,
      id: id,
      name: name,
      pax: Number.isInteger(pax) && pax >= 1 ? Math.min(pax, MAX_PAX_CEILING) : 1,
      phone: String(r[3]).replace(/\D/g, ''),
      opened: cellText_(r[OPENED_COL - 1])
    });
  });
  return out;
}

/** Every reply row in Responses, oldest first. */
function readResponses_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_RESPONSES);
  if (!sheet) throw new Error('E_NO_RESPONSES_TAB');
  const last = lastDataRow_(sheet, 1);
  if (last < 2) return [];
  const cols = responseCols_(sheet);
  return sheet.getRange(2, 1, last - 1, cols).getValues().map(function (r) {
    return {
      timestamp: cellText_(r[0]),
      guestId: String(r[1]).trim().toUpperCase(),
      name: String(r[2]),
      attending: String(r[3]) === 'yes' ? 'yes' : 'no',
      pax: Number(r[4]) || 0,
      note: String(r[5]),
      source: String(r[6]) === 'vip' ? 'vip' : 'shared',
      verified: r[7] === true || String(r[7]).toUpperCase() === 'TRUE',
      updated: cellText_(r[8]),
      phone: cols === PHONE_COL ? String(r[9]) : ''
    };
  });
}

/** Config tab: key in column A, value in column B, header in row 1. */
function readConfig_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_CONFIG);
  const cfg = {};
  if (!sheet) return cfg;
  const last = sheet.getLastRow();
  if (last < 2) return cfg;
  sheet.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
    const key = String(r[0]).trim();
    if (key) cfg[key] = r[1];
  });
  return cfg;
}

function findGuest_(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_GUESTS);
  if (!sheet) throw new Error('E_NO_GUESTS_TAB');
  const last = lastDataRow_(sheet, 1);
  if (last < 2) return null;
  const rows = sheet.getRange(2, 1, last - 1, OPENED_COL).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toUpperCase() === id) {
      const name = String(rows[i][1]).trim();
      const pax = Number(rows[i][2]);
      if (!name || !Number.isInteger(pax) || pax < 1) throw new Error('E_BAD_GUEST_ROW');
      return { row: i + 2, name: name, pax: Math.min(pax, MAX_PAX_CEILING), opened: rows[i][OPENED_COL - 1] };
    }
  }
  return null;
}

function findResponse_(guestId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAB_RESPONSES);
  if (!sheet) throw new Error('E_NO_RESPONSES_TAB');
  const last = lastDataRow_(sheet, 1);
  if (last < 2) return null;
  const rows = sheet.getRange(2, 1, last - 1, 5).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][1]).trim().toUpperCase() === guestId) {
      return { row: i + 2, attending: String(rows[i][3]), pax: Number(rows[i][4]) };
    }
  }
  return null;
}

/** Columns the script owns in Responses: A to I, plus J when its header is "phone" (older Sheets may use J for a host formula). */
function responseCols_(sheet) {
  return String(sheet.getRange(1, PHONE_COL).getValue()).trim() === 'phone' ? PHONE_COL : PHONE_COL - 1;
}

/**
 * Last row holding a value in the given column. Used instead of getLastRow() so helper
 * formulas in other columns (for example the duplicate flag) do not push new rows downward.
 */
function lastDataRow_(sheet, col) {
  const max = sheet.getMaxRows();
  const vals = sheet.getRange(1, col, max, 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (vals[i][0] !== '') return i + 1;
  }
  return 1;
}

/* ------------------------------------------------------------------ */
/* Helpers                                                             */
/* ------------------------------------------------------------------ */

function normaliseId_(raw) {
  if (typeof raw !== 'string') return null;
  const id = raw.trim().toUpperCase();
  return ID_PATTERN.test(id) ? id : null;
}

/** Date cell or text -> 'yyyy-MM-dd', or '' if it is not a valid date. */
function deadlineString_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, TZ, 'yyyy-MM-dd');
  const s = String(value === undefined || value === null ? '' : value).trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}

function parseBool_(value) {
  if (value === true || value === false) return value;
  const s = String(value).trim().toUpperCase();
  if (s === 'TRUE') return true;
  if (s === 'FALSE') return false;
  return null;
}

/** Closed when rsvp_open is not TRUE, the deadline has passed, or the deadline is missing or invalid (fail closed). */
function isOpen_(cfg) {
  if (parseBool_(cfg.rsvp_open) !== true) return false;
  const deadline = deadlineString_(cfg.rsvp_deadline);
  if (!deadline) return false;
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  return today <= deadline;
}

function sharedMaxPax_(cfg) {
  const n = Number(cfg.shared_max_pax);
  return Number.isInteger(n) && n >= 1 && n <= MAX_PAX_CEILING ? n : DEFAULT_SHARED_MAX_PAX;
}

/** Config contact_link, only if it is a plain https:// link (it becomes a link on the page). Otherwise ''. */
function contactUrl_(cfg) {
  const s = String(cfg.contact_link === undefined || cfg.contact_link === null ? '' : cfg.contact_link).trim();
  return s.length <= MAX_CONTACT_URL && /^https:\/\/[^\s"'<>\\]+$/i.test(s) ? s : '';
}

/**
 * Removes control characters and neutralises leading formula characters (Sheets and CSV/Excel).
 * Uses a typographic apostrophe (U+2019) on purpose: Sheets silently strips a leading ASCII apostrophe
 * when a script writes a value, so it would not survive a CSV export opened in Excel.
 */
function sanitise_(text) {
  let s = String(text).replace(/[\u0000-\u001F\u007F]/g, ' ').trim();
  if (/^[=+\-@]/.test(s)) s = '\u2019' + s;
  return s;
}

function ok_(message, data) {
  const out = { status: true, message: message };
  if (data !== undefined) out.data = data;
  return json_(out);
}

function fail_(message) {
  return json_({ status: false, message: message });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Logs where it failed, never guest names or notes. */
function logError_(where, err) {
  console.error(where + ' failed: ' + (err && err.name ? err.name : 'Error') + ' ' + (err && err.message && /^E_[A-Z_]+$/.test(err.message) ? err.message : ''));
}

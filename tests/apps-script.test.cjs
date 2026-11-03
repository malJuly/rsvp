// Mock harness: runs apps-script/Code.gs against an in-memory Sheet. Test aid only, not shipped.
const vm = require('vm');
const fs = require('fs');
const crypto = require('crypto');

function makeSheet(rows, maxRows = 200) {
  const sh = { rows: rows.map(r => r.slice()), formats: {}, maxRows };
  const ensure = (r, c) => { while (sh.rows.length < r) sh.rows.push([]); for (let i = 0; i < r; i++) while (sh.rows[i].length < c) sh.rows[i].push(''); };
  sh.setFrozenRows = () => {};
  sh.getMaxRows = () => sh.maxRows;
  sh.getLastRow = () => { let l = 0; sh.rows.forEach((r, i) => { if (r.some(v => v !== '' && v !== undefined)) l = i + 1; }); return l; };
  sh.getRange = (r, c, nr = 1, nc = 1) => ({
    getValues: () => { const out = []; for (let i = 0; i < nr; i++) { const row = []; for (let j = 0; j < nc; j++) { const v = (sh.rows[r - 1 + i] || [])[c - 1 + j]; row.push(v === undefined ? '' : v); } out.push(row); } return out; },
    getFormulas: () => { const out = []; for (let i = 0; i < nr; i++) { const row = []; for (let j = 0; j < nc; j++) { const v = (sh.rows[r - 1 + i] || [])[c - 1 + j]; row.push(typeof v === 'string' && v.startsWith('=') ? v : ''); } out.push(row); } return out; },
    setValues: vals => { ensure(r - 1 + nr, c - 1 + nc); vals.forEach((row, i) => row.forEach((v, j) => { sh.rows[r - 1 + i][c - 1 + j] = v; })); },
    setFormulas: vals => { ensure(r - 1 + nr, c - 1 + nc); vals.forEach((row, i) => row.forEach((v, j) => { sh.rows[r - 1 + i][c - 1 + j] = v; })); },
    getValue: () => ((sh.rows[r - 1] || [])[c - 1] === undefined ? '' : sh.rows[r - 1][c - 1]),
    clearContent: () => { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) { if (sh.rows[r - 1 + i]) sh.rows[r - 1 + i][c - 1 + j] = ''; } },
    setValue: v => { ensure(r, c); sh.rows[r - 1][c - 1] = v; },
    setNumberFormat: f => { sh.formats[r + ',' + c] = f; }
  });
  return sh;
}

function build(opts = {}) {
  const sheets = opts.empty ? { Sheet1: makeSheet([]) } : {
    Config: makeSheet([
      ['key', 'value'],
      ['event_title', 'Garden Dinner'],
      ['rsvp_deadline', '2026-12-10'],
      ['site_url', 'https://example.pages.dev'],
      ['shared_max_pax', 4],
      ['msg_vip', 'Dear {name}, {event} {deadline} {link}'],
      ['msg_shared', 'x'],
      ['rsvp_open', true]
    ]),
    Guests: makeSheet([
      ['guest_id', 'name', 'pax_allowed', 'phone', 'link', 'whatsapp_link', 'reply', 'pax_confirmed', 'responded', 'first_opened', 'status'],
      ['', "Dato' Ashok", 2, 60123456789],
      ['', 'Puan Lina', 3, ''],
      ['', '', '', '']
    ]),
    Responses: makeSheet([['timestamp', 'guest_id', 'name', 'attending', 'pax', 'note', 'source', 'verified', 'updated', 'phone']])
  };
  const alerts = [];
  const ctx = {
    console,
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({ getSheetByName: n => sheets[n] || null, insertSheet: n => { if (ctx.__failInsert) throw new Error('boom'); return (sheets[n] = makeSheet([])); }, deleteSheet: sh => { for (const k of Object.keys(sheets)) if (sheets[k] === sh) delete sheets[k]; }, getSheets: () => Object.values(sheets) }),
      getUi: () => ({
        Button: { YES: 'YES', NO: 'NO', OK: 'OK', CANCEL: 'CANCEL' },
        ButtonSet: { YES_NO: 'YES_NO', OK_CANCEL: 'OK_CANCEL' },
        alert: (m, t) => { alerts.push(t ? t : m); return ctx.__yes === false ? 'NO' : 'YES'; },
        prompt: () => ({ getSelectedButton: () => (ctx.__promptButton || 'OK'), getResponseText: () => (ctx.__promptText === undefined ? 'RESET' : ctx.__promptText) }),
        createMenu: () => ({ addItem() { return this; }, addSeparator() { return this; }, addToUi() {} })
      }),
      flush() {}
    },
    LockService: { getScriptLock: () => ({ tryLock: () => ctx.__lockOk !== false, releaseLock() {} }) },
    // In-memory stand-in; tests expire entries by deleting them from ctx.__cache.
    CacheService: { getScriptCache: () => ({ get: k => (k in ctx.__cache ? ctx.__cache[k] : null), put: (k, v) => { ctx.__cache[k] = String(v); }, remove: k => { delete ctx.__cache[k]; } }) },
    Utilities: {
      getUuid: () => crypto.randomUUID(),
      formatDate: (d, tz, fmt) => {
        const p = new Intl.DateTimeFormat('en-CA', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false }).formatToParts(d).reduce((a, x) => (a[x.type] = x.value, a), {});
        return fmt === 'yyyy-MM-dd' ? `${p.year}-${p.month}-${p.day}` : `${p.year}-${p.month}-${p.day} ${p.hour}:${p.minute}:${p.second}`;
      }
    },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: s => ({ s, setMimeType() { return this; }, getContent() { return this.s; } }) }
  };
  ctx.__cache = {};
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync((process.argv[2] || require('path').join(__dirname, '..', 'apps-script', 'Code.gs')), 'utf8'), ctx);
  const call = (fn, arg) => JSON.parse(vm.runInContext(fn, Object.assign(ctx, { __a: arg })) && ctx.__r);
  const run = (code) => { ctx.__r = undefined; const out = vm.runInContext(code, ctx); return out && out.getContent ? JSON.parse(out.getContent()) : out; };
  return { ctx, sheets, alerts, run };
}

let pass = 0, fail = 0;
const t = (name, cond, extra) => { if (cond) { pass++; console.log('PASS', name); } else { fail++; console.log('FAIL', name, extra !== undefined ? JSON.stringify(extra) : ''); } };
const post = (h, body) => { h.ctx.__body = typeof body === 'string' ? body : JSON.stringify(body); return h.run('doPost({postData:{contents:__body}})'); };
const get = (h, params) => { h.ctx.__p = params; return h.run('doGet({parameter:__p})'); };

// --- Generate IDs
let h = build();
h.run('generateIds()');
const g = h.sheets.Guests.rows;
const id1 = g[1][0], id2 = g[2][0];
t('ids generated, 8 chars, valid alphabet', /^[ABCDEFGHJKLMNPQRSTUVWXYZ2-9]{8}$/.test(id1) && /^[ABCDEFGHJKLMNPQRSTUVWXYZ2-9]{8}$/.test(id2) && id1 !== id2, [id1, id2]);
t('blank-name row gets no id', g[3][0] === '');
t('link formula written', String(g[1][4]).startsWith('=IF(A2'));
const before = id1;
h.run('generateIds()');
t('re-run keeps existing ids', h.sheets.Guests.rows[1][0] === before);
t('typed value preserved on re-run', (() => { h.sheets.Guests.rows[2][7] = 'x'; h.run('generateIds()'); return h.sheets.Guests.rows[2][7] === 'x'; })());

// --- GET
h = build(); h.run('generateIds()');
const vid = h.sheets.Guests.rows[1][0];
let r = get(h, { action: 'guest', id: vid });
t('guest lookup ok, minimal data', r.status && r.data.name === "Dato' Ashok" && r.data.pax === 2 && r.data.replied === null && Object.keys(r.data).sort().join() === 'contactUrl,deadline,name,open,pax,replied', r);
t('lowercase id accepted', get(h, { action: 'guest', id: vid.toLowerCase() }).status === true);
const nf1 = get(h, { action: 'guest', id: 'ZZZZZZZZ' }), nf2 = get(h, { action: 'guest', id: 'bad' }), nf3 = get(h, { action: 'guest' });
t('unknown/malformed/missing id identical message', nf1.message === 'Invitation not found' && nf2.message === nf1.message && nf3.message === nf1.message && !nf1.data);
r = get(h, { action: 'config' });
t('config endpoint', r.status && r.data.open === true && r.data.sharedMaxPax === 4 && r.data.deadline === '2026-12-10', r);
t('unknown action', get(h, { action: 'x' }).status === false);

// --- POST VIP
r = post(h, { action: 'rsvp', id: vid, attending: 'yes', pax: 2, note: 'hi', website: '' });
t('vip rsvp saved', r.status && r.data.pax === 2, r);
let rows = h.sheets.Responses.rows;
t('row written with name from Sheet', rows[1][2] === "Dato' Ashok" && rows[1][6] === 'vip' && rows[1][7] === true, rows[1]);
r = post(h, { action: 'rsvp', id: vid, attending: 'yes', pax: 3 });
t('over-cap pax rejected', r.status === false, r);
r = post(h, { action: 'rsvp', id: vid, attending: 'no', pax: 99 });
t('vip change to no -> updated in place', r.status && r.data.pax === 0 && h.sheets.Responses.rows.length === 2 && h.sheets.Responses.rows[1][3] === 'no' && h.sheets.Responses.rows[1][8] !== '', h.sheets.Responses.rows);
t('lookup now shows replied', get(h, { action: 'guest', id: vid }).data.replied.attending === 'no');
r = post(h, { action: 'rsvp', id: 'ZZZZZZZZ', attending: 'no' });
t('unknown id rejected', r.status === false && r.message === 'Invitation not found');
r = post(h, { action: 'rsvp', id: vid, name: 'Evil Name', attending: 'yes', pax: 1 });
t('client-sent name ignored for vip', h.sheets.Responses.rows[1][2] === "Dato' Ashok");

// --- POST shared
r = post(h, { action: 'rsvp', id: null, name: '=HYPERLINK("http://x")', attending: 'yes', pax: 2, note: '+cmd', website: '' });
rows = h.sheets.Responses.rows;
t('shared saved, source shared, unverified', r.status && rows[2][6] === 'shared' && rows[2][7] === false, rows[2]);
t('formula injection neutralised (name and note)', rows[2][2].startsWith('\u2019=') && rows[2][5].startsWith('\u2019+') && !/^[=+\-@]/.test(rows[2][2]) && !/^[=+\-@]/.test(rows[2][5]), rows[2]);
t('leading - and @ also neutralised', (() => { post(h, { action: 'rsvp', name: '-2+3', attending: 'no', note: '@SUM(1)' }); const last = h.sheets.Responses.rows[h.sheets.Responses.rows.length - 1]; return !/^[=+\-@]/.test(last[2]) && !/^[=+\-@]/.test(last[5]); })());
t('text format set on name/note cells', h.sheets.Responses.formats['3,3'] === '@' && h.sheets.Responses.formats['3,6'] === '@');
t('shared over cap rejected', post(h, { action: 'rsvp', name: 'A', attending: 'yes', pax: 5 }).status === false);
t('shared needs name', post(h, { action: 'rsvp', name: '  ', attending: 'no' }).status === false);
t('name too long rejected', post(h, { action: 'rsvp', name: 'x'.repeat(121), attending: 'no' }).status === false);
t('note too long rejected', post(h, { action: 'rsvp', name: 'A', attending: 'no', note: 'x'.repeat(301) }).status === false);
t('name of 120 and note of 300 accepted', post(h, { action: 'rsvp', name: 'x'.repeat(120), attending: 'no', note: 'y'.repeat(300) }).status === true);
t('pax string rejected', post(h, { action: 'rsvp', name: 'A', attending: 'yes', pax: '2' }).status === false);
t('pax float rejected', post(h, { action: 'rsvp', name: 'A', attending: 'yes', pax: 1.5 }).status === false);
t('bad attending rejected', post(h, { action: 'rsvp', name: 'A', attending: 'maybe' }).status === false);
t('non-JSON rejected', post(h, 'not json').status === false);
t('array body rejected', post(h, '[]').status === false);
t('wrong action rejected', post(h, { action: 'x' }).status === false);
const n = h.sheets.Responses.rows.length;
r = post(h, { action: 'rsvp', name: 'Bot', attending: 'yes', pax: 1, website: 'http://spam' });
t('honeypot: success but nothing written', r.status === true && h.sheets.Responses.rows.length === n);
r = post(h, { action: 'rsvp', name: 'New Shared', attending: 'no' });
t('shared rows always append', r.status && h.sheets.Responses.rows.length === n + 1);

// --- closed / lock
h.sheets.Config.rows[7][1] = false;
t('rsvp_open FALSE closes', post(h, { action: 'rsvp', name: 'A', attending: 'no' }).message === 'RSVP is closed');
h.sheets.Config.rows[7][1] = true; h.sheets.Config.rows[2][1] = '2026-01-01';
t('past deadline closes', post(h, { action: 'rsvp', name: 'A', attending: 'no' }).message === 'RSVP is closed');
t('past deadline blocks VIP change', post(h, { action: 'rsvp', id: vid, attending: 'yes', pax: 1 }).message === 'RSVP is closed');
t('guest endpoint reports closed', get(h, { action: 'guest', id: vid }).data.open === false);
h.sheets.Config.rows[2][1] = 'garbage';
t('invalid deadline fails closed', post(h, { action: 'rsvp', name: 'A', attending: 'no' }).message === 'RSVP is closed');
h.sheets.Config.rows[2][1] = '2026-12-10';
h.ctx.__lockOk = false;
t('lock failure -> Busy', post(h, { action: 'rsvp', name: 'A', attending: 'no' }).message === 'Busy, please try again');
h.ctx.__lockOk = true;

// --- helper formula column does not push rows down
h = build(); h.run('generateIds()');
for (let i = 2; i <= 60; i++) { h.sheets.Responses.rows[i - 1] = h.sheets.Responses.rows[i - 1] || []; h.sheets.Responses.rows[i - 1][9] = '=dup'; }
post(h, { action: 'rsvp', name: 'First', attending: 'no' });
t('append lands in row 2 despite col J content', h.sheets.Responses.rows[1][2] === 'First');

// --- setupSheet
h = build({ empty: true }); h.run('setupSheet()');
t('setup creates 3 tabs and drops empty Sheet1', ['Config', 'Guests', 'Responses'].every(n => h.sheets[n]) && !h.sheets.Sheet1, Object.keys(h.sheets));
t('setup headers and fake guests', h.sheets.Guests.rows[0][0] === 'guest_id' && h.sheets.Guests.rows.length === 4 && /^Test Guest/.test(h.sheets.Guests.rows[1][1]) && h.sheets.Responses.rows[0][8] === 'updated');
h.run('selfTest()');
t('setup output passes selfTest', h.alerts[h.alerts.length - 1] === 'Setup looks good.', h.alerts);
h.run('generateIds()');
t('generate works on fresh setup', /^[A-Z2-9]{8}$/.test(h.sheets.Guests.rows[1][0]) && String(h.sheets.Guests.rows[1][4]).startsWith('=IF(A2') && h.sheets.Guests.rows[1][4].includes('site_url'));
h.sheets.Guests.rows[1][1] = 'Changed';
h.run('setupSheet()');
t('setup never overwrites existing tabs', h.sheets.Guests.rows[1][1] === 'Changed' && /Nothing to do/.test(h.alerts[h.alerts.length - 1]), h.alerts[h.alerts.length - 1]);

// --- tracking: responded / opened / status
h = build(); h.run('generateIds()');
const G = h.sheets.Guests;
t('headers I-K exist', G.rows[0][8] === 'responded' && G.rows[0][9] === 'first_opened' && G.rows[0][10] === 'status');
t('responded + status formulas written', String(G.rows[1][8]).includes('COUNTIF(Responses!B:B') && String(G.rows[1][10]).includes('Not opened') && G.rows[1][9] === '', [G.rows[1][8], G.rows[1][10]]);
const oid = G.rows[1][0];
get(h, { action: 'guest', id: oid });
const firstStamp = G.rows[1][9];
t('first open stamps first_opened', /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(firstStamp), firstStamp);
G.rows[1][9] = '2020-01-01 00:00:00';
get(h, { action: 'guest', id: oid });
t('second open does not overwrite', G.rows[1][9] === '2020-01-01 00:00:00', G.rows[1][9]);
get(h, { action: 'guest', id: 'ZZZZZZZZ' });
t('unknown id stamps nothing', !G.rows[2][9] && !G.rows[3][9], [G.rows[2][9], G.rows[3][9]]);
h.ctx.__lockOk = false;
r = get(h, { action: 'guest', id: G.rows[2][0] });
t('lookup still works if lock is busy (no stamp)', r.status === true && G.rows[2][9] === '', r);
h.ctx.__lockOk = true;
post(h, { action: 'rsvp', id: oid, attending: 'yes', pax: 1 });
t('post by VIP does not break tracking cells', G.rows[1][9] === '2020-01-01 00:00:00');

// old sheet (8 headers, no I-K): lookups work, nothing stamped, Generate IDs upgrades it
h = build(); h.sheets.Guests.rows.forEach(r => { r.length = 8; });
h.run('generateIds()');
t('generate upgrades an older Sheet (adds I-K headers and formulas)', h.sheets.Guests.rows[0][8] === 'responded' && String(h.sheets.Guests.rows[1][10]).startsWith('=IF(A2'), h.sheets.Guests.rows[0]);
h = build(); h.run('generateIds()'); h.sheets.Guests.rows[0][9] = 'something_else';
const oid2 = h.sheets.Guests.rows[1][0];
r = get(h, { action: 'guest', id: oid2 });
t('missing first_opened header: lookup fine, no stamp', r.status === true && h.sheets.Guests.rows[1][9] === '', r);

// bad and duplicate IDs are reported
h = build(); h.run('generateIds()');
h.sheets.Guests.rows[2][0] = h.sheets.Guests.rows[1][0] + '1';
h.run('generateIds()');
t('invalid (9 char) id reported', /Row 3: guest_id .* is not valid/.test(h.alerts[h.alerts.length - 1]), h.alerts[h.alerts.length - 1]);
h.sheets.Guests.rows[2][0] = h.sheets.Guests.rows[1][0];
h.run('generateIds()');
t('duplicate id reported', /already used in row 2/.test(h.alerts[h.alerts.length - 1]), h.alerts[h.alerts.length - 1]);

// --- clear / reset
h = build(); h.run('generateIds()');
const gid = h.sheets.Guests.rows[1][0];
get(h, { action: 'guest', id: gid });
post(h, { action: 'rsvp', id: gid, attending: 'yes', pax: 1, note: 'x' });
post(h, { action: 'rsvp', name: 'Shared One', attending: 'no' });
t('setup: data exists before clearing', h.sheets.Responses.rows.length === 3 && h.sheets.Guests.rows[1][9] !== '');
h.ctx.__yes = false; h.run('clearTestData()');
t('clear: NO answer deletes nothing', h.sheets.Responses.rows.length === 3 && h.sheets.Guests.rows[1][9] !== '' && /Nothing was changed/.test(h.alerts[h.alerts.length - 1]));
h.ctx.__yes = true; h.run('clearTestData()');
t('clear: responses emptied, header kept', h.sheets.Responses.rows.slice(1).every(r => r.every(v => v === '' || v === undefined)) && h.sheets.Responses.rows[0][0] === 'timestamp');
t('clear: first_opened emptied', h.sheets.Guests.rows[1][9] === '');
t('clear: guests, IDs and Config kept', h.sheets.Guests.rows[1][0] === gid && h.sheets.Guests.rows[1][1] === "Dato' Ashok" && h.sheets.Config.rows.length === 8);
t('clear: personal link still works after clearing', get(h, { action: 'guest', id: gid }).data.replied === null);
t('clear: new replies append at row 2 again', (() => { post(h, { action: 'rsvp', name: 'Fresh', attending: 'no' }); return h.sheets.Responses.rows[1][2] === 'Fresh'; })());

h = build(); h.run('generateIds()');
const gid2 = h.sheets.Guests.rows[1][0];
post(h, { action: 'rsvp', id: gid2, attending: 'yes', pax: 1 });
h.ctx.__promptText = 'reset'; h.run('resetForNewEvent()');
t('reset: wrong word (lowercase) deletes nothing', h.sheets.Guests.rows[1][1] === "Dato' Ashok" && h.sheets.Responses.rows[1][1] === gid2);
h.ctx.__promptText = 'RESET'; h.ctx.__promptButton = 'CANCEL'; h.run('resetForNewEvent()');
t('reset: Cancel button deletes nothing', h.sheets.Guests.rows[1][1] === "Dato' Ashok");
h.ctx.__promptButton = 'OK'; h.run('resetForNewEvent()');
t('reset: guests A-D and responses wiped', h.sheets.Guests.rows.slice(1).every(r => [0, 1, 2, 3].every(c => r[c] === '' || r[c] === undefined)) && !h.sheets.Responses.rows[1][1]);
t('reset: headers, Config, formulas kept', h.sheets.Guests.rows[0][1] === 'name' && h.sheets.Config.rows[1][1] === 'Garden Dinner' && String(h.sheets.Guests.rows[1][4]).startsWith('=IF(A2'));
t('reset: old personal link now invalid', get(h, { action: 'guest', id: gid2 }).message === 'Invitation not found');
h.sheets.Guests.rows[1][1] = 'New Guest'; h.sheets.Guests.rows[1][2] = 2;
h.run('generateIds()');
t('reset: new event guests get fresh IDs', /^[A-Z2-9]{8}$/.test(h.sheets.Guests.rows[1][0]) && h.sheets.Guests.rows[1][0] !== gid2);
h.ctx.__promptText = undefined;

// --- soft delete (archive before clearing)
h = build(); h.run('generateIds()');
const sid = h.sheets.Guests.rows[1][0];
get(h, { action: 'guest', id: sid });
post(h, { action: 'rsvp', id: sid, attending: 'yes', pax: 2, note: 'keep this' });
post(h, { action: 'rsvp', name: '=EVIL()', attending: 'no' });
const beforeClear = JSON.stringify(h.sheets.Responses.rows);
h.ctx.__yes = true; h.run('clearTestData()');
const archName = Object.keys(h.sheets).find(n => /^Archive .* Responses$/.test(n));
t('clear: creates an Archive Responses tab', !!archName, Object.keys(h.sheets));
t('archive holds header + every reply, same column order', archName && h.sheets[archName].rows.length === 3 && h.sheets[archName].rows[0][0] === 'timestamp' && h.sheets[archName].rows[1][1] === sid && h.sheets[archName].rows[1][5] === 'keep this', archName && h.sheets[archName].rows);
t('archive text columns set to plain text', archName && h.sheets[archName].formats['1,3'] === '@' && h.sheets[archName].formats['1,6'] === '@');
t('archive of clear keeps opened time in a Guests snapshot', (() => { const g = Object.keys(h.sheets).find(n => /^Archive .* Guests$/.test(n)); return g && h.sheets[g].rows[0][4] === 'first_opened' && h.sheets[g].rows[1][0] === sid && /^\d{4}-/.test(h.sheets[g].rows[1][4]); })());
t('originals cleared after archive', h.sheets.Responses.rows.slice(1).every(r => r.every(v => v === '' || v === undefined)) && h.sheets.Guests.rows[1][9] === '');
t('required tabs untouched', ['Config', 'Guests', 'Responses'].every(n => h.sheets[n]));

// archive failure aborts everything
h = build(); h.run('generateIds()');
post(h, { action: 'rsvp', name: 'Keep Me', attending: 'no' });
h.ctx.__yes = true; h.ctx.__failInsert = true; h.run('clearTestData()');
t('archive failure: nothing deleted', h.sheets.Responses.rows[1][2] === 'Keep Me' && /NOTHING was cleared/.test(h.alerts[h.alerts.length - 1]), h.alerts[h.alerts.length - 1]);
h.ctx.__failInsert = false;

// reset archives guests (with IDs) and replies
h = build(); h.run('generateIds()');
const rid = h.sheets.Guests.rows[1][0], rid2 = h.sheets.Guests.rows[2][0];
post(h, { action: 'rsvp', id: rid, attending: 'yes', pax: 1 });
h.ctx.__promptText = 'RESET'; h.ctx.__promptButton = 'OK'; h.run('resetForNewEvent()');
const gArch = Object.keys(h.sheets).find(n => /^Archive .* Guests$/.test(n));
const rArch = Object.keys(h.sheets).find(n => /^Archive .* Responses$/.test(n));
t('reset: Guests archive has every guest with id, name, pax', gArch && h.sheets[gArch].rows.length === 3 && h.sheets[gArch].rows[1][0] === rid && h.sheets[gArch].rows[2][0] === rid2 && h.sheets[gArch].rows[1][1] === "Dato' Ashok" && h.sheets[gArch].rows[1][2] === 2, gArch && h.sheets[gArch].rows);
t('reset: Responses archive made, originals wiped', rArch && h.sheets[rArch].rows[1][1] === rid && !h.sheets.Guests.rows[1][1] && !h.sheets.Responses.rows[1][1]);
// second archive in the same minute gets a distinct name
h.sheets.Guests.rows[1][1] = 'Again'; h.sheets.Guests.rows[1][2] = 1; h.sheets.Guests.rows[1][0] = 'ZZZZZZZZ';
h.run('resetForNewEvent()');
t('second archive in the same minute does not overwrite the first', Object.keys(h.sheets).filter(n => /^Archive .* Guests/.test(n)).length === 2, Object.keys(h.sheets));
h.ctx.__promptText = undefined;

// --- selfTest
h = build(); h.run('selfTest()');
t('selfTest passes a good setup', h.alerts[0] === 'Setup looks good.', h.alerts);
h = build(); h.sheets.Config.rows[2][1] = 'bad'; h.sheets.Guests.rows[0][1] = 'oops'; h.run('selfTest()');
t('selfTest reports problems', /rsvp_deadline/.test(h.alerts[0]) && /Guests column 2/.test(h.alerts[0]), h.alerts);

// --- phone (optional, shared links only, column J only when its header is "phone")
h = build(); h.run('generateIds()');
r = post(h, { action: 'rsvp', name: 'Shared Phone', attending: 'no', phone: '+60 12-345 6789' });
rows = h.sheets.Responses.rows;
t('shared phone saved in column J as plain text', r.status && rows[1][9] === '+60 12-345 6789' && h.sheets.Responses.formats['2,10'] === '@', rows[1]);
t('shared without phone saves empty phone', post(h, { action: 'rsvp', name: 'No Phone', attending: 'no' }).status && h.sheets.Responses.rows[2][9] === '');
t('phone with letters rejected', post(h, { action: 'rsvp', name: 'A', attending: 'no', phone: '=HYPERLINK(1)' }).message === 'Please enter a valid phone number');
t('phone with too few digits rejected', post(h, { action: 'rsvp', name: 'A', attending: 'no', phone: '1234' }).status === false);
t('phone that is not a string rejected', post(h, { action: 'rsvp', name: 'A', attending: 'no', phone: 60123456789 }).status === false);
const pid = h.sheets.Guests.rows[1][0];
post(h, { action: 'rsvp', id: pid, attending: 'no', phone: '0123456789' });
t('vip phone ignored', h.sheets.Responses.rows[3][1] === pid && h.sheets.Responses.rows[3][9] === '', h.sheets.Responses.rows[3]);

// older Sheet: no phone header, host formula in J -> J never written or cleared
h = build(); h.run('generateIds()');
h.sheets.Responses.rows[0].length = 9;
h.sheets.Responses.rows[1] = ['', '', '', '', '', '', '', '', '', '=dup'];
post(h, { action: 'rsvp', name: 'Old Sheet', attending: 'no', phone: '0123456789' });
t('older Sheet: reply saved, J formula untouched', h.sheets.Responses.rows[1][2] === 'Old Sheet' && h.sheets.Responses.rows[1][9] === '=dup', h.sheets.Responses.rows[1]);
h.ctx.__yes = true; h.run('clearTestData()');
t('older Sheet: clear keeps the J formula', h.sheets.Responses.rows[1][2] === '' && h.sheets.Responses.rows[1][9] === '=dup', h.sheets.Responses.rows[1]);
h.run('selfTest()');
t('older Sheet: selfTest asks for the phone column', /Responses column 10 should be "phone"/.test(h.alerts[h.alerts.length - 1]), h.alerts[h.alerts.length - 1]);

// --- contact_link
h = build();
t('contactUrl empty when not set', get(h, { action: 'config' }).data.contactUrl === '');
h.sheets.Config.rows.push(['contact_link', ' https://wa.link/abc123 ']);
t('config returns contact_link', get(h, { action: 'config' }).data.contactUrl === 'https://wa.link/abc123');
h.run('generateIds()');
t('guest lookup returns contact_link', get(h, { action: 'guest', id: h.sheets.Guests.rows[1][0] }).data.contactUrl === 'https://wa.link/abc123');
for (const bad of ['javascript:alert(1)', 'http://wa.link/x', 'https://x.com/a b', 'https://x.com/"><b>']) {
  h.sheets.Config.rows[h.sheets.Config.rows.length - 1][1] = bad;
  t('unsafe contact_link not returned: ' + bad, get(h, { action: 'config' }).data.contactUrl === '');
}
h.run('selfTest()');
t('selfTest flags a bad contact_link', /contact_link/.test(h.alerts[h.alerts.length - 1]), h.alerts[h.alerts.length - 1]);
h = build({ empty: true }); h.run('setupSheet()');
t('setup adds contact_link row and phone header', h.sheets.Config.rows.some(r => r[0] === 'contact_link') && h.sheets.Responses.rows[0][9] === 'phone');

// --- admin page
const KEY = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
const admin = (h, body) => post(h, Object.assign({ action: 'admin' }, body));
const loginAs = (h) => admin(h, { op: 'login', key: KEY }).data.token;
const withKey = (h) => { h.sheets.Config.rows.push(['admin_key', KEY]); return h; };

h = build(); h.run('generateIds()');
t('admin: not set up without a key', admin(h, { op: 'login', key: 'x' }).message === 'Admin is not set up. In the Sheet, run RSVP tools > Generate admin key');
h.sheets.Config.rows.push(['admin_key', 'short']);
t('admin: a short key counts as not set up', admin(h, { op: 'login', key: 'short' }).status === false);
h.run('selfTest()');
t('selfTest warns about a short admin_key', /admin_key/.test(h.alerts[h.alerts.length - 1]), h.alerts[h.alerts.length - 1]);

h = withKey(build()); h.run('generateIds()');
t('admin: wrong key denied', admin(h, { op: 'login', key: 'nope' }).message === 'Wrong admin key');
r = admin(h, { op: 'login', key: '  ' + KEY + ' ' });
t('admin: right key gives a 64-hex session token', r.status && /^[0-9a-f]{64}$/.test(r.data.token) && r.data.expiresInSeconds === 7200, r);
const tok = r.data.token;
t('admin: no token -> session expired', admin(h, { op: 'overview' }).message === 'Session expired. Please log in again');
t('admin: made-up token -> session expired', admin(h, { op: 'overview', token: 'f'.repeat(64) }).status === false);
t('admin: unknown op rejected', admin(h, { op: 'dropTables', token: tok }).status === false);

// overview
const G0 = h.sheets.Guests.rows;
const gidA = G0[1][0], gidB = G0[2][0];
get(h, { action: 'guest', id: gidA });
post(h, { action: 'rsvp', id: gidA, attending: 'yes', pax: 2, note: 'see you' });
post(h, { action: 'rsvp', name: 'Shared Sam', attending: 'yes', pax: 3, phone: '0123456789' });
post(h, { action: 'rsvp', name: 'Shared No', attending: 'no' });
r = admin(h, { op: 'overview', token: tok });
const ov = r.data;
t('overview: never returns admin_key', r.status && !JSON.stringify(r).includes(KEY));
const ga = ov.guests.find(g => g.id === gidA), gb = ov.guests.find(g => g.id === gidB);
t('overview: personal link built from site_url', ga.link === 'https://example.pages.dev/?id=' + gidA, ga.link);
t('overview: WhatsApp link to the phone with the filled message', ga.whatsappUrl.startsWith('https://wa.me/60123456789?text=') && decodeURIComponent(ga.whatsappUrl.split('text=')[1]) === "Dear Dato' Ashok, Garden Dinner 10th December 2026 https://example.pages.dev/?id=" + gidA, ga.whatsappUrl);
t('overview: no phone -> WhatsApp lets you pick the contact', gb.whatsappUrl.startsWith('https://wa.me/?text='), gb.whatsappUrl);
t('overview: statuses', ga.status === 'replied' && ga.reply.attending === 'yes' && ga.reply.pax === 2 && gb.status === 'not_opened' && /^\d{4}-/.test(ga.firstOpened), [ga, gb]);
t('overview: summary counts', JSON.stringify(ov.summary) === JSON.stringify({ invited: 2, replied: 1, notOpened: 1, attendingReplies: 2, declinedReplies: 1, peopleComing: 5, sharedReplies: 2 }), ov.summary);
t('overview: responses include shared phone', ov.responses.length === 3 && ov.responses[1].phone === '0123456789' && ov.responses[1].source === 'shared' && ov.responses[1].verified === false);
t('overview: settings', ov.settings.deadline === '2026-12-10' && ov.settings.rsvpOpen === true && ov.settings.sharedMaxPax === 4 && ov.settings.sharedWhatsappUrl.startsWith('https://wa.me/?text='), ov.settings);
t('overview: does not stamp first_opened', (() => { admin(h, { op: 'overview', token: tok }); return h.sheets.Guests.rows[2][9] === ''; })());

// add / update
r = admin(h, { op: 'addGuest', token: tok, name: '  YBhg Datin Shivajini  ', pax: 3, phone: '+60 12-999 8888' });
const newRow = h.sheets.Guests.rows.find(x => x[1] === 'YBhg Datin Shivajini');
t('addGuest: row with a fresh valid ID, digits-only phone', r.status && newRow && /^[ABCDEFGHJKLMNPQRSTUVWXYZ2-9]{8}$/.test(newRow[0]) && newRow[0] === r.data.id && newRow[2] === 3 && newRow[3] === '60129998888', newRow);
t('addGuest: link and status formulas written', String(newRow[4]).startsWith('=IF(A') && String(newRow[10]).includes('Not opened'));
t('addGuest: new guest link works at once', get(h, { action: 'guest', id: r.data.id }).data.name === 'YBhg Datin Shivajini');
t('addGuest: formula-looking name neutralised', (() => { const x = admin(h, { op: 'addGuest', token: tok, name: '=IMPORTXML("x")', pax: 1 }); const row = h.sheets.Guests.rows.find(y => y[0] === x.data.id); return row[1].startsWith('\u2019='); })());
t('addGuest: pax 0 and 21 rejected, 20 accepted', admin(h, { op: 'addGuest', token: tok, name: 'A', pax: 0 }).status === false && admin(h, { op: 'addGuest', token: tok, name: 'A', pax: 21 }).status === false && admin(h, { op: 'addGuest', token: tok, name: 'Twenty', pax: 20 }).status === true);
t('addGuest: phone with leading 0 or letters rejected', admin(h, { op: 'addGuest', token: tok, name: 'A', pax: 1, phone: '0123456789' }).status === false && admin(h, { op: 'addGuest', token: tok, name: 'A', pax: 1, phone: 'abc12345678' }).status === false);
t('addGuest: empty name rejected', admin(h, { op: 'addGuest', token: tok, name: '   ', pax: 1 }).status === false);
t('addGuest: lock busy -> Busy', (() => { h.ctx.__lockOk = false; const x = admin(h, { op: 'addGuest', token: tok, name: 'B', pax: 1 }); h.ctx.__lockOk = true; return x.message === 'Busy, please try again'; })());
r = admin(h, { op: 'updateGuest', token: tok, id: gidB.toLowerCase(), name: 'Puan Lina Binti Ali', pax: 4, phone: '' });
t('updateGuest: name, pax, phone changed; ID kept', r.status && h.sheets.Guests.rows[2][0] === gidB && h.sheets.Guests.rows[2][1] === 'Puan Lina Binti Ali' && h.sheets.Guests.rows[2][2] === 4 && h.sheets.Guests.rows[2][3] === '', h.sheets.Guests.rows[2]);
t('updateGuest: unknown ID rejected', admin(h, { op: 'updateGuest', token: tok, id: 'ZZZZZZZZ', name: 'A', pax: 1 }).message === 'Invitation not found');

// settings
t('setConfig: close RSVP', admin(h, { op: 'setConfig', token: tok, rsvpOpen: false }).status && post(h, { action: 'rsvp', name: 'Late', attending: 'no' }).message === 'RSVP is closed');
t('setConfig: reopen and move deadline', admin(h, { op: 'setConfig', token: tok, rsvpOpen: true, deadline: '2026-12-20' }).status && get(h, { action: 'config' }).data.deadline === '2026-12-20' && get(h, { action: 'config' }).data.open === true);
t('setConfig: impossible date rejected', admin(h, { op: 'setConfig', token: tok, deadline: '2026-02-30' }).status === false);
t('setConfig: rsvpOpen must be true/false', admin(h, { op: 'setConfig', token: tok, rsvpOpen: 'yes' }).status === false);
t('setConfig: nothing to change rejected', admin(h, { op: 'setConfig', token: tok }).status === false);

// expiry, key change, lockout
delete h.ctx.__cache['admin_s_' + tok];
t('admin: expired session refused', admin(h, { op: 'overview', token: tok }).message === 'Session expired. Please log in again');
const tok2 = loginAs(h);
h.run('generateAdminKey()');
const newKey = h.sheets.Config.rows.find(x => x[0] === 'admin_key')[1];
t('generateAdminKey: 32 chars from the ID alphabet, replaces the old key', /^[ABCDEFGHJKLMNPQRSTUVWXYZ2-9]{32}$/.test(newKey) && newKey !== KEY && h.alerts[h.alerts.length - 1].includes(newKey), newKey);
t('generateAdminKey: old sessions logged out', admin(h, { op: 'overview', token: tok2 }).status === false);
t('generateAdminKey: new key logs in', admin(h, { op: 'login', key: newKey }).status === true);
h.ctx.__yes = false; h.run('generateAdminKey()'); h.ctx.__yes = undefined;
t('generateAdminKey: cancel keeps the key', h.sheets.Config.rows.find(x => x[0] === 'admin_key')[1] === newKey);
h = build(); h.run('generateAdminKey()');
t('generateAdminKey: adds the Config row when missing', /^[A-Z2-9]{32}$/.test((h.sheets.Config.rows.find(x => x[0] === 'admin_key') || [])[1]));

h = withKey(build());
for (let i = 0; i < 9; i++) admin(h, { op: 'login', key: 'guess' + i });
t('lockout: 10th wrong key locks', admin(h, { op: 'login', key: 'guess9' }).message === 'Too many wrong keys. Try again in 15 minutes');
t('lockout: even the right key is refused while locked', admin(h, { op: 'login', key: KEY }).message === 'Too many wrong keys. Try again in 15 minutes');
delete h.ctx.__cache.admin_locked;
t('lockout: after it expires the right key works', admin(h, { op: 'login', key: KEY }).status === true);
t('public RSVP still works next to admin', post(h, { action: 'rsvp', name: 'Still OK', attending: 'no' }).status === true);

// --- ordinal dates (WhatsApp {deadline})
h = build();
t('ordinal_: st, nd, rd, th and the 11th to 13th', [1, 2, 3, 4, 11, 12, 13, 21, 22, 23, 24, 30, 31].map(n => h.run('ordinal_(' + n + ')')).join() === '1st,2nd,3rd,4th,11th,12th,13th,21st,22nd,23rd,24th,30th,31st');
t('longDate_ uses ordinals', h.run("longDate_('2026-11-06')") === '6th November 2026' && h.run("longDate_('2026-11-22')") === '22nd November 2026');
h.run('generateIds()');
t('Sheet WhatsApp formula writes the deadline with an ordinal', /CHOOSE\(MOD\(DAY\(/.test(String(h.sheets.Guests.rows[1][5])) && !/d mmmm yyyy/.test(String(h.sheets.Guests.rows[1][5])), h.sheets.Guests.rows[1][5]);

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);

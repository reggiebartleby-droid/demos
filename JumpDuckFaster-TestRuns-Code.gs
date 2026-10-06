/**
 * Jump Duck Faster: test-run receiver + friends scoreboard (v44+)
 * Paste this whole file into Extensions > Apps Script of your test-runs Google Sheet,
 * replacing whatever is there. Then update the EXISTING deployment so the URL stays the same:
 * Deploy > Manage deployments > pencil (Edit) > Version: New version > Deploy.
 *
 * POST from the game: appends one row to the "Runs" tab (unchanged, plus 4 new columns at the end).
 * GET ?board=1: returns the friends scoreboard as JSON (new in v44).
 * GET with no parameters: "is it running?" text, as before.
 */

const SHEET_NAME = 'Runs';
const MAX_BODY = 60000;          // bytes; anything bigger is junk, not a run
const COLUMNS = [
  'received', 'name', 'build', 'level', 'chapter', 'boss', 'result', 'place', 'finishPlace',
  'stumbles', 'seconds', 'topSpeed', 'reachedOvertime', 'overtimeSec', 'jumps', 'duckPresses',
  'godMode', 'botMinus', 'device', 'screen', 'session', 'runInSession', 'details',
  // v44: added at the END so older rows keep their columns (they're just blank here)
  'dodges', 'bestStreak', 'eggs', 'godSession'
];
const ACC_MIN = 20;              // accuracy only shows once a player has 20+ counted obstacles

function doPost(e) {
  try {
    const body = (e && e.postData && e.postData.contents) || '';
    if (!body || body.length > MAX_BODY) return reply('rejected: size');
    const run = JSON.parse(body);

    // Light sanity checks. These only filter accidents and lazy junk: anyone can read the
    // game's source, copy this URL and send rows that pass. Treat the sheet as "trusted friends" data.
    if (typeof run.build !== 'string' || !/^v\d+-\d{4}-\d{2}-\d{2}$/.test(run.build)) return reply('rejected: build');
    if (typeof run.result !== 'string' || run.result.length > 20) return reply('rejected: result');

    const row = COLUMNS.map(c => c === 'received' ? new Date() : clean(run[c], c === 'details' ? 45000 : 60));

    const lock = LockService.getScriptLock();   // two friends finishing at once can't overwrite each other
    lock.waitLock(10000);
    try {
      sheetFor().appendRow(row);
    } finally {
      lock.releaseLock();
    }
    return reply('ok');
  } catch (err) {
    return reply('error: ' + err);
  }
}

function doGet(e) {
  if (e && e.parameter && e.parameter.board) {
    try {
      return json({ ok: true, players: board() });
    } catch (err) {
      return json({ ok: false, error: String(err), players: [] });
    }
  }
  // Opening the /exec URL in a browser: a quick "is it deployed?" check.
  return reply('Jump Duck Faster test-run receiver is running.');
}

// Friends scoreboard, overall (all levels), one entry per name (case-insensitive).
// Only rows from v44+ count (older rows have no dodges column).
// Not counted: rows with godMode, botMinus or godSession, AND every row from a session
// in which any row had god mode (god mode at any point taints the whole session).
// Those rows still stay in the sheet; they're just left off the board.
function board() {
  const values = sheetFor().getDataRange().getValues();
  if (values.length < 2) return [];
  const col = {};
  values[0].forEach((h, i) => { col[h] = i; });
  const rows = values.slice(1);
  const yes = v => v === true || String(v).toUpperCase() === 'TRUE';
  const num = v => (typeof v === 'number' && isFinite(v)) ? v : (v !== '' && isFinite(Number(v)) ? Number(v) : null);

  const tainted = {};
  rows.forEach(r => { if (yes(r[col.godMode]) || yes(r[col.godSession])) tainted[r[col.session]] = true; });

  const byName = {};
  rows.forEach(r => {
    const name = String(r[col.name] || '').trim();
    const dodges = num(r[col.dodges]);
    if (!name || name === '(no name)' || dodges === null) return;
    if (yes(r[col.godMode]) || yes(r[col.botMinus]) || yes(r[col.godSession]) || tainted[r[col.session]]) return;
    const key = name.toLowerCase();
    const p = byName[key] || (byName[key] = { name, devices: [], runs: 0, dodges: 0, speed: 0, streak: 0,
                                              ot: 0, eggs: 0, td: 0, ts: 0, last: 0 });
    const when = r[col.received] instanceof Date ? r[col.received].getTime() : 0;
    if (when >= p.last) { p.last = when; p.name = name; }          // show the most recent spelling
    const dev = String(r[col.device] || 'other');
    p.devices = [dev].concat(p.devices.filter(d => d !== dev));    // most recent device first (rows are in time order)
    p.runs++;
    p.dodges = Math.max(p.dodges, dodges);
    p.speed = Math.max(p.speed, num(r[col.topSpeed]) || 0);
    p.streak = Math.max(p.streak, num(r[col.bestStreak]) || 0);
    p.ot = Math.max(p.ot, num(r[col.overtimeSec]) || 0);
    p.eggs += num(r[col.eggs]) || 0;
    p.td += dodges;
    p.ts += num(r[col.stumbles]) || 0;
  });
  return Object.keys(byName).map(k => {
    const p = byName[k];
    return { name: p.name, devices: p.devices.slice(0, 4), runs: p.runs, dodges: p.dodges, speed: p.speed,
             streak: p.streak, ot: p.ot, eggs: p.eggs,
             acc: (p.td + p.ts) >= ACC_MIN ? Math.round(100 * p.td / (p.td + p.ts)) : null };
  });
}

// Turns any value into something safe for a cell.
// Strings that start with = + - @ are prefixed with ' so Sheets shows them as text instead of
// running them as formulas (formula/CSV injection).
function clean(v, maxLen) {
  if (v === undefined || v === null) return '';
  if (typeof v === 'number' || typeof v === 'boolean') return v;
  let s = String(v).slice(0, maxLen);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function sheetFor() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(COLUMNS);
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() < COLUMNS.length ||
             sh.getRange(1, COLUMNS.length).getValue() !== COLUMNS[COLUMNS.length - 1]) {
    // v44: the header row is from an older version: rewrite it so the new columns are labeled.
    sh.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
  }
  return sh;
}

function reply(text) {
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.TEXT);
}
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Run this once from the editor (select testSetup, click Run) to approve permissions
// and create the Runs tab with its header row before any friend plays.
function testSetup() {
  sheetFor().appendRow(COLUMNS.map(c => c === 'received' ? new Date() : (c === 'name' ? 'SETUP TEST (delete me)' : '')));
}

// v44: run this from the editor to see the board JSON in the execution log without a phone.
function testBoard() {
  Logger.log(JSON.stringify(board(), null, 1));
}

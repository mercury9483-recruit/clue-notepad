/**
 * 클루 메모판 기록 수집용 Apps Script
 * 스프레드시트에서 [확장 프로그램 > Apps Script]로 열어 이 코드를 붙여넣고
 * 웹 앱으로 배포합니다. 자세한 순서는 README.md 를 보세요.
 */

// index.html 의 TOKEN 과 같은 값이어야 합니다. 바꾸려면 두 파일을 같이 바꾸세요.
var TOKEN = 'clue-996e205f7cfb3e0f';

var RAW = 'Raw_기록';
var GAMES = 'Games';
var RAW_HEAD = ['시각', '게임ID', '구분', '카드', '플레이어', '표시', '비고', '이벤트ID'];
var GAMES_HEAD = ['게임ID', '시작시각', '인원', '플레이어', '최종수정'];
var TIME_FMT = 'yyyy-mm-dd hh:mm:ss';
var DEDUP_WINDOW = 500;   // 중복 확인에 쓰는 최근 행 수
var MAX_EVENTS = 200;     // 한 번에 받을 최대 건수

/** 처음 한 번 편집기에서 실행하세요. 권한 허용 창이 뜨고 시트 두 개가 만들어집니다. */
function setup() {
  sheet_(RAW, RAW_HEAD);
  sheet_(GAMES, GAMES_HEAD);
}

/** 웹앱 주소를 브라우저로 열었을 때 상태를 보여 줍니다. */
function doGet() {
  var raw = sheet_(RAW, RAW_HEAD);
  return out_({ ok: true, service: 'clue-notepad', rows: Math.max(0, raw.getLastRow() - 1) });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var body = JSON.parse(e.postData.contents);
    if (!body || body.token !== TOKEN) return out_({ ok: false, error: 'token' });
    var events = body.events;
    if (!Array.isArray(events) || events.length === 0 || events.length > MAX_EVENTS) {
      return out_({ ok: false, error: 'events' });
    }

    lock.waitLock(20000);
    var raw = sheet_(RAW, RAW_HEAD);
    var games = sheet_(GAMES, GAMES_HEAD);

    var seen = recentIds_(raw);
    var rows = [];
    var gameEvents = [];

    events.forEach(function (ev) {
      if (!ev || typeof ev.id !== 'string' || seen[ev.id]) return;
      seen[ev.id] = true;
      var t = new Date(ev.t);
      if (isNaN(t.getTime())) t = new Date();
      if (ev.k === 'game') {
        gameEvents.push({ t: t, g: safe_(ev.g), n: Number(ev.n) || '', names: safe_(ev.names) });
      } else if (ev.k === 'mark' || ev.k === 'strike') {
        rows.push([t, safe_(ev.g), safe_(ev.c), safe_(ev.n), safe_(ev.p), safe_(ev.v), safe_(ev.note), ev.id.slice(0, 40)]);
      }
    });

    if (rows.length) {
      var start = raw.getLastRow() + 1;
      raw.getRange(start, 1, rows.length, RAW_HEAD.length).setValues(rows);
      raw.getRange(start, 1, rows.length, 1).setNumberFormat(TIME_FMT);
    }
    gameEvents.forEach(function (g) { upsertGame_(games, g); });

    return out_({ ok: true, added: rows.length, games: gameEvents.length });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function upsertGame_(sh, g) {
  var last = sh.getLastRow();
  var row = 0;
  if (last >= 2) {
    var ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (ids[i][0] === g.g) { row = i + 2; break; }
    }
  }
  if (row) {
    sh.getRange(row, 3, 1, 3).setValues([[g.n, g.names, g.t]]);
    sh.getRange(row, 5).setNumberFormat(TIME_FMT);
  } else {
    row = last + 1;
    sh.getRange(row, 1, 1, GAMES_HEAD.length).setValues([[g.g, g.t, g.n, g.names, g.t]]);
    sh.getRange(row, 2).setNumberFormat(TIME_FMT);
    sh.getRange(row, 5).setNumberFormat(TIME_FMT);
  }
}

function recentIds_(raw) {
  var seen = {};
  var last = raw.getLastRow();
  if (last < 2) return seen;
  var from = Math.max(2, last - DEDUP_WINDOW + 1);
  var vals = raw.getRange(from, RAW_HEAD.length, last - from + 1, 1).getValues();
  vals.forEach(function (r) { if (r[0]) seen[String(r[0])] = true; });
  return seen;
}

function sheet_(name, head) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// 수식으로 해석되는 입력(=, +, -, @ 로 시작)을 막고 길이를 제한합니다.
function safe_(v) {
  var s = (v === undefined || v === null) ? '' : String(v).slice(0, 100);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

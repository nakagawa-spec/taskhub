/**
 * タスクハブ データAPI(Google Apps Script)
 *
 * - このスクリプトは Googleスプレッドシートに紐づけて使う(拡張機能 → Apps Script)。
 * - シート「tasks」は初回アクセス時に自動で作られる。1行 = 1タスク(中身はJSON)。
 * - 画面(GitHub Pages)からは POST(本文はJSON文字列、Content-Type: text/plain)で呼ぶ。
 * - 読み書きには合言葉が必要。スクリプト プロパティ PASSCODE に設定する。
 *   PASSCODE が未設定のあいだは、安全のためすべての読み書きを断る。
 */

var MAX_FAIL = 10;          // この回数まちがえると一時ロック
var LOCK_SEC = 600;         // ロック時間(秒)
var COLS = ['id', 'json', 'updatedAt', 'deleted'];

// ---------------------------------------------------------------- 入口

function doGet() {
  return json_({ ok: true, app: 'taskhub', time: Date.now() });
}

function doPost(e) {
  var req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'リクエストの形式が正しくありません' });
  }
  try {
    var ng = checkPass_(req.pass);
    if (ng) return json_(ng);
    ensureSheet_();
    var fn = ACTIONS[req.action];
    if (!fn) return json_({ ok: false, error: '不明な操作です: ' + req.action });
    return json_(fn(req));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

var ACTIONS = {
  load: load_,
  changes: changes_,
  batch: batch_,
  garoon: garoon_
};

// ---------------------------------------------------------------- 合言葉

function checkPass_(pass) {
  var code = String(PropertiesService.getScriptProperties().getProperty('PASSCODE') || '').trim();
  if (!code) return { ok: false, error: 'GAS のスクリプト プロパティ PASSCODE が設定されていません', code: 'NO_PASSCODE' };
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('fail') || 0);
  if (fails >= MAX_FAIL) return { ok: false, error: '合言葉を何度もまちがえたため、10分ほどロックしています', code: 'LOCKED' };
  if (String(pass || '') !== code) {
    cache.put('fail', String(fails + 1), LOCK_SEC);
    return { ok: false, error: '合言葉が違います', code: 'PASS' };
  }
  return null;
}

// ---------------------------------------------------------------- データ

function load_() {
  var tasks = {};
  var max = 0;
  rows_().forEach(function (r) {
    max = Math.max(max, Number(r.updatedAt) || 0);
    if (r.deleted === true) return;
    var v = parse_(r.json);
    if (v) tasks[r.id] = v;
  });
  return { ok: true, tasks: tasks, version: max };
}

/** since より後に変わったものだけ返す(削除は null) */
function changes_(req) {
  var since = Number(req.since) || 0;
  var tasks = {};
  var max = since;
  rows_().forEach(function (r) {
    var t = Number(r.updatedAt) || 0;
    if (t <= since) return;
    max = Math.max(max, t);
    tasks[r.id] = r.deleted === true ? null : parse_(r.json);
  });
  return { ok: true, tasks: tasks, version: max };
}

/**
 * まとめて書き込む。ops = [{ id, value }]。value が null なら削除。
 * 同じ id が複数あれば最後のものが勝つ。
 */
function batch_(req) {
  var ops = req.ops || [];
  if (!ops.length) return { ok: true, version: 0 };
  if (ops.length > 2000) return { ok: false, error: '一度に保存できる件数を超えました' };
  var last = {};
  ops.forEach(function (op) {
    var id = String(op.id || '').trim();
    if (!id || id.length > 100) throw new Error('タスクIDが正しくありません');
    last[id] = op.value == null ? null : JSON.stringify(op.value);
    if (last[id] && last[id].length > 45000) throw new Error('データが大きすぎて保存できません(1件あたりの上限を超えました)');
  });
  return withLock_(function () {
    var sh = sheet_();
    var values = sh.getDataRange().getValues();
    var rowOf = {};
    for (var i = 1; i < values.length; i++) rowOf[String(values[i][0])] = i;
    var t = nextVersion_();
    var appends = [];
    Object.keys(last).forEach(function (id) {
      var row = [id, last[id] || '', t, last[id] === null];
      if (id in rowOf) values[rowOf[id]] = row;
      else appends.push(row);
    });
    if (values.length > 1) sh.getRange(2, 1, values.length - 1, COLS.length).setValues(values.slice(1));
    if (appends.length) sh.getRange(values.length + 1, 1, appends.length, COLS.length).setValues(appends);
    return { ok: true, version: t };
  });
}

/** 同じミリ秒に2回書いても順番が崩れないように、必ず前回より大きい値を返す */
function nextVersion_() {
  var props = PropertiesService.getScriptProperties();
  var last = Number(props.getProperty('lastVersion') || 0);
  var t = Math.max(Date.now(), last + 1);
  props.setProperty('lastVersion', String(t));
  return t;
}

// ---------------------------------------------------------------- ガルーン(読み取りのみ)
// スクリプト プロパティ GAROON_URL(例 https://xxxx.cybozu.com)・GAROON_USER・GAROON_PASS を使う。
// cybozu.com に Basic 認証をかけている場合だけ GAROON_BASIC(「ID:パスワード」)も設定する。

/** date(YYYY-MM-DD)の1日分の予定を返す */
function garoon_(req) {
  var props = PropertiesService.getScriptProperties();
  var base = String(props.getProperty('GAROON_URL') || '').trim().replace(/\/+$/, '');
  var user = String(props.getProperty('GAROON_USER') || '').trim();
  var pass = String(props.getProperty('GAROON_PASS') || '');
  if (!base || !user || !pass) return { ok: false, error: 'ガルーンの設定(GAROON_URL・GAROON_USER・GAROON_PASS)がまだです', code: 'NO_GAROON' };
  var date = String(req.date || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { ok: false, error: '日付の形式が正しくありません' };

  var p = date.split('-');
  var next = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]) + 1);
  var nextStr = Utilities.formatDate(next, 'Asia/Tokyo', 'yyyy-MM-dd');
  var query = [
    'rangeStart=' + encodeURIComponent(date + 'T00:00:00+09:00'),
    'rangeEnd=' + encodeURIComponent(nextStr + 'T00:00:00+09:00'),
    'orderBy=' + encodeURIComponent('start asc'),
    'limit=1000'
  ].join('&');
  var headers = { 'X-Cybozu-Authorization': Utilities.base64Encode(user + ':' + pass, Utilities.Charset.UTF_8) };
  var basic = String(props.getProperty('GAROON_BASIC') || '');
  if (basic) headers.Authorization = 'Basic ' + Utilities.base64Encode(basic, Utilities.Charset.UTF_8);

  var res = UrlFetchApp.fetch(base + '/g/api/v1/schedule/events?' + query, { method: 'get', headers: headers, muteHttpExceptions: true });
  var status = res.getResponseCode();
  var body = parse_(res.getContentText()) || {};
  if (status === 401) return { ok: false, error: 'ガルーンにログインできません(ログイン名・パスワードを確認してください)', code: 'GAROON_AUTH' };
  if (status !== 200) return { ok: false, error: 'ガルーンから予定を取得できませんでした(' + status + (body.message ? ' ' + body.message : '') + ')' };

  var events = (body.events || []).map(function (ev) {
    return {
      id: String(ev.id),
      subject: ev.subject || '(件名なし)',
      menu: ev.eventMenu || '',
      start: ev.start && ev.start.dateTime || '',
      end: ev.end && ev.end.dateTime || '',
      allDay: !!ev.isAllDay,
      startOnly: !!ev.isStartOnly,
      facilities: (ev.facilities || []).map(function (f) { return f.name; }).filter(Boolean),
      url: base + '/g/schedule/view.csp?event=' + encodeURIComponent(ev.id) + '&bdate=' + date
    };
  });
  return { ok: true, date: date, events: events };
}

/** エディタから実行して、ガルーンにつながるか確かめる(今日の予定の件数がログに出る) */
function testGaroon() {
  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  var r = garoon_({ date: today });
  Logger.log(r.ok ? ('今日の予定: ' + r.events.length + '件 ' + r.events.map(function (e) { return e.subject; }).join(' / ')) : r.error);
}

function parse_(s) {
  try { return JSON.parse(s); } catch (e) { return null; }
}

// ---------------------------------------------------------------- シート操作

function ensureSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('tasks');
  if (sh) return;
  sh = ss.insertSheet('tasks');
  sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  // id が数値・日付に変換されないよう文字列にしておく
  sh.getRange('A:B').setNumberFormat('@');
}

function sheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('tasks');
}

function rows_() {
  var values = sheet_().getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === '') continue;
    var o = {};
    COLS.forEach(function (c, j) { o[c] = values[i][j]; });
    o.id = String(o.id);
    out.push(o);
  }
  return out;
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: '混み合っています。少し待ってからもう一度お試しください' };
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** 動作確認用:エディタから実行するとシートを作る */
function setup() {
  ensureSheet_();
  Logger.log('シートの準備ができました。PASSCODE の設定も忘れずに');
}

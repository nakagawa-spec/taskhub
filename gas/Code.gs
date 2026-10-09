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
  batch: batch_
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

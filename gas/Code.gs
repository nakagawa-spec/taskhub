/**
 * タスクハブ データAPI(Google Apps Script)
 *
 * - このスクリプトは Googleスプレッドシートに紐づけて使う(拡張機能 → Apps Script)。
 * - シート「tasks」は初回アクセス時に自動で作られる。1行 = 1タスク(中身はJSON)。
 * - 画面(GitHub Pages)からは POST(本文はJSON文字列、Content-Type: text/plain)で呼ぶ。
 * - 読み書きには合言葉が必要。スクリプト プロパティ PASSCODE に設定する。
 *   PASSCODE が未設定のあいだは、安全のためすべての読み書きを断る。
 * - 例外はプロジェクトの共有ページ(share.html)。合言葉の代わりに、プロジェクトごとの
 *   共有キー(推測できない長い文字列)で、そのプロジェクトのタスクの閲覧と完了チェックだけできる。
 *
 * 2026-11-05 から：会社用の画面(このスクリプトの Webアプリ、会社のアカウントだけ)に移した。
 * - doGet が画面(index.html)を返す。画面からは google.script.run で th(action, extra) を呼ぶ。
 * - 開いた人は会社のグーグルアカウントで見分ける(合言葉は使わない)。初めての人は名前を入れて登録(users シート)。
 * - tasks シートの5列目 user に持ち主のアドレスを入れ、自分の行だけ読み書きできる。
 *   user が空の行(移す前のもの・議事録の候補)はオーナー(このスクリプトの持ち主＝中川さん)のもの。
 * - 古い公開(@5、合言葉と GitHub Pages の画面)は共有ページ(share.html)のためだけに残す。
 *   doPost は共有ページの操作だけ受け付け、ほかは「移りました」と断る。
 */

var MAX_FAIL = 10;          // この回数まちがえると一時ロック
var LOCK_SEC = 600;         // ロック時間(秒)
var COLS = ['id', 'json', 'updatedAt', 'deleted', 'user'];
var USER_COLS = ['email', 'name', 'registeredAt'];

// ---------------------------------------------------------------- 入口

function doGet(e) {
  if (e && e.parameter && e.parameter.ping) return json_({ ok: true, app: 'taskhub', time: Date.now() });
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('タスクハブ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

// ---------------------------------------------------------------- 会社用の画面から呼ぶ入口

/** 画面の api(action, extra) から呼ばれる。返すものは前の doPost と同じ形 {ok, ...} */
function th(action, extra) {
  var req = extra || {};
  try {
    var me = activeEmail_();
    if (!me) return { ok: false, error: '会社のグーグルアカウントで開いてください', code: 'NOLOGIN' };
    ensureSheet_();
    if (action === 'whoami') return whoami_(me);
    if (action === 'register') return register_(me, req.name);
    if (!userOf_(me)) return { ok: false, error: '初めての方は登録してください', code: 'REGISTER' };
    var fn = ACTIONS[action];
    if (!fn) return { ok: false, error: '不明な操作です: ' + action };
    return fn(req, me);
  } catch (err) {
    return { ok: false, error: String(err && err.message || err) };
  }
}

function activeEmail_() {
  return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
}

/** このスクリプトの持ち主(中川さん)。user が空の行の持ち主 */
function ownerEmail_() {
  return String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
}

function rowUser_(r) {
  return String(r.user || '').trim().toLowerCase() || ownerEmail_();
}

function usersSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('users');
  if (!sh) {
    sh = ss.insertSheet('users');
    sh.getRange(1, 1, 1, USER_COLS.length).setValues([USER_COLS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** 登録している人なら {email, name}。オーナーは登録していなくても使える */
function userOf_(email) {
  var sh = usersSheet_();
  var n = sh.getLastRow();
  if (n > 1) {
    var rows = sh.getRange(2, 1, n - 1, USER_COLS.length).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim().toLowerCase() === email) return { email: email, name: String(rows[i][1] || '') };
    }
  }
  if (email === ownerEmail_()) return { email: email, name: '' };
  return null;
}

function whoami_(me) {
  if (me === ownerEmail_()) stopPasscode();
  var u = userOf_(me);
  return { ok: true, email: me, registered: !!u, name: u ? u.name : '', isOwner: me === ownerEmail_() };
}

function register_(me, name) {
  name = String(name || '').trim().slice(0, 30);
  if (!name) return { ok: false, error: '名前を入れてください' };
  return withLock_(function () {
    var sh = usersSheet_();
    var n = sh.getLastRow();
    var rows = n > 1 ? sh.getRange(2, 1, n - 1, USER_COLS.length).getValues() : [];
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim().toLowerCase() === me) {
        sh.getRange(i + 2, 2).setValue(name);
        return whoami_(me);
      }
    }
    sh.appendRow([me, name, new Date()]);
    return whoami_(me);
  });
}

function doPost(e) {
  var req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'リクエストの形式が正しくありません' });
  }
  try {
    var share = SHARE_ACTIONS[req.action];
    if (share) {
      ensureSheet_();
      return json_(share(req));
    }
    // 合言葉での読み書きはやめた(会社用の画面へ移したため)。古い画面には移った先を知らせる
    return json_({ ok: false, error: 'タスクハブは会社用の画面に移りました。新しいURLから開いてください', code: 'MOVED' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

var ACTIONS = {
  load: load_,
  changes: changes_,
  batch: batch_
};

// 共有ページ用(合言葉なし。共有キーで確認する)
var SHARE_ACTIONS = {
  shareLoad: shareLoad_,
  shareToggle: shareToggle_
};

// ---------------------------------------------------------------- 合言葉

/** 古い公開でも合言葉を使えなくする */
function stopPasscode() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('PASSCODE') !== null) {
    props.deleteProperty('PASSCODE');
    Logger.log('古い画面の合言葉を停止しました（PASSCODE を削除）');
  }
}

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

function load_(req, me) {
  var tasks = {};
  var max = 0;
  rows_().forEach(function (r) {
    max = Math.max(max, Number(r.updatedAt) || 0);
    if (rowUser_(r) !== me) return;
    if (r.deleted === true) return;
    var v = parse_(r.json);
    if (v) tasks[r.id] = v;
  });
  return { ok: true, tasks: tasks, version: max };
}

/** since より後に変わったものだけ返す(削除は null) */
function changes_(req, me) {
  var since = Number(req.since) || 0;
  var tasks = {};
  var max = since;
  rows_().forEach(function (r) {
    var t = Number(r.updatedAt) || 0;
    if (t <= since) return;
    max = Math.max(max, t);
    if (rowUser_(r) !== me) return;
    tasks[r.id] = r.deleted === true ? null : parse_(r.json);
  });
  return { ok: true, tasks: tasks, version: max };
}

/**
 * まとめて書き込む。ops = [{ id, value }]。value が null なら削除。
 * 同じ id が複数あれば最後のものが勝つ。
 * me：書き込む人。ほかの人の行は変えられない。省いたとき(議事録の取り込みなど)はオーナー
 */
function batch_(req, me) {
  var owner = ownerEmail_();
  me = me || owner;
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
    for (var i = 1; i < values.length; i++) {
      while (values[i].length < COLS.length) values[i].push('');
      rowOf[String(values[i][0])] = i;
    }
    var userAt = function (i) { return String(values[i][4] || '').trim().toLowerCase() || owner; };
    var skipped = [];
    Object.keys(last).forEach(function (id) {
      if (id in rowOf && userAt(rowOf[id]) !== me) {
        skipped.push(id);
        delete last[id];
      }
    });
    var t = nextVersion_();
    var appends = [];
    var removedProjects = {};
    Object.keys(last).forEach(function (id) {
      if (last[id] === null && id in rowOf && values[rowOf[id]][3] !== true) {
        var old = parse_(values[rowOf[id]][1]);
        if (old && old.kind === 'project') removedProjects[id] = true;
      }
      var row = [id, last[id] || '', t, last[id] === null, id in rowOf ? values[rowOf[id]][4] : (me === owner ? '' : me)];
      if (id in rowOf) values[rowOf[id]] = row;
      else appends.push(row);
    });
    // プロジェクトを消したら、そのタスクも全部消す(画面にまだ届いていないタスクも残さないため)
    if (Object.keys(removedProjects).length) {
      for (var j = 1; j < values.length; j++) {
        if (values[j][3] === true || userAt(j) !== me) continue;
        var v = parse_(values[j][1]);
        if (v && v.projectId && removedProjects[v.projectId]) values[j] = [values[j][0], '', t, true, values[j][4]];
      }
    }
    if (values.length > 1) sh.getRange(2, 1, values.length - 1, COLS.length).setValues(values.slice(1));
    if (appends.length) sh.getRange(values.length + 1, 1, appends.length, COLS.length).setValues(appends);
    return { ok: true, version: t, skipped: skipped };
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

// ---------------------------------------------------------------- プロジェクトの共有ページ
// プロジェクトは tasks シートに kind:'project' の行として入っている。
// プロジェクトのタスクは json.projectId でつながる。共有ページに返すのは見せてよい項目だけ。

var SHARE_FIELDS = ['title', 'owner', 'startDate', 'dueDate', 'completed', 'waiting', 'order'];
var SHARE_NG = { ok: false, error: 'このプロジェクトは見つかりません(削除されたか、URLが正しくありません)', code: 'SHARE' };

/** 共有キーが合っていればプロジェクトを返す。合わなければ null */
function findSharedProject_(rows, req) {
  var pid = String(req.project || '');
  var key = String(req.key || '');
  if (!pid || key.length < 16) return null;
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (r.id !== pid || r.deleted === true) continue;
    var p = parse_(r.json);
    if (p && p.kind === 'project' && p.shareKey && p.shareKey === key) return p;
    return null;
  }
  return null;
}

function shareLoad_(req) {
  var rows = rows_();
  var proj = findSharedProject_(rows, req);
  if (!proj) return SHARE_NG;
  var pid = String(req.project);
  var tasks = [];
  rows.forEach(function (r) {
    if (r.deleted === true) return;
    var t = parse_(r.json);
    if (!t || t.kind === 'project' || t.projectId !== pid) return;
    var o = { id: r.id };
    SHARE_FIELDS.forEach(function (f) { if (t[f] !== undefined) o[f] = t[f]; });
    tasks.push(o);
  });
  return { ok: true, project: { name: proj.name || '', note: proj.note || '' }, tasks: tasks };
}

/** 共有ページからの完了チェック。そのプロジェクトのタスクの completed だけ変えられる */
function shareToggle_(req) {
  var tid = String(req.task || '');
  if (!tid) return { ok: false, error: 'タスクが指定されていません' };
  return withLock_(function () {
    var rows = rows_();
    if (!findSharedProject_(rows, req)) return SHARE_NG;
    var sh = sheet_();
    var values = sh.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]) !== tid) continue;
      var t = values[i][3] === true ? null : parse_(values[i][1]);
      if (!t || t.projectId !== String(req.project)) break;
      if (t.recurrence && t.recurrence !== 'none') return { ok: false, error: '繰り返しタスクは共有ページから完了にできません' };
      var done = !!req.done;
      var now = new Date().toISOString();
      t.completed = done;
      t.completedAt = done ? now : null;
      t.updatedAt = now;
      t.isNew = false;
      var v = nextVersion_();
      sh.getRange(i + 1, 1, 1, 4).setValues([[tid, JSON.stringify(t), v, false]]); // 5列目(持ち主)はそのまま
      return { ok: true, version: v };
    }
    return { ok: false, error: 'このタスクは見つかりません(削除された可能性があります)', code: 'NOTASK' };
  });
}

// ---------------------------------------------------------------- シート操作

function ensureSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('tasks');
  if (sh) {
    // 5列目(user)を足す前のシートには見出しだけ足す(中身が空ならオーナーのもの)
    if (String(sh.getRange(1, 5).getValue()) !== 'user') sh.getRange(1, 5).setValue('user').setFontWeight('bold');
    return;
  }
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
  Logger.log('シートの準備ができました');
}

/** 許可を出すために、オーナーが編集画面から1回だけ実行する(会社のアカウントを見分ける許可 userinfo.email など) */
function allowAccess() {
  ensureSheet_();
  usersSheet_();
  Logger.log('許可できました。持ち主：' + ownerEmail_());
}

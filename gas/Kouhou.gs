/** 広報室の共有プロジェクト。個人タスクと同じ行を読み書きする。 */
var KOUHOU_DB_ID = '1Cjv_jZEZ24wsOqG_Avgs2in1NX3XC163oeh7xRhE_gg';

function kpMembers_() {
  var id = PropertiesService.getScriptProperties().getProperty('KOUHOU_DB_ID') || KOUHOU_DB_ID;
  var sh = SpreadsheetApp.openById(id).getSheetByName('members');
  if (!sh) throw new Error('広報室の名簿が見つかりません');
  var rows = sh.getDataRange().getValues(), headers = rows.shift();
  var email = headers.indexOf('email'), name = headers.indexOf('name'), active = headers.indexOf('active');
  if (email < 0 || name < 0 || active < 0) throw new Error('広報室の名簿の列を確認してください');
  var seen = {};
  return rows.filter(function (r) {
    var e = String(r[email] || '').trim().toLowerCase();
    if (!e || seen[e] || !/^(true|1)$/i.test(String(r[active]))) return false;
    seen[e] = true;
    return true;
  }).map(function (r) { return { email: String(r[email]).trim().toLowerCase(), name: String(r[name] || r[email]) }; });
}

function kpAuth_() {
  var me = activeEmail_(), members = kpMembers_();
  if (!me || (me !== ownerEmail_() && !members.some(function (m) { return m.email === me; }))) {
    throw new Error('AI広報室のメンバーだけが使えます');
  }
  return { me: me, members: members };
}

function kpPage_() {
  try {
    kpAuth_();
    return HtmlService.createHtmlOutputFromFile('kouhou').setTitle('広報室プロジェクト')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (e) {
    return HtmlService.createHtmlOutput('AI広報室のメンバーだけが使えます').setTitle('広報室プロジェクト');
  }
}

/** リンクは埋め込み画面の住所ではなく、公開済みの入口から作る。 */
function taskhubUrl() {
  if (!activeEmail_()) throw new Error('会社のグーグルアカウントで開いてください');
  return ScriptApp.getService().getUrl();
}

function kp(action, req) {
  try {
    var auth = kpAuth_();
    req = req || {};
    return withLock_(function () {
      ensureSheet_();
      if (action === 'load' || action === 'changes') return kpLoad_(auth, action, req);
      return kpWrite_(action, req, auth);
    });
  } catch (e) { return { ok: false, error: String(e.message || e) }; }
}

function kpLoad_(auth, action, req) {
  var projects = {}, tasks = {}, version = 0;
  rows_().forEach(function (r) {
    version = Math.max(version, Number(r.updatedAt) || 0);
    if (r.deleted === true) return;
    var v = parse_(r.json);
    if (!v) return;
    if (v.kind === 'kouhouProject' && r.user === '__kouhou__') projects[r.id] = v;
    else if (v.kouhouProjectId) tasks[r.id] = v;
  });
  Object.keys(tasks).forEach(function (id) { if (!projects[tasks[id].kouhouProjectId]) delete tasks[id]; });
  if (action === 'changes' && version <= (Number(req.since) || 0)) return { ok: true, changed: false, version: version };
  return { ok: true, changed: true, projects: projects, tasks: tasks, members: auth.members,
    me: auth.me, version: version, taskhubUrl: ScriptApp.getService().getUrl() };
}

function kpText_(v, max) { return String(v == null ? '' : v).trim().slice(0, max); }
function kpMember_(email, auth) {
  email = kpText_(email, 254).toLowerCase();
  var member = auth.members.filter(function (m) { return m.email === email; })[0];
  if (!member) throw new Error('有効な広報室メンバーを選んでください');
  return member;
}
function kpDate_(v) {
  v = kpText_(v, 20);
  if (v && (!/^\d{4}-\d{2}-\d{2}$/.test(v) || isNaN(Date.parse(v)) || new Date(v).toISOString().slice(0, 10) !== v)) throw new Error('日付を確認してください');
  return v;
}

function kpWrite_(action, req, auth) {
  if (['saveProject', 'deleteProject', 'saveTask', 'deleteTask', 'toggleTask'].indexOf(action) < 0) throw new Error('不明な操作です');
  var sh = sheet_(), rows = rows_(), byId = {};
  rows.forEach(function (r, i) { byId[r.id] = r; });
  var values = sh.getDataRange().getValues(), rowAt = {};
  for (var i = 1; i < values.length; i++) rowAt[String(values[i][0])] = i + 1;
  var id = kpText_(req.id, 100), oldRow = id ? byId[id] : null;
  var old = oldRow && oldRow.deleted !== true ? parse_(oldRow.json) : null;
  if (id && !old) throw new Error('対象が見つかりません。読み直してください');
  var isProject = action === 'saveProject' || action === 'deleteProject';
  if (old && (isProject ? old.kind !== 'kouhouProject' || oldRow.user !== '__kouhou__' : !old.kouhouProjectId)) throw new Error('共有プロジェクトの対象ではありません');
  if (!id && action !== 'saveProject' && action !== 'saveTask') throw new Error('対象を選んでください');
  var now = new Date().toISOString(), writes = [];
  id = id || (isProject ? 'kp_' : 'kt_') + Utilities.getUuid();
  function put(tid, value, user, deleted) {
    if (value) value.updatedAt = now;
    var json = JSON.stringify(value || {});
    if (json.length > 45000) throw new Error('データが大きすぎて保存できません');
    writes.push([tid, json, 0, !!deleted, user === ownerEmail_() ? '' : user]);
  }
  function move(tid, t, email, previous) {
    var before = rowUser_(previous);
    t.kouhouPreviousUsers = Array.isArray(t.kouhouPreviousUsers) ? t.kouhouPreviousUsers.slice() : [];
    if (before !== email && t.kouhouPreviousUsers.indexOf(before) < 0) t.kouhouPreviousUsers.push(before);
    put(tid, t, email, false);
  }
  var input = req.value || {};
  if (isProject) {
    var p;
    if (action === 'saveProject') {
      var leader = kpMember_(input.leader, auth);
      p = Object.assign({}, old || {}, { kind: 'kouhouProject', name: kpText_(input.name, 120), client: kpText_(input.client, 120),
        leader: leader.email, leaderName: leader.name, notes: kpText_(input.notes, 10000),
        order: old ? old.order : Date.now(), createdAt: old ? old.createdAt : now });
      if (!p.name) throw new Error('プロジェクト名を入れてください');
      put(id, p, '__kouhou__', false);
    } else put(id, old, '__kouhou__', true);
    rows.forEach(function (r) {
      var t = r.deleted === true ? null : parse_(r.json);
      if (!t || t.kouhouProjectId !== id) return;
      if (action === 'deleteProject') { put(r.id, t, r.user, true); return; }
      t.kouhouProjectName = p.name;
      if (t.assigneeType === 'client' && p.client) { t.assigneeName = p.client; t.owner = '先方：' + p.client; }
      move(r.id, t, t.assigneeType === 'member' ? rowUser_(r) : p.leader, r);
    });
  } else {
    var pid = old ? old.kouhouProjectId : kpText_(input.kouhouProjectId, 100);
    var pr = byId[pid], project = pr && pr.deleted !== true ? parse_(pr.json) : null;
    if (!project || project.kind !== 'kouhouProject' || pr.user !== '__kouhou__') throw new Error('プロジェクトが見つかりません');
    if (action === 'deleteTask') put(id, old, oldRow.user, true);
    else if (action === 'toggleTask') {
      old.completed = !!req.done; old.completedAt = old.completed ? now : null; old.overdueFrom = ''; old.isNew = false;
      put(id, old, oldRow.user, false);
    } else {
      var type = input.assigneeType, email = '', name = '', target = project.leader;
      if (type === 'member') { var member = kpMember_(input.assigneeEmail, auth); email = member.email; name = member.name; target = email; }
      else if (type === 'client' || type === 'other') name = type === 'client' && project.client ? project.client : kpText_(input.assigneeName, 120);
      else throw new Error('担当を選んでください');
      if (!name) throw new Error('担当の名前を入れてください');
      var t = Object.assign({ completed: false, completedAt: null, projectId: '', url: '', priority: 'none', waiting: false,
        recurrence: 'none', weekdays: [], monthDay: 1, subtasks: [], progress: '', isNew: true, order: Date.now(), createdAt: now }, old || {});
      t.title = kpText_(input.title, 500); t.startDate = kpDate_(input.startDate); t.dueDate = kpDate_(input.dueDate);
      if (!t.title) throw new Error('題名を入れてください');
      if (t.startDate && t.dueDate && t.startDate > t.dueDate) throw new Error('開始日が期限より後になっています');
      if (old && t.dueDate !== old.dueDate) t.overdueFrom = '';
      t.notes = kpText_(input.notes, 10000); t.kouhouProjectId = pid; t.kouhouProjectName = project.name;
      t.assigneeType = type; t.assigneeEmail = email; t.assigneeName = name; t.owner = type === 'client' ? '先方：' + name : name;
      if (oldRow) move(id, t, target, oldRow); else put(id, t, target, false);
    }
  }
  var version = nextVersion_();
  writes.forEach(function (r) {
    r[2] = version;
    var row = rowAt[r[0]] || sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, COLS.length).setValues([r]);
  });
  SpreadsheetApp.flush();
  return { ok: true, id: id, version: version };
}

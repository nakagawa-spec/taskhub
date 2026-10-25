/**
 * 議事録 → タスク候補(受け箱)
 *
 * - 「議事録」フォルダの Googleドキュメントを定期的に見て、まだ読んでいない議事録の
 *   「ToDo」の表(内容・担当・期限)を、tasks シートに kind:'candidate' の行として入れる。
 * - 画面の受け箱で「追加」するとタスクになり、「不要」なら消える。
 * - 読んだ議事録はシート「minutes」に記録し、2回は読まない(候補を消しても戻ってこない)。
 * - 最初の1回だけ、エディタから setupMinutes を実行する(権限の許可と、15分ごとの自動実行の登録)。
 */

var MINUTES_FOLDER_ID = '1qomrB_hPtz8-GLwUkCS13W-c5bkO8xHw';
var MINUTES_COLS = ['docId', 'title', 'importedAt', 'count'];

/** 最初に1回だけ実行する */
function setupMinutes() {
  ensureSheet_();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'importMinutes') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('importMinutes').timeBased().everyMinutes(15).create();
  var n = importMinutes();
  Logger.log('15分ごとの自動実行を登録しました。今回の候補:' + n + '件');
}

/** まだ読んでいない議事録から候補を作る。作った候補の件数を返す */
function importMinutes() {
  ensureSheet_();
  var log = minutesSheet_();
  var done = {};
  log.getDataRange().getValues().slice(1).forEach(function (r) { done[String(r[0])] = true; });

  var files = DriveApp.getFolderById(MINUTES_FOLDER_ID).getFilesByType(MimeType.GOOGLE_DOCS);
  var targets = [];
  while (files.hasNext()) {
    var f = files.next();
    if (done[f.getId()]) continue;
    // 録音アプリが書き終える前に読まないよう、作られてから2分たったものだけ読む
    if (Date.now() - f.getDateCreated().getTime() < 2 * 60 * 1000) continue;
    targets.push(f);
  }
  if (!targets.length) return 0;

  var total = 0;
  targets.forEach(function (f) {
    var items = [];
    try {
      items = readTodos_(f.getId());
    } catch (err) {
      Logger.log('読めませんでした: ' + f.getName() + ' ' + err);
      return; // 次回またやり直す
    }
    var title = f.getName().replace(/^議事録[:：]\s*/, '');
    var meetingDate = dateFromTitle_(title) || f.getDateCreated();
    var now = new Date().toISOString();
    var ops = items.map(function (it, i) {
      return {
        id: 'c_' + f.getId().slice(0, 20) + '_' + i,
        value: {
          kind: 'candidate',
          title: it.title,
          owner: it.owner,
          dueDate: parseDue_(it.due, meetingDate),
          dueText: it.due,
          meeting: title,
          meetingDate: fmtYmd_(meetingDate),
          docUrl: f.getUrl(),
          order: i,
          createdAt: now
        }
      };
    });
    if (ops.length) {
      var res = batch_({ ops: ops });
      if (!res.ok) { Logger.log('保存できませんでした: ' + res.error); return; }
    }
    log.appendRow([f.getId(), f.getName(), now, ops.length]);
    total += ops.length;
  });
  return total;
}

/** 「ToDo」見出しの後ろにある最初の表を読む。1行目(内容・担当・期限)は飛ばす */
function readTodos_(docId) {
  var body = DocumentApp.openById(docId).getBody();
  var n = body.getNumChildren();
  var inTodo = false;
  for (var i = 0; i < n; i++) {
    var el = body.getChild(i);
    var type = el.getType();
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      var p = el.asParagraph();
      if (p.getHeading() !== DocumentApp.ParagraphHeading.NORMAL) {
        inTodo = /todo|やること|宿題/i.test(p.getText());
      }
      continue;
    }
    if (!inTodo || type !== DocumentApp.ElementType.TABLE) continue;
    var table = el.asTable();
    var out = [];
    for (var r = 0; r < table.getNumRows(); r++) {
      var row = table.getRow(r);
      var cell = function (c) { return c < row.getNumCells() ? clean_(row.getCell(c).getText()) : ''; };
      var title = cell(0);
      if (!title || title === '内容') continue;
      out.push({ title: title, owner: unknown_(cell(1)), due: unknown_(cell(2)) });
    }
    return out;
  }
  return [];
}

function clean_(s) { return String(s || '').replace(/\*\*/g, '').replace(/\s+/g, ' ').trim(); }
function unknown_(s) { return /^(未定|不明|-|ー|―)?$/.test(s) ? '' : s; }

/** タイトルの「20261024_…」から会議の日付を取る */
function dateFromTitle_(title) {
  var m = String(title).match(/(20\d\d)(\d\d)(\d\d)/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
}

function fmtYmd_(d) {
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
}

/**
 * 期限の書き方(「10/31」「11月21日」「次の木曜日」「来週月曜」「明日」「今週中」など)を
 * 会議の日を基準に日付にする。読めないときは空(画面で元の書き方を見せる)。
 */
function parseDue_(text, base) {
  var s = String(text || '');
  if (!s) return '';
  var b = new Date(base.getFullYear(), base.getMonth(), base.getDate());
  var m = s.match(/(\d{1,2})\s*[\/月]\s*(\d{1,2})/);
  if (m) {
    var d = new Date(b.getFullYear(), Number(m[1]) - 1, Number(m[2]));
    if (d < b && b - d > 60 * 864e5) d.setFullYear(d.getFullYear() + 1); // 年をまたぐ
    return fmtYmd_(d);
  }
  var add = function (n) { var d = new Date(b); d.setDate(d.getDate() + n); return d; };
  if (/明後日|あさって/.test(s)) return fmtYmd_(add(2));
  if (/明日/.test(s)) return fmtYmd_(add(1));
  if (/本日|今日/.test(s)) return fmtYmd_(b);
  var w = s.match(/([月火水木金土日])曜/);
  var nextWeek = /来週|次週/.test(s);
  var nextMon = add(((8 - b.getDay()) % 7) || 7); // 会議の週の翌週(月曜はじまり)の月曜
  var inNextWeek = function (offset) { var d = new Date(nextMon); d.setDate(nextMon.getDate() + offset); return fmtYmd_(d); };
  if (w) {
    var target = '日月火水木金土'.indexOf(w[1]);
    if (nextWeek) return inNextWeek((target + 6) % 7);
    return fmtYmd_(add((target - b.getDay() + 7) % 7 || 7)); // 会議の日より後の、いちばん近いその曜日
  }
  if (/今週/.test(s)) return fmtYmd_(add((5 - b.getDay() + 7) % 7));       // 今週の金曜
  if (nextWeek) return inNextWeek(4);                                     // 来週の金曜
  if (/月末/.test(s)) {
    var plus = /来月/.test(s) ? 2 : 1;
    return fmtYmd_(new Date(b.getFullYear(), b.getMonth() + plus, 0));
  }
  return '';
}

function minutesSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('minutes');
  if (sh) return sh;
  sh = ss.insertSheet('minutes');
  sh.getRange(1, 1, 1, MINUTES_COLS.length).setValues([MINUTES_COLS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('@');
  return sh;
}

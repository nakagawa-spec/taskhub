/**
 * 期限切れのタスクを、期限を今日へずらす(毎朝)。
 *
 * - 期限が過ぎた未完了のタスクは、期限を今日にずらし、元の期限を overdueFrom に残す。
 *   画面は overdueFrom がある間、赤い印「期限切れ ○/○から」を出す。
 * - 15分ごとに動く importMinutes の先頭から呼ぶ。日付が変わったあとの最初の実行でずれる。
 *   (新しい自動実行の登録は要らない)
 * - 画面(index.html の rollOverdue)も同じ処理をする。どちらが先でも結果は同じ。
 * - プロジェクト・候補・完了済み・期限なしは対象外。全員のタスクが対象(持ち主の列は変えない)。
 */

/** ずらしたタスクの件数を返す */
function rollOverdueTasks_() {
  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  return withLock_(function () {
    var sh = sheet_();
    var values = sh.getDataRange().getValues();
    var n = 0;
    var t = null;
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] === '' || values[i][3] === true) continue;
      var v = parse_(values[i][1]);
      if (!v || v.kind === 'project' || v.kind === 'candidate' || v.kind === 'kouhouProject') continue;
      if (v.completed || !v.dueDate || String(v.dueDate) >= today) continue;
      v.overdueFrom = v.overdueFrom || v.dueDate;
      v.dueDate = today;
      v.updatedAt = new Date().toISOString();
      if (t === null) t = nextVersion_();
      sh.getRange(i + 1, 1, 1, 4).setValues([[String(values[i][0]), JSON.stringify(v), t, false]]); // 5列目(持ち主)はそのまま
      n++;
    }
    return n;
  });
}

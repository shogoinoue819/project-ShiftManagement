// シフト作成シートをアップデート
function updateSheets() {
  // SSをまとめて取得
  const [ss, manageSheet, templateSheet, allSheets, ui] = getCommonSheets();

  // 確認ダイアログを表示
  const confirm = ui.alert(
    "⚠️確認",
    "この操作で、現在の日程のシフト作成シートが全て削除されます。\n\n本当に実行してよろしいですか？",
    ui.ButtonSet.OK_CANCEL
  );
  // OKが押されなければキャンセル
  if (confirm !== ui.Button.OK) {
    ui.alert("❌ 操作はキャンセルされました");
    return;
  }

  // メンバーリスト表示をテンプレートに反映
  linkMemberDisplay();
  // 日程リストの取得
  const dateList = getDateList();

  // 各日程において、
  for (const row of dateList) {
    // 日程を取得
    const date = row[0];
    // 日程を文字列(M/d)にフォーマット
    const dateStr = formatDateToString(date, "M/d");

    // 同じ名前のシートが既に存在する場合は削除
    try {
      const existingSheet = ss.getSheetByName(dateStr);
      if (existingSheet) {
        ss.deleteSheet(existingSheet);
        Logger.log(`${dateStr}: 既存シートを削除しました`);
      }
    } catch (e) {
      // シートが存在しない場合は何もしない
    }

    // テンプレートシートをコピーし、日程をシート名にセットしてシフト作成シートを生成
    const newSheet = templateSheet.copyTo(ss).setName(dateStr);
    // A1に日程をセット
    newSheet
      .getRange(SHIFT_TEMPLATE_SHEET.DATE_ROW, SHIFT_TEMPLATE_SHEET.DATE_COL)
      .setValue(date);

    // 出退勤自動記録欄の保護
    const protectionRange = newSheet.getRange(
      SHIFT_TEMPLATE_SHEET.ROWS.START_TIME,
      1,
      SHIFT_TEMPLATE_SHEET.ROWS.WORKING_TIME -
        SHIFT_TEMPLATE_SHEET.ROWS.START_TIME +
        1,
      newSheet.getMaxColumns()
    );
    const protection = protectionRange.protect();
    protection.setDescription("出退勤自動記録欄の保護");
    protection.setWarningOnly(true);

    Logger.log(`${dateStr}: 完了`);
  }

  ui.alert("✅ シフト作成シートをすべて更新しました！");
}

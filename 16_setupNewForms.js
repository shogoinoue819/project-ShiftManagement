function setupNewForms() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNow = ss.getSheetByName(MANAGE_SHEET);
  const sheetPre = ss.getSheetByName(MANAGE_SHEET_PRE);
  const TEMP_NAME = "TEMP_OLD";

  if (!sheetNow || !sheetPre) {
    throw new Error("❌ 管理シートまたは前回分シートが見つかりません");
  }

  // 【追加】日程リストの先頭日付を取得し、ユーザー確認
  const firstDateValue = sheetPre.getRange(MANAGE_DATE_ROW_START, MANAGE_DATE_COLUMN).getValue();
  if (!(firstDateValue instanceof Date)) {
    throw new Error("❌ 日程リストに日付が正しく設定されていません");
  }
  const formattedDate = Utilities.formatDate(firstDateValue, Session.getScriptTimeZone(), "yyyy/MM/dd");

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "確認",
    `シフト希望表を「${formattedDate}」から始まる日程に更新します。\nよろしいですか？`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response !== ui.Button.OK) {
    ui.alert("キャンセルされました。処理を中止します。");
    return;
  }

  // シート名を一時リネーム
  sheetPre.setName(TEMP_NAME);
  sheetNow.setName(MANAGE_SHEET_PRE);
  ss.getSheetByName(TEMP_NAME).setName(MANAGE_SHEET);

  // シートの順序を調整（左から順に MANAGE_SHEET → MANAGE_SHEET_PRE）
  const manageSheet = ss.getSheetByName(MANAGE_SHEET);
  const manageSheetPre = ss.getSheetByName(MANAGE_SHEET_PRE);
  ss.setActiveSheet(manageSheet);
  ss.moveActiveSheet(1); // 一番左へ
  ss.setActiveSheet(manageSheetPre);
  ss.moveActiveSheet(2); // 次に移動

  Logger.log("✅ 管理シートと前回分シートを入れ替えました");

  // 新しい日程リストをテンプレートに反映
  reflectDateList();
}

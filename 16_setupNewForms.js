function setupNewForms() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNow = ss.getSheetByName(SHEET_NAMES.SHIFT_MANAGEMENT);
  const sheetPre = ss.getSheetByName(SHEET_NAMES.SHIFT_MANAGEMENT_PREVIOUS);
  const TEMP_NAME = "TEMP_OLD";

  if (!sheetNow || !sheetPre) {
    throw new Error("❌ 管理シートまたは前回分シートが見つかりません");
  }

  // 【追加】日程リストの先頭日付を取得し、ユーザー確認
  const firstDateValue = sheetPre
    .getRange(MANAGE_DATE_ROW_START, MANAGE_DATE_COLUMN)
    .getValue();
  if (!(firstDateValue instanceof Date)) {
    throw new Error("❌ 日程リストに日付が正しく設定されていません");
  }
  const formattedDate = Utilities.formatDate(
    firstDateValue,
    Session.getScriptTimeZone(),
    "yyyy/MM/dd"
  );

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
  sheetNow.setName(SHEET_NAMES.SHIFT_MANAGEMENT_PREVIOUS);
  ss.getSheetByName(TEMP_NAME).setName(SHEET_NAMES.SHIFT_MANAGEMENT);

  // シートの順序を調整（左から順に SHIFT_MANAGEMENT → SHIFT_MANAGEMENT_PREVIOUS）
  const manageSheet = ss.getSheetByName(SHEET_NAMES.SHIFT_MANAGEMENT);
  const manageSheetPre = ss.getSheetByName(
    SHEET_NAMES.SHIFT_MANAGEMENT_PREVIOUS
  );
  ss.setActiveSheet(manageSheet);
  ss.moveActiveSheet(1); // 一番左へ
  ss.setActiveSheet(manageSheetPre);
  ss.moveActiveSheet(2); // 次に移動

  Logger.log("✅ 管理シートと前回分シートを入れ替えました");

  // 新しい日程リストをテンプレートに反映
  reflectDateList();
}

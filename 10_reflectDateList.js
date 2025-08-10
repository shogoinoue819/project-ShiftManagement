// テンプレートファイルに日程リストを反映
function reflectDateList() {
  // 日程リストの取得
  const dateList = getDateList(); // [[8/25(月)], [8/26(火)], ...] の形式で取得される想定
  const numDates = dateList.length;

  // テンプレートファイルを取得
  const templateFile = SpreadsheetApp.openById(TEMPLATE_FILE_ID);
  const targetSheet = templateFile.getSheetByName(SHEET_NAMES.SHIFT_FORM);
  if (!targetSheet) {
    throw new Error("❌ シフト希望表_テンプレート シートが見つかりません");
  }

  // SSをまとめて取得
  const [ss, manageSheet, templateSheet, allSheets, ui] = getCommonSheets();

  // A列に日付をセット
  targetSheet
    .getRange(FORM_ROW_START, FORM_COLUMN_DATE, numDates, 1)
    .setValues(dateList);

  // B列（完了チェック）を FALSE で初期化
  const falseValues = Array(numDates).fill([false]);
  manageSheet
    .getRange(
      SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW,
      SHIFT_MANAGEMENT_SHEET.DATE_LIST.COMPLETE_COL,
      numDates,
      1
    )
    .setValues(falseValues);

  // C列（共有ステータス）を "未共有" で初期化
  const shareValues = Array(numDates).fill([`${SHARE_FALSE}`]);
  manageSheet
    .getRange(
      SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW,
      SHIFT_MANAGEMENT_SHEET.DATE_LIST.SHARE_COL,
      numDates,
      1
    )
    .setValues(shareValues);

  // 不要な行を削除
  const maxRow = targetSheet.getMaxRows();
  const deleteStart = FORM_ROW_START + numDates;
  if (deleteStart <= maxRow) {
    const numToDelete = maxRow - deleteStart + 1;
    targetSheet.deleteRows(deleteStart, numToDelete);
    Logger.log(`✅ ${deleteStart}行目から ${numToDelete}行分 を削除`);
  } else {
    Logger.log(
      "⚠️ 削除対象の行がシート範囲外だったため、削除をスキップしました"
    );
  }

  Logger.log(
    `✅ 日程 ${numDates} 件をテンプレートに反映し、完了・共有列を初期化しました`
  );
}

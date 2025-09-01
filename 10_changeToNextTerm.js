// シフト管理シートを更新し、新しい日程リストを反映する（reflectDateListの拡張版）
function changeToNextTerm() {
  // SSをまとめて取得
  const ss = getSpreadsheet();
  const ui = getUI();

  // 現在の管理シートと前回分シートを取得
  const sheetNow = ss.getSheetByName(SHEET_NAMES.SHIFT_MANAGEMENT);
  const sheetPre = ss.getSheetByName(SHEET_NAMES.SHIFT_MANAGEMENT_PREVIOUS);

  if (!sheetNow || !sheetPre) {
    throw new Error("❌ 管理シートまたは前回分シートが見つかりません");
  }

  // 【追加】日程リストの先頭日付を取得し、ユーザー確認
  const firstDateValue = sheetPre
    .getRange(
      SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW,
      SHIFT_MANAGEMENT_SHEET.DATE_LIST.COL
    )
    .getValue();

  if (!(firstDateValue instanceof Date)) {
    throw new Error("❌ 日程リストに日付が正しく設定されていません");
  }

  const formattedDate = Utilities.formatDate(
    firstDateValue,
    Session.getScriptTimeZone(),
    "yyyy/MM/dd"
  );

  const response = ui.alert(
    "確認",
    `シフト希望表を「${formattedDate}」から始まる日程に更新します。\nよろしいですか？`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response !== ui.Button.OK) {
    ui.alert("キャンセルされました。処理を中止します。");
    return;
  }

  // シートの入れ替え処理
  swapManagementSheets(ss, sheetNow, sheetPre);

  // 新しい管理シートで日程リストを反映
  const newManageSheet = ss.getSheetByName(SHEET_NAMES.SHIFT_MANAGEMENT);
  reflectDateListInternal(newManageSheet);

  Logger.log("✅ 管理シートの更新と日程リストの反映が完了しました");
}

/**
 * 管理シートと前回分シートを入れ替える
 *
 * @param {Spreadsheet} ss - スプレッドシートオブジェクト
 * @param {Sheet} sheetNow - 現在の管理シート
 * @param {Sheet} sheetPre - 前回分シート
 */
function swapManagementSheets(ss, sheetNow, sheetPre) {
  const TEMP_NAME = "TEMP_OLD";

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
}

/**
 * 日程リストをテンプレートに反映する（内部処理）
 *
 * @param {Sheet} manageSheet - 管理シート
 */
function reflectDateListInternal(manageSheet) {
  // 日程リストの取得
  const dateList = getDateList(manageSheet);
  const numDates = dateList.length;

  if (numDates === 0) {
    throw new Error("❌ 日程リストが取得できませんでした");
  }

  // テンプレートファイルを取得
  const templateFile = SpreadsheetApp.openById(TEMPLATE_FILE_ID);
  const targetSheet = templateFile.getSheetByName(SHEET_NAMES.SHIFT_FORM);

  if (!targetSheet) {
    throw new Error("❌ シフト希望表_テンプレート シートが見つかりません");
  }

  // A列に日付をセット
  targetSheet
    .getRange(
      SHIFT_FORM_TEMPLATE.DATA.START_ROW,
      SHIFT_FORM_TEMPLATE.DATA.DATE_COL,
      numDates,
      1
    )
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
  const shareValues = Array(numDates).fill([`${STATUS_STRINGS.SHARE.FALSE}`]);
  manageSheet
    .getRange(
      SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW,
      SHIFT_MANAGEMENT_SHEET.DATE_LIST.SHARE_COL,
      numDates,
      1
    )
    .setValues(shareValues);

  // 【追加】新しく管理シートにした方のチェック欄と反映欄を全てリセット
  resetMemberListColumns(manageSheet);

  // 不要な行を削除
  const maxRow = targetSheet.getMaxRows();
  const deleteStart = SHIFT_FORM_TEMPLATE.DATA.START_ROW + numDates;

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

/**
 * メンバーリストのチェック欄と反映欄をリセットする
 *
 * @param {Sheet} manageSheet - 管理シート
 */
function resetMemberListColumns(manageSheet) {
  // メンバーリストの最終行を取得
  const lastMemberRow = getLastRowInColumn(
    manageSheet,
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
  );

  if (lastMemberRow < SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW) {
    Logger.log("⚠️ メンバーリストが存在しないため、リセットをスキップしました");
    return;
  }

  const memberCount =
    lastMemberRow - SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW + 1;

  // チェック欄（I列）を FALSE でリセット
  const falseValues = Array(memberCount).fill([false]);
  manageSheet
    .getRange(
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW,
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.CHECK_COL,
      memberCount,
      1
    )
    .setValues(falseValues);

  // 反映欄（J列）を "未反映" でリセット
  const reflectValues = Array(memberCount).fill([STATUS_STRINGS.REFLECT.FALSE]);
  manageSheet
    .getRange(
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW,
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.REFLECT_COL,
      memberCount,
      1
    )
    .setValues(reflectValues);

  Logger.log(
    `✅ メンバー ${memberCount} 名のチェック欄と反映欄をリセットしました`
  );
}

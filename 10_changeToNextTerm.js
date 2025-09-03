// シフト管理シートを更新し、新しい日程リストを反映する（日付入力機能付き）
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

  // 開始日時の入力
  const startDateResponse = ui.prompt(
    "📅 開始日時の入力",
    "新しいシフト期間の開始日時を入力してください。\n形式: M/d (例: 4/1, 12/15)",
    ui.ButtonSet.OK_CANCEL
  );

  if (startDateResponse.getSelectedButton() !== ui.Button.OK) {
    ui.alert("キャンセルされました。処理を中止します。");
    return;
  }

  const startDateStr = startDateResponse.getResponseText().trim();
  const startDate = parseMDDate(startDateStr);

  if (!startDate) {
    ui.alert(
      "❌ エラー",
      "開始日時の形式が正しくありません。\nM/d形式で入力してください (例: 4/1)",
      ui.ButtonSet.OK
    );
    return;
  }

  // 終了日時の入力
  const endDateResponse = ui.prompt(
    "📅 終了日時の入力",
    "新しいシフト期間の終了日時を入力してください。\n形式: M/d (例: 4/30, 12/31)",
    ui.ButtonSet.OK_CANCEL
  );

  if (endDateResponse.getSelectedButton() !== ui.Button.OK) {
    ui.alert("キャンセルされました。処理を中止します。");
    return;
  }

  const endDateStr = endDateResponse.getResponseText().trim();
  const endDate = parseMDDate(endDateStr);

  if (!endDate) {
    ui.alert(
      "❌ エラー",
      "終了日時の形式が正しくありません。\nM/d形式で入力してください (例: 4/30)",
      ui.ButtonSet.OK
    );
    return;
  }

  // 日付の妥当性チェック
  if (endDate <= startDate) {
    ui.alert(
      "❌ エラー",
      "終了日時は開始日時より後の日付を入力してください。",
      ui.ButtonSet.OK
    );
    return;
  }

  // 確認ダイアログ
  const startFormatted = Utilities.formatDate(
    startDate,
    Session.getScriptTimeZone(),
    "M/d"
  );
  const endFormatted = Utilities.formatDate(
    endDate,
    Session.getScriptTimeZone(),
    "M/d"
  );

  const confirmResponse = ui.alert(
    "⚠️ 確認",
    `シフト期間を以下の日程に更新します：\n\n開始: ${startFormatted}\n終了: ${endFormatted}\n\nよろしいですか？`,
    ui.ButtonSet.OK_CANCEL
  );

  if (confirmResponse !== ui.Button.OK) {
    ui.alert("キャンセルされました。処理を中止します。");
    return;
  }

  // シートの入れ替え処理
  swapManagementSheets(ss, sheetNow, sheetPre);

  // 新しい管理シートで日程リストを生成・反映
  const newManageSheet = ss.getSheetByName(SHEET_NAMES.SHIFT_MANAGEMENT);
  generateAndReflectDateList(newManageSheet, startDate, endDate);

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
 * M/d形式の文字列をDateオブジェクトに変換する
 *
 * @param {string} dateStr - M/d形式の日付文字列
 * @returns {Date|null} 変換されたDateオブジェクト、無効な場合はnull
 */
function parseMDDate(dateStr) {
  // M/d形式のパターンをチェック
  const pattern = /^(\d{1,2})\/(\d{1,2})$/;
  const match = dateStr.match(pattern);

  if (!match) {
    return null;
  }

  const month = parseInt(match[1], 10);
  const day = parseInt(match[2], 10);

  // 月の妥当性チェック (1-12)
  if (month < 1 || month > 12) {
    return null;
  }

  // 日の妥当性チェック (1-31)
  if (day < 1 || day > 31) {
    return null;
  }

  // 現在の年を取得
  const currentYear = new Date().getFullYear();

  try {
    // Dateオブジェクトを作成（月は0ベースなので-1）
    const date = new Date(currentYear, month - 1, day);

    // 作成された日付が入力値と一致するかチェック（2月30日などの無効な日付を検出）
    if (date.getMonth() !== month - 1 || date.getDate() !== day) {
      return null;
    }

    return date;
  } catch (error) {
    return null;
  }
}

/**
 * 開始日時と終了日時から日程リストを生成し、管理シートとテンプレートに反映する
 *
 * @param {Sheet} manageSheet - 管理シート
 * @param {Date} startDate - 開始日時
 * @param {Date} endDate - 終了日時
 */
function generateAndReflectDateList(manageSheet, startDate, endDate) {
  // 日程リストを生成
  const dateList = generateDateList(startDate, endDate);
  const numDates = dateList.length;

  if (numDates === 0) {
    throw new Error("❌ 日程リストが生成できませんでした");
  }

  // 管理シートの日程リスト部分をクリアしてから新しい日程を設定
  clearAndSetDateList(manageSheet, dateList, numDates);

  // B列（完了チェック）とC列（共有ステータス）をクリアして初期化
  clearAndInitializeDateStatusColumns(manageSheet, numDates);

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

  // 新しく管理シートにした方のチェック欄と反映欄を全てリセット
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

  Logger.log(`✅ 日程 ${numDates} 件を生成し、テンプレートに反映しました`);
}

/**
 * 開始日時と終了日時から日程リストを生成する
 *
 * @param {Date} startDate - 開始日時
 * @param {Date} endDate - 終了日時
 * @returns {Array<Array<Date>>} 日程リストの配列
 */
function generateDateList(startDate, endDate) {
  const dateList = [];
  const currentDate = new Date(startDate);

  while (currentDate <= endDate) {
    dateList.push([new Date(currentDate)]);
    currentDate.setDate(currentDate.getDate() + 1);
  }

  return dateList;
}

/**
 * 管理シートの日程リスト部分をクリアしてから新しい日程を設定する
 * 既存の日程が新しい日程より多い場合、余分な部分を完全にクリアする
 *
 * @param {Sheet} manageSheet - 管理シート
 * @param {Array<Array<Date>>} dateList - 新しい日程リスト
 * @param {number} numDates - 新しい日程数
 */
function clearAndSetDateList(manageSheet, dateList, numDates) {
  const startRow = SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW;
  const dateCol = SHIFT_MANAGEMENT_SHEET.DATE_LIST.COL;

  // 既存の日程リストの範囲を取得（最大100行まで想定）
  const maxExistingRows = 100;
  const existingRange = manageSheet.getRange(
    startRow,
    dateCol,
    maxExistingRows,
    1
  );

  // 既存の日程を取得
  const existingDates = existingRange.getValues();

  // 既存の日程が新しい日程より多い場合、余分な部分を完全にクリア
  if (existingDates.length > numDates) {
    const clearStartRow = startRow + numDates;
    const clearRowCount = existingDates.length - numDates;

    // 余分な日程の内容のみをクリア（書式は保持）
    const clearRange = manageSheet.getRange(
      clearStartRow,
      dateCol,
      clearRowCount,
      1
    );
    clearRange.clearContent();

    Logger.log(`✅ 余分な日程 ${clearRowCount} 行の内容をクリアしました`);
  }

  // 新しい日程リストを設定
  manageSheet.getRange(startRow, dateCol, numDates, 1).setValues(dateList);

  Logger.log(`✅ 日程リスト ${numDates} 件を設定しました`);
}

/**
 * 管理シートの完了チェックと共有ステータス列をクリアして初期化する
 * 既存のデータが新しい日程より多い場合、余分な部分を完全にクリアする
 *
 * @param {Sheet} manageSheet - 管理シート
 * @param {number} numDates - 新しい日程数
 */
function clearAndInitializeDateStatusColumns(manageSheet, numDates) {
  const startRow = SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW;
  const completeCol = SHIFT_MANAGEMENT_SHEET.DATE_LIST.COMPLETE_COL;
  const shareCol = SHIFT_MANAGEMENT_SHEET.DATE_LIST.SHARE_COL;

  // 既存の完了チェック列の範囲を取得（最大100行まで想定）
  const maxExistingRows = 100;
  const existingCompleteRange = manageSheet.getRange(
    startRow,
    completeCol,
    maxExistingRows,
    1
  );

  // 既存の完了チェックデータを取得
  const existingCompleteData = existingCompleteRange.getValues();

  // 既存のデータが新しい日程より多い場合、余分な部分を完全にクリア
  if (existingCompleteData.length > numDates) {
    const clearStartRow = startRow + numDates;
    const clearRowCount = existingCompleteData.length - numDates;

    // 余分な完了チェック列の内容のみをクリア（書式は保持）
    const clearCompleteRange = manageSheet.getRange(
      clearStartRow,
      completeCol,
      clearRowCount,
      1
    );
    clearCompleteRange.clearContent();

    // 余分な共有ステータス列の内容のみをクリア（書式は保持）
    const clearShareRange = manageSheet.getRange(
      clearStartRow,
      shareCol,
      clearRowCount,
      1
    );
    clearShareRange.clearContent();

    Logger.log(
      `✅ 余分なステータス列 ${clearRowCount} 行の内容をクリアしました`
    );
  }

  // B列（完了チェック）を FALSE で初期化
  const falseValues = Array(numDates).fill([false]);
  manageSheet
    .getRange(startRow, completeCol, numDates, 1)
    .setValues(falseValues);

  // C列（共有ステータス）を "未共有" で初期化
  const shareValues = Array(numDates).fill([`${STATUS_STRINGS.SHARE.FALSE}`]);
  manageSheet.getRange(startRow, shareCol, numDates, 1).setValues(shareValues);

  Logger.log(`✅ ステータス列 ${numDates} 件を初期化しました`);
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

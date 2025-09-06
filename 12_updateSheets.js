/**
 * シフト作成シートをアップデート
 * 新しく作成する日程と同じ名前のシートがある場合のみ更新（削除→再作成）します
 * 既存の他のシートは削除されません
 */
function updateSheets() {
  Logger.log("🔄 シフト作成シートアップデート処理を開始");

  // SSをまとめて取得
  const ss = getSpreadsheet();
  const manageSheet = getManageSheet();
  const ui = getUI();

  Logger.log("📋 スプレッドシートとシートの取得完了");

  // 表示名の空白チェック（最初に実行）
  if (!validateMemberNames(manageSheet, ui)) {
    Logger.log("❌ 表示名の検証に失敗したため、処理を中断します");
    return;
  }

  // 確認ダイアログを表示
  if (!confirmSheetUpdate(ui)) {
    return;
  }

  // 日程リストの取得
  const dateList = getDateList(manageSheet);
  Logger.log(`📅 日程リスト取得成功: ${dateList.length}件`);

  // 進捗表示の初期化（UIでOKを押した直後）
  initializeSheetProgressDisplay(dateList.length);

  // メンバーリスト表示をテンプレートに反映
  const memberDisplaySuccess = updateMemberDisplay();
  if (!memberDisplaySuccess) {
    Logger.log("❌ メンバーリスト表示の更新に失敗したため、処理を中断します");
    return;
  }

  // 各日程のシートを処理
  processDateSheets(dateList);

  Logger.log("🎉 シフト作成シートアップデート処理が完了しました");
  ui.alert("✅ シフト作成シートをすべて更新しました！");
}

/**
 * シート更新の確認
 * ユーザーにシート更新の実行確認を求めます
 * @param {GoogleAppsScript.Base.UI} ui - UIオブジェクト
 * @return {boolean} 確認が取れた場合はtrue、キャンセルの場合はfalse
 */
function confirmSheetUpdate(ui) {
  const confirm = ui.alert(
    "⚠️確認",
    "この操作で、新しく作成する日程と同じ名前のシートがある場合、それらのシートが更新（削除→再作成）されます。\n\n既存の他のシートは削除されません。\n\n本当に実行してよろしいですか？",
    ui.ButtonSet.OK_CANCEL
  );

  if (confirm !== ui.Button.OK) {
    Logger.log("❌ ユーザーにより操作がキャンセルされました");
    ui.alert("❌ 操作はキャンセルされました");
    return false;
  }

  return true;
}

/**
 * 表示名の空白チェック
 * 管理シートと前回分シートの両方で表示名の空白をチェックします
 * @param {GoogleAppsScript.Spreadsheet.Sheet} manageSheet - 管理シート
 * @param {GoogleAppsScript.Base.UI} ui - UIオブジェクト
 * @return {boolean} 検証が成功した場合はtrue、失敗した場合はfalse
 */
function validateMemberNames(manageSheet, ui) {
  Logger.log("🔍 表示名の空白チェックを開始");

  // 管理シートのチェック
  const currentSheetResult = checkMemberNamesInSheet(manageSheet, "管理シート");
  if (!currentSheetResult.isValid) {
    ui.alert(
      "⚠️ 表示名リストに空白のセルがあります！",
      `管理シートの${currentSheetResult.blankRows.join(
        ", "
      )}行目に空白があります。\n` +
        "すべてのメンバーに名前を入力してください。",
      ui.ButtonSet.OK
    );
    return false;
  }

  // 前回分シートのチェック
  const ss = getSpreadsheet();
  const previousSheet = ss.getSheetByName(
    SHEET_NAMES.SHIFT_MANAGEMENT_PREVIOUS
  );
  if (previousSheet) {
    const previousSheetResult = checkMemberNamesInSheet(
      previousSheet,
      "管理シート<前回分>"
    );
    if (!previousSheetResult.isValid) {
      ui.alert(
        "⚠️ 表示名リストに空白のセルがあります！",
        `管理シート<前回分>の${previousSheetResult.blankRows.join(
          ", "
        )}行目に空白があります。\n` +
          "すべてのメンバーに名前を入力してください。",
        ui.ButtonSet.OK
      );
      return false;
    }
  }

  Logger.log("✅ 表示名の空白チェックが完了しました");
  return true;
}

/**
 * 指定されたシートの表示名の空白チェック
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - チェック対象のシート
 * @param {string} sheetName - シート名（ログ用）
 * @return {Object} チェック結果 {isValid: boolean, blankRows: Array<number>}
 */
function checkMemberNamesInSheet(sheet, sheetName) {
  try {
    // 最終行を取得
    const lastRow = getLastRowInColumn(
      sheet,
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
    );

    if (lastRow < SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW) {
      Logger.log(`⚠️ ${sheetName}: メンバーリストが存在しません`);
      return { isValid: true, blankRows: [] };
    }

    // 表示名リストを取得
    const nameRange = sheet.getRange(
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW,
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.DISPLAY_NAME_COL,
      lastRow - SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW + 1,
      1
    );
    const rawNames = nameRange.getValues().flat();

    // 空白セルの行番号を特定
    const blankRows = [];
    rawNames.forEach((name, index) => {
      if (name === "" || name === null || name === undefined) {
        const actualRow = SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW + index;
        blankRows.push(actualRow);
      }
    });

    if (blankRows.length > 0) {
      Logger.log(
        `❌ ${sheetName}: ${
          blankRows.length
        }箇所に空白があります (行: ${blankRows.join(", ")})`
      );
      return { isValid: false, blankRows: blankRows };
    }

    Logger.log(`✅ ${sheetName}: 表示名に空白はありません`);
    return { isValid: true, blankRows: [] };
  } catch (error) {
    Logger.log(`⚠️ ${sheetName}: 表示名チェックでエラー: ${error.message}`);
    return { isValid: false, blankRows: [] };
  }
}

/**
 * メンバーリスト表示の更新
 * シフトテンプレートにメンバーリストを反映します
 * @return {boolean} 成功した場合はtrue、失敗した場合はfalse
 */
function updateMemberDisplay() {
  Logger.log("👥 メンバーリスト表示の更新を開始");
  const success = linkMemberDisplay();
  if (success) {
    Logger.log("✅ メンバーリスト表示の更新が完了しました");
  } else {
    Logger.log("❌ メンバーリスト表示の更新に失敗しました");
  }
  return success;
}

/**
 * 日程シートの処理
 * 各日程のシフト作成シートを作成・更新します
 * @param {Array} dateList - 日程の配列
 */
function processDateSheets(dateList) {
  const ss = getSpreadsheet();
  const templateSheet = getTemplateSheet();
  let successCount = 0;
  let errorCount = 0;
  const errors = [];
  const totalDates = dateList.length;

  Logger.log(`🚀 日程シートの処理を開始: ${totalDates}件`);

  for (const row of dateList) {
    try {
      // 日程を取得
      const date = row[0];
      // 日程を文字列(M/d)にフォーマット
      const dateStr = formatDateToString(date, "M/d");

      createDateSheet(ss, date, dateStr, templateSheet);
      successCount++;

      // 進捗を更新（設定された間隔ごと、または最後の処理）
      const currentProcessed = successCount + errorCount;
      if (
        currentProcessed % UI_DISPLAY.PROGRESS_UPDATE_INTERVAL === 0 ||
        currentProcessed === totalDates
      ) {
        updateSheetProgressDisplay(currentProcessed, totalDates, dateStr);
      }

      Logger.log(`✅ ${dateStr}完了`);
    } catch (e) {
      errorCount++;
      const errorInfo = {
        date: row[0],
        dateStr: formatDateToString(row[0], "M/d"),
        error: e.message,
      };
      errors.push(errorInfo);
      Logger.log(`❌ エラー: ${errorInfo.dateStr || "不明"} - ${e.message}`);
    }
  }

  // 結果サマリーをログ出力
  Logger.log(
    `📊 日程シート処理完了サマリー: 成功 ${successCount}件, エラー ${errorCount}件`
  );

  // エラーが発生した場合の詳細ログ
  if (errors.length > 0) {
    Logger.log("⚠️ エラーが発生した日程:");
    errors.forEach(({ dateStr, error }) => {
      Logger.log(`  - ${dateStr}: ${error}`);
    });
  }

  // 進捗表示をクリア
  clearSheetProgressDisplay();
}

/**
 * 個別の日程シート作成
 * 指定された日程のシフト作成シートを作成します
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - 対象スプレッドシート
 * @param {Date} date - 日程
 * @param {string} dateStr - フォーマットされた日程文字列
 * @param {GoogleAppsScript.Spreadsheet.Sheet} templateSheet - テンプレートシート
 */
function createDateSheet(ss, date, dateStr, templateSheet) {
  // 同じ名前のシートが既に存在する場合は削除（更新）
  const existingSheet = ss.getSheetByName(dateStr);
  if (existingSheet) {
    try {
      ss.deleteSheet(existingSheet);
      Logger.log(`${dateStr}: 既存シートを削除して更新します`);
    } catch (e) {
      Logger.log(`⚠️ ${dateStr}: 既存シートの削除に失敗: ${e.message}`);
      throw new Error(`既存シートの削除に失敗: ${e.message}`);
    }
  } else {
    Logger.log(`${dateStr}: 新規シートを作成します`);
  }

  // テンプレートシートをコピーし、日程をシート名にセットしてシフト作成シートを生成
  const newSheet = templateSheet.copyTo(ss).setName(dateStr);

  // シートの初期化処理を一括で実行
  initializeDateSheet(newSheet, date, dateStr);
}

/**
 * シートの初期化処理を一括で実行
 * 日程の設定とシート保護を順次実行します
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {Date} date - 日程
 * @param {string} dateStr - フォーマットされた日程文字列
 */
function initializeDateSheet(sheet, date, dateStr) {
  // 初期化タスクの定義
  const INITIALIZATION_TASKS = [
    {
      task: () => {
        sheet
          .getRange(
            SHIFT_TEMPLATE_SHEET.DATE_ROW,
            SHIFT_TEMPLATE_SHEET.DATE_COL
          )
          .setValue(date);
      },
      description: "日程の設定",
    },
    {
      task: () => protectWorkingTimeRange(sheet),
      description: "出退勤自動記録欄の保護",
    },
  ];

  // 各初期化タスクを実行
  INITIALIZATION_TASKS.forEach(({ task, description }) => {
    try {
      task();
      Logger.log(`✅ ${dateStr}: ${description}完了`);
    } catch (e) {
      Logger.log(`❌ ${dateStr}: ${description}失敗 - ${e.message}`);
      throw e; // エラーを上位に伝播
    }
  });
}

/**
 * 出退勤自動記録欄の保護
 * シートの出退勤時間入力欄を保護します
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 */
function protectWorkingTimeRange(sheet) {
  // 保護範囲の計算
  const PROTECTION_CONFIG = {
    START_COL: 1,
    ROW_COUNT:
      SHIFT_TEMPLATE_SHEET.ROWS.WORKING_TIME -
      SHIFT_TEMPLATE_SHEET.ROWS.START_TIME +
      1,
    DESCRIPTION: "出退勤自動記録欄の保護",
  };

  const protectionRange = sheet.getRange(
    SHIFT_TEMPLATE_SHEET.ROWS.START_TIME,
    PROTECTION_CONFIG.START_COL,
    PROTECTION_CONFIG.ROW_COUNT,
    sheet.getMaxColumns()
  );

  const protection = protectionRange.protect();
  protection.setDescription(PROTECTION_CONFIG.DESCRIPTION);
  protection.setWarningOnly(true);
}

/**
 * メンバーリスト表示をシフトテンプレートにリンクさせる
 * 管理シートのメンバー情報をテンプレートシートに反映します
 * @return {boolean} 成功した場合はtrue、失敗した場合はfalse
 */
function linkMemberDisplay() {
  Logger.log("👥 メンバーリスト表示のリンク処理を開始");

  // SSをまとめて取得
  const ss = getSpreadsheet();
  const manageSheet = getManageSheet();
  const templateSheet = getTemplateSheet();
  const ui = getUI();

  // メンバー情報の取得と検証
  const memberInfo = getMemberInfoForUpdate(manageSheet, ui);
  if (!memberInfo) {
    Logger.log("❌ メンバー情報の取得に失敗しました");
    return false;
  }

  const { names, bgColors } = memberInfo;
  Logger.log(`📋 メンバー情報取得成功: ${names.length}名`);

  try {
    // メインシートの更新
    updateMainTemplateSheet(templateSheet, names, bgColors);

    // 曜日別テンプレートシートの更新
    updateWeekdayTemplateSheets(names, bgColors);

    // 数式の設定
    setWorkingTimeFormulas(templateSheet, names);

    Logger.log("✅ メンバーリスト表示のリンク処理が完了しました");
    return true;
  } catch (error) {
    Logger.log(`❌ メンバーリスト表示の更新でエラー: ${error.message}`);
    return false;
  }
}

/**
 * メンバー情報の取得と検証
 * @param {GoogleAppsScript.Spreadsheet.Sheet} manageSheet - 管理シート
 * @param {GoogleAppsScript.Base.UI} ui - UIオブジェクト
 * @return {Object|null} メンバー情報（names, bgColors）またはnull
 */
function getMemberInfoForUpdate(manageSheet, ui) {
  // 最終行を取得
  const lastRow = getLastRowInColumn(
    manageSheet,
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
  );

  // 表示名リストを取得
  const nameRange = manageSheet.getRange(
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW,
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.DISPLAY_NAME_COL,
    lastRow - SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW + 1,
    1
  );
  const rawNames = nameRange.getValues().flat();

  // 空白セルが存在するかチェック（既に最初にチェック済みなので、ここでは単純にnullを返す）
  if (
    rawNames.some((name) => name === "" || name === null || name === undefined)
  ) {
    Logger.log("❌ 表示名に空白が検出されました（既にチェック済み）");
    return null;
  }

  // 空白を除いた有効な名前リスト
  const names = rawNames.filter((name) => name);

  // 背景色リストを取得
  const rawColors = nameRange.getBackgrounds().flat();
  const bgColors = rawColors.map((color) => (color ? color : "white"));

  return { names, bgColors };
}

/**
 * メインシートの更新
 * @param {GoogleAppsScript.Spreadsheet.Sheet} templateSheet - テンプレートシート
 * @param {Array} names - メンバー名の配列
 * @param {Array} bgColors - 背景色の配列
 */
function updateMainTemplateSheet(templateSheet, names, bgColors) {
  const lastCol = templateSheet.getLastColumn();

  // テンプレートシートのメンバー欄を取得
  const targetRange = templateSheet.getRange(
    SHIFT_TEMPLATE_SHEET.ROWS.MEMBERS,
    SHIFT_TEMPLATE_SHEET.MEMBER_START_COL,
    1,
    lastCol - 1
  );

  // 内容と背景色をクリア
  targetRange.clearContent();
  targetRange.setBackground(null);

  // 灰色背景をクリア
  templateSheet
    .getRange(
      SHIFT_TEMPLATE_SHEET.ROWS.DATA_START,
      SHIFT_TEMPLATE_SHEET.MEMBER_START_COL,
      SHIFT_TEMPLATE_SHEET.ROWS.DATA_END -
        SHIFT_TEMPLATE_SHEET.ROWS.DATA_START +
        1,
      lastCol - 1
    )
    .setBackground(null);

  // テンプレートシートに氏名と背景色をセット
  for (let i = 0; i < names.length; i++) {
    templateSheet
      .getRange(
        SHIFT_TEMPLATE_SHEET.ROWS.MEMBERS,
        i + SHIFT_TEMPLATE_SHEET.MEMBER_START_COL
      )
      .setValue(names[i]);
    templateSheet
      .getRange(
        SHIFT_TEMPLATE_SHEET.ROWS.MEMBERS,
        i + SHIFT_TEMPLATE_SHEET.MEMBER_START_COL
      )
      .setBackground(bgColors[i]);
  }

  // 背景を灰色に
  templateSheet
    .getRange(
      SHIFT_TEMPLATE_SHEET.ROWS.DATA_START,
      SHIFT_TEMPLATE_SHEET.MEMBER_START_COL,
      SHIFT_TEMPLATE_SHEET.ROWS.DATA_END -
        SHIFT_TEMPLATE_SHEET.ROWS.DATA_START +
        1,
      names.length
    )
    .setBackground(TIME_SETTINGS.UNAVAILABLE_BACKGROUND_COLOR);

  Logger.log("📝 メインシートの更新が完了しました");
}

/**
 * 曜日別テンプレートシートの更新
 * @param {Array} names - メンバー名の配列
 * @param {Array} bgColors - 背景色の配列
 */
function updateWeekdayTemplateSheets(names, bgColors) {
  const ss = getSpreadsheet();
  const allSheets = ss.getSheets();

  const WEEKDAY_TEMPLATES = {
    Mon: SHEET_NAMES.LESSON_TEMPLATES.MON,
    Tue: SHEET_NAMES.LESSON_TEMPLATES.TUE,
    Wed: SHEET_NAMES.LESSON_TEMPLATES.WED,
    Thu: SHEET_NAMES.LESSON_TEMPLATES.THU,
    Fri: SHEET_NAMES.LESSON_TEMPLATES.FRI,
  };

  // 各曜日のテンプレートシートに氏名＋背景色を反映
  for (const day in WEEKDAY_TEMPLATES) {
    const sheetName = WEEKDAY_TEMPLATES[day];
    const sheet = allSheets.find((s) => s.getName() === sheetName);
    if (!sheet) continue;

    updateWeekdaySheet(sheet, names, bgColors);
  }

  Logger.log("📅 曜日別テンプレートシートの更新が完了しました");
}

/**
 * 個別の曜日シートの更新
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {Array} names - メンバー名の配列
 * @param {Array} bgColors - 背景色の配列
 */
function updateWeekdaySheet(sheet, names, bgColors) {
  const lastCol = sheet.getLastColumn();

  // メンバー欄の内容・背景色をリセット（2列目以降）
  if (lastCol >= 2) {
    const targetRange = sheet.getRange(
      SHIFT_TEMPLATE_SHEET.ROWS.MEMBERS,
      SHIFT_TEMPLATE_SHEET.MEMBER_START_COL,
      1,
      lastCol - 1
    );
    targetRange.clearContent();
    targetRange.setBackground(null);
  }

  // 氏名と背景色を1人ずつ反映
  for (let i = 0; i < names.length; i++) {
    sheet
      .getRange(
        SHIFT_TEMPLATE_SHEET.ROWS.MEMBERS,
        i + SHIFT_TEMPLATE_SHEET.MEMBER_START_COL
      )
      .setValue(names[i]);
    sheet
      .getRange(
        SHIFT_TEMPLATE_SHEET.ROWS.MEMBERS,
        i + SHIFT_TEMPLATE_SHEET.MEMBER_START_COL
      )
      .setBackground(bgColors[i]);
  }
}

/**
 * 勤務時間の数式を設定
 * @param {GoogleAppsScript.Spreadsheet.Sheet} templateSheet - テンプレートシート
 * @param {Array} names - メンバー名の配列
 */
function setWorkingTimeFormulas(templateSheet, names) {
  for (let i = 0; i < names.length; i++) {
    const col = i + SHIFT_TEMPLATE_SHEET.MEMBER_START_COL;
    const colLetter = convertColumnToLetter(col);

    // 出勤・退勤・勤務時間の数式を設定
    setWorkStartFormula(templateSheet, col, colLetter);
    setWorkEndFormula(templateSheet, col, colLetter);
    setWorkingTimeFormula(templateSheet, col, colLetter);
  }

  Logger.log("🧮 勤務時間の数式設定が完了しました");
}

/**
 * 出勤時間の数式を設定
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {number} col - 列番号
 * @param {string} colLetter - 列文字
 */
function setWorkStartFormula(sheet, col, colLetter) {
  sheet.getRange(SHIFT_TEMPLATE_SHEET.ROWS.WORK_START, col).setFormula(
    `=LET(
  r, ${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.DATA_START - 1}:${colLetter}${
      SHIFT_TEMPLATE_SHEET.ROWS.DATA_END + 1
    },
  norm, MAP(r, LAMBDA(x,
    IF(
      TO_TEXT(x)="開室",
      TIME(8,0,0) + (ROW(x)-ROW($${colLetter}$${
      SHIFT_TEMPLATE_SHEET.ROWS.DATA_START - 1
    }))*TIME(0,30,0),
      IFERROR(
        IF(REGEXMATCH(TO_TEXT(x),"^\\d{3,4}$"),
          TIME(VALUE(LEFT(TO_TEXT(x), LEN(TO_TEXT(x))-2)), VALUE(RIGHT(TO_TEXT(x),2)), 0),
          IF(ISNUMBER(x),
            IF(x<1, x, TIME(INT(x/100), MOD(x,100), 0)),
            TIMEVALUE(x)
          )
        ),
        NA()
      )
    )
  )),
  t, FILTER(norm, ISNUMBER(norm)),
  IFERROR(TEXT(INDEX(t, 1), "H:MM"), "")
)`
  );
}

/**
 * 退勤時間の数式を設定
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {number} col - 列番号
 * @param {string} colLetter - 列文字
 */
function setWorkEndFormula(sheet, col, colLetter) {
  sheet.getRange(SHIFT_TEMPLATE_SHEET.ROWS.WORK_END, col).setFormula(
    `=LET(
  r, ${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.DATA_START - 1}:${colLetter}${
      SHIFT_TEMPLATE_SHEET.ROWS.DATA_END + 1
    },
  norm, MAP(r, LAMBDA(x,
    IF(
      TO_TEXT(x)="閉室",
      TIME(8,0,0) + (ROW(x)-ROW($${colLetter}$${
      SHIFT_TEMPLATE_SHEET.ROWS.DATA_START - 1
    }) - 1)*TIME(0,30,0),
      IFERROR(
        IF(REGEXMATCH(TO_TEXT(x),"^\\d{3,4}$"),
          TIME(VALUE(LEFT(TO_TEXT(x), LEN(TO_TEXT(x))-2)), VALUE(RIGHT(TO_TEXT(x),2)), 0),
          IF(ISNUMBER(x),
            IF(x<1, x, TIME(INT(x/100), MOD(x,100), 0)),
            TIMEVALUE(x)
          )
        ),
        NA()
      )
    )
  )),
  t, FILTER(norm, ISNUMBER(norm)),
  IFERROR(TEXT(INDEX(t, ROWS(t)), "H:MM"), "")
)`
  );
}

/**
 * 勤務時間の数式を設定
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {number} col - 列番号
 * @param {string} colLetter - 列文字
 */
function setWorkingTimeFormula(sheet, col, colLetter) {
  sheet.getRange(SHIFT_TEMPLATE_SHEET.ROWS.WORKING_TIME, col).setFormula(
    `=IF(
  AND(ISNUMBER(TIMEVALUE(${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.WORK_END})), ISNUMBER(TIMEVALUE(${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.WORK_START}))),
  IF(
    (TIMEVALUE(${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.WORK_END}) - TIMEVALUE(${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.WORK_START})) > TIME(8,0,0),
    TEXT((TIMEVALUE(${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.WORK_END}) - TIMEVALUE(${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.WORK_START})) - TIME(1,0,0), "h:mm"),
    TEXT(TIMEVALUE(${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.WORK_END}) - TIMEVALUE(${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.WORK_START}), "h:mm")
  ),
  ""
)`
  );
}

// シート作成進捗表示の初期化
function initializeSheetProgressDisplay(totalDates) {
  try {
    const { progressCell, statusCell } = getProgressCells();

    // A1は空、B1に準備中を表示
    progressCell.clearContent();
    statusCell.setValue(UI_DISPLAY.SHEET_MESSAGES.PREPARING);

    SpreadsheetApp.flush();
    Logger.log("📊 シート作成進捗表示を初期化しました");
  } catch (error) {
    Logger.log(`⚠️ シート作成進捗表示初期化でエラー: ${error.message}`);
  }
}

// シート作成進捗表示を更新
function updateSheetProgressDisplay(current, total, currentDate) {
  try {
    const { progressCell, statusCell } = getProgressCells();
    const percentage = Math.round((current / total) * 100);

    // A1に進捗、B1に実行中を表示
    progressCell.setValue(`${current}/${total}日 (${percentage}%)`);
    statusCell.setValue(UI_DISPLAY.SHEET_MESSAGES.PROCESSING);

    SpreadsheetApp.flush();
  } catch (error) {
    Logger.log(`⚠️ シート作成進捗表示更新でエラー: ${error.message}`);
  }
}

// シート作成進捗表示をクリア
function clearSheetProgressDisplay() {
  try {
    const { progressCell, statusCell } = getProgressCells();

    // A1とB1の両方をクリア
    progressCell.clearContent();
    statusCell.clearContent();

    SpreadsheetApp.flush();
  } catch (error) {
    Logger.log(`⚠️ シート作成進捗表示クリアでエラー: ${error.message}`);
  }
}

// 進捗表示用セルの取得（共通処理）
function getProgressCells() {
  const ss = getSpreadsheet();
  const manageSheet = getManageSheet();

  return {
    progressCell: manageSheet.getRange(
      UI_DISPLAY.PROGRESS.ROW,
      UI_DISPLAY.PROGRESS.COL
    ),
    statusCell: manageSheet.getRange(
      UI_DISPLAY.STATUS.ROW,
      UI_DISPLAY.STATUS.COL
    ),
  };
}

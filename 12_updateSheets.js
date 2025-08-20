/**
 * シフト作成シートをアップデート
 * 現在の日程のシフト作成シートを全て削除し、新しく作成します
 */
function updateSheets() {
  Logger.log("🔄 シフト作成シートアップデート処理を開始");

  // SSをまとめて取得
  const ss = getSpreadsheet();
  const manageSheet = getManageSheet();
  const ui = getUI();

  Logger.log("📋 スプレッドシートとシートの取得完了");

  // 確認ダイアログを表示
  if (!confirmSheetUpdate(ui)) {
    return;
  }

  // メンバーリスト表示をテンプレートに反映
  updateMemberDisplay();

  // 日程リストの取得
  const dateList = getDateList();
  Logger.log(`📅 日程リスト取得成功: ${dateList.length}件`);

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
    "この操作で、現在の日程のシフト作成シートが全て削除されます。\n\n本当に実行してよろしいですか？",
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
 * メンバーリスト表示の更新
 * シフトテンプレートにメンバーリストを反映します
 */
function updateMemberDisplay() {
  Logger.log("👥 メンバーリスト表示の更新を開始");
  linkMemberDisplay();
  Logger.log("✅ メンバーリスト表示の更新が完了しました");
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

  Logger.log(`🚀 日程シートの処理を開始: ${dateList.length}件`);

  for (const row of dateList) {
    try {
      // 日程を取得
      const date = row[0];
      // 日程を文字列(M/d)にフォーマット
      const dateStr = formatDateToString(date, "M/d");

      createDateSheet(ss, date, dateStr, templateSheet);
      successCount++;
      Logger.log(`✅ ${dateStr}: 完了`);
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
  // 同じ名前のシートが既に存在する場合は削除
  const existingSheet = ss.getSheetByName(dateStr);
  if (existingSheet) {
    try {
      ss.deleteSheet(existingSheet);
      Logger.log(`${dateStr}: 既存シートを削除しました`);
    } catch (e) {
      Logger.log(`⚠️ ${dateStr}: 既存シートの削除に失敗: ${e.message}`);
    }
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
 */
function linkMemberDisplay() {
  Logger.log("👥 メンバーリスト表示のリンク処理を開始");

  // SSをまとめて取得
  const ss = getSpreadsheet();
  const manageSheet = getManageSheet();
  const templateSheet = getTemplateSheet();
  const ui = getUI();

  // メンバー情報の取得と検証
  const memberInfo = getMemberInfo(manageSheet, ui);
  if (!memberInfo) {
    return;
  }

  const { names, bgColors } = memberInfo;
  Logger.log(`📋 メンバー情報取得成功: ${names.length}名`);

  // メインシートの更新
  updateMainTemplateSheet(templateSheet, names, bgColors);

  // 曜日別テンプレートシートの更新
  updateWeekdayTemplateSheets(names, bgColors);

  // 数式の設定
  setWorkingTimeFormulas(templateSheet, names);

  Logger.log("✅ メンバーリスト表示のリンク処理が完了しました");
}

/**
 * メンバー情報の取得と検証
 * @param {GoogleAppsScript.Spreadsheet.Sheet} manageSheet - 管理シート
 * @param {GoogleAppsScript.Base.UI} ui - UIオブジェクト
 * @return {Object|null} メンバー情報（names, bgColors）またはnull
 */
function getMemberInfo(manageSheet, ui) {
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

  // 空白セルが存在するかチェック
  if (
    rawNames.some((name) => name === "" || name === null || name === undefined)
  ) {
    ui.alert(
      "⚠️ 表示名リストに空白のセルがあります。\nすべてのメンバーに名前を入力してください。"
    );
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
  // 授業割テンプレートシートの定義
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
  sheet
    .getRange(SHIFT_TEMPLATE_SHEET.ROWS.WORK_START, col)
    .setFormula(
      `=IFERROR(TO_TEXT(INDEX(${colLetter}${
        SHIFT_TEMPLATE_SHEET.ROWS.DATA_START - 1
      }:${colLetter}${
        SHIFT_TEMPLATE_SHEET.ROWS.DATA_END + 1
      }, MATCH(TRUE, ISNUMBER(SEARCH(":" , TO_TEXT(${colLetter}${
        SHIFT_TEMPLATE_SHEET.ROWS.DATA_START - 1
      }:${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.DATA_END + 1}))), 0))), ""`
    );
}

/**
 * 退勤時間の数式を設定
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {number} col - 列番号
 * @param {string} colLetter - 列文字
 */
function setWorkEndFormula(sheet, col, colLetter) {
  sheet
    .getRange(SHIFT_TEMPLATE_SHEET.ROWS.WORK_END, col)
    .setFormula(
      `=IFERROR(TO_TEXT(INDEX(${colLetter}${
        SHIFT_TEMPLATE_SHEET.ROWS.DATA_START - 1
      }:${colLetter}${
        SHIFT_TEMPLATE_SHEET.ROWS.DATA_END + 1
      }, MAX(FILTER(ROW(${colLetter}${
        SHIFT_TEMPLATE_SHEET.ROWS.DATA_START - 1
      }:${colLetter}${
        SHIFT_TEMPLATE_SHEET.ROWS.DATA_END + 1
      })-ROW(${colLetter}${
        SHIFT_TEMPLATE_SHEET.ROWS.DATA_START - 1
      })+1, ISNUMBER(SEARCH(":" , TO_TEXT(${colLetter}${
        SHIFT_TEMPLATE_SHEET.ROWS.DATA_START - 1
      }:${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.DATA_END + 1}))))))), ""`
    );
}

/**
 * 勤務時間の数式を設定
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {number} col - 列番号
 * @param {string} colLetter - 列文字
 */
function setWorkingTimeFormula(sheet, col, colLetter) {
  sheet
    .getRange(SHIFT_TEMPLATE_SHEET.ROWS.WORKING_TIME, col)
    .setFormula(
      `=IF(AND(ISNUMBER(TIMEVALUE(${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.WORK_END})), ISNUMBER(TIMEVALUE(${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.WORK_START}))), TEXT(TIMEVALUE(${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.WORK_END}) - TIMEVALUE(${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.WORK_START}), "h:mm"), ""`
    );
}

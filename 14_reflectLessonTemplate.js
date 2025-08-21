// 曜日別授業割を反映
function reflectLessonTemplate() {
  // SSをまとめて取得
  const ss = getSpreadsheet();
  const ui = getUI();

  // ターゲットはシフト作成シート
  const targetSheets = getTargetSheets(ss);

  // 各日程のシフト作成シートにおいて、
  targetSheets.forEach((dailySheet) => {
    processDailySheet(dailySheet, ss);
  });

  ui.alert("✅ 授業割テンプレを反映しました！");
}

/**
 * 対象となるシフト作成シートを取得
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - スプレッドシート
 * @returns {GoogleAppsScript.Spreadsheet.Sheet[]} 対象シートの配列
 */
function getTargetSheets(ss) {
  const allSheets = ss.getSheets();
  return allSheets.filter((s) => /^\d{1,2}\/\d{1,2}$/.test(s.getName()));
}

/**
 * 各日程シートの処理を実行
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dailySheet - 日程シート
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - スプレッドシート
 */
function processDailySheet(dailySheet, ss) {
  // シート名を取得
  const sheetName = dailySheet.getName();

  // 日付から曜日を取得
  const dayOfWeek = getDayOfWeekFromSheet(dailySheet);

  // 月〜金に含まれる場合のみ処理
  if (!isWeekday(dayOfWeek)) {
    return;
  }

  // 曜日に対応した授業割シートを取得
  const lessonTemplateSheet = getLessonTemplateSheet(ss, dayOfWeek);
  if (!lessonTemplateSheet) {
    Logger.log(`テンプレートが見つかりません: ${dayOfWeek}`);
    return;
  }

  // テンプレートデータをコピー
  copyTemplateData(dailySheet, lessonTemplateSheet);

  Logger.log(`テンプレートを適用: ${sheetName}`);
}

/**
 * シートから曜日を取得
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dailySheet - 日程シート
 * @returns {string} 曜日（Mon, Tue, Wed, Thu, Fri, Sat, Sun）
 */
function getDayOfWeekFromSheet(dailySheet) {
  const date = dailySheet
    .getRange(SHIFT_TEMPLATE_SHEET.DATE_ROW, SHIFT_TEMPLATE_SHEET.DATE_COL)
    .getValue();
  return formatDateToString(date, "E");
}

/**
 * 曜日が平日（月〜金）かどうかを判定
 * @param {string} dayOfWeek - 曜日（Mon, Tue, Wed, Thu, Fri, Sat, Sun）
 * @returns {boolean} 平日の場合true
 */
function isWeekday(dayOfWeek) {
  const weekdayMap = {
    Mon: true,
    Tue: true,
    Wed: true,
    Thu: true,
    Fri: true,
    Sat: false,
    Sun: false,
  };
  return weekdayMap[dayOfWeek] || false;
}

/**
 * 曜日に対応した授業割テンプレートシートを取得
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - スプレッドシート
 * @param {string} dayOfWeek - 曜日
 * @returns {GoogleAppsScript.Spreadsheet.Sheet|null} テンプレートシート
 */
function getLessonTemplateSheet(ss, dayOfWeek) {
  const templateMap = {
    Mon: SHEET_NAMES.LESSON_TEMPLATES.MON,
    Tue: SHEET_NAMES.LESSON_TEMPLATES.TUE,
    Wed: SHEET_NAMES.LESSON_TEMPLATES.WED,
    Thu: SHEET_NAMES.LESSON_TEMPLATES.THU,
    Fri: SHEET_NAMES.LESSON_TEMPLATES.FRI,
  };

  const templateSheetName = templateMap[dayOfWeek];
  if (!templateSheetName) {
    return null;
  }

  return ss.getSheetByName(templateSheetName);
}

/**
 * テンプレートデータを日程シートにコピー
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dailySheet - 日程シート
 * @param {GoogleAppsScript.Spreadsheet.Sheet} lessonTemplateSheet - 授業割テンプレートシート
 */
function copyTemplateData(dailySheet, lessonTemplateSheet) {
  // 取得する列数を計算
  const columnCount =
    dailySheet.getLastColumn() - SHIFT_TEMPLATE_SHEET.MEMBER_START_COL + 1;

  // データ範囲を取得
  const { sourceRange, targetRange } = getDataRanges(
    dailySheet,
    lessonTemplateSheet,
    columnCount
  );

  // セルの書式とデータをコピー
  copyCellProperties(sourceRange, targetRange);

  // 結合セルの処理
  handleMergedCells(sourceRange, targetRange, dailySheet);

  // ボーダーの適用
  applyBordersToRange(dailySheet, columnCount);
}

/**
 * データ範囲を取得
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dailySheet - 日程シート
 * @param {GoogleAppsScript.Spreadsheet.Sheet} lessonTemplateSheet - 授業割テンプレートシート
 * @param {number} columnCount - 列数
 * @returns {Object} ソース範囲とターゲット範囲
 */
function getDataRanges(dailySheet, lessonTemplateSheet, columnCount) {
  const rowCount =
    SHIFT_TEMPLATE_SHEET.ROWS.DATA_END -
    SHIFT_TEMPLATE_SHEET.ROWS.DATA_START +
    1;

  const sourceRange = lessonTemplateSheet.getRange(
    SHIFT_TEMPLATE_SHEET.ROWS.DATA_START,
    SHIFT_TEMPLATE_SHEET.MEMBER_START_COL,
    rowCount,
    columnCount
  );

  const targetRange = dailySheet.getRange(
    SHIFT_TEMPLATE_SHEET.ROWS.DATA_START,
    SHIFT_TEMPLATE_SHEET.MEMBER_START_COL,
    rowCount,
    columnCount
  );

  return { sourceRange, targetRange };
}

/**
 * セルの書式とデータをコピー
 * @param {GoogleAppsScript.Spreadsheet.Range} sourceRange - ソース範囲
 * @param {GoogleAppsScript.Spreadsheet.Range} targetRange - ターゲット範囲
 */
function copyCellProperties(sourceRange, targetRange) {
  // すべてのプロパティを一括取得
  const values = sourceRange.getValues();
  const backgrounds = sourceRange.getBackgrounds();
  const fontColors = sourceRange.getFontColors();
  const fontSizes = sourceRange.getFontSizes();
  const fontWeights = sourceRange.getFontWeights();

  // 背景色の処理（白背景は保持）
  const processedBackgrounds = processBackgrounds(backgrounds, targetRange);

  // 一括でプロパティを設定
  targetRange.setBackgrounds(processedBackgrounds);
  targetRange.setValues(values);
  targetRange.setFontColors(fontColors);
  targetRange.setFontSizes(fontSizes);
  targetRange.setFontWeights(fontWeights);
}

/**
 * 背景色を処理（白背景は元の背景を保持）
 * @param {Array} sourceBackgrounds - ソースの背景色配列
 * @param {GoogleAppsScript.Spreadsheet.Range} targetRange - ターゲット範囲
 * @returns {Array} 処理済みの背景色配列
 */
function processBackgrounds(sourceBackgrounds, targetRange) {
  // 元の背景色を取得
  const currentBackgrounds = targetRange.getBackgrounds();

  // 新しい背景色配列を作成
  return sourceBackgrounds.map((row, i) =>
    row.map((sourceColor, j) => {
      // 白背景（#ffffff）またはnullの場合は元の背景を保持
      if (sourceColor === "#ffffff" || sourceColor === null) {
        return currentBackgrounds[i][j];
      }
      return sourceColor;
    })
  );
}

/**
 * 結合セルの処理
 * @param {GoogleAppsScript.Spreadsheet.Range} sourceRange - ソース範囲
 * @param {GoogleAppsScript.Spreadsheet.Range} targetRange - ターゲット範囲
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dailySheet - 日程シート
 */
function handleMergedCells(sourceRange, targetRange, dailySheet) {
  const mergedRanges = sourceRange.getMergedRanges();

  mergedRanges.forEach((range) => {
    const rowOffset = range.getRow() - SHIFT_TEMPLATE_SHEET.ROWS.DATA_START;
    const colOffset = range.getColumn() - SHIFT_TEMPLATE_SHEET.MEMBER_START_COL;

    const targetRange = dailySheet.getRange(
      SHIFT_TEMPLATE_SHEET.ROWS.DATA_START + rowOffset,
      SHIFT_TEMPLATE_SHEET.MEMBER_START_COL + colOffset,
      range.getNumRows(),
      range.getNumColumns()
    );

    targetRange.merge();
  });
}

/**
 * ボーダーを範囲に適用
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dailySheet - 日程シート
 * @param {number} columnCount - 列数
 */
function applyBordersToRange(dailySheet, columnCount) {
  const rowCount =
    SHIFT_TEMPLATE_SHEET.ROWS.DATA_END -
    SHIFT_TEMPLATE_SHEET.ROWS.DATA_START +
    1;

  const targetRange = dailySheet.getRange(
    SHIFT_TEMPLATE_SHEET.ROWS.DATA_START,
    SHIFT_TEMPLATE_SHEET.MEMBER_START_COL,
    rowCount,
    columnCount
  );

  applyBorders(targetRange);
}

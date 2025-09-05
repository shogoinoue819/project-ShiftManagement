// ===== 設定定数 =====
const ENABLE_BORDER_PROCESSING = false; // true: ボーダー処理あり, false: ボーダー処理なし

// 曜日別授業割を反映
function reflectLessonTemplate() {
  try {
    Logger.log("🔄 授業割テンプレ反映開始");

    // SSをまとめて取得
    const ss = getSpreadsheet();
    const ui = getUI();

    // ターゲットはシフト作成シート
    const targetSheets = getTargetSheets(ss);
    Logger.log(`📋 対象シート数: ${targetSheets.length}`);

    // 全曜日のテンプレートデータを事前にキャッシュ
    const templateCache = buildTemplateCache(ss);
    Logger.log("📦 テンプレートデータキャッシュ完了");

    // 各日程のシフト作成シートにおいて、
    targetSheets.forEach((dailySheet) => {
      processDailySheetWithCache(dailySheet, templateCache);
    });

    Logger.log("✅ 授業割テンプレ反映完了");
    ui.alert("✅ 授業割テンプレを反映しました！");
  } catch (error) {
    Logger.log(`❌ エラーが発生しました: ${error.message}`);
    throw error;
  }
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
 * 全曜日のテンプレートデータをキャッシュとして構築
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - スプレッドシート
 * @returns {Object} 曜日別のテンプレートデータキャッシュ
 */
function buildTemplateCache(ss) {
  const cache = {};
  const weekdays = ["Mon", "Tue", "Wed", "Thu", "Fri"];

  weekdays.forEach((dayOfWeek) => {
    const lessonTemplateSheet = getLessonTemplateSheet(ss, dayOfWeek);
    if (lessonTemplateSheet) {
      // 最大列数を取得（最初のシートから）
      const firstTargetSheet = getTargetSheets(ss)[0];
      if (firstTargetSheet) {
        const columnCount =
          firstTargetSheet.getLastColumn() -
          SHIFT_TEMPLATE_SHEET.MEMBER_START_COL +
          1;
        const templateData = getLessonTemplateData(
          lessonTemplateSheet,
          columnCount
        );
        cache[dayOfWeek] = templateData;
        Logger.log(`📦 ${dayOfWeek}のテンプレートデータをキャッシュしました`);
      }
    }
  });

  return cache;
}

/**
 * 授業割テンプレートシートからデータを取得
 * @param {GoogleAppsScript.Spreadsheet.Sheet} lessonTemplateSheet - 授業割テンプレートシート
 * @param {number} columnCount - 列数
 * @returns {Object} テンプレートデータ
 */
function getLessonTemplateData(lessonTemplateSheet, columnCount) {
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

  return {
    values: sourceRange.getValues(),
    backgrounds: sourceRange.getBackgrounds(),
    fontColors: sourceRange.getFontColors(),
    fontSizes: sourceRange.getFontSizes(),
    fontWeights: sourceRange.getFontWeights(),
    mergedRanges: sourceRange.getMergedRanges(),
    rowCount: rowCount,
    columnCount: columnCount,
  };
}

/**
 * 各日程シートの処理を実行（キャッシュ使用版）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dailySheet - 日程シート
 * @param {Object} templateCache - テンプレートデータキャッシュ
 */
function processDailySheetWithCache(dailySheet, templateCache) {
  try {
    // シート名を取得
    const sheetName = dailySheet.getName();

    // 日付から曜日を取得
    const dayOfWeek = getDayOfWeekFromSheet(dailySheet);

    // 月〜金に含まれる場合のみ処理
    if (!isWeekday(dayOfWeek)) {
      return;
    }

    // キャッシュから該当曜日のデータを取得
    const templateData = templateCache[dayOfWeek];
    if (!templateData) {
      Logger.log(
        `⚠️ ${dayOfWeek}のテンプレートデータが見つかりません: ${sheetName}`
      );
      return;
    }

    // キャッシュされたテンプレートデータをコピー
    copyTemplateDataFromCache(dailySheet, templateData);
  } catch (error) {
    Logger.log(`❌ シート処理でエラー: ${sheetName} - ${error.message}`);
    throw error;
  }
}

/**
 * 各日程シートの処理を実行（旧版 - 互換性のため残す）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dailySheet - 日程シート
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - スプレッドシート
 */
function processDailySheet(dailySheet, ss) {
  try {
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
      return;
    }

    // テンプレートデータをコピー
    copyTemplateData(dailySheet, lessonTemplateSheet);
  } catch (error) {
    Logger.log(`❌ シート処理でエラー: ${sheetName} - ${error.message}`);
    throw error;
  }
}

/**
 * シートから曜日を取得
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dailySheet - 日程シート
 * @returns {string} 曜日（Mon, Tue, Wed, Thu, Fri, Sat, Sun）
 */
function getDayOfWeekFromSheet(dailySheet) {
  try {
    const date = dailySheet
      .getRange(SHIFT_TEMPLATE_SHEET.DATE_ROW, SHIFT_TEMPLATE_SHEET.DATE_COL)
      .getValue();

    if (!(date instanceof Date)) {
      return null;
    }

    // 曜日略称を取得（Fri, Mon, Tue等）
    const dayOfWeek = Utilities.formatDate(
      date,
      Session.getScriptTimeZone(),
      "EEE"
    );

    return dayOfWeek;
  } catch (error) {
    Logger.log(`❌ 曜日取得でエラー: ${error.message}`);
    throw error;
  }
}

/**
 * 曜日が平日（月〜金）かどうかを判定
 * @param {string} dayOfWeek - 曜日（Fri, Mon, Tue, Wed, Thu, Sat, Sun）
 * @returns {boolean} 平日の場合true
 */
function isWeekday(dayOfWeek) {
  if (!dayOfWeek) {
    return false;
  }

  const weekdayMap = {
    Mon: true,
    Tue: true,
    Wed: true,
    Thu: true,
    Fri: true,
    Sat: false,
    Sun: false,
  };

  const result = weekdayMap[dayOfWeek] || false;

  return result;
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
 * キャッシュされたテンプレートデータを日程シートにコピー
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dailySheet - 日程シート
 * @param {Object} templateData - キャッシュされたテンプレートデータ
 */
function copyTemplateDataFromCache(dailySheet, templateData) {
  // ターゲット範囲を取得
  const targetRange = dailySheet.getRange(
    SHIFT_TEMPLATE_SHEET.ROWS.DATA_START,
    SHIFT_TEMPLATE_SHEET.MEMBER_START_COL,
    templateData.rowCount,
    templateData.columnCount
  );

  // セルの書式とデータをコピー
  copyCellPropertiesFromCache(templateData, targetRange);

  // 結合セルの処理
  handleMergedCellsFromCache(templateData, dailySheet);

  // ボーダーの適用（設定に応じて）
  if (ENABLE_BORDER_PROCESSING) {
    applyBordersToRangeFromCache(dailySheet, templateData);
  }
}

/**
 * テンプレートデータを日程シートにコピー（旧版 - 互換性のため残す）
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
 * キャッシュされたセルの書式とデータをコピー
 * @param {Object} templateData - キャッシュされたテンプレートデータ
 * @param {GoogleAppsScript.Spreadsheet.Range} targetRange - ターゲット範囲
 */
function copyCellPropertiesFromCache(templateData, targetRange) {
  // 背景色の処理（白背景は保持）
  const processedBackgrounds = processBackgroundsFromCache(
    templateData.backgrounds,
    targetRange
  );

  // 一括でプロパティを設定
  targetRange.setBackgrounds(processedBackgrounds);
  targetRange.setValues(templateData.values);
  targetRange.setFontColors(templateData.fontColors);
  targetRange.setFontSizes(templateData.fontSizes);
  targetRange.setFontWeights(templateData.fontWeights);
}

/**
 * セルの書式とデータをコピー（旧版 - 互換性のため残す）
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
 * キャッシュされた背景色を処理（白背景は元の背景を保持）
 * @param {Array} sourceBackgrounds - ソースの背景色配列
 * @param {GoogleAppsScript.Spreadsheet.Range} targetRange - ターゲット範囲
 * @returns {Array} 処理済みの背景色配列
 */
function processBackgroundsFromCache(sourceBackgrounds, targetRange) {
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
 * 背景色を処理（白背景は元の背景を保持）（旧版 - 互換性のため残す）
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
 * キャッシュされた結合セルの処理
 * @param {Object} templateData - キャッシュされたテンプレートデータ
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dailySheet - 日程シート
 */
function handleMergedCellsFromCache(templateData, dailySheet) {
  templateData.mergedRanges.forEach((range) => {
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
 * キャッシュ版のボーダーを範囲に適用
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dailySheet - 日程シート
 * @param {Object} templateData - キャッシュされたテンプレートデータ
 */
function applyBordersToRangeFromCache(dailySheet, templateData) {
  const targetRange = dailySheet.getRange(
    SHIFT_TEMPLATE_SHEET.ROWS.DATA_START,
    SHIFT_TEMPLATE_SHEET.MEMBER_START_COL,
    templateData.rowCount,
    templateData.columnCount
  );

  applyBorders(targetRange);
}

/**
 * 結合セルの処理（旧版 - 互換性のため残す）
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

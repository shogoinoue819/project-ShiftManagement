// ===== 03_utils.js =====
// シフト管理システムのユーティリティ関数群
// シート操作、データ処理、フォーマット、スタイル適用などの共通機能を提供

// ===== 1. シート・UI取得 =====

/**
 * 共通で使用されるシート・UIオブジェクトをまとめて取得
 *
 * 後方互換性のため、従来の5つの値を配列で返す形式も維持しています。
 * パフォーマンスを重視する場合は、個別の関数を使用することを推奨します。
 *
 * @param {Spreadsheet|null} spreadsheet - 対象のスプレッドシート（テスト時用、省略時はアクティブなSS）
 * @returns {Array} [ss, manageSheet, templateSheet, allSheets, ui]
 *
 * @example
 * // 従来の使用方法（5つ全て取得）
 * const [ss, manageSheet, templateSheet, allSheets, ui] = getCommonSheets();
 *
 * // 個別取得の推奨方法
 * const ss = getSpreadsheet();
 * const manageSheet = getManageSheet();
 * const ui = getUI();
 *
 * @see getSpreadsheet, getManageSheet, getTemplateSheet, getAllSheets, getUI
 */
function getCommonSheets(spreadsheet = null) {
  // SSを取得（テスト時はパラメータを使用）
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  // シフト管理シートを取得
  const manageSheet = ss.getSheetByName(SHEET_NAMES.SHIFT_MANAGEMENT);
  // シフトテンプレートシートを取得
  const templateSheet = ss.getSheetByName(SHEET_NAMES.SHIFT_TEMPLATE);
  // 全てのシートを取得
  const allSheets = ss.getSheets();
  // UIを取得
  const ui = SpreadsheetApp.getUi();

  return [ss, manageSheet, templateSheet, allSheets, ui];
}

/**
 * スプレッドシートオブジェクトを取得
 *
 * @param {Spreadsheet|null} spreadsheet - 対象のスプレッドシート（テスト時用、省略時はアクティブなSS）
 * @returns {Spreadsheet} スプレッドシートオブジェクト
 */
function getSpreadsheet(spreadsheet = null) {
  return spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * シフト管理シートを取得
 *
 * @param {Spreadsheet|null} spreadsheet - 対象のスプレッドシート（テスト時用、省略時はアクティブなSS）
 * @returns {Sheet|null} シフト管理シート（存在しない場合はnull）
 */
function getManageSheet(spreadsheet = null) {
  const ss = getSpreadsheet(spreadsheet);
  return ss.getSheetByName(SHEET_NAMES.SHIFT_MANAGEMENT);
}

/**
 * シフトテンプレートシートを取得
 *
 * @param {Spreadsheet|null} spreadsheet - 対象のスプレッドシート（テスト時用、省略時はアクティブなSS）
 * @returns {Sheet|null} シフトテンプレートシート（存在しない場合はnull）
 */
function getTemplateSheet(spreadsheet = null) {
  const ss = getSpreadsheet(spreadsheet);
  return ss.getSheetByName(SHEET_NAMES.SHIFT_TEMPLATE);
}

/**
 * 全てのシートを取得
 *
 * @param {Spreadsheet|null} spreadsheet - 対象のスプレッドシート（テスト時用、省略時はアクティブなSS）
 * @returns {Sheet[]} 全てのシートの配列
 */
function getAllSheets(spreadsheet = null) {
  const ss = getSpreadsheet(spreadsheet);
  return ss.getSheets();
}

/**
 * UIオブジェクトを取得
 *
 * @returns {Ui} UIオブジェクト
 */
function getUI() {
  return SpreadsheetApp.getUi();
}

// SSをまとめて取得（本番環境用）
const [ss, manageSheet, templateSheet, allSheets, ui] = getCommonSheets();

// ===== 2. セル・範囲処理 =====

/**
 * 特定の列の最終行を取得
 *
 * 指定された列でデータが存在する最後の行番号を効率的に取得します。
 * パフォーマンスを考慮し、実際にデータが存在する範囲のみを処理します。
 *
 * @param {Sheet} sheet - 対象のシート
 * @param {number} col - 対象の列番号（1から開始）
 * @returns {number} 最終行番号（データが存在しない場合は0）
 *
 * @example
 * // シフト管理シートのメンバーリスト列の最終行を取得
 * const lastRow = getLastRowInColumn(manageSheet, 5);
 * console.log(`最終行: ${lastRow}`);
 *
 * // データの存在確認
 * if (lastRow > 0) {
 *   const data = sheet.getRange(1, 5, lastRow, 1).getValues();
 * }
 *
 * @note
 * - 列番号は1から開始（Google Apps Scriptの仕様）
 * - 空のセルは無視され、実際にデータが存在する行のみカウント
 * - パフォーマンス向上のため、getMaxRows()ではなくgetLastRow()を使用
 *
 * @see getLastColumnInRow, isValidSheetAndColumn
 */
function getLastRowInColumn(sheet, col) {
  // パラメータの検証
  if (!isValidSheetAndColumn(sheet, col)) {
    return UTILS_CONSTANTS.DEFAULTS.ZERO;
  }

  // 効率的な最終行取得：getLastRow()を使用して範囲を限定
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) return UTILS_CONSTANTS.DEFAULTS.ZERO;

  // 実際にデータが存在する範囲のみを取得
  const values = sheet
    .getRange(UTILS_CONSTANTS.ROWS.START_INDEX, col, lastRow)
    .getValues();
  return findLastNonEmptyRow(values);
}

/**
 * 特定の行の最終列を取得
 *
 * 指定された行でデータが存在する最後の列番号を効率的に取得します。
 * パフォーマンスを考慮し、実際にデータが存在する範囲のみを処理します。
 *
 * @param {Sheet} sheet - 対象のシート
 * @param {number} row - 対象の行番号（1から開始）
 * @returns {number} 最終列番号（データが存在しない場合は0）
 *
 * @example
 * // シフト管理シートの1行目の最終列を取得
 * const lastCol = getLastColumnInRow(manageSheet, 1);
 * console.log(`最終列: ${lastCol}`);
 *
 * @note
 * - 行番号は1から開始（Google Apps Scriptの仕様）
 * - 空のセルは無視され、実際にデータが存在する列のみカウント
 * - パフォーマンス向上のため、getMaxColumns()ではなくgetLastColumn()を使用
 *
 * @see getLastRowInColumn, isValidSheetAndRow
 */
function getLastColumnInRow(sheet, row) {
  // パラメータの検証
  if (!isValidSheetAndRow(sheet, row)) {
    return UTILS_CONSTANTS.DEFAULTS.ZERO;
  }

  // 効率的な最終列取得：getLastColumn()を使用して範囲を限定
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) return UTILS_CONSTANTS.DEFAULTS.ZERO;

  // 実際にデータが存在する範囲のみを取得
  const values = sheet
    .getRange(row, UTILS_CONSTANTS.ROWS.START_INDEX, 1, lastColumn)
    .getValues()[0];
  return findLastNonEmptyColumn(values);
}

/**
 * シートと列の妥当性を検証
 *
 * @param {Sheet} sheet - 検証対象のシート
 * @param {number} col - 検証対象の列番号
 * @returns {boolean} 妥当性の結果
 */
function isValidSheetAndColumn(sheet, col) {
  if (!sheet || !col || col < UTILS_CONSTANTS.ROWS.MIN_INDEX) {
    console.warn("getLastRowInColumn: 無効なパラメータ", {
      sheet: !!sheet,
      col,
    });
    return false;
  }
  return true;
}

/**
 * シートと行の妥当性を検証
 *
 * @param {Sheet} sheet - 検証対象のシート
 * @param {number} row - 検証対象の行番号
 * @returns {boolean} 妥当性の結果
 */
function isValidSheetAndRow(sheet, row) {
  if (!sheet || !row || row < UTILS_CONSTANTS.ROWS.MIN_INDEX) {
    console.warn("getLastColumnInRow: 無効なパラメータ", {
      sheet: !!sheet,
      row,
    });
    return false;
  }
  return true;
}

/**
 * 配列から最後の非空行を検索（効率化）
 *
 * @param {Array<Array>} values - 検索対象の2次元配列
 * @returns {number} 最後の非空行のインデックス（1から開始）
 */
function findLastNonEmptyRow(values) {
  // 逆順で検索して最初に見つかった非空行を返す
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "") {
      return i + 1; // インデックスは0スタートなので+1
    }
  }
  return UTILS_CONSTANTS.DEFAULTS.ZERO; // 空列の場合
}

/**
 * 配列から最後の非空列を検索（効率化）
 *
 * @param {Array} values - 検索対象の配列
 * @returns {number} 最後の非空列のインデックス（1から開始）
 */
function findLastNonEmptyColumn(values) {
  // 逆順で検索して最初に見つかった非空列を返す
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i] !== "") {
      return i + 1; // インデックスは0スタートなので+1
    }
  }
  return UTILS_CONSTANTS.DEFAULTS.ZERO; // 空行の場合
}

// ===== 3. メンバー管理 =====

/**
 * ランダムな6桁のメンバーIDを生成
 *
 * 英数字を組み合わせたユニークなメンバーIDを生成します。
 * パフォーマンスを考慮し、文字列連結を避けて配列で構築します。
 *
 * @returns {string} "usr_" + 6桁のランダム文字列
 *
 * @example
 * // 新しいメンバーIDを生成
 * const memberId = generateRandomMemberId();
 * console.log(`生成されたID: ${memberId}`); // 例: "usr_aB3x9K"
 *
 * // メンバー登録時に使用
 * const newMember = {
 *   id: generateRandomMemberId(),
 *   name: "田中太郎",
 *   email: "tanaka@example.com"
 * };
 *
 * @note
 * - IDの長さはUTILS_CONSTANTS.ID_GENERATION.MEMBER_ID_LENGTHで定義
 * - 使用可能文字: 英字（大文字・小文字）と数字
 * - 重複の可能性は極めて低いが、完全な保証はない
 * - パフォーマンス向上のため、配列構築後にjoin()を使用
 *
 * @see UTILS_CONSTANTS.ID_GENERATION
 */
function generateRandomMemberId() {
  const chars =
    "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  const charsLength = chars.length;
  let randomPart = "";

  // 文字列連結を避けて配列で構築
  const randomChars = [];
  for (let i = 0; i < UTILS_CONSTANTS.ID_GENERATION.MEMBER_ID_LENGTH; i++) {
    randomChars.push(chars.charAt(Math.floor(Math.random() * charsLength)));
  }

  return `usr_${randomChars.join("")}`;
}

/**
 * メンバーリストからデータを取得する共通ヘルパー関数
 *
 * シフト管理シートのメンバーリストから指定された列数のデータを取得します。
 * この関数は他のメンバー関連関数の基盤となり、データ取得の重複を防ぎます。
 *
 * @param {number} [columns=2] - 取得する列数（デフォルト: ID列と氏名列の2列）
 * @param {Sheet} [sheet=manageSheet] - 対象シート（テスト用）
 * @returns {Array<Array>} メンバーデータの2次元配列
 *
 * @example
 * // ID列と氏名列を取得（デフォルト）
 * const memberData = getMemberListData();
 * // 結果: [["usr_abc123", "田中太郎"], ["usr_def456", "佐藤花子"], ...]
 *
 * // ID列のみを取得
 * const idOnly = getMemberListData(1);
 * // 結果: [["usr_abc123"], ["usr_def456"], ...]
 *
 * // 3列分のデータを取得
 * const extendedData = getMemberListData(3);
 * // 結果: [["usr_abc123", "田中太郎", "田中"], ...]
 *
 * // テスト用：特定のシートを指定
 * const testData = getMemberListData(2, mockSheet);
 *
 * @note
 * - パラメータの妥当性チェックを自動実行
 * - データが存在しない場合は空配列を返す
 * - パフォーマンス向上のため、範囲を限定してデータ取得
 * - この関数は他のメンバー関数から呼び出されることを想定
 * - テスト時は外部依存を最小化するため、パラメータでシートを指定可能
 *
 * @see getMemberNameById, getMemberIdByName, getMemberOrderById
 * @see UTILS_CONSTANTS.COLUMNS
 */
function getMemberListData(
  columns = UTILS_CONSTANTS.COLUMNS.ID_AND_NAME,
  sheet = manageSheet
) {
  // パラメータの検証
  if (!isValidMemberListParams(sheet, columns)) {
    return UTILS_CONSTANTS.DEFAULTS.EMPTY_ARRAY;
  }

  const lastRow = getLastRowInColumn(
    sheet,
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
  );

  // データが存在しない場合
  if (!hasValidMemberData(lastRow)) {
    return UTILS_CONSTANTS.DEFAULTS.EMPTY_ARRAY;
  }

  return fetchMemberListData(lastRow, columns, sheet);
}

/**
 * メンバーリストパラメータの妥当性を検証
 *
 * @param {Sheet} sheet - 検証対象のシート
 * @param {number} columns - 検証対象の列数
 * @returns {boolean} 妥当性の結果
 */
function isValidMemberListParams(sheet, columns) {
  if (!sheet || !columns || columns < UTILS_CONSTANTS.ROWS.MIN_INDEX) {
    console.warn("getMemberListData: 無効なパラメータ", {
      sheet: !!sheet,
      columns,
    });
    return false;
  }
  return true;
}

/**
 * メンバーデータの存在確認
 *
 * @param {number} lastRow - 最終行番号
 * @returns {boolean} データの存在確認結果
 */
function hasValidMemberData(lastRow) {
  return lastRow >= SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW;
}

/**
 * メンバーリストデータの取得（効率化）
 *
 * @param {number} lastRow - 最終行番号
 * @param {number} columns - 取得する列数
 * @param {Sheet} sheet - 対象シート
 * @returns {Array<Array>} メンバーデータの2次元配列
 */
function fetchMemberListData(lastRow, columns, sheet) {
  const startRow = SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW;
  const rowCount = lastRow - startRow + 1;

  // 行数が0以下の場合は空配列を返す
  if (rowCount <= 0) {
    return UTILS_CONSTANTS.DEFAULTS.EMPTY_ARRAY;
  }

  return sheet
    .getRange(
      startRow,
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.ID_COL,
      rowCount,
      columns
    )
    .getValues();
}

/**
 * IDから氏名を取得
 *
 * @param {string} id - メンバーID
 * @param {Sheet} [sheet=manageSheet] - 対象シート（テスト用）
 * @returns {string|null} メンバー氏名、見つからない場合はnull
 */
function getMemberNameById(id, sheet = manageSheet) {
  // パラメータの検証
  if (!id) {
    console.warn("getMemberNameById: IDが指定されていません");
    return null;
  }

  const data = getMemberListData(UTILS_CONSTANTS.COLUMNS.ID_AND_NAME, sheet); // ID列と氏名列を取得
  return findMemberNameById(data, id);
}

/**
 * 氏名からIDを取得
 *
 * @param {string} name - メンバー氏名
 * @param {Sheet} [sheet=manageSheet] - 対象シート（テスト用）
 * @returns {string|null} メンバーID、見つからない場合はnull
 */
function getMemberIdByName(name, sheet = manageSheet) {
  // パラメータの検証
  if (!isValidMemberName(name)) {
    return null;
  }

  const data = getMemberListData(UTILS_CONSTANTS.COLUMNS.ID_AND_NAME, sheet); // ID列と氏名列を取得
  return findMemberIdByName(data, name);
}

/**
 * IDからorderを取得
 *
 * @param {string} id - メンバーID
 * @param {Sheet} [sheet=manageSheet] - 対象シート（テスト用）
 * @returns {number} メンバーの順序（0から開始、見つからない場合は-1）
 */
function getMemberOrderById(id, sheet = manageSheet) {
  // パラメータの検証
  if (!id) {
    console.warn("getMemberOrderById: IDが指定されていません");
    return UTILS_CONSTANTS.DEFAULTS.NOT_FOUND;
  }

  const data = getMemberListData(UTILS_CONSTANTS.COLUMNS.ID_ONLY, sheet); // ID列のみ取得
  return findMemberOrderById(data, id);
}

/**
 * メンバー名の妥当性を検証
 *
 * @param {string} name - 検証対象の氏名
 * @returns {boolean} 妥当性の結果
 */
function isValidMemberName(name) {
  if (!name || typeof name !== "string") {
    console.warn("getMemberIdByName: 無効な氏名", { name });
    return false;
  }
  return true;
}

/**
 * データからIDで氏名を検索（効率化）
 *
 * @param {Array<Array>} data - 検索対象のデータ
 * @param {string} id - 検索対象のID
 * @returns {string|null} 見つかった氏名、見つからない場合はnull
 */
function findMemberNameById(data, id) {
  const idStr = String(id);
  // 早期リターンで効率化
  for (const [vId, vName] of data) {
    if (String(vId) === idStr) {
      return vName;
    }
  }
  return null;
}

/**
 * データから氏名でIDを検索（効率化）
 *
 * @param {Array<Array>} data - 検索対象のデータ
 * @param {string} name - 検索対象の氏名
 * @returns {string|null} 見つかったID、見つからない場合はnull
 */
function findMemberIdByName(data, name) {
  const nameStr = String(name);
  // 早期リターンで効率化
  for (const [vId, vName] of data) {
    if (String(vName) === nameStr) {
      return vId;
    }
  }
  return null;
}

/**
 * データからIDでorderを検索（効率化）
 *
 * @param {Array<Array>} data - 検索対象のデータ
 * @param {string} id - 検索対象のID
 * @returns {number} 見つかった順序（0から開始、見つからない場合は-1）
 */
function findMemberOrderById(data, id) {
  const idStr = String(id);
  // 早期リターンで効率化
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === idStr) {
      return i; // 0から始まるorder
    }
  }
  return UTILS_CONSTANTS.DEFAULTS.NOT_FOUND; // 見つからなければ -1
}

/**
 * メンバーマップ作成
 *
 * @param {Sheet} [sheet=manageSheet] - 対象シート（テスト用）
 * @returns {Object} メンバーIDをキーとしたオブジェクト
 */
function createMemberMap(sheet = manageSheet) {
  if (!sheet) {
    console.warn("createMemberMap: 管理シートが取得できません");
    return UTILS_CONSTANTS.DEFAULTS.EMPTY_OBJECT;
  }

  const lastRow = getLastRowInColumn(
    sheet,
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
  );

  // データが存在しない場合
  if (!hasValidMemberData(lastRow)) {
    return UTILS_CONSTANTS.DEFAULTS.EMPTY_OBJECT;
  }

  const memberData = fetchMemberListData(
    lastRow,
    UTILS_CONSTANTS.COLUMNS.ID_AND_NAME,
    sheet
  );
  const urlData = fetchMemberUrlData(lastRow, sheet);

  return buildMemberMap(memberData, urlData);
}

/**
 * メンバーURLデータの取得（効率化）
 *
 * @param {number} lastRow - 最終行番号
 * @param {Sheet} sheet - 対象シート
 * @returns {Array<Array>} URLデータの2次元配列
 */
function fetchMemberUrlData(lastRow, sheet) {
  const startRow = SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW;
  const rowCount = lastRow - startRow + 1;

  // 行数が0以下の場合は空配列を返す
  if (rowCount <= 0) {
    return UTILS_CONSTANTS.DEFAULTS.EMPTY_ARRAY;
  }

  return sheet
    .getRange(
      startRow,
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.URL_COL,
      rowCount,
      UTILS_CONSTANTS.COLUMNS.ID_ONLY
    )
    .getFormulas();
}

/**
 * メンバーマップの構築（効率化）
 *
 * @param {Array<Array>} memberData - メンバーデータ
 * @param {Array<Array>} urlData - URLデータ
 * @returns {Object} メンバーマップ
 */
function buildMemberMap(memberData, urlData) {
  const memberMap = {};
  const length = Math.min(memberData.length, urlData.length);

  // ループを最適化
  for (let i = 0; i < length; i++) {
    const [id, name] = memberData[i];
    if (id && name) {
      // 有効なデータのみ処理
      memberMap[id] = {
        name,
        url: urlData[i][0] || UTILS_CONSTANTS.DEFAULTS.EMPTY_STRING,
      };
    }
  }
  return memberMap; // { id1: { name: ..., url: ... }, ... }
}

// ===== 4. 日付・時間処理 =====

/**
 * 日程リストから指定日付の順序（order）を取得
 *
 * シフト管理シートの日程リストで、指定された日付が何番目に配置されているかを
 * 0ベースのインデックスで返します。日付が見つからない場合は-1を返します。
 *
 * @param {Date|string} date - 検索対象の日付（Date型または文字列）
 * @param {Sheet} [sheet=manageSheet] - 対象シート（テスト用）
 * @returns {number} 日程リストでの順序（0から開始、見つからない場合は-1）
 *
 * @example
 * // Date型で検索
 * const order1 = getDateOrderByDate(new Date(2024, 0, 15)); // 1月15日
 * console.log(`1月15日の順序: ${order1}`); // 例: 10
 *
 * // 文字列で検索
 * const order2 = getDateOrderByDate("1/15");
 * console.log(`1/15の順序: ${order2}`); // 例: 10
 *
 * // 日付が見つからない場合
 * const order3 = getDateOrderByDate("12/31");
 * console.log(`12/31の順序: ${order3}`); // -1
 *
 * @note
 * - 日付の順序は0から開始（配列のインデックス）
 * - 文字列の場合は"M/d"形式を想定
 * - パフォーマンス向上のため、範囲を限定してデータ取得
 * - 日付が見つからない場合はUTILS_CONSTANTS.DEFAULTS.NOT_FOUND（-1）を返す
 *
 * @see getDateList, convertDateToString, findDateOrder
 * @see UTILS_CONSTANTS.DEFAULTS.NOT_FOUND
 */
function getDateOrderByDate(date, sheet = manageSheet) {
  // パラメータの検証
  if (!date) {
    console.warn("getDateOrderByDate: 日付が指定されていません");
    return UTILS_CONSTANTS.DEFAULTS.NOT_FOUND;
  }

  const dateStr = convertDateToString(date);
  if (!dateStr) {
    return UTILS_CONSTANTS.DEFAULTS.NOT_FOUND;
  }

  const lastRow = getLastRowInColumn(
    sheet,
    SHIFT_MANAGEMENT_SHEET.DATE_LIST.COL
  );

  // データが存在しない場合
  if (!hasValidDateData(lastRow)) {
    return UTILS_CONSTANTS.DEFAULTS.NOT_FOUND;
  }

  const dateValues = fetchDateListData(lastRow, sheet);
  return findDateOrder(dateValues, dateStr);
}

/**
 * 日程リスト作成
 *
 * @param {Sheet} [sheet=manageSheet] - 対象シート（テスト用）
 * @returns {Array<Array>} 日程データの2次元配列
 */
function getDateList(sheet = manageSheet) {
  if (!sheet) {
    console.warn("getDateList: 管理シートが取得できません");
    return UTILS_CONSTANTS.DEFAULTS.EMPTY_ARRAY;
  }

  const lastRow = getLastRowInColumn(
    sheet,
    SHIFT_MANAGEMENT_SHEET.DATE_LIST.COL
  );

  // データが存在しない場合
  if (!hasValidDateData(lastRow)) {
    return UTILS_CONSTANTS.DEFAULTS.EMPTY_ARRAY;
  }

  const dateRange = fetchDateListData(lastRow, sheet);
  return processDateListData(dateRange);
}

/**
 * 日付データの存在確認
 *
 * @param {number} lastRow - 最終行番号
 * @returns {boolean} データの存在確認結果
 */
function hasValidDateData(lastRow) {
  return lastRow >= SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW;
}

/**
 * 日付リストデータの取得（効率化）
 *
 * @param {number} lastRow - 最終行番号
 * @param {Sheet} sheet - 対象シート
 * @returns {Array<Array>} 日付データの2次元配列
 */
function fetchDateListData(lastRow, sheet) {
  const startRow = SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW;
  const rowCount = lastRow - startRow + 1;

  // 行数が0以下の場合は空配列を返す
  if (rowCount <= 0) {
    return UTILS_CONSTANTS.DEFAULTS.EMPTY_ARRAY;
  }

  return sheet
    .getRange(
      startRow,
      SHIFT_MANAGEMENT_SHEET.DATE_LIST.COL,
      rowCount,
      UTILS_CONSTANTS.COLUMNS.DATE_ONLY
    )
    .getValues();
}

/**
 * 日付文字列への変換
 *
 * @param {Date|string} date - 変換対象の日付
 * @returns {string|null} 変換された日付文字列、失敗時はnull
 */
function convertDateToString(date) {
  const dateStr =
    date instanceof Date
      ? formatDateToString(date, UTILS_CONSTANTS.DATE_FORMATS.DEFAULT)
      : date;

  if (!dateStr) {
    console.warn("getDateOrderByDate: 日付の変換に失敗しました", { date });
    return null;
  }

  return dateStr;
}

/**
 * 日付orderの検索（効率化）
 *
 * @param {Array<Array>} dateValues - 日付データの配列
 * @param {string} dateStr - 検索対象の日付文字列
 * @returns {number} 見つかった順序（0から開始、見つからない場合は-1）
 */
function findDateOrder(dateValues, dateStr) {
  // 早期リターンで効率化
  for (let i = 0; i < dateValues.length; i++) {
    const d = dateValues[i][0];
    if (d instanceof Date) {
      const currentStr = formatDateToString(
        d,
        UTILS_CONSTANTS.DATE_FORMATS.DEFAULT
      );
      if (currentStr === dateStr) {
        return i; // 0から始まるorder
      }
    }
  }
  return UTILS_CONSTANTS.DEFAULTS.NOT_FOUND; // 見つからなければ -1
}

/**
 * 日付リストデータの処理（効率化）
 *
 * @param {Array<Array>} dateRange - 日付範囲データ
 * @returns {Array<Array>} 処理された日付データ
 */
function processDateListData(dateRange) {
  const result = [];
  const length = dateRange.length;

  // ループを最適化
  for (let i = 0; i < length; i++) {
    const date = dateRange[i][0];
    if (date instanceof Date) {
      result.push([date]);
    }
  }

  return result;
}

// ===== 5. フォーマット・変換 =====

/**
 * 日付を指定されたフォーマットの文字列に変換
 *
 * Date型のオブジェクトを指定されたフォーマットの文字列に変換します。
 * デフォルトでは"M/d"形式（例: "1/15"）で出力されます。
 *
 * @param {Date} date - 変換対象の日付
 * @param {string} [format="M/d"] - 出力フォーマット（Google Apps Scriptの日付フォーマット）
 * @returns {string} フォーマットされた日付文字列、無効な日付の場合は空文字列
 *
 * @example
 * // 基本的な使用方法（デフォルトフォーマット）
 * const date = new Date(2024, 0, 15); // 1月15日
 * const formatted = formatDateToString(date);
 * console.log(formatted); // "1/15"
 *
 * // カスタムフォーマット
 * const longFormat = formatDateToString(date, "yyyy年M月d日");
 * console.log(longFormat); // "2024年1月15日"
 *
 * // 英語フォーマット
 * const englishFormat = formatDateToString(date, "MMM dd, yyyy");
 * console.log(englishFormat); // "Jan 15, 2024"
 *
 * // 無効な日付の場合
 * const invalidDate = new Date("invalid");
 * const result = formatDateToString(invalidDate);
 * console.log(result); // ""
 *
 * @note
 * - フォーマットはGoogle Apps Scriptの日付フォーマット仕様に準拠
 * - タイムゾーンは現在のスクリプトのタイムゾーンを使用
 * - 無効な日付の場合は空文字列を返す
 * - パフォーマンス向上のため、日付の妥当性を事前チェック
 *
 * @see isValidDate, UTILS_CONSTANTS.DATE_FORMATS
 * @see https://developers.google.com/apps-script/reference/utilities/utilities#formatdatedate-timezone-format
 */
function formatDateToString(
  date,
  format = UTILS_CONSTANTS.DATE_FORMATS.DEFAULT
) {
  if (!isValidDate(date)) {
    return UTILS_CONSTANTS.DEFAULTS.EMPTY_STRING;
  }
  return Utilities.formatDate(date, Session.getScriptTimeZone(), format);
}

// sting→dateのフォーマット
function formatStringToDate(str) {
  // パラメータの検証
  if (!isValidDateString(str)) {
    return null;
  }

  try {
    const { month, day } = parseDateString(str);

    if (!isValidMonthAndDay(month, day)) {
      return null;
    }

    const result = createDateFromMonthDay(month, day);

    if (!isValidDate(result)) {
      console.warn("formatStringToDate: 日付の作成に失敗", {
        str,
        month,
        day,
        result,
      });
      return null;
    }

    return result;
  } catch (e) {
    console.error("formatStringToDate: エラーが発生しました", {
      str,
      error: e.message,
    });
    return null;
  }
}

// 列番号からアルファベットへ変換
function convertColumnToLetter(column) {
  // パラメータの検証
  if (!isValidColumnNumber(column)) {
    return UTILS_CONSTANTS.DEFAULTS.EMPTY_STRING;
  }

  return convertColumnToLetterInternal(column);
}

// 時間を日付に連結させる
function normalizeTimeToDate(baseDate, timeValue) {
  // パラメータの検証
  if (!isValidBaseDate(baseDate)) {
    return null;
  }

  if (!timeValue) {
    console.warn("normalizeTimeToDate: 時間値が指定されていません");
    return null;
  }

  // timeValueがDate型の場合
  if (isValidTimeDate(timeValue)) {
    return createTimeDate(baseDate, timeValue);
  }

  // timeValueがstring型の場合
  if (typeof timeValue === "string") {
    return createTimeDateFromString(baseDate, timeValue);
  }

  // 無効な場合は null
  console.warn("normalizeTimeToDate: 無効な時間値", { timeValue });
  return null;
}

// 日付の妥当性を検証
function isValidDate(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    console.warn("formatDateToString: 無効な日付", { date });
    return false;
  }
  return true;
}

// 日付文字列の妥当性を検証
function isValidDateString(str) {
  if (!str || typeof str !== "string") {
    console.warn("formatStringToDate: 無効な文字列", { str });
    return false;
  }
  return true;
}

// 日付文字列の解析
function parseDateString(str) {
  const [month, day] = str.split("/").map(Number);
  return { month, day };
}

// 月と日の妥当性を検証
function isValidMonthAndDay(month, day) {
  if (
    isNaN(month) ||
    isNaN(day) ||
    month < UTILS_CONSTANTS.DATE_LIMITS.MIN_MONTH ||
    month > UTILS_CONSTANTS.DATE_LIMITS.MAX_MONTH ||
    day < UTILS_CONSTANTS.DATE_LIMITS.MIN_DAY ||
    day > UTILS_CONSTANTS.DATE_LIMITS.MAX_DAY
  ) {
    console.warn("formatStringToDate: 無効な日付形式", { month, day });
    return false;
  }
  return true;
}

// 月と日から日付を作成
function createDateFromMonthDay(month, day) {
  const year = new Date().getFullYear(); // 今年の年
  return new Date(year, month - 1, day); // JSの月は0始まり
}

// 列番号の妥当性を検証
function isValidColumnNumber(column) {
  if (
    !column ||
    column < UTILS_CONSTANTS.ROWS.MIN_INDEX ||
    !Number.isInteger(column)
  ) {
    console.warn("convertColumnToLetter: 無効な列番号", { column });
    return false;
  }
  return true;
}

// 列番号からアルファベットへの内部変換処理（効率化）
function convertColumnToLetterInternal(column) {
  let letter = "";
  let temp;

  // ループを最適化
  while (column > 0) {
    temp = (column - 1) % UTILS_CONSTANTS.ID_GENERATION.ALPHABET_BASE;
    letter =
      String.fromCharCode(temp + UTILS_CONSTANTS.ID_GENERATION.ALPHABET_START) +
      letter;
    column = (column - temp - 1) / UTILS_CONSTANTS.ID_GENERATION.ALPHABET_BASE;
  }
  return letter;
}

// 基準日付の妥当性を検証
function isValidBaseDate(baseDate) {
  if (!(baseDate instanceof Date) || isNaN(baseDate.getTime())) {
    console.warn("normalizeTimeToDate: 無効な基準日付", { baseDate });
    return false;
  }
  return true;
}

// 時間日付の妥当性を検証
function isValidTimeDate(timeValue) {
  return timeValue instanceof Date && !isNaN(timeValue.getTime());
}

// 基準日付と時間日付から新しい日付を作成
function createTimeDate(baseDate, timeValue) {
  return new Date(
    baseDate.getFullYear(),
    baseDate.getMonth(),
    baseDate.getDate(),
    timeValue.getHours(),
    timeValue.getMinutes()
  );
}

// 基準日付と時間文字列から新しい日付を作成
function createTimeDateFromString(baseDate, timeValue) {
  const match = timeValue.match(UTILS_CONSTANTS.REGEX_PATTERNS.TIME_FORMAT);
  if (match) {
    const h = Number(match[1]);
    const m = Number(match[2]);
    if (isValidHourAndMinute(h, m)) {
      return new Date(
        baseDate.getFullYear(),
        baseDate.getMonth(),
        baseDate.getDate(),
        h,
        m
      );
    }
  }
  return null;
}

// 時と分の妥当性を検証
function isValidHourAndMinute(h, m) {
  return (
    h >= UTILS_CONSTANTS.TIME_LIMITS.MIN_HOUR &&
    h < UTILS_CONSTANTS.TIME_LIMITS.MAX_HOUR + 1 &&
    m >= UTILS_CONSTANTS.TIME_LIMITS.MIN_MINUTE &&
    m < UTILS_CONSTANTS.TIME_LIMITS.MAX_MINUTE + 1
  );
}

// ===== 6. スタイル・背景処理 =====

/**
 * シートの背景色を一括削除
 *
 * シート内の勤務不可を示す背景色を一括で削除し、デフォルトの背景色に戻します。
 * エラーハンドリングを実装し、処理の安全性を確保しています。
 *
 * @param {Sheet} sheet - 背景色を削除する対象シート
 * @returns {void}
 *
 * @example
 * // シフト管理シートの背景色を削除
 * clearBackgrounds(manageSheet);
 *
 * // 特定のシートの背景色を削除
 * const targetSheet = ss.getSheetByName("シフト希望表");
 * if (targetSheet) {
 *   clearBackgrounds(targetSheet);
 * }
 *
 * // エラーハンドリング付きで実行
 * try {
 *   clearBackgrounds(manageSheet);
 *   console.log("背景色の削除が完了しました");
 * } catch (error) {
 *   console.error("背景色の削除に失敗しました:", error);
 * }
 *
 * @note
 * - 勤務不可背景色（TIME_SETTINGS.UNAVAILABLE_BACKGROUND_COLOR）のみを対象
 * - データ範囲（getDataRange()）内の全セルを処理
 * - エラーが発生した場合はログに記録し、処理を継続
 * - パフォーマンス向上のため、一括で背景色を更新
 *
 * @see processBackgroundColors, TIME_SETTINGS.UNAVAILABLE_BACKGROUND_COLOR
 * @see applyBorders, protectSheet
 */
function clearBackgrounds(sheet) {
  // パラメータの検証
  if (!sheet) {
    console.warn("clearBackgrounds: シートが指定されていません");
    return;
  }

  try {
    const range = sheet.getDataRange();
    const backgrounds = range.getBackgrounds();
    const updatedBackgrounds = processBackgroundColors(backgrounds);
    range.setBackgrounds(updatedBackgrounds);
  } catch (e) {
    console.error("clearBackgrounds: エラーが発生しました", {
      error: e.message,
    });
  }
}

// ボーダーをセット
function applyBorders(range) {
  // パラメータの検証
  if (!range) {
    console.warn("applyBorders: 範囲が指定されていません");
    return;
  }

  try {
    const mergedRanges = range.getMergedRanges();
    mergedRanges.forEach((merged) => {
      if (shouldApplyBorder(merged)) {
        applyBorderToRange(merged);
      }
    });
  } catch (e) {
    console.error("applyBorders: エラーが発生しました", { error: e.message });
  }
}

// 背景色の処理（効率化）
function processBackgroundColors(backgrounds) {
  const rows = backgrounds.length;
  const cols = backgrounds[0]?.length || 0;

  // ループを最適化
  for (let i = 0; i < rows; i++) {
    for (let j = 0; j < cols; j++) {
      const bgColor = backgrounds[i][j];
      // 背景色が勤務不可背景色ならば、
      if (bgColor === TIME_SETTINGS.UNAVAILABLE_BACKGROUND_COLOR) {
        // 背景色をnullにする
        backgrounds[i][j] = null;
      }
    }
  }
  return backgrounds;
}

// ボーダーを適用すべきかどうかを判定
function shouldApplyBorder(merged) {
  const bg = merged.getBackground();
  return (
    bg !== TIME_SETTINGS.UNAVAILABLE_BACKGROUND_COLOR &&
    bg !== UTILS_CONSTANTS.COLORS.WHITE &&
    bg !== null
  );
}

// 範囲にボーダーを適用
function applyBorderToRange(merged) {
  merged.setBorder(
    true,
    true,
    true,
    true,
    true,
    true,
    UTILS_CONSTANTS.COLORS.BLACK,
    SpreadsheetApp.BorderStyle.SOLID
  );
}

// ===== 7. シート保護・セキュリティ =====

/**
 * シートを保護して編集を制限
 *
 * 指定されたシートを保護し、編集権限を制限します。
 * ドメイン編集も無効化し、セキュリティを強化します。
 *
 * @param {Sheet} sheet - 保護する対象シート
 * @param {string} [description="シートの保護"] - 保護の説明文
 * @returns {void}
 *
 * @example
 * // 基本的な使用方法
 * protectSheet(manageSheet);
 *
 * // カスタム説明付きで保護
 * protectSheet(templateSheet, "シフトテンプレートの保護");
 *
 * // 複数シートを保護
 * const sheetsToProtect = [manageSheet, templateSheet];
 * sheetsToProtect.forEach(sheet => {
 *   if (sheet) {
 *     protectSheet(sheet, `${sheet.getName()}の保護`);
 *   }
 * });
 *
 * // エラーハンドリング付きで実行
 * try {
 *   protectSheet(manageSheet);
 *   console.log("シートの保護が完了しました");
 * } catch (error) {
 *   console.error("シートの保護に失敗しました:", error);
 * }
 *
 * @note
 * - 保護されたシートは編集できなくなる
 * - 既存の編集者は自動的に削除される
 * - ドメイン編集権限も無効化される
 * - エラーが発生した場合はログに記録し、処理を継続
 * - 保護の解除は手動で行う必要がある
 *
 * @see clearBackgrounds, applyBorders
 * @see https://developers.google.com/apps-script/reference/spreadsheet/sheet#protect
 */
function protectSheet(sheet, description = "シートの保護") {
  // パラメータの検証
  if (!sheet) {
    console.warn("protectSheet: シートが指定されていません");
    return;
  }

  try {
    const protection = sheet.protect();
    protection.setDescription(description);
    protection.removeEditors(protection.getEditors());
    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }
  } catch (e) {
    console.error("protectSheet: エラーが発生しました", { error: e.message });
  }
}

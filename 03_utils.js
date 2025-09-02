// ===== 03_utils.js =====
// シフト管理システムのユーティリティ関数群
// シート操作、データ処理、フォーマット、スタイル適用などの共通機能を提供

// ===== 1. シート・UI取得 =====

/**
 * スプレッドシートオブジェクトを取得
 * @param {Spreadsheet|null} spreadsheet - 対象のスプレッドシート（省略時はアクティブなSS）
 * @returns {Spreadsheet} スプレッドシートオブジェクト
 */
function getSpreadsheet(spreadsheet = null) {
  return spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * シフト管理シートを取得
 * @param {Spreadsheet|null} spreadsheet - 対象のスプレッドシート（省略時はアクティブなSS）
 * @returns {Sheet|null} シフト管理シート（存在しない場合はnull）
 */
function getManageSheet(spreadsheet = null) {
  const ss = getSpreadsheet(spreadsheet);
  return ss.getSheetByName(SHEET_NAMES.SHIFT_MANAGEMENT);
}

/**
 * シフトテンプレートシートを取得
 * @param {Spreadsheet|null} spreadsheet - 対象のスプレッドシート（省略時はアクティブなSS）
 * @returns {Sheet|null} シフトテンプレートシート（存在しない場合はnull）
 */
function getTemplateSheet(spreadsheet = null) {
  const ss = getSpreadsheet(spreadsheet);
  return ss.getSheetByName(SHEET_NAMES.SHIFT_TEMPLATE);
}

/**
 * 全てのシートを取得
 * @param {Spreadsheet|null} spreadsheet - 対象のスプレッドシート（省略時はアクティブなSS）
 * @returns {Sheet[]} 全てのシートの配列
 */
function getAllSheets(spreadsheet = null) {
  const ss = getSpreadsheet(spreadsheet);
  return ss.getSheets();
}

/**
 * UIオブジェクトを取得
 * @returns {Ui} UIオブジェクト
 */
function getUI() {
  return SpreadsheetApp.getUi();
}

// ===== 2. セル・範囲処理 =====

/**
 * 特定の列の最終行を取得
 * @param {Sheet} sheet - 対象のシート
 * @param {number} col - 対象の列番号（1から開始）
 * @returns {number} 最終行番号（データが存在しない場合は0）
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
 * @param {Sheet} sheet - 対象のシート
 * @param {number} row - 対象の行番号（1から開始）
 * @returns {number} 最終列番号（データが存在しない場合は0）
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
 * 配列から最後の非空行を検索
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
 * 配列から最後の非空列を検索
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
 * メンバー管理システム
 * メンバーID、氏名、URL、ステータスなどの情報を統一的に管理
 */
class MemberManager {
  /**
   * コンストラクタ
   * @param {Sheet} sheet - 管理シート
   */
  constructor(sheet) {
    if (!sheet) {
      throw new Error("MemberManager: シートが指定されていません");
    }
    this.sheet = sheet;
    this.memberMap = null;
    this.lastUpdate = null;
  }

  /**
   * メンバーデータを初期化・更新
   *
   * @param {boolean} [forceRefresh=false] - 強制更新フラグ
   * @returns {boolean} 初期化の成功/失敗
   */
  initialize(forceRefresh = false) {
    try {
      // シートの妥当性チェック
      if (!this.sheet) {
        console.warn("MemberManager: 管理シートが取得できません");
        return false;
      }

      // 強制更新でない場合、既存データの有効性をチェック
      if (!forceRefresh && this.isValidCache()) {
        return true;
      }

      // データの取得と初期化
      const success = this.refreshData();
      if (success) {
        this.lastUpdate = new Date();
        this.isInitialized = true;
      }

      return success;
    } catch (error) {
      console.error("MemberManager.initialize: エラーが発生しました", {
        error: error.message,
      });
      return false;
    }
  }

  /**
   * キャッシュの有効性をチェック
   *
   * @returns {boolean} キャッシュが有効かどうか
   */
  isValidCache() {
    return (
      this.isInitialized &&
      this.memberData &&
      this.memberMap &&
      this.lastUpdate &&
      new Date() - this.lastUpdate < 300000
    ); // 5分以内
  }

  /**
   * メンバーデータを更新
   *
   * @returns {boolean} 更新の成功/失敗
   */
  refreshData() {
    try {
      const lastRow = getLastRowInColumn(
        this.sheet,
        SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
      );

      // データが存在しない場合
      if (!this.hasValidMemberData(lastRow)) {
        this.clearData();
        return false;
      }

      // 一括でデータを取得
      const memberData = this.fetchMemberListData(lastRow);
      const urlData = this.fetchMemberUrlData(lastRow);

      // データの構築
      this.buildMemberData(memberData, urlData);
      return true;
    } catch (error) {
      console.error("MemberManager.refreshData: エラーが発生しました", {
        error: error.message,
      });
      this.clearData();
      return false;
    }
  }

  /**
   * メンバーデータの存在確認
   *
   * @param {number} lastRow - 最終行番号
   * @returns {boolean} データの存在確認結果
   */
  hasValidMemberData(lastRow) {
    return lastRow >= SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW;
  }

  /**
   * メンバーリストデータの取得
   *
   * @param {number} lastRow - 最終行番号
   * @returns {Array<Array>} メンバーデータの2次元配列
   */
  fetchMemberListData(lastRow) {
    const startRow = SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW;
    const rowCount = lastRow - startRow + 1;

    if (rowCount <= 0) {
      return [];
    }

    return this.sheet
      .getRange(
        startRow,
        SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.ID_COL,
        rowCount,
        UTILS_CONSTANTS.COLUMNS.ID_AND_NAME
      )
      .getValues();
  }

  /**
   * メンバーURLデータの取得
   *
   * @param {number} lastRow - 最終行番号
   * @returns {Array<Array>} URLデータの2次元配列
   */
  fetchMemberUrlData(lastRow) {
    const startRow = SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW;
    const rowCount = lastRow - startRow + 1;

    if (rowCount <= 0) {
      return [];
    }

    return this.sheet
      .getRange(
        startRow,
        SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.URL_COL,
        rowCount,
        UTILS_CONSTANTS.COLUMNS.ID_ONLY
      )
      .getFormulas();
  }

  /**
   * メンバーデータの構築
   *
   * @param {Array<Array>} memberData - メンバーデータ
   * @param {Array<Array>} urlData - URLデータ
   */
  buildMemberData(memberData, urlData) {
    this.memberData = memberData;
    this.memberMap = {};
    this.idToIndexMap = new Map();
    this.nameToIdMap = new Map();

    const length = Math.min(memberData.length, urlData.length);

    for (let i = 0; i < length; i++) {
      const [id, name] = memberData[i];
      if (id && name) {
        // メンバーマップの構築
        this.memberMap[id] = {
          name,
          url: urlData[i][0] || UTILS_CONSTANTS.DEFAULTS.EMPTY_STRING,
          order: i,
        };

        // 高速検索用のマップを構築
        this.idToIndexMap.set(String(id), i);
        this.nameToIdMap.set(String(name), id);
      }
    }
  }

  /**
   * データをクリア
   */
  clearData() {
    this.memberData = null;
    this.memberMap = null;
    this.idToIndexMap = null;
    this.nameToIdMap = null;
    this.isInitialized = false;
  }

  /**
   * IDから氏名を取得
   *
   * @param {string} id - メンバーID
   * @returns {string|null} メンバー氏名、見つからない場合はnull
   */
  getNameById(id) {
    if (!this.ensureInitialized() || !id) {
      return null;
    }

    const member = this.memberMap[id];
    return member ? member.name : null;
  }

  /**
   * 氏名からIDを取得
   *
   * @param {string} name - メンバー氏名
   * @returns {string|null} メンバーID、見つからない場合はnull
   */
  getIdByName(name) {
    if (!this.ensureInitialized() || !name) {
      return null;
    }

    return this.nameToIdMap.get(String(name)) || null;
  }

  /**
   * IDからorderを取得
   *
   * @param {string} id - メンバーID
   * @returns {number} メンバーの順序（0から開始、見つからない場合は-1）
   */
  getOrderById(id) {
    if (!this.ensureInitialized() || !id) {
      return UTILS_CONSTANTS.DEFAULTS.NOT_FOUND;
    }

    const index = this.idToIndexMap.get(String(id));
    return index !== undefined ? index : UTILS_CONSTANTS.DEFAULTS.NOT_FOUND;
  }

  /**
   * メンバー情報を取得
   *
   * @param {string} id - メンバーID
   * @returns {Object|null} メンバー情報、見つからない場合はnull
   */
  getMemberInfo(id) {
    if (!this.ensureInitialized() || !id) {
      return null;
    }

    return this.memberMap[id] || null;
  }

  /**
   * 全メンバーのリストを取得
   *
   * @returns {Array<Object>} 全メンバーの情報配列
   */
  getAllMembers() {
    if (!this.ensureInitialized()) {
      return [];
    }

    return Object.values(this.memberMap);
  }

  /**
   * メンバー数を取得
   *
   * @returns {number} メンバー数
   */
  getMemberCount() {
    if (!this.ensureInitialized()) {
      return 0;
    }

    return Object.keys(this.memberMap).length;
  }

  /**
   * 初期化の確認
   *
   * @returns {boolean} 初期化済みかどうか
   */
  ensureInitialized() {
    if (!this.isInitialized) {
      return this.initialize();
    }
    return true;
  }

  /**
   * メンバーの存在確認
   *
   * @param {string} id - メンバーID
   * @returns {boolean} 存在するかどうか
   */
  exists(id) {
    if (!this.ensureInitialized() || !id) {
      return false;
    }

    return id in this.memberMap;
  }

  /**
   * メンバー名の存在確認
   *
   * @param {string} name - メンバー氏名
   * @returns {boolean} 存在するかどうか
   */
  existsByName(name) {
    if (!this.ensureInitialized() || !name) {
      return false;
    }

    return this.nameToIdMap.has(String(name));
  }

  /**
   * メンバーマップを作成
   *
   * @returns {Object} メンバーマップ
   */
  createMemberMap() {
    const memberMap = {};
    const length = Math.min(this.memberData.length, this.urlData.length);

    // ループを最適化
    for (let i = 0; i < length; i++) {
      const [id, name] = this.memberData[i];
      if (id && name) {
        // 有効なデータのみ処理
        memberMap[id] = {
          name,
          url: this.urlData[i][0] || UTILS_CONSTANTS.DEFAULTS.EMPTY_STRING,
        };
      }
    }
    return memberMap; // { id1: { name: ..., url: ... }, ... }
  }
}

// グローバルなメンバーマネージャーインスタンス
let globalMemberManager = null;

/**
 * メンバーマネージャーを取得
 *
 * @param {Sheet} sheet - 管理シート
 * @returns {MemberManager} メンバーマネージャーインスタンス
 */
function getMemberManager(sheet) {
  if (!sheet) {
    console.warn("getMemberManager: シートが指定されていません");
    return null;
  }
  return new MemberManager(sheet);
}

// ===== 後方互換性のための関数 =====

/**
 * ランダムな6桁のメンバーIDを生成
 * @returns {string} "usr_" + 6桁のランダム文字列
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
 * メンバーリストからデータを取得
 * @param {number} [columns=2] - 取得する列数（デフォルト: ID列と氏名列の2列）
 * @param {Sheet} sheet - 対象シート
 * @returns {Array<Array>} メンバーデータの2次元配列
 * @deprecated 新しいMemberManagerクラスの使用を推奨
 */
function getMemberListData(columns, sheet) {
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
 * メンバーIDから氏名を取得
 *
 * @param {string} id - メンバーID
 * @param {Sheet} sheet - 管理シート
 * @returns {string|null} 氏名（見つからない場合はnull）
 */
function getMemberNameById(id, sheet) {
  if (!sheet) {
    console.warn("getMemberNameById: シートが指定されていません");
    return null;
  }

  const memberManager = getMemberManager(sheet);
  if (!memberManager) {
    return null;
  }

  return memberManager.getNameById(id);
}

/**
 * メンバー氏名からIDを取得
 *
 * @param {string} name - メンバー氏名
 * @param {Sheet} sheet - 管理シート
 * @returns {string|null} メンバーID（見つからない場合はnull）
 */
function getMemberIdByName(name, sheet) {
  if (!sheet) {
    console.warn("getMemberIdByName: シートが指定されていません");
    return null;
  }

  const memberManager = getMemberManager(sheet);
  if (!memberManager) {
    return null;
  }

  return memberManager.getIdByName(name);
}

/**
 * メンバーIDから順序（行番号）を取得
 *
 * @param {string} id - メンバーID
 * @param {Sheet} sheet - 管理シート
 * @returns {number} 順序（0から開始、見つからない場合は-1）
 */
function getMemberOrderById(id, sheet) {
  if (!sheet) {
    console.warn("getMemberOrderById: シートが指定されていません");
    return -1;
  }

  const memberManager = getMemberManager(sheet);
  if (!memberManager) {
    return -1;
  }

  return memberManager.getOrderById(id);
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
 * @param {Sheet} sheet - 対象シート
 * @returns {Object} メンバーマップ
 * @deprecated 新しいMemberManagerクラスの使用を推奨
 */
function createMemberMap(sheet) {
  if (!sheet) {
    console.warn("createMemberMap: シートが指定されていません");
    return UTILS_CONSTANTS.DEFAULTS.EMPTY_OBJECT;
  }

  const memberManager = getMemberManager(sheet);
  return memberManager.memberMap || UTILS_CONSTANTS.DEFAULTS.EMPTY_OBJECT;
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

/**
 * メンバー情報を取得する
 *
 * 指定された行からメンバーの名前とファイルIDを取得します。
 * URLからファイルIDを抽出し、オブジェクトとして返します。
 *
 * @param {number} row - メンバーリストの行番号
 * @param {Sheet} sheet - 対象のシート
 * @returns {Object|null} メンバー情報 {name: string, fileId: string} または null（失敗時）
 *
 * @example
 * // 基本的な使用方法
 * const manageSheet = getManageSheet();
 * const memberInfo = getMemberInfo(5, manageSheet);
 * if (memberInfo) {
 *   console.log(`名前: ${memberInfo.name}, ファイルID: ${memberInfo.fileId}`);
 * }
 *
 * // 特定のシートから取得
 * const memberInfo = getMemberInfo(5, otherSheet);
 *
 * // エラーハンドリング付きで実行
 * const manageSheet = getManageSheet();
 * const memberInfo = getMemberInfo(5, manageSheet);
 * if (!memberInfo) {
 *   console.error("メンバー情報の取得に失敗しました");
 *   return;
 * }
 */
function getMemberInfo(row, sheet) {
  if (!sheet) {
    console.warn("getMemberInfo: シートが指定されていません");
    Logger.log(`❌ getMemberInfo: シートが指定されていません`);
    return null;
  }

  try {
    // 氏名とURLを取得
    const name = sheet
      .getRange(row, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.NAME_COL)
      .getValue();

    const url = sheet
      .getRange(row, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.URL_COL)
      .getFormula();

    // URLからファイルIDを抽出
    const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!match || !match[1]) {
      console.warn(`getMemberInfo: URL抽出失敗: ${name} (URL: ${url})`);
      Logger.log(`❌ URL抽出失敗: ${name}`);
      return null;
    }

    const fileId = match[1];

    const result = {
      name: name,
      fileId: fileId,
    };

    return result;
  } catch (e) {
    console.error("getMemberInfo: エラーが発生しました", {
      error: e.message,
      row: row,
    });
    Logger.log(`❌ getMemberInfoエラー: 行${row} - ${e}`);
    return null;
  }
}

// ===== 4. 日付・時間処理 =====

/**
 * 日程リストから指定日付の順序（order）を取得
 *
 * シフト管理シートの日程リストで、指定された日付が何番目に配置されているかを
 * 0ベースのインデックスで返します。日付が見つからない場合は-1を返します。
 *
 * @param {Date|string} date - 検索対象の日付（Date型または文字列）
 * @param {Sheet} sheet - 対象シート
 * @returns {number} 日程リストでの順序（0から開始、見つからない場合は-1）
 *
 * @example
 * // Date型で検索
 * const order1 = getDateOrderByDate(new Date(2024, 0, 15), manageSheet); // 1月15日
 * console.log(`1月15日の順序: ${order1}`); // 例: 10
 *
 * // 文字列で検索
 * const order2 = getDateOrderByDate("1/15", manageSheet);
 * console.log(`1/15の順序: ${order2}`); // 例: 10
 *
 * // 日付が見つからない場合
 * const order3 = getDateOrderByDate("12/31", manageSheet);
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
function getDateOrderByDate(date, sheet) {
  // パラメータの検証
  if (!date) {
    console.warn("getDateOrderByDate: 日付が指定されていません");
    return UTILS_CONSTANTS.DEFAULTS.NOT_FOUND;
  }

  if (!sheet) {
    console.warn("getDateOrderByDate: シートが指定されていません");
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
 * @param {Sheet} sheet - 対象シート
 * @returns {Array<Array>} 日程データの2次元配列
 */
function getDateList(sheet) {
  if (!sheet) {
    console.warn("getDateList: 管理シートが指定されていません");
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
 * @param {number} lastRow - 最終行番号
 * @returns {boolean} データの存在確認結果
 */
function hasValidDateData(lastRow) {
  return lastRow >= SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW;
}

/**
 * 日付リストデータの取得
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
 * 日付orderの検索
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
 * 日付リストデータの処理
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
 * @param {Date} date - 変換対象の日付
 * @param {string} [format="M/d"] - 出力フォーマット（Google Apps Scriptの日付フォーマット）
 * @returns {string} フォーマットされた日付文字列、無効な日付の場合は空文字列
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
 * @param {Sheet} sheet - 背景色を削除する対象シート
 * @returns {void}
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

/**
 * ボーダーをセット
 * @param {Range} range - 対象の範囲
 * @returns {void}
 */
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

/**
 * 背景色の処理
 * @param {Array<Array>} backgrounds - 背景色の2次元配列
 * @returns {Array<Array>} 処理された背景色配列
 */
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

/**
 * ボーダーを適用すべきかどうかを判定
 * @param {Range} merged - マージされた範囲
 * @returns {boolean} ボーダーを適用すべきかどうか
 */
function shouldApplyBorder(merged) {
  const bg = merged.getBackground();
  return (
    bg !== TIME_SETTINGS.UNAVAILABLE_BACKGROUND_COLOR &&
    bg !== UTILS_CONSTANTS.COLORS.WHITE &&
    bg !== null
  );
}

/**
 * 範囲にボーダーを適用
 * @param {Range} merged - マージされた範囲
 * @returns {void}
 */
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
 * @param {Sheet} sheet - 保護する対象シート
 * @param {string} [description="シートの保護"] - 保護の説明文
 * @returns {boolean} 保護の成功/失敗
 */
function protectSheet(sheet, description = "シートの保護") {
  // パラメータの検証
  if (!sheet) {
    console.warn("protectSheet: シートが指定されていません");
    Logger.log(`❌ protectSheet: シートが指定されていません`);
    return false;
  }

  try {
    const protection = sheet.protect();
    protection.setDescription(description);
    protection.removeEditors(protection.getEditors());
    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }

    // 保護が正しく設定されたか確認
    const protections = sheet.getProtections(
      SpreadsheetApp.ProtectionType.SHEET
    );

    if (protections.length === 0) {
      console.error("protectSheet: 保護の設定に失敗しました");
      Logger.log(`❌ 保護の設定に失敗: ${sheet.getName()}`);
      return false;
    }

    return true;
  } catch (e) {
    console.error("protectSheet: エラーが発生しました", { error: e.message });
    Logger.log(`❌ protectSheetエラー: ${sheet.getName()} - ${e}`);
    return false;
  }
}

/**
 * 特定のシートを保護する
 *
 * 指定されたスプレッドシート内の特定のシートを保護します。
 * 既存の保護がある場合は一旦解除してから新しく保護を設定します。
 *
 * @param {Spreadsheet} spreadsheet - 対象のスプレッドシート
 * @param {string} sheetName - 保護するシート名
 * @param {string} description - 保護の説明文
 * @param {string} [memberName=""] - メンバー名（ログ用）
 * @returns {boolean} 保護が成功したかどうか
 *
 * @example
 * // 基本的な使用方法
 * const success = protectSheetByName(targetFile, "シフト希望表", "チェックによるロック");
 *
 * // メンバー名付きで実行
 * const success = protectSheetByName(targetFile, "シフト希望表", "チェックによるロック", "田中太郎");
 *
 * // エラーハンドリング付きで実行
 * if (!protectSheetByName(targetFile, "シフト希望表", "保護")) {
 *   console.error("シートの保護に失敗しました");
 * }
 */
function protectSheetByName(
  spreadsheet,
  sheetName,
  description,
  memberName = ""
) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    console.warn(
      `protectSheetByName: ${sheetName}が見つかりません${
        memberName ? `: ${memberName}` : ""
      }`
    );
    Logger.log(`❌ シートが見つかりません: ${sheetName}`);
    return false;
  }

  try {
    // 既存の保護がある場合のみ削除（パフォーマンス改善）
    const protections = sheet.getProtections(
      SpreadsheetApp.ProtectionType.SHEET
    );

    if (protections.length > 0) {
      protections.forEach((p) => p.remove());
    }

    // 新しく保護を設定
    const success = protectSheet(sheet, description);

    if (!success) {
      console.error(
        `protectSheetByName: 保護の設定に失敗しました${
          memberName ? ` (${memberName})` : ""
        }`
      );
      Logger.log(`❌ 保護の設定に失敗: ${sheetName}`);
      return false;
    }

    return true;
  } catch (e) {
    console.error(
      `protectSheetByName: エラーが発生しました${
        memberName ? ` (${memberName})` : ""
      }`,
      { error: e.message }
    );
    Logger.log(`❌ protectSheetByNameエラー: ${sheetName} - ${e}`);
    return false;
  }
}

/**
 * 特定のシートの保護を解除する
 *
 * 指定されたスプレッドシート内の特定のシートの保護を解除します。
 *
 * @param {Spreadsheet} spreadsheet - 対象のスプレッドシート
 * @param {string} sheetName - 保護を解除するシート名
 * @param {string} [memberName=""] - メンバー名（ログ用）
 * @returns {boolean} 保護解除が成功したかどうか
 *
 * @example
 * // 基本的な使用方法
 * const success = unprotectSheetByName(targetFile, "シフト希望表");
 *
 * // メンバー名付きで実行
 * const success = unprotectSheetByName(targetFile, "シフト希望表", "田中太郎");
 *
 * // エラーハンドリング付きで実行
 * if (!unprotectSheetByName(targetFile, "シフト希望表")) {
 *   console.error("シートの保護解除に失敗しました");
 * }
 */
function unprotectSheetByName(spreadsheet, sheetName, memberName = "") {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    console.warn(
      `unprotectSheetByName: ${sheetName}が見つかりません${
        memberName ? `: ${memberName}` : ""
      }`
    );
    return false;
  }

  try {
    // 保護を削除
    const protections = sheet.getProtections(
      SpreadsheetApp.ProtectionType.SHEET
    );
    protections.forEach((p) => p.remove());
    return true;
  } catch (e) {
    console.error(
      `unprotectSheetByName: エラーが発生しました${
        memberName ? ` (${memberName})` : ""
      }`,
      { error: e.message }
    );
    return false;
  }
}

// ===== 8. メンバーシート管理 =====

/**
 * メンバーマネージャーの初期化（共通処理）
 *
 * @returns {MemberManager|null} 初期化されたメンバーマネージャー、失敗時はnull
 */
function initializeMemberManager() {
  try {
    const manageSheet = getManageSheet();
    const memberManager = getMemberManager(manageSheet);

    if (!memberManager.ensureInitialized()) {
      Logger.log("❌ メンバーデータの初期化に失敗しました");
      return null;
    }

    const memberMap = memberManager.memberMap;
    if (!memberMap || Object.keys(memberMap).length === 0) {
      Logger.log("❌ メンバーデータが取得できませんでした");
      return null;
    }

    return memberManager;
  } catch (e) {
    Logger.log(`❌ メンバー管理初期化エラー: ${e.message}`);
    return null;
  }
}

/**
 * メンバーシートの整理（共通処理）
 *
 * @param {Spreadsheet} memberSS - メンバーのスプレッドシート
 * @param {string} memberName - メンバー名
 */
function organizeMemberSheets(memberSS, memberName) {
  try {
    const allSheets = memberSS.getSheets();
    const targetSheetNames = [
      SHEET_NAMES.SHIFT_FORM, // ①シフト希望表
      SHEET_NAMES.SHIFT_FORM_INFO, // ②今後の勤務希望
      SHEET_NAMES.SHIFT_FORM_PREVIOUS, // ③前回分
    ];

    // 不要なシートを削除
    for (const sheet of allSheets) {
      const sheetName = sheet.getName();
      if (!targetSheetNames.includes(sheetName)) {
        try {
          memberSS.deleteSheet(sheet);
          Logger.log(`🗑️ ${memberName} さんの不要シート削除: "${sheetName}"`);
        } catch (deleteError) {
          Logger.log(
            `⚠️ ${memberName} さんのシート削除失敗: "${sheetName}" - ${deleteError.message}`
          );
        }
      }
    }

    // シートの順番を整理
    let currentPosition = 1;
    for (const targetSheetName of targetSheetNames) {
      const targetSheet = memberSS.getSheetByName(targetSheetName);
      if (targetSheet) {
        try {
          memberSS.setActiveSheet(targetSheet);
          memberSS.moveActiveSheet(currentPosition);
          currentPosition++;
        } catch (moveError) {
          Logger.log(
            `⚠️ ${memberName} さんのシート移動失敗: "${targetSheetName}" - ${moveError.message}`
          );
        }
      }
    }

    Logger.log(`✅ ${memberName} さんのシート整理完了`);
  } catch (e) {
    Logger.log(`⚠️ ${memberName} さんのシート整理でエラー: ${e.message}`);
  }
}

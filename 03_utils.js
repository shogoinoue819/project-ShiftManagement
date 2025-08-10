// ===== 一括取得 =====

// SS・管理シート・テンプレートシート・全てのシート・UIをまとめて取得
function getCommonSheets() {
  // SSを取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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

// SSをまとめて取得
const [ss, manageSheet, templateSheet, allSheets, ui] = getCommonSheets();

// ===== セル処理 =====

// 特定の列の最終行を取得する関数
function getLastRowInCol(sheet, col) {
  const values = sheet.getRange(1, col, sheet.getMaxRows()).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "") {
      return i + 1; // インデックスは0スタートなので+1
    }
  }
  return 0; // 空列の場合
}

// 特定の行の最終列を取得する関数
function getLastColInRow(sheet, row) {
  const values = sheet
    .getRange(row, 1, 1, sheet.getMaxColumns())
    .getValues()[0];
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i] !== "") {
      return i + 1; // インデックスは0スタートなので+1
    }
  }
  return 0; // 空行の場合
}

// ===== ランダムメソッド =====

// ランダムな6桁のメンバーIDを生成するメソッド
function generateMemberId() {
  const chars =
    "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  let randomPart = "";
  for (let i = 0; i < 6; i++) {
    randomPart += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return `usr_${randomPart}`;
}

// ===== 管理シートリスト処理 =====

// IDから氏名を取得するメソッド
function getNameById(id) {
  // 最終行を取得
  const lastRow = getLastRowInCol(
    manageSheet,
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
  );
  // メンバーリストからデータを取得[[id, 氏名], ...]
  const data = manageSheet
    .getRange(
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW,
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.ID_COL,
      lastRow - SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW + 1,
      2
    )
    .getValues();
  // 引数IDとIDが一致するデータを探す
  const match = data.find(([vId]) => String(vId) === String(id));
  // 一致したデータの氏名、なければnullを返す
  return match ? match[1] : null;
}

// 氏名からIDを取得するメソッド
function getIdByName(name) {
  // 最終行を取得
  const lastRow = getLastRowInCol(
    manageSheet,
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
  );
  // メンバーリストからデータを取得[[id, 氏名], ...]
  const data = manageSheet
    .getRange(
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW,
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.ID_COL,
      lastRow - SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW + 1,
      2
    )
    .getValues();
  // 引数氏名と氏名が一致するものを探す
  const match = data.find(([, vName]) => String(vName) === String(name));
  // 一致したデータのID、なければnullを返す
  return match ? match[0] : null;
}

// IDからorderを取得するメソッド
function getOrderById(id) {
  // 最終行を取得
  const lastRow = getLastRowInCol(
    manageSheet,
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
  );
  // メンバーリストからデータを取得[[id], ...]
  const data = manageSheet
    .getRange(
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW,
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.ID_COL,
      lastRow - SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW + 1,
      1
    )
    .getValues();
  // IDを検索して、その行インデックス（0始まり）を取得
  const index = data.findIndex(([vId]) => String(vId) === String(id));
  //見つからなければ -1 を返す（見つかれば 0,1,2...）
  return index;
}

// 日程リストからorderを取得
function getOrderByDate(date) {
  // 引数がDate型なら "M/d" 形式に変換、文字列ならそのまま使う
  const dateStr = date instanceof Date ? formatDateToString(date, "M/d") : date;
  // 日程リストの最終行を取得
  const lastRow = getLastRowInCol(
    manageSheet,
    SHIFT_MANAGEMENT_SHEET.DATE_LIST.COL
  );
  // 日程データを取得
  const dateValues = manageSheet
    .getRange(
      SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW,
      SHIFT_MANAGEMENT_SHEET.DATE_LIST.COL,
      lastRow - SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW + 1,
      1
    )
    .getValues();

  // 日程リストで一致した日付のorderを取得
  const index = dateValues.findIndex(([d]) => {
    if (!(d instanceof Date)) return false;
    const currentStr = formatDateToString(d, "M/d");
    return currentStr === dateStr;
  });
  return index; // 0から始まるorder、見つからなければ -1
}

// ===== フォーマッター =====

// date→stingのフォーマット(デフォルトは"M/d")
function formatDateToString(date, format = "M/d") {
  if (!(date instanceof Date)) return "";
  return Utilities.formatDate(date, Session.getScriptTimeZone(), format);
}

// sting→dateのフォーマット
function formatStringToDate(str) {
  try {
    const [month, day] = str.split("/").map(Number);
    const year = new Date().getFullYear(); // 今年の年
    return new Date(year, month - 1, day); // JSの月は0始まり
  } catch (e) {
    return null;
  }
}

// 列番号からアルファベットへ変換
function columnToLetter(column) {
  let temp = "";
  let letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

// 時間を日付に連結させる
function normalizeTimeToDate(baseDate, timeValue) {
  // timeValueがDate型の場合
  if (timeValue instanceof Date && !isNaN(timeValue)) {
    return new Date(
      baseDate.getFullYear(),
      baseDate.getMonth(),
      baseDate.getDate(),
      timeValue.getHours(),
      timeValue.getMinutes()
    );
  }
  // timeValueがstring型の場合
  if (typeof timeValue === "string") {
    // 半角数字 + コロン形式 "H:mm" or "HH:mm"
    const timeRegex = /^(\d{1,2}):(\d{2})$/;
    const match = timeValue.match(timeRegex);
    if (match) {
      const h = Number(match[1]);
      const m = Number(match[2]);
      if (h >= 0 && h < 24 && m >= 0 && m < 60) {
        return new Date(
          baseDate.getFullYear(),
          baseDate.getMonth(),
          baseDate.getDate(),
          h,
          m
        );
      }
    }
  }
  // 無効な場合は null
  return null;
}

// ===== 汎用メソッド =====

// 日程リスト作成
function getDateList() {
  const lastRow = getLastRowInCol(
    manageSheet,
    SHIFT_MANAGEMENT_SHEET.DATE_LIST.COL
  );
  const dateRange = manageSheet
    .getRange(
      SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW,
      SHIFT_MANAGEMENT_SHEET.DATE_LIST.COL,
      lastRow - SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW + 1,
      1
    )
    .getValues();
  return dateRange
    .flat()
    .filter((date) => date instanceof Date)
    .map((date) => [date]);
}

// メンバーマップ作成
function createMemberMap() {
  const lastRow = getLastRowInCol(
    manageSheet,
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
  );
  // IDと氏名だけ取得（必要列は2列）
  const data = manageSheet
    .getRange(
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW,
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.ID_COL,
      lastRow - SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW + 1,
      2
    )
    .getValues();
  // URL列をformulasで取得（HYPERLINK保持）
  const urls = manageSheet
    .getRange(
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW,
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.URL_COL,
      lastRow - SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW + 1,
      1
    )
    .getFormulas();

  const memberMap = {};
  data.forEach(([id, name], i) => {
    memberMap[id] = {
      name,
      url: urls[i][0],
    };
  });
  return memberMap; // { id1: { name: ..., url: ... }, ... }
}

// ===== 背景処理 =====

// 背景色を削除
function clearBackgrounds(sheet) {
  // シートの背景を取得
  const range = sheet.getDataRange();
  const backgrounds = range.getBackgrounds();

  // 全てのセルにおいて、
  for (let i = 0; i < backgrounds.length; i++) {
    for (let j = 0; j < backgrounds[i].length; j++) {
      const bgColor = backgrounds[i][j];
      // 背景色が勤務不可背景色ならば、
      if (bgColor === UNAVAILABLE_COLOR) {
        // 背景色をnullにする
        backgrounds[i][j] = null;
      }
    }
  }
  range.setBackgrounds(backgrounds);
}

// ボーダーをセット
function applyBorders(range) {
  // 結合範囲を取得
  const mergedRanges = range.getMergedRanges();
  // 結合範囲の各セルにおいて
  mergedRanges.forEach((merged) => {
    // 背景が灰色または白でない場合にだけ枠線を適用
    const bg = merged.getBackground();
    if (bg !== UNAVAILABLE_COLOR && bg !== "#ffffff" && bg !== null) {
      merged.setBorder(
        true,
        true,
        true,
        true,
        true,
        true,
        "#000000",
        SpreadsheetApp.BorderStyle.SOLID
      );
    }
  });
}

// ===== シート処理 =====

// シートを保護
function protectSheet(sheet, description = "シートの保護") {
  const protection = sheet.protect();
  protection.setDescription(description);
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}

// 新規メンバー作成
function createNewMember() {
  // SSをまとめて取得
  const ss = getSpreadsheet();
  const manageSheet = getManageSheet();
  const ui = getUI();

  // ===== 管理者入力 =====

  // 氏名の入力
  const responseName = ui.prompt(
    "追加するメンバーの氏名を入力してください",
    ui.ButtonSet.OK_CANCEL
  );
  if (responseName.getSelectedButton() !== ui.Button.OK) {
    ui.alert("キャンセルされました");
    return;
  }
  // 空白などをトリミングして入力された氏名を取得
  const inputName = responseName.getResponseText().trim();
  // 未入力ならアラート
  if (!inputName) {
    ui.alert("❌ 氏名が入力されていません");
    return;
  }
  // 既に同名シートが存在するか確認
  if (ss.getSheetByName(inputName)) {
    ui.alert(`❌「${inputName}」さんのシートは既に存在しています`);
    return;
  }
  // 氏名をセット
  const name = inputName;

  // ===== 個別ファイルの作成 =====

  // シフト希望表個別フォルダを取得
  const folder = DriveApp.getFolderById(PERSONAL_FORM_FOLDER_ID);
  // テンプレートファイルから提出用ファイルを作成(ファイル名は"シフト希望表_{氏名}")
  const newFile = DriveApp.getFileById(TEMPLATE_FILE_ID).makeCopy(
    `${SHEET_NAMES.SHIFT_FORM}_${name}`,
    folder
  );

  // 提出用ファイルのSSを取得
  const newSS = SpreadsheetApp.openById(newFile.getId());
  // シフト希望表シートの作成とリネーム
  const newSheet = newSS.getSheetByName(SHEET_NAMES.SHIFT_FORM);
  if (newSheet) {
    newSheet.setName(SHEET_NAMES.SHIFT_FORM);
  } else {
    throw new Error(`❌ シート '${SHEET_NAMES.SHIFT_FORM}' が見つかりません。`);
  }
  // シートに氏名を記入
  newSheet
    .getRange(
      SHIFT_FORM_TEMPLATE.HEADER.ROW,
      SHIFT_FORM_TEMPLATE.HEADER.NAME_COL
    )
    .setValue(name);

  // 勤務希望表シートの作成とリネーム
  const infoSheet = newSS.getSheetByName(SHEET_NAMES.SHIFT_FORM_INFO);
  if (infoSheet) {
    infoSheet.setName(SHEET_NAMES.SHIFT_FORM_INFO);
  } else {
    throw new Error(
      `❌ シート '${SHEET_NAMES.SHIFT_FORM_INFO}' が見つかりません。`
    );
  }

  // 既存シートがあれば削除
  newSS.getSheets().forEach((s) => {
    if (
      s.getName() !== SHEET_NAMES.SHIFT_FORM &&
      s.getName() !== SHEET_NAMES.SHIFT_FORM_INFO
    )
      newSS.deleteSheet(s);
  });

  // 「前回分」シートを複製して追加(保護も追加)
  const originalSheet = newSS.getSheetByName(SHEET_NAMES.SHIFT_FORM);
  const previousSheet = originalSheet.copyTo(newSS);
  previousSheet.setName(SHEET_NAMES.SHIFT_FORM_PREVIOUS);
  protectSheet(previousSheet);

  // シートの並び順を調整する（original → info → previous）
  newSS.setActiveSheet(originalSheet);
  newSS.moveActiveSheet(1); // 一番先頭へ
  newSS.setActiveSheet(infoSheet);
  newSS.moveActiveSheet(2); // 二番目
  newSS.setActiveSheet(previousSheet);
  newSS.moveActiveSheet(3); // 三番目

  // ===== 個別シートの作成 =====

  // 個別シートを作成
  const personalSheet = ss.insertSheet(name);
  // 提出用ファイルのurlを取得
  const personalUrl = newSS.getUrl();

  // 個別シートにデータをimport
  const endColumnLetter = convertColumnToLetter(
    SHIFT_FORM_TEMPLATE.DATA.NOTE_COL
  );
  const formula = `=IMPORTRANGE("${personalUrl}", "${SHEET_NAMES.SHIFT_FORM}!A1:${endColumnLetter}")`;
  personalSheet.getRange(1, 1).setFormula(formula);

  // ===== メンバーリストに追加 =====

  // 再利用する参照を一度だけ算出
  const nameColLetter = convertColumnToLetter(
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.NAME_COL
  );
  const checkCell =
    convertColumnToLetter(SHIFT_FORM_TEMPLATE.HEADER.CHECK_COL) +
    SHIFT_FORM_TEMPLATE.HEADER.ROW;
  const infoCell =
    convertColumnToLetter(SHIFT_FORM_TEMPLATE.HEADER.INFO_COL) +
    SHIFT_FORM_TEMPLATE.HEADER.ROW;

  // 最終行を取得
  const lastRow = getLastRowInColumn(
    manageSheet,
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
  );
  // 新規メンバーを追加する行を最終行の1つ下としてセット
  const newRow = lastRow + 1;

  // IDを生成
  const uniqueId = generateRandomMemberId();
  setupMemberRow(
    manageSheet,
    newRow,
    uniqueId,
    name,
    personalUrl,
    nameColLetter,
    checkCell,
    infoCell
  );

  // ===== 前回用管理シート =====

  // 前回用管理シートを取得
  const manageSheetPre = ss.getSheetByName(
    SHEET_NAMES.SHIFT_MANAGEMENT_PREVIOUS
  );
  if (manageSheetPre) {
    // 最終行を取得
    const lastRowPre = getLastRowInColumn(
      manageSheetPre,
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
    );
    // 新規メンバーを追加する行を最終行の1つ下としてセット
    const newRowPre = lastRowPre + 1;

    setupMemberRow(
      manageSheetPre,
      newRowPre,
      uniqueId,
      name,
      personalUrl,
      nameColLetter,
      checkCell,
      infoCell
    );
  } else {
    Logger.log(
      `⚠️ 前回用管理シートが見つかりません: ${SHEET_NAMES.SHIFT_MANAGEMENT_PREVIOUS}`
    );
  }

  ui.alert(`✅「${name}」さんの個別ファイルと個別シートを作成しました！`);
}

// ===== ヘルパー関数 =====
function setupMemberRow(
  sheet,
  row,
  uniqueId,
  name,
  personalUrl,
  nameColLetter,
  checkCell,
  infoCell
) {
  // ID・氏名
  sheet
    .getRange(row, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.ID_COL)
    .setValue(uniqueId);
  sheet
    .getRange(row, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.NAME_COL)
    .setValue(name);

  // 提出ステータス（個別シートのチェック参照）
  sheet
    .getRange(row, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.SUBMIT_COL)
    .setFormula(
      `=IF(INDIRECT("'" & ${nameColLetter}${row} & "'!${checkCell}") = TRUE, "${STATUS_STRINGS.SUBMIT.TRUE}", "${STATUS_STRINGS.SUBMIT.FALSE}")`
    );

  // チェックボックス・反映ステータス
  sheet
    .getRange(row, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.CHECK_COL)
    .setValue(false);
  sheet
    .getRange(row, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.REFLECT_COL)
    .setValue(STATUS_STRINGS.REFLECT.FALSE);

  // URL（HYPERLINK）
  sheet
    .getRange(row, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.URL_COL)
    .setFormula(`=HYPERLINK("${personalUrl}", "シートリンク")`);

  // 勤務日数・労働時間（週1〜4）
  sheet
    .getRange(row, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.WORK_DATES_1_COL)
    .setFormula(WORK_CALCULATION_FORMULAS.WEEK_1.DATES);
  sheet
    .getRange(row, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.WORK_TIMES_1_COL)
    .setFormula(WORK_CALCULATION_FORMULAS.WEEK_1.TIMES);
  sheet
    .getRange(row, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.WORK_DATES_2_COL)
    .setFormula(WORK_CALCULATION_FORMULAS.WEEK_2.DATES);
  sheet
    .getRange(row, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.WORK_TIMES_2_COL)
    .setFormula(WORK_CALCULATION_FORMULAS.WEEK_2.TIMES);
  sheet
    .getRange(row, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.WORK_DATES_3_COL)
    .setFormula(WORK_CALCULATION_FORMULAS.WEEK_3.DATES);
  sheet
    .getRange(row, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.WORK_TIMES_3_COL)
    .setFormula(WORK_CALCULATION_FORMULAS.WEEK_3.TIMES);
  sheet
    .getRange(row, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.WORK_DATES_4_COL)
    .setFormula(WORK_CALCULATION_FORMULAS.WEEK_4.DATES);
  sheet
    .getRange(row, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.WORK_TIMES_4_COL)
    .setFormula(WORK_CALCULATION_FORMULAS.WEEK_4.TIMES);

  // 勤務日数希望
  sheet
    .getRange(row, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.WORK_DATES_REQUEST_COL)
    .setFormula(`=INDIRECT("'" & ${nameColLetter}${row} & "'!${infoCell}")`);

  // 新規追加行の枠線を削除（表機能による自動枠線を無効化）
  removeTableBorders(sheet, row);
}

/**
 * 指定行の枠線を削除（表機能による自動枠線を無効化）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {number} row - 対象行
 */
function removeTableBorders(sheet, row) {
  try {
    // メンバーリストの全列範囲を取得
    const startCol = SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.ID_COL;
    const endCol = SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.WORK_DATES_REQUEST_COL;

    // 対象行の全列範囲を取得
    const targetRange = sheet.getRange(row, startCol, 1, endCol - startCol + 1);

    // 上の枠線のみを削除（表の構造は維持）
    targetRange.setBorder(
      false, // top - 上の枠線を削除
      null, // left - 左の枠線は変更しない
      null, // bottom - 下の枠線は変更しない（表の構造維持）
      null, // right - 右の枠線は変更しない
      null, // vertical - 縦の枠線は変更しない
      null, // horizontal - 横の枠線は変更しない
      "white", // color
      SpreadsheetApp.BorderStyle.SOLID
    );

    Logger.log(`✅ 行${row}の枠線を削除しました`);
  } catch (error) {
    Logger.log(`⚠️ 枠線削除でエラー: ${error.message}`);
  }
}

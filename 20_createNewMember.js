// 新規メンバー作成
function createNewMember() {
  // SSをまとめて取得
  const ss = getSpreadsheet();
  const manageSheet = getManageSheet();
  const templateSheet = getTemplateSheet();
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

  // // メールアドレスの入力
  // const inputEmail = ui.prompt("追加するメンバーのメールアドレスを入力してください", ui.ButtonSet.OK_CANCEL);
  // if (inputEmail.getSelectedButton() !== ui.Button.OK) {
  //   ui.alert("キャンセルされました");
  //   return;
  // }
  // // 空白などをトリミングして入力されたメールアドレスをセット
  // const email = inputEmail.getResponseText().trim();

  // ===== 個別ファイルの作成 =====

  // シフト希望表個別フォルダを取得
  const folder = DriveApp.getFolderById(PERSONAL_FORM_FOLDER_ID);
  // テンプレートファイルから提出用ファイルを作成(ファイル名は"シフト希望表_{氏名}")
  const newFile = DriveApp.getFileById(TEMPLATE_FILE_ID).makeCopy(
    `${SHEET_NAMES.SHIFT_FORM}_${name}`,
    folder
  );

  // try {
  //   // 編集権限を付加
  //   newFile.addEditor(email);

  // } catch (e) {
  //   Logger.log(`❌ ${name} さんのファイル共有に失敗しました: ${e}`);
  //   ui.alert(`❌ メールアドレスの空白もしくは不適により、${name} さんのファイル共有は未実行です。\nエラー内容: ${e.message}`);
  // }

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
  const endColumnLetter = columnToLetter(SHIFT_FORM_TEMPLATE.DATA.NOTE_COL);
  const formula = `=IMPORTRANGE("${personalUrl}", "${SHEET_NAMES.SHIFT_FORM}!A1:${endColumnLetter}")`;
  personalSheet.getRange(1, 1).setFormula(formula);

  // ===== メンバーリストに追加 =====

  // 最終行を取得
  const lastRow = getLastRowInCol(
    manageSheet,
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
  );
  // 新規メンバーを追加する行を最終行の1つ下としてセット
  const newRow = lastRow + 1;

  // IDを生成
  const uniqueId = generateMemberId();
  // IDをセット
  manageSheet
    .getRange(newRow, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.ID_COL)
    .setValue(uniqueId);
  // 氏名をセット
  manageSheet
    .getRange(newRow, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.NAME_COL)
    .setValue(name);
  // 提出ステータスをセット(個別シートとリンク)
  const nameColLetter = columnToLetter(
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.NAME_COL
  );
  const checkCell =
    columnToLetter(SHIFT_FORM_TEMPLATE.HEADER.CHECK_COL) +
    SHIFT_FORM_TEMPLATE.HEADER.ROW;
  manageSheet
    .getRange(newRow, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.SUBMIT_COL)
    .setFormula(
      `=IF(INDIRECT("'" & ${nameColLetter}${newRow} & "'!${checkCell}") = TRUE, "${STATUS_STRINGS.SUBMIT.TRUE}", "${STATUS_STRINGS.SUBMIT.FALSE}")`
    );
  // チェックボックスをセット
  manageSheet
    .getRange(newRow, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.CHECK_COL)
    .setValue(false);
  // 反映ステータスをセット
  manageSheet
    .getRange(newRow, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.REFLECT_COL)
    .setValue(STATUS_STRINGS.REFLECT.FALSE);
  // URLをセット(HYPERLINK形式)
  manageSheet
    .getRange(newRow, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.URL_COL)
    .setFormula(`=HYPERLINK("${personalUrl}", "シートリンク")`);
  // 勤務日数①をセット
  manageSheet
    .getRange(newRow, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.WORK_DATES_1_COL)
    .setFormula(WORK_CALCULATION_FORMULAS.WEEK_1.DATES);
  // 労働時間①をセット
  manageSheet
    .getRange(newRow, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.WORK_TIMES_1_COL)
    .setFormula(WORK_CALCULATION_FORMULAS.WEEK_1.TIMES);
  // 勤務日数②をセット
  manageSheet
    .getRange(newRow, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.WORK_DATES_2_COL)
    .setFormula(WORK_CALCULATION_FORMULAS.WEEK_2.DATES);
  // 労働時間②をセット
  manageSheet
    .getRange(newRow, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.WORK_TIMES_2_COL)
    .setFormula(WORK_CALCULATION_FORMULAS.WEEK_2.TIMES);
  // 勤務日数③をセット
  manageSheet
    .getRange(newRow, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.WORK_DATES_3_COL)
    .setFormula(WORK_CALCULATION_FORMULAS.WEEK_3.DATES);
  // 労働時間③をセット
  manageSheet
    .getRange(newRow, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.WORK_TIMES_3_COL)
    .setFormula(WORK_CALCULATION_FORMULAS.WEEK_3.TIMES);
  // 勤務日数④をセット
  manageSheet
    .getRange(newRow, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.WORK_DATES_4_COL)
    .setFormula(WORK_CALCULATION_FORMULAS.WEEK_4.DATES);
  // 労働時間④をセット
  manageSheet
    .getRange(newRow, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.WORK_TIMES_4_COL)
    .setFormula(WORK_CALCULATION_FORMULAS.WEEK_4.TIMES);
  // 勤務日数希望をセット
  const infoCell =
    columnToLetter(SHIFT_FORM_TEMPLATE.HEADER.INFO_COL) +
    SHIFT_FORM_TEMPLATE.HEADER.ROW;
  manageSheet
    .getRange(newRow, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.WORK_DATES_REQUEST_COL)
    .setFormula(`=INDIRECT("'" & ${nameColLetter}${newRow} & "'!${infoCell}")`);

  // // メアドをセット
  // manageSheet.getRange(newRow, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.EMAIL_COL).setValue(email);

  // ===== 前回用管理シート =====

  // 前回用管理シートを取得
  const manageSheetPre = ss.getSheetByName(
    SHEET_NAMES.SHIFT_MANAGEMENT_PREVIOUS
  );
  // 最終行を取得
  const lastRowPre = getLastRowInCol(
    manageSheetPre,
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
  );
  // 新規メンバーを追加する行を最終行の1つ下としてセット
  const newRowPre = lastRowPre + 1;

  // IDをセット
  manageSheetPre
    .getRange(newRowPre, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.ID_COL)
    .setValue(uniqueId);
  // 氏名をセット
  manageSheetPre
    .getRange(newRowPre, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.NAME_COL)
    .setValue(name);
  // 提出ステータスをセット(個別シートとリンク)
  manageSheetPre
    .getRange(newRowPre, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.SUBMIT_COL)
    .setFormula(
      `=IF(INDIRECT("'" & ${nameColLetter}${newRowPre} & "'!${checkCell}") = TRUE, "${STATUS_STRINGS.SUBMIT.TRUE}", "${STATUS_STRINGS.SUBMIT.FALSE}")`
    );
  // チェックボックスをセット
  manageSheetPre
    .getRange(newRowPre, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.CHECK_COL)
    .setValue(false);
  // 反映ステータスをセット
  manageSheetPre
    .getRange(newRowPre, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.REFLECT_COL)
    .setValue(STATUS_STRINGS.REFLECT.FALSE);
  // URLをセット(HYPERLINK形式)
  manageSheetPre
    .getRange(newRowPre, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.URL_COL)
    .setFormula(`=HYPERLINK("${personalUrl}", "シートリンク")`);
  // 勤務日数①をセット
  manageSheetPre
    .getRange(newRowPre, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.WORK_DATES_1_COL)
    .setFormula(WORK_CALCULATION_FORMULAS.WEEK_1.DATES);
  // 労働時間①をセット
  manageSheetPre
    .getRange(newRowPre, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.WORK_TIMES_1_COL)
    .setFormula(WORK_CALCULATION_FORMULAS.WEEK_1.TIMES);
  // 勤務日数②をセット
  manageSheetPre
    .getRange(newRowPre, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.WORK_DATES_2_COL)
    .setFormula(WORK_CALCULATION_FORMULAS.WEEK_2.DATES);
  // 労働時間②をセット
  manageSheetPre
    .getRange(newRowPre, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.WORK_TIMES_2_COL)
    .setFormula(WORK_CALCULATION_FORMULAS.WEEK_2.TIMES);
  // 勤務日数③をセット
  manageSheetPre
    .getRange(newRowPre, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.WORK_DATES_3_COL)
    .setFormula(WORK_CALCULATION_FORMULAS.WEEK_3.DATES);
  // 労働時間③をセット
  manageSheetPre
    .getRange(newRowPre, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.WORK_TIMES_3_COL)
    .setFormula(WORK_CALCULATION_FORMULAS.WEEK_3.TIMES);
  // 勤務日数④をセット
  manageSheetPre
    .getRange(newRowPre, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.WORK_DATES_4_COL)
    .setFormula(WORK_CALCULATION_FORMULAS.WEEK_4.DATES);
  // 労働時間④をセット
  manageSheetPre
    .getRange(newRowPre, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.WORK_TIMES_4_COL)
    .setFormula(WORK_CALCULATION_FORMULAS.WEEK_4.TIMES);
  // 勤務日数希望をセット
  manageSheetPre
    .getRange(
      newRowPre,
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.WORK_DATES_REQUEST_COL
    )
    .setFormula(
      `=INDIRECT("'" & ${nameColLetter}${newRowPre} & "'!${infoCell}")`
    );

  ui.alert(`✅「${name}」さんの個別ファイルと個別シートを作成しました！`);
}

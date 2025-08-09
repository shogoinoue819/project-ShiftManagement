// チェックボックスを押すたびにロック関数を動作させる
function onEdit(e) {

  // チェックボックスが押された行列を取得
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();

  // チェック欄でチェックされた場合
  if (col === COLUMN_CHECK && row >= ROW_START) {
    if (e.value === 'TRUE') {
      lockSelectedMember(row);
    } else if (e.value === 'FALSE') {
      unlockSelectedMember(row);
    }
  }
  Logger.log(`onEdit 発火: row=${row}, col=${col}, value=${value}`);
}



// 選択されたメンバーをロック
function lockSelectedMember(row) {

  // SSをまとめて取得
  const [ss, manageSheet, templateSheet, allSheets, ui] = getCommonSheets();

  // 氏名とURLを取得
  const name = manageSheet.getRange(row, COLUMN_NAME).getValue();
  const url = manageSheet.getRange(row, COLUMN_URL).getFormula();

  // URLからファイルIDを抽出
  const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (!match || !match[1]) {
    Logger.log(`⚠️ URL抽出失敗: ${name}`);
    return null;
  }
  // ファイルIDを取得
  const fileId = match[1];

  try {
    // ファイルIDから提出用SSを取得
    const targetFile = SpreadsheetApp.openById(fileId);

    // シフト希望表シートのロック
    const targetSheet = targetFile.getSheetByName(FORM_SHEET_NAME);
    if (targetSheet) {
      // 既存の保護がある場合は一旦解除
      const protections = targetSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      protections.forEach(p => p.remove());
      // 新しく保護を設定
      protectSheet(targetSheet, "チェックによるロック");
    } else {
      Logger.log(`⚠️ シフト希望表が見つかりません: ${name}`);
    }

    // 今後の勤務希望シートのロック
    const infoSheet = targetFile.getSheetByName(FORM_INFO_SHEET_NAME);
    if (infoSheet) {
      // 既存の保護がある場合は解除（情報シート）
      const infoProtections = infoSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      infoProtections.forEach(p => p.remove());
      // 新しく保護を設定（情報シート）
      protectSheet(infoSheet, "チェックによるロック（今後の勤務希望）");
    } else {
      Logger.log(`⚠️ 情報シートが見つかりません: ${name}`);
    }

    Logger.log(`🔒 ${name} をロックしました`);
  } catch (e) {
    Logger.log(`❌ ロック失敗: ${name} - ${e}`);
  }
}



// 選択されたメンバーのロックを解除
function unlockSelectedMember(row) {

  const [ss, manageSheet, templateSheet, allSheets, ui] = getCommonSheets();

  // 氏名とURLを取得
  const name = manageSheet.getRange(row, COLUMN_NAME).getValue();
  const url = manageSheet.getRange(row, COLUMN_URL).getFormula();

  const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (!match || !match[1]) {
    Logger.log(`⚠️ URL抽出失敗: ${name}`);
    return null;
  }
  const fileId = match[1];

  try {
    // ファイルIDから提出用スプレッドシートを取得
    const targetFile = SpreadsheetApp.openById(fileId);

    // フォームシート（シフト希望表）を取得
    const targetSheet = targetFile.getSheetByName(FORM_SHEET_NAME);
    if (!targetSheet) {
      Logger.log(`⚠️ シフト希望表が見つかりません: ${name}`);
      return;
    }
    // 情報シート（フォーム情報）を取得
    const infoSheet = targetFile.getSheetByName(FORM_INFO_SHEET_NAME);
    if (!infoSheet) {
      Logger.log(`⚠️ 情報シートが見つかりません: ${name}`);
      return;
    }
    // シフト希望表の保護を削除
    const protections1 = targetSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    protections1.forEach(p => p.remove());
    // 情報シートの保護を削除
    const protections2 = infoSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    protections2.forEach(p => p.remove());
    
    Logger.log(`🔓 ${name} のロックを解除しました`);
    ui.alert(`🔓 ${name}さんのロックを解除しました`);
  } catch (e) {
    Logger.log(`❌ アンロック失敗: ${name} - ${e}`);
  }
}



// 提出済みのメンバーを全てチェックする
function checkAllSubmittedMembers() {

  // SSをまとめて取得
  const [ss, manageSheet, templateSheet, allSheets, ui] = getCommonSheets();

  // 最終行を取得
  const lastRow = getLastRowInCol(manageSheet, COLUMN_START);
  // メンバーリストからデータを取得[[id, name, shiftName, submit, check, reflect], ...]
  const data = manageSheet.getRange(ROW_START, COLUMN_START, lastRow - ROW_START + 1, COLUMN_REFLECT - COLUMN_START + 1).getValues();

  // 人数カウンター
  let count = 0;
  // データの各メンバーにおいて、
  data.forEach((row, i) => {
    // 提出ステータスとチェックを取得
    const submitStatus = row[COLUMN_SUBMIT - COLUMN_START]; 
    const isChecked = row[COLUMN_CHECK - COLUMN_START];
    // 提出済みかつチェックされていなければ、
    if (submitStatus === SUBMIT_TRUE && isChecked !== true) {
      // ロック処理
      lockSelectedMember(ROW_START + i);
      // シートにチェックを入れる
      manageSheet.getRange(ROW_START + i, COLUMN_CHECK).setValue(true); 
      // 人数を1人増やす
      count++;
    }
  });

  if (count === 0) {
    ui.alert(`❌ 新たにチェックできるメンバーはいません`);
  } else {
    ui.alert(`✅ 提出済みのメンバー${count}人をチェックしました`);
  }

}

// PDFとしてエクスポート
function exportPDF() {
  // SSをまとめて取得
  const [ss, manageSheet, templateSheet, allSheets, ui] = getCommonSheets();

  // 作成済みシフトPDFフォルダを取得
  const pdfFolder = DriveApp.getFolderById(SHIFT_PDF_FOLDER_ID);
  // 作成済みシフトSSフォルダを取得
  const ssFolder = DriveApp.getFolderById(SHIFT_SS_FOLDER_ID);

  // 日付形式のシートだけ抽出
  const targetSheets = allSheets.filter((sheet) =>
    /^\d{1,2}\/\d{1,2}$/.test(sheet.getName())
  );
  if (targetSheets.length === 0) {
    Logger.log("対象のシフトシートがありません。");
    return;
  }

  // 最初のシートの日付を取得し、ファイル名に使う
  const startSheetName = targetSheets[0].getName();
  const [smonth, sday] = startSheetName.split("/");
  const startDate = `${THIS_YEAR}/${String(smonth).padStart(2, "0")}/${String(
    sday
  ).padStart(2, "0")}`;

  // 最後のシートの日付を取得し、ファイル名に使う
  const endSheetName = targetSheets[targetSheets.length - 1].getName();
  const [emonth, eday] = endSheetName.split("/");
  const endDate = `${THIS_YEAR}/${String(emonth).padStart(2, "0")}/${String(
    eday
  ).padStart(2, "0")}`;

  // ファイルを作成
  const newSpreadsheet = SpreadsheetApp.create(
    `シフト_${startDate}~${endDate}`
  );
  newSpreadsheet.setSpreadsheetLocale("ja_JP");
  const newSpreadsheetId = newSpreadsheet.getId();
  const newFile = DriveApp.getFileById(newSpreadsheetId);

  // 指定フォルダに移動
  ssFolder.addFile(newFile);
  DriveApp.getRootFolder().removeFile(newFile);

  // 各日程のシフト作成シートにおいて、
  targetSheets.forEach((sheet) => {
    // シートをコピーして作成
    const copiedSheet = sheet.copyTo(newSpreadsheet);
    // 名前をコピーしてセット
    copiedSheet.setName(sheet.getName());
    // シフト希望メモを非表示
    copiedSheet.hideRows(
      SHIFT_ROW_START_TIME,
      SHIFT_ROW_NOTE - SHIFT_ROW_START_TIME + 1
    );
    // シフト希望反映背景色をクリア
    clearBackgrounds(copiedSheet);
    // 行の高さを1.5倍に設定
    const originalHeight = copiedSheet.getRowHeight(SHIFT_ROW_START);
    const newHeight = Math.floor(originalHeight * 1.5);
    copiedSheet.setRowHeights(
      SHIFT_ROW_MEMBERS,
      SHIFT_ROW_END - SHIFT_ROW_MEMBERS + 1,
      newHeight
    );
  });

  // デフォルトの空白シート（初期生成されるもの）を削除
  const newSheets = newSpreadsheet.getSheets();
  if (newSheets.length > targetSheets.length) {
    newSpreadsheet.deleteSheet(newSheets[0]); // 最初の空白シート
  }

  // PDF出力用のURL生成
  const url =
    "https://docs.google.com/spreadsheets/d/" +
    newSpreadsheetId +
    "/export?format=pdf" +
    "&portrait=false" + // 横向き
    "&size=A4" + // A4サイズ
    "&fitw=true" + // 幅にフィット（これだけでは不十分）
    "&scale=4" + // ❗全体を1ページに収める
    "&sheetnames=false" + // シート名は非表示にするならfalse
    "&printtitle=false" +
    "&pagenumbers=false" +
    "&gridlines=false" +
    "&fzr=false";

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: "Bearer " + token,
    },
  });

  // PDFファイルを作成し、ファイリング
  const pdfBlob = response
    .getBlob()
    .setName(`シフト_${startDate}~${endDate}.pdf`);
  pdfFolder.createFile(pdfBlob);

  // シフト希望メモを再表示
  // const newnewSheets = newSpreadsheet.getSheets();
  // newnewSheets.forEach(sheet => {
  // sheet.showRows(SHIFT_ROW_TIME_START, SHIFT_ROW_WORKING - SHIFT_ROW_TIME_START + 1);
  // });

  ui.alert("✅ 完成したシフトのエクスポートを完了しました！");
}

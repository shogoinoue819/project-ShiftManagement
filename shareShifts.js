// 共有用ファイルに完成済みシフトを反映
function shareShifts() {
  // SSをまとめて取得
  const [ss, manageSheet, templateSheet, allSheets, ui] = getCommonSheets();

  // シフト共有用ファイルを取得
  const shareFile = SpreadsheetApp.openById(SHARE_FILE_ID);

  // シフト日程のデータを取得
  const data = manageSheet
    .getRange(
      MANAGE_DATE_ROW_START,
      MANAGE_DATE_COLUMN,
      getLastRowInCol(manageSheet, MANAGE_DATE_COLUMN) -
        MANAGE_DATE_ROW_START +
        1,
      MANAGE_SHARE_COLUMN - MANAGE_DATE_COLUMN + 1
    )
    .getValues();

  // 作成済みシフトPDFフォルダを取得
  const pdfFolder = DriveApp.getFolderById(SHIFT_PDF_FOLDER_ID);
  // 作成済みシフトSSフォルダを取得
  const ssFolder = DriveApp.getFolderById(SHIFT_SS_FOLDER_ID);

  // ファイルを作成
  const now = new Date();
  const timestamp = Utilities.formatDate(
    now,
    "Asia/Tokyo",
    "yyyy-MM-dd_HH-mm-ss"
  );
  const newSpreadsheet = SpreadsheetApp.create(`シフト作成日時_${timestamp}`);
  newSpreadsheet.setSpreadsheetLocale("ja_JP");
  const newSpreadsheetId = newSpreadsheet.getId();
  const newFile = DriveApp.getFileById(newSpreadsheetId);

  // 指定フォルダに移動
  ssFolder.addFile(newFile);
  DriveApp.getRootFolder().removeFile(newFile);

  // 共有日数カウンタ
  let sharedCount = 0;

  // 取得した日程において、
  data.forEach((row, i) => {
    const [date, isComplete, isShare] = row;

    // 完成済みかつ未共有の日付のみ処理
    if (isComplete === true && isShare === SHARE_FALSE) {
      // Stringにフォーマット
      const dateStr = formatDateToString(date);
      const sheetName = dateStr;

      try {
        // ===== スプレッドシート共有の処理 =====

        // その日付のシフト作成シートを取得
        const dailySheet = ss.getSheetByName(sheetName);
        if (!dailySheet)
          throw new Error(`シート「${sheetName}」が見つかりません`);

        // コピー先に同名シートがあれば削除
        const existingSheet = shareFile.getSheetByName(sheetName);
        if (existingSheet) shareFile.deleteSheet(existingSheet);
        // コピーしてリネーム
        const copiedSheet = dailySheet.copyTo(shareFile).setName(sheetName);

        // シフト希望メモを非表示
        copiedSheet.hideRows(
          SHIFT_ROW_START_TIME,
          SHIFT_ROW_NOTE - SHIFT_ROW_START_TIME + 1
        );
        // シフト希望反映背景色をクリア
        clearBackgrounds(copiedSheet);

        // ===== PDF共有の処理 =====

        // シートをコピーして作成
        const pdfSheet = dailySheet.copyTo(newSpreadsheet);
        // 名前をコピーしてセット
        pdfSheet.setName(dailySheet.getName());
        // シフト希望メモを非表示
        pdfSheet.hideRows(
          SHIFT_ROW_START_TIME,
          SHIFT_ROW_NOTE - SHIFT_ROW_START_TIME + 1
        );
        // シフト希望反映背景色をクリア
        clearBackgrounds(pdfSheet);
        // 行の高さを1.5倍に設定
        const originalHeight = pdfSheet.getRowHeight(SHIFT_ROW_START);
        const newHeight = Math.floor(originalHeight * 1.5);
        pdfSheet.setRowHeights(
          SHIFT_ROW_MEMBERS,
          SHIFT_ROW_END - SHIFT_ROW_MEMBERS + 1,
          newHeight
        );

        // 「共有済」に更新
        manageSheet
          .getRange(i + MANAGE_DATE_ROW_START, MANAGE_SHARE_COLUMN)
          .setValue(SHARE_TRUE);
        // カウンタ増加
        sharedCount++;

        Logger.log(`${dateStr}: 完了`);
      } catch (e) {
        Logger.log("エラー: " + e.message);
      }
    }
  });

  // 共有する日程があれば、
  if (sharedCount > 0) {
    // 共有ファイルから全てのシートを取得
    const sheets = shareFile.getSheets();

    // M/d形式の日付だけを抽出して、日付順にソート
    const sortedSheets = sheets
      .filter((sheet) => /^\d{1,2}\/\d{1,2}$/.test(sheet.getName()))
      .map((sheet) => ({
        sheet: sheet,
        date: formatStringToDate(sheet.getName()),
      }))
      .filter((obj) => obj.date !== null)
      .sort((a, b) => a.date - b.date);

    // 並べ替え（index順にset）
    sortedSheets.forEach((obj, index) => {
      shareFile.setActiveSheet(obj.sheet);
      shareFile.moveActiveSheet(index + 1); // 0番はマスターシート等がある可能性を想定して1始まりに
    });

    // デフォルトの空白シート（初期生成されるもの）を削除
    const newSheets = newSpreadsheet.getSheets();
    if (newSheets.length > sharedCount) {
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
    const pdfBlob = response.getBlob().setName(`シフト作成日時_${timestamp}`);
    pdfFolder.createFile(pdfBlob);

    ui.alert(
      `✅ 完成済みのシフトを ${sharedCount} 日分、共有ファイルに反映しました！`
    );
  } else {
    // 作成したスプレッドシートを削除（ゴミ箱に移動）
    DriveApp.getFileById(newSpreadsheet.getId()).setTrashed(true);

    ui.alert("⚠️ 現在、共有対象のシフトはありませんでした。");
  }
}

// その日程のシフトのみ共有用ファイルに反映
function shareOnlyOneShift() {
  // ===== スプレッドシート共有の処理 =====

  // 今開いているシート名から日程を取得
  const dailySheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const dateStr = dailySheet.getName();

  // シフト共有用ファイルを取得
  const shareFile = SpreadsheetApp.openById(SHARE_FILE_ID);

  // コピー先に同名シートがあれば削除
  const existingSheet = shareFile.getSheetByName(dateStr);
  if (existingSheet) shareFile.deleteSheet(existingSheet);
  // コピーしてリネーム
  const copiedSheet = dailySheet.copyTo(shareFile).setName(dateStr);

  // シフト希望メモを非表示
  copiedSheet.hideRows(
    SHIFT_ROW_START_TIME,
    SHIFT_ROW_NOTE - SHIFT_ROW_START_TIME + 1
  );
  // シフト希望反映背景色をクリア
  clearBackgrounds(copiedSheet);

  // 「共有済」に更新
  manageSheet
    .getRange(
      getOrderByDate(dateStr) + MANAGE_DATE_ROW_START,
      MANAGE_SHARE_COLUMN
    )
    .setValue(SHARE_TRUE);

  // 共有ファイルから全てのシートを取得
  const sheets = shareFile.getSheets();
  // M/d形式の日付だけを抽出して、日付順にソート
  const sortedSheets = sheets
    .filter((sheet) => /^\d{1,2}\/\d{1,2}$/.test(sheet.getName()))
    .map((sheet) => ({
      sheet: sheet,
      date: formatStringToDate(sheet.getName()),
    }))
    .filter((obj) => obj.date !== null)
    .sort((a, b) => a.date - b.date);

  // 並べ替え（index順にset）
  sortedSheets.forEach((obj, index) => {
    shareFile.setActiveSheet(obj.sheet);
    shareFile.moveActiveSheet(index + 1); // 0番はマスターシート等がある可能性を想定して1始まりに
  });

  // ===== PDF共有の処理 =====

  // 作成済みシフトPDFフォルダを取得
  const pdfFolder = DriveApp.getFolderById(SHIFT_PDF_FOLDER_ID);
  // 作成済みシフトSSフォルダを取得
  const ssFolder = DriveApp.getFolderById(SHIFT_SS_FOLDER_ID);

  // ファイルを作成
  const now = new Date();
  const timestamp = Utilities.formatDate(
    now,
    "Asia/Tokyo",
    "yyyy-MM-dd_HH-mm-ss"
  );
  const newSpreadsheet = SpreadsheetApp.create(`シフト作成日時_${timestamp}`);
  newSpreadsheet.setSpreadsheetLocale("ja_JP");
  const newSpreadsheetId = newSpreadsheet.getId();
  const newFile = DriveApp.getFileById(newSpreadsheetId);
  // 指定フォルダに移動
  ssFolder.addFile(newFile);
  DriveApp.getRootFolder().removeFile(newFile);

  // シートをコピーして作成
  const pdfSheet = dailySheet.copyTo(newSpreadsheet);
  // 名前をコピーしてセット
  pdfSheet.setName(dailySheet.getName());
  // シフト希望メモを非表示
  pdfSheet.hideRows(
    SHIFT_ROW_START_TIME,
    SHIFT_ROW_NOTE - SHIFT_ROW_START_TIME + 1
  );
  // シフト希望反映背景色をクリア
  clearBackgrounds(pdfSheet);
  // 行の高さを1.5倍に設定
  const originalHeight = pdfSheet.getRowHeight(SHIFT_ROW_START);
  const newHeight = Math.floor(originalHeight * 1.5);
  pdfSheet.setRowHeights(
    SHIFT_ROW_MEMBERS,
    SHIFT_ROW_END - SHIFT_ROW_MEMBERS + 1,
    newHeight
  );

  // デフォルトの空白シート（初期生成されるもの）を削除
  const newSheets = newSpreadsheet.getSheets();
  newSpreadsheet.deleteSheet(newSheets[0]); // 最初の空白シート

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
  const pdfBlob = response.getBlob().setName(`シフト作成日時_${timestamp}`);
  pdfFolder.createFile(pdfBlob);

  Logger.log(`${dateStr}: 完了`);

  ui.alert(`✅ ${dateStr}のシフトを共有ファイルに反映しました！`);
}

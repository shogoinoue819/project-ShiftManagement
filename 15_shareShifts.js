/**
 * 管理シート名を指定して共有処理を実行（前回分/現在分 どちらでも可）
 * @param {string} manageSheetName - 例: SHEET_NAMES.SHIFT_MANAGEMENT_PREVIOUS or SHEET_NAMES.SHIFT_MANAGEMENT
 * @return {number} sharedCount - 共有した日数
 */
function shareShiftsFromManageSheet(manageSheetName) {
  // 共有に必要な共通オブジェクト取得
  const [ss, , , allSheets, ui] = getCommonSheets(); // manageSheetは後で取り直す
  const manageSheet = ss.getSheetByName(manageSheetName);
  if (!manageSheet) {
    ui.alert(
      `管理シート「${manageSheetName}」が見つかりません。処理を中断します。`
    );
    return 0;
  }

  // 共有先ファイル・保存フォルダ
  const shareFile = SpreadsheetApp.openById(SHARE_FILE_ID);
  const pdfFolder = DriveApp.getFolderById(SHIFT_PDF_FOLDER_ID);
  const ssFolder = DriveApp.getFolderById(SHIFT_SS_FOLDER_ID);

  // 日程リスト取得
  const last = getLastRowInCol(
    manageSheet,
    SHIFT_MANAGEMENT_SHEET.DATE_LIST.COL
  );
  if (last < SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW) return 0;

  const data = manageSheet
    .getRange(
      SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW,
      SHIFT_MANAGEMENT_SHEET.DATE_LIST.COL,
      last - SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW + 1,
      SHIFT_MANAGEMENT_SHEET.DATE_LIST.SHARE_COL -
        SHIFT_MANAGEMENT_SHEET.DATE_LIST.COL +
        1
    )
    .getValues();

  // 共有用の空スプレッドシートを作成（このバッチ分まとめPDF化）
  const now = new Date();
  const timestamp = Utilities.formatDate(
    now,
    "Asia/Tokyo",
    "yyyy-MM-dd_HH-mm-ss"
  );
  const workSS = SpreadsheetApp.create(`シフト作成日時_${timestamp}`);
  workSS.setSpreadsheetLocale("ja_JP");
  const workId = workSS.getId();
  const workFile = DriveApp.getFileById(workId);
  ssFolder.addFile(workFile);
  DriveApp.getRootFolder().removeFile(workFile);

  let sharedCount = 0;

  data.forEach((row, i) => {
    const [date, isComplete, isShare] = row;

    // 完成済み & 未共有のみ対象
    if (isComplete === true && isShare === SHARE_FALSE) {
      const dateStr = formatDateToString(date);
      const sheetName = dateStr;

      try {
        // 元のシート
        const dailySheet = ss.getSheetByName(sheetName);
        if (!dailySheet)
          throw new Error(`シート「${sheetName}」が見つかりません`);

        // 共有スプレッドシートへコピー（既存同名があれば置換）
        const exists = shareFile.getSheetByName(sheetName);
        if (exists) shareFile.deleteSheet(exists);

        const copiedToShare = dailySheet.copyTo(shareFile).setName(sheetName);
        // ビュー調整
        copiedToShare.hideRows(
          SHIFT_ROW_START_TIME,
          SHIFT_ROW_NOTE - SHIFT_ROW_START_TIME + 1
        );
        clearBackgrounds(copiedToShare);

        // PDF用ワークSSにもコピー
        const copiedToWork = dailySheet.copyTo(workSS).setName(sheetName);
        copiedToWork.hideRows(
          SHIFT_ROW_START_TIME,
          SHIFT_ROW_NOTE - SHIFT_ROW_START_TIME + 1
        );
        clearBackgrounds(copiedToWork);
        // 行高1.5倍
        const h = copiedToWork.getRowHeight(SHIFT_ROW_START);
        copiedToWork.setRowHeights(
          SHIFT_ROW_MEMBERS,
          SHIFT_ROW_END - SHIFT_ROW_MEMBERS + 1,
          Math.floor(h * 1.5)
        );

        // 「共有済」に更新
        manageSheet
          .getRange(
            i + SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW,
            SHIFT_MANAGEMENT_SHEET.DATE_LIST.SHARE_COL
          )
          .setValue(SHARE_TRUE);
        sharedCount++;
        Logger.log(`${manageSheetName} / ${dateStr}: 共有完了`);
      } catch (e) {
        Logger.log(`エラー (${manageSheetName} / ${sheetName}): ${e.message}`);
      }
    }
  });

  if (sharedCount > 0) {
    // 共有先のシートを M/d で抽出→日付ソート→並べ替え
    const sheets = shareFile.getSheets();
    const sorted = sheets
      .filter((s) => /^\d{1,2}\/\d{1,2}$/.test(s.getName()))
      .map((s) => ({ sheet: s, date: formatStringToDate(s.getName()) }))
      .filter((x) => x.date !== null)
      .sort((a, b) => a.date - b.date);

    sorted.forEach((x, idx) => {
      shareFile.setActiveSheet(x.sheet);
      shareFile.moveActiveSheet(idx + 1); // 0番はマスター等の可能性を考慮
    });

    // ワークSSの初期空白シートを削除（余っていれば）
    const newSheets = workSS.getSheets();
    if (newSheets.length > sharedCount) {
      workSS.deleteSheet(newSheets[0]);
    }

    // PDF化（1ファイルにまとめる）
    const url =
      "https://docs.google.com/spreadsheets/d/" +
      workId +
      "/export?format=pdf" +
      "&portrait=false" +
      "&size=A4" +
      "&fitw=true" +
      "&scale=4" +
      "&sheetnames=false" +
      "&printtitle=false" +
      "&pagenumbers=false" +
      "&gridlines=false" +
      "&fzr=false";

    const token = ScriptApp.getOAuthToken();
    const res = UrlFetchApp.fetch(url, {
      headers: { Authorization: "Bearer " + token },
    });

    const pdfBlob = res.getBlob().setName(`シフト作成日時_${timestamp}`);
    pdfFolder.createFile(pdfBlob);

    ui.alert(
      `✅ 「${manageSheetName}」から ${sharedCount} 日分を共有しました。`
    );
  } else {
    // 共有対象なし → ワークSSは削除
    DriveApp.getFileById(workId).setTrashed(true);
    Logger.log(`共有対象なし（${manageSheetName}）`);
  }

  return sharedCount;
}

/**
 * ① 前回分（SHIFT_MANAGEMENT_PREVIOUS）→ ② 現在分（SHIFT_MANAGEMENT）の順で共有
 * 両方に対象があるケースは稀だが想定して順次処理する
 */
function shareShiftsAll() {
  const totalPre = shareShiftsFromManageSheet(
    SHEET_NAMES.SHIFT_MANAGEMENT_PREVIOUS
  );
  const totalCurr = shareShiftsFromManageSheet(SHEET_NAMES.SHIFT_MANAGEMENT);
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    `完了：前回分 ${totalPre} 日、現在分 ${totalCurr} 日を共有しました。`
  );
}

/**
 * 管理シート（前回分→現在分の順）で該当日付の行番号を探す
 * 見つかれば絶対行番号を返し、見つからなければ null を返す
 */
function findDateRow(manageSheet, dateStr) {
  if (!manageSheet) return null;
  const last = getLastRowInCol(
    manageSheet,
    SHIFT_MANAGEMENT_SHEET.DATE_LIST.COL
  );
  if (last < SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW) return null;

  const values = manageSheet
    .getRange(
      SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW,
      SHIFT_MANAGEMENT_SHEET.DATE_LIST.COL,
      last - SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW + 1,
      1
    )
    .getValues()
    .map((v) => v[0]);

  // セルが Date の可能性にも対応
  for (let i = 0; i < values.length; i++) {
    const v = values[i];
    const asStr = v instanceof Date ? formatDateToString(v) : String(v);
    if (asStr === dateStr) {
      return SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW + i; // 絶対行番号で返す
    }
  }
  return null;
}

/**
 * 今開いている日次シート1枚だけを共有用ファイルへ反映し、PDFも作成。
 * 共有済みフラグは「前回分→現在分」の順で該当日付を探して更新する。
 */
function shareOnlyOneShift() {
  const ui = SpreadsheetApp.getUi();

  // 今開いているシート（日付名想定）
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dailySheet = ss.getActiveSheet();
  const dateStr = dailySheet.getName();

  // 共有用ファイル
  const shareFile = SpreadsheetApp.openById(SHARE_FILE_ID);

  // ====== スプレッドシート共有（1枚） ======
  const existing = shareFile.getSheetByName(dateStr);
  if (existing) shareFile.deleteSheet(existing);

  const copied = dailySheet.copyTo(shareFile).setName(dateStr);
  copied.hideRows(
    SHIFT_ROW_START_TIME,
    SHIFT_ROW_NOTE - SHIFT_ROW_START_TIME + 1
  );
  clearBackgrounds(copied);

  // 並べ替え（M/dのシートを日付昇順）
  const sheets = shareFile.getSheets();
  const sorted = sheets
    .filter((s) => /^\d{1,2}\/\d{1,2}$/.test(s.getName()))
    .map((s) => ({ sheet: s, date: formatStringToDate(s.getName()) }))
    .filter((x) => x.date !== null)
    .sort((a, b) => a.date - b.date);

  sorted.forEach((x, idx) => {
    shareFile.setActiveSheet(x.sheet);
    shareFile.moveActiveSheet(idx + 1);
  });

  // ====== PDF 共有（1枚） ======
  const pdfFolder = DriveApp.getFolderById(SHIFT_PDF_FOLDER_ID);
  const ssFolder = DriveApp.getFolderById(SHIFT_SS_FOLDER_ID);

  const now = new Date();
  const timestamp = Utilities.formatDate(
    now,
    "Asia/Tokyo",
    "yyyy-MM-dd_HH-mm-ss"
  );

  const workSS = SpreadsheetApp.create(`シフト作成日時_${timestamp}`);
  workSS.setSpreadsheetLocale("ja_JP");
  const workId = workSS.getId();
  const workFile = DriveApp.getFileById(workId);
  ssFolder.addFile(workFile);
  DriveApp.getRootFolder().removeFile(workFile);

  // デフォルト空白シートを先に削除してからコピーしてもOK（好み）
  // workSS.deleteSheet(workSS.getSheets()[0]);

  const pdfSheet = dailySheet.copyTo(workSS).setName(dateStr);
  pdfSheet.hideRows(
    SHIFT_ROW_START_TIME,
    SHIFT_ROW_NOTE - SHIFT_ROW_START_TIME + 1
  );
  clearBackgrounds(pdfSheet);
  const h = pdfSheet.getRowHeight(SHIFT_ROW_START);
  pdfSheet.setRowHeights(
    SHIFT_ROW_MEMBERS,
    SHIFT_ROW_END - SHIFT_ROW_MEMBERS + 1,
    Math.floor(h * 1.5)
  );

  // 初期空白シートが残っていれば削除
  const ws = workSS.getSheets();
  if (ws.length > 1) {
    // 先頭がデフォルト空白の想定
    workSS.deleteSheet(ws[0]);
  }

  const url =
    "https://docs.google.com/spreadsheets/d/" +
    workId +
    "/export?format=pdf" +
    "&portrait=false" +
    "&size=A4" +
    "&fitw=true" +
    "&scale=4" +
    "&sheetnames=false" +
    "&printtitle=false" +
    "&pagenumbers=false" +
    "&gridlines=false" +
    "&fzr=false";

  const token = ScriptApp.getOAuthToken();
  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + token },
  });
  const pdfBlob = res.getBlob().setName(`シフト作成日時_${timestamp}`);
  pdfFolder.createFile(pdfBlob);

  // ====== 共有済みフラグ更新（前回分→現在分） ======
  const manageSheetPre = ss.getSheetByName(
    SHEET_NAMES.SHIFT_MANAGEMENT_PREVIOUS
  );
  const manageSheet = ss.getSheetByName(SHEET_NAMES.SHIFT_MANAGEMENT);

  let updated = false;
  const preRow = findDateRow(manageSheetPre, dateStr);
  if (preRow) {
    manageSheetPre
      .getRange(preRow, SHIFT_MANAGEMENT_SHEET.DATE_LIST.SHARE_COL)
      .setValue(SHARE_TRUE);
    updated = true;
  } else {
    const currRow = findDateRow(manageSheet, dateStr);
    if (currRow) {
      manageSheet
        .getRange(currRow, SHIFT_MANAGEMENT_SHEET.DATE_LIST.SHARE_COL)
        .setValue(SHARE_TRUE);
      updated = true;
    }
  }

  if (!updated) {
    Logger.log(
      `⚠️ 管理シート上に ${dateStr} が見つからず、共有済みフラグを更新できませんでした。`
    );
  }

  Logger.log(`${dateStr}: 完了`);
  ui.alert(`✅ ${dateStr} のシフトを共有＆PDF化しました！`);
}

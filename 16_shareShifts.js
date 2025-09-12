// シフト共有機能

// ====== メイン関数 ======

/**
 * ① 前回分（SHIFT_MANAGEMENT_PREVIOUS）→ ② 現在分（SHIFT_MANAGEMENT）の順で共有
 * 両方に対象があるケースは稀だが想定して順次処理する
 */
function shareShiftsAll() {
  const resultPre = shareShiftsFromManageSheet(
    SHEET_NAMES.SHIFT_MANAGEMENT_PREVIOUS
  );
  const resultCurr = shareShiftsFromManageSheet(SHEET_NAMES.SHIFT_MANAGEMENT);

  const ui = SpreadsheetApp.getUi();

  // 詳細レポートを作成
  let report = "=== シフト共有完了レポート ===\n\n";

  // 前回分の結果
  report += `【前回分】\n`;
  if (resultPre.successCount > 0) {
    report += `✅ 共有成功: ${resultPre.successCount}日\n`;
    if (resultPre.successDates.length > 0) {
      report += `   日程: ${resultPre.successDates.join(", ")}\n`;
    }
  }
  if (resultPre.failedCount > 0) {
    report += `❌ 共有失敗: ${resultPre.failedCount}日\n`;
    if (resultPre.failedDates.length > 0) {
      report += `   日程: ${resultPre.failedDates.join(", ")}\n`;
    }
  }
  if (resultPre.successCount === 0 && resultPre.failedCount === 0) {
    report += `ℹ️ 共有対象なし\n`;
  }

  report += `\n【現在分】\n`;
  if (resultCurr.successCount > 0) {
    report += `✅ 共有成功: ${resultCurr.successCount}日\n`;
    if (resultCurr.successDates.length > 0) {
      report += `   日程: ${resultCurr.successDates.join(", ")}\n`;
    }
  }
  if (resultCurr.failedCount > 0) {
    report += `❌ 共有失敗: ${resultCurr.failedCount}日\n`;
    if (resultCurr.failedDates.length > 0) {
      report += `   日程: ${resultCurr.failedDates.join(", ")}\n`;
    }
  }
  if (resultCurr.successCount === 0 && resultCurr.failedCount === 0) {
    report += `ℹ️ 共有対象なし\n`;
  }

  const totalSuccess = resultPre.successCount + resultCurr.successCount;
  const totalFailed = resultPre.failedCount + resultCurr.failedCount;

  report += `\n【総合結果】\n`;
  report += `✅ 成功: ${totalSuccess}日\n`;
  report += `❌ 失敗: ${totalFailed}日\n`;

  ui.alert(report);
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
  configureSheetForSharing(copied, false); // シート共有用：行高調整なし

  // 並べ替え（M/dのシートを日付昇順）
  sortSheetsByDate(shareFile);

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
  configureSheetForSharing(pdfSheet, true);

  // 初期空白シートが残っていれば削除
  const ws = workSS.getSheets();
  if (ws.length > 1) {
    // 先頭がデフォルト空白の想定
    workSS.deleteSheet(ws[0]);
  }

  createPdfFromSpreadsheet(
    workId,
    `シフト作成日時_${timestamp}`,
    SHIFT_PDF_FOLDER_ID
  );

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
      .setValue(STATUS_STRINGS.SHARE.TRUE);
    updated = true;
  } else {
    const currRow = findDateRow(manageSheet, dateStr);
    if (currRow) {
      manageSheet
        .getRange(currRow, SHIFT_MANAGEMENT_SHEET.DATE_LIST.SHARE_COL)
        .setValue(STATUS_STRINGS.SHARE.TRUE);
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

// ====== ヘルパー関数 ======

/**
 * 管理シート名を指定して共有処理を実行（前回分/現在分 どちらでも可）
 * @param {string} manageSheetName - 例: SHEET_NAMES.SHIFT_MANAGEMENT_PREVIOUS or SHEET_NAMES.SHIFT_MANAGEMENT
 * @return {number} sharedCount - 共有した日数
 */
function shareShiftsFromManageSheet(manageSheetName) {
  // 共有に必要な共通オブジェクト取得
  const ss = getSpreadsheet();
  const allSheets = getAllSheets();
  const ui = getUI(); // manageSheetは後で取り直す
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
  const last = getLastRowInColumn(
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
  let successDates = [];
  let failedDates = [];
  let failedReasons = [];

  data.forEach((row, i) => {
    const [date, , isComplete, isShare] = row;

    // 完成済み & 未共有のみ対象
    if (isComplete === true && isShare === STATUS_STRINGS.SHARE.FALSE) {
      const dateStr = formatDateToString(date);
      const sheetName = dateStr;

      try {
        // 元のシート
        const dailySheet = ss.getSheetByName(sheetName);
        if (!dailySheet) {
          throw new Error(`シート「${sheetName}」が見つかりません`);
        }

        // 共有スプレッドシートへコピー（既存同名があれば置換）
        const exists = shareFile.getSheetByName(sheetName);
        if (exists) shareFile.deleteSheet(exists);

        const copiedToShare = dailySheet.copyTo(shareFile).setName(sheetName);
        // ビュー調整（シート共有用：行高調整なし）
        configureSheetForSharing(copiedToShare, false);

        // PDF用ワークSSにもコピー
        const copiedToWork = dailySheet.copyTo(workSS).setName(sheetName);
        configureSheetForSharing(copiedToWork, true);

        // 「共有済」に更新
        manageSheet
          .getRange(
            i + SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW,
            SHIFT_MANAGEMENT_SHEET.DATE_LIST.SHARE_COL
          )
          .setValue(STATUS_STRINGS.SHARE.TRUE);
        sharedCount++;
        successDates.push(dateStr);
        Logger.log(`✅ ${dateStr}: 共有完了`);
      } catch (e) {
        Logger.log(`❌ ${dateStr}: 共有失敗 - ${e.message}`);
        failedDates.push(dateStr);
        failedReasons.push(e.message);
      }
    }
  });

  if (sharedCount > 0) {
    // 共有先のシートを日付順に並び替え
    sortSheetsByDate(shareFile);

    // ワークSSの初期空白シートを削除（余っていれば）
    const newSheets = workSS.getSheets();
    if (newSheets.length > sharedCount) {
      workSS.deleteSheet(newSheets[0]);
    }

    // PDF化（1ファイルにまとめる）
    createPdfFromSpreadsheet(
      workId,
      `シフト作成日時_${timestamp}`,
      SHIFT_PDF_FOLDER_ID
    );

    Logger.log(`✅ ${sharedCount} 日分を共有しました`);
  } else {
    // 共有対象なし → ワークSSは削除
    DriveApp.getFileById(workId).setTrashed(true);
    Logger.log(`共有対象なし（${manageSheetName}）`);
  }

  return {
    successCount: sharedCount,
    failedCount: failedDates.length,
    successDates: successDates,
    failedDates: failedDates,
    failedReasons: failedReasons,
  };
}

/**
 * PDFエクスポート用のURLを構築
 * @param {string} spreadsheetId - スプレッドシートID
 * @return {string} PDFエクスポートURL
 */
function buildPdfExportUrl(spreadsheetId) {
  const baseUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=pdf`;
  const params = [
    `portrait=${PDF_EXPORT_CONFIG.PORTRAIT}`,
    `size=${PDF_EXPORT_CONFIG.SIZE}`,
    `fitw=${PDF_EXPORT_CONFIG.FIT_WIDTH}`,
    `scale=${PDF_EXPORT_CONFIG.SCALE}`,
    `sheetnames=${PDF_EXPORT_CONFIG.SHOW_SHEET_NAMES}`,
    `printtitle=${PDF_EXPORT_CONFIG.SHOW_TITLE}`,
    `pagenumbers=${PDF_EXPORT_CONFIG.SHOW_PAGE_NUMBERS}`,
    `gridlines=${PDF_EXPORT_CONFIG.SHOW_GRIDLINES}`,
    `fzr=${PDF_EXPORT_CONFIG.FIX_ROW_HEIGHT}`,
  ];
  return `${baseUrl}&${params.join("&")}`;
}

/**
 * シートを共有用に設定（行の非表示、背景クリア、行高調整）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {boolean} adjustRowHeight - 行高調整を行うかどうか
 */
function configureSheetForSharing(sheet, adjustRowHeight = false) {
  // 行の非表示
  sheet.hideRows(
    SHIFT_TEMPLATE_SHEET.ROWS.START_TIME,
    SHIFT_TEMPLATE_SHEET.ROWS.NOTE - SHIFT_TEMPLATE_SHEET.ROWS.START_TIME + 1
  );

  // 背景クリア
  clearBackgrounds(sheet);

  // 行高調整（PDF用の場合のみ）
  if (adjustRowHeight) {
    const baseHeight = sheet.getRowHeight(SHIFT_TEMPLATE_SHEET.ROWS.DATA_START);
    const adjustedHeight = Math.floor(baseHeight * ROW_HEIGHT_MULTIPLIER);
    sheet.setRowHeights(
      SHIFT_TEMPLATE_SHEET.ROWS.MEMBERS,
      SHIFT_TEMPLATE_SHEET.ROWS.DATA_END -
        SHIFT_TEMPLATE_SHEET.ROWS.MEMBERS +
        1,
      adjustedHeight
    );
  }
}

/**
 * 日付形式のシートを日付順に並び替え
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - 対象スプレッドシート
 */
function sortSheetsByDate(spreadsheet) {
  const sheets = spreadsheet.getSheets();
  const sorted = sheets
    .filter((s) => /^\d{1,2}\/\d{1,2}$/.test(s.getName()))
    .map((s) => ({ sheet: s, date: formatStringToDate(s.getName()) }))
    .filter((x) => x.date !== null)
    .sort((a, b) => a.date - b.date);

  sorted.forEach((x, idx) => {
    spreadsheet.setActiveSheet(x.sheet);
    spreadsheet.moveActiveSheet(idx + 1); // 0番はマスター等の可能性を考慮
  });
}

/**
 * スプレッドシートからPDFを作成して指定フォルダに保存
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} pdfFileName - PDFファイル名
 * @param {string} folderId - 保存先フォルダID
 * @return {GoogleAppsScript.Drive.File} 作成されたPDFファイル
 */
function createPdfFromSpreadsheet(spreadsheetId, pdfFileName, folderId) {
  const url = buildPdfExportUrl(spreadsheetId);
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + token },
  });

  const pdfBlob = response.getBlob().setName(pdfFileName);
  const folder = DriveApp.getFolderById(folderId);
  return folder.createFile(pdfBlob);
}

/**
 * 管理シート（前回分→現在分の順）で該当日付の行番号を探す
 * 見つかれば絶対行番号を返し、見つからなければ null を返す
 */
function findDateRow(manageSheet, dateStr) {
  if (!manageSheet) return null;
  const last = getLastRowInColumn(
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

// clearBackgrounds関数は03_utils.jsで定義済み
// 重複関数は03_utils.jsのものを使用:
// - getSpreadsheet()
// - getAllSheets()
// - getUI()
// - getLastRowInColumn()
// - formatStringToDate()
// - formatDateToString()

// 定数は01_consts.jsで定義済み:
// - PDF_EXPORT_CONFIG
// - ROW_HEIGHT_MULTIPLIER

// デバッグ用今後の勤務希望テンプレート反映

// ===== 設定定数 =====
const INFO_SHEET_PROCESSING_CONFIG = {
  LIMIT_COUNT: 30, // 処理対象人数の制限
  PROCESS_FIRST_HALF: true, // true: 前半処理, false: 後半処理
  // 前半処理: 1-30人目まで処理
  // 後半処理: 31人目以降を処理
};

function reReflectTemplateInfoSheet() {
  // テンプレートシートの取得と検証
  const templateSheet = getInfoSheetTemplateSheet();
  if (!templateSheet) {
    throw new Error("❌ テンプレートシートの取得に失敗しました");
  }

  // メンバー管理の初期化
  const memberManager = initializeMemberManager();
  if (!memberManager) {
    throw new Error("❌ メンバーデータの初期化に失敗しました");
  }

  let count = 0;
  let index = 0;

  for (const [id, { name, url }] of Object.entries(memberManager.memberMap)) {
    // 前半・後半の処理分岐
    if (INFO_SHEET_PROCESSING_CONFIG.PROCESS_FIRST_HALF) {
      // 前半処理: 制限人数まで処理
      if (index >= INFO_SHEET_PROCESSING_CONFIG.LIMIT_COUNT) break;
    } else {
      // 後半処理: 制限人数まではスキップ
      if (index < INFO_SHEET_PROCESSING_CONFIG.LIMIT_COUNT) {
        index++;
        continue;
      }
    }

    try {
      const success = processMemberSheet(
        name,
        url,
        templateSheet,
        SHEET_NAMES.SHIFT_FORM_INFO,
        2
      );
      if (success) {
        count++;
        Logger.log(`✅ ${name} さんに「今後の勤務希望」シートを再反映しました`);
      }
    } catch (e) {
      Logger.log(`❌ ${name} さんの処理中にエラー: ${e.message}`);
    }
    index++;
  }

  Logger.log(
    `✅ 完了：${count} 名に「今後の勤務希望」シートを上書き反映しました`
  );

  // 処理設定の表示
  const processType = INFO_SHEET_PROCESSING_CONFIG.PROCESS_FIRST_HALF
    ? "前半"
    : "後半";
  Logger.log(
    `📋 処理設定: ${processType}処理 (制限人数: ${INFO_SHEET_PROCESSING_CONFIG.LIMIT_COUNT}人)`
  );
}

// ===== ヘルパー関数 =====
function getInfoSheetTemplateSheet() {
  try {
    const templateSS = SpreadsheetApp.openById(TEMPLATE_FILE_ID);

    // デバッグ: テンプレートファイル内の全シート名を確認
    const allSheets = templateSS.getSheets();
    Logger.log("🔍 テンプレートファイル内の全シート名:");
    allSheets.forEach((sheet, index) => {
      Logger.log(`  ${index + 1}: "${sheet.getName()}"`);
    });

    const templateSheet = templateSS.getSheetByName(
      SHEET_NAMES.SHIFT_FORM_INFO
    );
    if (!templateSheet) {
      Logger.log(
        `⚠️ テンプレートにシート '${SHEET_NAMES.SHIFT_FORM_INFO}' が見つかりません`
      );
      return null;
    }

    // デバッグ: 実際に取得されたシート名を確認
    Logger.log(
      `🔍 テンプレートから取得したシート名: "${templateSheet.getName()}"`
    );
    Logger.log(`🔍 期待されるシート名: "${SHEET_NAMES.SHIFT_FORM_INFO}"`);

    return templateSheet;
  } catch (e) {
    Logger.log(`❌ テンプレートシート取得エラー: ${e.message}`);
    return null;
  }
}

// initializeMemberManager()関数は03_utils.jsで定義済み

function processMemberSheet(
  memberName,
  url,
  templateSheet,
  sheetName,
  movePosition
) {
  try {
    // URLからファイルIDを抽出
    const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!match || !match[1]) {
      Logger.log(`❌ ${memberName} さんのURLが不正です: ${url}`);
      return false;
    }

    const fileId = match[1];
    const memberSS = SpreadsheetApp.openById(fileId);

    // 既存シートを削除
    const existingSheet = memberSS.getSheetByName(sheetName);
    if (existingSheet) {
      memberSS.deleteSheet(existingSheet);
    }

    // コピーしてリネーム
    const copiedSheet = templateSheet.copyTo(memberSS);
    copiedSheet.setName(sheetName);
    memberSS.setActiveSheet(copiedSheet);
    memberSS.moveActiveSheet(movePosition);

    // シート整理処理
    organizeMemberSheets(memberSS, memberName);

    Logger.log(`✅ ${memberName} さんのシート処理完了`);
    return true;
  } catch (e) {
    Logger.log(`❌ ${memberName} さんのシート処理エラー: ${e.message}`);
    return false;
  }
}

// organizeMemberSheets()関数は03_utils.jsで定義済み

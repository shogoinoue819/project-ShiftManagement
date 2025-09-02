// デバッグ用シフト希望表テンプレート反映

// ===== 設定定数 =====
const SHIFT_FORM_PROCESSING_CONFIG = {
  LIMIT_COUNT: 30, // 処理対象人数の制限
  PROCESS_FIRST_HALF: true, // true: 前半処理, false: 後半処理
  // 前半処理: 1-30人目まで処理
  // 後半処理: 31人目以降を処理
};

function reReflectTemplateSheet() {
  // テンプレートシートの取得と検証
  const templateSheet = getShiftFormTemplateSheet();
  if (!templateSheet) {
    throw new Error("❌ テンプレートシートの取得に失敗しました");
  }

  // メンバー管理の初期化
  const memberManager = initializeMemberManager();
  if (!memberManager) {
    throw new Error("❌ メンバーデータの初期化に失敗しました");
  }

  let count = 0;

  // 提出ステータス列を取得
  const submitValues = getSubmitStatusValues();

  let index = 0;
  for (const [id, { name, url }] of Object.entries(memberManager.memberMap)) {
    // 前半・後半の処理分岐
    if (SHIFT_FORM_PROCESSING_CONFIG.PROCESS_FIRST_HALF) {
      // 前半処理: 制限人数まで処理
      if (index >= SHIFT_FORM_PROCESSING_CONFIG.LIMIT_COUNT) break;
    } else {
      // 後半処理: 制限人数まではスキップ
      if (index < SHIFT_FORM_PROCESSING_CONFIG.LIMIT_COUNT) {
        index++;
        continue;
      }
    }

    // 未提出以外はスキップ
    const submit = submitValues[index];
    if (submit !== STATUS_STRINGS.SUBMIT.FALSE) {
      index++;
      continue;
    }

    try {
      const success = processShiftFormMemberSheet(
        name,
        url,
        templateSheet,
        SHEET_NAMES.SHIFT_FORM,
        1
      );
      if (success) {
        count++;
        Logger.log(`✅ ${name} さんに「シフト希望表」シートを再反映しました`);
      }
    } catch (e) {
      Logger.log(`❌ ${name} さんの処理中にエラー: ${e.message}`);
    }
    index++;
  }

  Logger.log(
    `✅ 完了：${count} 名に「シフト希望表」シートを上書き反映しました`
  );

  // 処理設定の表示
  const processType = SHIFT_FORM_PROCESSING_CONFIG.PROCESS_FIRST_HALF
    ? "前半"
    : "後半";
  Logger.log(
    `📋 処理設定: ${processType}処理 (制限人数: ${SHIFT_FORM_PROCESSING_CONFIG.LIMIT_COUNT}人)`
  );
}

// ===== ヘルパー関数 =====
function getShiftFormTemplateSheet() {
  try {
    const templateSS = SpreadsheetApp.openById(TEMPLATE_FILE_ID);

    // デバッグ: テンプレートファイル内の全シート名を確認
    const allSheets = templateSS.getSheets();
    Logger.log("🔍 テンプレートファイル内の全シート名:");
    allSheets.forEach((sheet, index) => {
      Logger.log(`  ${index + 1}: "${sheet.getName()}"`);
    });

    const templateSheet = templateSS.getSheetByName(SHEET_NAMES.SHIFT_FORM);
    if (!templateSheet) {
      Logger.log(
        `⚠️ テンプレートにシート '${SHEET_NAMES.SHIFT_FORM}' が見つかりません`
      );
      return null;
    }

    // デバッグ: 実際に取得されたシート名を確認
    Logger.log(
      `🔍 テンプレートから取得したシート名: "${templateSheet.getName()}"`
    );
    Logger.log(`🔍 期待されるシート名: "${SHEET_NAMES.SHIFT_FORM}"`);

    return templateSheet;
  } catch (e) {
    Logger.log(`❌ テンプレートシート取得エラー: ${e.message}`);
    return null;
  }
}

// initializeMemberManager()関数は03_utils.jsで定義済み

function getSubmitStatusValues() {
  const manageSheet = getManageSheet();
  const lastRow = getLastRowInColumn(
    manageSheet,
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
  );

  return manageSheet
    .getRange(
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW,
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.SUBMIT_COL,
      lastRow - SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW + 1,
      1
    )
    .getValues()
    .flat();
}

function processShiftFormMemberSheet(
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

    // シフト希望表の場合のみ初期化処理
    if (sheetName === SHEET_NAMES.SHIFT_FORM) {
      try {
        // 名前を設定
        copiedSheet
          .getRange(
            SHIFT_FORM_TEMPLATE.HEADER.ROW,
            SHIFT_FORM_TEMPLATE.HEADER.NAME_COL
          )
          .setValue(memberName);

        Logger.log(`✅ ${memberName} さんの初期化処理完了`);
      } catch (initError) {
        Logger.log(
          `⚠️ ${memberName} さんの初期化処理でエラー: ${initError.message}`
        );
        // 初期化エラーでも処理は継続
      }
    }

    // シート整理処理
    organizeMemberSheets(memberSS, memberName);

    return true;
  } catch (e) {
    Logger.log(`❌ ${memberName} さんのシート処理エラー: ${e.message}`);
    return false;
  }
}

// organizeMemberSheets()関数は03_utils.jsで定義済み

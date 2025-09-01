// 個別ファイルのシフト希望表をアップデート
function updateForms() {
  Logger.log("🔄 シフト希望表アップデート処理を開始");

  // SSをまとめて取得
  const ss = getSpreadsheet();
  const manageSheet = getManageSheet();
  const ui = getUI();

  // 確認ダイアログを表示
  if (!confirmUpdateOperation(ui)) {
    Logger.log("❌ ユーザーにより操作がキャンセルされました");
    return;
  }

  // メンバーデータの初期化と検証
  const memberMap = initializeAndValidateMembers(ui);
  if (!memberMap) {
    Logger.log("❌ メンバーデータの初期化に失敗しました");
    return;
  }

  Logger.log(`📋 メンバーデータ取得成功: ${Object.keys(memberMap).length}件`);

  // 管理シートのリセット
  resetManagementSheet(manageSheet, memberMap);

  // テンプレートデータの取得
  const templateData = getTemplateData();
  Logger.log("📄 テンプレートデータ取得成功");

  // 各メンバーの個別ファイルをアップデート
  updateAllMemberForms(memberMap, templateData);

  Logger.log("🎉 シフト希望表アップデート処理が完了しました");
  ui.alert("✅ シフト希望表の個別ファイルをすべて更新しました！");
}

// 更新操作の確認
function confirmUpdateOperation(ui) {
  const confirm = ui.alert(
    "⚠️確認",
    "この操作で、全ての個別ファイルの中身が更新されます。\n個別ファイルの現在の入力内容は前回分として保存されます。\n\n本当に実行してよろしいですか？",
    ui.ButtonSet.OK_CANCEL
  );

  if (confirm !== ui.Button.OK) {
    ui.alert("❌ 操作はキャンセルされました");
    return false;
  }
  return true;
}

// メンバーデータの初期化と検証
function initializeAndValidateMembers(ui) {
  const manageSheet = getManageSheet();
  const memberManager = getMemberManager(manageSheet);

  // 初期化を確実に行う
  if (!memberManager.ensureInitialized()) {
    ui.alert("❌ メンバーデータの初期化に失敗しました");
    return null;
  }

  const memberMap = memberManager.memberMap;

  // メンバーマップの妥当性チェック
  if (!memberMap || Object.keys(memberMap).length === 0) {
    ui.alert("❌ メンバーデータが取得できませんでした");
    return null;
  }

  return memberMap;
}

// 管理シートのリセット
function resetManagementSheet(manageSheet, memberMap) {
  const memberCount = Object.keys(memberMap).length;
  const startRow = SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW;

  // バッチ処理でチェック列と反映ステータス列を同時にリセット
  const ranges = [
    {
      range: manageSheet.getRange(
        startRow,
        SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.CHECK_COL,
        memberCount,
        1
      ),
      value: false,
    },
    {
      range: manageSheet.getRange(
        startRow,
        SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.REFLECT_COL,
        memberCount,
        1
      ),
      value: STATUS_STRINGS.REFLECT.FALSE,
    },
  ];

  // 一括で値を設定
  ranges.forEach(({ range, value }) => {
    range.setValue(value);
  });

  Logger.log(`📊 管理シートをリセット: ${memberCount}件のメンバー`);
}

// テンプレートデータの取得
function getTemplateData() {
  const templateFile = SpreadsheetApp.openById(TEMPLATE_FILE_ID);
  const formTemplateSheet = templateFile.getSheetByName(SHEET_NAMES.SHIFT_FORM);

  const templateRange = formTemplateSheet.getDataRange();
  const numRows = templateRange.getNumRows();
  const numCols = templateRange.getNumColumns();
  const values = templateRange.getValues();

  return {
    file: templateFile,
    sheet: formTemplateSheet,
    numRows,
    numCols,
    values,
  };
}

// 全メンバーの個別ファイルをアップデート
function updateAllMemberForms(memberMap, templateData) {
  const totalMembers = Object.keys(memberMap).length;
  let successCount = 0;
  let errorCount = 0;
  const errors = [];

  Logger.log(`🚀 個別ファイルの更新を開始: ${totalMembers}件のメンバー`);

  for (const [id, { name, url }] of Object.entries(memberMap)) {
    try {
      updateIndividualForm(name, url, templateData);
      successCount++;
      Logger.log(`✅ アップデート完了: ${name}`);
    } catch (e) {
      errorCount++;
      const errorInfo = { name, error: e.message };
      errors.push(errorInfo);
      Logger.log(`❌ エラー: ${name} - ${e.message}`);
    }
  }

  // 結果サマリーをログ出力
  Logger.log(
    `📊 更新完了サマリー: 成功 ${successCount}件, エラー ${errorCount}件`
  );

  if (errors.length > 0) {
    Logger.log("⚠️ エラーが発生したメンバー:");
    errors.forEach(({ name, error }) => {
      Logger.log(`  - ${name}: ${error}`);
    });
  }
}

// 個別ファイルのアップデート処理
function updateIndividualForm(memberName, memberUrl, templateData) {
  const fileId = extractFileIdFromUrl(memberUrl);
  if (!fileId) {
    throw new Error(`ファイルIDの抽出に失敗: ${memberUrl}`);
  }

  let memberSS;
  try {
    memberSS = SpreadsheetApp.openById(fileId);
  } catch (e) {
    throw new Error(`スプレッドシートの開封に失敗: ${e.message}`);
  }

  try {
    // 各処理ステップを実行
    const { currentFormSheet, previousSheet } = processPreviousSheet(
      memberSS,
      templateData,
      memberName
    );
    const newFormSheet = createNewFormSheet(
      memberSS,
      templateData,
      previousSheet
    );
    const infoSheet = updateInfoSheet(memberSS, templateData, memberName);

    // シート順の整理
    organizeSheetOrder(memberSS, newFormSheet, infoSheet, currentFormSheet);

    // シート構成の整理（不要なシートの削除と順番の整理）
    organizeUpdateFormsSheets(memberSS, memberName);

    // 初期化処理
    initializeFormSheet(newFormSheet, memberName);
  } catch (e) {
    throw new Error(`シート処理中にエラー: ${e.message}`);
  }
}

// URLからファイルIDを抽出
function extractFileIdFromUrl(url) {
  const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
}

// 前回分シートの処理
function processPreviousSheet(ss, templateData, memberName) {
  // === ① 「前回分」シートの処理 ===
  let previousSheet = ss.getSheetByName(SHEET_NAMES.SHIFT_FORM_PREVIOUS);
  if (previousSheet) {
    previousSheet.setName("TEMP_OLD");
    previousSheet
      .getProtections(SpreadsheetApp.ProtectionType.SHEET)
      .forEach((protection) => protection.remove());
  }

  // === ② 現在のシフト希望表を「前回分」にリネーム＆保護 ===
  let currentFormSheet = ss.getSheetByName(SHEET_NAMES.SHIFT_FORM);
  if (!currentFormSheet) {
    // 現在のシフト希望表が存在しない場合は、テンプレートからコピーして作成
    currentFormSheet = templateData.sheet.copyTo(ss);
    currentFormSheet.setName(SHEET_NAMES.SHIFT_FORM_PREVIOUS);
    protectSheet(currentFormSheet, "前回分シートのロック");
    Logger.log(`📝 テンプレートから前回分シートを作成: ${memberName}`);
  } else {
    currentFormSheet.setName(SHEET_NAMES.SHIFT_FORM_PREVIOUS);
    protectSheet(currentFormSheet, "前回分シートのロック");
  }

  return {
    currentFormSheet,
    previousSheet,
  };
}

// 新しい提出用シートの作成
function createNewFormSheet(ss, templateData, previousSheet) {
  // === ③ 新しい提出用シートを作成 ===
  let newFormSheet = previousSheet
    ? previousSheet
    : templateData.sheet.copyTo(ss);
  newFormSheet.setName(SHEET_NAMES.SHIFT_FORM);

  // 2行目以降のデータを貼り付け（1行目は変更しない）
  const dataRows = templateData.values.slice(1); // 1行目を除く
  const dataRowCount = dataRows.length;
  newFormSheet
    .getRange(2, 1, dataRowCount, templateData.numCols)
    .setValues(dataRows);

  // 余分な行を削除
  const lastRow = newFormSheet.getLastRow();
  if (lastRow > dataRowCount + 1) {
    newFormSheet.deleteRows(dataRowCount + 2, lastRow - dataRowCount - 1);
  }

  return newFormSheet;
}

// 今後の勤務希望シートの更新
function updateInfoSheet(ss, templateData, memberName) {
  // === ④ 「今後の勤務希望」シートの取得 ===
  let infoSheet = ss.getSheetByName(SHEET_NAMES.SHIFT_FORM_INFO);
  if (!infoSheet) {
    // 今後の勤務希望シートが存在しない場合は、テンプレートからコピーして作成
    infoSheet = templateData.sheet.copyTo(ss);
    infoSheet.setName(SHEET_NAMES.SHIFT_FORM_INFO);
    Logger.log(`📝 テンプレートから今後の勤務希望シートを作成: ${memberName}`);
  } else {
    // 🔓 シート保護を解除
    const protections = infoSheet.getProtections(
      SpreadsheetApp.ProtectionType.SHEET
    );
    protections.forEach((protection) => protection.remove());
  }

  // リセット
  resetInfoSheetContent(infoSheet);

  return infoSheet;
}

// 今後の勤務希望シートの内容をリセット
function resetInfoSheetContent(infoSheet) {
  // バッチ処理で複数範囲を同時にクリア
  const RANGES_TO_CLEAR = {
    WORK_DAYS: "D1", // 希望勤務日数
    SCHOOL_INFO: "B5:C7", // 校舎情報
    BASIC_SHIFT: "F5:H11", // 基本シフト
    LESSON_DUTY: "K5:P11", // 授業担当
  };

  Object.values(RANGES_TO_CLEAR).forEach((range) => {
    infoSheet.getRange(range).clearContent();
  });

  Logger.log("🧹 今後の勤務希望シートの内容をリセット");
}

// シート順の整理
function organizeSheetOrder(ss, newFormSheet, infoSheet, currentFormSheet) {
  // === ⑤ シート順の整理 ===
  const SHEET_ORDER = {
    SUBMISSION_FORM: 1, // 提出用
    FUTURE_PREFERENCES: 2, // 今後の勤務希望
    PREVIOUS_FORM: 3, // 前回分
  };

  const moveSheet = (sheet, index) => {
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(index);
  };

  moveSheet(newFormSheet, SHEET_ORDER.SUBMISSION_FORM);
  moveSheet(infoSheet, SHEET_ORDER.FUTURE_PREFERENCES);
  moveSheet(currentFormSheet, SHEET_ORDER.PREVIOUS_FORM);
}

// シート構成の整理（不要なシートの削除と順番の整理）
function organizeUpdateFormsSheets(memberSS, memberName) {
  try {
    const allSheets = memberSS.getSheets();

    // 保持するシート名のリスト（順番通り）
    const targetSheetNames = [
      SHEET_NAMES.SHIFT_FORM, // ①シフト希望表
      SHEET_NAMES.SHIFT_FORM_INFO, // ②今後の勤務希望
      SHEET_NAMES.SHIFT_FORM_PREVIOUS, // ③前回分
    ];

    // 不要なシートを削除
    for (const sheet of allSheets) {
      const sheetName = sheet.getName();
      if (!targetSheetNames.includes(sheetName)) {
        try {
          memberSS.deleteSheet(sheet);
          Logger.log(`🗑️ ${memberName} さんの不要シート削除: "${sheetName}"`);
        } catch (deleteError) {
          Logger.log(
            `⚠️ ${memberName} さんのシート削除失敗: "${sheetName}" - ${deleteError.message}`
          );
        }
      }
    }

    // シートの順番を整理
    let currentPosition = 1;
    for (const targetSheetName of targetSheetNames) {
      const targetSheet = memberSS.getSheetByName(targetSheetName);
      if (targetSheet) {
        try {
          memberSS.setActiveSheet(targetSheet);
          memberSS.moveActiveSheet(currentPosition);
          currentPosition++;
        } catch (moveError) {
          Logger.log(
            `⚠️ ${memberName} さんのシート移動失敗: "${targetSheetName}" - ${moveError.message}`
          );
        }
      }
    }

    Logger.log(`✅ ${memberName} さんのシート構成整理完了`);
    return true;
  } catch (e) {
    Logger.log(`❌ ${memberName} さんのシート構成整理でエラー: ${e.message}`);
    return false;
  }
}

// フォームシートの初期化
function initializeFormSheet(newFormSheet, memberName) {
  // === ⑥ 初期化処理 ===
  const headerRow = SHIFT_FORM_TEMPLATE.HEADER.ROW;

  // バッチ処理で初期値を設定
  const initialValues = [
    {
      range: newFormSheet.getRange(
        headerRow,
        SHIFT_FORM_TEMPLATE.HEADER.NAME_COL
      ),
      value: memberName,
    },
    {
      range: newFormSheet.getRange(
        headerRow,
        SHIFT_FORM_TEMPLATE.HEADER.CHECK_COL
      ),
      value: false,
    },
  ];

  // 一括で値を設定
  initialValues.forEach(({ range, value }) => {
    range.setValue(value);
  });

  Logger.log(`✏️ フォームシートを初期化: ${memberName}`);
}

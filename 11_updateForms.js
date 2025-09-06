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

// テンプレートデータと書式の取得
function getTemplateData() {
  const templateFile = SpreadsheetApp.openById(TEMPLATE_FILE_ID);
  const formTemplateSheet = templateFile.getSheetByName(SHEET_NAMES.SHIFT_FORM);

  const templateRange = formTemplateSheet.getDataRange();
  const numRows = templateRange.getNumRows();
  const numCols = templateRange.getNumColumns();
  const values = templateRange.getValues();

  // 日程リスト部分の書式を事前に取得
  const dateListFormatting = getDateListFormatting(formTemplateSheet, numRows);

  return {
    file: templateFile,
    sheet: formTemplateSheet,
    numRows,
    numCols,
    values,
    dateListFormatting, // 書式データを追加
  };
}

/**
 * 日程リスト部分の書式を事前に取得する
 *
 * @param {Sheet} templateSheet - テンプレートシート
 * @param {number} numRows - 総行数
 * @returns {Object} 書式データ
 */
function getDateListFormatting(templateSheet, numRows) {
  try {
    const dateListStartRow = SHIFT_FORM_TEMPLATE.DATA.START_ROW; // 4行目
    const dateListRowCount = numRows - (dateListStartRow - 1); // 4行目以降の行数

    if (dateListRowCount <= 0) {
      return null;
    }

    const templateRange = templateSheet.getRange(
      dateListStartRow,
      1,
      dateListRowCount,
      templateSheet.getLastColumn()
    );

    // 書式を一括取得（最適化）
    const fontColors = templateRange.getFontColors();
    const backgrounds = templateRange.getBackgrounds();
    const fontWeights = templateRange.getFontWeights();
    const fontStyles = templateRange.getFontStyles();

    const formatting = [];
    for (let row = 0; row < dateListRowCount; row++) {
      const rowFormatting = [];
      for (let col = 0; col < templateRange.getNumColumns(); col++) {
        rowFormatting.push({
          fontColor: fontColors[row][col],
          backgroundColor: backgrounds[row][col],
          fontWeight: fontWeights[row][col],
          fontStyle: fontStyles[row][col],
        });
      }
      formatting.push(rowFormatting);
    }

    // Logger.log(`📋 日程リスト部分の書式を事前取得: ${dateListRowCount}行`);
    return {
      startRow: dateListStartRow,
      rowCount: dateListRowCount,
      colCount: templateRange.getNumColumns(),
      formatting: formatting,
    };
  } catch (error) {
    Logger.log(`⚠️ 書式取得でエラー: ${error.message}`);
    return null;
  }
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
      Logger.log(`✅ 処理完了: ${name}`);
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
  // === ① 残存シートのクリーンアップ ===
  // 前回の処理で残った可能性のあるシートを削除
  const cleanupSheetNames = ["TEMP_OLD", "TEMP_NEW", "TEMP"];
  cleanupSheetNames.forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      try {
        ss.deleteSheet(sheet);
        Logger.log(`🧹 残存シートを削除: ${sheetName} (${memberName})`);
      } catch (e) {
        Logger.log(
          `⚠️ 残存シート削除失敗: ${sheetName} (${memberName}) - ${e.message}`
        );
      }
    }
  });

  // === ② 「前回分」シートの処理 ===
  let previousSheet = ss.getSheetByName(SHEET_NAMES.SHIFT_FORM_PREVIOUS);
  if (previousSheet) {
    try {
      previousSheet.setName("TEMP_OLD");
      previousSheet
        .getProtections(SpreadsheetApp.ProtectionType.SHEET)
        .forEach((protection) => protection.remove());
    } catch (e) {
      Logger.log(`⚠️ 前回分シート処理でエラー: ${memberName} - ${e.message}`);
      // エラーが発生した場合は、シートを削除して続行
      try {
        ss.deleteSheet(previousSheet);
        previousSheet = null;
      } catch (deleteError) {
        Logger.log(
          `⚠️ 前回分シート削除失敗: ${memberName} - ${deleteError.message}`
        );
      }
    }
  }

  // === ③ 現在のシフト希望表を「前回分」にリネーム＆保護 ===
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
  const targetRange = newFormSheet.getRange(
    2,
    1,
    dataRowCount,
    templateData.numCols
  );

  // 値を設定
  targetRange.setValues(dataRows);

  // 事前に取得した書式データを適用
  // 書式データが存在する場合のみ適用（最適化）
  if (templateData.dateListFormatting) {
    applyDateListFormatting(templateData.dateListFormatting, newFormSheet);
  }

  // 余分な行を削除
  const lastRow = newFormSheet.getLastRow();
  if (lastRow > dataRowCount + 1) {
    newFormSheet.deleteRows(dataRowCount + 2, lastRow - dataRowCount - 1);
  }

  return newFormSheet;
}

/**
 * 事前に取得した書式データを日程リスト部分に適用する
 *
 * @param {Object} dateListFormatting - 事前に取得した書式データ
 * @param {Sheet} targetSheet - 対象シート
 */
function applyDateListFormatting(dateListFormatting, targetSheet) {
  try {
    if (!dateListFormatting) {
      Logger.log("⚠️ 書式データが存在しないため、書式適用をスキップしました");
      return;
    }

    const { startRow, rowCount, colCount, formatting } = dateListFormatting;

    // 書式を一括適用（最適化）
    const targetRange = targetSheet.getRange(startRow, 1, rowCount, colCount);

    // 2次元配列を準備
    const fontColors = [];
    const backgrounds = [];
    const fontWeights = [];
    const fontStyles = [];

    for (let row = 0; row < rowCount; row++) {
      const fontColorRow = [];
      const backgroundRow = [];
      const fontWeightRow = [];
      const fontStyleRow = [];

      for (let col = 0; col < colCount; col++) {
        const cellFormatting = formatting[row][col];
        fontColorRow.push(cellFormatting.fontColor);
        backgroundRow.push(cellFormatting.backgroundColor);
        fontWeightRow.push(cellFormatting.fontWeight);
        fontStyleRow.push(cellFormatting.fontStyle);
      }

      fontColors.push(fontColorRow);
      backgrounds.push(backgroundRow);
      fontWeights.push(fontWeightRow);
      fontStyles.push(fontStyleRow);
    }

    // 一括で書式を適用
    targetRange.setFontColors(fontColors);
    targetRange.setBackgrounds(backgrounds);
    targetRange.setFontWeights(fontWeights);
    targetRange.setFontStyles(fontStyles);

    // Logger.log(`✅ 日程リスト部分（${startRow}行目以降）の書式を適用しました`);
  } catch (error) {
    Logger.log(`⚠️ 書式適用でエラー: ${error.message}`);
    // エラーが発生しても処理は続行
  }
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

  // Logger.log("🧹 今後の勤務希望シートの内容をリセット");
}

// シート順の整理
function organizeSheetOrder(ss, newFormSheet, infoSheet, currentFormSheet) {
  // === ⑤ シート順の整理 ===
  const SHEET_ORDER = {
    SUBMISSION_FORM: 1, // 提出用
    FUTURE_PREFERENCES: 2, // 今後の勤務希望
    PREVIOUS_FORM: 3, // 前回分
  };

  // シート移動を一括実行（最適化）
  const sheetsToMove = [
    { sheet: newFormSheet, index: SHEET_ORDER.SUBMISSION_FORM },
    { sheet: infoSheet, index: SHEET_ORDER.FUTURE_PREFERENCES },
    { sheet: currentFormSheet, index: SHEET_ORDER.PREVIOUS_FORM },
  ];

  sheetsToMove.forEach(({ sheet, index }) => {
    try {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(index);
    } catch (error) {
      // エラーは無視（シート移動は重要度が低い）
    }
  });
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

    // 不要なシートを一括削除（最適化）
    const sheetsToDelete = allSheets.filter(
      (sheet) => !targetSheetNames.includes(sheet.getName())
    );

    sheetsToDelete.forEach((sheet) => {
      try {
        memberSS.deleteSheet(sheet);
        // Logger.log(`🗑️ ${memberName} さんの不要シート削除: "${sheet.getName()}"`);
      } catch (deleteError) {
        // エラーは無視（シート削除は重要度が低い）
      }
    });

    // シートの順番を一括整理（最適化）
    let currentPosition = 1;
    targetSheetNames.forEach((targetSheetName) => {
      const targetSheet = memberSS.getSheetByName(targetSheetName);
      if (targetSheet) {
        try {
          memberSS.setActiveSheet(targetSheet);
          memberSS.moveActiveSheet(currentPosition);
          currentPosition++;
        } catch (moveError) {
          // エラーは無視（シート移動は重要度が低い）
        }
      }
    });

    // Logger.log(`✅ ${memberName} さんのシート構成整理完了`);
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

  // 初期値を一括設定（最適化）
  const nameRange = newFormSheet.getRange(
    headerRow,
    SHIFT_FORM_TEMPLATE.HEADER.NAME_COL
  );
  const checkRange = newFormSheet.getRange(
    headerRow,
    SHIFT_FORM_TEMPLATE.HEADER.CHECK_COL
  );

  // 並列で値を設定
  nameRange.setValue(memberName);
  checkRange.setValue(false);

  // Logger.log(`✏️ フォームシートを初期化: ${memberName}`);
}

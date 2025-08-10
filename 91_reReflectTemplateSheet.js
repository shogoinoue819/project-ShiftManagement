// デバッグ用シフト希望表テンプレート反映
function reReflectTemplateSheet() {
  const templateSS = SpreadsheetApp.openById(TEMPLATE_FILE_ID);
  const templateSheet = templateSS.getSheetByName(SHEET_NAMES.SHIFT_FORM);
  if (!templateSheet) {
    throw new Error(
      `❌ テンプレートにシート '${SHEET_NAMES.SHIFT_FORM}' が見つかりません`
    );
  }

  const memberMap = createMemberMap();
  let count = 0;

  // 提出ステータス列を取得（提出列は SUBMIT 列）
  const lastRow = getLastRowInCol(
    manageSheet,
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
  );
  const submitValues = manageSheet
    .getRange(
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW,
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.SUBMIT_COL,
      lastRow - SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW + 1,
      1
    )
    .getValues()
    .flat();

  let index = 0;
  for (const [id, { name, url }] of Object.entries(memberMap)) {
    // ===== !前後半スイッチ！ =====
    // 前半
    // if (index >= 30) break; // 30人目以降はスキップ
    // 後半
    if (index < 30) {
      index++;
      continue; // 前半30人はスキップ
    }

    // SUBMIT_FALSE 以外はスキップ
    const submit = submitValues[index];
    if (submit !== SUBMIT_FALSE) {
      index++;
      continue;
    }

    const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!match || !match[1]) {
      Logger.log(`❌ ${name} さんのURLが不正です: ${url}`);
      index++;
      continue;
    }

    const fileId = match[1];
    try {
      const memberSS = SpreadsheetApp.openById(fileId);

      // 既存の「シフト希望表」シートを削除
      const existingSheet = memberSS.getSheetByName(SHEET_NAMES.SHIFT_FORM);
      if (existingSheet) {
        memberSS.deleteSheet(existingSheet);
      }

      // コピーしてリネーム
      const copiedSheet = templateSheet.copyTo(memberSS);
      copiedSheet.setName(SHEET_NAMES.SHIFT_FORM);
      memberSS.setActiveSheet(copiedSheet);
      memberSS.moveActiveSheet(1); // 位置を調整

      // 初期化処理
      copiedSheet
        .getRange(
          SHIFT_FORM_TEMPLATE.HEADER.ROW,
          SHIFT_FORM_TEMPLATE.HEADER.NAME_COL
        )
        .setValue(name);
      copiedSheet
        .getRange(
          SHIFT_FORM_TEMPLATE.HEADER.ROW,
          SHIFT_FORM_TEMPLATE.HEADER.CHECK_COL
        )
        .setValue(false);

      Logger.log(`✅ ${name} さんに「シフト希望表」シートを再反映しました`);
      count++;
      index++;
    } catch (e) {
      Logger.log(`❌ ${name} さんの処理中にエラー: ${e.message}`);
      index++;
    }
  }
  Logger.log(
    `✅ 完了：${count} 名に「シフト希望表」シートを上書き反映しました`
  );
}

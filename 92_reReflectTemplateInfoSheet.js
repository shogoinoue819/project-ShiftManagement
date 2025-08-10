// デバッグ用今後の勤務希望テンプレート反映
function reReflectTemplateInfoSheet() {
  const templateSS = SpreadsheetApp.openById(TEMPLATE_FILE_ID);
  const templateSheet = templateSS.getSheetByName(SHEET_NAMES.SHIFT_FORM_INFO);
  if (!templateSheet) {
    throw new Error(
      `❌ テンプレートにシート '${SHEET_NAMES.SHIFT_FORM_INFO}' が見つかりません`
    );
  }

  const memberMap = createMemberMap();
  let count = 0;

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

    const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!match || !match[1]) {
      Logger.log(`❌ ${name} さんのURLが不正です: ${url}`);
      continue;
    }

    const fileId = match[1];
    try {
      const memberSS = SpreadsheetApp.openById(fileId);

      // 既存の「今後の勤務希望」シートを削除
      const existingSheet = memberSS.getSheetByName(
        SHEET_NAMES.SHIFT_FORM_INFO
      );
      if (existingSheet) {
        memberSS.deleteSheet(existingSheet);
      }

      // コピーしてリネーム
      const copiedSheet = templateSheet.copyTo(memberSS);
      copiedSheet.setName(SHEET_NAMES.SHIFT_FORM_INFO);
      memberSS.setActiveSheet(copiedSheet);
      memberSS.moveActiveSheet(2); // 位置を調整

      Logger.log(`✅ ${name} さんに「今後の勤務希望」シートを再反映しました`);
      count++;
      index++;
    } catch (e) {
      Logger.log(`❌ ${name} さんの処理中にエラー: ${e.message}`);
    }
  }
  Logger.log(
    `✅ 完了：${count} 名に「今後の勤務希望」シートを上書き反映しました`
  );
}

// 個別ファイルのシフト希望表をアップデート
function updateForms() {

  // SSをまとめて取得
  const [ss, manageSheet, templateSheet, allSheets, ui] = getCommonSheets();
  
  // 確認ダイアログを表示
  const confirm = ui.alert(
    "⚠️確認",
    "この操作で、全ての個別ファイルの中身が更新されます。\n個別ファイルの現在の入力内容は前回分として保存されます。\n\n本当に実行してよろしいですか？",
    ui.ButtonSet.OK_CANCEL
  );

  // OKが押されなければキャンセル
  if (confirm !== ui.Button.OK) {
    ui.alert("❌ 操作はキャンセルされました");
    return;
  }


  // ====== 個別ファイルのシフト希望表をアップデート ======

  // メンバーマップを作成
  const memberMap = createMemberMap();

  // チェック列をリセット
  manageSheet.getRange(ROW_START, COLUMN_CHECK, Object.keys(memberMap).length, 1).setValue(false);
  // 反映ステータス列をリセット
  manageSheet.getRange(ROW_START, COLUMN_REFLECT, Object.keys(memberMap).length, 1).setValue(REFLECT_FALSE);

  // テンプレートファイルとシフト希望表テンプレートシートを取得
  const templateFile = SpreadsheetApp.openById(TEMPLATE_FILE_ID);
  const formTemplateSheet = templateFile.getSheetByName(FORM_SHEET_NAME);

  // テンプレートの値と行列数だけ取得してコピー
  const templateRange = formTemplateSheet.getDataRange();
  const numRows = templateRange.getNumRows();
  const numCols = templateRange.getNumColumns();
  const values = templateRange.getValues();

  // 各メンバーにおいて
  for (const [id, { name, url }] of Object.entries(memberMap)) {
    // 個別ファイルを取得
    const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!match || !match[1]) continue;
    const fileId = match[1];

    try {
      const ss = SpreadsheetApp.openById(fileId);

      // === ① 「前回分」シートの処理 ===
      let prevSheet = ss.getSheetByName(FORM_PREVIOUS_SHEET_NAME);
      if (prevSheet) {
        prevSheet.setName("TEMP_OLD");
        prevSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
      }

      // === ② 現在のシフト希望表を「前回分」にリネーム＆保護 ===
      const currSheet = ss.getSheetByName(FORM_SHEET_NAME);
      if (!currSheet) throw new Error("❌ シフト希望表シートが存在しません");
      currSheet.setName(FORM_PREVIOUS_SHEET_NAME);
      protectSheet(currSheet, "前回分シートのロック");

      // === ③ 新しい提出用シートを作成 ===
      let newFormSheet = prevSheet 
        ? prevSheet 
        : formTemplateSheet.copyTo(ss);
      newFormSheet.setName(FORM_SHEET_NAME);
      // 2行目以降のデータを貼り付け（1行目は変更しない）
      const dataOnly = values.slice(1); // 1行目を除く
      const dataNumRows = dataOnly.length;
      newFormSheet.getRange(2, 1, dataNumRows, numCols).setValues(dataOnly);
      // 余分な行を削除
      const maxRow = newFormSheet.getLastRow();
      if (maxRow > dataNumRows + 1) {
        newFormSheet.deleteRows(dataNumRows + 2, maxRow - dataNumRows - 1);
      }

      // === ④ 「今後の勤務希望」シートの取得 ===
      const infoSheet = ss.getSheetByName(FORM_INFO_SHEET_NAME);
      if (!infoSheet) throw new Error("❌ 今後の勤務希望シートが存在しません");
      // 🔓 シート保護を解除
      const protections = infoSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      protections.forEach(p => p.remove());
      // リセット
      infoSheet.getRange("D1").clearContent();           // 希望勤務日数
      infoSheet.getRange("B5:C7").clearContent();        // 校舎情報
      infoSheet.getRange("F5:H11").clearContent();       // 基本シフト
      infoSheet.getRange("K5:P11").clearContent();       // 授業担当

      // === ⑤ シート順の整理 ===
      const moveSheet = (sheet, index) => {
        ss.setActiveSheet(sheet);
        ss.moveActiveSheet(index);
      };
      moveSheet(newFormSheet, 1);   // 提出用
      moveSheet(infoSheet, 2);      // 今後の勤務希望
      moveSheet(currSheet, 3);      // 前回分

      // === ⑥ 初期化処理 ===
      newFormSheet.getRange(FORM_ROW_HEAD, FORM_COLUMN_NAME).setValue(name);
      newFormSheet.getRange(FORM_ROW_HEAD, FORM_COLUMN_CHECK).setValue(false);

      Logger.log(`✅ アップデート完了: ${name}`);

    } catch (e) {
      Logger.log(`❌ エラー: ${name} - ${e.message}`);
    }
  }

  ui.alert("✅ シフト希望表の個別ファイルをすべて更新しました！");
}



// 入力されたメンバーのファイルとシートを削除
function deleteSelectedMember() {
  // SSをまとめて取得
  const ss = getSpreadsheet();
  const manageSheet = getManageSheet();
  const templateSheet = getTemplateSheet();
  const ui = getUI();

  // 氏名の入力
  const response = ui.prompt(
    "削除対象の氏名を入力してください",
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert("キャンセルされました");
    return;
  }
  // 空白などをトリミングして入力された氏名を取得
  const inputName = response.getResponseText().trim();
  // 未入力ならばアラート
  if (!inputName) {
    ui.alert("❌ 氏名が入力されていません");
    return;
  }

  // 入力された氏名から個別シートを取得
  const targetSheet = ss.getSheetByName(inputName);
  // 取得できれば、削除
  if (targetSheet) {
    ss.deleteSheet(targetSheet);
  }

  // メンバーマップ作成
  const memberManager = getMemberManager(manageSheet);
  // 初期化を確実に行う
  if (!memberManager.ensureInitialized()) {
    ui.alert("❌ メンバーデータの初期化に失敗しました");
    return;
  }
  const memberMap = memberManager.memberMap;

  // メンバーマップの妥当性チェック
  if (!memberMap || Object.keys(memberMap).length === 0) {
    ui.alert("❌ メンバーデータが取得できませんでした");
    return;
  }

  // 各メンバーについて、
  for (const [id, { name, url }] of Object.entries(memberMap)) {
    // 氏名が一致すれば、
    if (name === inputName) {
      // URLからファイルIDを抽出
      const match = url.match(
        /https:\/\/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/
      );
      // ファイルIDがあれば、
      if (match && match[1]) {
        // ファイルIDを取得
        const fileId = match[1];
        // ドライブから、ファイルIDに一致するファイルを探し、削除
        try {
          DriveApp.getFileById(fileId).setTrashed(true);
          // エラー処理
        } catch (e) {
          ui.alert(`⚠️ 個別ファイル削除失敗: ${e.message}`);
        }
      }
      // シフト管理シートからその人の行を削除（A列の修正もセットで）
      const order = memberManager.getOrderById(id);
      if (order !== -1) {
        // A列（SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW以降）の日程リストを保存
        const dateValues = manageSheet
          .getRange(
            SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW,
            SHIFT_MANAGEMENT_SHEET.DATE_LIST.COL,
            getLastRowInColumn(
              manageSheet,
              SHIFT_MANAGEMENT_SHEET.DATE_LIST.COL
            ) -
              SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW +
              1,
            1
          )
          .getValues();
        // 対象の行を deleteRow
        manageSheet.deleteRow(
          order + SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW
        );
        // 日程リストをA列に書き戻し
        manageSheet
          .getRange(
            SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW,
            SHIFT_MANAGEMENT_SHEET.DATE_LIST.COL,
            dateValues.length,
            1
          )
          .setValues(dateValues);

        // ===== 前回用管理シート =====
        // 前回用管理シートを取得
        const manageSheetPre = ss.getSheetByName(
          SHEET_NAMES.SHIFT_MANAGEMENT_PREVIOUS
        );
        // A列（SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW以降）の日程リストを保存
        const dateValuesPre = manageSheetPre
          .getRange(
            SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW,
            SHIFT_MANAGEMENT_SHEET.DATE_LIST.COL,
            getLastRowInColumn(
              manageSheetPre,
              SHIFT_MANAGEMENT_SHEET.DATE_LIST.COL
            ) -
              SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW +
              1,
            1
          )
          .getValues();
        // 対象の行を deleteRow
        manageSheetPre.deleteRow(
          order + SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW
        );
        // 日程リストをA列に書き戻し
        manageSheetPre
          .getRange(
            SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW,
            SHIFT_MANAGEMENT_SHEET.DATE_LIST.COL,
            dateValuesPre.length,
            1
          )
          .setValues(dateValuesPre);
      }

      ui.alert(
        `✅「${inputName}」さんの個別ファイルと個別シートを削除しました！`
      );
      return;
    }
  }

  // 削除対象者がいなかった場合
  ui.alert(`❌「${inputName}」さんは管理リストに見つかりませんでした。`);
}

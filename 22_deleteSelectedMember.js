// 入力されたメンバーのファイルとシートを削除
function deleteSelectedMember() {
  // SSをまとめて取得
  const ss = getSpreadsheet();
  const manageSheet = getManageSheet();
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
          Logger.log(`✅ ファイル削除成功: ${inputName} (${fileId})`);
        } catch (e) {
          Logger.log(
            `❌ ファイル削除失敗: ${inputName} (${fileId}) - ${e.message}`
          );
          ui.alert(`⚠️ 個別ファイル削除失敗: ${e.message}`);
        }
      }
      // シフト管理シートからその人の行を削除（A列の修正もセットで）
      const order = memberManager.getOrderById(id);
      if (order !== -1) {
        try {
          // メイン管理シートから削除
          deleteMemberRowWithDateListPreservation(manageSheet, order);
          Logger.log(
            `✅ メイン管理シートから削除: ${inputName} (行${
              order + SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW
            })`
          );

          // ===== 前回用管理シート =====
          const manageSheetPre = ss.getSheetByName(
            SHEET_NAMES.SHIFT_MANAGEMENT_PREVIOUS
          );
          if (manageSheetPre) {
            deleteMemberRowWithDateListPreservation(manageSheetPre, order);
            Logger.log(`✅ 前回用管理シートから削除: ${inputName}`);
          } else {
            Logger.log(
              `⚠️ 前回用管理シートが見つかりません: ${SHEET_NAMES.SHIFT_MANAGEMENT_PREVIOUS}`
            );
          }
        } catch (e) {
          Logger.log(`❌ 管理シート削除エラー: ${inputName} - ${e.message}`);
          ui.alert(`⚠️ 管理シートからの削除に失敗しました: ${e.message}`);
          return;
        }
      } else {
        Logger.log(`⚠️ メンバーの順序が見つかりません: ${inputName}`);
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

// ===== ヘルパー関数 =====
function deleteMemberRowWithDateListPreservation(sheet, memberOrder) {
  try {
    // A列（日程リスト）を保存
    const dateValues = sheet
      .getRange(
        SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW,
        SHIFT_MANAGEMENT_SHEET.DATE_LIST.COL,
        getLastRowInColumn(sheet, SHIFT_MANAGEMENT_SHEET.DATE_LIST.COL) -
          SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW +
          1,
        1
      )
      .getValues();

    // 対象の行を削除
    sheet.deleteRow(memberOrder + SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW);

    // 日程リストをA列に書き戻し
    sheet
      .getRange(
        SHIFT_MANAGEMENT_SHEET.DATE_LIST.START_ROW,
        SHIFT_MANAGEMENT_SHEET.DATE_LIST.COL,
        dateValues.length,
        1
      )
      .setValues(dateValues);
  } catch (e) {
    Logger.log(`❌ メンバー行削除エラー (${sheet.getName()}): ${e.message}`);
    throw e;
  }
}

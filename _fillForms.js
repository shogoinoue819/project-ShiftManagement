// テスト用フォーム入力
function fillForms() {
  // SSをまとめて取得
  const [ss, manageSheet, templateSheet, allSheets, ui] = getCommonSheets();

  // 希望ステータスのオプション
  const statusOptions = [STATUS_TRUE, STATUS_TRUE, STATUS_FALSE];
  // 希望時間のオプション
  const timeOptions = [
    ["13:00", "指定なし"],
    ["10:00", "指定なし"],
    ["指定なし", "18:00"],
    ["14:00", "20:00"],
    ["指定なし", "17:30"],
    ["11:00", "16:00"],
    ["13:00", "指定なし"],
    ["18:00", "22:00"],
    ["16:00", "22:00"],
    [null, null],
    [null, null]
  ];

  // 最終行を取得
  const lastRow = getLastRowInCol(manageSheet, COLUMN_START);
  // IDと氏名のデータを取得
  const data = manageSheet.getRange(ROW_START, COLUMN_ID, lastRow - ROW_START + 1, COLUMN_NAME - COLUMN_ID + 1).getValues();
  // URLデータを取得
  const urls = manageSheet.getRange(ROW_START, COLUMN_URL, lastRow - ROW_START + 1, 1).getFormulas();
  // 提出ステータスデータを取得
  const submits = manageSheet.getRange(ROW_START, COLUMN_SUBMIT, lastRow - ROW_START + 1, 1).getValues();

  // メンバーマップ作成(提出ステータス付き)
  const memberMap = {};
  data.forEach(([id, name], i) => {
    memberMap[id] = {
      name,
      url: urls[i][0],
      submit: submits[i][0]
    };
  });

  // 全てのメンバーにおいて
  for (const [id, { name, url, submit }] of Object.entries(memberMap)) {
    // 提出済みであれば終了
    if (submit !== SUBMIT_FALSE || !url) continue;
    // URLから個別ファイルを取得
    const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!match || !match[1]) {
      Logger.log(`⚠️ URL抽出失敗: ${name}`);
      continue;
    }
    try {
      const fileId = match[1];
      const file = SpreadsheetApp.openById(fileId);
      // シフト希望表シートを取得
      const sheet = file.getSheetByName(FORM_SHEET_NAME);
      if (!sheet) {
        Logger.log(`⚠️ シートが見つかりません: ${name}`);
        continue;
      }
      // 最終行を取得
      const lastRow = sheet.getLastRow();
      // 表開始行以降の全ての行において
      for (let r = FORM_ROW_START; r <= lastRow; r++) {
        // 希望ステータスをランダムにセット
        const status = statusOptions[Math.floor(Math.random() * statusOptions.length)];
        sheet.getRange(r, FORM_COLUMN_STATUS).setValue(status);
        //　希望ステータスが◯なら
        if (status === STATUS_TRUE) {
          // 開始時間と終了時間をランダムにセット
          const [start, end] = timeOptions[Math.floor(Math.random() * timeOptions.length)];
          sheet.getRange(r, FORM_COLUMN_START_TIME, 1, 2).setValues([[start, end]]);
        } else {
          sheet.getRange(r, FORM_COLUMN_START_TIME, 1, 2).clearContent();
        }
        // 提出チェックをつける
        sheet.getRange(FORM_ROW_HEAD, FORM_COLUMN_CHECK).setValue(true);
      }

      Logger.log(`✅ テスト用記入完了: ${name}`);
    } catch (e) {
      Logger.log(`❌ エラー: ${name} - ${e}`);
    }
  }
}

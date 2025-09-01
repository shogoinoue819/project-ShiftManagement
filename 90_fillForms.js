// テスト用フォーム入力
function fillForms() {
  // SSをまとめて取得
  const manageSheet = getManageSheet();

  // テストデータ設定
  const TEST_DATA = {
    statusOptions: [
      STATUS_STRINGS.SHIFT_WISH.TRUE,
      STATUS_STRINGS.SHIFT_WISH.TRUE,
      STATUS_STRINGS.SHIFT_WISH.FALSE,
    ],
    timeOptions: [
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
      [null, null],
    ],
  };

  // 最終行を取得
  const lastRow = getLastRowInColumn(
    manageSheet,
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
  );
  // IDと氏名のデータを取得
  const data = manageSheet
    .getRange(
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW,
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.ID_COL,
      lastRow - SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW + 1,
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.NAME_COL -
        SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.ID_COL +
        1
    )
    .getValues();
  // URLデータを取得
  const urls = manageSheet
    .getRange(
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW,
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.URL_COL,
      lastRow - SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW + 1,
      1
    )
    .getFormulas();
  // 提出ステータスデータを取得
  const submits = manageSheet
    .getRange(
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW,
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.SUBMIT_COL,
      lastRow - SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW + 1,
      1
    )
    .getValues();

  // メンバーマップ作成(提出ステータス付き)
  const memberMap = {};
  data.forEach(([id, name], i) => {
    memberMap[id] = {
      name,
      url: urls[i][0],
      submit: submits[i][0],
    };
  });

  // 全てのメンバーにおいて
  let processedCount = 0;
  let errorCount = 0;

  for (const [id, { name, url, submit }] of Object.entries(memberMap)) {
    // 提出済みであれば終了
    if (submit !== STATUS_STRINGS.SUBMIT.FALSE || !url) continue;

    try {
      const success = fillMemberForm(name, url, TEST_DATA);
      if (success) {
        processedCount++;
        Logger.log(`✅ テスト用記入完了: ${name}`);
      } else {
        errorCount++;
      }
    } catch (e) {
      errorCount++;
      Logger.log(`❌ エラー: ${name} - ${e.message}`);
    }
  }

  Logger.log(`✅ 処理完了: ${processedCount}件成功, ${errorCount}件失敗`);
}

// ===== ヘルパー関数 =====
function fillMemberForm(memberName, url, testData) {
  try {
    // URLから個別ファイルを取得
    const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!match || !match[1]) {
      Logger.log(`⚠️ URL抽出失敗: ${memberName}`);
      return false;
    }

    const fileId = match[1];
    const file = SpreadsheetApp.openById(fileId);

    // シフト希望表シートを取得
    const sheet = file.getSheetByName(SHEET_NAMES.SHIFT_FORM);
    if (!sheet) {
      Logger.log(`⚠️ シートが見つかりません: ${memberName}`);
      return false;
    }

    // 最終行を取得
    const lastRow = sheet.getLastRow();

    // 表開始行以降の全ての行において
    for (let r = SHIFT_FORM_TEMPLATE.DATA.START_ROW; r <= lastRow; r++) {
      // 希望ステータスをランダムにセット
      const status =
        testData.statusOptions[
          Math.floor(Math.random() * testData.statusOptions.length)
        ];
      sheet.getRange(r, SHIFT_FORM_TEMPLATE.DATA.STATUS_COL).setValue(status);

      // 希望ステータスが◯なら
      if (status === STATUS_STRINGS.SHIFT_WISH.TRUE) {
        // 開始時間と終了時間をランダムにセット
        const [start, end] =
          testData.timeOptions[
            Math.floor(Math.random() * testData.timeOptions.length)
          ];
        sheet
          .getRange(r, SHIFT_FORM_TEMPLATE.DATA.START_TIME_COL, 1, 2)
          .setValues([[start, end]]);
      } else {
        sheet
          .getRange(r, SHIFT_FORM_TEMPLATE.DATA.START_TIME_COL, 1, 2)
          .clearContent();
      }

      // 提出チェックをつける
      sheet
        .getRange(
          SHIFT_FORM_TEMPLATE.HEADER.ROW,
          SHIFT_FORM_TEMPLATE.HEADER.CHECK_COL
        )
        .setValue(true);
    }

    return true;
  } catch (e) {
    Logger.log(`❌ フォーム記入エラー (${memberName}): ${e.message}`);
    return false;
  }
}

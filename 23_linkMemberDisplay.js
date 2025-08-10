// メンバーリスト表示をシフトテンプレートにリンクさせる
function linkMemberDisplay() {
  // SSをまとめて取得
  const [ss, manageSheet, templateSheet, allSheets, ui] = getCommonSheets();

  // 最終行を取得
  const lastRow = getLastRowInCol(
    manageSheet,
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
  );
  // 表示名リストを取得
  const nameRange = manageSheet.getRange(
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW,
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.DISPLAY_NAME_COL,
    lastRow - SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW + 1,
    1
  );
  const rawNames = nameRange.getValues().flat();
  // 空白セルが存在するかチェック
  if (
    rawNames.some((name) => name === "" || name === null || name === undefined)
  ) {
    ui.alert(
      "⚠️ 表示名リストに空白のセルがあります。\nすべてのメンバーに名前を入力してください。"
    );
    return; // 処理を中止
  }
  // 空白を除いた有効な名前リスト
  const names = rawNames.filter((name) => name);
  //　背景色リストを取得
  const rawColors = nameRange.getBackgrounds().flat();
  const bgColors = rawColors.map((color) => (color ? color : "white"));

  // テンプレートシートの最終列を取得
  const lastCol = templateSheet.getLastColumn();

  // テンプレートシートのメンバー欄を取得
  const targetRange = templateSheet.getRange(
    SHIFT_TEMPLATE_SHEET.ROWS.MEMBERS,
    SHIFT_TEMPLATE_SHEET.MEMBER_START_COL,
    1,
    lastCol - 1
  );
  // 内容をクリア
  targetRange.clearContent();
  // 背景色をクリア
  targetRange.setBackground(null);
  // 灰色背景をクリア
  templateSheet
    .getRange(
      SHIFT_TEMPLATE_SHEET.ROWS.DATA_START,
      SHIFT_TEMPLATE_SHEET.MEMBER_START_COL,
      SHIFT_TEMPLATE_SHEET.ROWS.DATA_END -
        SHIFT_TEMPLATE_SHEET.ROWS.DATA_START +
        1,
      lastCol - 1
    )
    .setBackground(null);

  // テンプレートシートに氏名と背景色をセット
  for (let i = 0; i < names.length; i++) {
    templateSheet
      .getRange(
        SHIFT_TEMPLATE_SHEET.ROWS.MEMBERS,
        i + SHIFT_TEMPLATE_SHEET.MEMBER_START_COL
      )
      .setValue(names[i]);
    templateSheet
      .getRange(
        SHIFT_TEMPLATE_SHEET.ROWS.MEMBERS,
        i + SHIFT_TEMPLATE_SHEET.MEMBER_START_COL
      )
      .setBackground(bgColors[i]);
  }

  // 背景を灰色に
  templateSheet
    .getRange(
      SHIFT_TEMPLATE_SHEET.ROWS.DATA_START,
      SHIFT_TEMPLATE_SHEET.MEMBER_START_COL,
      SHIFT_TEMPLATE_SHEET.ROWS.DATA_END -
        SHIFT_TEMPLATE_SHEET.ROWS.DATA_START +
        1,
      names.length
    )
    .setBackground(TIME_SETTINGS.UNAVAILABLE_BACKGROUND_COLOR);

  // 授業割テンプレートシートにも反映
  const templateMap = {
    Mon: SHEET_NAMES.LESSON_TEMPLATES.MON,
    Tue: SHEET_NAMES.LESSON_TEMPLATES.TUE,
    Wed: SHEET_NAMES.LESSON_TEMPLATES.WED,
    Thu: SHEET_NAMES.LESSON_TEMPLATES.THU,
    Fri: SHEET_NAMES.LESSON_TEMPLATES.FRI,
  };

  // 各曜日のテンプレートシートに氏名＋背景色を反映
  for (const day in templateMap) {
    const sheetName = templateMap[day];
    const sheet = allSheets.find((s) => s.getName() === sheetName);
    if (!sheet) continue;
    const lastCol = sheet.getLastColumn();
    // メンバー欄の内容・背景色をリセット（2列目以降）
    if (lastCol >= 2) {
      const targetRange = sheet.getRange(
        SHIFT_TEMPLATE_SHEET.ROWS.MEMBERS,
        SHIFT_TEMPLATE_SHEET.MEMBER_START_COL,
        1,
        lastCol - 1
      );
      targetRange.clearContent();
      targetRange.setBackground(null);
    }
    // 氏名と背景色を1人ずつ反映
    for (let i = 0; i < names.length; i++) {
      sheet
        .getRange(
          SHIFT_TEMPLATE_SHEET.ROWS.MEMBERS,
          i + SHIFT_TEMPLATE_SHEET.MEMBER_START_COL
        )
        .setValue(names[i]);
      sheet
        .getRange(
          SHIFT_TEMPLATE_SHEET.ROWS.MEMBERS,
          i + SHIFT_TEMPLATE_SHEET.MEMBER_START_COL
        )
        .setBackground(bgColors[i]);
    }
  }

  // 出勤・退勤・勤務時間の数式をセット
  for (let i = 0; i < names.length; i++) {
    const col = i + SHIFT_TEMPLATE_SHEET.MEMBER_START_COL;
    const colLetter = columnToLetter(col);

    // 出勤
    templateSheet
      .getRange(SHIFT_TEMPLATE_SHEET.ROWS.WORK_START, col)
      .setFormula(
        `=IFERROR(TO_TEXT(INDEX(${colLetter}${
          SHIFT_TEMPLATE_SHEET.ROWS.DATA_START - 1
        }:${colLetter}${
          SHIFT_TEMPLATE_SHEET.ROWS.DATA_END + 1
        }, MATCH(TRUE, ISNUMBER(SEARCH(":" , TO_TEXT(${colLetter}${
          SHIFT_TEMPLATE_SHEET.ROWS.DATA_START - 1
        }:${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.DATA_END + 1}))), 0))), "")`
      );

    // 退勤
    templateSheet
      .getRange(SHIFT_TEMPLATE_SHEET.ROWS.WORK_END, col)
      .setFormula(
        `=IFERROR(TO_TEXT(INDEX(${colLetter}${
          SHIFT_TEMPLATE_SHEET.ROWS.DATA_START - 1
        }:${colLetter}${
          SHIFT_TEMPLATE_SHEET.ROWS.DATA_END + 1
        }, MAX(FILTER(ROW(${colLetter}${
          SHIFT_TEMPLATE_SHEET.ROWS.DATA_START - 1
        }:${colLetter}${
          SHIFT_TEMPLATE_SHEET.ROWS.DATA_END + 1
        })-ROW(${colLetter}${
          SHIFT_TEMPLATE_SHEET.ROWS.DATA_START - 1
        })+1, ISNUMBER(SEARCH(":" , TO_TEXT(${colLetter}${
          SHIFT_TEMPLATE_SHEET.ROWS.DATA_START - 1
        }:${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.DATA_END + 1}))))))), "")`
      );

    // 勤務時間
    templateSheet
      .getRange(SHIFT_TEMPLATE_SHEET.ROWS.WORKING_TIME, col)
      .setFormula(
        `=IF(AND(ISNUMBER(TIMEVALUE(${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.WORK_END})), ISNUMBER(TIMEVALUE(${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.WORK_START}))), TEXT(TIMEVALUE(${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.WORK_END}) - TIMEVALUE(${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.WORK_START}), "h:mm"), "")`
      );
  }
}

// シフト表末尾に新規メンバーを追加
function addNewMember() {
  const ss = getSpreadsheet();
  const manageSheet = getManageSheet();
  const templateSheet = getTemplateSheet();
  const allSheets = ss.getSheets();
  const ui = getUI();

  // 1. 氏名を入力
  const response = ui.prompt(
    "新規追加するメンバーの表示名を入力してください",
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert("キャンセルされました");
    return;
  }
  const inputName = response.getResponseText().trim();
  if (!inputName) {
    ui.alert("❌ 表示名が入力されていません");
    return;
  }

  // 2. 管理シートから表示名と背景色を取得
  const memberInfo = findMemberInfo(manageSheet, inputName);
  if (!memberInfo) {
    ui.alert("⚠️ 入力された表示名が管理シートに存在しません");
    return;
  }

  const { displayName, bgColor } = memberInfo;
  Logger.log(`メンバー情報取得: ${displayName}, 背景色: ${bgColor}`);

  // 3. テンプレートシートの現在の最終列を取得（+1で新列に）
  const newCol = getLastColumnInRow(templateSheet, 1) + 1;

  // 4-6. テンプレートシートに新メンバー列を追加
  addMemberColumnToSheet(templateSheet, newCol, displayName, bgColor);

  // 7. すべての日付形式シートに同様に追加
  let processedCount = 0;
  for (const sheet of allSheets) {
    const name = sheet.getName();
    if (isDateFormatSheet(name)) {
      try {
        addMemberColumnToSheet(sheet, newCol, displayName, bgColor);
        processedCount++;
        Logger.log(`✅ ${name} に追加完了`);
      } catch (e) {
        Logger.log(`❌ ${name} への追加失敗: ${e.message}`);
      }
    }
  }

  Logger.log(`✅ ${processedCount}個の日付シートに追加完了`);

  // 8. 授業テンプレートシートにも追加
  const lessonTemplates = [
    SHEET_NAMES.LESSON_TEMPLATES.MON,
    SHEET_NAMES.LESSON_TEMPLATES.TUE,
    SHEET_NAMES.LESSON_TEMPLATES.WED,
    SHEET_NAMES.LESSON_TEMPLATES.THU,
    SHEET_NAMES.LESSON_TEMPLATES.FRI,
  ];

  let lessonCount = 0;
  for (const templateName of lessonTemplates) {
    const lessonSheet = ss.getSheetByName(templateName);
    if (lessonSheet) {
      try {
        addMemberColumnToSheet(lessonSheet, newCol, displayName, bgColor);
        lessonCount++;
        Logger.log(`✅ ${templateName} に追加完了`);
      } catch (e) {
        Logger.log(`❌ ${templateName} への追加失敗: ${e.message}`);
      }
    } else {
      Logger.log(`⚠️ ${templateName} シートが見つかりません`);
    }
  }

  Logger.log(`✅ ${lessonCount}個の授業テンプレートシートに追加完了`);

  ui.alert(`✅ ${inputName} さんを各シートの末尾に追加しました`);
}

// ===== ヘルパー関数 =====
function addMemberColumnToSheet(sheet, newCol, displayName, bgColor) {
  // 表示名と背景色を設定
  const memberCell = sheet.getRange(SHIFT_TEMPLATE_SHEET.ROWS.MEMBERS, newCol);
  memberCell.setValue(displayName);
  memberCell.setBackground(bgColor);

  // 授業テンプレートシートの場合はグレー背景設定をスキップ
  const sheetName = sheet.getName();
  const isLessonTemplate = [
    SHEET_NAMES.LESSON_TEMPLATES.MON,
    SHEET_NAMES.LESSON_TEMPLATES.TUE,
    SHEET_NAMES.LESSON_TEMPLATES.WED,
    SHEET_NAMES.LESSON_TEMPLATES.THU,
    SHEET_NAMES.LESSON_TEMPLATES.FRI,
  ].includes(sheetName);

  if (!isLessonTemplate) {
    // 勤務エリアに灰色背景を設定（授業テンプレート以外）
    sheet
      .getRange(
        SHIFT_TEMPLATE_SHEET.ROWS.DATA_START,
        newCol,
        SHIFT_TEMPLATE_SHEET.ROWS.DATA_END -
          SHIFT_TEMPLATE_SHEET.ROWS.DATA_START +
          1
      )
      .setBackground(TIME_SETTINGS.UNAVAILABLE_BACKGROUND_COLOR);
  }

  // 出勤・退勤・勤務時間の数式を設定
  const colLetter = convertColumnToLetter(newCol);

  sheet
    .getRange(SHIFT_TEMPLATE_SHEET.ROWS.WORK_START, newCol)
    .setFormula(generateWorkStartFormula(colLetter));

  sheet
    .getRange(SHIFT_TEMPLATE_SHEET.ROWS.WORK_END, newCol)
    .setFormula(generateWorkEndFormula(colLetter));

  sheet
    .getRange(SHIFT_TEMPLATE_SHEET.ROWS.WORKING_TIME, newCol)
    .setFormula(generateWorkingTimeFormula(colLetter));
}

function isDateFormatSheet(sheetName) {
  return /^\d{1,2}\/\d{1,2}$/.test(sheetName);
}

function findMemberInfo(manageSheet, inputName) {
  try {
    const lastRow = getLastRowInColumn(
      manageSheet,
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
    );

    if (lastRow < SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW) {
      Logger.log("⚠️ 管理シートにメンバーデータがありません");
      return null;
    }

    // 名前列を一括取得
    const nameRange = manageSheet.getRange(
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW,
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.NAME_COL,
      lastRow - SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW + 1,
      1
    );
    const names = nameRange.getValues().flat();

    // 対象メンバーのインデックスを検索
    const index = names.findIndex((name) => String(name).trim() === inputName);
    if (index === -1) {
      Logger.log(`⚠️ メンバーが見つかりません: ${inputName}`);
      return null;
    }

    // 表示名と背景色を取得
    const targetRow = SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW + index;
    const displayRange = manageSheet.getRange(
      targetRow,
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.DISPLAY_NAME_COL
    );

    return {
      displayName: displayRange.getValue(),
      bgColor: displayRange.getBackground(),
    };
  } catch (e) {
    Logger.log(`❌ メンバー情報取得エラー: ${e.message}`);
    return null;
  }
}

function generateWorkStartFormula(colLetter) {
  const startRow = SHIFT_TEMPLATE_SHEET.ROWS.DATA_START - 1;
  const endRow = SHIFT_TEMPLATE_SHEET.ROWS.DATA_END + 1;
  const range = `${colLetter}${startRow}:${colLetter}${endRow}`;

  return `=IFERROR(TO_TEXT(INDEX(${range}, MATCH(TRUE, ISNUMBER(SEARCH(":" , TO_TEXT(${range}))), 0))), "")`;
}

function generateWorkEndFormula(colLetter) {
  const startRow = SHIFT_TEMPLATE_SHEET.ROWS.DATA_START - 1;
  const endRow = SHIFT_TEMPLATE_SHEET.ROWS.DATA_END + 1;
  const range = `${colLetter}${startRow}:${colLetter}${endRow}`;

  return `=IFERROR(TO_TEXT(INDEX(${range}, MAX(FILTER(ROW(${range})-ROW(${colLetter}${startRow})+1, ISNUMBER(SEARCH(":" , TO_TEXT(${range}))))))), "")`;
}

function generateWorkingTimeFormula(colLetter) {
  const workStart = `${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.WORK_START}`;
  const workEnd = `${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.WORK_END}`;

  return `=IF(AND(ISNUMBER(TIMEVALUE(${workEnd})), ISNUMBER(TIMEVALUE(${workStart}))), TEXT(TIMEVALUE(${workEnd}) - TIMEVALUE(${workStart}), "h:mm"), "")`;
}

// シフト表末尾に新規メンバーを追加
function addNewMember() {
  const ss = getSpreadsheet();
  const manageSheet = getManageSheet();
  const templateSheet = getTemplateSheet();
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
  const lastRow = getLastRowInColumn(
    manageSheet,
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
  );
  const nameRange = manageSheet.getRange(
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW,
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.NAME_COL,
    lastRow - SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW + 1,
    1
  );
  const names = nameRange.getValues().flat();

  const index = names.findIndex((name) => name === inputName);
  if (index === -1) {
    ui.alert("⚠️ 入力された表示名が管理シートに存在しません");
    return;
  }
  const displayRange = manageSheet.getRange(
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW + index,
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.DISPLAY_NAME_COL
  );
  const displayName = displayRange.getValue();
  const bgColor = displayRange.getBackground();
  Logger.log(`${bgColor}`);

  // 3. テンプレートシートの現在の最終列を取得（+1で新列に）
  const newCol = getLastColumnInRow(templateSheet, 1) + 1;

  // 4. 表示名と背景色を追加
  templateSheet
    .getRange(SHIFT_TEMPLATE_SHEET.ROWS.MEMBERS, newCol)
    .setValue(displayName);
  templateSheet
    .getRange(SHIFT_TEMPLATE_SHEET.ROWS.MEMBERS, newCol)
    .setBackground(bgColor);

  // 5. 灰色背景を勤務エリアにセット
  templateSheet
    .getRange(
      SHIFT_TEMPLATE_SHEET.ROWS.DATA_START,
      newCol,
      SHIFT_TEMPLATE_SHEET.ROWS.DATA_END -
        SHIFT_TEMPLATE_SHEET.ROWS.DATA_START +
        1
    )
    .setBackground(TIME_SETTINGS.UNAVAILABLE_BACKGROUND_COLOR);

  // 6. 出勤・退勤・勤務時間の数式をセット
  const colLetter = convertColumnToLetter(newCol);
  templateSheet
    .getRange(SHIFT_TEMPLATE_SHEET.ROWS.WORK_START, newCol)
    .setFormula(
      `=IFERROR(TO_TEXT(INDEX(${colLetter}${
        SHIFT_TEMPLATE_SHEET.ROWS.DATA_START - 1
      }:${colLetter}${
        SHIFT_TEMPLATE_SHEET.ROWS.DATA_END + 1
      }, MATCH(TRUE, ISNUMBER(SEARCH(":" , TO_TEXT(${colLetter}${
        SHIFT_TEMPLATE_SHEET.ROWS.DATA_START - 1
      }:${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.DATA_END + 1}))), 0))), "")`
    );
  templateSheet
    .getRange(SHIFT_TEMPLATE_SHEET.ROWS.WORK_END, newCol)
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
  templateSheet
    .getRange(SHIFT_TEMPLATE_SHEET.ROWS.WORKING_TIME, newCol)
    .setFormula(
      `=IF(AND(ISNUMBER(TIMEVALUE(${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.WORK_END})), ISNUMBER(TIMEVALUE(${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.WORK_START}))), TEXT(TIMEVALUE(${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.WORK_END}) - TIMEVALUE(${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.WORK_START}), "h:mm"), "")`
    );

  // 7. すべての日付形式シートに同様に追加
  for (const sheet of allSheets) {
    const name = sheet.getName();
    if (/^\d{1,2}\/\d{1,2}$/.test(name)) {
      // 表示名と背景色をセット
      sheet
        .getRange(SHIFT_TEMPLATE_SHEET.ROWS.MEMBERS, newCol)
        .setValue(displayName);
      sheet
        .getRange(SHIFT_TEMPLATE_SHEET.ROWS.MEMBERS, newCol)
        .setBackground(bgColor);

      // 勤務エリアに灰色
      sheet
        .getRange(
          SHIFT_TEMPLATE_SHEET.ROWS.DATA_START,
          newCol,
          SHIFT_TEMPLATE_SHEET.ROWS.DATA_END -
            SHIFT_TEMPLATE_SHEET.ROWS.DATA_START +
            1
        )
        .setBackground(TIME_SETTINGS.UNAVAILABLE_BACKGROUND_COLOR);

      // 数式セット
      const colLetter = convertColumnToLetter(newCol);
      sheet
        .getRange(SHIFT_TEMPLATE_SHEET.ROWS.WORK_START, newCol)
        .setFormula(
          `=IFERROR(TO_TEXT(INDEX(${colLetter}${
            SHIFT_TEMPLATE_SHEET.ROWS.DATA_START - 1
          }:${colLetter}${
            SHIFT_TEMPLATE_SHEET.ROWS.DATA_END + 1
          }, MATCH(TRUE, ISNUMBER(SEARCH(":" , TO_TEXT(${colLetter}${
            SHIFT_TEMPLATE_SHEET.ROWS.DATA_START - 1
          }:${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.DATA_END + 1}))), 0))), "")`
        );
      sheet
        .getRange(SHIFT_TEMPLATE_SHEET.ROWS.WORK_END, newCol)
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
      sheet
        .getRange(SHIFT_TEMPLATE_SHEET.ROWS.WORKING_TIME, newCol)
        .setFormula(
          `=IF(AND(ISNUMBER(TIMEVALUE(${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.WORK_END})), ISNUMBER(TIMEVALUE(${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.WORK_START}))), TEXT(TIMEVALUE(${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.WORK_END}) - TIMEVALUE(${colLetter}${SHIFT_TEMPLATE_SHEET.ROWS.WORK_START}), "h:mm"), "")`
        );
    }

    Logger.log(`${name}完了`);
  }

  ui.alert(`✅ ${inputName} さんをテンプレートシートの末尾に追加しました`);
}

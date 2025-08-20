// 曜日別授業割を反映
function reflectLessonTemplate() {
  // SSをまとめて取得
  const ss = getSpreadsheet();
  const manageSheet = getManageSheet();
  const templateSheet = getTemplateSheet();
  const allSheets = ss.getSheets();
  const ui = getUI();

  // 曜日とテンプレート名の対応マップ
  const templateMap = {
    Mon: SHEET_NAMES.LESSON_TEMPLATES.MON,
    Tue: SHEET_NAMES.LESSON_TEMPLATES.TUE,
    Wed: SHEET_NAMES.LESSON_TEMPLATES.WED,
    Thu: SHEET_NAMES.LESSON_TEMPLATES.THU,
    Fri: SHEET_NAMES.LESSON_TEMPLATES.FRI,
  };
  // ターゲットはシフト作成シート
  const targetSheets = allSheets.filter((s) =>
    /^\d{1,2}\/\d{1,2}$/.test(s.getName())
  );

  // 各日程のシフト作成シートにおいて、
  targetSheets.forEach((dailySheet) => {
    // シート名を取得
    const sheetName = dailySheet.getName();
    // A1から日付を取得
    const date = dailySheet
      .getRange(SHIFT_TEMPLATE_SHEET.DATE_ROW, SHIFT_TEMPLATE_SHEET.DATE_COL)
      .getValue();
    // 日付から曜日を取得
    const dayOfWeek = formatDateToString(date, "E"); // Mon ~ Sun
    // 月〜金に含まれる場合のみ処理
    if (templateMap.hasOwnProperty(dayOfWeek)) {
      // 曜日に対応した授業割シートを取得
      const lessonTemplateSheet = ss.getSheetByName(templateMap[dayOfWeek]);
      if (!lessonTemplateSheet) {
        Logger.log(`テンプレートが見つかりません: ${templateMap[dayOfWeek]}`);
        return;
      }
      // 取得する列数
      const numCols =
        dailySheet.getLastColumn() - SHIFT_TEMPLATE_SHEET.MEMBER_START_COL + 1;
      // 授業割シートのデータ元範囲を取得
      const sourceRange = lessonTemplateSheet.getRange(
        SHIFT_TEMPLATE_SHEET.ROWS.DATA_START,
        SHIFT_TEMPLATE_SHEET.MEMBER_START_COL,
        SHIFT_TEMPLATE_SHEET.ROWS.DATA_END -
          SHIFT_TEMPLATE_SHEET.ROWS.DATA_START +
          1,
        numCols
      );
      // シフト作成シートの貼り付け先範囲を取得
      const destRange = dailySheet.getRange(
        SHIFT_TEMPLATE_SHEET.ROWS.DATA_START,
        SHIFT_TEMPLATE_SHEET.MEMBER_START_COL,
        SHIFT_TEMPLATE_SHEET.ROWS.DATA_END -
          SHIFT_TEMPLATE_SHEET.ROWS.DATA_START +
          1,
        numCols
      );

      // 授業割シートの結合セル情報を取得
      const mergedRanges = sourceRange.getMergedRanges();
      // すべてのプロパティを取得
      const values = sourceRange.getValues();
      const backgrounds = sourceRange.getBackgrounds();
      const fontColors = sourceRange.getFontColors();
      const fontSizes = sourceRange.getFontSizes();
      const fontWeights = sourceRange.getFontWeights();

      // 元の背景（上書き対象）を取得
      const currentBackgrounds = destRange.getBackgrounds();
      // 貼り付け用の新背景配列を元の背景からコピー
      const newBackgrounds = currentBackgrounds.map((row) => [...row]);

      // 各セルにおいて、
      for (let i = 0; i < backgrounds.length; i++) {
        for (let j = 0; j < backgrounds[i].length; j++) {
          // 反映する背景色
          const sourceColor = backgrounds[i][j];
          // 元が白背景（#ffffff）もしくはnullの場合、何もしない（元の背景を維持）
          if (sourceColor !== "#ffffff" && sourceColor !== null) {
            newBackgrounds[i][j] = sourceColor;
          }
        }
      }
      // 新しい背景色だけをセット
      destRange.setBackgrounds(newBackgrounds);

      // 他のプロパティの貼り付け実行
      destRange.setValues(values);
      destRange.setFontColors(fontColors);
      destRange.setFontSizes(fontSizes);
      destRange.setFontWeights(fontWeights);
      // ※setBorder は別途処理が必要（高度）

      // セルの結合
      mergedRanges.forEach((range) => {
        const rowOffset = range.getRow() - SHIFT_TEMPLATE_SHEET.ROWS.DATA_START;
        const colOffset =
          range.getColumn() - SHIFT_TEMPLATE_SHEET.MEMBER_START_COL;
        const targetRange = dailySheet.getRange(
          SHIFT_TEMPLATE_SHEET.ROWS.DATA_START + rowOffset,
          SHIFT_TEMPLATE_SHEET.MEMBER_START_COL + colOffset,
          range.getNumRows(),
          range.getNumColumns()
        );
        targetRange.merge();
      });

      // ボーダーをセット
      const pastedRange = dailySheet.getRange(
        SHIFT_TEMPLATE_SHEET.ROWS.DATA_START,
        SHIFT_TEMPLATE_SHEET.MEMBER_START_COL,
        SHIFT_TEMPLATE_SHEET.ROWS.DATA_END -
          SHIFT_TEMPLATE_SHEET.ROWS.DATA_START +
          1,
        numCols
      );
      applyBorders(pastedRange);

      Logger.log(`テンプレートを適用: ${sheetName}`);
    }
  });

  ui.alert("✅ 授業割テンプレを反映しました！");
}

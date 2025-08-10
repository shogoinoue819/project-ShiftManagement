// シフト希望を反映
function reflectShiftForms() {
  // SSを取得
  const [ss, manageSheet, templateSheet, allSheets, ui] = getCommonSheets();
  // 最終行を取得
  const lastRow = getLastRowInCol(
    manageSheet,
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
  );
  // 反映ステータスまでのデータを取得
  const data = manageSheet
    .getRange(
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW,
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL,
      lastRow - SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW + 1,
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.REFLECT_COL -
        SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL +
        1
    )
    .getValues();

  // 提出済み、チェック済み、未反映でフィルタリングしてマップを作成
  const filtered = data
    .map((row, i) => ({
      id: row[
        SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.ID_COL -
          SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
      ],
      name: row[
        SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.NAME_COL -
          SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
      ],
      order: i,
      submit:
        row[
          SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.SUBMIT_COL -
            SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
        ],
      check:
        row[
          SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.CHECK_COL -
            SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
        ],
      reflect:
        row[
          SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.REFLECT_COL -
            SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
        ],
    }))
    .filter(
      ({ submit, check, reflect }) =>
        submit === SUBMIT_TRUE && check === true && reflect === REFLECT_FALSE
    )
    .slice(0, 15); // ← 上限制限

  // ターゲットはシフト作成シート
  const targetSheets = allSheets.filter((s) =>
    /^\d{1,2}\/\d{1,2}$/.test(s.getName())
  );

  // 各日程のシフト作成シートにおいて、
  targetSheets.forEach((dailySheet) => {
    // シート名から日程の文字列を取得
    const dateStr = dailySheet.getName();
    // A1から日程を取得
    const date = dailySheet
      .getRange(SHIFT_ROW_DATE, SHIFT_COLUMN_DATE)
      .getValue();
    // シフト希望表のその日程の行を取得
    const dateRow = getOrderByDate(date) + SHIFT_FORM_TEMPLATE.DATA.START_ROW;

    // フィルタリングされたメンバーにおいて、
    filtered.forEach(({ id, name, order }) => {
      // シフト作成シートにおけるその人の列を取得
      const personalCol = order + SHIFT_COLUMN_START;
      // 個別シートを取得
      const personalSheet = ss.getSheetByName(name);
      if (!personalSheet) return;
      // その日程のデータを取得
      const rowData = personalSheet
        .getRange(dateRow, 1, 1, SHIFT_FORM_TEMPLATE.DATA.NOTE_COL)
        .getValues()[0];

      // 備考だけ先に取得してセット
      const note = rowData[SHIFT_FORM_TEMPLATE.DATA.NOTE_COL - 1];
      dailySheet.getRange(SHIFT_ROW_NOTE, personalCol).setValue(note);
      // 希望ステータスを取得
      const status = rowData[SHIFT_FORM_TEMPLATE.DATA.STATUS_COL - 1];
      // ◯でないなら、ここで終了
      if (status != STATUS_TRUE) return;

      // 開始時間と終了時間を取得
      const start = rowData[SHIFT_FORM_TEMPLATE.DATA.START_TIME_COL - 1];
      const end = rowData[SHIFT_FORM_TEMPLATE.DATA.END_TIME_COL - 1];

      // 背景色の配列
      const bgArray = Array(timeList.length).fill(UNAVAILABLE_COLOR);

      // 書式によって分類
      const formatTime = (value) => {
        // nillもしくは"指定なし"なら、fullを返す
        if (!value || value === "指定なし") {
          return "full";
        }
        // 有効なDate型の日程なら、フォーマット化して返す
        if (
          Object.prototype.toString.call(value) === "[object Date]" &&
          !isNaN(value)
        ) {
          return Utilities.formatDate(value, "Asia/Tokyo", "H:mm");
        }
        return "error";
      };
      // [開始時間, 終了時間] の配列をセット
      const timeArray = [formatTime(start), formatTime(end)];
      dailySheet
        .getRange(SHIFT_ROW_START_TIME, personalCol, 2, 1)
        .setValues(timeArray.map((v) => [v]));

      // 開始時間と終了時間を日付に紐づけて取得
      const base = new Date(
        date.getFullYear(),
        date.getMonth(),
        date.getDate()
      );
      const startTime =
        normalizeTimeToDate(base, start) ||
        new Date(
          base.getFullYear(),
          base.getMonth(),
          base.getDate(),
          DEFAULT_OPEN_HOUR,
          DEFAULT_OPEN_MINUTE
        );
      const endTime =
        normalizeTimeToDate(base, end) ||
        new Date(
          base.getFullYear(),
          base.getMonth(),
          base.getDate(),
          DEFAULT_CLOSE_HOUR,
          DEFAULT_CLOSE_MINUTE
        );

      // 各時間帯において、
      timeList.forEach((t, ti) => {
        // セルの時間帯を取得
        const [h, m] = t.split(":").map(Number);
        const cellTime = new Date(
          base.getFullYear(),
          base.getMonth(),
          base.getDate(),
          h,
          m
        );
        // 開始時間と終了時間の間ならば、背景色をnullにセット
        if (startTime <= cellTime && cellTime < endTime) {
          bgArray[ti] = null;
        }
      });
      // 背景色を一括で反映
      dailySheet
        .getRange(SHIFT_ROW_START, personalCol, timeList.length, 1)
        .setBackgrounds(bgArray.map((c) => [c]));
    });

    Logger.log(`✅ ${dateStr}の色反映が完了しました。`);
  });

  filtered.forEach(({ order }) => {
    manageSheet
      .getRange(
        order + SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW,
        SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.REFLECT_COL
      )
      .setValue(REFLECT_TRUE);
  });

  Logger.log("✅ シフト希望の色反映が完了しました。");
  ui.alert("✅ チェック済みのシフト希望を反映しました！");
}

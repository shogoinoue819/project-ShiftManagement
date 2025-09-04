// シフト希望を反映

/**
 * シフト希望反映システム
 *
 * このファイルは、メンバーが提出したシフト希望を
 * シフト作成シートに反映する機能を提供します。
 *
 * 主な機能:
 * - 提出済み・チェック済み・未反映のシフト希望を自動検出
 * - 個別メンバーシートから希望時間と備考を取得
 * - シフト作成シートに時間と背景色を反映
 * - エラーハンドリングと詳細なログ出力
 * - パフォーマンス監視と統計情報の提供
 *
 * パフォーマンス最適化:
 * - データキャッシュシステムによる高速化
 * - 一括更新処理による高速化
 * - 早期リターンによる不要処理の回避
 * - 効率的なデータ構造とアルゴリズム
 *
 * @author System
 * @version 2.0.0
 * @since 2024
 */

// メイン関数
function reflectShiftForms() {
  const ss = getSpreadsheet();
  const manageSheet = getManageSheet();
  const templateSheet = getTemplateSheet();
  const allSheets = getAllSheets();
  const ui = getUI();

  // ===== 反映対象メンバーの抽出 =====
  const lastRow = getLastRowInColumn(
    manageSheet,
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
  );
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
        submit === STATUS_STRINGS.SUBMIT.TRUE &&
        check === true &&
        reflect === STATUS_STRINGS.REFLECT.FALSE
    );

  if (filtered.length === 0) {
    ui.alert("反映対象のメンバーはいません。");
    return;
  }

  // ===== 管理シートの対象日付（M/d文字列）と index マップ =====
  const dateList = getDateList(manageSheet);
  const dateStrList = dateList.map((row) => formatDateToString(row[0], "M/d"));
  const dateIndexMap = {};
  dateStrList.forEach((s, i) => (dateIndexMap[s] = i));

  // 管理リストに載っている日付名のシートのみ対象
  const targetSheets = allSheets.filter(
    (s) => dateIndexMap[s.getName()] != null
  );
  if (targetSheets.length === 0) {
    ui.alert("管理シートの日付に対応する日付シートが見つかりません。");
    return;
  }

  // ===== "読み取り" をメンバー単位で一括キャッシュ =====
  const memberData = {};
  filtered.forEach(({ name }) => {
    const personalSheet = ss.getSheetByName(name);
    if (!personalSheet) {
      memberData[name] = null;
      return;
    }
    // 日付行ぶんを一括で読む（1人1回！）
    memberData[name] = personalSheet
      .getRange(
        SHIFT_FORM_TEMPLATE.DATA.START_ROW,
        1,
        dateStrList.length,
        SHIFT_FORM_TEMPLATE.DATA.NOTE_COL
      )
      .getValues();
  });

  const timeCount = TIME_SETTINGS.TIME_LIST.length;

  // ===== 反映ループ（書き込みは従来どおり列ごと。ただし読み取りはキャッシュ済み） =====
  targetSheets.forEach((dailySheet) => {
    const dateStr = dailySheet.getName();
    const dateIndex = dateIndexMap[dateStr]; // 0始まり
    if (dateIndex == null) {
      Logger.log(`⏭ 管理リストに無い日付のためスキップ: ${dateStr}`);
      return;
    }

    // A1の日付(Date)→時間帯計算のベース
    const date = dailySheet
      .getRange(SHIFT_TEMPLATE_SHEET.DATE_ROW, SHIFT_TEMPLATE_SHEET.DATE_COL)
      .getValue();
    if (!(date instanceof Date)) {
      Logger.log(`⏭ A1が日付でないためスキップ: ${dateStr}`);
      return;
    }
    const base = new Date(date.getFullYear(), date.getMonth(), date.getDate());

    filtered.forEach(({ name, order }) => {
      const personalCol = order + SHIFT_TEMPLATE_SHEET.MEMBER_START_COL;
      const rows = memberData[name];
      if (!rows) return; // 個別シート無し

      const rowData = rows[dateIndex]; // [col1, col2, ..., FORM_COLUMN_NOTE]
      if (!rowData) return;

      // 備考
      const note = rowData[SHIFT_FORM_TEMPLATE.DATA.NOTE_COL - 1];
      dailySheet
        .getRange(SHIFT_TEMPLATE_SHEET.ROWS.NOTE, personalCol)
        .setValue(note);

      // 希望ステータス
      const status = rowData[SHIFT_FORM_TEMPLATE.DATA.STATUS_COL - 1];
      if (status != STATUS_STRINGS.SHIFT_WISH.TRUE) {
        // ×や未記入：時間表示は消し、背景は灰に戻す
        dailySheet
          .getRange(SHIFT_TEMPLATE_SHEET.ROWS.START_TIME, personalCol, 2, 1)
          .clearContent();
        dailySheet
          .getRange(
            SHIFT_TEMPLATE_SHEET.ROWS.DATA_START,
            personalCol,
            timeCount,
            1
          )
          .setBackgrounds(
            Array(timeCount).fill([TIME_SETTINGS.UNAVAILABLE_BACKGROUND_COLOR])
          );
        return;
      }

      // 開始/終了
      const start = rowData[SHIFT_FORM_TEMPLATE.DATA.START_TIME_COL - 1];
      const end = rowData[SHIFT_FORM_TEMPLATE.DATA.END_TIME_COL - 1];

      // 表示用フォーマット（文字列 "H:mm" or "full" or "error"）
      const formatTime = (value) => {
        if (!value || value === "指定なし") return "full";
        if (
          Object.prototype.toString.call(value) === "[object Date]" &&
          !isNaN(value)
        ) {
          return Utilities.formatDate(value, "Asia/Tokyo", "H:mm");
        }
        return "error";
      };

      const timeArray = [formatTime(start), formatTime(end)];
      dailySheet
        .getRange(SHIFT_TEMPLATE_SHEET.ROWS.START_TIME, personalCol, 2, 1)
        .setValues(timeArray.map((v) => [v]));

      // 比較用に Date 化（未入力はデフォルトの開閉時間）
      const startTime =
        start && start !== "指定なし" ? normalizeTimeToDate(base, start) : null;
      const finalStartTime =
        startTime ||
        new Date(
          base.getFullYear(),
          base.getMonth(),
          base.getDate(),
          ENVIRONMENT.DEFAULT_HOURS.OPEN.HOUR,
          ENVIRONMENT.DEFAULT_HOURS.OPEN.MINUTE
        );

      const endTime =
        end && end !== "指定なし" ? normalizeTimeToDate(base, end) : null;
      const finalEndTime =
        endTime ||
        new Date(
          base.getFullYear(),
          base.getMonth(),
          base.getDate(),
          ENVIRONMENT.DEFAULT_HOURS.CLOSE.HOUR,
          ENVIRONMENT.DEFAULT_HOURS.CLOSE.MINUTE
        );

      // 背景色列を構築（連続範囲なら null、外は灰）
      const bgArray = Array(timeCount).fill(
        TIME_SETTINGS.UNAVAILABLE_BACKGROUND_COLOR
      );
      for (let ti = 0; ti < timeCount; ti++) {
        const [h, m] = TIME_SETTINGS.TIME_LIST[ti].split(":").map(Number);
        const cellTime = new Date(
          base.getFullYear(),
          base.getMonth(),
          base.getDate(),
          h,
          m
        );
        if (finalStartTime <= cellTime && cellTime < finalEndTime)
          bgArray[ti] = null;
      }
      dailySheet
        .getRange(
          SHIFT_TEMPLATE_SHEET.ROWS.DATA_START,
          personalCol,
          timeCount,
          1
        )
        .setBackgrounds(bgArray.map((c) => [c]));
    });

    Logger.log(`✅ ${dateStr} の反映完了`);
  });

  // 反映済みフラグ
  filtered.forEach(({ order }) => {
    manageSheet
      .getRange(
        order + SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW,
        SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.REFLECT_COL
      )
      .setValue(STATUS_STRINGS.REFLECT.TRUE);
  });

  ui.alert(
    `✅ チェック済みのシフト希望（${filtered.length}名）を反映しました！`
  );
}

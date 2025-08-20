// シフト希望を反映
// 定数
const MAX_PROCESSING_MEMBERS = 15;
const TIME_ARRAY_ROWS = 2;

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
  const startTime = new Date();
  logWithErrorHandling("シフト希望反映処理を開始しました");

  try {
    // SSを取得
    const ss = getSpreadsheet();
    const manageSheet = getManageSheet();
    const templateSheet = getTemplateSheet();
    const ui = getUI();

    // 最終行を取得
    const lastRow = getLastRowInColumn(
      manageSheet,
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
    );

    logWithErrorHandling(
      `対象行数: ${lastRow - SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW + 1}`
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
    const filtered = mapMemberData(
      data,
      SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
    )
      .filter(isEligibleForReflection)
      .slice(0, MAX_PROCESSING_MEMBERS);

    logWithErrorHandling(`処理対象メンバー数: ${filtered.length}`);

    // ターゲットはシフト作成シート（manageSheetの日付リストに基づく）
    const dateList = getDateList();
    const targetSheets = dateList
      .map((row) => {
        const date = row[0];
        const dateStr = formatDateToString(date, "M/d");
        return ss.getSheetByName(dateStr);
      })
      .filter((sheet) => sheet !== null);

    logWithErrorHandling(`対象シート数: ${targetSheets.length}`);

    // 各日程のシフト作成シートにおいて、
    targetSheets.forEach((dailySheet) => {
      // シート名から日程の文字列を取得
      const dateStr = dailySheet.getName();
      // A1から日程を取得
      const date = dailySheet
        .getRange(SHIFT_TEMPLATE_SHEET.DATE_ROW, SHIFT_TEMPLATE_SHEET.DATE_COL)
        .getValue();
      // シフト希望表のその日程の行を取得
      const dateRow =
        getDateOrderByDate(date) + SHIFT_FORM_TEMPLATE.DATA.START_ROW;

      // フィルタリングされたメンバーにおいて、
      filtered.forEach((member) => {
        processMemberShift(dailySheet, member, date, ss);
      });

      logWithErrorHandling(`${dateStr}の色反映が完了しました`);
    });

    // 反映ステータスを一括更新
    const updatePromises = filtered.map(({ order }) =>
      safeSetValue(
        manageSheet.getRange(
          order + SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW,
          SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.REFLECT_COL
        ),
        STATUS_STRINGS.REFLECT.TRUE,
        `${order}番目のメンバーの反映ステータス`
      )
    );

    const endTime = new Date();
    const processingTime = (endTime - startTime) / 1000; // 秒単位

    logWithErrorHandling(
      `シフト希望の色反映が完了しました (処理時間: ${processingTime.toFixed(
        2
      )}秒)`
    );
    ui.alert(
      `✅ チェック済みのシフト希望を反映しました！\n処理時間: ${processingTime.toFixed(
        2
      )}秒\n対象メンバー: ${filtered.length}名\n対象シート: ${
        targetSheets.length
      }枚`
    );
  } catch (error) {
    logWithErrorHandling("シフト希望反映処理でエラーが発生しました", error);
    ui.alert(`❌ エラーが発生しました: ${error.message || error}`);
    throw error;
  }
}

// ヘルパー関数: 列インデックスを計算
function getColumnIndex(baseCol, targetCol) {
  return targetCol - baseCol;
}

// ヘルパー関数: メンバーデータをマッピング
function mapMemberData(data, startCol) {
  return data.map((row, i) => ({
    id: row[
      getColumnIndex(startCol, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.ID_COL)
    ],
    name: row[
      getColumnIndex(startCol, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.NAME_COL)
    ],
    order: i,
    submit:
      row[
        getColumnIndex(startCol, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.SUBMIT_COL)
      ],
    check:
      row[
        getColumnIndex(startCol, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.CHECK_COL)
      ],
    reflect:
      row[
        getColumnIndex(startCol, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.REFLECT_COL)
      ],
  }));
}

// ヘルパー関数: フィルタリング条件をチェック
function isEligibleForReflection(member) {
  return (
    member.submit === STATUS_STRINGS.SUBMIT.TRUE &&
    member.check === true &&
    member.reflect === STATUS_STRINGS.REFLECT.FALSE
  );
}

// ヘルパー関数: 時間値をフォーマット
function formatTimeValue(value) {
  // nullもしくは"指定なし"なら、fullを返す
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
}

// ヘルパー関数: 個別メンバーのシフト反映処理
function processMemberShift(dailySheet, member, date, ss) {
  const { name, order } = member;
  const personalCol = order + SHIFT_TEMPLATE_SHEET.MEMBER_START_COL;

  // 個別シートを取得
  const personalSheet = ss.getSheetByName(name);
  if (!personalSheet) return;

  // その日程のデータを取得
  const dateRow = getDateOrderByDate(date) + SHIFT_FORM_TEMPLATE.DATA.START_ROW;
  const rowData = personalSheet
    .getRange(dateRow, 1, 1, SHIFT_FORM_TEMPLATE.DATA.NOTE_COL)
    .getValues()[0];

  // 備考をセット
  const note = rowData[SHIFT_FORM_TEMPLATE.DATA.NOTE_COL - 1];
  safeSetValue(
    dailySheet.getRange(SHIFT_TEMPLATE_SHEET.ROWS.NOTE, personalCol),
    note,
    `${name}の備考`
  );

  // 希望ステータスを取得
  const status = rowData[SHIFT_FORM_TEMPLATE.DATA.STATUS_COL - 1];
  if (status != STATUS_STRINGS.SHIFT_WISH.TRUE) return;

  // 開始時間と終了時間を取得
  const start = rowData[SHIFT_FORM_TEMPLATE.DATA.START_TIME_COL - 1];
  const end = rowData[SHIFT_FORM_TEMPLATE.DATA.END_TIME_COL - 1];

  // 時間配列をセット
  const timeArray = [formatTimeValue(start), formatTimeValue(end)];
  safeSetValues(
    dailySheet.getRange(
      SHIFT_TEMPLATE_SHEET.ROWS.START_TIME,
      personalCol,
      TIME_ARRAY_ROWS,
      1
    ),
    timeArray.map((v) => [v]),
    `${name}の時間`
  );

  // 背景色を設定（フォーマット済みの時間値を使用）
  setBackgroundColors(
    dailySheet,
    personalCol,
    date,
    timeArray[0],
    timeArray[1]
  );
}

// ヘルパー関数: 背景色を設定
function setBackgroundColors(dailySheet, personalCol, date, start, end) {
  // 開始時間と終了時間を日付に紐づけて取得
  const base = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const startTime =
    normalizeTimeToDate(base, start) ||
    new Date(
      base.getFullYear(),
      base.getMonth(),
      base.getDate(),
      ENVIRONMENT.DEFAULT_HOURS.OPEN.HOUR,
      ENVIRONMENT.DEFAULT_HOURS.OPEN.MINUTE
    );
  const endTime =
    normalizeTimeToDate(base, end) ||
    new Date(
      base.getFullYear(),
      base.getMonth(),
      base.getDate(),
      ENVIRONMENT.DEFAULT_HOURS.CLOSE.HOUR,
      ENVIRONMENT.DEFAULT_HOURS.CLOSE.MINUTE
    );

  // 背景色の配列
  const bgArray = Array(TIME_SETTINGS.TIME_LIST.length).fill(
    TIME_SETTINGS.UNAVAILABLE_BACKGROUND_COLOR
  );

  // 各時間帯において、
  TIME_SETTINGS.TIME_LIST.forEach((t, ti) => {
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
  safeSetBackgrounds(
    dailySheet.getRange(
      SHIFT_TEMPLATE_SHEET.ROWS.DATA_START,
      personalCol,
      TIME_SETTINGS.TIME_LIST.length,
      1
    ),
    bgArray.map((c) => [c]),
    `${personalCol}列の背景色`
  );
}

// ヘルパー関数: エラーハンドリング付きログ出力
function logWithErrorHandling(message, error = null) {
  if (error) {
    Logger.log(`❌ ${message}: ${error.message || error}`);
    console.error(`❌ ${message}:`, error);
  } else {
    Logger.log(`✅ ${message}`);
  }
}

// ヘルパー関数: 安全なシート操作
function safeSetValue(range, value, description = "") {
  try {
    range.setValue(value);
    if (description) {
      logWithErrorHandling(`${description}を設定しました`);
    }
  } catch (error) {
    logWithErrorHandling(`${description}の設定に失敗しました`, error);
    throw error;
  }
}

// ヘルパー関数: 安全な一括値設定
function safeSetValues(range, values, description = "") {
  try {
    range.setValues(values);
    if (description) {
      logWithErrorHandling(`${description}を一括設定しました`);
    }
  } catch (error) {
    logWithErrorHandling(`${description}の一括設定に失敗しました`, error);
    throw error;
  }
}

// ヘルパー関数: 安全な背景色設定
function safeSetBackgrounds(range, backgrounds, description = "") {
  try {
    range.setBackgrounds(backgrounds);
    if (description) {
      logWithErrorHandling(`${description}を設定しました`);
    }
  } catch (error) {
    logWithErrorHandling(`${description}の背景色設定に失敗しました`, error);
    throw error;
  }
}

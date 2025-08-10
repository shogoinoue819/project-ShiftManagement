// ===== 環境共通定数ファイル =====
//
// 注意: 環境依存の値（ファイルID等）は consts-env.js に分離されました
// このファイルには全環境で共通の定数のみを記載してください

// ===== シート名 =====
const SHEET_NAMES = {
  // シフト管理シート
  SHIFT_MANAGEMENT: "シフト管理",
  SHIFT_MANAGEMENT_PREVIOUS: "シフト管理<前回分>",

  // シフト希望表テンプレート
  SHIFT_FORM: "シフト希望表",
  SHIFT_FORM_INFO: "今後の勤務希望",
  SHIFT_FORM_PREVIOUS: "前回分",

  // シフトテンプレート
  SHIFT_TEMPLATE: "シフトテンプレート",

  // 授業割テンプレート
  LESSON_TEMPLATES: {
    MON: "授業割(月)",
    TUE: "授業割(火)",
    WED: "授業割(水)",
    THU: "授業割(木)",
    FRI: "授業割(金)",
  },
};

// ===== シフト管理シート設定 =====
const SHIFT_MANAGEMENT_SHEET = {
  // 日程リスト
  DATE_LIST: {
    COL: 1,
    START_ROW: 4,
    COMPLETE_COL: 2,
    SHARE_COL: 3,
  },

  // メンバーリスト
  MEMBER_LIST: {
    START_ROW: 4,
    START_COL: 5,
    ID_COL: 5,
    NAME_COL: 6,
    DISPLAY_NAME_COL: 7,
    SUBMIT_COL: 8,
    CHECK_COL: 9,
    REFLECT_COL: 10,
    URL_COL: 11,
    WORK_DATES_1_COL: 12,
    WORK_TIMES_1_COL: 13,
    WORK_DATES_2_COL: 14,
    WORK_TIMES_2_COL: 15,
    WORK_DATES_3_COL: 16,
    WORK_TIMES_3_COL: 17,
    WORK_DATES_4_COL: 18,
    WORK_TIMES_4_COL: 19,
    WORK_DATES_REQUEST_COL: 20,
    EMAIL_COL: 21,
  },
};

// ===== シフト希望表テンプレート設定 =====
const SHIFT_FORM_TEMPLATE = {
  // ヘッダー
  HEADER: {
    ROW: 1,
    NAME_COL: 2,
    INFO_COL: 3,
    CHECK_COL: 4,
  },

  // データ部分
  DATA: {
    START_ROW: 4,
    DATE_COL: 1,
    STATUS_COL: 2,
    START_TIME_COL: 3,
    END_TIME_COL: 4,
    NOTE_COL: 5,
    CONTACT_COL: 6,
  },
};

// ===== シフトテンプレートシート設定 =====
const SHIFT_TEMPLATE_SHEET = {
  // 基本設定
  DATE_ROW: 1,
  DATE_COL: 1,
  MEMBER_START_COL: 2,

  // 行設定
  ROWS: {
    MEMBERS: 1,
    START_TIME: 2,
    END_TIME: 3,
    NOTE: 4,
    WORK_START: 5,
    WORK_END: 6,
    WORKING_TIME: 7,
    DATA_START: 9,
    DATA_END: 36,
  },
};

// ===== 環境共通設定 =====
const ENVIRONMENT = {
  YEAR: 2025,

  // デフォルト開閉室時間
  DEFAULT_HOURS: {
    OPEN: {
      HOUR: 8,
      MINUTE: 0,
    },
    CLOSE: {
      HOUR: 22,
      MINUTE: 0,
    },
  },
};

// ===== 時間設定 =====
const TIME_SETTINGS = {
  // 時間帯リスト
  TIME_LIST: [
    "8:00",
    "8:30",
    "9:00",
    "9:30",
    "10:00",
    "10:30",
    "11:00",
    "11:30",
    "12:00",
    "12:30",
    "13:00",
    "13:30",
    "14:00",
    "14:30",
    "15:00",
    "15:30",
    "16:00",
    "16:30",
    "17:00",
    "17:30",
    "18:00",
    "18:30",
    "19:00",
    "19:30",
    "20:00",
    "20:30",
    "21:00",
    "21:30",
  ],

  // 勤務不可背景色
  UNAVAILABLE_BACKGROUND_COLOR: "#d3d3d3",
};

// ===== ステータス文字列 =====
const STATUS_STRINGS = {
  // 共有ステータス
  SHARE: {
    TRUE: "✅共有済み",
    FALSE: "未共有",
  },

  // 提出ステータス
  SUBMIT: {
    TRUE: "✅提出済み",
    FALSE: "未提出",
  },

  // 反映ステータス
  REFLECT: {
    TRUE: "✅反映済み",
    FALSE: "未反映",
  },

  // シフト希望ステータス
  SHIFT_WISH: {
    TRUE: "◯",
    FALSE: "×",
  },
};

// ===== 勤務日数・労働時間計算関数 =====
const WORK_CALCULATION_FORMULAS = {
  // 第1週
  WEEK_1: {
    DATES: `
      SUM(
        ARRAYFORMULA({
          IFERROR(N(INDIRECT("'" & TEXT($A$4, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4)) <> ""), 0),
          IFERROR(N(INDIRECT("'" & TEXT($A$5, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4)) <> ""), 0),
          IFERROR(N(INDIRECT("'" & TEXT($A$6, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4)) <> ""), 0),
          IFERROR(N(INDIRECT("'" & TEXT($A$7, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4)) <> ""), 0),
          IFERROR(N(INDIRECT("'" & TEXT($A$8, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4)) <> ""), 0),
          IFERROR(N(INDIRECT("'" & TEXT($A$9, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4)) <> ""), 0),
          IFERROR(N(INDIRECT("'" & TEXT($A$10, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4)) <> ""), 0)
        })
      )
    `,
    TIMES: `
      SUM(
        ARRAYFORMULA({
          IFERROR(TIMEVALUE(INDIRECT("'" & TEXT($A$4, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4))), 0),
          IFERROR(TIMEVALUE(INDIRECT("'" & TEXT($A$5, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4))), 0),
          IFERROR(TIMEVALUE(INDIRECT("'" & TEXT($A$6, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4))), 0),
          IFERROR(TIMEVALUE(INDIRECT("'" & TEXT($A$7, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4))), 0),
          IFERROR(TIMEVALUE(INDIRECT("'" & TEXT($A$8, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4))), 0),
          IFERROR(TIMEVALUE(INDIRECT("'" & TEXT($A$9, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4))), 0),
          IFERROR(TIMEVALUE(INDIRECT("'" & TEXT($A$10, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4))), 0)
        })
      )
    `,
  },

  // 第2週
  WEEK_2: {
    DATES: `
      SUM(
        ARRAYFORMULA({
          IFERROR(N(INDIRECT("'" & TEXT($A$11, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4)) <> ""), 0),
          IFERROR(N(INDIRECT("'" & TEXT($A$12, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4)) <> ""), 0),
          IFERROR(N(INDIRECT("'" & TEXT($A$13, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4)) <> ""), 0),
          IFERROR(N(INDIRECT("'" & TEXT($A$14, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4)) <> ""), 0),
          IFERROR(N(INDIRECT("'" & TEXT($A$15, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4)) <> ""), 0),
          IFERROR(N(INDIRECT("'" & TEXT($A$16, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4)) <> ""), 0),
          IFERROR(N(INDIRECT("'" & TEXT($A$17, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4)) <> ""), 0)
        })
      )
    `,
    TIMES: `
      SUM(
        ARRAYFORMULA({
          IFERROR(TIMEVALUE(INDIRECT("'" & TEXT($A$11, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4))), 0),
          IFERROR(TIMEVALUE(INDIRECT("'" & TEXT($A$12, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4))), 0),
          IFERROR(TIMEVALUE(INDIRECT("'" & TEXT($A$13, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4))), 0),
          IFERROR(TIMEVALUE(INDIRECT("'" & TEXT($A$14, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4))), 0),
          IFERROR(TIMEVALUE(INDIRECT("'" & TEXT($A$15, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4))), 0),
          IFERROR(TIMEVALUE(INDIRECT("'" & TEXT($A$16, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4))), 0),
          IFERROR(TIMEVALUE(INDIRECT("'" & TEXT($A$17, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4))), 0)
        })
      )
    `,
  },

  // 第3週
  WEEK_3: {
    DATES: `
      SUM(
        ARRAYFORMULA({
          IFERROR(N(INDIRECT("'" & TEXT($A$18, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4)) <> ""), 0),
          IFERROR(N(INDIRECT("'" & TEXT($A$19, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4)) <> ""), 0),
          IFERROR(N(INDIRECT("'" & TEXT($A$20, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4)) <> ""), 0),
          IFERROR(N(INDIRECT("'" & TEXT($A$21, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4)) <> ""), 0),
          IFERROR(N(INDIRECT("'" & TEXT($A$22, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4)) <> ""), 0),
          IFERROR(N(INDIRECT("'" & TEXT($A$23, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4)) <> ""), 0),
          IFERROR(N(INDIRECT("'" & TEXT($A$24, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4)) <> ""), 0)
        })
      )
    `,
    TIMES: `
      SUM(
        ARRAYFORMULA({
          IFERROR(TIMEVALUE(INDIRECT("'" & TEXT($A$18, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4))), 0),
          IFERROR(TIMEVALUE(INDIRECT("'" & TEXT($A$19, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4))), 0),
          IFERROR(TIMEVALUE(INDIRECT("'" & TEXT($A$20, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4))), 0),
          IFERROR(TIMEVALUE(INDIRECT("'" & TEXT($A$21, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4))), 0),
          IFERROR(TIMEVALUE(INDIRECT("'" & TEXT($A$22, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4))), 0),
          IFERROR(TIMEVALUE(INDIRECT("'" & TEXT($A$23, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4))), 0),
          IFERROR(TIMEVALUE(INDIRECT("'" & TEXT($A$24, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4))), 0)
        })
      )
    `,
  },

  // 第4週
  WEEK_4: {
    DATES: `
      SUM(
        ARRAYFORMULA({
          IFERROR(N(INDIRECT("'" & TEXT($A$25, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4)) <> ""), 0),
          IFERROR(N(INDIRECT("'" & TEXT($A$26, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4)) <> ""), 0),
          IFERROR(N(INDIRECT("'" & TEXT($A$27, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4)) <> ""), 0),
          IFERROR(N(INDIRECT("'" & TEXT($A$28, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4)) <> ""), 0),
          IFERROR(N(INDIRECT("'" & TEXT($A$29, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4)) <> ""), 0),
          IFERROR(N(INDIRECT("'" & TEXT($A$30, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4)) <> ""), 0),
          IFERROR(N(INDIRECT("'" & TEXT($A$31, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4)) <> ""), 0)
        })
      )
    `,
    TIMES: `
      SUM(
        ARRAYFORMULA({
          IFERROR(TIMEVALUE(INDIRECT("'" & TEXT($A$25, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4))), 0),
          IFERROR(TIMEVALUE(INDIRECT("'" & TEXT($A$26, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4))), 0),
          IFERROR(TIMEVALUE(INDIRECT("'" & TEXT($A$27, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4))), 0),
          IFERROR(TIMEVALUE(INDIRECT("'" & TEXT($A$28, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4))), 0),
          IFERROR(TIMEVALUE(INDIRECT("'" & TEXT($A$29, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4))), 0),
          IFERROR(TIMEVALUE(INDIRECT("'" & TEXT($A$30, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4))), 0),
          IFERROR(TIMEVALUE(INDIRECT("'" & TEXT($A$31, "M/d") & "'!" & ADDRESS(7, ROW() - 2, 4))), 0)
        })
      )
    `,
  },
};

// ===== 後方互換性のための定数エイリアス =====
// 既存のコードが動作するように、古い定数名を新しいオブジェクトのプロパティにマッピング

// シフトテンプレートシート
const SHIFT_ROW_DATE = SHIFT_TEMPLATE_SHEET.DATE_ROW;
const SHIFT_COLUMN_DATE = SHIFT_TEMPLATE_SHEET.DATE_COL;
const SHIFT_COLUMN_START = SHIFT_TEMPLATE_SHEET.MEMBER_START_COL;
const SHIFT_ROW_MEMBERS = SHIFT_TEMPLATE_SHEET.ROWS.MEMBERS;
const SHIFT_ROW_START_TIME = SHIFT_TEMPLATE_SHEET.ROWS.START_TIME;
const SHIFT_ROW_END_TIME = SHIFT_TEMPLATE_SHEET.ROWS.END_TIME;
const SHIFT_ROW_NOTE = SHIFT_TEMPLATE_SHEET.ROWS.NOTE;
const SHIFT_ROW_WORK_START = SHIFT_TEMPLATE_SHEET.ROWS.WORK_START;
const SHIFT_ROW_WORK_END = SHIFT_TEMPLATE_SHEET.ROWS.WORK_END;
const SHIFT_ROW_WORKING = SHIFT_TEMPLATE_SHEET.ROWS.WORKING_TIME;
const SHIFT_ROW_START = SHIFT_TEMPLATE_SHEET.ROWS.DATA_START;
const SHIFT_ROW_END = SHIFT_TEMPLATE_SHEET.ROWS.DATA_END;

// 環境共通設定
const THIS_YEAR = ENVIRONMENT.YEAR;
const DEFAULT_OPEN_HOUR = ENVIRONMENT.DEFAULT_HOURS.OPEN.HOUR;
const DEFAULT_OPEN_MINUTE = ENVIRONMENT.DEFAULT_HOURS.OPEN.MINUTE;
const DEFAULT_CLOSE_HOUR = ENVIRONMENT.DEFAULT_HOURS.CLOSE.HOUR;
const DEFAULT_CLOSE_MINUTE = ENVIRONMENT.DEFAULT_HOURS.CLOSE.MINUTE;

// 時間設定
const timeList = TIME_SETTINGS.TIME_LIST;
const UNAVAILABLE_COLOR = TIME_SETTINGS.UNAVAILABLE_BACKGROUND_COLOR;

// ステータス文字列
const SHARE_TRUE = STATUS_STRINGS.SHARE.TRUE;
const SHARE_FALSE = STATUS_STRINGS.SHARE.FALSE;
const SUBMIT_TRUE = STATUS_STRINGS.SUBMIT.TRUE;
const SUBMIT_FALSE = STATUS_STRINGS.SUBMIT.FALSE;
const REFLECT_TRUE = STATUS_STRINGS.REFLECT.TRUE;
const REFLECT_FALSE = STATUS_STRINGS.REFLECT.FALSE;
const STATUS_TRUE = STATUS_STRINGS.SHIFT_WISH.TRUE;
const STATUS_FALSE = STATUS_STRINGS.SHIFT_WISH.FALSE;

// 勤務日数・労働時間計算関数
const WORK_DATES_1 = WORK_CALCULATION_FORMULAS.WEEK_1.DATES;
const WORK_TIMES_1 = WORK_CALCULATION_FORMULAS.WEEK_1.TIMES;
const WORK_DATES_2 = WORK_CALCULATION_FORMULAS.WEEK_2.DATES;
const WORK_TIMES_2 = WORK_CALCULATION_FORMULAS.WEEK_2.TIMES;
const WORK_DATES_3 = WORK_CALCULATION_FORMULAS.WEEK_3.DATES;
const WORK_TIMES_3 = WORK_CALCULATION_FORMULAS.WEEK_3.TIMES;
const WORK_DATES_4 = WORK_CALCULATION_FORMULAS.WEEK_4.DATES;
const WORK_TIMES_4 = WORK_CALCULATION_FORMULAS.WEEK_4.TIMES;

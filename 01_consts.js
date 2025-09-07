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

// ===== ユーティリティ関数用定数 =====
const UTILS_CONSTANTS = {
  // 列・行の基本設定
  COLUMNS: {
    ID_AND_NAME: 2, // ID列と氏名列の2列
    ID_ONLY: 1, // ID列のみの1列
    DATE_ONLY: 1, // 日付列のみの1列
  },

  // 行の基本設定
  ROWS: {
    START_INDEX: 1, // 開始行インデックス
    MIN_INDEX: 1, // 最小行インデックス
  },

  // 日付・時間の制限値
  DATE_LIMITS: {
    MIN_MONTH: 1, // 最小月
    MAX_MONTH: 12, // 最大月
    MIN_DAY: 1, // 最小日
    MAX_DAY: 31, // 最大日
  },

  // 時間の制限値
  TIME_LIMITS: {
    MIN_HOUR: 0, // 最小時
    MAX_HOUR: 23, // 最大時
    MIN_MINUTE: 0, // 最小分
    MAX_MINUTE: 59, // 最大分
  },

  // 文字・ID生成
  ID_GENERATION: {
    MEMBER_ID_LENGTH: 6, // メンバーIDの長さ
    ALPHABET_START: 65, // アルファベット開始コード（A）
    ALPHABET_BASE: 26, // アルファベット基数
  },

  // デフォルト値
  DEFAULTS: {
    EMPTY_STRING: "", // 空文字列
    EMPTY_ARRAY: [], // 空配列
    EMPTY_OBJECT: {}, // 空オブジェクト
    NOT_FOUND: -1, // 見つからない場合の値
    ZERO: 0, // ゼロ値
  },

  // 色コード
  COLORS: {
    WHITE: "#ffffff", // 白色
    BLACK: "#000000", // 黒色
  },

  // 日付フォーマット
  DATE_FORMATS: {
    DEFAULT: "M/d", // デフォルト日付フォーマット
  },

  // 正規表現パターン
  REGEX_PATTERNS: {
    TIME_FORMAT: /^(\d{1,2}):(\d{2})$/, // 時間形式（H:mm または HH:mm）
  },
};

// ===== PDF出力設定 =====
const PDF_EXPORT_CONFIG = {
  PORTRAIT: false,
  SIZE: "A4",
  FIT_WIDTH: true,
  SCALE: 4,
  SHOW_SHEET_NAMES: false,
  SHOW_TITLE: false,
  SHOW_PAGE_NUMBERS: false,
  SHOW_GRIDLINES: false,
  FIX_ROW_HEIGHT: false,
};

// ===== 行高調整設定 =====
const ROW_HEIGHT_MULTIPLIER = 1.5;

// ===== UI表示設定 =====
const UI_DISPLAY = {
  // 進捗表示の更新間隔（人単位）
  PROGRESS_UPDATE_INTERVAL: 5,

  // 進捗表示用のセル位置
  PROGRESS: {
    ROW: 1,
    COL: 1,
  },

  // 実行状況表示用のセル位置
  STATUS: {
    ROW: 1,
    COL: 2,
  },

  // 進捗表示のメッセージ
  PROGRESS_MESSAGES: {
    // シフト希望表更新用
    FORM_UPDATE: {
      PREPARING: "②準備中...",
      PROCESSING: "②実行中...",
    },

    // シート作成用
    SHEET_CREATE: {
      PREPARING: "③準備中...",
      PROCESSING: "③実行中...",
    },

    // チェック処理用
    MEMBER_CHECK: {
      PREPARING: "④準備中...",
      PROCESSING: "④実行中...",
    },

    // シフト希望反映用
    SHIFT_REFLECT: {
      PREPARING: "⑤準備中...",
      PROCESSING: "⑤実行中...",
    },

    // 授業テンプレ反映用
    LESSON_TEMPLATE: {
      PREPARING: "⑥準備中...",
      PROCESSING: "⑥実行中...",
    },
  },
};

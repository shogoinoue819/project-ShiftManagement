// ===== 各種ファイルID =====

// テンプレートファイルID
const TEMPLATE_FILE_ID = "1EgSTvXTuu33kGmpeQOloD7W2hSPy3WHykRnSKdDA-_s";

// シフト表共有ファイルID
const SHARE_FILE_ID = "1CpqkCCMt-CLzKl8bTnfZORRkhWKc_rrN1HpEPZcLkbs";

// 作成済みシフトPDFフォルダID
const SHIFT_PDF_FOLDER_ID = "12PzpZ71xOQhnq4cinpGLH6vFL51Y83rX";

// 作成済みシフトSSフォルダID
const SHIFT_SS_FOLDER_ID = "1YSaTPB611Jme7nDIi5sNL0B4Uf3GUU1v";

// シフト希望表個別フォルダID
const PERSONAL_FORM_FOLDER_ID = "1Gg9p9aEXo1BEstj7DTFa9ba_xeT4mFGs";



// ===== シフト管理シート =====

// シート名
const MANAGE_SHEET = "シフト管理";
const MANAGE_SHEET_PRE = "シフト管理<前回分>";

// 行列インデックス
const MANAGE_DATE_ROW_START = 4; // 日程リスト開始行
const MANAGE_DATE_COLUMN = 1; // 日程リスト列
const MANAGE_COMPLETE_COLUMN = 2; //完成列
const MANAGE_SHARE_COLUMN = 3; // 共有列

const ROW_START = 4; // メンバーリスト開始行
const COLUMN_START = 5; // メンバーリスト開始列
const COLUMN_ID = 5; // ID列
const COLUMN_NAME = 6; // 氏名列
const COLUMN_DISPLAYNAME = 7; // 表示名列
const COLUMN_SUBMIT = 8; // 提出ステータス列
const COLUMN_CHECK = 9; // チェック列
const COLUMN_REFLECT = 10; // 反映ステータス列
const COLUMN_URL = 11; // URL列
const COLUMN_WORK_DATES_1 = 12; // 勤務日数①列
const COLUMN_WORK_TIMES_1 = 13; // 労働時間①列
const COLUMN_WORK_DATES_2 = 14; // 勤務日数②列
const COLUMN_WORK_TIMES_2 = 15; // 労働時間②列
const COLUMN_WORK_DATES_3 = 16; // 勤務日数③列
const COLUMN_WORK_TIMES_3 = 17; // 労働時間③列
const COLUMN_WORK_DATES_4 = 18; // 勤務日数④列
const COLUMN_WORK_TIMES_4 = 19; // 労働時間④列
const COLUMN_WORK_DATES_REQ = 20; // 勤務日数希望列
const COLUMN_EMAIL = 21; // メアド列

// Boolean文字列
const SHARE_TRUE = "✅共有済み"; // 共有true
const SHARE_FALSE = "未共有"; // 共有false
const SUBMIT_TRUE = "✅提出済み"; // 提出true
const SUBMIT_FALSE = "未提出"; // 提出false
const REFLECT_TRUE = "✅反映済み"; // 反映true
const REFLECT_FALSE = "未反映"; // 反映false



// ===== シフト希望表テンプレートファイル =====

// シート名
const FORM_SHEET_NAME = "シフト希望表";
const FORM_INFO_SHEET_NAME = "今後の勤務希望";
const FORM_PREVIOUS_SHEET_NAME = "前回分";

// const FORM_PARSONAL_NAME = "個人用シフト表";


// 行列インデックス
const FORM_ROW_HEAD = 1; // ヘッダー行
const FORM_COLUMN_NAME = 2; // 氏名列
const FORM_COLUMN_INFO = 3; // 勤務日数希望列(隠す)
const FORM_COLUMN_CHECK = 4; // チェック列

const FORM_ROW_START = 4; // 表の開始行
const FORM_COLUMN_DATE = 1; // 日程列
const FORM_COLUMN_STATUS = 2; // ステータス列
const FORM_COLUMN_START_TIME = 3; // 開始時間列
const FORM_COLUMN_END_TIME = 4; // 終了時間列
const FORM_COLUMN_NOTE = 5; // 備考列
const FORM_COLUMN_CONTACT = 6; // 連絡事項列

// Boolean文字列
const STATUS_TRUE = "◯";
const STATUS_FALSE = "×";




// ===== シフトテンプレートシート =====

// シフトテンプレートシート名
const SHIFT_SHEET_NAME = "シフトテンプレート";

// 行列インデックス
const SHIFT_ROW_DATE = 1; // 日程行
const SHIFT_COLUMN_DATE = 1; // 日程列

const SHIFT_COLUMN_START = 2; // メンバーリストの開始列

const SHIFT_ROW_MEMBERS = 1; // メンバーリスト行
const SHIFT_ROW_START_TIME = 2; // 開始時間行
const SHIFT_ROW_END_TIME = 3; // 終了時間行
const SHIFT_ROW_NOTE = 4; // 備考行
const SHIFT_ROW_WORK_START = 5; // 出勤時間
const SHIFT_ROW_WORK_END = 6; // 退勤時間
const SHIFT_ROW_WORKING = 7; // 勤務時間
const SHIFT_ROW_START = 9; // シフトの開始行
const SHIFT_ROW_END = 36; // シフトの終了行



// ===== 授業割テンプレートシート =====

// シート名
const LESSON_MON = "授業割(月)";
const LESSON_TUE = "授業割(火)";
const LESSON_WED = "授業割(水)";
const LESSON_THU = "授業割(木)";
const LESSON_FRI = "授業割(金)";



// ===== 環境設定 =====

// 年
const THIS_YEAR = 2025; // 今年（西暦）

// 時間帯リスト
const timeList = [
  "8:00", "8:30", "9:00", "9:30", "10:00", "10:30", "11:00", "11:30",
  "12:00", "12:30", "13:00", "13:30", "14:00", "14:30", "15:00", "15:30",
  "16:00", "16:30", "17:00", "17:30", "18:00", "18:30", "19:00", "19:30",
  "20:00", "20:30", "21:00", "21:30"
];

// デフォルト開閉室時間
const DEFAULT_OPEN_HOUR = 8;
const DEFAULT_OPEN_MINUTE = 0;
const DEFAULT_CLOSE_HOUR = 22;
const DEFAULT_CLOSE_MINUTE = 0;



// ===== 勤務不可背景色設定 ======
const UNAVAILABLE_COLOR = "#d3d3d3";



// ===== 勤務日数・労働時間計算関数 =====

const WORK_DATES_1 = `
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
`;

const WORK_TIMES_1 = `
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
`;

const WORK_DATES_2 = `
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
`;

const WORK_TIMES_2 = `
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
`;

const WORK_DATES_3 = `
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
`;

const WORK_TIMES_3 = `
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
`;

const WORK_DATES_4 = `
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
`;

const WORK_TIMES_4 = `
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
`;






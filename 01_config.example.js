/**
 * 環境設定テンプレートファイル
 *
 * 使い方:
 * 1. このファイルを config.env.js にコピー
 * 2. 各環境（本番・テスト）の実際のIDに置き換え
 * 3. config.env.js は Git に含めない（.gitignore で除外済み）
 *
 * 注意: このファイル(config.example.js)は Git に含まれるため、
 *      実際のIDは記載しないでください
 */

const CONFIG_TEMPLATE = {
  // === スプレッドシート・ファイルID ===

  // テンプレートファイルID（シフト希望表のベース）
  TEMPLATE_FILE_ID: "YOUR_TEMPLATE_FILE_ID_HERE",

  // シフト表共有ファイルID（完成したシフトを共有するファイル）
  SHARE_FILE_ID: "YOUR_SHARE_FILE_ID_HERE",

  // === フォルダID ===

  // 作成済みシフトPDFフォルダID
  SHIFT_PDF_FOLDER_ID: "YOUR_SHIFT_PDF_FOLDER_ID_HERE",

  // 作成済みシフトSSフォルダID（スプレッドシート保存用）
  SHIFT_SS_FOLDER_ID: "YOUR_SHIFT_SS_FOLDER_ID_HERE",

  // シフト希望表個別フォルダID（個人用フォームの保存先）
  PERSONAL_FORM_FOLDER_ID: "YOUR_PERSONAL_FORM_FOLDER_ID_HERE",

  // === 環境識別 ===

  // 環境名（"production" または "test"）
  ENVIRONMENT: "production", // または "test"

  // === その他の設定 ===

  // 年（西暦）
  THIS_YEAR: 2025,

  // デフォルト開閉室時間
  DEFAULT_OPEN_HOUR: 8,
  DEFAULT_OPEN_MINUTE: 0,
  DEFAULT_CLOSE_HOUR: 22,
  DEFAULT_CLOSE_MINUTE: 0,
};

// GAS環境用のエクスポート
if (typeof module !== "undefined" && module.exports) {
  module.exports = CONFIG_TEMPLATE;
} else {
  // GAS環境では global に設定
  this.CONFIG_TEMPLATE = CONFIG_TEMPLATE;
}

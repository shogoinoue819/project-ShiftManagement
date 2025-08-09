/**
 * 環境設定管理システム
 *
 * 設定の優先順位:
 * 1. config.env.js (最優先・Git除外)
 * 2. Script Properties (GAS環境)
 * 3. config.example.js (デフォルト・フォールバック)
 *
 * 使い方:
 * import { CONFIG } from "./config.js";
 * SpreadsheetApp.openById(CONFIG.TEMPLATE_FILE_ID);
 */

/**
 * 設定を読み込む関数
 * @returns {Object} マージされた設定オブジェクト
 */
function loadConfig() {
  let config = {};

  // 1. デフォルト設定（config.example.js）を読み込み
  try {
    if (typeof CONFIG_TEMPLATE !== "undefined") {
      config = { ...CONFIG_TEMPLATE };
    }
  } catch (e) {
    console.warn("config.example.js の読み込みに失敗:", e.message);
  }

  // 2. Script Properties から設定を読み込み（GAS環境のみ）
  try {
    if (typeof PropertiesService !== "undefined") {
      const scriptProperties = PropertiesService.getScriptProperties();
      const properties = scriptProperties.getProperties();

      // プロパティが存在する場合のみマージ
      Object.keys(properties).forEach((key) => {
        if (properties[key]) {
          // 数値型の変換
          if (
            key.includes("YEAR") ||
            key.includes("HOUR") ||
            key.includes("MINUTE")
          ) {
            config[key] = parseInt(properties[key], 10);
          } else {
            config[key] = properties[key];
          }
        }
      });
    }
  } catch (e) {
    console.warn("Script Properties の読み込みに失敗:", e.message);
  }

  // 3. 環境固有設定（config.env.js）を読み込み（最優先）
  try {
    if (typeof CONFIG_ENV !== "undefined") {
      config = { ...config, ...CONFIG_ENV };
    }
  } catch (e) {
    console.warn("config.env.js の読み込みに失敗:", e.message);
  }

  // 設定の妥当性チェック
  validateConfig(config);

  return config;
}

/**
 * 設定の妥当性をチェック
 * @param {Object} config 設定オブジェクト
 */
function validateConfig(config) {
  const requiredKeys = [
    "TEMPLATE_FILE_ID",
    "SHARE_FILE_ID",
    "SHIFT_PDF_FOLDER_ID",
    "SHIFT_SS_FOLDER_ID",
    "PERSONAL_FORM_FOLDER_ID",
  ];

  const missingKeys = requiredKeys.filter(
    (key) => !config[key] || config[key].includes("YOUR_") || config[key] === ""
  );

  if (missingKeys.length > 0) {
    const message = `設定が不完全です。以下の項目を確認してください: ${missingKeys.join(
      ", "
    )}`;
    console.error(message);

    // 開発環境でない場合はエラーを投げる
    if (config.ENVIRONMENT === "production") {
      throw new Error(message);
    }
  }
}

/**
 * 現在の設定情報を表示（デバッグ用）
 */
function showCurrentConfig() {
  const config = loadConfig();
  console.log("=== 現在の設定 ===");
  console.log(`環境: ${config.ENVIRONMENT || "未設定"}`);
  console.log(`テンプレートファイルID: ${config.TEMPLATE_FILE_ID}`);
  console.log(`共有ファイルID: ${config.SHARE_FILE_ID}`);
  console.log(`PDFフォルダID: ${config.SHIFT_PDF_FOLDER_ID}`);
  console.log(`SSフォルダID: ${config.SHIFT_SS_FOLDER_ID}`);
  console.log(`個人フォームフォルダID: ${config.PERSONAL_FORM_FOLDER_ID}`);
  console.log("==================");
}

// 設定オブジェクトをエクスポート
const CONFIG = loadConfig();

// GAS環境用のエクスポート
if (typeof module !== "undefined" && module.exports) {
  module.exports = { CONFIG, showCurrentConfig };
} else {
  // GAS環境では global に設定
  this.CONFIG = CONFIG;
  this.showCurrentConfig = showCurrentConfig;
}

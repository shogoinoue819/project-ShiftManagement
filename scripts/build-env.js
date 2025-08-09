#!/usr/bin/env node

const fs = require("fs");
const path = require("path");

/**
 * 環境別のconsts-env.jsを生成するスクリプト
 */

function buildEnvConstants(environment) {
  const configPath = path.join(__dirname, "../config/env-config.json");
  const outputPath = path.join(__dirname, "../consts-env.js");

  try {
    // 設定ファイルを読み込み
    if (!fs.existsSync(configPath)) {
      console.error("❌ 設定ファイルが見つかりません:", configPath);
      console.log(
        "💡 config/env-config.example.json をコピーして config/env-config.json を作成してください"
      );
      process.exit(1);
    }

    const config = JSON.parse(fs.readFileSync(configPath, "utf8"));

    if (!config[environment]) {
      console.error("❌ 無効な環境名:", environment);
      console.log("💡 使用可能な環境:", Object.keys(config).join(", "));
      process.exit(1);
    }

    const envConfig = config[environment];

    // consts-env.js の内容を生成
    const constContent = `/**
 * 環境依存定数ファイル
 * 
 * ⚠️ このファイルは自動生成されます。直接編集しないでください。
 * 環境: ${environment}
 * 生成日時: ${new Date().toLocaleString("ja-JP")}
 */

// ===== 環境依存の各種ファイルID =====

// テンプレートファイルID
const TEMPLATE_FILE_ID = "${envConfig.TEMPLATE_FILE_ID}";

// シフト表共有ファイルID
const SHARE_FILE_ID = "${envConfig.SHARE_FILE_ID}";

// 作成済みシフトPDFフォルダID
const SHIFT_PDF_FOLDER_ID = "${envConfig.SHIFT_PDF_FOLDER_ID}";

// 作成済みシフトSSフォルダID
const SHIFT_SS_FOLDER_ID = "${envConfig.SHIFT_SS_FOLDER_ID}";

// シフト希望表個別フォルダID
const PERSONAL_FORM_FOLDER_ID = "${envConfig.PERSONAL_FORM_FOLDER_ID}";

// 現在の環境
const CURRENT_ENVIRONMENT = "${environment}";
`;

    // ファイルを書き出し
    fs.writeFileSync(outputPath, constContent);

    console.log(`✅ ${environment} 環境用の consts-env.js を生成しました`);
    console.log(`📁 出力先: ${outputPath}`);

    // 設定内容を表示
    console.log("\n📋 設定内容:");
    Object.entries(envConfig).forEach(([key, value]) => {
      console.log(`  ${key}: ${value}`);
    });
  } catch (error) {
    console.error("❌ ビルドエラー:", error.message);
    process.exit(1);
  }
}

// コマンドライン引数から環境を取得
const environment = process.argv[2];

if (!environment) {
  console.error("❌ 環境名を指定してください");
  console.log("💡 使用方法: node scripts/build-env.js <test|production>");
  process.exit(1);
}

buildEnvConstants(environment);

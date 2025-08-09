# シフト管理システム (Google Apps Script)

Google Apps Script で作成したシフト管理システムです。本番環境とテスト環境の 2 環境で安全に運用できるよう設計されています。

## 🚀 初期セットアップ

### 1. 環境設定ファイルの作成

```bash
# config.example.js を config.env.js にコピー
cp config.example.js config.env.js
```

`config.env.js` を編集して、実際の環境に合わせて ID を設定してください：

```javascript
const CONFIG_ENV = {
  ENVIRONMENT: "test", // "production" または "test"
  TEMPLATE_FILE_ID: "実際のテンプレートファイルID",
  SHARE_FILE_ID: "実際の共有ファイルID",
  SHIFT_PDF_FOLDER_ID: "実際のPDFフォルダID",
  SHIFT_SS_FOLDER_ID: "実際のSSフォルダID",
  PERSONAL_FORM_FOLDER_ID: "実際の個人フォームフォルダID",
  // ... その他の設定
};
```

### 2. clasp 設定

`.clasp-test.json` を編集してテスト環境の scriptId を設定：

```json
{ "scriptId": "テスト環境のscriptId", "rootDir": "/path/to/project" }
```

### 3. 依存関係のインストール（任意）

npm scripts を使用する場合：

```bash
npm install
```

## 🔄 環境切り替えとデプロイ

### 基本的な方法（clasp 直接使用）

```bash
# 本番環境にpush
clasp push --project .clasp.json

# テスト環境にpush
clasp push --project .clasp-test.json
```

### npm scripts を使用（推奨）

```bash
# 本番環境にpush
npm run push:prod

# テスト環境にpush
npm run push:test

# 現在の設定を確認
npm run status

# セットアップ手順を表示
npm run setup
```

## 📋 現在の push 先を確認する方法

### 1. clasp 設定ファイルの確認

```bash
# 本番環境の設定
cat .clasp.json

# テスト環境の設定
cat .clasp-test.json
```

### 2. GAS エディタでの確認

GAS エディタのスクリプトエディタで、以下の関数を実行：

```javascript
showCurrentConfig();
```

実行結果のログで、現在の環境設定を確認できます。

### 3. 現在アクティブな.clasp.json の確認

```bash
# 現在の.clasp.jsonの内容を表示
cat .clasp.json
```

## ⚠️ 安全な運用のための注意事項

### push 前の確認事項

1. **環境の確認**: 必ず `--project` オプションで正しい設定ファイルを指定
2. **設定の確認**: GAS エディタで `showCurrentConfig()` を実行して環境設定を確認
3. **テスト実行**: テスト環境で動作確認してから本番環境に push

### 本番環境への誤 push を防ぐチェックリスト

- [ ] `--project .clasp-test.json` でテスト環境を指定しているか
- [ ] `config.env.js` で `ENVIRONMENT: "test"` に設定しているか
- [ ] テスト環境のスプレッドシート/フォルダ ID が設定されているか
- [ ] テスト実行で正常に動作することを確認したか

## 🛠️ トラブルシューティング

### 設定が読み込まれない場合

1. `config.env.js` が正しく作成されているか確認
2. ファイルの文法エラーがないか確認
3. GAS エディタで `showCurrentConfig()` を実行して設定値を確認

### push 時にエラーが発生する場合

1. `.clasp.json` または `.clasp-test.json` の scriptId が正しいか確認
2. GAS プロジェクトへのアクセス権限があるか確認
3. `clasp login` でログイン状態を確認

### 環境の設定値が反映されない場合

設定の優先順位を確認：

1. `config.env.js`（最優先）
2. Script Properties
3. `config.example.js`（フォールバック）

## 📁 ファイル構成

```
project-ShiftManagement/
├── config.js              # 設定管理システム
├── config.example.js      # 設定テンプレート（Git管理対象）
├── config.env.js          # 環境固有設定（Git除外）
├── .clasp.json            # 本番環境clasp設定（Git除外）
├── .clasp-test.json       # テスト環境clasp設定（Git除外）
├── package.json           # npm scripts定義
├── consts.js              # 定数定義（環境依存値は削除済み）
├── (その他の.jsファイル)   # 各機能の実装
└── README.md              # このファイル
```

## 🔒 セキュリティ

- 機密情報（ファイル ID、フォルダ ID）は `config.env.js` に分離
- `.gitignore` で機密ファイルを Git 管理から除外
- 環境ごとに異なる clasp 設定ファイルを使用
- 設定の妥当性チェック機能を内蔵

## 📝 開発時の注意

- ハードコードされた ID は使用せず、必ず `CONFIG.TEMPLATE_FILE_ID` などを使用
- 新しい環境依存の値が必要な場合は、`config.example.js` と `config.env.js` の両方に追加
- 機能追加時は適切な環境でテストしてから本番環境に push

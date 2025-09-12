# シフト管理システム (Google Apps Script)

Google Apps Script で作成したシフト管理システムです。本番環境とテスト環境の 2 環境で安全に運用できるよう設計されています。

## 🚀 初期セットアップ

### 1. 環境設定ファイルの作成

```bash
# 設定ファイルのテンプレートをコピー
cp config/env-config.example.json config/env-config.json
```

`config/env-config.json` を編集して、実際の環境に合わせて ID を設定してください：

```json
{
  "test": {
    "TEMPLATE_FILE_ID": "テスト環境のテンプレートファイルID",
    "SHARE_FILE_ID": "テスト環境の共有ファイルID",
    "SHIFT_PDF_FOLDER_ID": "テスト環境のPDFフォルダID",
    "SHIFT_SS_FOLDER_ID": "テスト環境のSSフォルダID",
    "PERSONAL_FORM_FOLDER_ID": "テスト環境の個人フォームフォルダID"
  },
  "production": {
    "TEMPLATE_FILE_ID": "本番環境のテンプレートファイルID",
    "SHARE_FILE_ID": "本番環境の共有ファイルID",
    "SHIFT_PDF_FOLDER_ID": "本番環境のPDFフォルダID",
    "SHIFT_SS_FOLDER_ID": "本番環境のSSフォルダID",
    "PERSONAL_FORM_FOLDER_ID": "本番環境の個人フォームフォルダID"
  }
}
```

### 2. clasp 設定

`.clasp.json` と `.clasp-test.json` を編集して各環境の scriptId を設定：

```json
// .clasp.json (本番環境)
{
  "scriptId": "本番環境のscriptId",
  "rootDir": "/path/to/project"
}

// .clasp-test.json (テスト環境)
{
  "scriptId": "テスト環境のscriptId",
  "rootDir": "/path/to/project"
}
```

### 3. 依存関係のインストール（任意）

npm scripts を使用する場合：

```bash
npm install
```

## 🔄 環境切り替えとデプロイ

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

# 現在の環境設定を表示
npm run config:show
```

### 基本的な方法（clasp 直接使用）

```bash
# 本番環境用にビルドしてpush
npm run build:prod
clasp push --project .clasp.json

# テスト環境用にビルドしてpush
npm run build:test
clasp push --project .clasp-test.json
```

## 📋 現在の push 先を確認する方法

### 1. clasp 設定ファイルの確認

```bash
# 本番環境の設定
cat .clasp.json

# テスト環境の設定
cat .clasp-test.json
```

### 2. 現在の環境設定を確認

```bash
# 現在の環境設定を表示
npm run config:show
```

### 3. 設定ファイルの確認

```bash
# 設定ファイルの内容を確認
cat config/env-config.json
```

## ⚠️ 安全な運用のための注意事項

### push 前の確認事項

1. **環境の確認**: 必ず正しい npm script を使用
2. **設定の確認**: `npm run config:show` で現在の環境設定を確認
3. **テスト実行**: テスト環境で動作確認してから本番環境に push

### 本番環境への誤 push を防ぐチェックリスト

- [ ] `npm run push:test` でテスト環境を指定しているか
- [ ] `config/env-config.json` で正しい ID が設定されているか
- [ ] テスト環境のスプレッドシート/フォルダ ID が設定されているか
- [ ] テスト実行で正常に動作することを確認したか

## 🛠️ トラブルシューティング

### 設定が読み込まれない場合

1. `config/env-config.json` が正しく作成されているか確認
2. ファイルの文法エラーがないか確認
3. `npm run config:show` で現在の設定値を確認

### push 時にエラーが発生する場合

1. `.clasp.json` または `.clasp-test.json` の scriptId が正しいか確認
2. GAS プロジェクトへのアクセス権限があるか確認
3. `clasp login` でログイン状態を確認

### 環境の設定値が反映されない場合

1. `npm run build:test` または `npm run build:prod` を実行
2. `02_consts-env.js` が正しく生成されているか確認

## 📁 ファイル構成

```
project-ShiftManagement/
├── 01_consts.js              # 共通定数定義
├── 02_consts-env.js          # 環境依存定数（自動生成）
├── 03_utils.js               # ユーティリティ関数
├── 04_createMenu.js          # 管理メニュー作成
├── 10_changeToNextTerm.js    # 次回用シフト準備
├── 11_updateForms.js         # シフト希望表配布
├── 12_updateSheets.js        # 各日程シート作成
├── 13_checkMethods.js        # 一括チェック機能
├── 14_reflectShiftForms.js   # シフト希望反映
├── 15_reflectLessonTemplate.js # 授業割テンプレート反映
├── 16_shareShifts.js         # シフト共有機能
├── 20_createNewMember.js     # 新規メンバー追加
├── 21_deleteSelectedMember.js # メンバー削除
├── 22_addNewMember.js        # 臨時メンバー追加
├── 30_sendReminderMail.js    # リマインダーメール送信
├── 90_fillForms.js           # フォーム埋め込み
├── 91_reReflectTemplateSheet.js # テンプレートシート再反映
├── 92_reReflectTemplateInfoSheet.js # テンプレート情報シート再反映
├── config/
│   ├── env-config.example.json # 設定テンプレート（Git管理対象）
│   └── env-config.json        # 環境固有設定（Git除外）
├── scripts/
│   └── build-env.js           # 環境別ビルドスクリプト
├── .clasp.json                # 本番環境clasp設定（Git除外）
├── .clasp-test.json           # テスト環境clasp設定（Git除外）
├── package.json               # npm scripts定義
└── README.md                  # このファイル
```

## 🔒 セキュリティ

- 機密情報（ファイル ID、フォルダ ID）は `config/env-config.json` に分離
- `.gitignore` で機密ファイルを Git 管理から除外
- 環境ごとに異なる clasp 設定ファイルを使用
- 環境別の自動ビルドシステム

## 📝 開発時の注意

- ハードコードされた ID は使用せず、必ず `TEMPLATE_FILE_ID` などを使用
- 新しい環境依存の値が必要な場合は、`config/env-config.example.json` と `config/env-config.json` の両方に追加
- 機能追加時は適切な環境でテストしてから本番環境に push
- `02_consts-env.js` は自動生成ファイルなので直接編集しない

## 🎯 主な機能

### シフト管理機能

- **シフト希望表の自動配布** - メンバー個別のシフト希望表を自動生成・配布
- **シフト希望の自動反映** - 提出されたシフト希望をシフト作成シートに自動反映
- **一括チェック機能** - 提出状況の一括確認と管理
- **シフト共有機能** - 完成したシフト表の自動共有と PDF 生成

### メンバー管理

- **新規メンバー追加** - 簡単な操作でメンバーを追加
- **メンバー削除** - 不要なメンバーの安全な削除
- **臨時メンバー追加** - シフト表末尾への一時的な追加

### 通知機能

- **リマインダーメール** - シフト提出期限の自動通知

### 授業割管理

- **授業割テンプレート反映** - 授業スケジュールの自動反映

## ⚡ パフォーマンス最適化

- **データキャッシュシステム** - 高速なデータ処理
- **一括更新処理** - 効率的なシート操作
- **Google Apps Script 6 分制限対応** - 実行時間の最適化
- **環境別自動ビルド** - 効率的なデプロイメント

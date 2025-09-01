// 未提出者にリマインダーメールを送信

// 件名・本文（ファイル先頭に定義）
const REMINDER_SUBJECT = "シフト希望提出のお願い";
const REMINDER_BASE_BODY =
  "シフト希望表が未提出になっております。\n早急に提出をお願いします。";

function sendReminderMail() {
  // SSをまとめて取得
  const manageSheet = getManageSheet();
  const ui = getUI();

  // 最終行取得
  const lastRow = getLastRowInColumn(
    manageSheet,
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
  );
  // 送信者リスト
  const sentNames = [];

  for (
    let row = SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW;
    row <= lastRow;
    row++
  ) {
    // 氏名、提出ステータス、メアド、URLを取得
    const name = manageSheet
      .getRange(row, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.NAME_COL)
      .getValue();
    const status = manageSheet
      .getRange(row, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.SUBMIT_COL)
      .getValue();
    const email = manageSheet
      .getRange(row, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.EMAIL_COL)
      .getValue();
    const urlFormula = manageSheet
      .getRange(row, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.URL_COL)
      .getFormula();
    const directUrl = extractDirectUrlFromFormula(urlFormula);
    // 未提出かつメアドとURLがあれば、
    if (isUnsubmitted(status) && isDeliverable(email, directUrl)) {
      try {
        // メール本文を生成（URL付き）
        const body = buildReminderBody(directUrl);
        // メール送信
        GmailApp.sendEmail(email, REMINDER_SUBJECT, body);
        Logger.log(`✅ メール送信: ${email}`);
        // 送信者リストに追加
        sentNames.push(name);
      } catch (e) {
        Logger.log(`❌ メール送信失敗: ${email} (${e})`);
      }
    }
  }

  // アラートメッセージ整形
  const nameList =
    sentNames.length > 0
      ? "\n\n【送信対象】\n- " + sentNames.join("\n- ")
      : "\n\n送信対象者はいませんでした。";

  ui.alert(
    `✅ メール送信完了\n${sentNames.length}件の未提出者にメールを送信しました。${nameList}`
  );
}

// ===== ヘルパー関数（メイン関数の下に配置） =====
function isUnsubmitted(status) {
  return status === STATUS_STRINGS.SUBMIT.FALSE;
}

function isDeliverable(email, url) {
  if (!email || !url) return false;
  const e = String(email).trim();
  const u = String(url).trim();
  return e.includes("@") && u.length > 0;
}

function buildReminderBody(url) {
  const link = String(url).trim();
  return `${REMINDER_BASE_BODY}\n\n以下のリンクから提出してください：\n${link}`;
}

function extractDirectUrlFromFormula(formula) {
  if (!formula) return "";
  const str = String(formula).trim();
  // =HYPERLINK("https://...", "text")
  const m1 = str.match(/^=HYPERLINK\(\s*"([^"]+)"/i);
  if (m1 && m1[1]) return m1[1];
  // =HYPERLINK('https://...', 'text')
  const m2 = str.match(/^=HYPERLINK\(\s*'([^']+)'/i);
  if (m2 && m2[1]) return m2[1];
  // 一般的なURLパターンを抽出
  const m3 = str.match(/https?:\/\/[^\s",)]+/i);
  if (m3 && m3[0]) return m3[0];
  // 見つからない場合は空
  return "";
}

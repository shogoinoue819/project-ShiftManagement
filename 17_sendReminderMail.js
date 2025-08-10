// 未提出者にリマインダーメールを送信
function sendReminderMail() {
  // SSをまとめて取得
  const [ss, manageSheet, templateSheet, allSheets, ui] = getCommonSheets();

  // タイトルと本文
  const SUBJECT = "シフト希望提出のお願い";
  const BASE_BODY = `
    シフト希望表が未提出になっております。
    早急に提出をお願いします。
  `;

  // 最終行取得
  const lastRow = getLastRowInCol(
    manageSheet,
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
  );
  // 送信人数
  let sentCount = 0;
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
    const url = manageSheet
      .getRange(row, SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.URL_COL)
      .getFormula();
    // 未提出かつメアドとURLがあれば、
    if (status === SUBMIT_FALSE && email && url) {
      try {
        // メール本文を生成（URL付き）
        const body = `${BASE_BODY}\n\n以下のリンクから提出してください：\n${url}`;
        // メール送信
        GmailApp.sendEmail(email, SUBJECT, body);
        Logger.log(`✅ メール送信: ${email}`);
        // 送信者リストに追加
        sentNames.push(name);
        sentCount++;
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
    `✅ メール送信完了\n${sentCount}件の未提出者にメールを送信しました。${nameList}`
  );
}

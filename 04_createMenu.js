// 管理メニューを作成
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("管理メニュー")
    .addItem("①次回用シフト希望表準備", "changeToNextTerm")
    .addItem("②シフト希望表配布", "updateForms")
    .addItem("③各日程シート作成", "updateSheets")
    .addItem("④一括チェック", "checkAllSubmittedMembers")
    .addItem("⑤シフト希望反映", "reflectShiftForms")
    .addItem("⑥授業割テンプレ反映", "reflectLessonTemplate")
    .addItem("⑦完成済みシフト一括共有", "shareShiftsAll")
    .addItem("⑧作業中シート限定更新", "shareOnlyOneShift")
    .addSeparator()
    .addItem("👥 新規メンバー追加", "createNewMember")
    .addItem("🗑️ メンバー削除", "deleteSelectedMember")
    .addItem("➕ シフト表末尾に追加(臨時)", "addNewMember")
    .addSeparator()
    .addItem("📧 リマインダーメール送信", "sendReminderMail")
    .addToUi();
}

// 管理メニューを作成
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("管理メニュー")
    .addItem("①シフト日程を反映", "reflectDateList")
    .addItem("②シフト希望表をアップデート", "updateForms")
    .addItem("③シフト作成シートをアップデート", "updateSheets")
    .addItem("④提出済みメンバーを一括チェック", "checkAllSubmittedMembers")
    .addItem("⑤シフト希望を反映", "reflectShiftForms")
    .addItem("⑥授業割テンプレを反映", "reflectLessonTemplate")
    .addItem("⑦完成したシフトを共有(一括更新)", "shareShiftsAll")
    .addItem("⑧開いている日程のシフトを更新(限定更新)", "shareOnlyOneShift")
    .addItem("新規メンバーを追加", "createNewMember")
    .addItem("メンバーを削除", "deleteSelectedMember")
    .addItem("未提出者にリマインダーメールを送信", "sendReminderMail")
    .addItem("<新規>シフト表末尾にメンバーを追加", "addNewMember")
    .addToUi();
}

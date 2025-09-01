// チェックボックスを押すたびにロック関数を動作させる
function onEdit(e) {
  // チェックボックスが押された行列を取得
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();

  // チェック欄でチェックされた場合
  if (
    col === SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.CHECK_COL &&
    row >= SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW
  ) {
    if (e.value === "TRUE") {
      lockSelectedMember(row);
    } else if (e.value === "FALSE") {
      unlockSelectedMember(row);
    }
  }
  Logger.log(`onEdit 発火: row=${row}, col=${col}, value=${e.value}`);
}

// シート保護の共通処理
function protectMemberSheets(fileId, memberName, isLock) {
  try {
    Logger.log(`🔒 ${isLock ? "ロック" : "アンロック"}処理開始: ${memberName}`);

    // ファイルIDから提出用SSを取得
    const targetFile = SpreadsheetApp.openById(fileId);

    if (isLock) {
      // ロック処理
      const formSuccess = protectSheetByName(
        targetFile,
        SHEET_NAMES.SHIFT_FORM,
        "チェックによるロック",
        memberName
      );

      const infoSuccess = protectSheetByName(
        targetFile,
        SHEET_NAMES.SHIFT_FORM_INFO,
        "チェックによるロック（今後の勤務希望）",
        memberName
      );

      if (formSuccess && infoSuccess) {
        Logger.log(`🔒 ${memberName} をロックしました`);
      } else {
        Logger.log(
          `⚠️ ${memberName} のロックが部分的に失敗しました (form: ${formSuccess}, info: ${infoSuccess})`
        );
        return false;
      }
    } else {
      // アンロック処理
      const formSuccess = unprotectSheetByName(
        targetFile,
        SHEET_NAMES.SHIFT_FORM,
        memberName
      );

      const infoSuccess = unprotectSheetByName(
        targetFile,
        SHEET_NAMES.SHIFT_FORM_INFO,
        memberName
      );

      if (formSuccess && infoSuccess) {
        Logger.log(`🔓 ${memberName} のロックを解除しました`);
      } else {
        Logger.log(
          `⚠️ ${memberName} のロック解除が部分的に失敗しました (form: ${formSuccess}, info: ${infoSuccess})`
        );
        return false;
      }
    }

    return true;
  } catch (e) {
    Logger.log(
      `❌ ${isLock ? "ロック" : "アンロック"}失敗: ${memberName} - ${e}`
    );
    return false;
  }
}

// 選択されたメンバーをロック
function lockSelectedMember(row) {
  try {
    const manageSheet = getManageSheet();
    const memberInfo = getMemberInfo(row, manageSheet);
    if (!memberInfo) {
      Logger.log(`⚠️ メンバー情報の取得に失敗: 行${row}`);
      return false;
    }

    Logger.log(`🔒 ロック処理開始: ${memberInfo.name}`);

    const success = protectMemberSheets(
      memberInfo.fileId,
      memberInfo.name,
      true
    );

    if (!success) {
      Logger.log(`⚠️ メンバーロックに失敗: ${memberInfo.name}`);
    } else {
      Logger.log(`✅ メンバーロック成功: ${memberInfo.name}`);
    }

    return success;
  } catch (e) {
    Logger.log(`❌ ロック処理でエラーが発生: 行${row} - ${e}`);
    return false;
  }
}

// 選択されたメンバーのロックを解除
function unlockSelectedMember(row) {
  try {
    const manageSheet = getManageSheet();
    const memberInfo = getMemberInfo(row, manageSheet);
    if (!memberInfo) {
      Logger.log(`⚠️ メンバー情報の取得に失敗: 行${row}`);
      return false;
    }

    Logger.log(`🔓 ロック解除処理開始: ${memberInfo.name}`);

    const success = protectMemberSheets(
      memberInfo.fileId,
      memberInfo.name,
      false
    );

    if (success) {
      Logger.log(`✅ メンバーロック解除成功: ${memberInfo.name}`);
      const ui = getUI();
      ui.alert(`🔓 ${memberInfo.name}さんのロックを解除しました`);
    } else {
      Logger.log(`⚠️ メンバーロック解除に失敗: ${memberInfo.name}`);
    }

    return success;
  } catch (e) {
    Logger.log(`❌ ロック解除処理でエラーが発生: 行${row} - ${e}`);
    return false;
  }
}

// 提出済みのメンバーを全てチェックする
function checkAllSubmittedMembers() {
  // SSをまとめて取得
  const ss = getSpreadsheet();
  const manageSheet = getManageSheet();
  const ui = getUI();

  // 最終行を取得
  const lastRow = getLastRowInColumn(
    manageSheet,
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_COL
  );

  // データが存在しない場合
  if (lastRow < SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW) {
    ui.alert(`❌ メンバーデータが存在しません`);
    return;
  }

  // 必要な列のみを取得（パフォーマンス改善）
  const startRow = SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW;
  const rowCount = lastRow - startRow + 1;

  // 提出ステータスとチェック状態の列のみを取得
  const submitCol = SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.SUBMIT_COL;
  const checkCol = SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.CHECK_COL;

  const data = manageSheet
    .getRange(startRow, submitCol, rowCount, 2) // submit列とcheck列のみ
    .getValues();

  // 人数カウンター
  let count = 0;
  const rowsToCheck = [];

  // データの各メンバーにおいて、提出済みかつ未チェックの行を特定
  data.forEach((row, i) => {
    const submitStatus = row[0]; // submit列
    const isChecked = row[1]; // check列

    // 提出済みかつチェックされていなければ、対象行として記録
    if (submitStatus === STATUS_STRINGS.SUBMIT.TRUE && isChecked !== true) {
      rowsToCheck.push(startRow + i);
    }
  });

  // 対象行がない場合
  if (rowsToCheck.length === 0) {
    ui.alert(`❌ 新たにチェックできるメンバーはいません`);
    return;
  }

  // 各対象メンバーをロック（先にロック処理を実行）
  const successfulRows = [];
  const failedRows = [];

  Logger.log(`🔒 対象メンバー数: ${rowsToCheck.length}人`);

  rowsToCheck.forEach((rowIndex) => {
    try {
      const success = lockSelectedMember(rowIndex);
      if (success) {
        successfulRows.push(rowIndex);
      } else {
        failedRows.push(rowIndex);
      }
    } catch (e) {
      Logger.log(`❌ 行${rowIndex}のロック処理でエラー: ${e}`);
      failedRows.push(rowIndex);
    }
  });

  // ロックに成功した行がない場合
  if (successfulRows.length === 0) {
    const failedNames = failedRows
      .map((row) => {
        try {
          const memberInfo = getMemberInfo(row, manageSheet);
          return memberInfo ? memberInfo.name : `行${row}`;
        } catch (e) {
          return `行${row}`;
        }
      })
      .join(", ");

    ui.alert(
      `❌ ロック処理に失敗したため、チェックを設定できませんでした\n\n失敗したメンバー: ${failedNames}`
    );
    return;
  }

  // 一括でチェックを設定（パフォーマンス改善）
  const checkRange = manageSheet.getRange(
    SHIFT_MANAGEMENT_SHEET.MEMBER_LIST.START_ROW,
    checkCol,
    rowCount,
    1
  );
  const checkValues = checkRange.getValues();

  // ロックに成功した行のみチェックを設定
  successfulRows.forEach((rowIndex) => {
    const relativeRow = rowIndex - startRow;
    checkValues[relativeRow][0] = true;
  });

  // 一括更新
  checkRange.setValues(checkValues);

  // 結果の表示
  if (successfulRows.length === rowsToCheck.length) {
    ui.alert(
      `✅ 提出済みのメンバー${successfulRows.length}人をチェックしました`
    );
  } else {
    const failedNames = failedRows
      .map((row) => {
        try {
          const memberInfo = getMemberInfo(row, manageSheet);
          return memberInfo ? memberInfo.name : `行${row}`;
        } catch (e) {
          return `行${row}`;
        }
      })
      .join(", ");

    ui.alert(
      `⚠️ 提出済みのメンバー${successfulRows.length}人をチェックしました（${
        rowsToCheck.length - successfulRows.length
      }人はロック処理に失敗）\n\n失敗したメンバー: ${failedNames}`
    );
  }
}

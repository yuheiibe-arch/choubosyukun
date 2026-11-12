/**
 * ============================================================================
 * 05c_D6_And_Notify_Exec.gs（Slack＆D6転記 3分割の3）
 * ・確認済行の一括実行（executeApprovedRows：日付昇順Slack同日集約更新＋D6メモ転記＋履歴記録＋行削除）
 * ・取込完了時のChatwork／Slack通知（コピペ自動リンク化防止対応済）
 * ============================================================================
 */

function buildChatworkRoomMessagesUrl_(roomId) {
  const host = ["api", "chatwork", "com"].join(".");
  return "https://" + host + "/v2/rooms/" + encodeURIComponent(String(roomId || "").trim()) + "/messages";
}

function executeApprovedRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const activeName = activeSheet ? activeSheet.getName() : "";
  const offlineSheetName = (typeof SHEET_NAMES !== "undefined" && SHEET_NAMES.D7_OFFLINE_REVIEW)
    ? SHEET_NAMES.D7_OFFLINE_REVIEW
    : "オフライン修正シート";

  const revSheet = (activeName === offlineSheetName)
    ? activeSheet
    : ss.getSheetByName(SHEET_NAMES.D7_REVIEW);

  const histSheet = ss.getSheetByName(SHEET_NAMES.D7_HISTORY);
  const logSheet = ss.getSheetByName(SHEET_NAMES.D7_LOG);
  if (!revSheet) return;
  const lastRow = revSheet.getLastRow();
  if (lastRow < 2) return;

  loadStaffMasterFromD7_(ss);

  const range = revSheet.getRange(2, 1, lastRow - 1, 24);
  const rawValues = range.getValues();
  const dispValues = range.getDisplayValues();

  const targetIndices = [];
  for (let i = 0; i < dispValues.length; i++) {
    const isChecked = (rawValues[i][0] === true || dispValues[i][0] === "TRUE");
    const statusStr = String(dispValues[i][1] || "").trim();
    if (isChecked && (statusStr === "未レビュー" || statusStr === "確認済み" || statusStr === "エラー" || statusStr === "")) {
      targetIndices.push(i);
    }
  }

  const ui = SpreadsheetApp.getUi();
  if (targetIndices.length === 0) {
    ui.alert(`シート「${revSheet.getName()}」に「確認済」にチェックが入っている対象行がありません。`);
    return;
  }

  const confirmResponse = ui.alert(
    `Slack投稿・転記の確認（対象シート：${revSheet.getName()}）`,
    "医師への送信なども済みですか？\n（同じ日付の投稿はSlackスレッド内の1つのメッセージに自動でまとめ・編集追記されます）\n投稿後、対象行はこのシートから削除されます。",
    ui.ButtonSet.OK_CANCEL
  );
  if (confirmResponse !== ui.Button.OK) {
    return;
  }

  const unfilledSs = SpreadsheetApp.openById(EXTERNAL_IDS.D6_D3_UNFILLED_SS);
  const d6Sheet = unfilledSs.getSheetByName(SHEET_NAMES.D6_APPROACH);
  const d6Headers = d6Sheet.getRange(1, 1, 1, d6Sheet.getLastColumn()).getDisplayValues()[0];
  const colDate = d6Headers.indexOf("日付") + 1;
  const colClinic = d6Headers.indexOf("拠点名") + 1;
  const colTime = d6Headers.indexOf("不在時間") + 1;
  const colMemo = d6Headers.indexOf("作業メモ") + 1;

  const d6LastRow = d6Sheet.getLastRow();
  const d6Keys = d6LastRow >= 3
    ? d6Sheet.getRange(1, 1, d6LastRow, Math.max(colDate, colClinic, colTime)).getDisplayValues()
    : [];
  const nowStr = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");
  const userEmail = Session.getActiveUser().getEmail();

  const rawToken = PropertiesService.getScriptProperties().getProperty("SLACK_BOT_TOKEN");
  const token = rawToken ? rawToken.trim() : "";
  const parentThreadMap = fetchSlackParentThreadMap_(token, SLACK_CHANNEL_ID);
  const threadRepliesCache = new Map();

  const dateSlackTasks = new Map();
  const rowExecutionPlans = [];

  for (const idx of targetIndices) {
    const sheetRowNum = idx + 2;
    const row = dispValues[idx];

    const dateList = String(row[3] || "").split(/\r?\n/).map(function(s) { return normalizeDateStr_(s); }).filter(Boolean);
    const clinicList = String(row[4] || "").split(/\r?\n/).map(function(s) { return s.trim(); }).filter(Boolean);
    const timeList = String(row[5] || "").split(/\r?\n/).map(function(s) { return normalizeTimeRange_(s); }).filter(Boolean);
    const targetDept = String(row[6] || "").trim();
    const approachType = String(row[7] || "").trim();
    const doctorsText = String(row[8] || "").trim();
    const rowIsSpecial = String(row[11] || "").includes("特別時給");
    const staffName = extractCleanStaffName_(String(row[15] || "").trim());
    const memoToAppend = String(row[17] || "").trim();
    const fallbackThreadTitle = String(row[18] || "").trim();
    const rawSlackBlocks = String(row[20] || "").split(/(?=`\d{1,2}月\s*\d{1,2}日)/).map(function(s) { return s.trim(); }).filter(Boolean);

    const slotCount = Math.max(dateList.length, clinicList.length, 1);
    const slotOps = [];

    for (let sIdx = 0; sIdx < slotCount; sIdx++) {
      const targetDate = dateList[sIdx] || dateList[0] || "";
      const targetClinic = clinicList[sIdx] || clinicList[0] || "";
      const targetTime = timeList[sIdx] || timeList[0] || "";
      const threadTitle = buildSlackThreadTitle_(targetDate) || fallbackThreadTitle;

      const blockForSlot = rawSlackBlocks[sIdx] || buildSingleSlotSlackBlockText_(targetDate, targetClinic, targetTime, rowIsSpecial, memoToAppend);
      const parsedBlock = parseSlackDateMessageStruct_(blockForSlot);

      if (targetDate) {
        if (!dateSlackTasks.has(targetDate)) {
          dateSlackTasks.set(targetDate, {
            threadTitle: threadTitle,
            struct: { sections: new Map(), resultLines: [] }
          });
        }
        const task = dateSlackTasks.get(targetDate);
        task.struct = mergeSlackDateMessageStructs_(task.struct, parsedBlock);
      }

      slotOps.push({
        targetDate: targetDate,
        targetClinic: targetClinic,
        targetTime: targetTime,
        targetDept: targetDept,
        approachType: approachType,
        doctorsText: doctorsText,
        staffName: staffName,
        memoToAppend: memoToAppend
      });
    }

    rowExecutionPlans.push({
      sheetRowNum: sheetRowNum,
      slotOps: slotOps
    });
  }

  // Step 1: 日付昇順（11/15 ➔ 11/16 ...）で1日1回ずつSlackへ投稿または既存日付メッセージを上書き編集
  const postIdByDate = new Map();
  const errorByDate = new Map();
  const sortedDates = Array.from(dateSlackTasks.keys()).sort();

  for (const targetDate of sortedDates) {
    const task = dateSlackTasks.get(targetDate);
    try {
      const postId = upsertDateStructToSlackThread_(
        targetDate,
        task.threadTitle,
        task.struct,
        parentThreadMap,
        threadRepliesCache
      );
      postIdByDate.set(targetDate, postId);
      Utilities.sleep(350);
    } catch (e) {
      errorByDate.set(targetDate, e.message);
    }
  }

  // Step 2: Slack反映に成功した行のD6作業メモ転記 ＆ 履歴・ログ記録
  const completedSheetRows = [];
  let d6Modified = false;

  for (const plan of rowExecutionPlans) {
    const failedSlot = plan.slotOps.find(function(op) { return errorByDate.has(op.targetDate); });
    if (failedSlot) {
      revSheet.getRange(plan.sheetRowNum, 2).setValue("エラー");
      revSheet.getRange(plan.sheetRowNum, 24).setValue(errorByDate.get(failedSlot.targetDate));
      continue;
    }

    try {
      for (const op of plan.slotOps) {
        const postId = postIdByDate.get(op.targetDate) || "";
        const targetNormClinic = normalizeClinicForMatch_(op.targetClinic);
        const exactTimeD6Rows = [];
        const sameClinicDateD6Rows = [];

        for (let r = 2; r < d6Keys.length; r++) {
          const rDate = normalizeDateStr_(d6Keys[r][colDate - 1]);
          if (rDate !== op.targetDate) continue;

          const rClinic = String(d6Keys[r][colClinic - 1] || "").trim();
          const rNormClinic = normalizeClinicForMatch_(rClinic);
          const rTime = normalizeTimeRange_(d6Keys[r][colTime - 1]);

          const isKitaOrKame = (targetNormClinic === "北葛西" || targetNormClinic === "亀有");
          const isClinicMatch = isKitaOrKame
            ? (rClinic === op.targetClinic || (rNormClinic === targetNormClinic && (!op.targetDept || rClinic.includes(op.targetDept))))
            : (rClinic === op.targetClinic || rNormClinic === targetNormClinic);

          if (isClinicMatch) {
            sameClinicDateD6Rows.push(r + 1);
            if (!op.targetTime || !rTime || rTime === op.targetTime || isTimeOverlapping_(rTime, op.targetTime)) {
              exactTimeD6Rows.push(r + 1);
            }
          }
        }

        const matchedD6Rows = exactTimeD6Rows.length > 0 ? exactTimeD6Rows : sameClinicDateD6Rows;

        if (matchedD6Rows.length > 0 && op.memoToAppend) {
          for (const matchedD6Row of matchedD6Rows) {
            const memoCell = d6Sheet.getRange(matchedD6Row, colMemo);
            const currentMemo = String(memoCell.getDisplayValue() || "").trim();
            const cleanTargetMemo = op.memoToAppend.replace(/^・/, "").trim();
            const lines = currentMemo ? currentMemo.split(/\r?\n/).map(function(l) { return l.trim(); }).filter(Boolean) : [];
            const alreadyExists = lines.some(function(l) { return l.replace(/^・/, "").trim() === cleanTargetMemo || l.includes(cleanTargetMemo); });
            if (!alreadyExists) {
              lines.push(op.memoToAppend);
              memoCell.setValue(lines.join("\n"));
              d6Modified = true;
            }
          }
        }

        if (histSheet && op.doctorsText && op.doctorsText !== "該当候補なし") {
          const docList = op.doctorsText.split(/\r?\n/).map(function(s) { return s.trim(); }).filter(Boolean);
          for (const docName of docList) {
            appendRowAndTrimBlank_(
              histSheet,
              [nowStr, op.targetDate, op.targetClinic, op.targetTime, op.targetDept, op.approachType, docName, op.staffName || userEmail, postId],
              9
            );
          }
        }

        if (logSheet) {
          const actionLabel = matchedD6Rows.length > 0
            ? `D6転記・Slack同日集約更新(${revSheet.getName()})`
            : `Slack同日集約更新(D6該当枠なし/${revSheet.getName()})`;
          appendRowAndTrimBlank_(
            logSheet,
            [nowStr, actionLabel, postId, `${op.targetDate} ${op.targetClinic}${op.targetTime}`, 1, `${op.memoToAppend} (${op.staffName || userEmail})`],
            6
          );
        }
      }

      completedSheetRows.push(plan.sheetRowNum);
    } catch (e) {
      revSheet.getRange(plan.sheetRowNum, 2).setValue("エラー");
      revSheet.getRange(plan.sheetRowNum, 24).setValue(e.message);
    }
  }

  if (d6Modified) {
    SpreadsheetApp.flush();
  }

  if (completedSheetRows.length > 0) {
    completedSheetRows.sort(function(a, b) { return b - a; });
    for (const rowNum of completedSheetRows) {
      if (revSheet.getMaxRows() > 2) {
        revSheet.deleteRow(rowNum);
      } else {
        revSheet.getRange(2, 1, 1, revSheet.getLastColumn())
          .clearContent()
          .clearDataValidations()
          .setBackground("#ffffff")
          .setFontColor("#000000");
      }
    }
  }

  trimBlankRowsAndCols_(revSheet, 24);
  applyLeftMiddleToSheet_(revSheet);
  if (histSheet) applyLeftMiddleToSheet_(histSheet);
  if (logSheet) applyLeftMiddleToSheet_(logSheet);
}

function sendImportCompletionNotifications_(sheetNameOpt) {
  const todayStr = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");
  const targetLabel = sheetNameOpt ? `（${sheetNameOpt}）` : "";
  const bodyText = `${todayStr} 超募集くんデータ作成済み${targetLabel}\nURL:${D7_REVIEW_SHEET_URL}`;
  const props = PropertiesService.getScriptProperties();

  const cwToken = props.getProperty("CHATWORK_API_KEY");
  if (cwToken && CHATWORK_NOTIFY_ROOM_ID) {
    try {
      const cwMessage = `[toall]\n[info]${bodyText}[/info]`;
      const cwUrl = buildChatworkRoomMessagesUrl_(CHATWORK_NOTIFY_ROOM_ID);
      const cwRes = UrlFetchApp.fetch(cwUrl, {
        method: "post",
        headers: { "X-ChatWorkToken": cwToken.trim() },
        payload: { body: cwMessage },
        muteHttpExceptions: true
      });
      console.log(`Chatwork通知送信完了 (Status: ${cwRes.getResponseCode()})`);
    } catch (e) {
      console.error("Chatwork通知エラー: " + e.message);
    }
  }

  const slackToken = props.getProperty("SLACK_BOT_TOKEN");
  if (slackToken && SLACK_NOTIFY_CHANNEL_ID) {
    try {
      const slackMessage = `<!here>\n\`\`\`\n${bodyText}\n\`\`\``;
      const slackUrl = buildSlackApiUrl_("chat.postMessage");
      const slackRes = UrlFetchApp.fetch(slackUrl, {
        method: "post",
        contentType: "application/json; charset=utf-8",
        headers: { Authorization: "Bearer " + slackToken.trim() },
        payload: JSON.stringify({
          channel: SLACK_NOTIFY_CHANNEL_ID,
          text: slackMessage
        }),
        muteHttpExceptions: true
      });
      const sData = parseSlackApiJsonSafely_(slackRes, "Slack完了通知");
      if (!sData.ok) {
        console.error("Slack通知APIエラー: " + sData.error);
      } else {
        console.log("Slack通知送信完了");
      }
    } catch (e) {
      console.error("Slack通知エラー: " + e.message);
    }
  }
}

function inspectCurrentOutput() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const revSheet = ss.getSheetByName(SHEET_NAMES.D7_REVIEW);
  const offlineSheet = ss.getSheetByName(SHEET_NAMES.D7_OFFLINE_REVIEW || "オフライン修正シート");

  const dumpReviewType = function(title, sh) {
    console.log("==================================================");
    console.log(title);
    console.log("==================================================");
    if (sh && sh.getLastRow() >= 2) {
      const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 24).getDisplayValues();
      rows.forEach(function(r, idx) {
        if (!r[3] && !r[4]) return;
        console.log(`\n--- [No.${idx + 1}] 状態:${r[1]} | 日付:${r[3].replace(/\n/g, ",")} | 拠点:${r[4].replace(/\n/g, ",")} | 医師:${r[8].replace(/\n/g, "・")} ---`);
        console.log(`  時給区分: ${r[11]} | 相談日給: ${r[13].replace(/\n/g, "/")} | 担当: ${r[15]}`);
        console.log(`  作業メモ案: ${r[17]}`);
        if (r[16]) {
          console.log(`  【送信文案】:\n${r[16]}`);
        }
      });
    } else {
      console.log("データがありません");
    }
  };

  dumpReviewType("■ 1-A. オフライン修正シート", offlineSheet);
  dumpReviewType("■ 1-B. 文案・Slack投稿レビュー", revSheet);
}
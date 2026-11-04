/**
 * ============================================================================
 * 05_Slack_And_D6Post.gs
 * 了承行のD6「作業メモ」転記・Slack投稿（1行集約された複数枠の一括転記・投稿対応版）
 * ============================================================================
 */

function fetchSlackParentThreadMap_(token, channelId) {
  const map = new Map();
  if (!token || !channelId) return map;

  try {
    const histUrl = cleanUrl_("https://" + "slack.com/api/conversations.history") +
      "?channel=" + encodeURIComponent(channelId) + "&limit=100";
    const res = UrlFetchApp.fetch(histUrl, {
      method: "get",
      headers: { Authorization: "Bearer " + token },
      muteHttpExceptions: true
    });
    const data = JSON.parse(res.getContentText());
    if (data.ok && Array.isArray(data.messages)) {
      data.messages.forEach(m => {
        const txt = String(m.text || "");
        if (txt.includes("超募集MTGログ") && (!m.thread_ts || m.thread_ts === m.ts)) {
          const match = txt.match(/(\d{1,2})月\s*(前半|後半)/);
          if (match) {
            const key = `${Number(match[1])}月${match[2]}`;
            if (!map.has(key)) {
              map.set(key, m.ts);
            }
          }
        }
      });
    }
  } catch (e) {
    console.error("Slack親スレッド検索エラー: " + e.message);
  }
  return map;
}

function getSlackThreadKeyAndHeader_(dateStr, fallbackTitle) {
  const norm = normalizeDateStr_(dateStr);
  if (!norm) {
    return { halfKey: "", headerText: fallbackTitle || "```超募集MTGログ```" };
  }
  const d = parseDate_(norm);
  const y = d.getFullYear();
  const m = d.getMonth() + 1;
  const day = d.getDate();
  const lastDay = new Date(y, m, 0).getDate();
  const isFirstHalf = day <= 15;
  const halfLabel = isFirstHalf ? "前半" : "後半";
  const rangeStr = isFirstHalf ? `${m}/1~${m}/15` : `${m}/16~${m}/${lastDay}`;
  return {
    halfKey: `${m}月${halfLabel}`,
    headerText: `\`\`\`【${m}月${halfLabel}】${rangeStr} 超募集MTGログ\`\`\``
  };
}

function formatSlackDateHeader_(dateStr) {
  const norm = normalizeDateStr_(dateStr);
  if (!norm) return "";
  const d = parseDate_(norm);
  if (!d || isNaN(d.getTime())) return norm;
  const days = ["日", "月", "火", "水", "木", "金", "土"];
  return `${d.getMonth() + 1}月${d.getDate()}日 (${days[d.getDay()]})`;
}

function normalizeSlackBodyDateHeader_(slackText, targetDate) {
  const header = formatSlackDateHeader_(targetDate);
  if (!header) return slackText;
  const s = String(slackText || "").trim();
  if (!s) return "";
  if (/^`[^`\r\n]+`/.test(s)) {
    return s.replace(/^`[^`\r\n]+`/, "`" + header + "`");
  }
  return "`" + header + "`\n" + s;
}

function postToSlackThread_(threadTitle, text, targetDate, parentThreadMap) {
  const token = PropertiesService.getScriptProperties().getProperty("SLACK_BOT_TOKEN");
  if (!token) return "SLACK_TOKEN未設定(作業メモ転記のみ完了)";

  const postUrl = cleanUrl_("https://" + "slack.com/api/chat.postMessage");
  const { halfKey, headerText } = getSlackThreadKeyAndHeader_(targetDate, threadTitle);
  const formattedText = normalizeSlackBodyDateHeader_(text, targetDate);

  let parentTs = (halfKey && parentThreadMap && parentThreadMap.has(halfKey))
    ? parentThreadMap.get(halfKey)
    : "";

  if (!parentTs) {
    const parentRes = UrlFetchApp.fetch(postUrl, {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: "Bearer " + token },
      payload: JSON.stringify({
        channel: SLACK_CHANNEL_ID,
        text: headerText
      }),
      muteHttpExceptions: true
    });
    const parentData = JSON.parse(parentRes.getContentText());
    if (!parentData.ok) {
      throw new Error("Slack親スレッド作成エラー: " + parentData.error);
    }
    parentTs = parentData.ts;
    if (halfKey && parentThreadMap) {
      parentThreadMap.set(halfKey, parentTs);
    }
    Utilities.sleep(500);
  }

  const res = UrlFetchApp.fetch(postUrl, {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + token },
    payload: JSON.stringify({
      channel: SLACK_CHANNEL_ID,
      thread_ts: parentTs,
      text: formattedText
    }),
    muteHttpExceptions: true
  });
  const data = JSON.parse(res.getContentText());
  if (!data.ok) throw new Error("Slack投稿エラー: " + data.error);
  return `${SLACK_CHANNEL_ID}:${data.ts}`;
}

/**
 * ★1行に集約された複数枠（改行区切りの日付・拠点・不在時間）を展開して全枠分のD6転記＆Slack投稿を行う関数
 */
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
    "医師への送信なども済みですか？\n投稿後、対象行はこのシートから削除されます。",
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

  const token = PropertiesService.getScriptProperties().getProperty("SLACK_BOT_TOKEN");
  const parentThreadMap = fetchSlackParentThreadMap_(token, SLACK_CHANNEL_ID);

  const completedSheetRows = [];

  for (const idx of targetIndices) {
    const sheetRowNum = idx + 2;
    const row = dispValues[idx];

    // ★1行に複数枠が改行区切りで集約されている場合に対応して展開する
    const dateList = String(row[3] || "").split(/\r?\n/).map(s => normalizeDateStr_(s)).filter(Boolean);
    const clinicList = String(row[4] || "").split(/\r?\n/).map(s => s.trim()).filter(Boolean);
    const timeList = String(row[5] || "").split(/\r?\n/).map(s => normalizeTimeRange_(s)).filter(Boolean);
    const targetDept = String(row[6] || "").trim();
    const approachType = String(row[7] || "").trim();
    const doctorsText = String(row[8] || "").trim();
    const staffName = extractCleanStaffName_(String(row[15] || "").trim());
    const memoToAppend = String(row[17] || "").trim();
    const fallbackThreadTitle = String(row[18] || "").trim();
    const slackBlocks = String(row[20] || "").split(/\n\s*\n/).map(s => s.trim()).filter(Boolean);

    const slotCount = Math.max(dateList.length, clinicList.length, 1);

    try {
      for (let sIdx = 0; sIdx < slotCount; sIdx++) {
        const targetDate = dateList[sIdx] || dateList[0] || "";
        const targetClinic = clinicList[sIdx] || clinicList[0] || "";
        const targetTime = timeList[sIdx] || timeList[0] || "";
        const slackText = slackBlocks[sIdx] || slackBlocks[0] || "";
        const threadTitle = buildSlackThreadTitle_(targetDate) || fallbackThreadTitle;

        const targetNormClinic = normalizeClinicForMatch_(targetClinic);
        const exactTimeD6Rows = [];
        const sameClinicDateD6Rows = [];

        for (let r = 2; r < d6Keys.length; r++) {
          const rDate = normalizeDateStr_(d6Keys[r][colDate - 1]);
          if (rDate !== targetDate) continue;

          const rClinic = String(d6Keys[r][colClinic - 1] || "").trim();
          const rNormClinic = normalizeClinicForMatch_(rClinic);
          const rTime = normalizeTimeRange_(d6Keys[r][colTime - 1]);

          const isKitaOrKame = (targetNormClinic === "北葛西" || targetNormClinic === "亀有");
          const isClinicMatch = isKitaOrKame
            ? (rClinic === targetClinic || (rNormClinic === targetNormClinic && (!targetDept || rClinic.includes(targetDept))))
            : (rClinic === targetClinic || rNormClinic === targetNormClinic);

          if (isClinicMatch) {
            sameClinicDateD6Rows.push(r + 1);
            if (!targetTime || !rTime || rTime === targetTime || isTimeOverlapping_(rTime, targetTime)) {
              exactTimeD6Rows.push(r + 1);
            }
          }
        }

        const matchedD6Rows = exactTimeD6Rows.length > 0 ? exactTimeD6Rows : sameClinicDateD6Rows;

        if (matchedD6Rows.length > 0 && memoToAppend) {
          for (const matchedD6Row of matchedD6Rows) {
            const memoCell = d6Sheet.getRange(matchedD6Row, colMemo);
            const currentMemo = String(memoCell.getDisplayValue() || "").trim();
            const cleanTargetMemo = memoToAppend.replace(/^・/, "").trim();
            const lines = currentMemo ? currentMemo.split(/\r?\n/).map(l => l.trim()).filter(Boolean) : [];
            const alreadyExists = lines.some(l => l.replace(/^・/, "").trim() === cleanTargetMemo || l.includes(cleanTargetMemo));
            if (!alreadyExists) {
              lines.push(memoToAppend);
              memoCell.setValue(lines.join("\n"));
            }
          }
          SpreadsheetApp.flush();
        }

        const postId = postToSlackThread_(threadTitle, slackText, targetDate, parentThreadMap);

        if (histSheet && doctorsText && doctorsText !== "該当候補なし") {
          const docList = doctorsText.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
          for (const docName of docList) {
            appendRowAndTrimBlank_(
              histSheet,
              [nowStr, targetDate, targetClinic, targetTime, targetDept, approachType, docName, staffName || userEmail, postId],
              9
            );
          }
        }

        if (logSheet) {
          const actionLabel = matchedD6Rows.length > 0
            ? `D6転記・Slack投稿完了(${revSheet.getName()})`
            : `Slack投稿完了(D6該当枠なし/${revSheet.getName()})`;
          appendRowAndTrimBlank_(
            logSheet,
            [nowStr, actionLabel, postId, `${targetDate} ${targetClinic}${targetTime}`, 1, `${memoToAppend} (${staffName || userEmail})`],
            6
          );
        }

        if (slotCount > 1) Utilities.sleep(500);
      }

      completedSheetRows.push(sheetRowNum);
      Utilities.sleep(600);
    } catch (e) {
      revSheet.getRange(sheetRowNum, 2).setValue("エラー");
      revSheet.getRange(sheetRowNum, 24).setValue(e.message);
    }
  }

  if (completedSheetRows.length > 0) {
    completedSheetRows.sort((a, b) => b - a);
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
      const cwUrl = cleanUrl_("https://" + "api.chatwork.com/v2/rooms/" + CHATWORK_NOTIFY_ROOM_ID + "/messages");
      const cwRes = UrlFetchApp.fetch(cwUrl, {
        method: "post",
        headers: { "X-ChatWorkToken": cwToken },
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
      const slackUrl = cleanUrl_("https://" + "slack.com/api/chat.postMessage");
      const slackRes = UrlFetchApp.fetch(slackUrl, {
        method: "post",
        contentType: "application/json",
        headers: { Authorization: "Bearer " + slackToken },
        payload: JSON.stringify({
          channel: SLACK_NOTIFY_CHANNEL_ID,
          text: slackMessage
        }),
        muteHttpExceptions: true
      });
      const sData = JSON.parse(slackRes.getContentText());
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
  const confSheet = ss.getSheetByName(SHEET_NAMES.D7_CONFIRM);

  const dumpReviewType = (title, sh) => {
    console.log("==================================================");
    console.log(title);
    console.log("==================================================");
    if (sh && sh.getLastRow() >= 2) {
      const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 24).getDisplayValues();
      rows.forEach((r, idx) => {
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
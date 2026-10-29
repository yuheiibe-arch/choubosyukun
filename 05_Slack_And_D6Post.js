/**
 * ============================================================================
 * 05_Slack_And_D6Post.gs
 * 了承行のD6「作業メモ」転記・Slack投稿（親スレッド自動検索＆返信）・完了通知・出力確認
 * ============================================================================
 */

/**
 * Slackチャンネルから既存の「超募集MTGログ」親スレッド（月・前半/後半）を検索して Map("10月後半" -> ts) を返す
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

/**
 * 対象日から「10月後半」などのキーと、親スレッド未作成時の公式ヘッダー文字列を作る
 */
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

/**
 * Slack本文用の日付見出し「10月24日 (土)」を生成する
 */
function formatSlackDateHeader_(dateStr) {
  const norm = normalizeDateStr_(dateStr);
  if (!norm) return "";
  const d = parseDate_(norm);
  if (!d || isNaN(d.getTime())) return norm;
  const days = ["日", "月", "火", "水", "木", "金", "土"];
  return `${d.getMonth() + 1}月${d.getDate()}日 (${days[d.getDay()]})`;
}

/**
 * 既存のSlack本文の先頭行（`10/24` など）を「`10月24日 (土)`」形式に正規化する
 */
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

/**
 * 親スレッド（thread_ts）へ返信する。親スレッドがまだ無い場合は親を作成してから返信する
 */
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
 * D7レビューシートの「確認済」チェック行を一括でD6転記・Slack投稿・履歴記録する
 */
function executeApprovedRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const revSheet = ss.getSheetByName(SHEET_NAMES.D7_REVIEW);
  const histSheet = ss.getSheetByName(SHEET_NAMES.D7_HISTORY);
  const logSheet = ss.getSheetByName(SHEET_NAMES.D7_LOG);
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
    ui.alert("「確認済」にチェックが入っている対象行がありません。");
    return;
  }

  const confirmResponse = ui.alert(
    "Slack投稿・転記の確認",
    "医師への送信なども済みですか？\n投稿後、このスレッドは削除されます。",
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
    const targetDate = normalizeDateStr_(row[3]);
    const targetClinic = String(row[4] || "").trim();
    const targetNormClinic = normalizeClinicForMatch_(targetClinic);
    const targetTime = normalizeTimeRange_(row[5]);
    const targetDept = String(row[6] || "").trim();
    const approachType = String(row[7] || "").trim();
    const doctorsText = String(row[8] || "").trim();
    const staffName = extractCleanStaffName_(String(row[15] || "").trim());
    const memoToAppend = String(row[17] || "").trim();
    const threadTitle = String(row[18] || "").trim();
    const slackText = String(row[20] || "").trim();

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

    try {
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
          ? "D6転記・Slack投稿完了"
          : "Slack投稿完了(D6該当枠なし)";
        appendRowAndTrimBlank_(
          logSheet,
          [nowStr, actionLabel, postId, `${targetDate} ${targetClinic}${targetTime}`, 1, `${memoToAppend} (${staffName || userEmail})`],
          6
        );
      }

      completedSheetRows.push(sheetRowNum);
      Utilities.sleep(800);
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

/**
 * 📢 データ取込完了時にChatwork（[toall] + [info]）とSlack（<!here> + コードブロック）へ通知を送る
 */
function sendImportCompletionNotifications_() {
  const todayStr = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");
  const bodyText = `${todayStr} 超募集くんデータ作成済み\nURL:${D7_REVIEW_SHEET_URL}`;
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
  } else {
    console.log("CHATWORK_API_KEY が未設定のためChatwork通知をスキップしました。");
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
  } else {
    console.log("SLACK_BOT_TOKEN が未設定のためSlack通知をスキップしました。");
  }
}

/**
 * デバッグ・確認用：現在のD7レビューシート・確認リストの内容をログ出力する
 */
function inspectCurrentOutput() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const revSheet = ss.getSheetByName(SHEET_NAMES.D7_REVIEW);
  const confSheet = ss.getSheetByName(SHEET_NAMES.D7_CONFIRM);

  console.log("==================================================");
  console.log("■ 1. 文案・Slack投稿レビュー（全件・全文出力）");
  console.log("==================================================");
  if (revSheet && revSheet.getLastRow() >= 2) {
    const rows = revSheet.getRange(2, 1, revSheet.getLastRow() - 1, 24).getDisplayValues();
    rows.forEach((r, idx) => {
      if (!r[3] && !r[4]) return;
      console.log(`\n--- [No.${idx + 1}] 状態:${r[1]} | 日付:${r[3]} | 拠点:${r[4]} | 不在:${r[5]} | 科:${r[6]} ---`);
      console.log(`  種別: ${r[7]} | 医師: ${r[8].replace(/\n/g, "・")} | 現在シフト: ${r[9]} -> 相談シフト: ${r[10]}`);
      console.log(`  時給区分: ${r[11]} | 現在日給: ${r[12]}円 | 相談日給: ${r[13]}円 | 依頼手当: ${r[14]}円 | 担当: ${r[15]}`);
      console.log(`  作業メモ案: ${r[17]}`);
      console.log(`  根拠: ${r[21].replace(/\n/g, " / ")}`);
      if (r[16]) {
        console.log(`  【送信文案】:\n${r[16]}`);
      } else {
        console.log(`  【送信文案】: (作成なし)`);
      }
    });
  } else {
    console.log("データがありません");
  }

  console.log("\n==================================================");
  console.log("■ 2. 確認リスト（全件出力）");
  console.log("==================================================");
  if (confSheet && confSheet.getLastRow() >= 2) {
    const cRows = confSheet.getRange(2, 1, confSheet.getLastRow() - 1, 11).getDisplayValues();
    cRows.forEach((r, idx) => {
      if (!r[1] && !r[2]) return;
      console.log(
        `[確認${idx + 1}] 日付:${r[1]} | 音声:${r[2]} -> 補正:${r[3]}(${r[4]}) | 時間:${r[5]} | 種別:${r[6]} | 医師:${r[7]} | 突合:${r[8]} | 理由:${r[9]} | 発言:${r[10]}`
      );
    });
  } else {
    console.log("データがありません");
  }
}
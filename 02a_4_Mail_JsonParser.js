/**
 * ============================================================================
 * 02a_4_Mail_JsonParser.gs（旧02b_Setup 4分割の4）
 * Gmailからの最新JSONメールバッチ収集（collectLatestJsonBatchFromGmail_）＆ 安全JSONパーサー
 * ============================================================================
 */

function collectLatestJsonBatchFromGmail_(isOfflineDocMode) {
  const searchQuery = isOfflineDocMode
    ? 'subject:"【Agent解析JSON】ドキュメント確定分"'
    : 'subject:"【Agent解析JSON"';

  const threads = GmailApp.search(searchQuery, 0, 8);
  if (!threads || threads.length === 0) return null;

  const validMsgs = [];
  for (const th of threads) {
    for (const m of th.getMessages()) {
      const subj = String(m.getSubject() || "");
      const isDocConfirmMail = subj.includes("ドキュメント確定分");

      if (isOfflineDocMode && !isDocConfirmMail) continue;
      if (!isOfflineDocMode && isDocConfirmMail) continue;

      const rawBody = String(m.getPlainBody() || "").trim();
      if (!rawBody || rawBody === "[]" || rawBody.includes("文字起こしなし")) continue;

      const parsed = parseEmailJsonSafely_(rawBody);
      if (Array.isArray(parsed) && parsed.length > 0) {
        const validItems = parsed.filter(function(item) {
          return item && item.status !== "文字起こしなし" && (item.date || item.clinic_name || item.raw_clinic_name);
        });
        if (validItems.length > 0) {
          validMsgs.push({
            msg: m,
            items: validItems,
            bodyHash: rawBody.length + "_" + rawBody.slice(0, 120)
          });
        }
      }
    }
  }

  if (validMsgs.length === 0) return null;

  validMsgs.sort(function(a, b) { return a.msg.getDate().getTime() - b.msg.getDate().getTime(); });
  const latestEntry = validMsgs[validMsgs.length - 1];
  const latestMsg = latestEntry.msg;
  const latestTime = latestMsg.getDate().getTime();
  const latestMsgDateYmd = Utilities.formatDate(latestMsg.getDate(), "Asia/Tokyo", "yyyy/MM/dd");

  const batchEntries = validMsgs.filter(function(e) {
    return Math.abs(latestTime - e.msg.getDate().getTime()) <= 20 * 60 * 1000;
  });

  const uniqueBatchEntries = [];
  const seenBodyHashes = new Set();
  for (const entry of batchEntries) {
    if (!seenBodyHashes.has(entry.bodyHash)) {
      seenBodyHashes.add(entry.bodyHash);
      uniqueBatchEntries.push(entry);
    }
  }

  const mergedItems = [];
  const seenItemKeys = new Set();

  for (const entry of uniqueBatchEntries) {
    for (const item of entry.items) {
      const dedupKey = [
        item.date || "",
        item.raw_clinic_name || "",
        item.target_time || "",
        item.approach_type || "",
        item.target_doctor || "",
        item.license_no || ""
      ].join("|");

      if (!seenItemKeys.has(dedupKey)) {
        seenItemKeys.add(dedupKey);
        mergedItems.push(item);
      }
    }
  }

  return {
    items: mergedItems,
    msgCount: uniqueBatchEntries.length,
    latestMsgId: latestMsg.getId(),
    latestSubject: latestMsg.getSubject(),
    latestMsgDateYmd: latestMsgDateYmd
  };
}

function parseEmailJsonSafely_(rawText) {
  if (!rawText) return null;
  const match = String(rawText).match(/\[[\s\S]*\]/);
  if (!match) return null;

  let jsonStr = match[0];
  let cleaned = "";
  let inString = false;
  let isEscaped = false;

  for (let i = 0; i < jsonStr.length; i++) {
    const ch = jsonStr[i];
    const code = jsonStr.charCodeAt(i);
    if (inString) {
      if (isEscaped) {
        cleaned += ch;
        isEscaped = false;
      } else if (ch === "\\") {
        cleaned += ch;
        isEscaped = true;
      } else if (ch === '"') {
        cleaned += ch;
        inString = false;
      } else if (code < 0x20) {
        continue;
      } else {
        cleaned += ch;
      }
    } else {
      if (ch === '"') inString = true;
      cleaned += ch;
    }
  }

  try {
    return JSON.parse(cleaned);
  } catch (e) {
    console.log("JSONパースエラーをスキップしました: " + e.message);
    return null;
  }
}
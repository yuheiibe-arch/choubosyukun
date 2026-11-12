/**
 * ============================================================================
 * 05b_Slack_ApiClient.gs（Slack＆D6転記 3分割の2）
 * ・コピペ時のMarkdown自動リンク化を100%防ぐURLビルダー搭載
 * ・同日1メッセージ自動編集追記（chat.update）＆ 新規投稿（chat.postMessage）API通信
 * ============================================================================
 */

var SLACK_BOT_IDENTITY_CACHE_ = null;

/**
 * ドットを含まない配列結合でURLを生成（コピペ時の [url](https://url) 化を完全防止）
 */
function buildSlackApiUrl_(apiMethod) {
  const host = ["slack", "com"].join(".");
  const method = String(apiMethod || "").replace(/^\/+/, "").trim();
  return "https://" + host + "/api/" + method;
}

function parseSlackApiJsonSafely_(res, apiLabel) {
  const code = res.getResponseCode();
  const text = String(res.getContentText() || "");
  if (text.trim().startsWith("<")) {
    throw new Error(`${apiLabel} がHTMLを返しました (HTTP ${code})`);
  }
  return JSON.parse(text);
}

function getSlackBotIdentityCached_(token) {
  if (SLACK_BOT_IDENTITY_CACHE_) return SLACK_BOT_IDENTITY_CACHE_;
  SLACK_BOT_IDENTITY_CACHE_ = { userId: "", botId: "" };
  if (!token) return SLACK_BOT_IDENTITY_CACHE_;

  try {
    const res = UrlFetchApp.fetch(buildSlackApiUrl_("auth.test"), {
      method: "get",
      headers: { Authorization: "Bearer " + token.trim() },
      muteHttpExceptions: true
    });
    const data = parseSlackApiJsonSafely_(res, "auth.test");
    if (data && data.ok) {
      SLACK_BOT_IDENTITY_CACHE_ = {
        userId: String(data.user_id || ""),
        botId: String(data.bot_id || "")
      };
    }
  } catch (e) {
    console.error("Slack Bot識別取得エラー: " + e.message);
  }
  return SLACK_BOT_IDENTITY_CACHE_;
}

function fetchSlackParentThreadMap_(token, channelId) {
  const map = new Map();
  if (!token || !channelId) return map;

  try {
    const histUrl = buildSlackApiUrl_("conversations.history") +
      "?channel=" + encodeURIComponent(channelId) + "&limit=100";
    const res = UrlFetchApp.fetch(histUrl, {
      method: "get",
      headers: { Authorization: "Bearer " + token.trim() },
      muteHttpExceptions: true
    });
    const data = parseSlackApiJsonSafely_(res, "conversations.history");
    if (data.ok && Array.isArray(data.messages)) {
      data.messages.forEach(function(m) {
        const txt = String(m.text || "");
        if (txt.includes("超募集MTGログ") && (!m.thread_ts || m.thread_ts === m.ts)) {
          const match = txt.match(/(\d{1,2})月\s*(前半|後半)/);
          if (match) {
            const key = Number(match[1]) + "月" + match[2];
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
 * 親スレッド内の既存返信を取得し、「M/D -> { ts, text }」（自Bot投稿分）のマップを返す
 */
function fetchSlackThreadDateRepliesMap_(token, channelId, parentTs, botIdentity) {
  const dateReplyMap = new Map();
  if (!token || !channelId || !parentTs) return dateReplyMap;

  try {
    const repUrl = buildSlackApiUrl_("conversations.replies") +
      "?channel=" + encodeURIComponent(channelId) +
      "&ts=" + encodeURIComponent(parentTs) +
      "&limit=200";
    const res = UrlFetchApp.fetch(repUrl, {
      method: "get",
      headers: { Authorization: "Bearer " + token.trim() },
      muteHttpExceptions: true
    });
    const data = parseSlackApiJsonSafely_(res, "conversations.replies");
    if (data && data.ok && Array.isArray(data.messages)) {
      const replies = data.messages.slice(1);
      replies.forEach(function(msg) {
        const isOwnBot =
          (botIdentity && botIdentity.userId && msg.user === botIdentity.userId) ||
          (botIdentity && botIdentity.botId && msg.bot_id === botIdentity.botId);
        if (!isOwnBot) return;

        const rawText = String(msg.text || "");
        const dm = rawText.match(/`(\d{1,2})月\s*(\d{1,2})日\s*[\(（][日月火水木金土][\)）]`/);
        if (dm) {
          const mdKey = Number(dm[1]) + "/" + Number(dm[2]);
          dateReplyMap.set(mdKey, { ts: msg.ts, text: rawText });
        }
      });
    }
  } catch (e) {
    console.error("Slackスレッド返信一覧取得エラー: " + e.message);
  }
  return dateReplyMap;
}

/**
 * 特定日付のメッセージを親スレッド内で検索し、存在すれば chat.update（追記編集）、無ければ chat.postMessage する
 */
function upsertDateStructToSlackThread_(targetDate, threadTitle, incomingStruct, parentThreadMap, threadRepliesCacheOpt) {
  const rawToken = PropertiesService.getScriptProperties().getProperty("SLACK_BOT_TOKEN");
  const token = rawToken ? rawToken.trim() : "";
  if (!token) return "SLACK_TOKEN未設定(作業メモ転記のみ完了)";

  const postUrl = buildSlackApiUrl_("chat.postMessage");
  const updateUrl = buildSlackApiUrl_("chat.update");
  const { halfKey, headerText } = getSlackThreadKeyAndHeader_(targetDate, threadTitle);

  let parentTs = (halfKey && parentThreadMap && parentThreadMap.has(halfKey))
    ? parentThreadMap.get(halfKey)
    : "";

  if (!parentTs) {
    const parentRes = UrlFetchApp.fetch(postUrl, {
      method: "post",
      contentType: "application/json; charset=utf-8",
      headers: { Authorization: "Bearer " + token },
      payload: JSON.stringify({
        channel: SLACK_CHANNEL_ID,
        text: headerText
      }),
      muteHttpExceptions: true
    });
    const parentData = parseSlackApiJsonSafely_(parentRes, "chat.postMessage(親スレッド)");
    if (!parentData.ok) {
      throw new Error("Slack親スレッド作成エラー: " + parentData.error);
    }
    parentTs = parentData.ts;
    if (halfKey && parentThreadMap) {
      parentThreadMap.set(halfKey, parentTs);
    }
    Utilities.sleep(400);
  }

  const botIdentity = getSlackBotIdentityCached_(token);
  let dateRepliesMap = (threadRepliesCacheOpt && threadRepliesCacheOpt.has(parentTs))
    ? threadRepliesCacheOpt.get(parentTs)
    : null;

  if (!dateRepliesMap) {
    dateRepliesMap = fetchSlackThreadDateRepliesMap_(token, SLACK_CHANNEL_ID, parentTs, botIdentity);
    if (threadRepliesCacheOpt) {
      threadRepliesCacheOpt.set(parentTs, dateRepliesMap);
    }
  }

  const mdKey = getSlackMdKeyFromDateStr_(targetDate);
  const existingReply = mdKey ? dateRepliesMap.get(mdKey) : null;

  if (existingReply && existingReply.ts) {
    const existingStruct = parseSlackDateMessageStruct_(existingReply.text);
    const mergedStruct = mergeSlackDateMessageStructs_(existingStruct, incomingStruct);
    const updatedText = serializeSlackDateMessageStruct_(targetDate, mergedStruct);

    if (updatedText.trim() === String(existingReply.text || "").replace(/&gt;/g, ">").trim()) {
      return `${SLACK_CHANNEL_ID}:${existingReply.ts}`;
    }

    const upRes = UrlFetchApp.fetch(updateUrl, {
      method: "post",
      contentType: "application/json; charset=utf-8",
      headers: { Authorization: "Bearer " + token },
      payload: JSON.stringify({
        channel: SLACK_CHANNEL_ID,
        ts: existingReply.ts,
        text: updatedText
      }),
      muteHttpExceptions: true
    });
    const upData = parseSlackApiJsonSafely_(upRes, "chat.update");
    if (upData && upData.ok) {
      dateRepliesMap.set(mdKey, { ts: existingReply.ts, text: updatedText });
      return `${SLACK_CHANNEL_ID}:${existingReply.ts}`;
    }
  }

  const freshText = serializeSlackDateMessageStruct_(targetDate, incomingStruct);
  const res = UrlFetchApp.fetch(postUrl, {
    method: "post",
    contentType: "application/json; charset=utf-8",
    headers: { Authorization: "Bearer " + token },
    payload: JSON.stringify({
      channel: SLACK_CHANNEL_ID,
      thread_ts: parentTs,
      text: freshText
    }),
    muteHttpExceptions: true
  });
  const data = parseSlackApiJsonSafely_(res, "chat.postMessage");
  if (!data.ok) throw new Error("Slack投稿エラー: " + data.error);

  if (mdKey) {
    dateRepliesMap.set(mdKey, { ts: data.ts, text: freshText });
  }
  return `${SLACK_CHANNEL_ID}:${data.ts}`;
}

/**
 * 単発投稿互換ラッパー（『文案くん』WebApp等からの呼び出しも自動で同日1メッセージ編集化）
 */
function postToSlackThread_(threadTitle, text, targetDate, parentThreadMap, threadRepliesCacheOpt) {
  const formattedText = normalizeSlackBodyDateHeader_(text, targetDate);
  const incomingStruct = parseSlackDateMessageStruct_(formattedText);
  if (incomingStruct.sections.size === 0) {
    const token = PropertiesService.getScriptProperties().getProperty("SLACK_BOT_TOKEN");
    if (!token) return "SLACK_TOKEN未設定(作業メモ転記のみ完了)";
  }
  return upsertDateStructToSlackThread_(targetDate, threadTitle, incomingStruct, parentThreadMap, threadRepliesCacheOpt);
}
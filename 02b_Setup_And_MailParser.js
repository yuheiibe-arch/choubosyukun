/**
 * ============================================================================
 * 02b_Setup_And_MailParser.gs
 * D7シート初期セットアップ・JSONメール収集＆パース・複数日まとめ統合・過去完了報告判定
 * ============================================================================
 */

function setupD7Workspace() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const reviewHeaders = [
    "確認済", "状態", "確認事項・候補", "日付", "拠点名(D6突合キー)",
    "不在時間(D6突合キー)", "診療科", "アプローチ種別", "対象医師名", "現在シフト(拠点・時間)",
    "相談シフト(拠点・時間)", "特別時給", "現在日給(円)", "相談日給(円)", "依頼手当(円)",
    "担当者", "送信文案（社内システム貼付用）", "作業メモ（D6転記内容）", "投稿先スレッド", "投稿種別",
    "Slack投稿本文", "根拠（参照シート・MTG発言）", "投稿ID・処理日時", "エラー内容"
  ];
  const revSheet = ensureSheet_(ss, SHEET_NAMES.D7_REVIEW, reviewHeaders, "#1a73e8");
  revSheet.setColumnWidth(9, 140);
  revSheet.setColumnWidth(17, 360);
  revSheet.setColumnWidth(18, 240);
  revSheet.setColumnWidth(21, 300);
  revSheet.setColumnWidth(22, 260);
  trimBlankRowsAndCols_(revSheet, reviewHeaders.length);
  applyLeftMiddleToSheet_(revSheet);

  const confirmHeaders = [
    "取込日時", "日付", "文字起こし上の拠点名", "補正後の拠点名候補", "診療科判定",
    "時間帯", "アプローチ種別", "抽出医師名", "突合結果", "要確認の理由", "MTG発言抜粋"
  ];
  const confSheet = ensureSheet_(ss, SHEET_NAMES.D7_CONFIRM, confirmHeaders, "#e65100");
  trimBlankRowsAndCols_(confSheet, confirmHeaders.length);
  applyLeftMiddleToSheet_(confSheet);

  const historyHeaders = [
    "記録日時", "対象日付", "拠点名(D6突合キー)", "不在時間", "診療科",
    "アプローチ種別", "打診医師名", "担当者", "証跡・Slack投稿ID"
  ];
  const histSheet = ensureSheet_(ss, SHEET_NAMES.D7_HISTORY, historyHeaders, "#00897b");
  trimBlankRowsAndCols_(histSheet, historyHeaders.length);
  applyLeftMiddleToSheet_(histSheet);

  const dictHeaders = ["区分(拠点名 / その他)", "誤変換・揺れ表記（変換前）", "正規表記（変換後）", "メモ・備考"];
  const dictSheet = ensureSheet_(ss, SHEET_NAMES.D7_DICT, dictHeaders, "#8e24aa");
  populateInitialClinicDictSheet_(dictSheet);
  trimBlankRowsAndCols_(dictSheet, dictHeaders.length);
  applyLeftMiddleToSheet_(dictSheet);

  const holidayHeaders = ["祝日日付(YYYY/MM/DD)", "祝日名称", "取得日時"];
  const holSheet = ensureSheet_(ss, SHEET_NAMES.D7_HOLIDAY, holidayHeaders, "#0f9d58");
  updateCabinetOfficeHolidays_(holSheet);
  trimBlankRowsAndCols_(holSheet, holidayHeaders.length);
  applyLeftMiddleToSheet_(holSheet);

  const logHeaders = ["処理日時", "種別", "メールID / 投稿ID", "件名 / スレッド名", "処理件数", "詳細"];
  const logSheet = ensureSheet_(ss, SHEET_NAMES.D7_LOG, logHeaders, "#5f6368");
  trimBlankRowsAndCols_(logSheet, logHeaders.length);
  applyLeftMiddleToSheet_(logSheet);

  const def = ss.getSheetByName("シート1");
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) {
    ss.deleteSheet(def);
  }
  console.log("「超募集くん（D7）」の初期セットアップ（『打診履歴』シート含む）が完了しました！");
}

function inferTimeFromCategory_(timeCategory, isKitaKasai) {
  const tc = String(timeCategory || "").trim();
  const nightEnd = isKitaKasai ? "20:00" : "21:00";
  const amEnd = isKitaKasai ? "12:30" : "13:00";
  const pmStart = isKitaKasai ? "14:00" : "15:00";
  if (tc === "午前") return "09:00-" + amEnd;
  if (tc === "午後") return pmStart + "-18:00";
  if (tc === "夜間") return "18:00-" + nightEnd;
  if (tc === "午後・夜間") return pmStart + "-" + nightEnd;
  if (tc === "午前・午後") return "09:00-18:00";
  if (tc === "終日") return "09:00-" + nightEnd;
  return "";
}

/**
 * 発言内容が「すでに完了した過去の調整の事後報告（移動してもらいました／埋まりました等）」のみであり、
 * 今後のアクション（しましょう／お願いしよう等）を含まない場合は true を返す
 */
function isPastCompletedReport_(evidenceText) {
  const s = String(evidenceText || "").trim();
  if (!s) return false;
  const hasFutureIntent = /しましょう|ましょう|お願い|打診|聞いて|相談|調整し|いけるか|もらえるよう/.test(s);
  if (hasFutureIntent) return false;
  return /してもらいました|いただきました|埋まりました|充足しました|変更済み|移動済み|完了しました/.test(s);
}

function consolidateMultiDateDoctorDrafts_(reviewRows, draftParamsByRowIdx) {
  const doctorGroups = new Map();

  for (let i = 0; i < draftParamsByRowIdx.length; i++) {
    const p = draftParamsByRowIdx[i];
    if (!p || !p.doctorName || p.doctorName === "{医師名}" || p.approachType !== "勤務地変更相談") continue;
    if (p.rowState !== "未レビュー") continue;

    if (!doctorGroups.has(p.doctorName)) {
      doctorGroups.set(p.doctorName, []);
    }
    const list = doctorGroups.get(p.doctorName);
    if (!list.some(existing => existing.dateStr === p.dateStr && existing.clinicName === p.clinicName)) {
      list.push(p);
    }
  }

  for (const [docName, entries] of doctorGroups.entries()) {
    if (entries.length < 2) continue;

    entries.sort((a, b) => a.dateStr.localeCompare(b.dateStr));
    const combinedDraft = buildMultiDateTemplate_(entries);

    const primaryEntry = entries[0];
    const primarySheetRow = primaryEntry.rowIndex + 2;
    const allDatesLabel = entries.map(e => e.dateStr.slice(5)).join(", ");

    reviewRows[primaryEntry.rowIndex][16] = combinedDraft;

    for (let k = 1; k < entries.length; k++) {
      const subIdx = entries[k].rowIndex;
      reviewRows[subIdx][16] = "【※" + primarySheetRow + "行目の" + docName + "先生の送信文案に複数日分（" + allDatesLabel + "）を1通に統合済み】";
    }
  }
}

/**
 * 分割送信された複数の【Agent解析JSON】メール（Part 1, Part 2...）をすべて収集・結合する
 * - 件名に [Part 1/2] が付いていても、付いていなくても（同件名で2通届いても）上書きせず全通を時系列順に統合する
 * - 空配列 [] や「文字起こしなし」のメールは無視する
 */
function collectLatestJsonBatchFromGmail_() {
  const threads = GmailApp.search('subject:"【Agent解析JSON"', 0, 5);
  if (!threads || threads.length === 0) return null;

  const validMsgs = [];
  for (const th of threads) {
    for (const m of th.getMessages()) {
      const rawBody = String(m.getPlainBody() || "").trim();
      if (!rawBody || rawBody === "[]" || rawBody.includes("文字起こしなし")) continue;

      const parsed = parseEmailJsonSafely_(rawBody);
      if (Array.isArray(parsed) && parsed.length > 0) {
        const validItems = parsed.filter(item => item && item.status !== "文字起こしなし" && (item.date || item.clinic_name || item.raw_clinic_name));
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

  validMsgs.sort((a, b) => a.msg.getDate().getTime() - b.msg.getDate().getTime());
  const latestEntry = validMsgs[validMsgs.length - 1];
  const latestMsg = latestEntry.msg;
  const latestTime = latestMsg.getDate().getTime();
  const latestMsgDateYmd = Utilities.formatDate(latestMsg.getDate(), "Asia/Tokyo", "yyyy/MM/dd");

  // 最新の有効メールから前後20分以内に届いた全分割メールを対象にする
  const batchEntries = validMsgs.filter(e => Math.abs(latestTime - e.msg.getDate().getTime()) <= 20 * 60 * 1000);

  // 同一本文の重複転送だけを除外し、Part 1（例:30件）と Part 2（例:8件）は両方とも残す
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
        item.target_doctor || ""
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
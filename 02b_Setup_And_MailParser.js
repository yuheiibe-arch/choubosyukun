/**
 * ============================================================================
 * 02b_Setup_And_MailParser.gs
 * （同一医師の複数枠を「1行」に完全集約・余計なシステム記載カット・シンプル文案版）
 * ============================================================================
 */

function getReviewHeadersCommon_() {
  return [
    "確認済", "状態", "確認事項・候補", "日付", "拠点名(D6突合キー)",
    "不在時間(D6突合キー)", "診療科", "アプローチ種別", "対象医師名", "現在シフト(拠点・時間)",
    "相談シフト(拠点・時間)", "特別時給", "現在日給(円)", "相談日給(円)", "依頼手当(円)",
    "担当者", "送信文案（社内システム貼付用）", "作業メモ（D6転記内容）", "投稿先スレッド", "投稿種別",
    "Slack投稿本文", "根拠（参照シート・MTG発言）", "投稿ID・処理日時", "エラー内容"
  ];
}

function ensureReviewStyleSheet_(ss, sheetName, headerColor) {
  const reviewHeaders = getReviewHeadersCommon_();
  const sh = ensureSheet_(ss, sheetName, reviewHeaders, headerColor || "#1a73e8");
  sh.setColumnWidth(9, 140);
  sh.setColumnWidth(17, 380);
  sh.setColumnWidth(18, 240);
  sh.setColumnWidth(21, 300);
  sh.setColumnWidth(22, 180);
  trimBlankRowsAndCols_(sh, reviewHeaders.length);
  applyLeftMiddleToSheet_(sh);
  return sh;
}

function setupD7Workspace() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureReviewStyleSheet_(ss, SHEET_NAMES.D7_REVIEW, "#1a73e8");

  const offlineSheetName = SHEET_NAMES.D7_OFFLINE_REVIEW || "オフライン修正シート";
  ensureReviewStyleSheet_(ss, offlineSheetName, "#0b8043");

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
  console.log("「超募集くん（D7）」の初期セットアップが完了しました！");
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

function isPastCompletedReport_(evidenceText) {
  const s = String(evidenceText || "").trim();
  if (!s) return false;
  if (s.includes("ドキュメント確定") || s.includes("編集くんカスタム確定")) return false;
  const hasFutureIntent = /しましょう|ましょう|お願い|打診|聞いて|相談|調整し|いけるか|もらえるよう/.test(s);
  if (hasFutureIntent) return false;
  return /してもらいました|いただきました|埋まりました|充足しました|変更済み|移動済み|完了しました/.test(s);
}

/**
 * ★V列（根拠欄）の余計なシステムログ（[MTG発言]、[D6]、[医師特定]、[時給表]など）を削ぎ落とし、
 * 1行の簡潔なメモだけにクリーンアップするヘルパー
 */
function cleanVerboseEvidenceCell_(rawEvidence) {
  const firstLine = String(rawEvidence || "").split(/\r?\n/)[0] || "";
  let s = firstLine
    .replace(/\[MTG発言\]\s*/g, "")
    .replace(/\[ドキュメント確定[：:]([^\]]+)\]/g, "【$1】")
    .replace(/\[編集くんカスタム確定\]/g, "【カスタム】")
    .replace(/【複数枠おまとめ打診.*?$/g, "")
    .replace(/【選択的打診.*?$/g, "")
    .replace(/\[医籍[：:][^\]]+\]/g, "")
    .trim();
  return s || "-";
}

/**
 * ★全面改修版：
 * 1. 全行の「根拠（22列目）」から余計なシステムログを除去して1行にスリム化
 * 2. 同一医師への複数枠の打診を **1行（primaryRow）に完全集約** し、2行目以降の余計な行を `reviewRows` から削除！
 */
function consolidateMultiDateDoctorDrafts_(reviewRows, draftParamsByRowIdx) {
  // まず全行の22列目（根拠欄）から余計なシステムログ（[D6]・[医師特定]・[時給表]など）を削除
  for (let i = 0; i < reviewRows.length; i++) {
    reviewRows[i][21] = cleanVerboseEvidenceCell_(reviewRows[i][21]);
  }

  const byDoctor = new Map();

  for (let i = 0; i < draftParamsByRowIdx.length; i++) {
    const p = draftParamsByRowIdx[i];
    if (!p || !p.doctorName || p.doctorName === "{医師名}" || p.doctorName.includes("\n")) continue;
    if (p.rowState !== "未レビュー") continue;
    if (p.approachType === "募集継続・ステイ") continue;

    const cleanDocKey = String(p.doctorName).replace(/先生$/g, "").trim();
    if (!cleanDocKey) continue;

    if (!byDoctor.has(cleanDocKey)) {
      byDoctor.set(cleanDocKey, []);
    }
    byDoctor.get(cleanDocKey).push(p);
  }

  const rowIndicesToDelete = new Set();

  for (const [docName, rawEntries] of byDoctor.entries()) {
    if (rawEntries.length < 2) continue;

    // 同一枠の完全重複を除外しつつ時系列ソート
    const entries = [];
    rawEntries.forEach(p => {
      const isDup = entries.some(ex =>
        ex.dateStr === p.dateStr &&
        ex.clinicName === p.clinicName &&
        ex.absenceStart === p.absenceStart &&
        ex.approachType === p.approachType
      );
      if (!isDup) {
        entries.push(p);
      } else {
        rowIndicesToDelete.add(p.rowIndex);
      }
    });

    if (entries.length < 2) continue;

    entries.sort((a, b) =>
      a.dateStr.localeCompare(b.dateStr) ||
      String(a.consultStart || a.absenceStart || "").localeCompare(String(b.consultStart || b.absenceStart || ""))
    );

    const uniqueDates = Array.from(new Set(entries.map(e => e.dateStr)));
    const allMovers = entries.every(e => e.approachType === "勤務地変更相談");
    const allIndividuals = entries.every(e => e.approachType === "個別アプローチ");

    let combinedDraft = "";
    if (uniqueDates.length === 1 && allMovers) {
      combinedDraft = buildSameDateMultiClinicMoverTemplate_(entries);
    } else if (allIndividuals) {
      combinedDraft = buildMultiDateIndividualTemplate_(entries);
    } else {
      combinedDraft = buildUnifiedMultiSlotDraftTemplate_(entries);
    }

    // ★先頭の行（primaryIdx）に全枠の情報を集約し、2枠目以降の行はシートから削除する
    const primaryIdx = entries[0].rowIndex;
    const primaryRow = reviewRows[primaryIdx];

    // D列(3): 日付（改行区切り）
    primaryRow[3] = entries.map(e => e.dateStr).join("\n");
    // E列(4): 拠点名（改行区切り）
    primaryRow[4] = entries.map(e => reviewRows[e.rowIndex][4]).join("\n");
    // F列(5): 不在時間（改行区切り）
    primaryRow[5] = entries.map(e => reviewRows[e.rowIndex][5]).join("\n");
    // G列(6): 診療科
    const uniqueDepts = Array.from(new Set(entries.map(e => reviewRows[e.rowIndex][6]).filter(Boolean)));
    primaryRow[6] = uniqueDepts.join("・") || "小児科";
    // H列(7): アプローチ種別
    const uniqueTypes = Array.from(new Set(entries.map(e => e.approachType).filter(Boolean)));
    primaryRow[7] = uniqueTypes.join("・");

    // J列(9): 現在シフト
    const curShifts = Array.from(new Set(entries.map(e => reviewRows[e.rowIndex][9]).filter(v => v && v !== "-")));
    primaryRow[9] = curShifts.length > 0 ? curShifts.join("\n") : "-";

    // K列(10): 相談シフト（日付＋拠点＋時間で見やすく一覧化）
    primaryRow[10] = entries.map(e => `${e.dateStr.slice(5)} ${reviewRows[e.rowIndex][10]}`).join("\n");

    // L列(11): 特別時給
    const anySpecial = entries.some(e => String(reviewRows[e.rowIndex][11] || "").includes("特別時給"));
    primaryRow[11] = anySpecial ? "特別時給" : "通常時給";

    // N列(13): 相談日給
    const wageLabels = entries.map(e => {
      const w = Number(String(reviewRows[e.rowIndex][13] || "0").replace(/,/g, ""));
      return w > 0 ? `${formatComma_(w)}円` : "-";
    });
    primaryRow[13] = wageLabels.every(w => w === "-") ? "" : wageLabels.join("\n");

    // Q列(16): 1通にまとめた送信文案
    primaryRow[16] = combinedDraft;

    // R列(17): 作業メモ（同日複数候補なら候補拠点名を付記、別日おまとめなら共通メモ）
    if (uniqueDates.length === 1 && allMovers) {
      const allClinicsLabel = entries.map(e => e.clinicName).join("・");
      if (primaryRow[17] && !String(primaryRow[17]).includes("候補：")) {
        primaryRow[17] += `（候補：${allClinicsLabel}）`;
      }
    }

    // U列(20): Slack投稿本文（全枠分を改行で統合）
    primaryRow[20] = entries.map(e => reviewRows[e.rowIndex][20]).filter(Boolean).join("\n\n");

    // V列(21): 根拠（シンプルに1行でおまとめ枠数を表示）
    const slotSummary = entries.map(e => `${e.dateStr.slice(5)} ${e.clinicName}`).join("・");
    primaryRow[21] = `おまとめ全${entries.length}枠（${slotSummary}）`;

    // 2枠目以降の行を削除対象に登録
    for (let k = 1; k < entries.length; k++) {
      rowIndicesToDelete.add(entries[k].rowIndex);
    }
  }

  // ★削除対象の行を後ろから splice して reviewRows と draftParamsByRowIdx から完全に除去
  const sortedDel = Array.from(rowIndicesToDelete).sort((a, b) => b - a);
  for (const delIdx of sortedDel) {
    reviewRows.splice(delIdx, 1);
    if (draftParamsByRowIdx && delIdx < draftParamsByRowIdx.length) {
      draftParamsByRowIdx.splice(delIdx, 1);
    }
  }
}

/**
 * パターン1：同日に複数拠点（AまたはB）へ勤務地変更相談を行う際のシンプルテンプレート
 */
function buildSameDateMultiClinicMoverTemplate_(entries) {
  const first = entries[0];
  const docName = String(first.doctorName || "{医師名}").replace(/先生$/g, "");
  const staff = first.staffName || "{担当者名}";
  const mdDay = typeof formatMdDay_ === "function" ? formatMdDay_(first.dateObj) : first.dateStr;
  const fromClinic = first.fromClinicName || "現在拠点";
  const curTime = `${first.currentStart || ""}~${first.currentEnd || ""}`;
  const curWageNum = Number(first.currentDailyWage || 0);
  const curWageStr = curWageNum > 0 ? ` (日給:${formatComma_(curWageNum)}円)` : "";

  const toHeaderStr = entries.map(e => `【${e.clinicName}】`).join(" または ");
  const circledNums = ["①", "②", "③", "④", "⑤", "⑥"];

  const consultLines = entries.map((e, idx) => {
    const numLabel = circledNums[idx] || (idx + 1);
    const conTime = `${e.consultStart || e.absenceStart || ""}~${e.consultEnd || e.absenceEnd || ""}`;
    const conWageNum = Number(e.consultDailyWage || 0);
    const allowSuffix = e.hasRequestAllow ? "+依頼手当10,000円" : "";
    const wagePart = conWageNum > 0
      ? ` (日給:${formatComma_(conWageNum)}円${allowSuffix})`
      : (allowSuffix ? ` (${allowSuffix.replace(/^\+/, "")})` : "");
    return `ご相談${numLabel}) 【${e.clinicName}】${conTime}${wagePart}`;
  });

  return [
    docName + "先生",
    "平素よりお世話になっております。",
    "ドクターサポート事務局の" + staff + "でございます。",
    "",
    "この度は、" + mdDay + "のご勤務につきましてご相談がありご連絡させていただきました。",
    "ご相談の案が多く大変申し訳ございません。",
    "当日のご勤務拠点を 【" + fromClinic + "】 より " + toHeaderStr + " へご変更いただくことは可能でしょうか。",
    "",
    "■" + mdDay,
    "現在)【" + fromClinic + "】 " + curTime + curWageStr,
    "↓",
    ...consultLines,
    "",
    "※上記募集は公開中のため、ご応募いただいた際に充足している可能性もございます。",
    "恐れ入りますが、充足の際はご容赦くださいませ。",
    "",
    "お忙しいところ、ご負担をおかけして恐縮ではございますが、ご検討いただけましたら幸いです。",
    "何卒よろしくお願い申し上げます。",
    "",
    "ドクターサポート事務局\u3000" + staff
  ].join("\n");
}

/**
 * パターン2：個別アプローチ（追加勤務）で複数枠をまとめて打診する際のシンプルテンプレート
 */
function buildMultiDateIndividualTemplate_(entries) {
  const first = entries[0];
  const docName = String(first.doctorName || "{医師名}").replace(/先生$/g, "");
  const staff = first.staffName || "{担当者名}";

  const slotLines = entries.map(e => {
    const ymdLabel = typeof formatYmdDay_ === "function" ? formatYmdDay_(e.dateObj) : e.dateStr;
    const absRange = (e.absenceStart || "") + "~" + (e.absenceEnd || "");
    const conWageNum = Number(e.consultDailyWage || 0);
    const wageText = conWageNum > 0 ? " (日給:" + formatComma_(conWageNum) + "円)" : "";
    return "・" + ymdLabel + " 【" + e.clinicName + "】" + absRange + wageText;
  });

  return [
    docName + "先生",
    "平素より大変お世話になっております。",
    "ドクターサポート事務局の" + staff + "でございます。",
    "",
    "この度は、下記日程のご勤務につきましてご相談がありご連絡させていただきました。",
    "ご相談の案が多く大変申し訳ございません。",
    "現在、下記枠にて医師不在となる可能性がある時間帯がございますため、",
    "もし先生のご都合がよろしければ、いずれかの日程でもご勤務いただくことは可能でしょうか。",
    "",
    "■ご相談日程",
    ...slotLines,
    "",
    "※上記募集は公開中のため、ご応募いただいた際に充足している可能性もございます。",
    "恐れ入りますが、充足の際はご容赦くださいませ。",
    "",
    "急なお願いとなり大変恐れ入りますが、ご無理のない範囲でご検討いただけますと幸いです。",
    "お忙しいところ、ご負担をおかけして恐縮ではございますが、何卒よろしくお願い申し上げます。",
    "",
    "ドクターサポート事務局\u3000" + staff
  ].join("\n");
}

/**
 * パターン3：複数日程の移動、または「移動＋個別」「延長＋別枠」など複数枠が混在する場合のシンプル統合テンプレート
 */
function buildUnifiedMultiSlotDraftTemplate_(entries) {
  const first = entries[0];
  const docName = String(first.doctorName || "{医師名}").replace(/先生$/g, "");
  const staff = first.staffName || "{担当者名}";
  const circledNums = ["①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧"];

  const blocks = entries.map((e, idx) => {
    const numLabel = circledNums[idx] || (idx + 1);
    const ymdLabel = typeof formatYmdDay_ === "function" ? formatYmdDay_(e.dateObj) : e.dateStr;
    const curTime = `${e.currentStart || ""}~${e.currentEnd || ""}`;
    const conTime = `${e.consultStart || e.absenceStart || ""}~${e.consultEnd || e.absenceEnd || ""}`;
    const curWageNum = Number(e.currentDailyWage || 0);
    const conWageNum = Number(e.consultDailyWage || 0);
    const curWageStr = curWageNum > 0 ? ` (日給:${formatComma_(curWageNum)}円)` : "";
    const allowSuffix = (e.approachType === "勤務地変更相談" && e.hasRequestAllow) ? "+依頼手当10,000円" : "";
    const conWageStr = conWageNum > 0
      ? ` (日給:${formatComma_(conWageNum)}円${allowSuffix})`
      : (allowSuffix ? ` (${allowSuffix.replace(/^\+/, "")})` : "");

    if (e.approachType === "勤務地変更相談") {
      const fromC = e.fromClinicName || "現在拠点";
      return [
        `■ご相談${numLabel}：${ymdLabel}`,
        `現在)【${fromC}】${curTime}${curWageStr}`,
        `↓`,
        `ご相談)【${e.clinicName}】${conTime}${conWageStr}`
      ].join("\n");
    } else if (e.approachType === "延長相談" || e.approachType === "早出相談") {
      return [
        `■ご相談${numLabel}：${ymdLabel}`,
        `現在)【${e.clinicName}】${curTime}${curWageStr}`,
        `↓`,
        `ご相談)【${e.clinicName}】${conTime}${conWageStr}`
      ].join("\n");
    } else {
      const absTime = `${e.absenceStart || ""}~${e.absenceEnd || ""}`;
      return [
        `■ご相談${numLabel}：${ymdLabel}`,
        `ご相談)【${e.clinicName}】${absTime}${conWageStr}`
      ].join("\n");
    }
  });

  return [
    docName + "先生",
    "平素より大変お世話になっております。",
    "ドクターサポート事務局の" + staff + "でございます。",
    "",
    "この度は、下記のご勤務につきましてご相談がありご連絡させていただきました。",
    "ご相談の案が多く大変申し訳ございません。",
    "もし先生のご都合がよろしければ、下記のいずれかの枠にてご協力をいただくことは可能でしょうか。",
    "",
    blocks.join("\n\n"),
    "",
    "※上記募集は公開中のため、ご応募いただいた際に充足している可能性もございます。",
    "恐れ入りますが、充足の際はご容赦くださいませ。",
    "",
    "急なお願いとなり大変恐れ入りますが、ご無理のない範囲でご検討いただけますと幸いです。",
    "お忙しいところ、ご負担をおかけして恐縮ではございますが、何卒よろしくお願い申し上げます。",
    "",
    "ドクターサポート事務局\u3000" + staff
  ].join("\n");
}

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

  const batchEntries = validMsgs.filter(e => Math.abs(latestTime - e.msg.getDate().getTime()) <= 20 * 60 * 1000);

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
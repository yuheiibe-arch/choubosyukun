/**
 * ============================================================================
 * 02b_Setup_And_MailParser.gs
 * （1名への複数枠打診の全種別1通統合・「案が多く申し訳ございません」枕詞自動付与・見やすい罫線レイアウト対応版）
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
  sh.setColumnWidth(22, 260);
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
  console.log("「超募集くん（D7）」の初期セットアップ（『オフライン修正シート』『打診履歴』含む）が完了しました！");
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
 * ★全面改良版：1名の医師に対して複数枠（2枠以上）の交渉がある場合、
 * アプローチ種別（移動・個別・延長・早出）や同日・別日を問わず、必ず1通の送信文案に統合する関数。
 * また、複数枠を打診する際は「ご相談の案が多く大変申し訳ございません。」の枕詞を添えて見やすく整形する。
 */
function consolidateMultiDateDoctorDrafts_(reviewRows, draftParamsByRowIdx) {
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
    const list = byDoctor.get(cleanDocKey);
    const isDup = list.some(ex =>
      ex.dateStr === p.dateStr &&
      ex.clinicName === p.clinicName &&
      ex.absenceStart === p.absenceStart &&
      ex.approachType === p.approachType
    );
    if (!isDup) {
      list.push(p);
    }
  }

  for (const [docName, entries] of byDoctor.entries()) {
    if (entries.length < 2) continue;

    entries.sort((a, b) =>
      a.dateStr.localeCompare(b.dateStr) ||
      String(a.consultStart || a.absenceStart || "").localeCompare(String(b.consultStart || b.absenceStart || ""))
    );

    const uniqueDates = Array.from(new Set(entries.map(e => e.dateStr)));
    const allMovers = entries.every(e => e.approachType === "勤務地変更相談");
    const allIndividuals = entries.every(e => e.approachType === "個別アプローチ");

    let combinedDraft = "";
    // パターン1：同じ日に複数の拠点へ「勤務地変更相談（AまたはBの選択的移動）」をかける場合
    if (uniqueDates.length === 1 && allMovers) {
      combinedDraft = buildSameDateMultiClinicMoverTemplate_(entries);
    }
    // パターン2：すべて「個別アプローチ（追加勤務）」で複数枠（同日または複数日）をかける場合
    else if (allIndividuals) {
      combinedDraft = buildMultiDateIndividualTemplate_(entries);
    }
    // パターン3：複数日程の移動、または「移動＋個別」「延長＋別枠」など複数種別・複数枠が混在する場合
    else {
      combinedDraft = buildUnifiedMultiSlotDraftTemplate_(entries);
    }

    const primaryEntry = entries[0];
    const primarySheetRow = primaryEntry.rowIndex + 2;
    const allSlotsShortLabel = entries.map(e => `${e.dateStr.slice(5)} ${e.clinicName}`).join("・");

    reviewRows[primaryEntry.rowIndex][16] = combinedDraft;

    // 同日複数拠点の移動の場合は、作業メモ欄にも「（候補：A・B）」を自動追記
    if (uniqueDates.length === 1 && allMovers) {
      const allClinicsLabel = entries.map(e => e.clinicName).join("・");
      entries.forEach(e => {
        if (reviewRows[e.rowIndex][17] && !String(reviewRows[e.rowIndex][17]).includes("候補：")) {
          reviewRows[e.rowIndex][17] += `（候補：${allClinicsLabel}）`;
        }
      });
    }

    for (let k = 1; k < entries.length; k++) {
      const subIdx = entries[k].rowIndex;
      reviewRows[subIdx][16] = `【※${primarySheetRow}行目の${docName}先生の送信文案に全${entries.length}枠分（${allSlotsShortLabel}）を1通にまとめて出力済み】`;
    }
  }
}

/**
 * パターン1：同日に複数拠点（AまたはB）へ勤務地変更相談を行う際の選択式テンプレート
 * （「ご相談の案が多く大変申し訳ございません。」＋見やすい罫線ボックス付き）
 */
function buildSameDateMultiClinicMoverTemplate_(entries) {
  const first = entries[0];
  const docName = String(first.doctorName || "{医師名}").replace(/先生$/g, "");
  const staff = first.staffName || "{担当者名}";
  const mdDay = typeof formatMdDay_ === "function" ? formatMdDay_(first.dateObj) : first.dateStr;
  const fromClinic = first.fromClinicName || "現在拠点";
  const curTime = `${first.currentStart || ""}~${first.currentEnd || ""}`;
  const curWage = formatComma_(first.currentDailyWage || 0);

  const toHeaderStr = entries.map(e => `【${e.clinicName}】`).join(" または ");
  const circledNums = ["①", "②", "③", "④", "⑤", "⑥"];

  const consultLines = entries.map((e, idx) => {
    const numLabel = circledNums[idx] || (idx + 1);
    const conTime = `${e.consultStart || e.absenceStart || ""}~${e.consultEnd || e.absenceEnd || ""}`;
    const conWage = formatComma_(e.consultDailyWage || 0);
    const allowSuffix = e.hasRequestAllow ? "＋移動依頼手当10,000円" : "";
    return `【候補${numLabel}】【${e.clinicName}】${conTime}（日給:${conWage}円${allowSuffix}）`;
  });

  return [
    docName + "先生",
    "平素より大変お世話になっております。",
    "ドクターサポート事務局の" + staff + "でございます。",
    "",
    "この度は、" + mdDay + "のご勤務につきましてご相談がありご連絡させていただきました。",
    "ご相談の案が多く大変申し訳ございません。",
    "当日のご勤務拠点を【" + fromClinic + "】より、下記の " + toHeaderStr + " のいずれかへご変更いただくことは可能でしょうか。",
    "",
    "━━━━━━━━━━━━━━━━━━━━",
    "■ " + mdDay + " ご相談内容",
    "【現在のご勤務】",
    "・【" + fromClinic + "】" + curTime + "（日給:" + curWage + "円）",
    " ↓",
    "【ご変更候補（いずれか1枠）】",
    ...consultLines,
    "━━━━━━━━━━━━━━━━━━━━",
    "",
    "※上記募集は公開中のため、ご返信いただいた際に充足している可能性もございます。",
    "恐れ入りますが、充足の際はご容赦くださいませ。",
    "",
    "お忙しいところ、ご負担をおかけして大変恐縮ではございますが、ご検討いただけましたら幸いです。",
    "何卒よろしくお願い申し上げます。",
    "",
    "ドクターサポート事務局\u3000" + staff
  ].join("\n");
}

/**
 * パターン2：個別アプローチ（追加勤務）で複数枠をまとめて打診する際のテンプレート
 * （「ご相談の案が多く大変申し訳ございません。」＋見やすい罫線＆番号リスト付き）
 */
function buildMultiDateIndividualTemplate_(entries) {
  const first = entries[0];
  const docName = String(first.doctorName || "{医師名}").replace(/先生$/g, "");
  const staff = first.staffName || "{担当者名}";
  const circledNums = ["①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧"];

  // 同日・同時間帯の重複があるか（AまたはBの選択式か）をチェック
  let hasSameTimeOverlap = false;
  for (let i = 0; i < entries.length; i++) {
    for (let j = i + 1; j < entries.length; j++) {
      if (entries[i].dateStr === entries[j].dateStr) {
        const rangeA = `${entries[i].absenceStart || ""}-${entries[i].absenceEnd || ""}`;
        const rangeB = `${entries[j].absenceStart || ""}-${entries[j].absenceEnd || ""}`;
        if (typeof isTimeOverlapping_ === "function" && isTimeOverlapping_(rangeA, rangeB)) {
          hasSameTimeOverlap = true;
        }
      }
    }
  }

  const slotLines = entries.map((e, idx) => {
    const ymdLabel = typeof formatYmdDay_ === "function" ? formatYmdDay_(e.dateObj) : e.dateStr;
    const absRange = (e.absenceStart || "") + "~" + (e.absenceEnd || "");
    const wageText = e.consultDailyWage ? "（日給:" + formatComma_(e.consultDailyWage) + "円）" : "";
    const numLabel = circledNums[idx] || (idx + 1);
    return `【候補${numLabel}】${ymdLabel} 【${e.clinicName}】${absRange}${wageText}`;
  });

  const askLine = hasSameTimeOverlap
    ? "現在、複数拠点にて医師不在となる可能性がある時間帯がございますため、もし先生のご都合がよろしければ、下記のいずれかの枠にてご勤務いただくことは可能でしょうか。"
    : "現在、複数日程（時間帯）にて医師不在となる可能性がある枠がございますため、もし先生のご都合がよろしければ、下記のいずれかの枠（または複数枠）にてご勤務いただくことは可能でしょうか。";

  return [
    docName + "先生",
    "平素より大変お世話になっております。",
    "ドクターサポート事務局の" + staff + "でございます。",
    "",
    "この度は、下記日程のご勤務につきましてご相談がありご連絡させていただきました。",
    "ご相談の案が多く大変申し訳ございません。",
    askLine,
    "",
    "━━━━━━━━━━━━━━━━━━━━",
    `■ ご相談日程一覧（計${entries.length}枠）`,
    ...slotLines,
    "━━━━━━━━━━━━━━━━━━━━",
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
 * パターン3：複数日程の移動、または「移動＋個別」「延長＋別枠」など複数枠が混在する場合の統合テンプレート
 * （「ご相談の案が多く大変申し訳ございません。」＋枠ごとに見やすく整理したブロック表示）
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
    const curWage = formatComma_(e.currentDailyWage || 0);
    const conWage = formatComma_(e.consultDailyWage || 0);
    const allowSuffix = (e.approachType === "勤務地変更相談" && e.hasRequestAllow) ? "＋移動依頼手当10,000円" : "";

    if (e.approachType === "勤務地変更相談") {
      const fromC = e.fromClinicName || "現在拠点";
      return [
        `【ご相談${numLabel}】${ymdLabel}（拠点移動のご相談）`,
        `・現在：【${fromC}】${curTime}（日給:${curWage}円）`,
        `・ご相談：【${e.clinicName}】${conTime}（日給:${conWage}円${allowSuffix}）`
      ].join("\n");
    } else if (e.approachType === "延長相談" || e.approachType === "早出相談") {
      const typeLabel = e.approachType === "延長相談" ? "時間延長のご相談" : "前倒し（早出）のご相談";
      return [
        `【ご相談${numLabel}】${ymdLabel}（${typeLabel}）`,
        `・現在：【${e.clinicName}】${curTime}（日給:${curWage}円）`,
        `・ご相談：【${e.clinicName}】${conTime}（日給:${conWage}円）`
      ].join("\n");
    } else {
      const absTime = `${e.absenceStart || ""}~${e.absenceEnd || ""}`;
      return [
        `【ご相談${numLabel}】${ymdLabel}（追加勤務のご相談）`,
        `・ご相談：【${e.clinicName}】${absTime}（日給:${conWage}円）`
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
    "もし先生のご都合がよろしければ、下記のいずれかの枠（または複数枠）にてご協力をいただくことは可能でしょうか。",
    "",
    "━━━━━━━━━━━━━━━━━━━━",
    `■ ご相談内容一覧（計${entries.length}件）`,
    "",
    blocks.join("\n\n"),
    "━━━━━━━━━━━━━━━━━━━━",
    "",
    "※上記募集は公開中のため、ご返信いただいた際に充足している可能性もございます。",
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
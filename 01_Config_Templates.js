/**
 * ============================================================================
 * 01_Config_Templates.gs
 * 外部参照シートID定義・『担当者』シート（A〜D列）参照・8章送信文案テンプレート
 * ============================================================================
 */

var EXTERNAL_IDS = {
  D6_D3_UNFILLED_SS: "1cbeXWojsxNMhQUo1c6VflF5hLUJUyfuOXCFbGP5jJEA",
  D4_D5_WAGE_SS:     "14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs",
  D9_ALLOWANCE_SS:   "1Cw2iHRglJdUf4aeCaE04DhGVBMPpyQdhpiC8U_OdVtl",
  MENTION_LIST_SS:   "1Q1bsHIUysw27aaPD-wTOAKWBbN7scD9RQIIdERYrylg",
  KINMU_MONTHLY_1:   "1JhLJuxkp6T5hEkfultUE57pmRXBI_lkN6ROp2-CuhCc",
  KINMU_MONTHLY_2:   "10Z8jg4o7Ri9Ggf7u_jauyDDZa6bVnc-iuuuqmkNDafw",
  KINMU_SHIFT_NEW:   "1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co",
  KINMU_AGENCY:      "1Fd8uOCE1SKvLCIPjZZ7sE2rFsQFOhjVjHqoaqs-pXqE"
};

var MENTION_LIST_GID = 2012574647;
var MENTION_TARGET_HEADERS = ["シフト戦略チーム", "アルバイトスタッフ"];
var VACCINE_DEDICATED_HOURLY_WAGE = 13000;

var SHEET_NAMES = {
  D6_APPROACH:    "アプローチシート",
  D3_PASTED:      "貼付用",
  D5_CLINICS:     "拠点名",
  D7_REVIEW:      "文案・Slack投稿レビュー",
  D7_CONFIRM:     "確認リスト",
  D7_DICT:        "拠点名・誤変換辞書",
  D7_STAFF:       "担当者",
  D7_NEGOTIATION: "交渉注意医師",
  D7_EXCLUDED:    "除外医師",
  D7_HISTORY:     "打診履歴",
  D7_HOLIDAY:     "祝日 (内閣府)",
  D7_LOG:         "取込・投稿ログ"
};

var CAO_HOLIDAY_CSV_URL = "https://" + "www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv";
var SLACK_CHANNEL_ID = "C0AJTMLF0UD";
var SLACK_NOTIFY_CHANNEL_ID = "C09TRHGU64D";
var CHATWORK_NOTIFY_ROOM_ID = "415381235";
var D7_REVIEW_SHEET_URL = "https://" + "docs.google.com/spreadsheets/d/1WNT1EWFP7lNzfO5hseFyfiw3BjK9SG42XrTFKMxz4kc/edit?hl=ja&gid=1198962830#gid=1198962830";

// 『担当者』シートのキャッシュ（A〜D列の呼び名 -> D列「プルダウン名（名字）」）
var STAFF_ALIAS_MAP = new Map();
var STAFF_RECORDS_LIST = [];

/**
 * D7内の『担当者』シート（A列:担当者氏名, B列:呼び名1, C列:呼び名2, D列:プルダウン名）を読み込む
 * - フルネームが指定された場合でも、必ずD列「プルダウン名（名字）」に紐付ける
 */
function loadStaffMasterFromD7_(d7Ss) {
  STAFF_ALIAS_MAP = new Map();
  STAFF_RECORDS_LIST = [];
  const dropdownNames = [];

  const registerAlias = (aliasStr, dropdownName) => {
    const clean = String(aliasStr || "").replace(/[\s\u3000]+/g, "").trim();
    if (!clean) return;
    STAFF_ALIAS_MAP.set(clean, dropdownName);
    const stem = clean.replace(/(さん|氏|ちゃん|くん|先生)$/g, "");
    if (stem) {
      STAFF_ALIAS_MAP.set(stem, dropdownName);
    }
  };

  try {
    const targetSs = d7Ss || SpreadsheetApp.getActiveSpreadsheet();
    const sheet = targetSs ? targetSs.getSheetByName(SHEET_NAMES.D7_STAFF) : null;
    if (sheet && sheet.getLastRow() >= 2) {
      const vals = sheet.getDataRange().getDisplayValues();
      for (let r = 1; r < vals.length; r++) {
        const colA = String(vals[r][0] || "").trim(); // 担当者氏名（フルネーム等）
        const colB = String(vals[r][1] || "").trim(); // 呼び名1
        const colC = String(vals[r][2] || "").trim(); // 呼び名2
        const colD = String(vals[r][3] || "").trim(); // プルダウン名（名字）
        if (!colA && !colD) continue;

        const surnameGuess = colA.split(/[\s\u3000]+/)[0].trim();
        const dropdownName = colD || surnameGuess || colA.replace(/[\s\u3000]+/g, "");
        if (!dropdownNames.includes(dropdownName)) {
          dropdownNames.push(dropdownName);
        }

        const rowAliases = [colA, surnameGuess, colB, colC, colD, `${dropdownName}さん`].filter(Boolean);
        rowAliases.forEach(a => registerAlias(a, dropdownName));

        STAFF_RECORDS_LIST.push({
          dropdownName,
          fullNameClean: colA.replace(/[\s\u3000]+/g, ""),
          aliases: rowAliases.map(a => a.replace(/[\s\u3000]+/g, "")).filter(Boolean)
        });
      }
    }
  } catch (e) {
    console.log("『担当者』シートの読込エラー: " + e.message);
  }

  return dropdownNames;
}

/**
 * 入力文字列（フルネーム・呼び名・名字）を『担当者』シートのD列「プルダウン名（名字）」に変換する
 */
function resolveStaffByAlias_(inputStr) {
  if (!inputStr) return "";
  if (STAFF_ALIAS_MAP.size === 0) {
    loadStaffMasterFromD7_(SpreadsheetApp.getActiveSpreadsheet());
  }

  const clean = String(inputStr)
    .replace(/\[To:\d+\]/gi, "")
    .replace(/[\s\u3000]+/g, "")
    .trim();
  if (!clean) return "";

  if (STAFF_ALIAS_MAP.has(clean)) return STAFF_ALIAS_MAP.get(clean);
  const withoutHonorific = clean.replace(/(さん|氏|ちゃん|くん|先生)$/g, "");
  if (STAFF_ALIAS_MAP.has(withoutHonorific)) return STAFF_ALIAS_MAP.get(withoutHonorific);

  for (const rec of STAFF_RECORDS_LIST) {
    if (!rec.dropdownName) continue;
    if (
      withoutHonorific.startsWith(rec.dropdownName) ||
      (rec.fullNameClean && rec.fullNameClean.startsWith(withoutHonorific))
    ) {
      return rec.dropdownName;
    }
  }

  return "";
}

/**
 * MTG発言（evidence）の中に『担当者』シートの呼び名やフルネームが含まれていればD列「プルダウン名（名字）」を返す
 */
function detectStaffInEvidence_(evidenceText) {
  if (STAFF_RECORDS_LIST.length === 0) {
    loadStaffMasterFromD7_(SpreadsheetApp.getActiveSpreadsheet());
  }
  const ev = String(evidenceText || "").replace(/[\s\u3000]+/g, "");
  if (!ev) return "";

  for (const rec of STAFF_RECORDS_LIST) {
    for (const alias of rec.aliases) {
      if (alias === rec.dropdownName) continue;
      if (alias.length >= 2 && ev.includes(alias)) {
        return rec.dropdownName;
      }
    }
  }
  return "";
}

function extractCleanStaffName_(rawText) {
  if (!rawText) return "";
  const resolved = resolveStaffByAlias_(rawText);
  if (resolved) return resolved;

  const firstPart = String(rawText || "")
    .replace(/\[To:\d+\]/gi, "")
    .replace(/[\s\u3000]*さん$/g, "")
    .trim()
    .split(/[\s\u3000]+/)[0]
    .trim();

  return resolveStaffByAlias_(firstPart) || firstPart;
}

/**
 * 送信文案の「1行目（〇〇先生）直後の空行削除」と「末尾署名の全角スペース（ドクターサポート事務局 〇〇）」を強制整形する
 */
function normalizeDraftMessageFormat_(draftText, staffNameRaw) {
  if (!draftText) return "";
  const cleanStaff = extractCleanStaffName_(staffNameRaw);
  let text = String(draftText);

  // 1. 先頭「〇〇先生」の直後に空行があれば詰める
  text = text.replace(/^([^\r\n]*?先生)[ \t]*\r?\n(?:\s*\r?\n)+/, "$1\n");

  // 2. 文中の担当者名をプルダウン名（名字）に統一し、末尾署名に全角スペース（\u3000）を入れる
  if (cleanStaff) {
    text = text
      .replace(/ドクターサポート(事務局|部)の.*?でございます。/g, `ドクターサポート$1の${cleanStaff}でございます。`)
      .replace(/ドクターサポート(事務局|部)の.*?です。/g, `ドクターサポート$1の${cleanStaff}です。`)
      .replace(/ドクターサポート(事務局|部)[\s\u3000]*[^\r\n]+$/g, `ドクターサポート$1\u3000${cleanStaff}`);
  } else {
    text = text.replace(/ドクターサポート(事務局|部)[\s\u3000]*([^\r\n]+)$/g, "ドクターサポート$1\u3000$2");
  }

  return text;
}

function getCircledNumber_(num) {
  const circled = ["①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩"];
  return circled[num - 1] || `(${num})`;
}

/**
 * 同一医師に対する複数日・または同日複数拠点の移動打診を1通のメッセージに統合する
 */
function buildMultiDateTemplate_(entries) {
  if (!entries || entries.length === 0) return "";
  if (entries.length === 1) return buildMessageTemplate_(entries[0]);

  const first = entries[0];
  const docHonorific = `${first.doctorName}先生`;
  const staffLabel = extractCleanStaffName_(first.staffName) || "{担当者名}";

  // 日付ごとにグループ化（同日に複数拠点が提案されている場合は1つの日付ブロック内で ご相談①, ② として並記する）
  const dateGroupMap = new Map();
  for (const e of entries) {
    if (!dateGroupMap.has(e.dateStr)) {
      dateGroupMap.set(e.dateStr, []);
    }
    dateGroupMap.get(e.dateStr).push(e);
  }
  const dateGroups = Array.from(dateGroupMap.values());
  const uniqueDateLabels = dateGroups.map(g => (g[0].isTomorrow ? "明日" : "") + formatMdDay_(g[0].dateObj)).join("、");

  if (first.isFullTime) {
    const blocks = dateGroups.map(group => {
      const gFirst = group[0];
      const ymd = formatYmdDay_(gFirst.dateObj);
      const fromC = gFirst.approachType === "勤務地変更相談" ? gFirst.fromClinicName : (gFirst.currentClinicLabel || gFirst.clinicName);
      const consultLines = group.map((e, idx) => {
        const prefix = group.length > 1 ? `ご相談${getCircledNumber_(idx + 1)})` : "ご相談)";
        return `${prefix} 【${e.clinicName}】${e.consultStart}~${e.consultEnd}`;
      });
      return [
        `・${ymd}`,
        `現在)【${fromC}】${gFirst.currentStart}~${gFirst.currentEnd}`,
        ...consultLines
      ].join("\n");
    }).join("\n\n");

    return [
      docHonorific,
      "お疲れ様です。",
      `ドクターサポート部の${staffLabel}です。`,
      "",
      `この度は、${uniqueDateLabels}のご勤務につきましてご相談がありご連絡させていただきました。`,
      "医師不在となる可能性がある拠点があり、大変恐縮ではございますがもし先生のご都合がよろしければ、下記のご勤務変更・ご調整をいただくことは可能でしょうか。",
      "",
      blocks,
      "",
      "急なお願いとなり大変恐れ入りますが、ご無理のない範囲でご検討いただけますと幸いです。",
      "お忙しいところ、ご負担をおかけして恐縮ではございますが、何卒よろしくお願い申し上げます。",
      "",
      `ドクターサポート部\u3000${staffLabel}`
    ].join("\n");
  }

  if (first.shiftType === "応募") {
    const m = first.dateObj.getMonth() + 1;
    const countPrefix = dateGroups.length === 1 ? "当日" : (dateGroups.length === 2 ? "両日とも" : "いずれも");
    const blocks = dateGroups.map(group => {
      const gFirst = group[0];
      const ymd = formatYmdDay_(gFirst.dateObj);
      const curW = formatComma_(gFirst.currentDailyWage);
      const consultLines = group.map((e, idx) => {
        const prefix = group.length > 1 ? `ご相談${getCircledNumber_(idx + 1)})` : "ご相談)";
        const conW = formatComma_(e.consultDailyWage);
        return `${prefix} 【${e.clinicName}】${e.consultStart}~${e.consultEnd} 日給${conW}円`;
      });
      return [
        `■${ymd}`,
        `現在)【${gFirst.fromClinicName}】${gFirst.currentStart}~${gFirst.currentEnd} 日給${curW}円`,
        ...consultLines
      ].join("\n");
    }).join("\n\n");

    return [
      docHonorific,
      "平素よりお世話になっております。",
      `ドクターサポート事務局の${staffLabel}でございます。`,
      "",
      `このたびは${m}月に多くのご応募いただき、誠にありがとうございます。`,
      `恐れ入りますが、${uniqueDateLabels}のご応募につきまして、ご相談差し上げたくご連絡いたしました。`,
      `${countPrefix}医師不在の拠点が多く、もし先生のご都合がよろしければ、ご勤務地を変更いただくことは可能でしょうか。`,
      "その際は時給を高時給の拠点に合わせてご案内させていただきます。",
      "",
      blocks,
      "",
      "難しい場合は現在の拠点にて採用させていただきますので、お手数をおかけいたしますが、ご検討いただけましたら幸いです。",
      "何卒よろしくお願い申し上げます。",
      "",
      `ドクターサポート事務局\u3000${staffLabel}`
    ].join("\n");
  }

  const uniqueFrom = Array.from(new Set(entries.map(e => e.fromClinicName)));
  const uniqueTo = Array.from(new Set(entries.map(e => e.clinicName)));
  let routeSentence = "当日のご勤務拠点を下記の通りご変更いただくことは可能でしょうか。";
  if (uniqueFrom.length === 1 && uniqueTo.length === 1) {
    routeSentence = `当日のご勤務拠点を 【${uniqueFrom[0]}】 より 【${uniqueTo[0]}】 へご変更いただくことは可能でしょうか。`;
  } else if (dateGroups.length === 1 && uniqueFrom.length === 1 && uniqueTo.length > 1) {
    const toListStr = uniqueTo.map(c => `【${c}】`).join(" または ");
    routeSentence = `当日のご勤務拠点を 【${uniqueFrom[0]}】 より ${toListStr} へご変更いただくことは可能でしょうか。`;
  }

  const blocks = dateGroups.map(group => {
    const gFirst = group[0];
    const md = formatMdDay_(gFirst.dateObj);
    const curW = formatComma_(gFirst.currentDailyWage);
    const fromC = gFirst.approachType === "勤務地変更相談" ? gFirst.fromClinicName : (gFirst.currentClinicLabel || gFirst.clinicName);
    const consultLines = group.map((e, idx) => {
      const prefix = group.length > 1 ? `ご相談${getCircledNumber_(idx + 1)})` : "ご相談)";
      const conW = formatComma_(e.consultDailyWage);
      const allowSuffix = e.hasRequestAllow ? "+依頼手当10,000円" : "";
      return `${prefix} 【${e.clinicName}】${e.consultStart}~${e.consultEnd}(日給:${conW}円${allowSuffix})`;
    });
    return [
      `■${md}`,
      `現在)【${fromC}】 ${gFirst.currentStart}~${gFirst.currentEnd} (日給:${curW}円)`,
      "↓",
      ...consultLines
    ].join("\n");
  }).join("\n\n");

  return [
    docHonorific,
    "平素よりお世話になっております。",
    `ドクターサポート事務局の${staffLabel}でございます。`,
    "",
    `ご勤務のご相談が重なり恐縮ですが、${uniqueDateLabels}のご勤務につきましてご相談でございます。`,
    routeSentence,
    "",
    blocks,
    "",
    "※上記募集は公開中のため、ご応募いただいた際に充足している可能性もございます。",
    "恐れ入りますが、充足の際はご容赦くださいませ。",
    "",
    "お忙しいところ、ご負担をおかけして恐縮ではございますが、ご検討いただけましたら幸いです。",
    "何卒よろしくお願い申し上げます。",
    "",
    `ドクターサポート事務局\u3000${staffLabel}`
  ].join("\n");
}

function buildMessageTemplate_(params) {
  const {
    isFullTime, doctorName, staffName, approachType, shiftType,
    dateObj, isTomorrow, clinicName, currentClinicLabel, fromClinicName,
    absenceStart, absenceEnd, currentStart, currentEnd,
    consultStart, consultEnd, consultNoteSuffix, diffHoursText,
    isMorningToNightExt, consultPm1DailyWage,
    currentDailyWage, consultDailyWage, hasRequestAllow, isPublicRecruit,
    multiTargets
  } = params;

  const docHonorific = `${doctorName || "{医師名}"}先生`;
  const staffLabel = extractCleanStaffName_(staffName) || "{担当者名}";
  const mdDay = formatMdDay_(dateObj);
  const ymdDay = formatYmdDay_(dateObj);
  const tomorrowPrefix = isTomorrow ? "明日" : "";
  const curWageStr = formatComma_(currentDailyWage);
  const pm1WageStr = formatComma_(consultPm1DailyWage);
  const conWageStr = formatComma_(consultDailyWage);
  const curClinicDisp = currentClinicLabel || clinicName;
  const noteSuffix = consultNoteSuffix ? ` ${consultNoteSuffix}` : "";
  const pm1NoteSuffix = noteSuffix ? noteSuffix.replace(/~\d{2}:\d{2}/, "~18:00") : "";

  if (isFullTime) {
    if (approachType === "延長相談" && isMorningToNightExt) {
      return [
        docHonorific,
        "お疲れ様です。",
        `ドクターサポート部の${staffLabel}です。`,
        "",
        `この度は、${tomorrowPrefix}${mdDay}のご勤務につきましてご相談がありご連絡させていただきました。`,
        `現在、【${clinicName}】${absenceStart}~${absenceEnd}の時間帯に医師不在となる可能性があるため、`,
        `長時間のご勤務となり大変恐縮ではございますが、もし先生のご都合がよろしければ、可能な限りご勤務時間を延長していただくことは可能でしょうか。`,
        `18:00まで、あるいは${consultEnd}までなど、可能な範囲でご検討いただけますと幸いです。`,
        "",
        `・${ymdDay}`,
        `現在)【${curClinicDisp}】${currentStart}~${currentEnd}`,
        `ご相談①) 【${clinicName}】${currentStart}~18:00${pm1NoteSuffix}`,
        `ご相談②) 【${clinicName}】${currentStart}~${consultEnd}${noteSuffix}`,
        "",
        "長時間にわたり大きなご負担をおかけするお願いとなり大変心苦しい限りではございますが、ご無理のない範囲でご検討いただけますと幸いです。",
        "お忙しいところ誠に恐縮ではございますが、何卒よろしくお願い申し上げます。",
        "",
        `ドクターサポート部\u3000${staffLabel}`
      ].join("\n");
    }

    let consultPhrase = "ご勤務いただくことは可能でしょうか。";
    let hasCurrentRow = false;
    let curLabelForFullTime = curClinicDisp;
    let targetStartForFullTime = consultStart || absenceStart;
    let targetEndForFullTime = consultEnd || absenceEnd;

    if (approachType === "延長相談") {
      consultPhrase = `ご勤務時間を${diffHoursText}延長していただくことは可能でしょうか。`;
      hasCurrentRow = true;
    } else if (approachType === "早出相談") {
      consultPhrase = `ご勤務時間を${diffHoursText}前倒ししていただくことは可能でしょうか。`;
      hasCurrentRow = true;
    } else if (approachType === "勤務地変更相談") {
      consultPhrase = `ご勤務拠点を 【${fromClinicName}】 より 【${clinicName}】 へご変更いただくことは可能でしょうか。`;
      hasCurrentRow = true;
      curLabelForFullTime = fromClinicName;
    } else {
      targetStartForFullTime = absenceStart;
      targetEndForFullTime = absenceEnd;
    }

    const lines = [
      docHonorific,
      "お疲れ様です。",
      `ドクターサポート部の${staffLabel}です。`,
      "",
      `この度は、${tomorrowPrefix}${mdDay}のご勤務につきましてご相談がありご連絡させていただきました。`,
      `現在、【${clinicName}】${absenceStart}~${absenceEnd}の時間帯に医師不在となる可能性があるため、`,
      `大変恐縮ではございますがもし先生のご都合がよろしければ、${consultPhrase}`,
      "",
      `・${ymdDay}`
    ];
    if (hasCurrentRow) lines.push(`現在)【${curLabelForFullTime}】${currentStart}~${currentEnd}`);
    lines.push(`ご相談) 【${clinicName}】${targetStartForFullTime}~${targetEndForFullTime}${noteSuffix}`);
    lines.push("");
    lines.push("急なお願いとなり大変恐れ入りますが、ご無理のない範囲でご検討いただけますと幸いです。");
    lines.push("お忙しいところ、ご負担をおかけして恐縮ではございますが、何卒よろしくお願い申し上げます。");
    lines.push("");
    lines.push(`ドクターサポート部\u3000${staffLabel}`);
    return lines.join("\n");
  }

  if (approachType === "早出相談") {
    return [
      docHonorific,
      "平素よりお世話になっております。",
      `ドクターサポート事務局の${staffLabel}でございます。`,
      "",
      `この度は、${tomorrowPrefix}${mdDay}のご勤務につきましてご相談がありご連絡させていただきました。`,
      `現在、【${clinicName}】${absenceStart}~${absenceEnd}の時間帯に医師不在となる可能性があるため、`,
      `大変恐縮ではございますがもし先生のご都合がよろしければ、ご勤務時間を${diffHoursText}前倒ししていただくことは可能でしょうか。`,
      "",
      `・${ymdDay}`,
      `現在)【${curClinicDisp}】${currentStart}~${currentEnd}`,
      `日給:${curWageStr}円`,
      `ご相談) 【${clinicName}】${consultStart}~${currentEnd}${noteSuffix}`,
      `日給:${conWageStr}円`,
      "",
      "急なお願いとなり大変恐れ入りますが、ご無理のない範囲でご検討いただけますと幸いです。",
      "お忙しいところ、ご負担をおかけして恐縮ではございますが、何卒よろしくお願い申し上げます。",
      "",
      `ドクターサポート事務局\u3000${staffLabel}`
    ].join("\n");
  }

  if (approachType === "延長相談") {
    if (isMorningToNightExt) {
      return [
        docHonorific,
        "平素よりお世話になっております。",
        `ドクターサポート事務局の${staffLabel}でございます。`,
        "",
        `この度は、${tomorrowPrefix}${mdDay}のご勤務につきましてご相談がありご連絡させていただきました。`,
        `現在、【${clinicName}】${absenceStart}~${absenceEnd}の時間帯に医師不在となる可能性があるため、`,
        "長時間のご勤務となり大変恐縮ではございますが、もし先生のご都合がよろしければ、可能な限りご勤務時間を延長していただくことは可能でしょうか。",
        `18:00まで、あるいは${consultEnd}までなど、可能な範囲でご検討いただけますと幸いです。`,
        "",
        `・${ymdDay}`,
        `現在)【${curClinicDisp}】${currentStart}~${currentEnd}`,
        `日給:${curWageStr}円`,
        "",
        `ご相談①) 【${clinicName}】${currentStart}~18:00${pm1NoteSuffix}`,
        `日給:${pm1WageStr}円`,
        `ご相談②) 【${clinicName}】${currentStart}~${consultEnd}${noteSuffix}`,
        `日給:${conWageStr}円`,
        "",
        "長時間にわたり大きなご負担をおかけするお願いとなり大変心苦しい限りではございますが、ご無理のない範囲でご検討いただけますと幸いです。",
        "お忙しいところ誠に恐縮ではございますが、何卒よろしくお願い申し上げます。",
        "",
        `ドクターサポート事務局\u3000${staffLabel}`
      ].join("\n");
    }

    return [
      docHonorific,
      "平素よりお世話になっております。",
      `ドクターサポート事務局の${staffLabel}でございます。`,
      "",
      `この度は、${tomorrowPrefix}${mdDay}のご勤務につきましてご相談がありご連絡させていただきました。`,
      `現在、【${clinicName}】${absenceStart}~${absenceEnd}の時間帯に医師不在となる可能性があるため、`,
      `大変恐縮ではございますがもし先生のご都合がよろしければ、ご勤務時間を${diffHoursText}延長していただくことは可能でしょうか。`,
      "",
      `・${ymdDay}`,
      `現在)【${curClinicDisp}】${currentStart}~${currentEnd}`,
      `日給:${curWageStr}円`,
      `ご相談) 【${clinicName}】${currentStart}~${consultEnd}${noteSuffix}`,
      `日給:${conWageStr}円`,
      "",
      "急なお願いとなり大変恐れ入りますが、ご無理のない範囲でご検討いただけますと幸いです。",
      "お忙しいところ、ご負担をおかけして恐縮ではございますが、何卒よろしくお願い申し上げます。",
      "",
      `ドクターサポート事務局\u3000${staffLabel}`
    ].join("\n");
  }

  if (approachType === "勤務地変更相談") {
    if (shiftType === "応募") {
      const m = dateObj.getMonth() + 1;
      return [
        docHonorific,
        "平素よりお世話になっております。",
        `ドクターサポート事務局の${staffLabel}でございます。`,
        "",
        `このたびは${m}月に多くのご応募いただき、誠にありがとうございます。`,
        `恐れ入りますが、${mdDay}のご応募につきまして、ご相談差し上げたくご連絡いたしました。`,
        "医師不在の拠点が多く、もし先生のご都合がよろしければ、ご勤務地を変更いただくことは可能でしょうか。",
        "その際は時給を高時給の拠点に合わせてご案内させていただきます。",
        "",
        `■${ymdDay}`,
        `現在)【${fromClinicName}】${currentStart}~${currentEnd} 日給${curWageStr}円`,
        `ご相談) 【${clinicName}】${consultStart}~${consultEnd} 日給${conWageStr}円`,
        "",
        "難しい場合は現在の拠点にて採用させていただきますので、お手数をおかけいたしますが、ご検討いただけましたら幸いです。",
        "何卒よろしくお願い申し上げます。",
        "",
        `ドクターサポート事務局\u3000${staffLabel}`
      ].join("\n");
    }

    const allowSuffix = hasRequestAllow ? "+依頼手当10,000円" : "";

    // 単一行で複数候補拠点（multiTargets）が渡された場合の並記出力
    if (Array.isArray(multiTargets) && multiTargets.length > 1) {
      const toNames = multiTargets.map(t => `【${t.clinicName}】`).join(" または ");
      const consultLines = multiTargets.map((t, idx) => {
        const cW = formatComma_(t.consultDailyWage);
        const aSuf = t.hasRequestAllow ? "+依頼手当10,000円" : allowSuffix;
        return `ご相談${getCircledNumber_(idx + 1)}) 【${t.clinicName}】${t.consultStart || consultStart}~${t.consultEnd || consultEnd}(日給:${cW}円${aSuf})`;
      });
      return [
        docHonorific,
        "平素よりお世話になっております。",
        `ドクターサポート事務局の${staffLabel}でございます。`,
        "",
        `ご勤務のご相談が重なり恐縮ですが、${tomorrowPrefix}${mdDay}のご勤務につきましてご相談でございます。`,
        `当日のご勤務拠点を 【${fromClinicName}】 より ${toNames} へご変更いただくことは可能でしょうか。`,
        "",
        `■${mdDay}`,
        `現在)【${fromClinicName}】 ${currentStart}~${currentEnd} (日給:${curWageStr}円)`,
        "↓",
        ...consultLines,
        "",
        "※上記募集は公開中のため、ご応募いただいた際に充足している可能性もございます。",
        "恐れ入りますが、充足の際はご容赦くださいませ。",
        "",
        "お忙しいところ、ご負担をおかけして恐縮ではございますが、ご検討いただけましたら幸いです。",
        "何卒よろしくお願い申し上げます。",
        "",
        `ドクターサポート事務局\u3000${staffLabel}`
      ].join("\n");
    }

    return [
      docHonorific,
      "平素よりお世話になっております。",
      `ドクターサポート事務局の${staffLabel}でございます。`,
      "",
      `ご勤務のご相談が重なり恐縮ですが、${tomorrowPrefix}${mdDay}のご勤務につきましてご相談でございます。`,
      `当日のご勤務拠点を 【${fromClinicName}】 より 【${clinicName}】 へご変更いただくことは可能でしょうか。`,
      "",
      `■${mdDay}`,
      `現在)【${fromClinicName}】 ${currentStart}~${currentEnd} (日給:${curWageStr}円)`,
      "↓",
      `ご相談) 【${clinicName}】${consultStart}~${consultEnd}(日給:${conWageStr}円${allowSuffix})`,
      "",
      "※上記募集は公開中のため、ご応募いただいた際に充足している可能性もございます。",
      "恐れ入りますが、充足の際はご容赦くださいませ。",
      "",
      "お忙しいところ、ご負担をおかけして恐縮ではございますが、ご検討いただけましたら幸いです。",
      "何卒よろしくお願い申し上げます。",
      "",
      `ドクターサポート事務局\u3000${staffLabel}`
    ].join("\n");
  }

  const lines = [
    docHonorific,
    "平素よりお世話になっております。",
    `ドクターサポート事務局の${staffLabel}でございます。`,
    "",
    `この度は、${tomorrowPrefix}${mdDay}のご勤務につきましてご相談がありご連絡させていただきました。`,
    `現在、【${clinicName}】${absenceStart}~${absenceEnd}の時間帯に医師不在となる可能性があるため、`,
    "大変恐縮ではございますがもし先生のご都合がよろしければ、ご勤務いただくことは可能でしょうか。",
    "",
    `・${ymdDay}`,
    `ご相談) 【${clinicName}】${absenceStart}~${absenceEnd}`,
    `日給:${conWageStr}円`
  ];
  if (isPublicRecruit !== false) {
    lines.push("");
    lines.push("※上記募集は公開中のため、ご応募いただいた際に充足している可能性もございます。");
    lines.push("恐れ入りますが、充足の際はご容赦くださいませ。");
  }
  lines.push("");
  lines.push("急なお願いとなり大変恐れ入りますが、ご無理のない範囲でご検討いただけますと幸いです。");
  lines.push("お忙しいところ、ご負担をおかけして恐縮ではございますが、何卒よろしくお願い申し上げます。");
  lines.push("");
  lines.push(`ドクターサポート事務局\u3000${staffLabel}`);
  return lines.join("\n");
}
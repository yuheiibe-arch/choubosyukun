/**
 * ============================================================================
 * 02b_2_Merge_Consolidator.gs（マージ管理 3分割の2）
 * おまとめ行のうち一部枠のみ充足した際の「単枠／残存枠文案への自動再構成（既存担当者維持）」専用
 * （※新規おまとめ集約本体は 02a_2_Consolidate_Rows.gs に一本化済）
 * ============================================================================
 */

function rebuildSurvivingSubsetRow_(origRow, survivingSlots, ctx) {
  const { d3Data, wageSs, holidaySet } = ctx;
  const rawApproach = String(origRow[7] || "個別アプローチ").split("・")[0].trim() || "個別アプローチ";
  const rawDocText = String(origRow[8] || "").trim();
  const docLines = rawDocText.split(/\r?\n/).map(function(s) { return s.trim(); }).filter(Boolean);
  const singleDocName = (docLines.length === 1 && docLines[0] !== "該当候補なし")
    ? docLines[0].replace(/先生$/g, "").trim()
    : "";
  const staffName = extractCleanStaffName_(origRow[15]);
  const postType = String(origRow[19] || "通常").trim() || "通常";

  const entries = [];
  const slotRowDatas = [];

  for (let i = 0; i < survivingSlots.length; i++) {
    const sl = survivingSlots[i];
    const baseClinic = sl.d6ClinicKey.replace(/[\(（].*?[\)）]/g, "").trim();
    const slotDept = sl.d6ClinicKey.includes("内科") || sl.dept === "内科" ? "内科" : "小児科";

    const docRes = resolveDoctorAndShift_(
      singleDocName,
      rawApproach,
      sl.dateStr,
      baseClinic,
      slotDept,
      "",
      sl.d6Row,
      d3Data,
      sl.absenceTime,
      "",
      ""
    );

    const effApproach = docRes.resolvedApproachType || rawApproach;
    const effFromClinic = docRes.fromClinic || "";
    const wageRes = calculateWagesAndAllowance_({
      wageSs: wageSs,
      dateStr: sl.dateStr,
      clinicName: baseClinic,
      dept: slotDept,
      fromClinicName: effFromClinic,
      approachType: effApproach,
      absenceTime: sl.absenceTime,
      currentStart: docRes.currentStart,
      currentEnd: docRes.currentEnd,
      consultStart: docRes.consultStart,
      consultEnd: docRes.consultEnd,
      isCurrentVaccine: Boolean(docRes.isCurrentVaccine),
      isSpecialFromMtg: sl.isSpecial,
      isSpecialFromD6: sl.d6Row ? sl.d6Row.isSpecialWage : false,
      isFullTime: Boolean(docRes.isFullTime),
      alreadyHasAllowance: Boolean(docRes.alreadyHasAllowance),
      holidaySet: holidaySet
    });

    const absParts = String(sl.absenceTime || "-").split("-");
    const entryParam = {
      rowIndex: i,
      rowState: "未レビュー",
      isDocResolved: (docRes.status === "OK"),
      dateStr: sl.dateStr,
      isFullTime: Boolean(docRes.isFullTime),
      doctorName: singleDocName || "{医師名}",
      staffName: staffName,
      approachType: effApproach,
      shiftType: "確定",
      dateObj: parseDate_(sl.dateStr),
      isTomorrow: isTomorrowDate_(sl.dateStr),
      clinicName: sl.d6ClinicKey,
      currentClinicLabel: docRes.currentClinicLabel || sl.d6ClinicKey,
      fromClinicName: docRes.currentClinicLabel || effFromClinic || "現在拠点",
      absenceStart: absParts[0] || "{不在開始}",
      absenceEnd: absParts[1] || "{不在終了}",
      currentStart: docRes.currentStart || "{現在開始}",
      currentEnd: docRes.currentEnd || "{現在終了}",
      consultStart: docRes.consultStart || absParts[0] || "{相談開始}",
      consultEnd: docRes.consultEnd || absParts[1] || "{相談終了}",
      consultNoteSuffix: docRes.consultNoteSuffix || "",
      diffHoursText: wageRes.diffHoursText || "{時間}",
      isMorningToNightExt: Boolean(wageRes.isMorningToNightExt),
      consultPm1DailyWage: wageRes.consultPm1DailyWage || 0,
      currentDailyWage: wageRes.currentDailyWage || 0,
      consultDailyWage: wageRes.consultDailyWage || 0,
      hasRequestAllow: (wageRes.requestAllowance || 0) > 0,
      isPublicRecruit: true
    };
    entries.push(entryParam);

    const mdShort = formatSlackDateHeader_(sl.dateStr);
    const targetDocLabel = rawDocText || "候補医師";
    const memoPlan = buildNaturalWorkMemo_(effApproach, targetDocLabel, effFromClinic);
    const slackBody = "`" + mdShort + "`\n```\n【" + sl.d6ClinicKey + "】" + sl.absenceTime + (wageRes.isSpecial ? "【特別時給】" : "") + "\n・" + memoPlan + "\n```";

    slotRowDatas.push({
      dateStr: sl.dateStr,
      d6ClinicKey: sl.d6ClinicKey,
      absenceTime: sl.absenceTime,
      dept: slotDept,
      approachType: effApproach,
      currentShiftLabel: docRes.currentShiftLabel || "-",
      consultShiftLabel: docRes.consultShiftLabel || (sl.d6ClinicKey + " " + sl.absenceTime),
      isSpecial: Boolean(wageRes.isSpecial),
      currentDailyWage: wageRes.currentDailyWage || "",
      consultDailyWage: wageRes.consultDailyWage || "",
      requestAllowance: wageRes.requestAllowance || "",
      memoPlan: memoPlan,
      slackBody: slackBody
    });
  }

  // 1枠だけ残った場合 ➔ 単枠用の文案・表示に完全再構成（既存の担当者 staffName も完全維持）
  if (entries.length === 1) {
    const e0 = entries[0];
    const s0 = slotRowDatas[0];
    let singleDraft = (e0.approachType === "募集継続・ステイ" || e0.approachType === "ステイ")
      ? ""
      : buildMessageTemplate_(e0);
    if (singleDraft) {
      singleDraft = normalizeDraftMessageFormat_(singleDraft, staffName);
    }

    return [
      false,
      "未レビュー",
      String(origRow[2] || ""),
      s0.dateStr,
      s0.d6ClinicKey,
      s0.absenceTime,
      s0.dept,
      s0.approachType,
      rawDocText,
      s0.currentShiftLabel,
      s0.consultShiftLabel,
      s0.isSpecial ? "特別時給" : "通常時給",
      s0.currentDailyWage,
      s0.consultDailyWage,
      s0.requestAllowance,
      staffName,
      singleDraft,
      s0.memoPlan,
      buildSlackThreadTitle_(s0.dateStr),
      postType,
      s0.slackBody,
      "※他枠充足により単枠文案へ自動再構成（" + s0.dateStr.slice(5) + " " + s0.d6ClinicKey + "）",
      "",
      ""
    ];
  }

  // 2枠以上残った場合 ➔ 残存枠のみでのおまとめ文案に再構成（既存の担当者 staffName も完全維持）
  const uniqueDates = Array.from(new Set(entries.map(function(e) { return e.dateStr; })));
  const allMovers = entries.every(function(e) { return e.approachType === "勤務地変更相談"; });
  const allIndividuals = entries.every(function(e) { return e.approachType === "個別アプローチ"; });

  let combinedDraft = "";
  if (uniqueDates.length === 1 && allMovers) {
    combinedDraft = buildSameDateMultiClinicMoverTemplate_(entries);
  } else if (allIndividuals) {
    combinedDraft = buildMultiDateIndividualTemplate_(entries);
  } else {
    combinedDraft = buildUnifiedMultiSlotDraftTemplate_(entries);
  }
  if (combinedDraft) {
    combinedDraft = normalizeDraftMessageFormat_(combinedDraft, staffName);
  }

  const curShifts = Array.from(new Set(slotRowDatas.map(function(s) { return s.currentShiftLabel; }).filter(function(v) { return v && v !== "-"; })));
  const wageLabels = slotRowDatas.map(function(s) {
    const w = Number(String(s.consultDailyWage || "0").replace(/,/g, ""));
    return w > 0 ? (formatComma_(w) + "円") : "-";
  });
  let memo0 = slotRowDatas[0].memoPlan;
  if (uniqueDates.length === 1 && allMovers) {
    const allClinicsLabel = entries.map(function(e) { return e.clinicName; }).join("・");
    if (memo0 && !memo0.includes("候補：")) memo0 += "（候補：" + allClinicsLabel + "）";
  }
  const slotSummary = entries.map(function(e) { return e.dateStr.slice(5) + " " + e.clinicName; }).join("・");

  return [
    false,
    "未レビュー",
    String(origRow[2] || ""),
    slotRowDatas.map(function(s) { return s.dateStr; }).join("\n"),
    slotRowDatas.map(function(s) { return s.d6ClinicKey; }).join("\n"),
    slotRowDatas.map(function(s) { return s.absenceTime; }).join("\n"),
    Array.from(new Set(slotRowDatas.map(function(s) { return s.dept; }))).join("・") || "小児科",
    Array.from(new Set(slotRowDatas.map(function(s) { return s.approachType; }))).join("・"),
    rawDocText,
    curShifts.length > 0 ? curShifts.join("\n") : "-",
    slotRowDatas.map(function(s) { return s.dateStr.slice(5) + " " + s.consultShiftLabel; }).join("\n"),
    slotRowDatas.some(function(s) { return s.isSpecial; }) ? "特別時給" : "通常時給",
    slotRowDatas[0].currentDailyWage || "",
    wageLabels.every(function(w) { return w === "-"; }) ? "" : wageLabels.join("\n"),
    slotRowDatas[0].requestAllowance || "",
    staffName,
    combinedDraft,
    memo0,
    buildSlackThreadTitle_(slotRowDatas[0].dateStr),
    postType,
    slotRowDatas.map(function(s) { return s.slackBody; }).join("\n\n"),
    "おまとめ全" + entries.length + "枠（" + slotSummary + " ※一部充足により再構成）",
    "",
    ""
  ];
}
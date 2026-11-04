/**
 * ============================================================================
 * 03b_1_D6D3_Loader.gs
 * D6（アプローチシート）・D3（貼付用シート）の読込 ＆ 拠点名解決・枠結合
 * ============================================================================
 */

function loadD6ApproachData_(sheet) {
  const vals = sheet.getDataRange().getDisplayValues();
  if (vals.length < 3) return [];
  const headers = vals[0];
  const idx = name => headers.indexOf(name);
  const result = [];

  for (let r = 2; r < vals.length; r++) {
    const row = vals[r];
    const dateStr = normalizeDateStr_(row[idx("日付")]);
    if (!dateStr) continue;
    const clinicRaw = String(row[idx("拠点名")] || "").trim();
    const baseClinic = clinicRaw.replace(/[\(（].*?[\)）]/g, "").trim();
    const deptMatch = clinicRaw.match(/[\(（](.*?)[\)）]/);
    const dept = deptMatch ? deptMatch[1].trim() : "小児科";

    result.push({
      rowNum: r + 1,
      rowNumsText: String(r + 1),
      dateStr,
      clinicRaw,
      baseClinic,
      dept,
      absenceTime: normalizeTimeRange_(row[idx("不在時間")]),
      listedWage: Number(String(row[idx("掲載時給")] || "0").replace(/,/g, "")),
      prevDocsRaw: String(row[idx("前の時間枠の医師")] || "").trim(),
      nextDocsRaw: String(row[idx("後ろの時間枠の医師")] || "").trim(),
      staff: String(row[idx("担当")] || "").trim(),
      status: String(row[idx("対応状況")] || "").trim(),
      isSpecialWage: String(row[idx("特別時給")] || "").includes("特別時給"),
      prevExtStatus: String(row[idx("前医延長")] || "").trim(),
      nextEarlyStatus: String(row[idx("後医早出")] || "").trim()
    });
  }
  return result;
}

function loadD3PastedShifts_(sheet) {
  const vals = sheet.getDataRange().getDisplayValues();
  if (vals.length < 2) return [];
  const headers = vals[0];
  const idx = name => headers.indexOf(name);
  const cId = idx("医籍番号");
  const cName = idx("名前");
  const cKana = idx("名前（カタカナ）");
  const cLast = idx("姓");
  const cFirst = idx("名");
  const cLastKana = idx("セイ");
  const cFirstKana = idx("メイ");
  const cClinic = idx("クリニック名");
  const cDept = idx("診療科");
  const cDate = idx("勤務日");
  const cStart = idx("勤務開始時間");
  const cEnd = idx("勤務終了時間");
  const cWageSum = idx("時給合計");

  const rawList = [];
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    const docName = String(row[cName] || "").replace(/\s+/g, "").trim();
    if (!docName || docName.includes("休館")) continue;

    let alreadyHasAllowance = false;
    for (let c = 0; c < headers.length; c++) {
      if (headers[c].includes("追加支給項目") && String(row[c]).includes("依頼手当")) {
        alreadyHasAllowance = true;
      }
    }

    const wageSumStr = cWageSum >= 0 ? String(row[cWageSum] || "").replace(/,/g, "").trim() : "";
    const wageSumVal = wageSumStr !== "" ? Number(wageSumStr) : -1;

    rawList.push({
      docId: cId >= 0 ? String(row[cId] || "").trim() : "",
      docName,
      normName: toStandardKanji_(docName),
      kanaName: String(row[cKana] || "").replace(/\s+/g, "").trim(),
      lastName: String(row[cLast] || "").trim(),
      normLast: toStandardKanji_(String(row[cLast] || "").trim()),
      normFirst: toStandardKanji_(cFirst >= 0 ? String(row[cFirst] || "").trim() : ""),
      lastKana: String(row[cLastKana] || "").trim(),
      firstKana: cFirstKana >= 0 ? String(row[cFirstKana] || "").trim() : "",
      clinicName: String(row[cClinic] || "").trim(),
      dept: String(row[cDept] || "").trim(),
      dateStr: normalizeDateStr_(row[cDate]),
      startTime: padTime_(row[cStart]),
      endTime: padTime_(row[cEnd]),
      isFullTime: wageSumVal === 0,
      alreadyHasAllowance
    });
  }

  const groupedMap = new Map();
  for (const s of rawList) {
    const key = [s.dateStr, s.docName, s.clinicName, s.dept].join("|");
    if (!groupedMap.has(key)) {
      groupedMap.set(key, { ...s });
    } else {
      const existing = groupedMap.get(key);
      if (toMinutes_(s.startTime) < toMinutes_(existing.startTime)) existing.startTime = s.startTime;
      if (toMinutes_(s.endTime) > toMinutes_(existing.endTime)) existing.endTime = s.endTime;
      if (s.alreadyHasAllowance) existing.alreadyHasAllowance = true;
    }
  }
  return Array.from(groupedMap.values());
}

function resolveAmbiguousClinic_(clinicName, rawClinicName, dateStr, targetTime, rawDocName, fromClinic, d6Data, d3Data) {
  if (clinicName && clinicName !== "要確認") {
    return { clinicName, candidatesText: clinicName };
  }
  const raw = String(rawClinicName || "");
  const dictMatched = normalizeClinicByDict_(raw);
  if (dictMatched && dictMatched !== raw && dictMatched !== "要確認") {
    return { clinicName: dictMatched, candidatesText: dictMatched };
  }

  if (raw.includes("相模")) {
    const sagamiCandidates = d6Data.filter(d =>
      d.dateStr === dateStr &&
      (d.baseClinic === "相模原" || d.baseClinic === "相模大野") &&
      (!targetTime || isTimeOverlapping_(d.absenceTime, targetTime))
    );
    const uniqueClinics = Array.from(new Set(sagamiCandidates.map(d => d.baseClinic)));
    if (uniqueClinics.length === 1) {
      return { clinicName: uniqueClinics[0], candidatesText: uniqueClinics[0] };
    }
    return { clinicName: "要確認", candidatesText: "相模原 / 相模大野", note: "相模原・相模大野の候補" };
  }

  if (raw.includes("武蔵") && rawDocName) {
    const docMatches = findDoctorsByScoring_(d3Data, dateStr, rawDocName, "武蔵");
    const musashiShift = docMatches.find(s => s.clinicName.startsWith("武蔵"));
    if (musashiShift) {
      return {
        clinicName: "要確認",
        inferredFromClinic: musashiShift.clinicName,
        candidatesText: `移動元:${musashiShift.clinicName} -> 移動先候補:亀有(小児科) / 海老名`,
        note: `${musashiShift.docName}先生(${musashiShift.clinicName} ${musashiShift.startTime}-${musashiShift.endTime})の移動先(亀有または海老名)を要確認`
      };
    }
  }

  if (fromClinic) {
    const fromDocs = d3Data.filter(s => s.dateStr === dateStr && s.clinicName === fromClinic);
    const docNames = fromDocs.map(s => `${s.docName}(${s.startTime}-${s.endTime})`).join(", ");
    return {
      clinicName: "要確認",
      candidatesText: `移動元:${fromClinic}(${docNames})`,
      note: `移動先「${raw}」が不明（当日${fromClinic}勤務医: ${docNames || "なし"}）`
    };
  }

  return { clinicName: "要確認", candidatesText: "要確認" };
}

function matchD6RowsSmart_(d6Data, dateStr, clinicName, dept, targetTime, isKitaOrKame) {
  if (!clinicName || clinicName === "要確認") return { mergedRow: null };

  const normKey = s => String(s || "").replace(/[ヶケが]/g, "ケ").trim();
  const targetNorm = normKey(clinicName);

  let matches = d6Data.filter(d => {
    if (d.dateStr !== dateStr) return false;
    if (d.baseClinic !== clinicName && normKey(d.baseClinic) !== targetNorm) return false;
    if (isKitaOrKame && dept !== "要確認" && dept !== "ワクチン専任" && d.dept !== dept) return false;
    if (!targetTime) return true;
    return d.absenceTime === targetTime || isTimeOverlapping_(d.absenceTime, targetTime);
  });

  if (matches.length === 0) return { mergedRow: null };
  if (matches.length === 1) return { mergedRow: matches[0] };

  const uniqueDepts = Array.from(new Set(matches.map(m => m.dept)));
  if (isKitaOrKame && uniqueDepts.length > 1 && dept === "要確認") {
    return { mergedRow: null };
  }

  matches.sort((a, b) => toMinutes_(a.absenceTime.split("-")[0]) - toMinutes_(b.absenceTime.split("-")[0]));
  const first = matches[0];
  const last = matches[matches.length - 1];
  const mergedStart = first.absenceTime.split("-")[0];
  const mergedEnd = last.absenceTime.split("-")[1];

  return {
    mergedRow: {
      ...first,
      rowNumsText: matches.map(m => m.rowNum).join("・"),
      absenceTime: `${mergedStart}-${mergedEnd}`,
      prevDocsRaw: matches.map(m => m.prevDocsRaw).filter(Boolean).join(", "),
      nextDocsRaw: matches.map(m => m.nextDocsRaw).filter(Boolean).join(", "),
      isSpecialWage: matches.some(m => m.isSpecialWage)
    }
  };
}
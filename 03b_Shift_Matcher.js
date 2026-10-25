/**
 * ============================================================================
 * 03b_Shift_Matcher.gs
 * D6（アプローチシート）・D3（貼付用シート）の読込 ＆ 当日シフト突合
 * - 移動相談時に発言内で「前倒し・早出・延長・通し」があれば15:00-21:00等へ自動拡大
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

/**
 * MTG発言（evidence）の中に、移動対象の医師本人に対する「前倒し・早出・延長・通し」等の時間拡大依頼が含まれているか判定する
 * ※「川口から移動してもらって川口延長で…」のように移動元拠点の穴埋めとして「(移動元)延長」と言っているケースは除外する
 */
function hasTimeExpansionIntentForMover_(evidenceText, fromClinicName) {
  const ev = String(evidenceText || "");
  if (!ev) return false;

  // 「とりあえず」「まずは」で夜間のみ等に絞っている発言は拡大しない
  if ((ev.includes("とりあえず") || ev.includes("まずは") || ev.includes("一旦")) &&
      !ev.includes("前倒") && !ev.includes("早出") && !ev.includes("通し") && !ev.includes("投資")) {
    return false;
  }

  // 前倒し・早出・通し（音声誤変換の「投資」含む）・フル勤務の明示キーワード
  if (/前倒|早出|通し|投資の勤務|フルで|終日|1日|一日|午後から|15時から|午前から|9時から/.test(ev)) {
    return true;
  }

  // 「延長」が含まれる場合、「(移動元拠点)延長」「(移動元拠点)は延長」でなければ移動先での延長依頼と判定
  if (ev.includes("延長")) {
    if (fromClinicName && (ev.includes(`${fromClinicName}延長`) || ev.includes(`${fromClinicName}は延長`) || ev.includes(`${fromClinicName}の延長`))) {
      return false;
    }
    return true;
  }

  return false;
}

function resolveDoctorAndShift_(rawDocName, approachType, dateStr, clinicName, dept, fromClinic, d6Row, d3Data, effectiveAbsenceTime, evidenceText) {
  const cleanTarget = (rawDocName || "").replace(/\s+/g, "");
  const [absStart, absEnd] = (effectiveAbsenceTime || "-").split("-");
  let resolvedApproachType = approachType;

  let targetDept = dept;
  if (targetDept === "ワクチン専任") {
    const absEndMin = toMinutes_(absEnd);
    const isPmVaccineAlreadyFilled = d3Data.some(s =>
      s.dateStr === dateStr &&
      s.clinicName === clinicName &&
      s.dept.includes("ワクチン") &&
      toMinutes_(s.endTime) >= 18 * 60
    );
    if (d6Row || (absEndMin && absEndMin > 19 * 60) || isPmVaccineAlreadyFilled) {
      targetDept = (d6Row && d6Row.dept) ? d6Row.dept : "小児科";
    }
  }

  if (d6Row && (resolvedApproachType === "延長相談" || resolvedApproachType === "早出相談")) {
    const sourceRaw = resolvedApproachType === "延長相談" ? d6Row.prevDocsRaw : d6Row.nextDocsRaw;
    const candidates = parseD6DoctorCell_(sourceRaw);

    let matched = candidates;
    if (cleanTarget && candidates.length > 0) {
      const scored = candidates.map(c => {
        const d3Doc = d3Data.find(s => s.dateStr === dateStr && toStandardKanji_(s.docName) === toStandardKanji_(c.name));
        const fakeRecord = d3Doc || {
          normName: toStandardKanji_(c.name),
          normLast: toStandardKanji_(c.name.slice(0, 2)),
          normFirst: toStandardKanji_(c.name.slice(2)),
          kanaName: "",
          lastKana: ""
        };
        return { cand: c, score: calcDoctorSimilarityScore_(cleanTarget, fakeRecord) };
      }).filter(x => x.score >= 55);

      scored.sort((a, b) => b.score - a.score);
      if (scored.length === 1 || (scored.length > 1 && scored[0].score > scored[1].score)) {
        matched = [scored[0].cand];
      } else {
        matched = scored.map(x => x.cand);
      }
    }

    if (matched.length === 1) {
      const m = matched[0];
      const d3Doc = d3Data.find(s =>
        s.dateStr === dateStr &&
        toStandardKanji_(s.docName) === toStandardKanji_(m.name) &&
        s.clinicName === clinicName
      ) || d3Data.find(s => s.dateStr === dateStr && toStandardKanji_(s.docName) === toStandardKanji_(m.name));

      const actualStart = d3Doc ? d3Doc.startTime : m.start;
      const actualEnd = d3Doc ? d3Doc.endTime : m.end;
      const isOrigVaccine = d3Doc ? d3Doc.dept.includes("ワクチン") : false;
      const origClinicLabel = isOrigVaccine ? `${clinicName}(ワクチン専任)` : (d6Row ? d6Row.clinicRaw : clinicName);
      const consultStart = resolvedApproachType === "早出相談" && absStart ? absStart : actualStart;
      const consultEnd = resolvedApproachType === "延長相談" && absEnd ? absEnd : actualEnd;

      return {
        status: "OK",
        resolvedApproachType,
        resolvedDept: targetDept,
        doctorName: d3Doc ? d3Doc.docName : m.name,
        fromClinic: clinicName,
        isCurrentVaccine: isOrigVaccine,
        currentClinicLabel: origClinicLabel,
        currentStart: actualStart,
        currentEnd: actualEnd,
        consultStart,
        consultEnd,
        hasTimeExpansion: true,
        consultNoteSuffix: (isOrigVaccine && targetDept !== "ワクチン専任") ? `(※${absStart}~${absEnd}は${targetDept})` : "",
        currentShiftLabel: `${origClinicLabel} ${actualStart}-${actualEnd}`,
        consultShiftLabel: (isOrigVaccine && targetDept !== "ワクチン専任")
          ? `${d6Row ? d6Row.clinicRaw : clinicName} ${consultStart}-${consultEnd} (${absStart}-${absEnd}${targetDept})`
          : `${d6Row ? d6Row.clinicRaw : clinicName} ${consultStart}-${consultEnd}`,
        isFullTime: d3Doc ? d3Doc.isFullTime : false,
        alreadyHasAllowance: d3Doc ? d3Doc.alreadyHasAllowance : false,
        evidenceNote: `D6前後医師欄＆D3シフト一致`
      };
    }
  }

  if (cleanTarget) {
    const clinicHint = (resolvedApproachType === "延長相談" || resolvedApproachType === "早出相談") ? clinicName : fromClinic;
    let sameDayDocs = findDoctorsByScoring_(d3Data, dateStr, cleanTarget, clinicHint, dept);

    const uniqueNames = Array.from(new Set(sameDayDocs.map(d => d.docName)));

    if (uniqueNames.length === 1) {
      const s = sameDayDocs[0];
      const isSameClinic = (s.clinicName === clinicName);
      const isOrigVaccine = s.dept.includes("ワクチン");
      const isTargetVaccine = (targetDept === "ワクチン専任");

      if (isSameClinic) {
        if (absEnd && toMinutes_(s.endTime) <= toMinutes_(absStart || absEnd)) {
          resolvedApproachType = "延長相談";
        } else if (absStart && toMinutes_(s.startTime) >= toMinutes_(absEnd || absStart)) {
          resolvedApproachType = "早出相談";
        }
      }

      let cStart = s.startTime;
      let cEnd = s.endTime;
      const hasExpansion = hasTimeExpansionIntentForMover_(evidenceText, s.clinicName);

      if (resolvedApproachType === "早出相談" && absStart) {
        cStart = absStart;
      } else if (resolvedApproachType === "延長相談" && absEnd) {
        cEnd = absEnd;
      } else if (resolvedApproachType === "勤務地変更相談" && absStart && absEnd) {
        // ★発言内に「前倒し」「早出」「延長」「通し」等の時間拡大依頼がある場合（またはワクチン専任から小児科へ移る場合）は、
        // 移動先の不在時間（例：15:00-21:00 や 09:00-21:00）へ相談シフトを拡大する！
        if (absStart === "09:00" && absEnd === "21:00") {
          cStart = absStart;
          cEnd = absEnd;
        } else if (isOrigVaccine && !isTargetVaccine) {
          cStart = absStart;
          cEnd = absEnd;
        } else if (hasExpansion) {
          cStart = toMinutes_(absStart) < toMinutes_(s.startTime) ? absStart : s.startTime;
          cEnd = toMinutes_(absEnd) > toMinutes_(s.endTime) ? absEnd : s.endTime;
        }
      }

      const originLabel = isOrigVaccine ? `${s.clinicName}(ワクチン専任)` : s.clinicName;
      const targetClinicDisp = d6Row ? d6Row.clinicRaw : (isTargetVaccine ? `${clinicName}(ワクチン専任)` : clinicName);
      const isVaccineToClinicExt = (resolvedApproachType === "延長相談" || resolvedApproachType === "早出相談") && isOrigVaccine && !isTargetVaccine;

      return {
        status: "OK",
        resolvedApproachType,
        resolvedDept: targetDept,
        doctorName: s.docName,
        fromClinic: s.clinicName,
        isCurrentVaccine: isOrigVaccine,
        currentClinicLabel: originLabel,
        currentStart: s.startTime,
        currentEnd: s.endTime,
        consultStart: cStart,
        consultEnd: cEnd,
        hasTimeExpansion: hasExpansion || (cStart !== s.startTime || cEnd !== s.endTime),
        consultNoteSuffix: isVaccineToClinicExt ? `(※${absStart}~${absEnd}は${targetDept})` : "",
        currentShiftLabel: `${originLabel} ${s.startTime}-${s.endTime}`,
        consultShiftLabel: isVaccineToClinicExt
          ? `${targetClinicDisp}${cStart}-${cEnd} (${absStart}-${absEnd}${targetDept})`
          : `${targetClinicDisp} ${cStart}-${cEnd}`,
        isFullTime: s.isFullTime,
        alreadyHasAllowance: s.alreadyHasAllowance,
        evidenceNote: `D3自動スコアリング一致(${originLabel})`
      };
    } else if (uniqueNames.length > 1) {
      return {
        status: "要確認",
        resolvedApproachType,
        resolvedDept: targetDept,
        reason: `同姓・複数候補あり(${uniqueNames.join(", ")})`,
        doctorName: rawDocName,
        currentShiftLabel: "複数候補",
        consultShiftLabel: ""
      };
    }
  }

  if (resolvedApproachType === "募集継続・ステイ" || resolvedApproachType === "個別アプローチ") {
    return {
      status: "OK",
      resolvedApproachType,
      resolvedDept: targetDept,
      doctorName: "",
      currentShiftLabel: "-",
      consultShiftLabel: resolvedApproachType === "個別アプローチ" ? `個別アプローチ(${effectiveAbsenceTime})` : "ステイ"
    };
  }

  return {
    status: "要確認",
    resolvedApproachType,
    resolvedDept: targetDept,
    reason: rawDocName ? `当日シフトに「${rawDocName}」先生が見つかりません` : "対象医師が未指定",
    doctorName: rawDocName,
    fromClinic: fromClinic || "",
    currentShiftLabel: "要確認",
    consultShiftLabel: effectiveAbsenceTime ? `${clinicName}${effectiveAbsenceTime}` : "要確認"
  };
}

function findDoctorsByScoring_(d3Data, dateStr, targetRaw, clinicHint, deptHint) {
  const sameDay = d3Data.filter(s => s.dateStr === dateStr);
  if (sameDay.length === 0) return [];

  let scored = sameDay.map(s => {
    let score = calcDoctorSimilarityScore_(targetRaw, s);
    if (score > 0) {
      if (clinicHint && clinicHint !== "要確認" && s.clinicName.startsWith(clinicHint)) {
        score += 20;
      }
      if (deptHint === "ワクチン専任" && s.dept.includes("ワクチン")) {
        score += 15;
      }
    }
    return { shift: s, score };
  }).filter(x => x.score >= 60);

  if (scored.length === 0) return [];

  scored.sort((a, b) => b.score - a.score);
  const topScore = scored[0].score;
  const topTier = scored.filter(x => x.score >= topScore - 8);
  return topTier.map(x => x.shift);
}

function parseD6DoctorCell_(cellText) {
  if (!cellText) return [];
  const results = [];
  const segments = cellText.split(/\n/);
  for (const seg of segments) {
    const timeMatch = seg.match(/(\d{1,2}:\d{2})-(\d{1,2}:\d{2})/);
    const start = timeMatch ? padTime_(timeMatch[1]) : "";
    const end = timeMatch ? padTime_(timeMatch[2]) : "";
    const namesPart = seg.replace(/[：:]?\s*\d{1,2}:\d{2}-\d{1,2}:\d{2}/g, "");
    const names = namesPart.split(/[、,，]/).map(n => n.replace(/[▲※★☆\s]/g, "").trim()).filter(Boolean);
    for (const name of names) {
      results.push({ name, start, end });
    }
  }
  return results;
}
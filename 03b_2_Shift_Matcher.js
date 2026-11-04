/**
 * ============================================================================
 * 03b_2_Shift_Matcher.gs
 * 当日シフト（D3・D6前後医師・バックアップ勤務依頼）突合メインロジック
 * ============================================================================
 */

function hasTimeExpansionIntentForMover_(evidenceText, fromClinicName) {
  const ev = String(evidenceText || "");
  if (!ev) return false;

  if ((ev.includes("とりあえず") || ev.includes("まずは") || ev.includes("一旦")) &&
      !ev.includes("前倒") && !ev.includes("早出") && !ev.includes("通し") && !ev.includes("投資")) {
    return false;
  }

  if (/前倒|早出|通し|投資の勤務|フルで|終日|1日|一日|午後から|15時から|午前から|9時から/.test(ev)) {
    return true;
  }

  if (ev.includes("延長")) {
    // ★「※残る野口優輔先生へ21:00まで延長打診を確認後に」は残存医の延長であり、移動医自身の時間拡大ではないため除外
    if (/残る.*?延長/.test(ev)) return false;
    if (fromClinicName && (ev.includes(`${fromClinicName}延長`) || ev.includes(`${fromClinicName}は延長`) || ev.includes(`${fromClinicName}の延長`))) {
      return false;
    }
    return true;
  }

  return false;
}

function resolveDoctorAndShift_(rawDocName, approachType, dateStr, clinicName, dept, fromClinic, d6Row, d3Data, effectiveAbsenceTime, evidenceText, licenseNo) {
  const cleanTarget = String(rawDocName || "").replace(/先生$/g, "").replace(/[\s ]+/g, "").trim();
  const cleanLic = String(licenseNo || "").trim();
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

  // ★0. 「バックアップ勤務依頼」の場合：関東バックアップ待機医師への出動依頼として正常通過させる
  if (resolvedApproachType === "バックアップ勤務依頼") {
    const targetClinicDisp = d6Row ? d6Row.clinicRaw : clinicName;
    const bkShift = d3Data.find(s =>
      s.dateStr === dateStr &&
      ((cleanLic && s.docId === cleanLic) || (cleanTarget && toStandardKanji_(s.docName) === toStandardKanji_(cleanTarget)))
    );
    const docDisp = bkShift ? bkShift.docName : (cleanTarget || "バックアップ医師");
    const curStart = bkShift ? bkShift.startTime : (absStart || "09:00");
    const curEnd = bkShift ? bkShift.endTime : (absEnd || "13:00");
    return {
      status: "OK",
      resolvedApproachType: "勤務地変更相談",
      resolvedDept: targetDept,
      doctorName: docDisp,
      fromClinic: "関東バックアップ",
      isCurrentVaccine: false,
      currentClinicLabel: "関東バックアップ",
      currentStart: curStart,
      currentEnd: curEnd,
      consultStart: absStart || curStart,
      consultEnd: absEnd || curEnd,
      hasTimeExpansion: false,
      consultNoteSuffix: "",
      currentShiftLabel: `関東バックアップ ${curStart}-${curEnd}`,
      consultShiftLabel: `${targetClinicDisp} ${absStart || curStart}-${absEnd || curEnd}`,
      isFullTime: bkShift ? bkShift.isFullTime : false,
      alreadyHasAllowance: true, // バックアップ医師は移動依頼手当(10,000円)対象外
      evidenceNote: `バックアップ医師に勤務依頼(${docDisp})`
    };
  }

  // 1. 個別アプローチ（新規打診）の場合 -> 03b_3_Individual_DocResolver.gs で処理
  if (resolvedApproachType === "個別アプローチ") {
    return resolveIndividualApproachSlot_(
      cleanTarget, cleanLic, dateStr, clinicName, targetDept, d6Row, d3Data, effectiveAbsenceTime, absStart, absEnd
    );
  }

  // 2. D6前後医師欄からの延長・早出特定
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
        const licBonus = (cleanLic && d3Doc && d3Doc.docId === cleanLic) ? 50 : 0;
        return { cand: c, score: calcDoctorSimilarityScore_(cleanTarget, fakeRecord) + licBonus };
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
      const d3Doc = (cleanLic ? d3Data.find(s => s.dateStr === dateStr && s.docId === cleanLic) : null) ||
        d3Data.find(s =>
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

  // 3. 当日D3シフトからの検索
  if (cleanTarget || cleanLic) {
    const clinicHint = (resolvedApproachType === "延長相談" || resolvedApproachType === "早出相談") ? clinicName : fromClinic;
    let sameDayDocs = [];

    if (cleanLic) {
      sameDayDocs = d3Data.filter(s => s.dateStr === dateStr && s.docId === cleanLic);
    }
    if (sameDayDocs.length === 0 && cleanTarget) {
      sameDayDocs = findDoctorsByScoring_(d3Data, dateStr, cleanTarget, clinicHint, dept);
    }

    const uniqueNames = Array.from(new Set(sameDayDocs.map(d => d.docName)));

    if (uniqueNames.length === 1) {
      const s = (clinicHint ? sameDayDocs.find(d => d.clinicName.startsWith(clinicHint)) : null) || sameDayDocs[0];
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
        evidenceNote: cleanLic ? `D3医籍番号一致(${cleanLic}/${originLabel})` : `D3自動スコアリング一致(${originLabel})`
      };
    } else if (uniqueNames.length > 1) {
      return {
        status: "要確認",
        resolvedApproachType,
        resolvedDept: targetDept,
        reason: `当日シフトに同姓・複数候補あり(${uniqueNames.join(", ")})`,
        doctorName: rawDocName,
        currentShiftLabel: "複数候補",
        consultShiftLabel: ""
      };
    }
  }

  if (resolvedApproachType === "募集継続・ステイ") {
    return {
      status: "OK",
      resolvedApproachType,
      resolvedDept: targetDept,
      doctorName: "",
      currentShiftLabel: "-",
      consultShiftLabel: "ステイ"
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
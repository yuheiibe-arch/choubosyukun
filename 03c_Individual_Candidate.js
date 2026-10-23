/**
 * ============================================================================
 * 03c_Individual_Candidate.gs
 * 個別アプローチ候補医師の絞り込み・優先順ソート・除外ルール判定
 * 1) 当日に他拠点や重複時間で勤務が入っている医師は完全除外
 * 2) 同じ拠点の前後枠（時間が重ならない枠）の医師は候補に出してOK（※別行で打診済なら除外）
 * 3) 『打診履歴』シート（過去のMTG等）ですでに同じシフト枠を打診済みの医師は除外
 * 4) 『除外医師』シートの医籍番号一致医師、および『交渉注意医師』NG条件一致医師は除外
 * ============================================================================
 */

function populateIndividualApproachCandidates_(
  reviewRows, draftParamsByRowIdx, targetRows, d3Data, negSheet, excludedSheet, histSheet, d5ClinicSheet, holidaySet
) {
  if (!targetRows || targetRows.length === 0) return;

  const startTimeMs = Date.now();
  const negRules = loadNegotiationRules_(negSheet);
  const excludedSet = loadExcludedDoctors_(excludedSheet);
  const historyRecords = loadApproachHistory_(histSheet);

  const neededClinics = new Set(targetRows.map(t => t.clinicName));
  const targetDateSet = new Set(targetRows.map(t => t.dateStr));

  // 03d_Kinmukun_Scanner.gs のキャッシュ付き高速スキャナを呼び出し
  const { statsByClinicKey, shiftNewOnTargetDates } = getKinmukunStatsWithCache_(neededClinics, targetDateSet, startTimeMs);

  // 今回のMTGの「別の行」ですでに同日に延長・早出・移動等の対象になっている医師（日付 -> Set(正規化医師名)）
  const alreadyTargetedInOtherRowsByDate = new Map();
  for (let rIdx = 0; rIdx < reviewRows.length; rIdx++) {
    const r = reviewRows[rIdx];
    const rDate = r[3];
    const rDoc = String(r[8] || "").trim();
    if (!rDate || !rDoc) continue;
    if (!alreadyTargetedInOtherRowsByDate.has(rDate)) {
      alreadyTargetedInOtherRowsByDate.set(rDate, new Set());
    }
    rDoc.split(/\r?\n/).forEach(name => {
      const n = toStandardKanji_(name);
      if (n) alreadyTargetedInOtherRowsByDate.get(rDate).add(n);
    });
  }

  for (const t of targetRows) {
    const { rowIndex, dateStr, clinicName, dept, d6ClinicKey, absenceTime } = t;
    const isKitaOrKame = (clinicName === "北葛西" || clinicName === "亀有");
    const clinicKey = isKitaOrKame ? `${clinicName}_${dept}` : clinicName;
    const docStatsList = statsByClinicKey.get(clinicKey) || [];

    const dateObj = parseDate_(dateStr);
    const dayOfWeek = dateObj.getDay();
    const isHoliday = holidaySet && holidaySet.has(dateStr);
    const targetFrames = getTargetFrameKeys_(absenceTime, clinicName === "北葛西");

    // 1. 当日のシフト状況から「除外すべき医師（他拠点勤務 または 同拠点で時間重複）」を特定
    //    ※同拠点の前後枠（時間が重ならない）は候補に出してOK
    const blockedDocIdsOnDate = new Set();
    const blockedDocNamesOnDate = new Set();
    const sameClinicAdjacentDocNames = new Set();

    const evaluateDayShift = (sClinic, sDept, sTimeRange, docId, normName) => {
      const isSameClinic = (sClinic === clinicName) && (!isKitaOrKame || !sDept || sDept.includes(dept) || sDept.includes("ワクチン"));
      if (!isSameClinic) {
        if (docId) blockedDocIdsOnDate.add(docId);
        if (normName) blockedDocNamesOnDate.add(normName);
      } else {
        if (!sTimeRange || !absenceTime || isTimeOverlapping_(sTimeRange, absenceTime)) {
          if (docId) blockedDocIdsOnDate.add(docId);
          if (normName) blockedDocNamesOnDate.add(normName);
        } else {
          if (normName) sameClinicAdjacentDocNames.add(normName);
        }
      }
    };

    d3Data.filter(s => s.dateStr === dateStr).forEach(s => {
      evaluateDayShift(s.clinicName, s.dept, `${s.startTime}-${s.endTime}`, s.docId, s.normName);
    });

    const shiftNewList = shiftNewOnTargetDates.get(dateStr) || [];
    shiftNewList.forEach(b => {
      evaluateDayShift(b.kyoten, b.shinryo, b.timeRange, b.id, b.normName);
    });

    const otherRowDocs = alreadyTargetedInOtherRowsByDate.get(dateStr) || new Set();
    const alreadyApproachedDocs = getAlreadyApproachedDocsForSlot_(
      historyRecords, dateStr, clinicName, d6ClinicKey, absenceTime
    );

    const candidates = [];
    for (const docStat of docStatsList) {
      if (docStat.totalCount < 3) continue;

      const dayFrameObj = (docStat.frames && docStat.frames[dayOfWeek]) || { am: 0, pm1: 0, pm2: 0 };
      const dayTotalCount = dayFrameObj.am + dayFrameObj.pm1 + dayFrameObj.pm2;
      if (dayTotalCount === 0) continue;

      if ((docStat.id && blockedDocIdsOnDate.has(docStat.id)) || blockedDocNamesOnDate.has(docStat.normName)) {
        continue;
      }
      if (otherRowDocs.has(docStat.normName)) continue;
      if (alreadyApproachedDocs.has(docStat.normName)) continue;
      if (isDoctorInExcludedSheet_(docStat.id, docStat.normName, excludedSet)) continue;
      if (isExcludedByNegotiationSheet_(docStat, clinicName, dept, dayOfWeek, isHoliday, absenceTime, targetFrames, negRules)) {
        continue;
      }

      let sameDaySameTimeCount = 0;
      let totalSameTimeCount = 0;
      for (const fk of targetFrames) {
        sameDaySameTimeCount += (dayFrameObj[fk] || 0);
        if (fk === "am") totalSameTimeCount += docStat.totalAm;
        if (fk === "pm1") totalSameTimeCount += docStat.totalPm1;
        if (fk === "pm2") totalSameTimeCount += docStat.totalPm2;
      }

      candidates.push({
        name: docStat.name,
        id: docStat.id,
        isSameClinicAdjacent: sameClinicAdjacentDocNames.has(docStat.normName),
        sameDaySameTimeCount,
        dayTotalCount,
        totalSameTimeCount,
        totalCount: docStat.totalCount
      });
    }

    candidates.sort((a, b) => {
      if (b.sameDaySameTimeCount !== a.sameDaySameTimeCount) return b.sameDaySameTimeCount - a.sameDaySameTimeCount;
      if (b.dayTotalCount !== a.dayTotalCount) return b.dayTotalCount - a.dayTotalCount;
      if (b.totalSameTimeCount !== a.totalSameTimeCount) return b.totalSameTimeCount - a.totalSameTimeCount;
      return b.totalCount - a.totalCount;
    });

    const topCandidates = candidates.slice(0, 10);
    const namesText = topCandidates.length > 0
      ? topCandidates.map(c => c.name).join("\n")
      : "該当候補なし";

    reviewRows[rowIndex][8] = namesText;

    if (draftParamsByRowIdx[rowIndex]) {
      draftParamsByRowIdx[rowIndex].doctorName = topCandidates.length === 1 ? topCandidates[0].name : "{医師名}";
      reviewRows[rowIndex][16] = buildMessageTemplate_(draftParamsByRowIdx[rowIndex]);
    }

    const memoInlineNames = topCandidates.length > 0
      ? topCandidates.map(c => c.name).join("・")
      : "該当候補なし";
    reviewRows[rowIndex][17] = reviewRows[rowIndex][17].replace("候補医師抽出対象", memoInlineNames);
    reviewRows[rowIndex][20] = reviewRows[rowIndex][20].replace("候補医師抽出対象", memoInlineNames);

    const detailEvidence = topCandidates.length > 0
      ? `[候補抽出] 同曜日実績あり・他拠点/既打診/除外医師/交渉注意除外済(${topCandidates.length}名): ` +
        topCandidates.map(c => `${c.name}${c.isSameClinicAdjacent ? "[同拠点前後枠]" : ""}(同曜同枠:${c.sameDaySameTimeCount}/同曜:${c.dayTotalCount}/総計:${c.totalCount})`).join(", ")
      : `[候補抽出] 条件一致の候補医師なし`;
    reviewRows[rowIndex][21] = `${reviewRows[rowIndex][21]}\n${detailEvidence}`;
  }
}

function loadApproachHistory_(histSheet) {
  const list = [];
  if (!histSheet || histSheet.getLastRow() < 2) return list;
  const vals = histSheet.getDataRange().getDisplayValues();
  for (let r = 1; r < vals.length; r++) {
    const dateStr = normalizeDateStr_(vals[r][1]);
    const clinicRaw = String(vals[r][2] || "").trim();
    const baseClinic = clinicRaw.replace(/[\(（].*?[\)）]/g, "").trim();
    const timeRange = normalizeTimeRange_(vals[r][3]);
    const docRaw = String(vals[r][6] || "").trim();
    if (!dateStr || !clinicRaw || !docRaw) continue;

    docRaw.split(/[、,，・\r\n]+/).forEach(n => {
      const norm = toStandardKanji_(n);
      if (norm && norm !== "該当候補なし" && norm !== "対象未定") {
        list.push({ dateStr, clinicRaw, baseClinic, timeRange, normName: norm });
      }
    });
  }
  return list;
}

function getAlreadyApproachedDocsForSlot_(historyRecords, dateStr, clinicName, d6ClinicKey, absenceTime) {
  const approached = new Set();
  if (!historyRecords || historyRecords.length === 0) return approached;

  for (const h of historyRecords) {
    if (h.dateStr !== dateStr) continue;
    if (h.clinicRaw !== d6ClinicKey && h.baseClinic !== clinicName) continue;
    if (!h.timeRange || !absenceTime || h.timeRange === absenceTime || isTimeOverlapping_(h.timeRange, absenceTime)) {
      approached.add(h.normName);
    }
  }
  return approached;
}

function loadExcludedDoctors_(excludedSheet) {
  const ids = new Set();
  const names = new Set();
  if (!excludedSheet || excludedSheet.getLastRow() < 2) {
    return { ids, names };
  }
  const vals = excludedSheet.getDataRange().getDisplayValues();
  for (let r = 1; r < vals.length; r++) {
    const rawName = String(vals[r][0] || "").replace(/[\s ]+/g, "").trim();
    const rawId = String(vals[r][1] || "").trim();
    if (rawId) ids.add(rawId);
    if (rawName) names.add(toStandardKanji_(rawName));
  }
  return { ids, names };
}

function isDoctorInExcludedSheet_(docId, normName, excludedSet) {
  if (!excludedSet) return false;
  if (docId && excludedSet.ids.has(String(docId).trim())) return true;
  if (normName && excludedSet.names.has(normName)) return true;
  return false;
}

function loadNegotiationRules_(negSheet) {
  if (!negSheet || negSheet.getLastRow() < 2) return [];
  const vals = negSheet.getDataRange().getDisplayValues();
  const rules = [];

  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    const rawName = String(row[0] || "").replace(/[\s ]+/g, "").trim();
    const licenseNo = String(row[1] || "").trim();
    if (!rawName && !licenseNo) continue;

    const ngLocations = [row[5], row[6], row[7], row[8], row[9]]
      .map(v => String(v || "").trim())
      .filter(Boolean);
    const ngDays = [row[11], row[12], row[13]]
      .map(v => String(v || "").trim())
      .filter(Boolean);

    rules.push({
      normName: toStandardKanji_(rawName),
      licenseNo,
      dept: String(row[2] || "").trim(),
      condition: String(row[3] || "").trim(),
      ngArea: String(row[4] || "").trim(),
      ngLocations,
      note: String(row[10] || "").trim(),
      ngDays,
      ngTime: String(row[14] || "").trim()
    });
  }
  return rules;
}

function isExcludedByNegotiationSheet_(docStat, clinicName, dept, dayOfWeek, isHoliday, absenceTime, targetFrames, negRules) {
  if (!negRules || negRules.length === 0) return false;

  const matchedRules = negRules.filter(r =>
    (r.licenseNo && docStat.id && r.licenseNo === docStat.id) ||
    (r.normName && r.normName === docStat.normName)
  );
  if (matchedRules.length === 0) return false;

  const clinicMeta = CLINIC_META_MAP.get(clinicName) || { area: "", group: "" };
  const dayNames = ["日", "月", "火", "水", "木", "金", "土"];
  const targetDayChar = dayNames[dayOfWeek];

  for (const rule of matchedRules) {
    if (rule.ngArea) {
      if (
        rule.ngArea === "すべて" ||
        rule.ngArea === "全拠点" ||
        (clinicMeta.area && rule.ngArea.includes(clinicMeta.area)) ||
        (clinicMeta.group && rule.ngArea.includes(clinicMeta.group))
      ) {
        return true;
      }
    }

    if (rule.ngLocations.length > 0) {
      for (const ngLoc of rule.ngLocations) {
        const cleanNg = ngLoc.replace(/[\(（].*?[\)）]/g, "").trim();
        if (cleanNg === clinicName || ngLoc.includes(clinicName)) {
          return true;
        }
      }
    }

    if (rule.condition.includes("不可") && !rule.ngArea && rule.ngLocations.length === 0) {
      return true;
    }

    if (rule.ngDays.length > 0) {
      for (const ngDay of rule.ngDays) {
        if (ngDay.includes(targetDayChar)) return true;
        if (isHoliday && ngDay.includes("祝")) return true;
      }
    }

    if (rule.ngTime) {
      const nt = rule.ngTime;
      if (targetFrames.includes("am") && (nt.includes("午前") || nt.includes("9") || nt.includes("09"))) return true;
      if (targetFrames.includes("pm1") && (nt.includes("午後") || nt.includes("15"))) return true;
      if (targetFrames.includes("pm2") && (nt.includes("夜") || nt.includes("18") || nt.includes("21"))) return true;
      if (absenceTime && nt.includes(absenceTime)) return true;
    }
  }

  return false;
}
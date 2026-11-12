/**
 * ============================================================================
 * 02b_1_Merge_SlotCheck.gs（マージ管理 3分割の1）
 * ・取込済メールID取得
 * ・枠の同一性・時間重複判定
 * ・最新D6『アプローチシート』との未充足照合ヘルパー
 * ============================================================================
 */

function getLoggedEmailIdsFromD7Log_(logSheet) {
  const loggedIds = new Set();
  if (!logSheet || logSheet.getLastRow() < 2) return loggedIds;
  logSheet
    .getRange(2, 3, logSheet.getLastRow() - 1, 1)
    .getDisplayValues()
    .forEach(function(r) {
      const id = String(r[0] || "").trim();
      if (id) loggedIds.add(id);
    });
  return loggedIds;
}

function makeSlotIdentityForMerge_(dateStr, clinicRaw, deptHint, timeStr) {
  const normDate = normalizeDateStr_(dateStr);
  const normClinic = normalizeClinicForMatch_(clinicRaw);
  const isKitaOrKame = (normClinic === "北葛西" || normClinic === "亀有");
  const isInt = String(clinicRaw || "").includes("内科") || String(deptHint || "").includes("内科");
  const deptPart = isInt ? "内科" : (isKitaOrKame ? "小児科" : "ALL");
  return {
    dateStr: normDate,
    normClinic: normClinic,
    deptPart: deptPart,
    timeRange: normalizeTimeRange_(timeStr)
  };
}

function isSameOrOverlappingSlotIdentity_(a, b) {
  if (!a || !b) return false;
  if (a.dateStr !== b.dateStr) return false;
  if (a.normClinic !== b.normClinic) return false;
  if (a.deptPart !== "ALL" && b.deptPart !== "ALL" && a.deptPart !== b.deptPart) return false;
  if (!a.timeRange || !b.timeRange || a.timeRange === b.timeRange) return true;
  return isTimeOverlapping_(a.timeRange, b.timeRange);
}

/**
 * 指定枠が最新D6『アプローチシート』上でまだ未充足として残っているかを判定する
 * （※D6シートが空＝貼付作業中の場合は誤削除防止のため未充足扱いで保護する）
 */
function checkSlotStillUnfilledInD6_(dateStr, d6ClinicKey, absenceTime, deptHint, d6Data, todayYmd) {
  const normDate = normalizeDateStr_(dateStr);
  if (!normDate || normDate < todayYmd) {
    return { unfilled: false, d6Row: null };
  }
  if (!d6Data || d6Data.length === 0) {
    return { unfilled: true, d6Row: null };
  }
  if (deptHint === "ワクチン専任" || String(d6ClinicKey || "").includes("ワクチン")) {
    return { unfilled: true, d6Row: null };
  }

  const baseClinic = String(d6ClinicKey || "").replace(/[\(（].*?[\)）]/g, "").trim();
  if (!baseClinic || baseClinic === "要確認") {
    return { unfilled: true, d6Row: null };
  }

  const isKitaOrKame = (baseClinic === "北葛西" || baseClinic === "亀有");
  const slotDept = String(d6ClinicKey || "").includes("内科") || String(deptHint || "").includes("内科")
    ? "内科"
    : "小児科";

  const matched = matchD6RowsSmart_(d6Data, normDate, baseClinic, slotDept, absenceTime, isKitaOrKame);
  return {
    unfilled: Boolean(matched && matched.mergedRow),
    d6Row: matched ? matched.mergedRow : null
  };
}

/**
 * 既存のSlack本文セルから、特定インデックス枠に【特別時給】が付いていたかを判定するヘルパー
 */
function detectSlotSpecialFromExistingRow_(origRow, slotIdx, slotCount, d6Row) {
  if (d6Row && d6Row.isSpecialWage) return true;
  const rowSp = String(origRow[11] || "").includes("特別時給");
  if (!rowSp) return false;
  if (slotCount <= 1) return true;

  const slackRaw = String(origRow[20] || "");
  const blocks = slackRaw.split(/(?=`\d{1,2}月\d{1,2}日)/).map(function(s) { return s.trim(); }).filter(Boolean);
  if (blocks[slotIdx]) {
    return blocks[slotIdx].includes("特別時給");
  }
  return rowSp;
}
/**
 * ============================================================================
 * 04_Wage_Calculator.gs
 * 内閣府祝日CSV取得・D4時給表自動選択・15分刻み日給計算
 * - 午前勤務医への通し延長依頼時は「18:00まで(consultPm1DailyWage)」と「21:00まで(consultDailyWage)」の2パターンを算出
 * ============================================================================
 */

let WAGE_SHEET_CACHE = {};

function updateCabinetOfficeHolidays_(holSheet) {
  try {
    const res = UrlFetchApp.fetch(CAO_HOLIDAY_CSV_URL, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return;

    const csvText = res.getContentText("Shift_JIS");
    const lines = csvText.split(/\r?\n/);
    const nowStr = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");
    const rows = [];

    for (let i = 1; i < lines.length; i++) {
      const cols = lines[i].split(",");
      if (cols.length < 2) continue;
      const dateStr = normalizeDateStr_(cols[0]);
      if (dateStr >= "2025/01/01") {
        rows.push([dateStr, cols[1].trim(), nowStr]);
      }
    }

    if (rows.length > 0) {
      writeDataExactRows_(holSheet, rows, 3);
    }
  } catch (e) {
    console.log("内閣府祝日CSV取得エラー: " + e.message);
  }
}

function loadHolidaySet_(holSheet) {
  const set = new Set();
  if (!holSheet || holSheet.getLastRow() < 2) return set;
  const vals = holSheet.getRange(2, 1, holSheet.getLastRow() - 1, 1).getDisplayValues();
  vals.forEach(r => {
    if (r[0]) set.add(normalizeDateStr_(r[0]));
  });
  return set;
}

function calculateWagesAndAllowance_(params) {
  const {
    wageSs, dateStr, clinicName, dept, fromClinicName,
    approachType, absenceTime, currentStart, currentEnd,
    consultStart, consultEnd, isCurrentVaccine,
    isSpecialFromMtg, isSpecialFromD6, isFullTime, alreadyHasAllowance, holidaySet
  } = params;

  const isSpecial = Boolean(isSpecialFromMtg || isSpecialFromD6);

  // 午前勤務（終了が13:30以前）の医師に、夜間（20:00または21:00）までの通し延長をお願いするケースか判定
  const curEndMin = toMinutes_(currentEnd);
  const conEndMin = toMinutes_(consultEnd);
  const isMorningToNightExt =
    approachType === "延長相談" &&
    !isNaN(curEndMin) && curEndMin <= 13 * 60 + 30 &&
    !isNaN(conEndMin) && conEndMin >= 20 * 60;

  if (isFullTime) {
    return {
      status: "OK",
      isSpecial,
      isMorningToNightExt,
      currentDailyWage: 0,
      consultPm1DailyWage: 0,
      consultDailyWage: 0,
      requestAllowance: 0,
      diffHoursText: calcDiffHoursText_(currentStart, currentEnd, consultStart, consultEnd),
      sheetNameUsed: "常勤医師(時給合計0)",
      wageNote: "常勤のため日給表記なし"
    };
  }

  if (!dateStr || !clinicName || clinicName === "要確認") {
    return { status: "要確認", reason: "日付または拠点名未確定", isSpecial };
  }

  const sheetNameUsed = selectWageSheetName_(dateStr, clinicName);
  const wageTable = loadWageTableCached_(wageSs, sheetNameUsed);
  if (!wageTable) {
    return { status: "要確認", reason: `時給表「${sheetNameUsed}」が見つかりません`, isSpecial };
  }

  const dayCategory = getDayCategory_(dateStr, holidaySet);
  const targetLookupKey = (clinicName === "北葛西" || clinicName === "亀有")
    ? `${clinicName}（${dept}）`
    : clinicName;
  const targetRates = wageTable.get(targetLookupKey) || wageTable.get(clinicName);

  if (!targetRates) {
    return { status: "要確認", reason: `時給表に「${targetLookupKey}」の行がありません`, isSpecial, sheetNameUsed };
  }

  let fromRates = targetRates;
  if (fromClinicName && fromClinicName !== clinicName && fromClinicName !== "ワクチン専任") {
    fromRates = wageTable.get(fromClinicName) || wageTable.get(`${fromClinicName}（${dept}）`) || targetRates;
  }

  // 1. 現在シフトの日給計算
  let currentDailyWage = 0;
  if (currentStart && currentEnd) {
    if (isCurrentVaccine) {
      const curMins = calcWorkingMinutesExcludeBreak_(currentStart, currentEnd);
      currentDailyWage = Math.round((curMins / 60) * VACCINE_DEDICATED_HOURLY_WAGE);
    } else {
      currentDailyWage = calcIntervalWage_(
        currentStart, currentEnd, fromRates[dayCategory], clinicName === "北葛西", null, 0
      );
    }
  }

  // 2. 相談シフトの日給計算（通し延長時は18:00までのパターン①と21:00までのパターン②の両方を算出）
  const specialBonus = isSpecial ? 2000 : 0;
  let consultDailyWage = 0;
  let consultPm1DailyWage = 0;

  if (approachType === "個別アプローチ" || approachType === "募集継続・ステイ") {
    const [absS, absE] = (absenceTime || "-").split("-");
    if (absS && absE) {
      consultDailyWage = calcIntervalWage_(
        absS, absE, targetRates[dayCategory], clinicName === "北葛西", absenceTime, specialBonus
      );
    }
  } else if (approachType === "勤務地変更相談") {
    const baseMaxRates = isCurrentVaccine
      ? {
          am: Math.max(VACCINE_DEDICATED_HOURLY_WAGE, targetRates[dayCategory].am),
          pm: Math.max(VACCINE_DEDICATED_HOURLY_WAGE, targetRates[dayCategory].pm),
          night: Math.max(VACCINE_DEDICATED_HOURLY_WAGE, targetRates[dayCategory].night)
        }
      : {
          am: Math.max(fromRates[dayCategory].am, targetRates[dayCategory].am),
          pm: Math.max(fromRates[dayCategory].pm, targetRates[dayCategory].pm),
          night: Math.max(fromRates[dayCategory].night, targetRates[dayCategory].night)
        };
    consultDailyWage = calcIntervalWage_(
      consultStart, consultEnd, baseMaxRates, clinicName === "北葛西", absenceTime, specialBonus
    );
  } else if ((approachType === "延長相談" || approachType === "早出相談") && isCurrentVaccine && dept !== "ワクチン専任") {
    const [absS, absE] = (absenceTime || "-").split("-");
    const addedClinicWage = (absS && absE)
      ? calcIntervalWage_(absS, absE, targetRates[dayCategory], clinicName === "北葛西", absenceTime, specialBonus)
      : 0;
    consultDailyWage = currentDailyWage + addedClinicWage;

    if (isMorningToNightExt && absS) {
      const addedPm1Wage = calcIntervalWage_(
        absS, "18:00", targetRates[dayCategory], clinicName === "北葛西", absenceTime, specialBonus
      );
      consultPm1DailyWage = currentDailyWage + addedPm1Wage;
    }
  } else {
    consultDailyWage = calcIntervalWage_(
      consultStart, consultEnd, targetRates[dayCategory], clinicName === "北葛西", absenceTime, specialBonus
    );
    if (isMorningToNightExt) {
      consultPm1DailyWage = calcIntervalWage_(
        consultStart, "18:00", targetRates[dayCategory], clinicName === "北葛西", absenceTime, specialBonus
      );
    }
  }

  let requestAllowance = 0;
  if (approachType === "勤務地変更相談" && !alreadyHasAllowance) {
    requestAllowance = 10000;
  }

  return {
    status: "OK",
    isSpecial,
    isMorningToNightExt,
    currentDailyWage,
    consultPm1DailyWage,
    consultDailyWage,
    requestAllowance,
    diffHoursText: calcDiffHoursText_(currentStart, currentEnd, consultStart, consultEnd),
    sheetNameUsed,
    wageNote: `${dayCategory === "weekday" ? "平日" : dayCategory === "sat" ? "土曜" : "日祝"} / 特別時給:${isSpecial ? "あり(通常時給+2,000円/h)" : "なし"}${isCurrentVaccine ? " / 現在枠:ワクチン専任(13,000円/h)" : ""}`
  };
}

function selectWageSheetName_(dateStr, clinicName) {
  const [y, m] = dateStr.split("/").map(Number);
  const term = (m >= 4 && m <= 9) ? "上期" : "下期";
  if (clinicName === "西宮北口" || clinicName === "あべのキューズ") {
    return "関西拠点時給";
  }
  return `${y}${term}時給`;
}

function loadWageTableCached_(wageSs, sheetName) {
  if (WAGE_SHEET_CACHE[sheetName]) return WAGE_SHEET_CACHE[sheetName];
  const sheet = wageSs.getSheetByName(sheetName);
  if (!sheet) return null;

  const vals = sheet.getDataRange().getDisplayValues();
  const map = new Map();
  const parseYen = s => Number(String(s || "0").replace(/[¥￥,，\s]/g, "")) || 0;

  for (let r = 2; r < vals.length; r++) {
    const row = vals[r];
    const name = String(row[1] || "").trim();
    if (!name) continue;
    map.set(name, {
      weekday: { am: parseYen(row[2]), pm: parseYen(row[3]), night: parseYen(row[4]) },
      sat:     { am: parseYen(row[5]), pm: parseYen(row[6]), night: parseYen(row[7]) },
      sunHol:  { am: parseYen(row[8]), pm: parseYen(row[9]), night: parseYen(row[10]) }
    });
  }
  WAGE_SHEET_CACHE[sheetName] = map;
  return map;
}

function getDayCategory_(dateStr, holidaySet) {
  if (holidaySet && holidaySet.has(dateStr)) return "sunHol";
  const d = parseDate_(dateStr);
  const dow = d.getDay();
  if (dow === 0) return "sunHol";
  if (dow === 6) return "sat";
  return "weekday";
}

function calcIntervalWage_(startStr, endStr, rates, isKitaKasai, specialRangeStr, specialBonusPerHour) {
  const startMin = toMinutes_(startStr);
  const endMin = toMinutes_(endStr);
  if (isNaN(startMin) || isNaN(endMin) || endMin <= startMin) return 0;

  let spStart = -1, spEnd = -1;
  if (specialRangeStr && specialRangeStr.includes("-") && specialBonusPerHour > 0) {
    const [s, e] = specialRangeStr.split("-");
    spStart = toMinutes_(s);
    spEnd = toMinutes_(e);
  }

  let totalYen = 0;
  for (let t = startMin; t < endMin; t += 15) {
    if (t >= 13 * 60 && t < 15 * 60) continue;

    let baseHourly = rates.am;
    if (isKitaKasai) {
      if (t < 12 * 60 + 30) baseHourly = rates.am;
      else if (t >= 14 * 60 && t < 18 * 60) baseHourly = rates.pm;
      else if (t >= 18 * 60) baseHourly = rates.night;
    } else {
      if (t < 13 * 60) baseHourly = rates.am;
      else if (t >= 15 * 60 && t < 18 * 60) baseHourly = rates.pm;
      else if (t >= 18 * 60) baseHourly = rates.night;
    }

    let bonus = 0;
    if (spStart !== -1 && t >= spStart && t < spEnd) {
      bonus = specialBonusPerHour;
    }
    totalYen += (baseHourly + bonus) * 0.25;
  }
  return Math.round(totalYen);
}
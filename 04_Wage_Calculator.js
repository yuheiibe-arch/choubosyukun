/**
 * ============================================================================
 * 04_Wage_Calculator.gs
 * 内閣府祝日CSV取得・D4時給表自動選択・15分刻み日給計算（動的ヘッダー＆全拠点対応版）
 * ============================================================================
 */

let WAGE_TABLES_BY_SHEET_CACHE = null;

function updateCabinetOfficeHolidays_(holSheet) {
  try {
    const csvUrl = typeof cleanUrl_ === "function"
      ? cleanUrl_(CAO_HOLIDAY_CSV_URL)
      : String(CAO_HOLIDAY_CSV_URL || "").replace(/^\[([^\]]+)\]\([^\)]+\)$/, "$1").trim();

    const res = UrlFetchApp.fetch(csvUrl, { muteHttpExceptions: true });
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

function cleanWageClinicKey_(name) {
  return String(name || "")
    .replace(/^[0-9０-９]+[\.．、\s\u3000]*/, "")
    .replace(/[\(（].*?[\)）]/g, "")
    .replace(/[ヶケが]/g, "ケ")
    .replace(/[\s\u3000]+/g, "")
    .trim();
}

function getAllWageTablesCached_(wageSs) {
  if (WAGE_TABLES_BY_SHEET_CACHE) return WAGE_TABLES_BY_SHEET_CACHE;
  if (typeof buildWageTablesForShuseikun_ === "function") {
    WAGE_TABLES_BY_SHEET_CACHE = buildWageTablesForShuseikun_(wageSs);
    return WAGE_TABLES_BY_SHEET_CACHE;
  }

  const wageTablesObj = {};
  const parseYen = val => Number(String(val || "").replace(/[^0-9]/g, "")) || 0;

  wageSs.getSheets().forEach(sh => {
    const sName = sh.getName();
    if (!sName.includes("時給") || sName.includes("医師") || sName.includes("特別")) return;
    if (sh.getLastRow() < 2) return;

    const vals = sh.getDataRange().getDisplayValues();
    const headers = vals[0].map(h => String(h || "").trim());
    let colClinic = headers.indexOf("拠点名");
    let colDept = headers.indexOf("科目");
    let colWAm = headers.indexOf("平日_午前");
    let colSAm = headers.indexOf("土曜_午前");
    let colHAm = headers.indexOf("日祝_午前");

    if (colClinic === -1) colClinic = 1;
    if (colDept === -1) colDept = 2;
    if (colWAm === -1) colWAm = 3;
    if (colSAm === -1) colSAm = 6;
    if (colHAm === -1) colHAm = 9;

    const obj = {};
    for (let r = 1; r < vals.length; r++) {
      const row = vals[r];
      const base = String(row[colClinic] || "").replace(/^[0-9０-９]+[\.．、\s\u3000]*/, "").replace(/[\(（].*?[\)）]/g, "").trim();
      if (!base || base === "拠点名") continue;
      const deptRaw = String(row[colDept] || "小児科").trim();
      const dept = deptRaw.includes("内科") ? "内科" : "小児科";
      const rateData = {
        weekday: { am: parseYen(row[colWAm]), pm: parseYen(row[colWAm + 1]), night: parseYen(row[colWAm + 2]) },
        sat:     { am: parseYen(row[colSAm]), pm: parseYen(row[colSAm + 1]), night: parseYen(row[colSAm + 2]) },
        sunHol:  { am: parseYen(row[colHAm]), pm: parseYen(row[colHAm + 1]), night: parseYen(row[colHAm + 2]) }
      };
      obj[`${base}|${dept}`] = rateData;
      obj[`${cleanWageClinicKey_(base)}|${dept}`] = rateData;
      if (base !== "北葛西" && base !== "亀有") {
        obj[base] = rateData;
        obj[cleanWageClinicKey_(base)] = rateData;
      }
    }
    wageTablesObj[sName] = obj;
  });

  WAGE_TABLES_BY_SHEET_CACHE = wageTablesObj;
  return wageTablesObj;
}

function findClinicRateInTables_(wageSs, dateStr, clinicName, dept) {
  if (!clinicName || clinicName === "要確認") return null;
  const allTables = getAllWageTablesCached_(wageSs);
  const availableSheets = Object.keys(allTables);
  if (availableSheets.length === 0) return null;

  const [y, m] = String(dateStr || "2026/10/01").split("/").map(Number);
  const fiscalYear = (m >= 1 && m <= 3) ? (y - 1) : y;
  const halfLabel = (m >= 4 && m <= 9) ? "上期" : "下期";

  const preferredOrder = [
    `${fiscalYear}${halfLabel}時給`,
    `${y}${halfLabel}時給`,
    "関西拠点時給",
    `${fiscalYear}時給`,
    `${y}時給`
  ];
  availableSheets.slice().sort().reverse().forEach(s => {
    if (!preferredOrder.includes(s)) preferredOrder.push(s);
  });

  const baseClinic = String(clinicName || "").replace(/^[0-9０-９]+[\.．、\s\u3000]*/, "").replace(/[\(（].*?[\)）]/g, "").trim();
  const normKey = cleanWageClinicKey_(baseClinic);
  const d = (dept === "内科") ? "内科" : "小児科";
  const isKitaOrKame = (normKey === "北葛西" || normKey === "亀有");

  for (const sName of preferredOrder) {
    const tbl = allTables[sName];
    if (!tbl) continue;

    const strictCandidates = [
      `${baseClinic}|${d}`,
      `${normKey}|${d}`,
      `${baseClinic}(${d})`,
      `${baseClinic}（${d}）`
    ];
    for (const k of strictCandidates) {
      if (tbl[k]) return { rates: tbl[k], sheetName: sName };
    }

    const keys = Object.keys(tbl);
    for (const k of keys) {
      const [kClinic, kDept] = k.split("|");
      if (kDept === d && cleanWageClinicKey_(kClinic) === normKey) {
        return { rates: tbl[k], sheetName: sName };
      }
    }

    if (!isKitaOrKame) {
      if (tbl[baseClinic]) return { rates: tbl[baseClinic], sheetName: sName };
      if (tbl[normKey]) return { rates: tbl[normKey], sheetName: sName };
      for (const k of keys) {
        const kClinic = k.split("|")[0];
        if (cleanWageClinicKey_(kClinic) === normKey) {
          return { rates: tbl[k], sheetName: sName };
        }
      }
    }
  }
  return null;
}

function calculateWagesAndAllowance_(params) {
  const {
    wageSs, dateStr, clinicName, dept, fromClinicName,
    approachType, absenceTime, currentStart, currentEnd,
    consultStart, consultEnd, isCurrentVaccine,
    isSpecialFromMtg, isSpecialFromD6, isFullTime, alreadyHasAllowance, holidaySet
  } = params;

  const isSpecial = Boolean(isSpecialFromMtg || isSpecialFromD6);

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

  const targetFound = findClinicRateInTables_(wageSs, dateStr, clinicName, dept);
  if (!targetFound || !targetFound.rates) {
    const targetLookupKey = (clinicName === "北葛西" || clinicName === "亀有")
      ? `${clinicName}（${dept}）`
      : clinicName;
    return { status: "要確認", reason: `時給表に「${targetLookupKey}」の行がありません`, isSpecial };
  }

  const targetRates = targetFound.rates;
  const sheetNameUsed = targetFound.sheetName;
  const dayCategory = getDayCategory_(dateStr, holidaySet);

  let fromRates = targetRates;
  if (fromClinicName && fromClinicName !== clinicName && fromClinicName !== "ワクチン専任") {
    const fromFound = findClinicRateInTables_(wageSs, dateStr, fromClinicName, dept);
    if (fromFound && fromFound.rates) {
      fromRates = fromFound.rates;
    }
  }

  let currentDailyWage = 0;
  if (currentStart && currentEnd) {
    if (isCurrentVaccine) {
      const curMins = calcWorkingMinutesExcludeBreak_(currentStart, currentEnd);
      currentDailyWage = Math.round((curMins / 60) * VACCINE_DEDICATED_HOURLY_WAGE);
    } else {
      const isFromKita = cleanWageClinicKey_(fromClinicName || clinicName) === "北葛西";
      currentDailyWage = calcIntervalWage_(
        currentStart, currentEnd, fromRates[dayCategory], isFromKita, null, 0
      );
    }
  }

  const specialBonus = isSpecial ? 2000 : 0;
  let consultDailyWage = 0;
  let consultPm1DailyWage = 0;
  const isTargetKita = cleanWageClinicKey_(clinicName) === "北葛西";

  if (approachType === "個別アプローチ" || approachType === "募集継続・ステイ") {
    const [absS, absE] = (absenceTime || "-").split("-");
    if (absS && absE) {
      consultDailyWage = calcIntervalWage_(
        absS, absE, targetRates[dayCategory], isTargetKita, absenceTime, specialBonus
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
      consultStart, consultEnd, baseMaxRates, isTargetKita, absenceTime, specialBonus
    );
  } else if ((approachType === "延長相談" || approachType === "早出相談") && isCurrentVaccine && dept !== "ワクチン専任") {
    const [absS, absE] = (absenceTime || "-").split("-");
    const addedClinicWage = (absS && absE)
      ? calcIntervalWage_(absS, absE, targetRates[dayCategory], isTargetKita, absenceTime, specialBonus)
      : 0;
    consultDailyWage = currentDailyWage + addedClinicWage;

    if (isMorningToNightExt && absS) {
      const addedPm1Wage = calcIntervalWage_(
        absS, "18:00", targetRates[dayCategory], isTargetKita, absenceTime, specialBonus
      );
      consultPm1DailyWage = currentDailyWage + addedPm1Wage;
    }
  } else {
    consultDailyWage = calcIntervalWage_(
      consultStart, consultEnd, targetRates[dayCategory], isTargetKita, absenceTime, specialBonus
    );
    if (isMorningToNightExt) {
      consultPm1DailyWage = calcIntervalWage_(
        consultStart, "18:00", targetRates[dayCategory], isTargetKita, absenceTime, specialBonus
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
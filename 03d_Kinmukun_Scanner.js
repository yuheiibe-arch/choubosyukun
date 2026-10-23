/**
 * ============================================================================
 * 03d_Kinmukun_Scanner.gs
 * 「勤務くん」シフトブック（MONTHLY_1, MONTHLY_2, SHIFT_NEW）の高速スキャン＆キャッシュ管理
 * ============================================================================
 */

var KINMUKUN_CACHE_TTL_SEC = 21600; // 6時間キャッシュ保持
var MAX_SAFE_EXECUTION_MS = 180 * 1000; // タイムアウト防止ブレーカー（3分）

function getKinmukunStatsWithCache_(neededClinics, targetDateSet, startTimeMs) {
  const cache = CacheService.getScriptCache();
  const resultMap = new Map();
  const uncachedClinics = new Set();

  for (const clinic of neededClinics) {
    if (clinic === "北葛西" || clinic === "亀有") {
      const k1 = `KINMU_V2_${clinic}_小児科`;
      const k2 = `KINMU_V2_${clinic}_内科`;
      const c1 = cache.get(k1);
      const c2 = cache.get(k2);
      if (c1 || c2) {
        if (c1) resultMap.set(`${clinic}_小児科`, JSON.parse(c1));
        if (c2) resultMap.set(`${clinic}_内科`, JSON.parse(c2));
      } else {
        uncachedClinics.add(clinic);
      }
    } else {
      const cachedJson = cache.get(`KINMU_V2_${clinic}`);
      if (cachedJson) {
        resultMap.set(clinic, JSON.parse(cachedJson));
      } else {
        uncachedClinics.add(clinic);
      }
    }
  }

  const today = new Date();
  const startDate = new Date(today.getFullYear() - 1, today.getMonth(), 1);
  const endDate = new Date(today.getFullYear() + 1, today.getMonth(), 1);

  if (uncachedClinics.size === 0) {
    console.log("候補医師の過去実績データをキャッシュから0.1秒で取得しました！");
    const shiftNewOnTargetDates = scanTargetDatesShiftsFromShiftNew_(targetDateSet);
    return { statsByClinicKey: resultMap, shiftNewOnTargetDates };
  }

  console.log(`未キャッシュ拠点(${Array.from(uncachedClinics).join(", ")})の過去実績を集計します...`);
  const { statsByClinicKey, shiftNewOnTargetDates } = scanKinmukunShiftBooksFast_(
    startDate, endDate, uncachedClinics, targetDateSet, startTimeMs
  );

  const toCacheObj = {};
  for (const [clinicKey, docMap] of statsByClinicKey.entries()) {
    const slimList = Array.from(docMap.values()).filter(d => d.totalCount >= 3);
    resultMap.set(clinicKey, slimList);
    try {
      const jsonStr = JSON.stringify(slimList);
      if (jsonStr.length < 95000) {
        toCacheObj[`KINMU_V2_${clinicKey}`] = jsonStr;
      }
    } catch (e) {}
  }

  if (Object.keys(toCacheObj).length > 0) {
    try {
      cache.putAll(toCacheObj, KINMUKUN_CACHE_TTL_SEC);
    } catch (e) {}
  }

  return { statsByClinicKey: resultMap, shiftNewOnTargetDates };
}

function formatCellTimeHHMM_(val) {
  if (!val) return "";
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return "";
    return `${("0" + val.getHours()).slice(-2)}:${("0" + val.getMinutes()).slice(-2)}`;
  }
  return padTime_(val);
}

function scanTargetDatesShiftsFromShiftNew_(targetDateSet) {
  const shiftNewOnTargetDates = new Map();
  if (!targetDateSet || targetDateSet.size === 0) return shiftNewOnTargetDates;

  try {
    const ss = SpreadsheetApp.openById(EXTERNAL_IDS.KINMU_SHIFT_NEW);
    for (const sName of ["確定シフト", "応募シフト"]) {
      const sheet = ss.getSheetByName(sName);
      if (!sheet) continue;
      const data = sheet.getDataRange().getValues();
      if (data.length < 2) continue;

      const headers = data[0];
      const iName = headers.indexOf("名前");
      const iId = headers.indexOf("医籍番号");
      const iKyoten = headers.indexOf("クリニック名");
      const iShinryo = headers.indexOf("診療科");
      const iDate = headers.indexOf("勤務日");
      const iStart = headers.indexOf("勤務開始時間");
      const iEnd = headers.indexOf("勤務終了時間");
      if (iName === -1 || iDate === -1) continue;

      for (let r = 1; r < data.length; r++) {
        const row = data[r];
        const rDate = row[iDate] instanceof Date ? row[iDate] : new Date(row[iDate]);
        if (isNaN(rDate.getTime())) continue;
        const dStr = Utilities.formatDate(rDate, "Asia/Tokyo", "yyyy/MM/dd");
        if (!targetDateSet.has(dStr)) continue;

        const rawName = String(row[iName] || "").replace(/\s+/g, "").trim();
        if (!rawName || rawName.includes("休館")) continue;

        const stStr = iStart !== -1 ? formatCellTimeHHMM_(row[iStart]) : "";
        const enStr = iEnd !== -1 ? formatCellTimeHHMM_(row[iEnd]) : "";

        if (!shiftNewOnTargetDates.has(dStr)) shiftNewOnTargetDates.set(dStr, []);
        shiftNewOnTargetDates.get(dStr).push({
          id: iId !== -1 ? String(row[iId] || "").trim() : "",
          normName: toStandardKanji_(rawName),
          kyoten: iKyoten !== -1 ? String(row[iKyoten] || "").trim() : "",
          shinryo: iShinryo !== -1 ? String(row[iShinryo] || "").trim() : "",
          timeRange: (stStr && enStr) ? `${stStr}-${enStr}` : ""
        });
      }
    }
  } catch (e) {}

  return shiftNewOnTargetDates;
}

function scanKinmukunShiftBooksFast_(startDate, endDate, neededClinics, targetDateSet, startTimeMs) {
  const statsByClinicKey = new Map();
  const shiftNewOnTargetDates = new Map();
  const startYmStr = Utilities.formatDate(startDate, "Asia/Tokyo", "yyyy/MM");
  const endYmStr = Utilities.formatDate(endDate, "Asia/Tokyo", "yyyy/MM");

  const parseHourFast = t => {
    if (!t) return null;
    if (t instanceof Date) return t.getHours() + t.getMinutes() / 60;
    const s = String(t);
    const colonIdx = s.indexOf(":");
    if (colonIdx > 0) {
      const h = Number(s.slice(Math.max(0, colonIdx - 2), colonIdx));
      const m = Number(s.slice(colonIdx + 1, colonIdx + 3));
      if (!isNaN(h) && !isNaN(m)) return h + m / 60;
    }
    return null;
  };

  const processBook = (ssId, isNewShiftBook) => {
    if (Date.now() - startTimeMs > MAX_SAFE_EXECUTION_MS) {
      console.log("⚠️ 実行時間上限に近づいたため、追加シフトブックの走査を安全スキップしました。");
      return;
    }
    try {
      const ss = SpreadsheetApp.openById(ssId);
      const sheets = ss.getSheets();

      for (const sheet of sheets) {
        if (Date.now() - startTimeMs > MAX_SAFE_EXECUTION_MS) break;

        const sName = sheet.getName();
        if (!isNewShiftBook) {
          if (!/^\d{4}\/\d{2}$/.test(sName)) continue;
          if (sName < startYmStr || sName > endYmStr) continue;
        } else {
          if (sName !== "確定シフト" && sName !== "応募シフト") continue;
        }

        const isObo = (sName === "応募シフト");
        const data = sheet.getDataRange().getValues();
        if (data.length < 2) continue;

        const headers = data[0];
        const iName = headers.indexOf("名前");
        const iId = headers.indexOf("医籍番号");
        const iKyoten = headers.indexOf("クリニック名");
        const iShinryo = headers.indexOf("診療科");
        const iDate = headers.indexOf("勤務日");
        let iStart = headers.indexOf("(元のシフト）勤務開始時間");
        let iEnd = headers.indexOf("(元のシフト)勤務終了時間");
        if (iStart === -1) iStart = headers.indexOf("勤務開始時間");
        if (iEnd === -1) iEnd = headers.indexOf("勤務終了時間");
        const iWage = headers.indexOf("時給合計");
        const iRemarks = headers.indexOf("業務内容（備考）");

        if (iName === -1 || iId === -1 || iKyoten === -1 || iDate === -1) continue;

        for (let r = 1; r < data.length; r++) {
          const row = data[r];
          const kyoten = row[iKyoten];

          if (isNewShiftBook && targetDateSet && targetDateSet.size > 0) {
            const rDateTmp = row[iDate] instanceof Date ? row[iDate] : new Date(row[iDate]);
            if (!isNaN(rDateTmp.getTime())) {
              const dStr = Utilities.formatDate(rDateTmp, "Asia/Tokyo", "yyyy/MM/dd");
              if (targetDateSet.has(dStr)) {
                const rNameTmp = String(row[iName] || "").replace(/\s+/g, "").trim();
                if (rNameTmp && !rNameTmp.includes("休館")) {
                  const stStr = formatCellTimeHHMM_(row[iStart]);
                  const enStr = formatCellTimeHHMM_(row[iEnd]);
                  if (!shiftNewOnTargetDates.has(dStr)) shiftNewOnTargetDates.set(dStr, []);
                  shiftNewOnTargetDates.get(dStr).push({
                    id: String(row[iId] || "").trim(),
                    normName: toStandardKanji_(rNameTmp),
                    kyoten: String(kyoten || "").trim(),
                    shinryo: iShinryo !== -1 ? String(row[iShinryo] || "").trim() : "",
                    timeRange: (stStr && enStr) ? `${stStr}-${enStr}` : ""
                  });
                }
              }
            }
          }

          if (!neededClinics.has(kyoten)) continue;

          const totalWage = iWage !== -1 ? Number(row[iWage]) : 0;
          if (!totalWage || totalWage === 0) continue;

          const remarks = iRemarks !== -1 ? String(row[iRemarks] || "") : "";
          if (!isObo && remarks && remarks.includes("【") && remarks.includes("曜日") && /\d{4}\/\d{2}\/\d{2}/.test(remarks)) {
            continue;
          }

          const rDate = row[iDate] instanceof Date ? row[iDate] : new Date(row[iDate]);
          if (isNaN(rDate.getTime()) || rDate < startDate || rDate > endDate) continue;

          const rawName = String(row[iName] || "").replace(/\s+/g, "").trim();
          if (!rawName || rawName.includes("休館")) continue;

          const docId = String(row[iId] || "").trim();
          const shinryo = iShinryo !== -1 ? String(row[iShinryo] || "").trim() : "";
          const isKitaOrKame = (kyoten === "北葛西" || kyoten === "亀有");
          const clinicKey = isKitaOrKame ? `${kyoten}_${shinryo}` : kyoten;

          let docMap = statsByClinicKey.get(clinicKey);
          if (!docMap) {
            docMap = new Map();
            statsByClinicKey.set(clinicKey, docMap);
          }

          const mapKey = docId || toStandardKanji_(rawName);
          let docStat = docMap.get(mapKey);
          if (!docStat) {
            docStat = {
              id: docId,
              name: rawName,
              normName: toStandardKanji_(rawName),
              totalCount: 0,
              totalAm: 0,
              totalPm1: 0,
              totalPm2: 0,
              frames: {
                0: { am: 0, pm1: 0, pm2: 0 },
                1: { am: 0, pm1: 0, pm2: 0 },
                2: { am: 0, pm1: 0, pm2: 0 },
                3: { am: 0, pm1: 0, pm2: 0 },
                4: { am: 0, pm1: 0, pm2: 0 },
                5: { am: 0, pm1: 0, pm2: 0 },
                6: { am: 0, pm1: 0, pm2: 0 }
              }
            };
            docMap.set(mapKey, docStat);
          }

          const st = parseHourFast(row[iStart]);
          const en = parseHourFast(row[iEnd]);
          if (st === null || en === null) continue;

          const isKitakasai = (kyoten === "北葛西");
          const fAm = (st < 13 && en > 9) ? 1 : 0;
          const fPm1 = (st < 18 && en > 15) ? 1 : 0;
          const fPm2 = (st < (isKitakasai ? 20 : 21) && en > 18) ? 1 : 0;
          const shiftSum = fAm + fPm1 + fPm2;

          if (shiftSum > 0) {
            const dow = rDate.getDay();
            docStat.frames[dow].am += fAm;
            docStat.frames[dow].pm1 += fPm1;
            docStat.frames[dow].pm2 += fPm2;
            docStat.totalAm += fAm;
            docStat.totalPm1 += fPm1;
            docStat.totalPm2 += fPm2;
            docStat.totalCount += shiftSum;
          }
        }
      }
    } catch (e) {
      console.log(`シフトブック読込スキップ (${ssId}): ` + e.message);
    }
  };

  processBook(EXTERNAL_IDS.KINMU_MONTHLY_1, false);
  processBook(EXTERNAL_IDS.KINMU_MONTHLY_2, false);
  processBook(EXTERNAL_IDS.KINMU_SHIFT_NEW, true);

  return { statsByClinicKey, shiftNewOnTargetDates };
}

function getTargetFrameKeys_(absenceTime, isKitakasai) {
  if (!absenceTime || !absenceTime.includes("-")) return ["am", "pm1", "pm2"];
  const [sStr, eStr] = absenceTime.split("-");
  const st = toMinutes_(sStr) / 60;
  const en = toMinutes_(eStr) / 60;
  const keys = [];
  if (st < 13 && en > 9) keys.push("am");
  if (st < 18 && en > 15) keys.push("pm1");
  if (st < (isKitakasai ? 20 : 21) && en > 18) keys.push("pm2");
  return keys.length > 0 ? keys : ["am", "pm1", "pm2"];
}
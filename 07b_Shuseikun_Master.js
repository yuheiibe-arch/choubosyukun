/**
 * ============================================================================
 * 07b_Shuseikun_Master.gs
 * 拠点マスタ（エリア・開院日）・動的時給表・医師マスタ構築ヘルパー
 * ============================================================================
 */

function cleanClinicBaseName_(raw) {
  return String(raw || "")
    .replace(/^[0-9０-９]+[\.．、\s ]*/, "")
    .replace(/[\(（].*?[\)）]/g, "")
    .trim();
}

function normalizeKanaKe_(name) {
  return String(name || "").replace(/ヶ/g, "ケ").replace(/が丘/g, "ヶ丘");
}

function normalizeOpenDateYmd_(rawVal) {
  var s = String(rawVal || "").trim();
  if (!s) return "";
  var m = s.match(/(\d{4})[\/\-年\.]\s*(\d{1,2})[\/\-月\.]\s*(\d{1,2})/);
  if (m) {
    return m[1] + "/" + ("0" + m[2]).slice(-2) + "/" + ("0" + m[3]).slice(-2);
  }
  if (typeof normalizeDateStr_ === "function") {
    var norm = normalizeDateStr_(s);
    if (/^\d{4}\/\d{2}\/\d{2}$/.test(norm)) return norm;
  }
  return "";
}

function buildClinicMasterMeta_(d5ClinicSheet) {
  var areaMap = {};
  var openDateMap = {};
  if (!d5ClinicSheet || d5ClinicSheet.getLastRow() < 2) {
    return { areaMap: areaMap, openDateMap: openDateMap };
  }

  var vals = d5ClinicSheet.getDataRange().getDisplayValues();
  var headers = vals[0].map(function(h) { return String(h || "").trim(); });

  var colArea = headers.indexOf("エリア");
  if (colArea === -1) colArea = 7;

  var colOpenDate = -1;
  for (var hIdx = 0; hIdx < headers.length; hIdx++) {
    if (headers[hIdx].indexOf("開院") !== -1 || headers[hIdx].indexOf("オープン") !== -1) {
      colOpenDate = hIdx;
      break;
    }
  }

  for (var r = 1; r < vals.length; r++) {
    var rawName = vals[r][0];
    var cleanName = cleanClinicBaseName_(rawName);
    if (!cleanName || cleanName === "拠点名") continue;

    var normKey = normalizeKanaKe_(cleanName);
    var altName = cleanName.replace(/ケ/g, "ヶ").replace(/が丘/g, "ヶ丘");

    var rawArea = String(vals[r][colArea] || "").trim();
    var area = "";
    if (rawArea.indexOf("関西") !== -1) area = "関西";
    else if (rawArea.indexOf("関東") !== -1) area = "関東";
    else area = rawArea;

    if (area) {
      areaMap[cleanName] = area;
      areaMap[normKey] = area;
      areaMap[altName] = area;
    }

    if (colOpenDate !== -1) {
      var openYmd = normalizeOpenDateYmd_(vals[r][colOpenDate]);
      if (openYmd) {
        openDateMap[cleanName] = openYmd;
        openDateMap[normKey] = openYmd;
        openDateMap[altName] = openYmd;
      }
    }
  }

  return { areaMap: areaMap, openDateMap: openDateMap };
}

function isClinicOpenOnDate_(clinicName, targetDateStr, openDateMap) {
  if (!clinicName || !targetDateStr || !openDateMap) return true;
  var clean = cleanClinicBaseName_(clinicName);
  var openYmd = openDateMap[clean] || openDateMap[normalizeKanaKe_(clean)] || "";
  if (openYmd && targetDateStr < openYmd) {
    return false;
  }
  return true;
}

function sanitizeExtractedDoctorName_(rawName, evidenceText, doctorMasterList) {
  var s = String(rawName || "").replace(/先生$/g, "").replace(/[\s ]+/g, "").trim();
  if (!s) return "";
  if (/^(ああ+|あー+|えー+|うー+ん?|ん+|ま+|そ+|ちょっと|一旦|要確認|未定|なし|不明|別アプローチ|個別アプローチ|ステイ)$/.test(s)) {
    return "";
  }
  if (resolveStaffByAlias_(s)) return "";
  var ev = String(evidenceText || "");
  var familyGuess = s.slice(0, 2);
  if (ev && familyGuess && ev.indexOf(familyGuess) === -1 && ev.indexOf(s) === -1) {
    return "";
  }
  if (Array.isArray(doctorMasterList) && doctorMasterList.length > 0) {
    var normS = toStandardKanji_(s);
    var existsInMaster = doctorMasterList.some(function(d) {
      return d.normName.indexOf(normS) === 0 || d.name.indexOf(s) === 0;
    });
    if (!existsInMaster) return "";
  }
  return s;
}

function buildWageTablesForShuseikun_(wageSs) {
  var wageTablesObj = {};
  var parseYen = function(val) {
    var n = Number(String(val || "").replace(/[^0-9]/g, ""));
    return isNaN(n) ? 0 : n;
  };

  var allSheets = wageSs.getSheets();
  allSheets.forEach(function(sh) {
    var sName = sh.getName();
    if (sName.indexOf("時給") === -1 || sName.indexOf("医師") !== -1 || sName.indexOf("特別") !== -1) return;
    if (sh.getLastRow() < 2) return;

    var vals = sh.getDataRange().getDisplayValues();
    var headers = vals[0].map(function(h) { return String(h || "").trim(); });

    var colClinic = headers.indexOf("拠点名");
    var colDept = headers.indexOf("科目");
    var colWAm = headers.indexOf("平日_午前");
    var colSAm = headers.indexOf("土曜_午前");
    var colHAm = headers.indexOf("日祝_午前");

    if (colClinic === -1) colClinic = 1;
    if (colDept === -1) colDept = 2;
    if (colWAm === -1) colWAm = 3;
    if (colSAm === -1) colSAm = 6;
    if (colHAm === -1) colHAm = 9;

    var obj = {};
    for (var r = 1; r < vals.length; r++) {
      var row = vals[r];
      var base = cleanClinicBaseName_(row[colClinic]);
      if (!base || base === "拠点名") continue;

      var deptRaw = String(row[colDept] || "小児科").trim();
      var dept = (deptRaw.indexOf("内科") !== -1) ? "内科" : "小児科";

      var rateData = {
        weekday: { am: parseYen(row[colWAm]), pm: parseYen(row[colWAm + 1]), night: parseYen(row[colWAm + 2]) },
        sat:     { am: parseYen(row[colSAm]), pm: parseYen(row[colSAm + 1]), night: parseYen(row[colSAm + 2]) },
        sunHol:  { am: parseYen(row[colHAm]), pm: parseYen(row[colHAm + 1]), night: parseYen(row[colHAm + 2]) }
      };

      var isKitaOrKame = (base === "北葛西" || base === "亀有");
      obj[base + "|" + dept] = rateData;
      obj[base + "(" + dept + ")"] = rateData;
      obj[base + "（" + dept + "）"] = rateData;

      var altBase = base.replace(/ケ/g, "ヶ").replace(/が丘/g, "ヶ丘");
      var altBase2 = base.replace(/ヶ/g, "ケ").replace(/ヶ丘/g, "が丘");
      obj[altBase + "|" + dept] = rateData;
      obj[altBase2 + "|" + dept] = rateData;

      if (!isKitaOrKame) {
        obj[base] = rateData;
        obj[altBase] = rateData;
        obj[altBase2] = rateData;
      }
    }
    wageTablesObj[sName] = obj;
  });

  return wageTablesObj;
}

function buildDoctorMasterForShuseikun_(d3Data, clinicAreaMap) {
  var docMap = new Map();

  var addDocEntry = function(id, name, kana, deptRaw, areaHint) {
    var cleanName = String(name || "").replace(/[\s ]+/g, "").trim();
    if (!cleanName || cleanName.indexOf("休館") !== -1) return;
    var dept = (String(deptRaw || "").indexOf("内科") !== -1) ? "内科" : "小児科";
    var key = cleanName + "|" + dept;
    if (!docMap.has(key)) {
      docMap.set(key, {
        id: String(id || "").trim(),
        name: cleanName,
        normName: toStandardKanji_(cleanName),
        kana: String(kana || "").replace(/[\s ]+/g, "").trim(),
        dept: dept,
        areas: areaHint ? [areaHint] : []
      });
    } else if (areaHint) {
      var existing = docMap.get(key);
      if (existing.areas.indexOf(areaHint) === -1) {
        existing.areas.push(areaHint);
      }
    }
  };

  for (var i = 0; i < d3Data.length; i++) {
    var s = d3Data[i];
    var cArea = s.area || (clinicAreaMap ? (clinicAreaMap[s.clinicName] || clinicAreaMap[normalizeKanaKe_(s.clinicName)] || "") : "");
    addDocEntry(s.docId, s.docName, s.kanaName, s.dept, cArea);
  }

  if (typeof neg_fetchDoctorMasterData_ === "function") {
    try {
      var masterDocs = neg_fetchDoctorMasterData_();
      for (var j = 0; j < masterDocs.length; j++) {
        var md = masterDocs[j];
        addDocEntry(md.licenseNo, md.fullName || md.normFullName, md.kana || "", md.department, "");
      }
    } catch (e) {}
  }

  try {
    var ssNew = SpreadsheetApp.openById(EXTERNAL_IDS.KINMU_SHIFT_NEW);
    ["確定シフト", "応募シフト"].forEach(function(sName) {
      var sh = ssNew.getSheetByName(sName);
      if (!sh) return;
      var vals = sh.getDataRange().getDisplayValues();
      if (vals.length < 2) return;
      var h = vals[0];
      var iName = h.indexOf("名前");
      var iId = h.indexOf("医籍番号");
      var iDept = h.indexOf("診療科");
      var iKana = h.indexOf("名前（カタカナ）");
      var iClinic = h.indexOf("クリニック名");
      if (iClinic === -1) iClinic = h.indexOf("拠点名");
      if (iName === -1) return;
      for (var r = 1; r < vals.length; r++) {
        var cRaw = iClinic !== -1 ? cleanClinicBaseName_(vals[r][iClinic]) : "";
        var aHint = (cRaw && clinicAreaMap) ? (clinicAreaMap[cRaw] || clinicAreaMap[normalizeKanaKe_(cRaw)] || "") : "";
        addDocEntry(iId !== -1 ? vals[r][iId] : "", vals[r][iName], iKana !== -1 ? vals[r][iKana] : "", iDept !== -1 ? vals[r][iDept] : "小児科", aHint);
      }
    });
  } catch (e) {}

  return Array.from(docMap.values());
}
/**
 * ============================================================================
 * 06_Sheet_And_Date_Utils.gs
 * シート行制御・書式設定・日付/時間計算・文字列正規化の共通ユーティリティ群
 * ============================================================================
 */

function cleanUrl_(rawUrl) {
  return String(rawUrl || "")
    .replace(/^\[([^\]]+)\]\([^\)]+\)$/, "$1")
    .replace(/[\[\]\(\)]/g, "")
    .trim();
}

function normalizeClinicForMatch_(name) {
  const s = typeof cleanClinicBaseName_ === "function"
    ? cleanClinicBaseName_(name)
    : String(name || "").replace(/^[0-9０-９]+[\.．、\s ]*/, "").replace(/[\(（].*?[\)）]/g, "").trim();
  return s.replace(/[ヶケが]/g, "ケ").replace(/[\s ]+/g, "");
}

function writeDataExactRows_(sheet, rows, numCols) {
  if (!sheet) return;
  const dataCount = rows ? rows.length : 0;
  const requiredTotalRows = Math.max(1 + dataCount, 2);
  const currentMaxRows = sheet.getMaxRows();

  if (currentMaxRows < requiredTotalRows) {
    sheet.insertRowsAfter(currentMaxRows, requiredTotalRows - currentMaxRows);
  } else if (currentMaxRows > requiredTotalRows) {
    sheet.deleteRows(requiredTotalRows + 1, currentMaxRows - requiredTotalRows);
  }

  sheet.getRange(2, 1, requiredTotalRows - 1, numCols).clearContent().clearDataValidations();

  if (dataCount > 0) {
    sheet.getRange(2, 1, dataCount, numCols).setValues(rows);
  }

  trimBlankRowsAndCols_(sheet, numCols);
}

function trimBlankRowsAndCols_(sheet, expectedCols) {
  if (!sheet) return;
  const lastDataRow = Math.max(sheet.getLastRow(), 2);
  const maxRows = sheet.getMaxRows();
  if (maxRows > lastDataRow) {
    sheet.deleteRows(lastDataRow + 1, maxRows - lastDataRow);
  }

  if (expectedCols) {
    const maxCols = sheet.getMaxColumns();
    if (maxCols > expectedCols) {
      sheet.deleteColumns(expectedCols + 1, maxCols - expectedCols);
    }
  }
}

function appendRowAndTrimBlank_(sheet, rowArray, numCols) {
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  const isRow2Empty = lastRow <= 1 || String(sheet.getRange(2, 1).getDisplayValue() || "").trim() === "";
  if (isRow2Empty && lastRow <= 2) {
    sheet.getRange(2, 1, 1, rowArray.length).setValues([rowArray]);
  } else {
    sheet.insertRowAfter(lastRow);
    sheet.getRange(lastRow + 1, 1, 1, rowArray.length).setValues([rowArray]);
  }
  trimBlankRowsAndCols_(sheet, numCols || rowArray.length);
}

function applyLeftMiddleToSheet_(sheet) {
  if (!sheet) return;
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  sheet.getRange(1, 1, lastRow, lastCol)
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle");
}

function ensureSheet_(ss, name, headers, color) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(color)
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle");
  sheet.setFrozenRows(1);
  return sheet;
}

function normalizeDateStr_(val) {
  if (!val) return "";
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return "";
    return Utilities.formatDate(val, "Asia/Tokyo", "yyyy/MM/dd");
  }
  const s = String(val).trim().replace(/[年月\.\-]/g, "/").replace(/日.*$/, "");
  const m = s.match(/(\d{4})\/(\d{1,2})\/(\d{1,2})/);
  if (!m) return s;
  return `${m[1]}/${("0" + m[2]).slice(-2)}/${("0" + m[3]).slice(-2)}` ;
}

function normalizeTimeRange_(val) {
  if (!val || !String(val).includes("-")) return String(val || "").trim();
  const [s, e] = String(val).split("-");
  return `${padTime_(s)}-${padTime_(e)}`;
}

function padTime_(t) {
  const s = String(t || "").trim();
  const m = s.match(/^(\d{1,2}):(\d{2})/);
  if (!m) return s;
  return `${("0" + m[1]).slice(-2)}:${m[2]}`;
}

function toMinutes_(hhmm) {
  const m = String(hhmm || "").match(/^(\d{1,2}):(\d{2})/);
  if (!m) return NaN;
  return Number(m[1]) * 60 + Number(m[2]);
}

function isTimeOverlapping_(rangeA, rangeB) {
  if (!rangeA || !rangeB || !rangeA.includes("-") || !rangeB.includes("-")) return false;
  const [aS, aE] = rangeA.split("-").map(toMinutes_);
  const [bS, bE] = rangeB.split("-").map(toMinutes_);
  return Math.max(aS, bS) < Math.min(aE, bE);
}

function calcDiffHoursText_(curS, curE, conS, conE) {
  const curMin = calcWorkingMinutesExcludeBreak_(curS, curE);
  const conMin = calcWorkingMinutesExcludeBreak_(conS, conE);
  const diff = Math.abs(conMin - curMin);
  if (!diff) return "";
  const h = Math.floor(diff / 60);
  const m = diff % 60;
  if (h > 0 && m > 0) return `${h}時間${m}分`;
  if (h > 0) return `${h}時間`;
  return `${m}分`;
}

function calcWorkingMinutesExcludeBreak_(startStr, endStr) {
  const s = toMinutes_(startStr);
  const e = toMinutes_(endStr);
  if (isNaN(s) || isNaN(e) || e <= s) return 0;
  let mins = 0;
  for (let m = s; m < e; m += 15) {
    if (m >= 13 * 60 && m < 15 * 60) continue;
    mins += 15;
  }
  return mins;
}

function parseDate_(dateStr) {
  const norm = normalizeDateStr_(dateStr);
  const [y, m, d] = String(norm || "").split("/").map(Number);
  return new Date(y, (m || 1) - 1, d || 1);
}

function isTomorrowDate_(dateStr) {
  const today = new Date();
  const tomorrow = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1);
  const tStr = Utilities.formatDate(tomorrow, "Asia/Tokyo", "yyyy/MM/dd");
  return normalizeDateStr_(dateStr) === tStr;
}

function formatMdDay_(d) {
  if (!d || isNaN(d.getTime())) return "{M月D日(曜)}";
  const days = ["日", "月", "火", "水", "木", "金", "土"];
  return `${d.getMonth() + 1}月${d.getDate()}日(${days[d.getDay()]})`;
}

function formatYmdDay_(d) {
  if (!d || isNaN(d.getTime())) return "{YYYY年M月D日(曜)}";
  const days = ["日", "月", "火", "水", "木", "金", "土"];
  return `${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日(${days[d.getDay()]})`;
}

function formatComma_(num) {
  return Number(num || 0).toLocaleString("ja-JP");
}

function buildSlackThreadTitle_(dateStr) {
  if (!dateStr) return "超募集MTGログ";
  const d = parseDate_(dateStr);
  const y = d.getFullYear();
  const m = d.getMonth() + 1;
  const day = d.getDate();
  const lastDay = new Date(y, m, 0).getDate();
  const half = day <= 15 ? "前半(1~15日)" : `後半(16~${lastDay}日)`;
  return `【${y}年${m}月 ${half}】超募集MTGログ`;
}
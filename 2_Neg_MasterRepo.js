/**
 * ============================================================
 * 2_Neg_MasterRepo.gs
 * 医師マスタ・拠点マスタからのデータ取得とフィルタリング
 * ============================================================
 */

/**
 * 医師マスタから「姓名」「セイメイ」「医籍登録番号」「診療科目」を動的に取得する
 * ※ワクチン専任しか持っていない医師は検索対象から除外する
 */
function neg_fetchDoctorMasterData_() {
  const ss = SpreadsheetApp.openById(NEGOTIATION_CONFIG.DOCTOR_MASTER_SS_ID);
  const sh = ss.getSheetByName(NEGOTIATION_CONFIG.DOCTOR_MASTER_SHEET_NAME) || ss.getSheets()[0];
  const data = sh.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];

  const headers = data[0].map(h => h.trim());
  const fullNameIdx = headers.indexOf(NEGOTIATION_CONFIG.DOC_HEADERS.FULL_NAME);
  const lastNameIdx = headers.indexOf(NEGOTIATION_CONFIG.DOC_HEADERS.LAST_NAME);
  const firstNameIdx = headers.indexOf(NEGOTIATION_CONFIG.DOC_HEADERS.FIRST_NAME);
  const kanaNameIdx = headers.indexOf(NEGOTIATION_CONFIG.DOC_HEADERS.KANA_NAME);
  const licenseIdx = headers.indexOf(NEGOTIATION_CONFIG.DOC_HEADERS.LICENSE_NO);
  const deptIdx = headers.indexOf(NEGOTIATION_CONFIG.DOC_HEADERS.DEPARTMENT);

  const result = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    let fullName = fullNameIdx !== -1 ? String(row[fullNameIdx] || '').trim() : '';
    if (!fullName && lastNameIdx !== -1 && firstNameIdx !== -1) {
      fullName = `${row[lastNameIdx]} ${row[firstNameIdx]}`.trim();
    }
    if (!fullName) continue;

    // 診療科目から「ワクチン専任」を除外して整形
    const rawDept = deptIdx !== -1 ? String(row[deptIdx] || '').trim() : '';
    const cleanedDept = neg_cleanDoctorDepartments_(rawDept);

    // ワクチン専任のみ（除外後に空になる）の医師は対象外
    if (!cleanedDept) continue;

    const kanaName = kanaNameIdx !== -1 ? String(row[kanaNameIdx] || '').trim() : '';
    const licenseNo = licenseIdx !== -1 ? String(row[licenseIdx] || '').trim() : '';

    result.push({
      fullName: fullName,
      normFullName: neg_normalizeSpace_(fullName),
      normKanaName: neg_normalizeSpace_(kanaName),
      licenseNo: licenseNo,
      department: cleanedDept
    });
  }
  return result;
}

/**
 * 医師マスタの診療科目文字列から「ワクチン専任」を含む項目を除外する
 */
function neg_cleanDoctorDepartments_(rawDeptStr) {
  if (!rawDeptStr) return '';

  const parts = String(rawDeptStr)
    .split(/[、,，\/／\n]+/)
    .map(s => s.trim())
    .filter(Boolean);

  const validDepts = parts.filter(part => !part.includes('ワクチン専任'));
  return [...new Set(validDepts)].join('、');
}

/**
 * 拠点マスタから指定エリア（関東 / 関西）かつ開院日条件を満たす拠点を動的に取得する
 * ※ requiredDept に '内科' が指定された場合は、拠点マスタの「診療科」が「内科」の拠点のみ抽出
 */
function neg_fetchLocationsByAreaAndDept_(targetArea, requiredDept) {
  const ss = SpreadsheetApp.openById(NEGOTIATION_CONFIG.LOCATION_MASTER_SS_ID);
  const sh = ss.getSheetByName(NEGOTIATION_CONFIG.LOCATION_MASTER_SHEET_NAME);
  if (!sh) return [];

  const data = sh.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];

  const headers = data[0].map(h => h.trim());
  const nameIdx = headers.indexOf(NEGOTIATION_CONFIG.LOC_HEADERS.NAME);
  const deptIdx = headers.indexOf(NEGOTIATION_CONFIG.LOC_HEADERS.DEPARTMENT);
  const areaIdx = headers.indexOf(NEGOTIATION_CONFIG.LOC_HEADERS.AREA);
  const openDateIdx = headers.indexOf(NEGOTIATION_CONFIG.LOC_HEADERS.OPEN_DATE);

  const today = new Date();
  const oneMonthLater = new Date(today);
  oneMonthLater.setMonth(oneMonthLater.getMonth() + 1);
  oneMonthLater.setHours(23, 59, 59, 999);

  const locations = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const locName = String(row[nameIdx] || '').trim();
    const areaVal = String(row[areaIdx] || '').trim();
    const openDateStr = String(row[openDateIdx] || '').trim();
    const locDept = deptIdx !== -1 ? String(row[deptIdx] || '').trim() : '';

    if (!locName || !openDateStr) continue;
    if (areaVal !== targetArea) continue;
    if (requiredDept && locDept !== requiredDept) continue;

    const openDate = neg_parseJapaneseDate_(openDateStr);
    if (!openDate || openDate > oneMonthLater) continue;

    if (!locations.includes(locName)) {
      locations.push(locName);
    }
  }

  return locations;
}

/**
 * 日付文字列をDateオブジェクトに変換する
 */
function neg_parseJapaneseDate_(dateStr) {
  const normalized = String(dateStr)
    .trim()
    .replace(/[年月\.]/g, '/')
    .replace(/日/g, '');
  const d = new Date(normalized);
  if (isNaN(d.getTime())) return null;
  d.setHours(0, 0, 0, 0);
  return d;
}
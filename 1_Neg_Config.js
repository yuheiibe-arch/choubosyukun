/**
 * ============================================================
 * 1_Neg_Config.gs
 * 設定定義とシートの動的列マッピング
 * ============================================================
 */
const NEGOTIATION_CONFIG = {
  // 作業用スプレッドシート『超募集くん』のIDと対象シート名
  WORK_SS_ID: '1WNT1EWFP7lNzfO5hseFyfiw3BjK9SG42XrTFKMxz4kc',
  TARGET_SHEET_NAME: '交渉注意医師',

  // 医師マスタ
  DOCTOR_MASTER_SS_ID: '1-Ss0Zo1TZkl01d208_zVpG9DSLWOwLh7DntsvvK81Fc',
  DOCTOR_MASTER_SHEET_NAME: '医師マスタ',

  // 拠点マスタ
  LOCATION_MASTER_SS_ID: '14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs',
  LOCATION_MASTER_SHEET_NAME: '拠点名',

  // 交渉注意医師シートのヘッダー名
  HEADERS: {
    DOCTOR_NAME: '医師名',
    LICENSE_NO: '医籍番号',
    DEPARTMENT: '診療科目',
    CONDITION: '交渉条件',
    NG_AREA: 'NGエリア',
    NG_LOCATION: 'NG拠点',
    NG_DAY_PREFIX: 'NG曜日' // 「NG曜日1」「NG曜日2」などをすべて動的に認識します
  },

  // 医師マスタシートのヘッダー名
  DOC_HEADERS: {
    FULL_NAME: '姓名',
    LAST_NAME: '姓',
    FIRST_NAME: '名',
    KANA_NAME: 'セイメイ',
    LICENSE_NO: '医籍登録番号',
    DEPARTMENT: '診療科目'
  },

  // 拠点マスタシートのヘッダー名
  LOC_HEADERS: {
    NAME: '正規記載',
    DEPARTMENT: '診療科',
    AREA: 'エリア',
    OPEN_DATE: '開院日'
  }
};

/**
 * シートの1行目から各列の位置（1始まり）を動的に取得する
 */
function neg_getDynamicColumnMap_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return null;

  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(h => h.trim());

  const doctorNameCol = headers.indexOf(NEGOTIATION_CONFIG.HEADERS.DOCTOR_NAME) + 1;
  const licenseNoCol = headers.indexOf(NEGOTIATION_CONFIG.HEADERS.LICENSE_NO) + 1;

  let departmentCol = headers.indexOf(NEGOTIATION_CONFIG.HEADERS.DEPARTMENT) + 1;
  if (departmentCol === 0) {
    departmentCol = headers.indexOf('診療科') + 1;
  }

  const conditionCol = headers.indexOf(NEGOTIATION_CONFIG.HEADERS.CONDITION) + 1;
  const ngAreaCol = headers.indexOf(NEGOTIATION_CONFIG.HEADERS.NG_AREA) + 1;

  const ngLocationCols = [];
  const ngDayCols = [];

  headers.forEach((h, idx) => {
    if (h === NEGOTIATION_CONFIG.HEADERS.NG_LOCATION) {
      ngLocationCols.push(idx + 1);
    }
    if (h.startsWith(NEGOTIATION_CONFIG.HEADERS.NG_DAY_PREFIX)) {
      ngDayCols.push(idx + 1);
    }
  });

  return {
    lastCol,
    doctorNameCol,
    licenseNoCol,
    departmentCol,
    conditionCol,
    ngAreaCol,
    ngLocationCols,
    ngDayCols
  };
}

/**
 * 文字列の全角・半角スペースを除去して検索用に正規化する
 */
function neg_normalizeSpace_(str) {
  return String(str || '').replace(/[\s ]+/g, '');
}
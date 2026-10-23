/**
 * ============================================================================
 * 03a_Clinic_And_Scorer.gs
 * 拠点名マスタ（D5・保険辞書）の読込 ＆ 医師名の文字・カナ類似度スコア計算
 * ============================================================================
 */

let CLINIC_DICT_MAP = new Map();
let CLINIC_META_MAP = new Map(); // { clinicName -> { area: "関東", group: "..." } }
let CUSTOM_TEXT_RULES = [];

function populateInitialClinicDictSheet_(sheet) {
  if (sheet.getLastRow() > 1) return;

  const initialRows = [
    ["拠点名", "天町", "天王町", "音声誤変換の保険"],
    ["拠点名", "天皇町", "天王町", "音声誤変換の保険"],
    ["拠点名", "天皇長", "天王町", "音声誤変換の保険"],
    ["拠点名", "雨", "亀有", "雨小児→亀有"],
    ["拠点名", "ためあり", "亀有", "音声誤変換の保険"],
    ["拠点名", "亀やあり", "亀有", "音声誤変換の保険"],
    ["拠点名", "今ざ", "今里", "音声誤変換の保険"],
    ["拠点名", "稲海岸", "稲毛海岸", "音声誤変換の保険"],
    ["拠点名", "大海山", "代官山", "音声誤変換の保険"],
    ["拠点名", "東川", "東品川", "音声誤変換の保険"],
    ["拠点名", "東長川", "東品川", "音声誤変換の保険"],
    ["拠点名", "星谷", "越谷レイクタウン", "音声誤変換の保険"],
    ["拠点名", "腰", "越谷レイクタウン", "音声誤変換の保険"],
    ["拠点名", "光が", "光が丘", "音声誤変換の保険"],
    ["拠点名", "八千", "八千代緑が丘", "音声誤変換の保険"],
    ["拠点名", "柏", "柏の葉", "音声誤変換の保険"],
    ["拠点名", "式", "志木", "音声誤変換の保険"],
    ["拠点名", "新鎌ヶ谷", "新鎌ケ谷", "表記揺れ統一"]
  ];

  sheet.getRange(2, 1, initialRows.length, initialRows[0].length).setValues(initialRows);
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 260);
}

function loadClinicDictionaries_(d5Sheet, d7DictSheet) {
  CLINIC_DICT_MAP = new Map();
  CLINIC_META_MAP = new Map();
  CUSTOM_TEXT_RULES = [];

  if (d5Sheet && d5Sheet.getLastRow() >= 2) {
    const vals = d5Sheet.getDataRange().getDisplayValues();
    const headers = vals[0];
    const iGroup = headers.indexOf("拠点グループ");
    const iArea = headers.indexOf("エリア");

    for (let r = 1; r < vals.length; r++) {
      const canonical = String(vals[r][0] || "").trim();
      if (!canonical) continue;
      CLINIC_DICT_MAP.set(canonical, canonical);
      for (let c = 1; c <= 4; c++) {
        const v = String(vals[r][c] || "").trim();
        if (v) CLINIC_DICT_MAP.set(v, canonical);
      }
      CLINIC_META_MAP.set(canonical, {
        group: iGroup >= 0 ? String(vals[r][iGroup] || "").trim() : "",
        area: iArea >= 0 ? String(vals[r][iArea] || "").trim() : ""
      });
    }
  }

  if (d7DictSheet && d7DictSheet.getLastRow() >= 2) {
    const rows = d7DictSheet.getRange(2, 1, d7DictSheet.getLastRow() - 1, 3).getDisplayValues();
    for (const [type, fromStr, toStr] of rows) {
      const f = String(fromStr || "").trim();
      const t = String(toStr || "").trim();
      if (!f || !t) continue;
      if (type.includes("拠点")) {
        CLINIC_DICT_MAP.set(f, t);
      } else {
        CUSTOM_TEXT_RULES.push({ from: f, to: t });
      }
    }
  }
}

function normalizeClinicByDict_(clinicStr) {
  const s = String(clinicStr || "").trim();
  if (!s) return "";
  if (CLINIC_DICT_MAP.has(s)) return CLINIC_DICT_MAP.get(s);
  for (const [k, v] of CLINIC_DICT_MAP.entries()) {
    if (k.length >= 2 && s.startsWith(k)) return v;
  }
  return s;
}

function toStandardKanji_(str) {
  if (!str) return "";
  let s = String(str)
    .normalize("NFKC")
    .replace(/[\s ]+/g, "")
    .replace(/[﨑嵜]/g, "崎")
    .replace(/髙/g, "高")
    .replace(/[瀧泷]/g, "滝")
    .replace(/[邊邉]/g, "辺")
    .replace(/[齋齊]/g, "斎")
    .replace(/澤/g, "沢")
    .replace(/廣/g, "広")
    .replace(/國/g, "国")
    .replace(/嶋/g, "島")
    .replace(/ヶ/g, "ケ");

  for (const r of CUSTOM_TEXT_RULES) {
    if (s.includes(r.from)) s = s.split(r.from).join(r.to);
  }
  return s;
}

function extractKatakanaPart_(str) {
  if (!str) return "";
  const converted = String(str).replace(/[\u3041-\u3096]/g, ch =>
    String.fromCharCode(ch.charCodeAt(0) + 0x60)
  );
  const m = converted.match(/[\u30A1-\u30FAー]+/g);
  return m ? m.join("") : "";
}

function extractKanjiPart_(str) {
  if (!str) return "";
  const m = toStandardKanji_(str).match(/[\u4E00-\u9FFF]+/g);
  return m ? m.join("") : "";
}

function calcDoctorSimilarityScore_(targetRaw, docRecord) {
  if (!targetRaw) return 0;

  const normTarget = toStandardKanji_(targetRaw);
  const targetKanji = extractKanjiPart_(normTarget);
  const targetKana = extractKatakanaPart_(normTarget);

  const docFull = docRecord.normName;
  const docLast = docRecord.normLast;
  const docFirst = docRecord.normFirst;
  const docKana = docRecord.kanaName;
  const docLastKana = docRecord.lastKana;

  if (docFull === normTarget) return 100;
  if (docLast && docLast === normTarget) return 92;
  if (docFull.startsWith(normTarget) && normTarget.length >= 2) return 90;
  if (docLast && normTarget.startsWith(docLast)) {
    let bonus = 0;
    const restKanji = targetKanji.replace(docLast, "");
    for (const ch of restKanji) {
      if (docFirst.includes(ch)) bonus += 8;
    }
    if (targetKana) {
      for (const ch of targetKana) {
        if (docKana.includes(ch)) bonus += 5;
      }
    }
    return 80 + Math.min(bonus, 18);
  }

  if (targetKanji.length > 0) {
    if (docLast && targetKanji.startsWith(docLast)) {
      let matchedChars = 0;
      for (const ch of targetKanji) {
        if (docFull.includes(ch)) matchedChars++;
      }
      let kanaBonus = 0;
      for (const ch of targetKana) {
        if (docKana.includes(ch)) kanaBonus += 4;
      }
      return 75 + (matchedChars * 5) + Math.min(kanaBonus, 12);
    }

    let commonKanji = 0;
    for (const ch of targetKanji) {
      if (docFull.includes(ch)) commonKanji++;
    }
    if (commonKanji >= 2 || (commonKanji === 1 && docFull[0] === targetKanji[0] && targetKanji.length === 1)) {
      const ratio = commonKanji / Math.max(targetKanji.length, 2);
      return Math.round(55 + ratio * 30);
    }
  }

  if (targetKana.length > 0 && targetKanji.length === 0) {
    if (docKana.startsWith(targetKana) || docLastKana === targetKana) return 85;
    if (docLastKana && docLastKana[0] === targetKana[0]) {
      let commonKana = 0;
      for (const ch of targetKana) {
        if (docKana.includes(ch)) commonKana++;
      }
      return 50 + commonKana * 12;
    }
  }

  return 0;
}
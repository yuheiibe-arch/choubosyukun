/**
 * ============================================================================
 * 01a_Config_And_Staff.gs（01系 3分割の1）
 * ・外部参照シートID・シート名定数・URL定義（コピペ時の自動リンク化防止済）
 * ・『担当者』シート（A〜D列）読込 ＆ 「未定／担当者未定」除外・署名フォーマット整形
 * ============================================================================
 */

var EXTERNAL_IDS = {
  D6_D3_UNFILLED_SS: "1cbeXWojsxNMhQUo1c6VflF5hLUJUyfuOXCFbGP5jJEA",
  D4_D5_WAGE_SS:     "14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs",
  D9_ALLOWANCE_SS:   "1Cw2iHRglJdUf4aeCaE04DhGVBMPpyQdhpiC8U_OdVtl",
  MENTION_LIST_SS:   "1Q1bsHIUysw27aaPD-wTOAKWBbN7scD9RQIIdERYrylg",
  KINMU_MONTHLY_1:   "1JhLJuxkp6T5hEkfultUE57pmRXBI_lkN6ROp2-CuhCc",
  KINMU_MONTHLY_2:   "10Z8jg4o7Ri9Ggf7u_jauyDDZa6bVnc-iuuuqmkNDafw",
  KINMU_SHIFT_NEW:   "1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co",
  KINMU_AGENCY:      "1Fd8uOCE1SKvLCIPjZZ7sE2rFsQFOhjVjHqoaqs-pXqE"
};

var MENTION_LIST_GID = 2012574647;
var MENTION_TARGET_HEADERS = ["シフト戦略チーム", "アルバイトスタッフ"];
var VACCINE_DEDICATED_HOURLY_WAGE = 13000;

var SHEET_NAMES = {
  D6_APPROACH:       "アプローチシート",
  D3_PASTED:         "貼付用",
  D5_CLINICS:        "拠点名",
  D7_REVIEW:         "文案・Slack投稿レビュー",
  D7_OFFLINE_REVIEW: "オフライン修正シート",
  D7_CONFIRM:        "確認リスト",
  D7_DICT:           "拠点名・誤変換辞書",
  D7_STAFF:          "担当者",
  D7_NEGOTIATION:    "交渉注意医師",
  D7_EXCLUDED:       "除外医師",
  D7_HISTORY:        "打診履歴",
  D7_HOLIDAY:        "祝日 (内閣府)",
  D7_LOG:            "取込・投稿ログ"
};

var CAO_HOLIDAY_CSV_URL = "https://" + ["www8", "cao", "go", "jp"].join(".") + "/chosei/shukujitsu/syukujitsu.csv";
var SLACK_CHANNEL_ID = "C0AJTMLF0UD";
var SLACK_NOTIFY_CHANNEL_ID = "C09TRHGU64D";
var CHATWORK_NOTIFY_ROOM_ID = "415381235";
var D7_REVIEW_SHEET_URL = "https://" + ["docs", "google", "com"].join(".") + "/spreadsheets/d/1WNT1EWFP7lNzfO5hseFyfiw3BjK9SG42XrTFKMxz4kc/edit?hl=ja&gid=1198962830#gid=1198962830";

var STAFF_ALIAS_MAP = new Map();
var STAFF_RECORDS_LIST = [];

function getCircledNumber_(num) {
  const circled = ["①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩"];
  return circled[num - 1] || `(${num})`;
}

/**
 * 「未定」「担当者未定」等の未割当プレースホルダーかどうかを判定するヘルパー
 */
function isUnassignedStaffPlaceholder_(val) {
  const s = String(val || "").replace(/[\s\u3000]+/g, "").trim();
  if (!s) return true;
  return /^(?:未定|担当者未定|\(未定\)|（未定）|\(未選択\)|（未選択）|\(担当不要\)|（担当不要）|あとで決定|\{担当者名\}|-|ー|―|なし)$/.test(s);
}

/**
 * D7内の『担当者』シート（A列:担当者氏名, B列:呼び名1, C列:呼び名2, D列:プルダウン名）を読み込む
 */
function loadStaffMasterFromD7_(d7Ss) {
  STAFF_ALIAS_MAP = new Map();
  STAFF_RECORDS_LIST = [];
  const dropdownNames = [];

  const registerAlias = function(aliasStr, dropdownName) {
    const clean = String(aliasStr || "").replace(/[\s\u3000]+/g, "").trim();
    if (!clean || isUnassignedStaffPlaceholder_(clean)) return;
    STAFF_ALIAS_MAP.set(clean, dropdownName);
    const stem = clean.replace(/(さん|氏|ちゃん|くん|先生)$/g, "");
    if (stem && !isUnassignedStaffPlaceholder_(stem)) {
      STAFF_ALIAS_MAP.set(stem, dropdownName);
    }
  };

  try {
    const targetSs = d7Ss || SpreadsheetApp.getActiveSpreadsheet();
    const sheet = targetSs ? targetSs.getSheetByName(SHEET_NAMES.D7_STAFF) : null;
    if (sheet && sheet.getLastRow() >= 2) {
      const vals = sheet.getDataRange().getDisplayValues();
      for (let r = 1; r < vals.length; r++) {
        const colA = String(vals[r][0] || "").trim();
        const colB = String(vals[r][1] || "").trim();
        const colC = String(vals[r][2] || "").trim();
        const colD = String(vals[r][3] || "").trim();
        if (!colA && !colD) continue;

        const surnameGuess = colA.split(/[\s\u3000]+/)[0].trim();
        const dropdownName = colD || surnameGuess || colA.replace(/[\s\u3000]+/g, "");
        if (!dropdownName || isUnassignedStaffPlaceholder_(dropdownName)) continue;

        if (!dropdownNames.includes(dropdownName)) {
          dropdownNames.push(dropdownName);
        }

        const rowAliases = [colA, surnameGuess, colB, colC, colD, `${dropdownName}さん`].filter(Boolean);
        rowAliases.forEach(function(a) { registerAlias(a, dropdownName); });

        STAFF_RECORDS_LIST.push({
          dropdownName: dropdownName,
          fullNameClean: colA.replace(/[\s\u3000]+/g, ""),
          aliases: rowAliases.map(function(a) { return a.replace(/[\s\u3000]+/g, ""); }).filter(Boolean)
        });
      }
    }
  } catch (e) {
    console.log("『担当者』シートの読込エラー: " + e.message);
  }

  return dropdownNames;
}

/**
 * 入力文字列（フルネーム・呼び名・名字）を『担当者』シートのD列「プルダウン名（名字）」に変換する
 */
function resolveStaffByAlias_(inputStr) {
  if (!inputStr || isUnassignedStaffPlaceholder_(inputStr)) return "";
  if (STAFF_ALIAS_MAP.size === 0) {
    loadStaffMasterFromD7_(SpreadsheetApp.getActiveSpreadsheet());
  }

  const clean = String(inputStr)
    .replace(/\[To:\d+\]/gi, "")
    .replace(/[\s\u3000]+/g, "")
    .trim();
  if (!clean || isUnassignedStaffPlaceholder_(clean)) return "";

  if (STAFF_ALIAS_MAP.has(clean)) return STAFF_ALIAS_MAP.get(clean);
  const withoutHonorific = clean.replace(/(さん|氏|ちゃん|くん|先生)$/g, "");
  if (!withoutHonorific || isUnassignedStaffPlaceholder_(withoutHonorific)) return "";
  if (STAFF_ALIAS_MAP.has(withoutHonorific)) return STAFF_ALIAS_MAP.get(withoutHonorific);

  for (const rec of STAFF_RECORDS_LIST) {
    if (!rec.dropdownName) continue;
    if (
      withoutHonorific.startsWith(rec.dropdownName) ||
      (rec.fullNameClean && rec.fullNameClean.startsWith(withoutHonorific))
    ) {
      return rec.dropdownName;
    }
  }

  return "";
}

/**
 * MTG発言（evidence）の中に『担当者』シートの呼び名やフルネームが含まれていればD列「プルダウン名（名字）」を返す
 */
function detectStaffInEvidence_(evidenceText) {
  if (STAFF_RECORDS_LIST.length === 0) {
    loadStaffMasterFromD7_(SpreadsheetApp.getActiveSpreadsheet());
  }
  const ev = String(evidenceText || "").replace(/[\s\u3000]+/g, "");
  if (!ev) return "";

  for (const rec of STAFF_RECORDS_LIST) {
    for (const alias of rec.aliases) {
      if (alias === rec.dropdownName) continue;
      if (alias.length >= 2 && ev.includes(alias)) {
        return rec.dropdownName;
      }
    }
  }
  return "";
}

function extractCleanStaffName_(rawText) {
  if (!rawText || isUnassignedStaffPlaceholder_(rawText)) return "";
  const resolved = resolveStaffByAlias_(rawText);
  if (resolved) return resolved;

  const firstPart = String(rawText || "")
    .replace(/\[To:\d+\]/gi, "")
    .replace(/[\s\u3000]*さん$/g, "")
    .trim()
    .split(/[\s\u3000]+/)[0]
    .trim();

  if (!firstPart || isUnassignedStaffPlaceholder_(firstPart)) return "";
  return resolveStaffByAlias_(firstPart) || firstPart;
}

/**
 * 送信文案の「1行目（〇〇先生）直後の空行削除」と「末尾署名の全角スペース（ドクターサポート事務局 〇〇）」を強制整形する
 * ※担当者未設定時は「未定」ではなく「{担当者名}」を維持する
 */
function normalizeDraftMessageFormat_(draftText, staffNameRaw) {
  if (!draftText) return "";
  const cleanStaff = extractCleanStaffName_(staffNameRaw);
  let text = String(draftText);

  text = text.replace(/^([^\r\n]*?先生)[ \t]*\r?\n(?:\s*\r?\n)+/, "$1\n");

  const targetLabel = cleanStaff || "{担当者名}";
  text = text
    .replace(/ドクターサポート(事務局|部)の.*?でございます。/g, `ドクターサポート$1の${targetLabel}でございます。`)
    .replace(/ドクターサポート(事務局|部)の.*?です。/g, `ドクターサポート$1の${targetLabel}です。`)
    .replace(/ドクターサポート(事務局|部)[\s\u3000]*[^\r\n]+$/g, `ドクターサポート$1\u3000${targetLabel}`);

  return text;
}
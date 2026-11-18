/**
 * ============================================================================
 * 02a_1_Workspace_Setup.gs（旧02b_Setup 4分割の1）
 * ・レビュー表ヘッダー定義・D7ワークスペース初期セットアップ（setupD7Workspace）
 * ・時間帯カテゴリ推定・過去完了報告判定・根拠欄（V列）クリーンアップ
 * ============================================================================
 */

function getReviewHeadersCommon_() {
  return [
    "確認済", "状態", "確認事項・候補", "日付", "拠点名(D6突合キー)",
    "不在時間(D6突合キー)", "診療科", "アプローチ種別", "対象医師名", "現在シフト(拠点・時間)",
    "相談シフト(拠点・時間)", "特別時給", "現在日給(円)", "相談日給(円)", "依頼手当(円)",
    "担当者", "送信文案（社内システム貼付用）", "作業メモ（D6転記内容）", "投稿先スレッド", "投稿種別",
    "Slack投稿本文", "根拠（参照シート・MTG発言）", "投稿ID・処理日時", "エラー内容"
  ];
}

function ensureReviewStyleSheet_(ss, sheetName, headerColor) {
  const reviewHeaders = getReviewHeadersCommon_();
  const sh = ensureSheet_(ss, sheetName, reviewHeaders, headerColor || "#1a73e8");
  sh.setColumnWidth(9, 140);
  sh.setColumnWidth(17, 380);
  sh.setColumnWidth(18, 240);
  sh.setColumnWidth(21, 300);
  sh.setColumnWidth(22, 180);
  trimBlankRowsAndCols_(sh, reviewHeaders.length);
  applyLeftMiddleToSheet_(sh);
  return sh;
}

function setupD7Workspace() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureReviewStyleSheet_(ss, SHEET_NAMES.D7_REVIEW, "#1a73e8");

  const offlineSheetName = SHEET_NAMES.D7_OFFLINE_REVIEW || "オフライン修正シート";
  ensureReviewStyleSheet_(ss, offlineSheetName, "#0b8043");

  const confirmHeaders = [
    "取込日時", "日付", "文字起こし上の拠点名", "補正後の拠点名候補", "診療科判定",
    "時間帯", "アプローチ種別", "抽出医師名", "突合結果", "要確認の理由", "MTG発言抜粋"
  ];
  const confSheet = ensureSheet_(ss, SHEET_NAMES.D7_CONFIRM, confirmHeaders, "#e65100");
  trimBlankRowsAndCols_(confSheet, confirmHeaders.length);
  applyLeftMiddleToSheet_(confSheet);

  const historyHeaders = [
    "記録日時", "対象日付", "拠点名(D6突合キー)", "不在時間", "診療科",
    "アプローチ種別", "打診医師名", "担当者", "証跡・Slack投稿ID"
  ];
  const histSheet = ensureSheet_(ss, SHEET_NAMES.D7_HISTORY, historyHeaders, "#00897b");
  trimBlankRowsAndCols_(histSheet, historyHeaders.length);
  applyLeftMiddleToSheet_(histSheet);

  const dictHeaders = ["区分(拠点名 / その他)", "誤変換・揺れ表記（変換前）", "正規表記（変換後）", "メモ・備考"];
  const dictSheet = ensureSheet_(ss, SHEET_NAMES.D7_DICT, dictHeaders, "#8e24aa");
  populateInitialClinicDictSheet_(dictSheet);
  trimBlankRowsAndCols_(dictSheet, dictHeaders.length);
  applyLeftMiddleToSheet_(dictSheet);

  const holidayHeaders = ["祝日日付(YYYY/MM/DD)", "祝日名称", "取得日時"];
  const holSheet = ensureSheet_(ss, SHEET_NAMES.D7_HOLIDAY, holidayHeaders, "#0f9d58");
  updateCabinetOfficeHolidays_(holSheet);
  trimBlankRowsAndCols_(holSheet, holidayHeaders.length);
  applyLeftMiddleToSheet_(holSheet);

  const logHeaders = ["処理日時", "種別", "メールID / 投稿ID", "件名 / スレッド名", "処理件数", "詳細"];
  const logSheet = ensureSheet_(ss, SHEET_NAMES.D7_LOG, logHeaders, "#5f6368");
  trimBlankRowsAndCols_(logSheet, logHeaders.length);
  applyLeftMiddleToSheet_(logSheet);

  const def = ss.getSheetByName("シート1");
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) {
    ss.deleteSheet(def);
  }
  console.log("「超募集くん（D7）」の初期セットアップが完了しました！");
}

function inferTimeFromCategory_(timeCategory, isKitaKasai) {
  const tc = String(timeCategory || "").trim();
  const nightEnd = isKitaKasai ? "20:00" : "21:00";
  const amEnd = isKitaKasai ? "12:30" : "13:00";
  const pmStart = isKitaKasai ? "14:00" : "15:00";
  if (tc === "午前") return "09:00-" + amEnd;
  if (tc === "午後") return pmStart + "-18:00";
  if (tc === "夜間") return "18:00-" + nightEnd;
  if (tc === "午後・夜間") return pmStart + "-" + nightEnd;
  if (tc === "午前・午後") return "09:00-18:00";
  if (tc === "終日") return "09:00-" + nightEnd;
  return "";
}

function isPastCompletedReport_(evidenceText) {
  const s = String(evidenceText || "").trim();
  if (!s) return false;
  if (s.includes("ドキュメント確定") || s.includes("編集くんカスタム確定")) return false;
  const hasFutureIntent = /しましょう|ましょう|お願い|打診|聞いて|相談|調整し|いけるか|もらえるよう/.test(s);
  if (hasFutureIntent) return false;
  return /してもらいました|いただきました|埋まりました|充足しました|変更済み|移動済み|完了しました/.test(s);
}

function cleanVerboseEvidenceCell_(rawEvidence) {
  const firstLine = String(rawEvidence || "").split(/\r?\n/)[0] || "";
  let s = firstLine
    .replace(/\[MTG発言\]\s*/g, "")
    .replace(/\[ドキュメント確定[：:]([^\]]+)\]/g, "【$1】")
    .replace(/\[編集くんカスタム確定\]/g, "【カスタム】")
    .replace(/【複数枠おまとめ打診.*?$/g, "")
    .replace(/【選択的打診.*?$/g, "")
    .replace(/\[医籍[：:][^\]]+\]/g, "")
    .trim();
  return s || "-";
}
/**
 * ============================================================================
 * 02a_2_Consolidate_Rows.gs（旧02b_Setup 4分割の2）
 * 同一医師への複数枠の打診を **1行（primaryRow）に完全集約** し、2行目以降を削除する本体関数
 * （★担当者優先採用・「未定」除外・未特定医師の誤結合防止ガード付・省略なし完全版）
 * ============================================================================
 */

function consolidateMultiDateDoctorDrafts_(reviewRows, draftParamsByRowIdx) {
  // 1. まず全行の22列目（根拠欄）から余計なシステムログ（[D6]・[医師特定]・[時給表]など）を削除
  for (let i = 0; i < reviewRows.length; i++) {
    reviewRows[i][21] = cleanVerboseEvidenceCell_(reviewRows[i][21]);
  }

  const byDoctor = new Map();

  for (let i = 0; i < draftParamsByRowIdx.length; i++) {
    const p = draftParamsByRowIdx[i];
    if (!p || !p.doctorName || p.doctorName === "{医師名}" || p.doctorName.includes("\n")) continue;
    if (p.isDocResolved === false) continue;
    if (p.rowState !== "未レビュー") continue;
    if (p.approachType === "募集継続・ステイ" || p.approachType === "ステイ") continue;

    const cleanDocKey = String(p.doctorName).replace(/先生$/g, "").trim();
    if (!cleanDocKey) continue;

    if (!byDoctor.has(cleanDocKey)) {
      byDoctor.set(cleanDocKey, []);
    }
    byDoctor.get(cleanDocKey).push(p);
  }

  const rowIndicesToDelete = new Set();

  for (const [docName, rawEntries] of byDoctor.entries()) {
    if (rawEntries.length < 2) continue;

    // 同一枠の完全重複を除外しつつ時系列ソート
    const entries = [];
    rawEntries.forEach(function(p) {
      const isDup = entries.some(function(ex) {
        return ex.dateStr === p.dateStr &&
          ex.clinicName === p.clinicName &&
          ex.absenceStart === p.absenceStart &&
          ex.approachType === p.approachType;
      });
      if (!isDup) {
        entries.push(p);
      } else {
        rowIndicesToDelete.add(p.rowIndex);
      }
    });

    if (entries.length < 2) continue;

    entries.sort(function(a, b) {
      return a.dateStr.localeCompare(b.dateStr) ||
        String(a.consultStart || a.absenceStart || "").localeCompare(String(b.consultStart || b.absenceStart || ""));
    });

    // ★複数枠のいずれかに有効な担当者名が入っていれば、集約行の担当者として優先採用する
    let pickedStaff = "";
    for (let k = 0; k < entries.length; k++) {
      const st = extractCleanStaffName_(reviewRows[entries[k].rowIndex][15] || entries[k].staffName);
      if (st) {
        pickedStaff = st;
        break;
      }
    }
    if (pickedStaff) {
      entries.forEach(function(e) { e.staffName = pickedStaff; });
    }

    const uniqueDates = Array.from(new Set(entries.map(function(e) { return e.dateStr; })));
    const allMovers = entries.every(function(e) { return e.approachType === "勤務地変更相談"; });
    const allIndividuals = entries.every(function(e) { return e.approachType === "個別アプローチ"; });

    let combinedDraft = "";
    if (uniqueDates.length === 1 && allMovers) {
      combinedDraft = buildSameDateMultiClinicMoverTemplate_(entries);
    } else if (allIndividuals) {
      combinedDraft = buildMultiDateIndividualTemplate_(entries);
    } else {
      combinedDraft = buildUnifiedMultiSlotDraftTemplate_(entries);
    }

    // ★先頭の行（primaryIdx）に全枠の情報を集約し、2枠目以降の行はシートから削除する
    const primaryIdx = entries[0].rowIndex;
    const primaryRow = reviewRows[primaryIdx];

    // D列(3): 日付（改行区切り）
    primaryRow[3] = entries.map(function(e) { return e.dateStr; }).join("\n");
    // E列(4): 拠点名（改行区切り）
    primaryRow[4] = entries.map(function(e) { return reviewRows[e.rowIndex][4]; }).join("\n");
    // F列(5): 不在時間（改行区切り）
    primaryRow[5] = entries.map(function(e) { return reviewRows[e.rowIndex][5]; }).join("\n");
    // G列(6): 診療科
    const uniqueDepts = Array.from(new Set(entries.map(function(e) { return reviewRows[e.rowIndex][6]; }).filter(Boolean)));
    primaryRow[6] = uniqueDepts.join("・") || "小児科";
    // H列(7): アプローチ種別
    const uniqueTypes = Array.from(new Set(entries.map(function(e) { return e.approachType; }).filter(Boolean)));
    primaryRow[7] = uniqueTypes.join("・");

    // J列(9): 現在シフト
    const curShifts = Array.from(new Set(entries.map(function(e) { return reviewRows[e.rowIndex][9]; }).filter(function(v) { return v && v !== "-"; })));
    primaryRow[9] = curShifts.length > 0 ? curShifts.join("\n") : "-";

    // K列(10): 相談シフト（日付＋拠点＋時間で見やすく一覧化）
    primaryRow[10] = entries.map(function(e) { return `${e.dateStr.slice(5)} ${reviewRows[e.rowIndex][10]}`; }).join("\n");

    // L列(11): 特別時給
    const anySpecial = entries.some(function(e) { return String(reviewRows[e.rowIndex][11] || "").includes("特別時給"); });
    primaryRow[11] = anySpecial ? "特別時給" : "通常時給";

    // N列(13): 相談日給
    const wageLabels = entries.map(function(e) {
      const w = Number(String(reviewRows[e.rowIndex][13] || "0").replace(/,/g, ""));
      return w > 0 ? `${formatComma_(w)}円` : "-";
    });
    primaryRow[13] = wageLabels.every(function(w) { return w === "-"; }) ? "" : wageLabels.join("\n");

    // P列(15): 担当者（実名があれば採用、未定なら空欄）
    primaryRow[15] = pickedStaff;

    // Q列(16): 1通にまとめた送信文案
    primaryRow[16] = combinedDraft;

    // R列(17): 作業メモ（同日複数候補なら候補拠点名を付記、別日おまとめなら共通メモ）
    if (uniqueDates.length === 1 && allMovers) {
      const allClinicsLabel = entries.map(function(e) { return e.clinicName; }).join("・");
      if (primaryRow[17] && !String(primaryRow[17]).includes("候補：")) {
        primaryRow[17] += `（候補：${allClinicsLabel}）`;
      }
    }

    // U列(20): Slack投稿本文（全枠分を改行で統合）
    primaryRow[20] = entries.map(function(e) { return reviewRows[e.rowIndex][20]; }).filter(Boolean).join("\n\n");

    // V列(21): 根拠（シンプルに1行でおまとめ枠数を表示）
    const slotSummary = entries.map(function(e) { return `${e.dateStr.slice(5)} ${e.clinicName}`; }).join("・");
    primaryRow[21] = `おまとめ全${entries.length}枠（${slotSummary}）`;

    // 2枠目以降の行を削除対象に登録
    for (let k = 1; k < entries.length; k++) {
      rowIndicesToDelete.add(entries[k].rowIndex);
    }
  }

  // ★削除対象の行を後ろから splice して reviewRows と draftParamsByRowIdx から完全に除去
  const sortedDel = Array.from(rowIndicesToDelete).sort(function(a, b) { return b - a; });
  for (const delIdx of sortedDel) {
    reviewRows.splice(delIdx, 1);
    if (draftParamsByRowIdx && delIdx < draftParamsByRowIdx.length) {
      draftParamsByRowIdx.splice(delIdx, 1);
    }
  }
}

// 02_Main_Controller 互換用エイリアス
function consolidateMultiDateDoctorDraftsV2_(reviewRows, draftParamsByRowIdx) {
  return consolidateMultiDateDoctorDrafts_(reviewRows, draftParamsByRowIdx);
}
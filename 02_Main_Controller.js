/**
 * ============================================================================
 * 02_Main_Controller.gs
 * メイン実行エントリーポイント（スマートマージ取込・確認済保護・充足自動削除・おまとめ再構成対応版）
 * ============================================================================
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🤖 超募集くんAgent")
    .addItem("📥 ① 最新データを取り込む（差分マージ・充足整理）", "importLatestJsonSmart")
    .addItem("🚀 ② 確認済の行をSlackへ投稿（同日1枠まとめ更新）", "executeApprovedRows")
    .addItem("🕹️ ③ 修正くんを開く", "openShuseikunModal")
    .addSeparator()
    .addItem("🧹 充足枠のみ最新D6で自動整理", "refreshActiveSheetAgainstD6Only")
    .addToUi();
}

function importLatestJsonSmart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_NAMES.D7_LOG);
  const loggedIds = getLoggedEmailIdsFromD7Log_(logSheet);

  const offlineBatch = collectLatestJsonBatchFromGmail_(true);
  const voiceBatch = collectLatestJsonBatchFromGmail_(false);

  const hasNewOffline = Boolean(offlineBatch && offlineBatch.latestMsgId && !loggedIds.has(offlineBatch.latestMsgId));
  const hasNewVoice = Boolean(voiceBatch && voiceBatch.latestMsgId && !loggedIds.has(voiceBatch.latestMsgId));

  if (hasNewOffline || hasNewVoice) {
    if (hasNewOffline) runJsonImportCore_(true, { skipIfLogged: true });
    if (hasNewVoice) runJsonImportCore_(false, { skipIfLogged: true });
    return;
  }

  // 新着の未取込JSONメールがない場合は、既存行の「D6充足削除＆片方充足おまとめ再構成」のみ安全に実行する
  const activeName = ss.getActiveSheet() ? ss.getActiveSheet().getName() : "";
  const offlineSheetName = SHEET_NAMES.D7_OFFLINE_REVIEW || "オフライン修正シート";
  const isOffline = (activeName === offlineSheetName);
  const stats = runJsonImportCore_(isOffline, { pruneOnly: true });
  if (stats) {
    ss.toast(
      `新着JSONメールはありません。最新D6と照合し、充足削除:${stats.removedFilledOrReplacedCount}件 / 単枠再構成:${stats.rebuiltBundleCount}件 / 残存保護:${stats.keptCheckedCount + stats.keptUncheckedCount}件を整理しました。`,
      "🧹 充足整理完了",
      6
    );
  }
}

function refreshActiveSheetAgainstD6Only() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeName = ss.getActiveSheet() ? ss.getActiveSheet().getName() : "";
  const offlineSheetName = SHEET_NAMES.D7_OFFLINE_REVIEW || "オフライン修正シート";
  const isOffline = (activeName === offlineSheetName);
  const stats = runJsonImportCore_(isOffline, { pruneOnly: true });
  if (stats) {
    ss.toast(
      `充足削除:${stats.removedFilledOrReplacedCount}件 / おまとめ再構成:${stats.rebuiltBundleCount}件 / 確認済保護:${stats.keptCheckedCount}件 / 未送信維持:${stats.keptUncheckedCount}件`,
      "✅ D6充足整理完了",
      6
    );
  }
}

function sortReviewRowsAndParamsByDate_(reviewRows, draftParamsByRowIdx) {
  if (!reviewRows || reviewRows.length <= 1) return;

  const combined = reviewRows.map(function(row, idx) {
    return {
      row: row,
      param: draftParamsByRowIdx[idx] || null
    };
  });

  combined.sort(function(a, b) {
    const dA = normalizeDateStr_(a.row[3]);
    const dB = normalizeDateStr_(b.row[3]);
    if (dA !== dB) return dA.localeCompare(dB);

    const tA = toMinutes_(String(a.row[5] || "").split("-")[0]);
    const tB = toMinutes_(String(b.row[5] || "").split("-")[0]);
    const minA = isNaN(tA) ? 9999 : tA;
    const minB = isNaN(tB) ? 9999 : tB;
    return minA - minB;
  });

  for (let i = 0; i < combined.length; i++) {
    reviewRows[i] = combined[i].row;
    draftParamsByRowIdx[i] = combined[i].param;
    if (draftParamsByRowIdx[i]) {
      draftParamsByRowIdx[i].rowIndex = i;
    }
  }
}

function importLatestJsonAndBuildReview() {
  return runJsonImportCore_(false, { skipIfLogged: false });
}

function importOfflineDocJsonAndBuildSheet() {
  return runJsonImportCore_(true, { skipIfLogged: false });
}

function resolveExactDocNameByLicenseNo_(licenseNo, fallbackName, dateStr, d3Data) {
  const lic = String(licenseNo || "").trim();
  if (!lic) return fallbackName;

  if (Array.isArray(d3Data)) {
    const onDayHit = d3Data.find(function(s) {
      return s.dateStr === dateStr && String(s.docId || "").trim() === lic;
    });
    if (onDayHit && onDayHit.docName) return onDayHit.docName;

    const anyHit = d3Data.find(function(s) {
      return String(s.docId || "").trim() === lic;
    });
    if (anyHit && anyHit.docName) return anyHit.docName;
  }

  if (typeof neg_fetchDoctorMasterData_ === "function") {
    try {
      const masterDocs = neg_fetchDoctorMasterData_();
      const mHit = masterDocs.find(function(d) {
        return String(d.licenseNo || "").trim() === lic;
      });
      if (mHit && mHit.fullName) return mHit.fullName.replace(/[\s\u3000]+/g, "");
    } catch (e) {}
  }

  return fallbackName;
}

function runJsonImportCore_(isOfflineDocMode, optionsOpt) {
  const opts = optionsOpt || {};
  const d7Ss = SpreadsheetApp.getActiveSpreadsheet();
  const offlineSheetName = SHEET_NAMES.D7_OFFLINE_REVIEW || "オフライン修正シート";
  const targetReviewSheetName = isOfflineDocMode ? offlineSheetName : SHEET_NAMES.D7_REVIEW;
  const headerColor = isOfflineDocMode ? "#0b8043" : "#1a73e8";

  let revSheet = d7Ss.getSheetByName(targetReviewSheetName);
  let confSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_CONFIRM);
  let dictSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_DICT);
  let histSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_HISTORY);
  let logSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_LOG);
  const negSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_NEGOTIATION);
  const excludedSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_EXCLUDED) || d7Ss.getSheets().find(function(s) {
    return s.getSheetId() === 1246480688;
  });

  if (!confSheet || !dictSheet || !histSheet) {
    setupD7Workspace();
    confSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_CONFIRM);
    dictSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_DICT);
    histSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_HISTORY);
    logSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_LOG);
  }
  if (!revSheet) {
    revSheet = ensureReviewStyleSheet_(d7Ss, targetReviewSheetName, headerColor);
  }

  revSheet.getRange(1, 1).setValue("確認済");

  const unfilledSs = SpreadsheetApp.openById(EXTERNAL_IDS.D6_D3_UNFILLED_SS);
  const wageSs = SpreadsheetApp.openById(EXTERNAL_IDS.D4_D5_WAGE_SS);
  const d5ClinicSheet = wageSs.getSheetByName(SHEET_NAMES.D5_CLINICS);

  loadClinicDictionaries_(d5ClinicSheet, dictSheet);

  const d6Data = loadD6ApproachData_(unfilledSs.getSheetByName(SHEET_NAMES.D6_APPROACH));
  const d3Data = loadD3PastedShifts_(unfilledSs.getSheetByName(SHEET_NAMES.D3_PASTED));
  const holidaySet = loadHolidaySet_(d7Ss.getSheetByName(SHEET_NAMES.D7_HOLIDAY));
  const staffList = loadStaffMasterFromD7_(d7Ss);

  // 整理のみモード、または既取込メールスキップ判定
  let batchResult = null;
  if (!opts.pruneOnly) {
    batchResult = collectLatestJsonBatchFromGmail_(isOfflineDocMode);
    if (batchResult && opts.skipIfLogged) {
      const loggedIds = getLoggedEmailIdsFromD7Log_(logSheet);
      if (batchResult.latestMsgId && loggedIds.has(batchResult.latestMsgId)) {
        batchResult = null;
      }
    }
  }

  if (!batchResult || !batchResult.items || batchResult.items.length === 0) {
    return mergeAndWriteReviewSheetSmart_({
      revSheet: revSheet,
      incomingRows: [],
      d6Data: d6Data,
      d3Data: d3Data,
      wageSs: wageSs,
      holidaySet: holidaySet,
      staffList: staffList
    });
  }

  const items = batchResult.items;
  console.log(batchResult.msgCount + " 通のJSONメールから合計 " + items.length + " 件を読み込み、『" + targetReviewSheetName + "』へスマートマージします...");

  const nowStr = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");
  const baseDateYmd = batchResult.latestMsgDateYmd || Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");

  const reviewRows = [];
  const confirmRows = [];
  const draftParamsByRowIdx = [];
  const individualTargetRows = [];
  const doctorClinicRowMap = new Map();

  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const dateStr = normalizeDateStr_(item.date);
    let clinicName = normalizeClinicByDict_(item.clinic_name || "");
    let dept = (item.department || "小児科").trim();
    let targetTime = normalizeTimeRange_(item.target_time || "");
    let approachType = (item.approach_type || "").trim();
    let rawDocName = String(item.target_doctor || "")
      .replace(/\r\n|\r/g, "\n")
      .split("\n")[0]
      .replace(/(?:→|⇒|⇨).*$/g, "")
      .replace(/^(?:【?案\s*[123１２３]?】?\s*)+/g, "")
      .replace(/(?:[\s\u3000]*(?:他|ほか|外)\s*\d+\s*名.*$)/g, "")
      .replace(/先生$/g, "")
      .trim();

    const itemLicenseNo = String(item.license_no || "").trim();
    let rawAssignee = (item.assignee || "").trim();

    if (/^(?:バックアップ医師|関東バックアップ|関西バックアップ|バックアップ)$/.test(rawDocName)) {
      if (!approachType) approachType = "バックアップ勤務依頼";
      rawDocName = "";
    }

    if (itemLicenseNo) {
      rawDocName = resolveExactDocNameByLicenseNo_(itemLicenseNo, rawDocName, dateStr, d3Data);
    }

    if (dateStr && dateStr < baseDateYmd) {
      continue;
    }
    if (isPastCompletedReport_(item.evidence)) {
      continue;
    }

    const staffFromDoc = resolveStaffByAlias_(rawDocName);
    if (staffFromDoc && !itemLicenseNo) {
      rawAssignee = rawAssignee || staffFromDoc;
      rawDocName = "";
      if (!approachType || approachType === "要確認") {
        approachType = "個別アプローチ";
      }
    }

    if (String(item.evidence || "").includes("別アプローチ") && approachType === "募集継続・ステイ") {
      approachType = "個別アプローチ";
    }

    if (
      approachType === "募集継続・ステイ" ||
      approachType === "ステイ" ||
      /^(?:ステイ|募集継続・ステイ|特別時給.*ステイ|充足|応募あり充足)/.test(rawDocName)
    ) {
      continue;
    }
    if (rawDocName && (/[\(\)（）⭕❌📩⚠️]|000円|^両方|要確認|返信待ち|辞退|承諾/.test(rawDocName) || /^\d+$/.test(rawDocName))) {
      continue;
    }

    const resolvedClinic = resolveAmbiguousClinic_(
      clinicName, item.raw_clinic_name, dateStr, targetTime, rawDocName, item.from_clinic, d6Data, d3Data
    );
    clinicName = resolvedClinic.clinicName;

    if (!targetTime || !targetTime.includes("-")) {
      targetTime = inferTimeFromCategory_(item.time_category, clinicName === "北葛西") || targetTime;
    }

    const isKitaOrKame = (clinicName === "北葛西" || clinicName === "亀有");
    let d6MatchResult = matchD6RowsSmart_(d6Data, dateStr, clinicName, dept, targetTime, isKitaOrKame);
    let d6Row = d6MatchResult.mergedRow;

    if (isKitaOrKame && dept === "要確認" && d6Row && d6Row.dept) {
      dept = d6Row.dept;
    }

    // 提出後にすでにD6から消えて充足していた枠は新規追加しない
    if (!d6Row && dept !== "ワクチン専任" && d6Data.length > 0 && isOfflineDocMode) {
      continue;
    }

    let d6ClinicKey = d6Row
      ? d6Row.clinicRaw
      : (isKitaOrKame && dept !== "要確認" ? clinicName + " (" + dept + ")" : clinicName);
    let effectiveAbsenceTime = d6Row ? d6Row.absenceTime : targetTime;

    const docResolution = resolveDoctorAndShift_(
      rawDocName, approachType, dateStr, clinicName, dept,
      item.from_clinic || resolvedClinic.inferredFromClinic,
      d6Row, d3Data, effectiveAbsenceTime, item.evidence, itemLicenseNo
    );

    approachType = docResolution.resolvedApproachType || approachType;
    dept = docResolution.resolvedDept || dept;
    const effectiveFromClinic = docResolution.fromClinic || item.from_clinic || resolvedClinic.inferredFromClinic || "";
    const finalDocName = docResolution.isUnresolvedBackup ? "" : (docResolution.doctorName || rawDocName);

    const mergeKey = finalDocName ? [dateStr, d6ClinicKey, finalDocName].join("|") : "";
    let existingRowIdx = (mergeKey && doctorClinicRowMap.has(mergeKey)) ? doctorClinicRowMap.get(mergeKey) : -1;

    let combinedEvidence = String(item.evidence || "");
    if (itemLicenseNo && !combinedEvidence.includes(itemLicenseNo)) {
      combinedEvidence += " [医籍:" + itemLicenseNo + "]";
    }

    if (existingRowIdx !== -1) {
      const prevEv = doctorClinicRowMap.get(mergeKey + "_EVIDENCE") || "";
      combinedEvidence = prevEv.includes(combinedEvidence) ? prevEv : prevEv + " / " + combinedEvidence;

      const prevAbsence = reviewRows[existingRowIdx][5];
      const isNarrowingNow = /とりあえず|まずは|一旦/.test(String(item.evidence || ""));
      if (!isNarrowingNow && hasTimeExpansionIntentForMover_(combinedEvidence, effectiveFromClinic)) {
        const prevParts = String(prevAbsence || "-").split("-");
        const curParts = String(effectiveAbsenceTime || "-").split("-");
        const pS = prevParts[0], pE = prevParts[1];
        const cS = curParts[0], cE = curParts[1];
        const mergedS = (pS && cS) ? (toMinutes_(pS) <= toMinutes_(cS) ? pS : cS) : (pS || cS);
        const mergedE = (pE && cE) ? (toMinutes_(pE) >= toMinutes_(cE) ? pE : cE) : (pE || cE);
        if (mergedS && mergedE) {
          effectiveAbsenceTime = mergedS + "-" + mergedE;
          const expandedD6 = matchD6RowsSmart_(d6Data, dateStr, clinicName, dept, effectiveAbsenceTime, isKitaOrKame).mergedRow;
          if (expandedD6) {
            d6Row = expandedD6;
            d6ClinicKey = expandedD6.clinicRaw;
            effectiveAbsenceTime = expandedD6.absenceTime;
          }
        }
      }

      const reResolved = resolveDoctorAndShift_(
        rawDocName, approachType, dateStr, clinicName, dept,
        effectiveFromClinic, d6Row, d3Data, effectiveAbsenceTime, combinedEvidence, itemLicenseNo
      );
      Object.assign(docResolution, reResolved);
      approachType = docResolution.resolvedApproachType || approachType;
    }

    if (approachType === "勤務地変更相談" && !docResolution.hasTimeExpansion && docResolution.consultStart && docResolution.consultEnd) {
      const actualConsultRange = docResolution.consultStart + "-" + docResolution.consultEnd;
      if (actualConsultRange !== effectiveAbsenceTime) {
        const exactD6 = matchD6RowsSmart_(d6Data, dateStr, clinicName, dept, actualConsultRange, isKitaOrKame).mergedRow;
        if (exactD6) {
          d6Row = exactD6;
          d6ClinicKey = exactD6.clinicRaw;
          effectiveAbsenceTime = exactD6.absenceTime;
        }
      }
    }

    const isNoChangeShift =
      docResolution.currentShiftLabel &&
      docResolution.currentShiftLabel !== "要確認" &&
      docResolution.currentShiftLabel === docResolution.consultShiftLabel;

    if (isNoChangeShift) {
      continue;
    }

    const isSpecialFromMtg = Boolean(item.is_special_wage) || String(combinedEvidence).includes("特別時給");
    const isSpecialFromD6 = d6Row ? d6Row.isSpecialWage : false;

    const wageResult = calculateWagesAndAllowance_({
      wageSs: wageSs,
      dateStr: dateStr,
      clinicName: clinicName,
      dept: (dept === "ワクチン専任" ? "小児科" : dept),
      fromClinicName: effectiveFromClinic,
      approachType: approachType,
      absenceTime: effectiveAbsenceTime,
      currentStart: docResolution.currentStart,
      currentEnd: docResolution.currentEnd,
      consultStart: docResolution.consultStart,
      consultEnd: docResolution.consultEnd,
      isCurrentVaccine: Boolean(docResolution.isCurrentVaccine),
      isSpecialFromMtg: isSpecialFromMtg,
      isSpecialFromD6: isSpecialFromD6,
      isFullTime: docResolution.isFullTime,
      alreadyHasAllowance: docResolution.alreadyHasAllowance,
      holidaySet: holidaySet,
      specialWageHintText: String(item.special_wage_addon || item.memo || ""),
      evidenceText: combinedEvidence
    });

    const reasons = [];
    if (clinicName === "要確認" || !clinicName) {
      reasons.push(resolvedClinic.note || item.note_for_review || "拠点名が未確定");
    }
    if (isKitaOrKame && dept === "要確認") {
      reasons.push("北葛西・亀有の診療科が未確定");
    }
    if (!d6Row && dept !== "ワクチン専任") {
      reasons.push("D6アプローチシートに一致枠なし");
    }
    if (docResolution.status !== "OK") {
      reasons.push(docResolution.reason);
    }
    if (wageResult.status !== "OK") {
      reasons.push(wageResult.reason);
    }

    if (reasons.length > 0 && !isOfflineDocMode) {
      confirmRows.push([
        nowStr, dateStr, item.raw_clinic_name,
        resolvedClinic.candidatesText || clinicName,
        dept, effectiveAbsenceTime || item.time_category, approachType,
        finalDocName,
        d6Row ? "D6一致(" + d6Row.rowNumsText + "行目)" : "D6不一致",
        reasons.join(" / "), combinedEvidence
      ]);
      continue;
    }

    const rowState = "未レビュー";
    const warnNote = (isOfflineDocMode && reasons.length > 0) ? reasons.join(" / ") : "";
    const initialStaff = extractCleanStaffName_(
      resolveStaffByAlias_(rawAssignee) ||
      detectStaffInEvidence_(combinedEvidence) ||
      rawAssignee ||
      (d6Row ? d6Row.staff : "")
    );
    const targetRowIdx = existingRowIdx !== -1 ? existingRowIdx : reviewRows.length;

    let draftParams = null;
    let draftMessage = "";
    if (finalDocName || approachType === "個別アプローチ" || docResolution.isUnresolvedBackup) {
      const absParts = (effectiveAbsenceTime || "-").split("-");
      const absStart = absParts[0];
      const absEnd = absParts[1];
      draftParams = {
        rowIndex: targetRowIdx,
        rowState: rowState,
        isDocResolved: (docResolution.status === "OK"),
        dateStr: dateStr,
        isFullTime: docResolution.isFullTime,
        doctorName: finalDocName || "{医師名}",
        staffName: initialStaff,
        approachType: approachType,
        shiftType: "確定",
        dateObj: parseDate_(dateStr),
        isTomorrow: isTomorrowDate_(dateStr),
        clinicName: d6ClinicKey,
        currentClinicLabel: docResolution.currentClinicLabel || d6ClinicKey,
        fromClinicName: docResolution.currentClinicLabel || effectiveFromClinic || "現在拠点",
        absenceStart: absStart || "{不在開始}",
        absenceEnd: absEnd || "{不在終了}",
        currentStart: docResolution.currentStart || "{現在開始}",
        currentEnd: docResolution.currentEnd || "{現在終了}",
        consultStart: docResolution.consultStart || absStart || "{相談開始}",
        consultEnd: docResolution.consultEnd || absEnd || "{相談終了}",
        consultNoteSuffix: docResolution.consultNoteSuffix || "",
        diffHoursText: wageResult.diffHoursText || "{時間}",
        isMorningToNightExt: Boolean(wageResult.isMorningToNightExt),
        consultPm1DailyWage: wageResult.consultPm1DailyWage || 0,
        currentDailyWage: wageResult.currentDailyWage || 0,
        consultDailyWage: wageResult.consultDailyWage || 0,
        hasRequestAllow: (wageResult.requestAllowance || 0) > 0,
        isPublicRecruit: true
      };
      draftMessage = buildMessageTemplate_(draftParams);
    }

    const mdShort = formatSlackDateHeader_(dateStr);
    const targetDocLabel = finalDocName || (approachType === "個別アプローチ" ? "候補医師抽出対象" : "対象未定");
    const memoPlan = (targetDocLabel === "候補医師抽出対象")
      ? "候補医師抽出対象に個別アプローチ交渉"
      : buildNaturalWorkMemo_(approachType, targetDocLabel, effectiveFromClinic);
    const threadTitle = buildSlackThreadTitle_(dateStr);
    const slackBody = "`" + mdShort + "`\n```\n【" + d6ClinicKey + "】" + effectiveAbsenceTime + (wageResult.isSpecial ? "【特別時給】" : "") + "\n・" + memoPlan + "\n```";

    const sheetNameUsed = wageResult.sheetNameUsed ? wageResult.sheetNameUsed : "-";
    const wageNoteUsed = wageResult.wageNote ? wageResult.wageNote : "";
    const evidenceText = [
      "[MTG発言] " + combinedEvidence,
      d6Row ? "[D6] " + d6Row.rowNumsText + "行目一致" : "[D6] 未一致",
      docResolution.evidenceNote ? "[医師特定] " + docResolution.evidenceNote : "",
      "[時給表] " + sheetNameUsed + " (" + wageNoteUsed + ")"
    ].filter(Boolean).join("\n");

    const rowData = [
      false, rowState, warnNote, dateStr, d6ClinicKey,
      effectiveAbsenceTime, dept, approachType, finalDocName,
      docResolution.currentShiftLabel || "-", docResolution.consultShiftLabel || (d6ClinicKey + " " + effectiveAbsenceTime),
      wageResult.isSpecial ? "特別時給" : "通常時給",
      wageResult.currentDailyWage || "", wageResult.consultDailyWage || "", wageResult.requestAllowance || "",
      initialStaff,
      draftMessage, memoPlan, threadTitle, "通常", slackBody, evidenceText, "", ""
    ];

    if (existingRowIdx !== -1) {
      reviewRows[existingRowIdx] = rowData;
      draftParamsByRowIdx[existingRowIdx] = draftParams;
      doctorClinicRowMap.set(mergeKey + "_EVIDENCE", combinedEvidence);
    } else {
      if (mergeKey) {
        doctorClinicRowMap.set(mergeKey, reviewRows.length);
        doctorClinicRowMap.set(mergeKey + "_EVIDENCE", combinedEvidence);
      }
      if (approachType === "個別アプローチ" && !finalDocName && clinicName && clinicName !== "要確認") {
        individualTargetRows.push({
          rowIndex: reviewRows.length,
          dateStr: dateStr,
          clinicName: clinicName,
          dept: (dept === "ワクチン専任" ? "小児科" : dept),
          d6ClinicKey: d6ClinicKey,
          absenceTime: effectiveAbsenceTime,
          d6Row: d6Row
        });
      }
      reviewRows.push(rowData);
      draftParamsByRowIdx.push(draftParams);
    }
  }

  if (individualTargetRows.length > 0) {
    populateIndividualApproachCandidates_(
      reviewRows, draftParamsByRowIdx, individualTargetRows,
      d3Data, negSheet, excludedSheet, histSheet, d5ClinicSheet, holidaySet
    );
  }

  sortReviewRowsAndParamsByDate_(reviewRows, draftParamsByRowIdx);

  consolidateMultiDateDoctorDraftsV2_(reviewRows, draftParamsByRowIdx);

  for (let r = 0; r < reviewRows.length; r++) {
    const cleanStaff = extractCleanStaffName_(reviewRows[r][15]);
    reviewRows[r][15] = cleanStaff;
    if (reviewRows[r][16]) {
      reviewRows[r][16] = normalizeDraftMessageFormat_(reviewRows[r][16], cleanStaff);
    }
  }

  // ★既存行（確認済保護＋未送信維持＋充足削除＋おまとめ再構成）と新規行をマージ出力
  const mergeStats = mergeAndWriteReviewSheetSmart_({
    revSheet: revSheet,
    incomingRows: reviewRows,
    d6Data: d6Data,
    d3Data: d3Data,
    wageSs: wageSs,
    holidaySet: holidaySet,
    staffList: staffList
  });

  if (!isOfflineDocMode) {
    if (confirmRows.length > 1) {
      confirmRows.sort(function(a, b) {
        return normalizeDateStr_(a[1]).localeCompare(normalizeDateStr_(b[1]));
      });
    }
    writeDataExactRows_(confSheet, confirmRows, 11);
    applyLeftMiddleToSheet_(confSheet);
  }

  const logTypeLabel = isOfflineDocMode ? "JSONマージ取込(オフライン修正)" : "JSONマージ取込(口頭MTG)";
  const summaryMemo = `出力先:${targetReviewSheetName} (新規:${mergeStats.addedNewCount} / 確認済保護:${mergeStats.keptCheckedCount} / 未送信維持:${mergeStats.keptUncheckedCount} / 再構成:${mergeStats.rebuiltBundleCount} / 充足削除:${mergeStats.removedFilledOrReplacedCount})`;
  appendRowAndTrimBlank_(logSheet, [nowStr, logTypeLabel, batchResult.latestMsgId, batchResult.latestSubject, mergeStats.totalCount, summaryMemo], 6);
  applyLeftMiddleToSheet_(logSheet);

  sendImportCompletionNotifications_(targetReviewSheetName);
  return mergeStats;
}

function autoImportTriggerHandler() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_NAMES.D7_LOG);
  const loggedIds = getLoggedEmailIdsFromD7Log_(logSheet);

  const offlineBatch = collectLatestJsonBatchFromGmail_(true);
  if (offlineBatch && offlineBatch.latestMsgId && offlineBatch.items.length > 0 && !loggedIds.has(offlineBatch.latestMsgId)) {
    runJsonImportCore_(true, { skipIfLogged: true });
    loggedIds.add(offlineBatch.latestMsgId);
  }

  const voiceBatch = collectLatestJsonBatchFromGmail_(false);
  if (voiceBatch && voiceBatch.latestMsgId && voiceBatch.items.length > 0 && !loggedIds.has(voiceBatch.latestMsgId)) {
    runJsonImportCore_(false, { skipIfLogged: true });
  }
}
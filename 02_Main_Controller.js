/**
 * ============================================================================
 * 02_Main_Controller.gs
 * メイン実行エントリーポイント（シンプル3行メニュー・自動振り分け・医籍番号直結照合対応版）
 * ============================================================================
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🤖 超募集くんAgent")
    .addItem("📥 ① 最新データを取り込む", "importLatestJsonSmart")
    .addItem("🚀 ② 確認済の行をSlackへ投稿", "executeApprovedRows")
    .addItem("🕹️ ③ 修正くんを開く", "openShuseikunModal")
    .addToUi();
}

function importLatestJsonSmart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_NAMES.D7_LOG);
  const loggedIds = new Set();

  if (logSheet && logSheet.getLastRow() >= 2) {
    logSheet
      .getRange(2, 3, logSheet.getLastRow() - 1, 1)
      .getDisplayValues()
      .forEach(function(r) {
        const id = String(r[0] || "").trim();
        if (id) loggedIds.add(id);
      });
  }

  const offlineBatch = collectLatestJsonBatchFromGmail_(true);
  const voiceBatch = collectLatestJsonBatchFromGmail_(false);

  const hasNewOffline = Boolean(offlineBatch && offlineBatch.latestMsgId && !loggedIds.has(offlineBatch.latestMsgId));
  const hasNewVoice = Boolean(voiceBatch && voiceBatch.latestMsgId && !loggedIds.has(voiceBatch.latestMsgId));

  if (hasNewOffline || hasNewVoice) {
    if (hasNewOffline) runJsonImportCore_(true);
    if (hasNewVoice) runJsonImportCore_(false);
    return;
  }

  const activeName = ss.getActiveSheet() ? ss.getActiveSheet().getName() : "";
  const offlineSheetName = SHEET_NAMES.D7_OFFLINE_REVIEW || "オフライン修正シート";
  if (activeName === offlineSheetName) {
    runJsonImportCore_(true);
  } else {
    runJsonImportCore_(false);
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
  return runJsonImportCore_(false);
}

function importOfflineDocJsonAndBuildSheet() {
  return runJsonImportCore_(true);
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

function runJsonImportCore_(isOfflineDocMode) {
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

  const batchResult = collectLatestJsonBatchFromGmail_(isOfflineDocMode);
  if (!batchResult || batchResult.items.length === 0) {
    console.log("対象メールが見つかりませんでした。");
    return;
  }

  const items = batchResult.items;
  console.log(batchResult.msgCount + " 通のJSONメールから合計 " + items.length + " 件を読み込み、『" + targetReviewSheetName + "』へ展開します...");

  const unfilledSs = SpreadsheetApp.openById(EXTERNAL_IDS.D6_D3_UNFILLED_SS);
  const wageSs = SpreadsheetApp.openById(EXTERNAL_IDS.D4_D5_WAGE_SS);
  const d5ClinicSheet = wageSs.getSheetByName(SHEET_NAMES.D5_CLINICS);

  loadClinicDictionaries_(d5ClinicSheet, dictSheet);

  const d6Data = loadD6ApproachData_(unfilledSs.getSheetByName(SHEET_NAMES.D6_APPROACH));
  const d3Data = loadD3PastedShifts_(unfilledSs.getSheetByName(SHEET_NAMES.D3_PASTED));
  const holidaySet = loadHolidaySet_(d7Ss.getSheetByName(SHEET_NAMES.D7_HOLIDAY));
  const staffList = loadStaffMasterFromD7_(d7Ss);

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
    let rawDocName = (item.target_doctor || "").replace(/先生$/g, "").trim();
    const itemLicenseNo = String(item.license_no || "").trim();
    let rawAssignee = (item.assignee || "").trim();

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

    if (approachType === "募集継続・ステイ" || approachType === "ステイ") {
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
    const finalDocName = docResolution.doctorName || rawDocName;

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
      holidaySet: holidaySet
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
        docResolution.doctorName || rawDocName,
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
    if (finalDocName || approachType === "個別アプローチ") {
      const absParts = (effectiveAbsenceTime || "-").split("-");
      const absStart = absParts[0];
      const absEnd = absParts[1];
      draftParams = {
        rowIndex: targetRowIdx,
        rowState: rowState,
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

  consolidateMultiDateDoctorDrafts_(reviewRows, draftParamsByRowIdx);

  for (let r = 0; r < reviewRows.length; r++) {
    const cleanStaff = extractCleanStaffName_(reviewRows[r][15]);
    reviewRows[r][15] = cleanStaff;
    if (reviewRows[r][16]) {
      reviewRows[r][16] = normalizeDraftMessageFormat_(reviewRows[r][16], cleanStaff);
    }
  }

  writeDataExactRows_(revSheet, reviewRows, 24);
  if (reviewRows.length > 0) {
    revSheet.getRange(2, 1, reviewRows.length, 24).setBackground("#ffffff").setFontColor("#000000");
    revSheet.getRange(2, 1, reviewRows.length, 1).insertCheckboxes();

    const specialRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["特別時給", "通常時給"], true)
      .setAllowInvalid(false)
      .build();
    revSheet.getRange(2, 12, reviewRows.length, 1).setDataValidation(specialRule);

    if (staffList.length > 0) {
      const staffRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(staffList, true)
        .setAllowInvalid(true)
        .build();
      revSheet.getRange(2, 16, reviewRows.length, 1).setDataValidation(staffRule);
    }

    const postTypeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["通常", "引用して更新"], true)
      .setAllowInvalid(false)
      .build();
    revSheet.getRange(2, 20, reviewRows.length, 1).setDataValidation(postTypeRule);
  }
  applyLeftMiddleToSheet_(revSheet);

  if (!isOfflineDocMode) {
    if (confirmRows.length > 1) {
      confirmRows.sort(function(a, b) {
        return normalizeDateStr_(a[1]).localeCompare(normalizeDateStr_(b[1]));
      });
    }
    writeDataExactRows_(confSheet, confirmRows, 11);
    applyLeftMiddleToSheet_(confSheet);
  }

  const logTypeLabel = isOfflineDocMode ? "JSON取込(オフライン修正)" : "JSON取込(口頭MTG)";
  appendRowAndTrimBlank_(logSheet, [nowStr, logTypeLabel, batchResult.latestMsgId, batchResult.latestSubject, reviewRows.length, "出力先:" + targetReviewSheetName], 6);
  applyLeftMiddleToSheet_(logSheet);

  sendImportCompletionNotifications_(targetReviewSheetName);
}

function autoImportTriggerHandler() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_NAMES.D7_LOG);
  const loggedIds = new Set();

  if (logSheet && logSheet.getLastRow() >= 2) {
    logSheet
      .getRange(2, 3, logSheet.getLastRow() - 1, 1)
      .getDisplayValues()
      .forEach(function(r) {
        const id = String(r[0] || "").trim();
        if (id) loggedIds.add(id);
      });
  }

  const offlineBatch = collectLatestJsonBatchFromGmail_(true);
  if (offlineBatch && offlineBatch.latestMsgId && offlineBatch.items.length > 0 && !loggedIds.has(offlineBatch.latestMsgId)) {
    importOfflineDocJsonAndBuildSheet();
    loggedIds.add(offlineBatch.latestMsgId);
  }

  const voiceBatch = collectLatestJsonBatchFromGmail_(false);
  if (voiceBatch && voiceBatch.latestMsgId && voiceBatch.items.length > 0 && !loggedIds.has(voiceBatch.latestMsgId)) {
    importLatestJsonAndBuildReview();
  }
}
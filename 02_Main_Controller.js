/**
 * ============================================================================
 * 02_Main_Controller.gs
 * メイン実行エントリーポイント（修正くん起動メニュー・自然文作業メモ・自動通知＆トリガー対応版）
 * ============================================================================
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🤖 超募集くんAgent")
    .addItem("① 最新の【Agent解析JSON】メールを取り込む", "importLatestJsonAndBuildReview")
    .addItem("② 確認済の行をSlackへまとめて投稿", "executeApprovedRows")
    .addItem("③ 修正くんを開く（確認リスト補正パネル）", "openShuseikunModal")
    .addSeparator()
    .addItem("⚙️ D7シート構造の初期化・祝日更新", "setupD7Workspace")
    .addToUi();
}

function importLatestJsonAndBuildReview() {
  const d7Ss = SpreadsheetApp.getActiveSpreadsheet();
  let revSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_REVIEW);
  let confSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_CONFIRM);
  let dictSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_DICT);
  let histSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_HISTORY);
  let logSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_LOG);
  const negSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_NEGOTIATION);
  const excludedSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_EXCLUDED) || d7Ss.getSheets().find(s => s.getSheetId() === 1246480688);

  if (!revSheet || !confSheet || !dictSheet || !histSheet) {
    setupD7Workspace();
    revSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_REVIEW);
    confSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_CONFIRM);
    dictSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_DICT);
    histSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_HISTORY);
    logSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_LOG);
  }

  revSheet.getRange(1, 1).setValue("確認済");

  const batchResult = collectLatestJsonBatchFromGmail_();
  if (!batchResult || batchResult.items.length === 0) {
    console.log("【Agent解析JSON】メールが見つからないか、JSON配列をパースできませんでした。");
    return;
  }

  const items = batchResult.items;
  console.log(batchResult.msgCount + " 通のJSONメールから合計 " + items.length + " 件を読み込みました。外部シートと突合します...");

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
    let rawAssignee = (item.assignee || "").trim();

    if (dateStr && dateStr < baseDateYmd) {
      continue;
    }
    if (isPastCompletedReport_(item.evidence)) {
      continue;
    }

    const staffFromDoc = resolveStaffByAlias_(rawDocName);
    if (staffFromDoc) {
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
      d6Row, d3Data, effectiveAbsenceTime, item.evidence
    );

    approachType = docResolution.resolvedApproachType || approachType;
    dept = docResolution.resolvedDept || dept;
    const effectiveFromClinic = docResolution.fromClinic || item.from_clinic || resolvedClinic.inferredFromClinic || "";
    const finalDocName = docResolution.doctorName || rawDocName;

    const mergeKey = finalDocName ? [dateStr, d6ClinicKey, finalDocName].join("|") : "";
    let existingRowIdx = (mergeKey && doctorClinicRowMap.has(mergeKey)) ? doctorClinicRowMap.get(mergeKey) : -1;

    let combinedEvidence = String(item.evidence || "");
    if (existingRowIdx !== -1) {
      const prevEv = doctorClinicRowMap.get(mergeKey + "_EVIDENCE") || "";
      combinedEvidence = prevEv.includes(combinedEvidence) ? prevEv : prevEv + " / " + combinedEvidence;

      const prevAbsence = reviewRows[existingRowIdx][5];
      const isNarrowingNow = /とりあえず|まずは|一旦/.test(String(item.evidence || ""));
      if (!isNarrowingNow && hasTimeExpansionIntentForMover_(combinedEvidence, effectiveFromClinic)) {
        const [pS, pE] = String(prevAbsence || "-").split("-");
        const [cS, cE] = String(effectiveAbsenceTime || "-").split("-");
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
        effectiveFromClinic, d6Row, d3Data, effectiveAbsenceTime, combinedEvidence
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
      wageSs,
      dateStr,
      clinicName,
      dept: (dept === "ワクチン専任" ? "小児科" : dept),
      fromClinicName: effectiveFromClinic,
      approachType,
      absenceTime: effectiveAbsenceTime,
      currentStart: docResolution.currentStart,
      currentEnd: docResolution.currentEnd,
      consultStart: docResolution.consultStart,
      consultEnd: docResolution.consultEnd,
      isCurrentVaccine: Boolean(docResolution.isCurrentVaccine),
      isSpecialFromMtg,
      isSpecialFromD6,
      isFullTime: docResolution.isFullTime,
      alreadyHasAllowance: docResolution.alreadyHasAllowance,
      holidaySet
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

    if (reasons.length > 0) {
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
      const [absStart, absEnd] = (effectiveAbsenceTime || "-").split("-");
      draftParams = {
        rowIndex: targetRowIdx,
        rowState,
        dateStr,
        isFullTime: docResolution.isFullTime,
        doctorName: finalDocName || "{医師名}",
        staffName: initialStaff,
        approachType,
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
        currentDailyWage: wageResult.currentDailyWage,
        consultDailyWage: wageResult.consultDailyWage,
        hasRequestAllow: wageResult.requestAllowance > 0,
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
      false, rowState, "", dateStr, d6ClinicKey,
      effectiveAbsenceTime, dept, approachType, finalDocName,
      docResolution.currentShiftLabel, docResolution.consultShiftLabel,
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
      if (approachType === "個別アプローチ" && clinicName && clinicName !== "要確認") {
        individualTargetRows.push({
          rowIndex: reviewRows.length,
          dateStr,
          clinicName,
          dept: (dept === "ワクチン専任" ? "小児科" : dept),
          d6ClinicKey,
          absenceTime: effectiveAbsenceTime,
          d6Row
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

  consolidateMultiDateDoctorDrafts_(reviewRows, draftParamsByRowIdx);

  // ★全行の「担当者（P列）」をプルダウンの名字に統一し、「送信文案（Q列）」の1行目改行と末尾全角スペースを正規化する
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

  writeDataExactRows_(confSheet, confirmRows, 11);
  applyLeftMiddleToSheet_(confSheet);

  appendRowAndTrimBlank_(logSheet, [nowStr, "JSON取込", batchResult.latestMsgId, batchResult.latestSubject, reviewRows.length, "確認リスト:" + confirmRows.length + "件"], 6);
  applyLeftMiddleToSheet_(logSheet);

  sendImportCompletionNotifications_();
}

function autoImportTriggerHandler() {
  const batchResult = collectLatestJsonBatchFromGmail_();
  if (!batchResult || !batchResult.latestMsgId || batchResult.items.length === 0) {
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_NAMES.D7_LOG);
  if (logSheet && logSheet.getLastRow() >= 2) {
    const loggedIds = logSheet
      .getRange(2, 3, logSheet.getLastRow() - 1, 1)
      .getDisplayValues()
      .map(r => String(r[0] || "").trim());

    if (loggedIds.includes(batchResult.latestMsgId)) {
      return;
    }
  }

  importLatestJsonAndBuildReview();
}
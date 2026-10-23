/**
 * ============================================================================
 * 07_Shuseikun_Backend.gs
 * 「修正くん」メイン処理・D6作業メモ同期・レビューシート移動処理（個別候補抽出＆日付順挿入対応）
 * ============================================================================
 */

/**
 * 確認リストのG列（アプローチ種別）の表記揺れを標準種別名に正規化する
 */
function normalizeApproachTypeForShuseikun_(rawApproach, evidenceText) {
  var s = String(rawApproach || "").trim();
  var ev = String(evidenceText || "").trim();

  if (s.indexOf("移動") !== -1 || s.indexOf("変更") !== -1) return "勤務地変更相談";
  if (s.indexOf("個別") !== -1 || s.indexOf("追加") !== -1) return "個別アプローチ";
  if (s.indexOf("早出") !== -1 || s.indexOf("前倒") !== -1) return "早出相談";
  if (s.indexOf("延長") !== -1) return "延長相談";

  // G列が「要確認」や空欄の場合はMTG発言から推定
  if (ev.indexOf("移動") !== -1 || ev.indexOf("から") !== -1) return "勤務地変更相談";
  if (ev.indexOf("個別") !== -1 || ev.indexOf("別アプローチ") !== -1 || ev.indexOf("追加") !== -1) return "個別アプローチ";
  if (ev.indexOf("早出") !== -1 || ev.indexOf("前倒") !== -1) return "早出相談";
  if (ev.indexOf("延長") !== -1) return "延長相談";

  return "勤務地変更相談";
}

function buildNaturalWorkMemo_(approachType, doctorNameRaw, fromClinicRaw) {
  var docs = String(doctorNameRaw || "")
    .split(/\r?\n|[、,，・]+/)
    .map(function(s) { return s.replace(/先生$/g, "").trim(); })
    .filter(function(s) { return s && s !== "該当候補なし" && s !== "対象未定" && s !== "候補医師抽出対象" && s !== "{医師名}"; });

  var docLabel = docs.length > 0
    ? docs.map(function(d) { return d + "先生"; }).join("・")
    : "候補医師";

  var cleanFrom = cleanClinicBaseName_(fromClinicRaw);

  if (approachType === "勤務地変更相談") {
    var fromPart = cleanFrom ? ("【" + cleanFrom + "】から") : "";
    return docLabel + "に" + fromPart + "移動依頼交渉";
  }
  if (approachType === "延長相談") {
    return docLabel + "に延長依頼交渉";
  }
  if (approachType === "早出相談") {
    return docLabel + "に早出依頼交渉";
  }
  if (approachType === "個別アプローチ" || approachType === "追加勤務") {
    return docLabel + "に個別アプローチ交渉";
  }
  return docLabel + "に" + (approachType || "シフト調整") + "交渉";
}

function handleReviewCheckboxToggle_(sheet, row, isChecked) {
  var rowVals = sheet.getRange(row, 1, 1, 24).getDisplayValues()[0];
  var targetDate = normalizeDateStr_(rowVals[3]);
  var targetClinic = String(rowVals[4] || "").trim();
  var targetTime = normalizeTimeRange_(rowVals[5]);
  var approachType = String(rowVals[7] || "").trim();
  var docName = String(rowVals[8] || "").trim();
  var curShift = String(rowVals[9] || "").trim();
  var memoText = String(rowVals[17] || "").trim();

  if (!memoText || memoText.indexOf("Agent]") !== -1) {
    var fromClinic = curShift.split(/\s+/)[0] || "";
    memoText = buildNaturalWorkMemo_(approachType, docName, fromClinic);
    sheet.getRange(row, 18).setValue(memoText);
  }

  var rowRange = sheet.getRange(row, 1, 1, 24);
  if (isChecked) {
    sheet.getRange(row, 2).setValue("確認済み");
    sheet.getRange(row, 3).setValue("確認済み");
    rowRange.setBackground("#efefef").setFontColor("#666666");
    syncMemoToD6ApproachSheet_(targetDate, targetClinic, targetTime, memoText, true);
  } else {
    sheet.getRange(row, 2).setValue("未レビュー");
    sheet.getRange(row, 3).setValue("");
    rowRange.setBackground("#ffffff").setFontColor("#000000");
    syncMemoToD6ApproachSheet_(targetDate, targetClinic, targetTime, memoText, false);
  }
}

function syncMemoToD6ApproachSheet_(targetDate, targetClinic, targetTime, memoText, isAdd) {
  if (!targetDate || !targetClinic || !memoText) return;

  try {
    var unfilledSs = SpreadsheetApp.openById(EXTERNAL_IDS.D6_D3_UNFILLED_SS);
    var d6Sheet = unfilledSs.getSheetByName(SHEET_NAMES.D6_APPROACH);
    if (!d6Sheet) return;

    var lastRow = d6Sheet.getLastRow();
    var lastCol = d6Sheet.getLastColumn();
    if (lastRow < 3) return;

    var headers = d6Sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    var colDate = headers.indexOf("日付") + 1;
    var colClinic = headers.indexOf("拠点名") + 1;
    var colTime = headers.indexOf("不在時間") + 1;
    var colMemo = headers.indexOf("作業メモ") + 1;
    if (colDate === 0 || colClinic === 0 || colTime === 0 || colMemo === 0) return;

    var data = d6Sheet.getRange(1, 1, lastRow, Math.max(colDate, colClinic, colTime, colMemo)).getDisplayValues();
    var normTargetClinic = typeof normalizeClinicForMatch_ === "function"
      ? normalizeClinicForMatch_(targetClinic)
      : cleanClinicBaseName_(targetClinic).replace(/[ヶケが]/g, "ケ");

    for (var r = 2; r < data.length; r++) {
      var rDate = normalizeDateStr_(data[r][colDate - 1]);
      var rClinic = String(data[r][colClinic - 1] || "").trim();
      var normRClinic = typeof normalizeClinicForMatch_ === "function"
        ? normalizeClinicForMatch_(rClinic)
        : cleanClinicBaseName_(rClinic).replace(/[ヶケが]/g, "ケ");
      var rTime = normalizeTimeRange_(data[r][colTime - 1]);

      if (rDate !== targetDate) continue;
      if (rClinic !== targetClinic && normRClinic !== normTargetClinic) continue;
      if (targetTime && rTime && rTime !== targetTime && !isTimeOverlapping_(rTime, targetTime)) continue;

      var memoCell = d6Sheet.getRange(r + 1, colMemo);
      var currentMemo = String(memoCell.getDisplayValue() || "").trim();
      var lines = currentMemo ? currentMemo.split(/\r?\n/).map(function(l) { return l.trim(); }).filter(Boolean) : [];
      var cleanTargetMemo = memoText.replace(/^・/, "").trim();

      if (isAdd) {
        var alreadyExists = lines.some(function(l) {
          return l.replace(/^・/, "").trim() === cleanTargetMemo || l.indexOf(cleanTargetMemo) !== -1;
        });
        if (!alreadyExists) {
          lines.push(memoText);
          memoCell.setValue(lines.join("\n"));
        }
      } else {
        var filtered = lines.filter(function(l) {
          return l.replace(/^・/, "").trim() !== cleanTargetMemo && l !== memoText;
        });
        if (filtered.length !== lines.length) {
          memoCell.setValue(filtered.join("\n"));
        }
      }
    }
    SpreadsheetApp.flush();
  } catch (e) {
    console.error("作業メモ同期エラー: " + e.message);
  }
}

function openShuseikunModal() {
  var html = HtmlService.createTemplateFromFile("08_Shuseikun_UI")
    .evaluate()
    .setWidth(1200)
    .setHeight(820)
    .setTitle("🎮 修正くん - 確認リスト補正システム");
  SpreadsheetApp.getUi().showModalDialog(html, "🎮 修正くん（確認リスト補正パネル）");
}

function includeHtml_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getShuseikunInitData() {
  var d7Ss = SpreadsheetApp.getActiveSpreadsheet();
  var confSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_CONFIRM);
  var dictSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_DICT);
  var holSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_HOLIDAY);

  var unfilledSs = SpreadsheetApp.openById(EXTERNAL_IDS.D6_D3_UNFILLED_SS);
  var wageSs = SpreadsheetApp.openById(EXTERNAL_IDS.D4_D5_WAGE_SS);
  var d5ClinicSheet = wageSs.getSheetByName(SHEET_NAMES.D5_CLINICS);

  loadClinicDictionaries_(d5ClinicSheet, dictSheet);
  var clinicMeta = buildClinicMasterMeta_(d5ClinicSheet);
  var clinicAreaMap = clinicMeta.areaMap;
  var clinicOpenDateMap = clinicMeta.openDateMap;

  var staffList = loadStaffMasterFromD7_(d7Ss);
  var holidaySet = loadHolidaySet_(holSheet);
  var d6Data = loadD6ApproachData_(unfilledSs.getSheetByName(SHEET_NAMES.D6_APPROACH));
  var d3DataRaw = loadD3PastedShifts_(unfilledSs.getSheetByName(SHEET_NAMES.D3_PASTED));
  var wageTablesObj = buildWageTablesForShuseikun_(wageSs);

  var clinicSet = new Set();
  var addCleanClinic = function(raw) {
    if (!raw || String(raw).indexOf("|") !== -1) return;
    var clean = cleanClinicBaseName_(raw);
    if (!clean || clean === "拠点名" || clean === "クリニック名" || /^[0-9０-９]+$/.test(clean)) return;
    clinicSet.add(clean);
  };

  Object.keys(clinicAreaMap).forEach(addCleanClinic);
  Object.keys(wageTablesObj).forEach(function(sName) {
    Object.keys(wageTablesObj[sName] || {}).forEach(function(k) {
      addCleanClinic(k.split("|")[0]);
    });
  });
  for (var dIdx = 0; dIdx < d6Data.length; dIdx++) {
    if (d6Data[dIdx].baseClinic) addCleanClinic(d6Data[dIdx].baseClinic);
  }
  var CANONICAL_CLINICS = Array.from(clinicSet);

  var validClinicNormSet = new Set(CANONICAL_CLINICS.map(normalizeKanaKe_));
  var d3Data = [];
  for (var sIdx = 0; sIdx < d3DataRaw.length; sIdx++) {
    var sh = d3DataRaw[sIdx];
    var c = cleanClinicBaseName_(sh.clinicName);
    if (!c) continue;
    if (/有給|有休|休暇|欠勤|休診|休館|研修|待機|会議|面談|本部|事務/.test(c)) continue;
    if (!validClinicNormSet.has(normalizeKanaKe_(c))) continue;
    sh.clinicName = c;
    sh.area = clinicAreaMap[c] || clinicAreaMap[normalizeKanaKe_(c)] || "";
    d3Data.push(sh);
  }

  var doctorMasterList = buildDoctorMasterForShuseikun_(d3Data, clinicAreaMap);

  var confirmItems = [];
  if (confSheet && confSheet.getLastRow() >= 2) {
    var cVals = confSheet.getRange(2, 1, confSheet.getLastRow() - 1, 11).getDisplayValues();
    for (var i = 0; i < cVals.length; i++) {
      var r = cVals[i];
      if (!r[1] && !r[2] && !r[10]) continue;
      var dateStr = normalizeDateStr_(r[1]);
      var rawClinic = String(r[2] || "").trim();
      var candClinic = cleanClinicBaseName_(r[3]);
      var rawDept = String(r[4] || "小児科").trim();
      var dept = (rawDept === "内科") ? "内科" : "小児科";
      var rawTime = String(r[5] || "").trim();
      var rawApproach = String(r[6] || "").trim();
      var evidenceText = String(r[10] || "").trim();
      var cleanedDoc = sanitizeExtractedDoctorName_(r[7], evidenceText, doctorMasterList);

      var guessedClinic = "";
      for (var cIdx = 0; cIdx < CANONICAL_CLINICS.length; cIdx++) {
        var canName = CANONICAL_CLINICS[cIdx];
        if (!isClinicOpenOnDate_(canName, dateStr, clinicOpenDateMap)) continue;
        if (candClinic === canName || String(r[3] || "").indexOf(canName) !== -1) {
          guessedClinic = canName;
          break;
        }
      }
      if (!guessedClinic) {
        var normC = cleanClinicBaseName_(normalizeClinicByDict_(rawClinic));
        if (CANONICAL_CLINICS.indexOf(normC) !== -1 && isClinicOpenOnDate_(normC, dateStr, clinicOpenDateMap)) {
          guessedClinic = normC;
        }
      }

      var isKita = (guessedClinic === "北葛西");
      var defaultTime = normalizeTimeRange_(rawTime);
      if (!defaultTime || defaultTime.indexOf("-") === -1) {
        if (rawTime.indexOf("午前") !== -1) {
          defaultTime = "09:00-13:00";
        } else if (rawTime.indexOf("夜") !== -1 && rawTime.indexOf("午後") !== -1) {
          defaultTime = isKita ? "15:00-20:00" : "15:00-21:00";
        } else if (rawTime.indexOf("夜") !== -1) {
          defaultTime = isKita ? "18:00-20:00" : "18:00-21:00";
        } else {
          defaultTime = "15:00-18:00";
        }
      }

      var detectedStaff = extractCleanStaffName_(detectStaffInEvidence_(evidenceText) || "");
      var safeDisplayClinic = guessedClinic || (isClinicOpenOnDate_(candClinic, dateStr, clinicOpenDateMap) ? candClinic : "");
      var resolvedApproach = normalizeApproachTypeForShuseikun_(rawApproach, evidenceText);

      confirmItems.push({
        sheetRow: i + 2,
        importedAt: r[0],
        dateStr: dateStr,
        rawClinicName: safeDisplayClinic,
        origRawClinicName: rawClinic,
        candidatesText: safeDisplayClinic,
        guessedClinic: guessedClinic,
        dept: dept,
        rawTime: rawTime,
        targetTime: defaultTime,
        approachType: resolvedApproach,
        extractedDoctor: cleanedDoc,
        matchStatus: "",
        reason: "",
        evidence: evidenceText,
        staffName: detectedStaff
      });
    }
  }

  return {
    confirmItems: confirmItems,
    clinics: CANONICAL_CLINICS,
    clinicAreaMap: clinicAreaMap,
    clinicOpenDateMap: clinicOpenDateMap,
    staffList: staffList,
    holidays: Array.from(holidaySet),
    d6Data: d6Data,
    d3Data: d3Data,
    doctorMaster: doctorMasterList,
    wageTables: wageTablesObj
  };
}

/**
 * 『文案・Slack投稿レビュー』シートへ日付順（同日内は開始時刻順）になる位置を特定して1行挿入する
 */
function insertReviewRowSortedByDate_(revSheet, newRow, numCols) {
  var lastRow = revSheet.getLastRow();
  var isRow2Empty = (lastRow <= 1) || (lastRow === 2 && String(revSheet.getRange(2, 4).getDisplayValue() || "").trim() === "");

  if (isRow2Empty) {
    revSheet.getRange(2, 1, 1, newRow.length).setValues([newRow]);
    trimBlankRowsAndCols_(revSheet, numCols || newRow.length);
    return 2;
  }

  var existingVals = revSheet.getRange(2, 1, lastRow - 1, 6).getDisplayValues();
  var newDate = normalizeDateStr_(newRow[3]);
  var newTime = normalizeTimeRange_(newRow[5]);
  var newStartMin = toMinutes_(String(newTime || "").split("-")[0]);
  if (isNaN(newStartMin)) newStartMin = 9999;

  var insertBeforeSheetRow = -1;
  for (var i = 0; i < existingVals.length; i++) {
    var rDate = normalizeDateStr_(existingVals[i][3]);
    if (!rDate) continue;
    var rTime = normalizeTimeRange_(existingVals[i][5]);
    var rStartMin = toMinutes_(String(rTime || "").split("-")[0]);
    if (isNaN(rStartMin)) rStartMin = 9999;

    if (newDate < rDate || (newDate === rDate && newStartMin < rStartMin)) {
      insertBeforeSheetRow = i + 2;
      break;
    }
  }

  var targetRowNum;
  if (insertBeforeSheetRow !== -1) {
    revSheet.insertRowBefore(insertBeforeSheetRow);
    targetRowNum = insertBeforeSheetRow;
  } else {
    revSheet.insertRowAfter(lastRow);
    targetRowNum = lastRow + 1;
  }

  revSheet.getRange(targetRowNum, 1, 1, newRow.length)
    .setValues([newRow])
    .setBackground("#ffffff")
    .setFontColor("#000000");

  trimBlankRowsAndCols_(revSheet, numCols || newRow.length);
  return targetRowNum;
}

function commitShuseikunItemToReview(payload) {
  var d7Ss = SpreadsheetApp.getActiveSpreadsheet();
  var revSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_REVIEW);
  var confSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_CONFIRM);
  var logSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_LOG);
  var dictSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_DICT);
  var histSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_HISTORY);
  var negSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_NEGOTIATION);
  var excludedSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_EXCLUDED) || d7Ss.getSheets().find(function(s) { return s.getSheetId() === 1246480688; });

  var staffList = loadStaffMasterFromD7_(d7Ss);
  var dateStr = normalizeDateStr_(payload.dateStr);
  var clinicName = cleanClinicBaseName_(payload.clinicName);
  var dept = String(payload.dept || "小児科").trim();
  var isKitaOrKame = (clinicName === "北葛西" || clinicName === "亀有");
  var d6ClinicKey = isKitaOrKame ? (clinicName + " (" + dept + ")") : clinicName;
  var absenceTime = normalizeTimeRange_(payload.absenceTime);
  var approachType = String(payload.approachType || "").trim();
  var doctorName = String(payload.doctorName || "").replace(/先生$/g, "").trim();
  var fromClinic = cleanClinicBaseName_(payload.fromClinic);
  var currentShiftLabel = String(payload.currentShiftLabel || "-").trim();
  var consultShiftLabel = String(payload.consultShiftLabel || (d6ClinicKey + " " + absenceTime)).trim();
  var isSpecial = Boolean(payload.isSpecial);
  var currentDailyWage = Number(payload.currentDailyWage || 0);
  var consultDailyWage = Number(payload.consultDailyWage || 0);
  var requestAllowance = Number(payload.requestAllowance || 0);
  var staffName = extractCleanStaffName_(payload.staffName || "");
  var draftMessage = normalizeDraftMessageFormat_(payload.draftMessage || "", staffName);
  var evidence = String(payload.evidence || "");

  var isAutoExtractIndividual = (approachType === "個別アプローチ") && (!doctorName || doctorName === "{医師名}" || doctorName === "候補医師");
  var targetDocForMemo = isAutoExtractIndividual ? "候補医師抽出対象" : doctorName;
  var naturalMemo = isAutoExtractIndividual
    ? "候補医師抽出対象に個別アプローチ交渉"
    : buildNaturalWorkMemo_(approachType, targetDocForMemo, fromClinic);

  var threadTitle = buildSlackThreadTitle_(dateStr);
  var mdShort = formatSlackDateHeader_(dateStr);
  var slackBody = "`" + mdShort + "`\n```\n【" + d6ClinicKey + "】" + absenceTime + (isSpecial ? "【特別時給】" : "") + "\n・" + naturalMemo + "\n```";
  var evidenceFull = "[修正くん補正済]\n[MTG発言] " + evidence;

  var newRow = [
    false, "未レビュー", "修正くんから移動", dateStr, d6ClinicKey,
    absenceTime, dept, approachType, doctorName, currentShiftLabel,
    consultShiftLabel, isSpecial ? "特別時給" : "通常時給",
    currentDailyWage || "", consultDailyWage || "", requestAllowance || "",
    staffName, draftMessage, naturalMemo, threadTitle, "通常",
    slackBody, evidenceFull, "", ""
  ];

  // ★個別アプローチで医師名が未指定の場合は、過去実績・同曜日・除外条件・打診履歴から候補医師を自動抽出する
  if (isAutoExtractIndividual && clinicName && typeof populateIndividualApproachCandidates_ === "function") {
    try {
      var unfilledSs = SpreadsheetApp.openById(EXTERNAL_IDS.D6_D3_UNFILLED_SS);
      var wageSs = SpreadsheetApp.openById(EXTERNAL_IDS.D4_D5_WAGE_SS);
      var d5ClinicSheet = wageSs.getSheetByName(SHEET_NAMES.D5_CLINICS);
      loadClinicDictionaries_(d5ClinicSheet, dictSheet);

      var d3Data = loadD3PastedShifts_(unfilledSs.getSheetByName(SHEET_NAMES.D3_PASTED));
      var holidaySet = loadHolidaySet_(d7Ss.getSheetByName(SHEET_NAMES.D7_HOLIDAY));

      // 既存のレビューシート行も一時配列に含めて同日重複打診を除外できるようにする
      var tempReviewRows = [newRow];
      if (revSheet.getLastRow() >= 2) {
        var existingRows = revSheet.getRange(2, 1, revSheet.getLastRow() - 1, 24).getDisplayValues();
        for (var er = 0; er < existingRows.length; er++) {
          tempReviewRows.push(existingRows[er]);
        }
      }

      var absParts = String(absenceTime || "-").split("-");
      var draftParams = {
        isFullTime: false,
        doctorName: "{医師名}",
        staffName: staffName,
        approachType: approachType,
        shiftType: "確定",
        dateObj: parseDate_(dateStr),
        isTomorrow: isTomorrowDate_(dateStr),
        clinicName: d6ClinicKey,
        currentClinicLabel: d6ClinicKey,
        fromClinicName: fromClinic || "現在拠点",
        absenceStart: absParts[0] || "{不在開始}",
        absenceEnd: absParts[1] || "{不在終了}",
        currentStart: "{現在開始}",
        currentEnd: "{現在終了}",
        consultStart: absParts[0] || "{相談開始}",
        consultEnd: absParts[1] || "{相談終了}",
        consultNoteSuffix: "",
        diffHoursText: "",
        isMorningToNightExt: false,
        consultPm1DailyWage: 0,
        currentDailyWage: currentDailyWage,
        consultDailyWage: consultDailyWage,
        hasRequestAllow: requestAllowance > 0,
        isPublicRecruit: true
      };

      populateIndividualApproachCandidates_(
        tempReviewRows,
        [draftParams],
        [{
          rowIndex: 0,
          dateStr: dateStr,
          clinicName: clinicName,
          dept: (dept === "ワクチン専任" ? "小児科" : dept),
          d6ClinicKey: d6ClinicKey,
          absenceTime: absenceTime,
          d6Row: null
        }],
        d3Data, negSheet, excludedSheet, histSheet, d5ClinicSheet, holidaySet
      );

      newRow = tempReviewRows[0];
      if (newRow[16]) {
        newRow[16] = normalizeDraftMessageFormat_(newRow[16], staffName);
      }
    } catch (e) {
      console.error("個別アプローチ候補抽出エラー: " + e.message);
    }
  }

  // ★日付順（同日内は開始時刻順）の正しい行位置に挿入する
  var addedRowNum = insertReviewRowSortedByDate_(revSheet, newRow, 24);
  revSheet.getRange(addedRowNum, 1).insertCheckboxes();

  var specialRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["特別時給", "通常時給"], true)
    .setAllowInvalid(false)
    .build();
  revSheet.getRange(addedRowNum, 12).setDataValidation(specialRule);

  if (staffList.length > 0) {
    var staffRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(staffList, true)
      .setAllowInvalid(true)
      .build();
    revSheet.getRange(addedRowNum, 16).setDataValidation(staffRule);
  }

  var postTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["通常", "引用して更新"], true)
    .setAllowInvalid(false)
    .build();
  revSheet.getRange(addedRowNum, 20).setDataValidation(postTypeRule);
  applyLeftMiddleToSheet_(revSheet);

  deleteConfirmRowByMatch_(confSheet, payload.sheetRow, dateStr, payload.origRawClinicName || payload.rawClinicName, evidence);

  if (logSheet) {
    var nowStr = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");
    appendRowAndTrimBlank_(logSheet, [nowStr, "修正くん移動", "-", dateStr + " " + d6ClinicKey, 1, newRow[17]], 6);
  }

  return { ok: true, addedRowNum: addedRowNum };
}

function discardShuseikunConfirmItem(payload) {
  var d7Ss = SpreadsheetApp.getActiveSpreadsheet();
  var confSheet = d7Ss.getSheetByName(SHEET_NAMES.D7_CONFIRM);
  deleteConfirmRowByMatch_(confSheet, payload.sheetRow, dateStr, payload.origRawClinicName || payload.rawClinicName, payload.evidence);
  return { ok: true };
}

function deleteConfirmRowByMatch_(confSheet, sheetRowHint, dateStr, rawClinicName, evidence) {
  if (!confSheet || confSheet.getLastRow() < 2) return;
  var lastRow = confSheet.getLastRow();
  var vals = confSheet.getRange(2, 1, lastRow - 1, 11).getDisplayValues();
  var targetRowToDelete = -1;

  for (var i = 0; i < vals.length; i++) {
    var rDate = normalizeDateStr_(vals[i][1]);
    var rEv = String(vals[i][10] || "").trim();
    if (rDate === normalizeDateStr_(dateStr) && rEv === String(evidence || "").trim()) {
      targetRowToDelete = i + 2;
      break;
    }
  }

  if (targetRowToDelete === -1 && sheetRowHint >= 2 && sheetRowHint <= lastRow) {
    targetRowToDelete = sheetRowHint;
  }

  if (targetRowToDelete !== -1) {
    if (confSheet.getMaxRows() > 2) {
      confSheet.deleteRow(targetRowToDelete);
    } else {
      confSheet.getRange(2, 1, 1, 11).clearContent();
    }
    trimBlankRowsAndCols_(confSheet, 11);
  }
}
/**
 * ============================================================================
 * 4_Neg_Triggers.gs（または 06_Edit_Triggers.gs）
 * トリガー登録・onEditイベント制御（『交渉注意医師』＆『文案・Slack投稿レビュー』両対応）
 * ============================================================================
 */

/**
 * 編集時に自動で発火するメイン関数（インストーラブルトリガー）
 */
function handleNegotiationSheetEdit(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    const startRow = e.range.getRow();
    const startCol = e.range.getColumn();
    if (startRow <= 1) return;

    // ========================================================================
    // A. 『文案・Slack投稿レビュー』シートが編集された場合（超募集くん連動）
    // ========================================================================
    if (typeof SHEET_NAMES !== 'undefined' && sheetName === SHEET_NAMES.D7_REVIEW) {
      // 0. A列（1列目：確認済チェックボックス）ON/OFF ➔ 行グレーアウト＆D6作業メモ双方向同期
      if (startCol === 1) {
        const isChecked = (e.value === 'TRUE' || e.range.getValue() === true);
        handleReviewCheckboxToggle_(sheet, startRow, isChecked);
        return;
      }
      // 1. P列（16列目：担当者）変更 ➔ 名字に正規化＆送信文案の担当者名を即時更新
      if (startCol === 16) {
        updateStaffNameInDraftCell_(sheet, startRow, e.value);
        return;
      }
      // 2. L列（12列目：特別時給 / 通常時給）変更 ➔ 日給・文案・作業メモ・Slack本文を再計算
      if (startCol === 12) {
        recalculateRowBySpecialWageToggle_(sheet, startRow, e.value);
        return;
      }
      // 3. I列（9列目：対象医師名）変更 ➔ 送信文案の宛名（{医師名}先生）や作業メモ・Slack本文を更新
      if (startCol === 9) {
        updateDoctorNameInDraftAndMemo_(sheet, startRow);
        return;
      }
      return;
    }

    // ========================================================================
    // B. 『交渉注意医師』シートが編集された場合（既存スクリプト連動）
    // ========================================================================
    if (typeof NEGOTIATION_CONFIG !== 'undefined' && sheetName === NEGOTIATION_CONFIG.TARGET_SHEET_NAME) {
      const colMap = neg_getDynamicColumnMap_(sheet);
      if (!colMap) return;

      const numRows = e.range.getNumRows();
      const numCols = e.range.getNumColumns();

      for (let r = 0; r < numRows; r++) {
        const currentRow = startRow + r;
        for (let c = 0; c < numCols; c++) {
          const currentCol = startCol + c;

          // 1. 「医師名」列が編集された場合
          if (currentCol === colMap.doctorNameCol) {
            neg_processDoctorNameEdit_(sheet, currentRow, colMap);
          }

          // 2. 「NGエリア」「NG拠点」「診療科目」列のいずれかが編集された場合
          if (
            currentCol === colMap.ngAreaCol ||
            currentCol === colMap.departmentCol ||
            colMap.ngLocationCols.includes(currentCol)
          ) {
            neg_processNgLocationEdit_(sheet, currentRow, currentCol, colMap);
          }
        }
      }
    }
  } catch (err) {
    console.error('エラー発生: ' + err.stack);
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast('エラー: ' + err.message, 'スクリプトエラー', 5);
    } catch (_) {}
  }
}

/**
 * 【初回のみ実行】トリガー登録用
 */
function installTrigger() {
  const functionName = 'handleNegotiationSheetEdit';
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === functionName || t.getHandlerFunction() === 'handleEditTrigger') {
      ScriptApp.deleteTrigger(t);
    }
  });

  const workSs = (typeof NEGOTIATION_CONFIG !== 'undefined' && NEGOTIATION_CONFIG.WORK_SS_ID)
    ? SpreadsheetApp.openById(NEGOTIATION_CONFIG.WORK_SS_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  ScriptApp.newTrigger(functionName)
    .forSpreadsheet(workSs)
    .onEdit()
    .create();

  Logger.log(`✅ スプレッドシート「${workSs.getName()}」にトリガー（${functionName}）を設定しました！`);
}

/**
 * 【既存行への一括反映用】（『交渉注意医師』シート用）
 */
function applyDropdownsToExistingRows() {
  if (typeof NEGOTIATION_CONFIG === 'undefined') return;
  const workSs = SpreadsheetApp.openById(NEGOTIATION_CONFIG.WORK_SS_ID);
  const sheet = workSs.getSheetByName(NEGOTIATION_CONFIG.TARGET_SHEET_NAME);
  if (!sheet) return;

  const colMap = neg_getDynamicColumnMap_(sheet);
  const lastRow = sheet.getLastRow();
  const doctors = neg_fetchDoctorMasterData_();

  for (let r = 2; r <= lastRow; r++) {
    const docName = String(sheet.getRange(r, colMap.doctorNameCol).getDisplayValue() || '').trim();
    if (!docName) continue;

    neg_ensureRowDropdownsFromTemplate_(sheet, r, colMap);

    if (colMap.departmentCol > 0) {
      const deptCell = sheet.getRange(r, colMap.departmentCol);
      const currentDept = String(deptCell.getDisplayValue() || '').trim();
      const currentLicense = colMap.licenseNoCol > 0
        ? String(sheet.getRange(r, colMap.licenseNoCol).getDisplayValue() || '').trim()
        : '';

      const matchedDoc = doctors.find(d =>
        (currentLicense && d.licenseNo === currentLicense) ||
        d.normFullName === neg_normalizeSpace_(docName)
      );

      const deptValue = currentDept || (matchedDoc ? matchedDoc.department : '');
      if (deptValue) {
        neg_setDepartmentValueWithChipStyle_(sheet, r, colMap.departmentCol, deptValue);
        neg_applyInternalMedicineRuleIfNeeded_(sheet, r, deptValue, colMap);
      }
    }
  }
  Logger.log('✅ 既存行への「診療科目」カプセル型プルダウン体裁と値の反映が完了しました！');
}

/* =========================================================================
 * 『文案・Slack投稿レビュー』シート用の編集連動ヘルパー関数群
 * ========================================================================= */

function updateStaffNameInDraftCell_(sheet, row, newStaffRaw) {
  loadStaffMasterFromD7_(sheet.getParent());
  const rawVal = newStaffRaw !== undefined
    ? String(newStaffRaw || '')
    : String(sheet.getRange(row, 16).getDisplayValue() || '');
  const cleanStaff = extractCleanStaffName_(rawVal) || '{担当者名}';

  // フルネーム等が入っていた場合はP列自体もプルダウンの名字に正規化する
  if (rawVal && cleanStaff !== '{担当者名}' && rawVal.trim() !== cleanStaff) {
    sheet.getRange(row, 16).setValue(cleanStaff);
  }

  const draftCell = sheet.getRange(row, 17);
  const currentDraft = String(draftCell.getDisplayValue() || '');
  if (!currentDraft) return;

  const updatedDraft = currentDraft
    .replace(/^([^\r\n]*?先生)[ \t]*\r?\n(?:\s*\r?\n)+/, '$1\n')
    .replace(/ドクターサポート(事務局|部)の.*?でございます。/g, `ドクターサポート$1の${cleanStaff}でございます。`)
    .replace(/ドクターサポート(事務局|部)の.*?です。/g, `ドクターサポート$1の${cleanStaff}です。`)
    .replace(/ドクターサポート(事務局|部)[\s\u3000]*[^\n\r]+$/g, `ドクターサポート$1\u3000${cleanStaff}`);

  draftCell.setValue(updatedDraft);
}

function updateDoctorNameInDraftAndMemo_(sheet, row) {
  const rowVals = sheet.getRange(row, 1, 1, 24).getDisplayValues()[0];
  const dateStr = normalizeDateStr_(rowVals[3]);
  const d6ClinicKey = String(rowVals[4] || '').trim();
  const absenceTime = normalizeTimeRange_(rowVals[5]);
  const approachType = String(rowVals[7] || '').trim();
  const docNameRaw = String(rowVals[8] || '').trim();
  const curShiftLabel = String(rowVals[9] || '').trim();
  const isSpecial = (String(rowVals[11] || '').trim() === '特別時給');

  const lines = docNameRaw.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  const singleDocHonorific = (lines.length === 1 && lines[0] !== '該当候補なし')
    ? `${lines[0].replace(/先生$/g, '')}先生`
    : '{医師名}先生';

  const draftCell = sheet.getRange(row, 17);
  const currentDraft = String(draftCell.getDisplayValue() || '');
  if (currentDraft) {
    const updatedDraft = currentDraft
      .replace(/^[^\n\r]*?先生/m, singleDocHonorific)
      .replace(/^([^\r\n]*?先生)[ \t]*\r?\n(?:\s*\r?\n)+/, '$1\n');
    draftCell.setValue(updatedDraft);
  }

  if (dateStr && d6ClinicKey) {
    const mdShort = formatSlackDateHeader_(dateStr);
    const fromClinic = curShiftLabel.split(/\s+/)[0] || '';
    const memoPlan = buildNaturalWorkMemo_(approachType, docNameRaw, fromClinic);
    const slackBody = `\`${mdShort}\`\n\`\`\`\n【${d6ClinicKey}】${absenceTime}${isSpecial ? '【特別時給】' : ''}\n・${memoPlan}\n\`\`\``;
    sheet.getRange(row, 18).setValue(memoPlan);
    sheet.getRange(row, 21).setValue(slackBody);
  }
}

function recalculateRowBySpecialWageToggle_(sheet, row, newSpecialVal) {
  loadStaffMasterFromD7_(sheet.getParent());
  const rowVals = sheet.getRange(row, 1, 1, 24).getDisplayValues()[0];
  const dateStr = normalizeDateStr_(rowVals[3]);
  const d6ClinicKey = String(rowVals[4] || '').trim();
  const clinicName = d6ClinicKey.replace(/[\(（].*?[\)）]/g, '').trim();
  const absenceTime = normalizeTimeRange_(rowVals[5]);
  const dept = String(rowVals[6] || '小児科').trim();
  const approachType = String(rowVals[7] || '').trim();
  const docName = String(rowVals[8] || '').trim();
  const curShiftLabel = String(rowVals[9] || '').trim();
  const conShiftLabel = String(rowVals[10] || '').trim();
  const staffName = extractCleanStaffName_(rowVals[15]);

  if (!dateStr || !clinicName) return;

  const isSpecial = (newSpecialVal === '特別時給');
  const curTimeMatch = curShiftLabel.match(/(\d{1,2}:\d{2})-(\d{1,2}:\d{2})/);
  const conTimeMatch = conShiftLabel.match(/(\d{1,2}:\d{2})-(\d{1,2}:\d{2})/);
  const currentStart = curTimeMatch ? padTime_(curTimeMatch[1]) : '';
  const currentEnd = curTimeMatch ? padTime_(curTimeMatch[2]) : '';
  const consultStart = conTimeMatch ? padTime_(conTimeMatch[1]) : '';
  const consultEnd = conTimeMatch ? padTime_(conTimeMatch[2]) : '';

  const isCurrentVaccine = curShiftLabel.includes('ワクチン専任');
  const fromClinicName = curShiftLabel.split(/\s+/)[0].replace(/[\(（].*?[\)）]/g, '').trim() || clinicName;

  const ss = sheet.getParent();
  const holidaySet = loadHolidaySet_(ss.getSheetByName(SHEET_NAMES.D7_HOLIDAY));
  const wageSs = SpreadsheetApp.openById(EXTERNAL_IDS.D4_D5_WAGE_SS);

  const wageResult = calculateWagesAndAllowance_({
    wageSs,
    dateStr,
    clinicName,
    dept: (dept === 'ワクチン専任' ? '小児科' : dept),
    fromClinicName,
    approachType,
    absenceTime,
    currentStart,
    currentEnd,
    consultStart,
    consultEnd,
    isCurrentVaccine,
    isSpecialFromMtg: isSpecial,
    isSpecialFromD6: false,
    isFullTime: false,
    alreadyHasAllowance: !(Number(String(rowVals[14] || '0').replace(/,/g, '')) > 0),
    holidaySet
  });

  if (wageResult.status !== 'OK') return;

  sheet.getRange(row, 13).setValue(wageResult.currentDailyWage || '');
  sheet.getRange(row, 14).setValue(wageResult.consultDailyWage || '');

  const [absStart, absEnd] = (absenceTime || '-').split('-');
  const lines = docName.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  const singleDocName = (lines.length === 1 && lines[0] !== '該当候補なし') ? lines[0] : '{医師名}';

  const draftMessage = buildMessageTemplate_({
    isFullTime: false,
    doctorName: singleDocName,
    staffName,
    approachType,
    shiftType: '確定',
    dateObj: parseDate_(dateStr),
    isTomorrow: isTomorrowDate_(dateStr),
    clinicName: d6ClinicKey,
    currentClinicLabel: isCurrentVaccine ? `${clinicName}(ワクチン専任)` : d6ClinicKey,
    fromClinicName: isCurrentVaccine ? `${clinicName}(ワクチン専任)` : fromClinicName,
    absenceStart: absStart || '{不在開始}',
    absenceEnd: absEnd || '{不在終了}',
    currentStart: currentStart || '{現在開始}',
    currentEnd: currentEnd || '{現在終了}',
    consultStart: consultStart || absStart || '{相談開始}',
    consultEnd: consultEnd || absEnd || '{相談終了}',
    consultNoteSuffix: (isCurrentVaccine && dept !== 'ワクチン専任' && (approachType === '延長相談' || approachType === '早出相談'))
      ? `(※${absStart}~${absEnd}は${dept})`
      : '',
    diffHoursText: wageResult.diffHoursText || '{時間}',
    isMorningToNightExt: Boolean(wageResult.isMorningToNightExt),
    consultPm1DailyWage: wageResult.consultPm1DailyWage || 0,
    currentDailyWage: wageResult.currentDailyWage,
    consultDailyWage: wageResult.consultDailyWage,
    hasRequestAllow: Number(String(rowVals[14] || '0').replace(/,/g, '')) > 0,
    isPublicRecruit: true
  });
  sheet.getRange(row, 17).setValue(normalizeDraftMessageFormat_(draftMessage, staffName));

  const mdShort = formatSlackDateHeader_(dateStr);
  const memoPlan = buildNaturalWorkMemo_(approachType, docName, fromClinicName);
  const slackBody = `\`${mdShort}\`\n\`\`\`\n【${d6ClinicKey}】${absenceTime}${isSpecial ? '【特別時給】' : ''}\n・${memoPlan}\n\`\`\``;

  sheet.getRange(row, 18).setValue(memoPlan);
  sheet.getRange(row, 21).setValue(slackBody);
}
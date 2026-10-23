/**
 * ============================================================
 * 3_Neg_Handlers.gs
 * 医師名の段階的絞り込み・NG拠点連動・全列プルダウン自動コピー制御
 * ============================================================
 */

/**
 * ① 医師名の段階的絞り込み・候補プルダウン生成・医籍番号＆診療科目の自動取得
 */
function neg_processDoctorNameEdit_(sheet, row, colMap) {
  const ss = sheet.getParent();
  const nameCell = sheet.getRange(row, colMap.doctorNameCol);
  const licenseCell = colMap.licenseNoCol > 0 ? sheet.getRange(row, colMap.licenseNoCol) : null;
  const deptCell = colMap.departmentCol > 0 ? sheet.getRange(row, colMap.departmentCol) : null;
  const rawInput = String(nameCell.getDisplayValue() || '').trim();

  // 空欄に消去された場合はクリア
  if (!rawInput) {
    nameCell.clearDataValidations();
    if (licenseCell) licenseCell.clearContent();
    if (deptCell) deptCell.clearContent();
    return;
  }

  // ★2行目（基準行）にあるすべてのプルダウン体裁（診療科目・交渉条件・NGエリア・NG曜日1〜3・勤務不可時間など）を対象行へ一括コピー
  neg_ensureRowDropdownsFromTemplate_(sheet, row, colMap);

  const doctors = neg_fetchDoctorMasterData_();

  // パターンA: プルダウンから「氏名 (医籍番号)」の形式で選択された場合 → 確定
  const bracketMatch = rawInput.match(/^(.+?)\s*[\(（]([^\)）]+)[\)）]$/);
  if (bracketMatch) {
    const cleanName = bracketMatch[1].trim();
    const extractedLicense = bracketMatch[2].trim();

    const matchedDoc = doctors.find(d => d.licenseNo === extractedLicense) ||
                       doctors.find(d => d.normFullName === neg_normalizeSpace_(cleanName));
    const deptValue = matchedDoc ? matchedDoc.department : '';

    nameCell.clearDataValidations();
    nameCell.setValue(cleanName);
    if (licenseCell) licenseCell.setValue(extractedLicense);
    if (deptCell) neg_setDepartmentValueWithChipStyle_(sheet, row, colMap.departmentCol, deptValue);

    neg_applyInternalMedicineRuleIfNeeded_(sheet, row, deptValue, colMap);
    ss.toast(`「${cleanName}」（医籍番号: ${extractedLicense} / 診療科: ${deptValue}）をセットしました。`, '選択完了', 3);
    return;
  }

  // パターンB: 文字が入力された場合 → 部分一致検索
  const normalizedInput = neg_normalizeSpace_(rawInput);
  const partialMatches = doctors.filter(d =>
    d.normFullName.includes(normalizedInput) ||
    d.normKanaName.includes(normalizedInput)
  );

  if (partialMatches.length === 0) {
    nameCell.clearDataValidations();
    if (licenseCell) licenseCell.clearContent();
    if (deptCell) deptCell.clearContent();
    ss.toast(`「${rawInput}」に一致する医師が見つかりませんでした。`, '該当なし', 4);
    return;
  }

  // 候補が1名だけに絞られた場合 → 自動確定
  if (partialMatches.length === 1) {
    const doc = partialMatches[0];
    nameCell.clearDataValidations();
    nameCell.setValue(doc.fullName);
    if (licenseCell) licenseCell.setValue(doc.licenseNo);
    if (deptCell) neg_setDepartmentValueWithChipStyle_(sheet, row, colMap.departmentCol, doc.department);

    neg_applyInternalMedicineRuleIfNeeded_(sheet, row, doc.department, colMap);
    ss.toast(`「${doc.fullName}」（${doc.licenseNo} / ${doc.department}）を自動反映しました。`, '医師特定完了', 3);
    return;
  }

  // 候補が複数いる場合 → プルダウンをセット（さらなる追加入力も許可）
  const candidateList = partialMatches.slice(0, 500).map(d =>
    d.licenseNo ? `${d.fullName} (${d.licenseNo})` : d.fullName
  );
  const uniqueCandidates = [...new Set(candidateList)];

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(uniqueCandidates, true)
    .setAllowInvalid(true)
    .build();

  nameCell.setDataValidation(rule);
  if (licenseCell) licenseCell.clearContent();
  if (deptCell) deptCell.clearContent();

  ss.toast(
    `「${rawInput}」の候補が ${uniqueCandidates.length} 件あります。▼から選択するか、文字を追加して絞れます。`,
    '候補の絞り込み',
    5
  );
}

/**
 * 「診療科目」セルにカプセル型（チップ形式）のプルダウン体裁を維持したまま値をセットする
 */
function neg_setDepartmentValueWithChipStyle_(sheet, row, deptCol, deptValue) {
  if (deptCol <= 0) return;
  const deptCell = sheet.getRange(row, deptCol);

  if (deptCell.getDataValidation() === null) {
    const templateRow = neg_findTemplateRowWithValidation_(sheet, deptCol, row);
    if (templateRow > 0) {
      sheet.getRange(templateRow, deptCol)
        .copyTo(deptCell, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    }
  }

  const rule = deptCell.getDataValidation();
  if (rule && deptValue) {
    const currentOptions = (rule.getCriteriaValues()[0] || []).map(String);
    if (!currentOptions.includes(deptValue)) {
      const updatedOptions = [...new Set([...currentOptions, deptValue])];
      const newRule = rule.copy()
        .requireValueInList(updatedOptions, true)
        .setAllowInvalid(false)
        .build();
      deptCell.setDataValidation(newRule);
    }
  } else if (!rule && deptValue) {
    const fallbackRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['小児科', '内科', deptValue], true)
      .setAllowInvalid(false)
      .build();
    deptCell.setDataValidation(fallbackRule);
  }

  deptCell.setValue(deptValue);
}

/**
 * ② NGエリア・NG拠点の連動プルダウン更新（内科特別ルール＆選択済み除外対応）
 */
function neg_processNgLocationEdit_(sheet, row, editedCol, colMap) {
  if (colMap.ngAreaCol <= 0 || colMap.ngLocationCols.length === 0) return;

  const deptStr = colMap.departmentCol > 0
    ? String(sheet.getRange(row, colMap.departmentCol).getDisplayValue() || '').trim()
    : '';
  const isInternalMed = neg_isInternalMedicineDept_(deptStr);

  const areaCell = sheet.getRange(row, colMap.ngAreaCol);
  let selectedArea = String(areaCell.getDisplayValue() || '').trim();

  if (isInternalMed && selectedArea && selectedArea !== '関東') {
    areaCell.setValue('関東');
    selectedArea = '関東';
  }

  const ngCols = colMap.ngLocationCols;

  if (!selectedArea) {
    ngCols.forEach(col => {
      const cell = sheet.getRange(row, col);
      cell.clearContent();
      cell.clearDataValidations();
    });
    return;
  }

  const areaLocations = neg_fetchLocationsByAreaAndDept_(selectedArea, isInternalMed ? '内科' : '');

  if (editedCol === colMap.ngAreaCol || editedCol === colMap.departmentCol) {
    ngCols.forEach(col => {
      const cell = sheet.getRange(row, col);
      const val = String(cell.getDisplayValue() || '').trim();
      if (val && !areaLocations.includes(val)) {
        cell.clearContent();
      }
    });
  }

  const selectedSoFar = new Set();
  const styleTemplateCell = sheet.getRange(row, colMap.ngAreaCol);

  ngCols.forEach(col => {
    const cell = sheet.getRange(row, col);
    let currentVal = String(cell.getDisplayValue() || '').trim();

    if (currentVal && selectedSoFar.has(currentVal)) {
      cell.clearContent();
      currentVal = '';
    }

    const availableOptions = areaLocations.filter(loc => !selectedSoFar.has(loc));

    if (availableOptions.length > 0) {
      neg_setDropdownKeepChipStyle_(cell, availableOptions, styleTemplateCell);
    } else {
      cell.clearDataValidations();
    }

    if (currentVal) {
      selectedSoFar.add(currentVal);
    }
  });
}

/**
 * 診療科目が「内科」の場合、NGエリアを「関東」で固定しNG拠点（亀有・北葛西）を更新する
 */
function neg_applyInternalMedicineRuleIfNeeded_(sheet, row, departmentStr, colMap) {
  if (colMap.ngAreaCol <= 0) return;

  const isInternalMedicineOnly = neg_isInternalMedicineDept_(departmentStr);
  const areaCell = sheet.getRange(row, colMap.ngAreaCol);

  if (isInternalMedicineOnly) {
    const tplRow = neg_findTemplateRowWithValidation_(sheet, colMap.ngAreaCol, row);
    const templateCell = tplRow > 0 ? sheet.getRange(tplRow, colMap.ngAreaCol) : null;
    neg_setDropdownKeepChipStyle_(areaCell, ['関東'], templateCell);
    areaCell.setValue('関東');
    neg_processNgLocationEdit_(sheet, row, colMap.ngAreaCol, colMap);
  } else {
    const templateRow = neg_findTemplateRowWithMultipleAreaOptions_(sheet, colMap.ngAreaCol, row);
    if (templateRow > 0) {
      sheet.getRange(templateRow, colMap.ngAreaCol)
        .copyTo(areaCell, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    } else {
      neg_setDropdownKeepChipStyle_(areaCell, ['関東', '関西'], null);
    }
    if (String(areaCell.getDisplayValue() || '').trim()) {
      neg_processNgLocationEdit_(sheet, row, colMap.ngAreaCol, colMap);
    }
  }
}

function neg_isInternalMedicineDept_(deptStr) {
  const s = String(deptStr || '').trim();
  return s.includes('内科') && !s.includes('小児科');
}

/**
 * ★全列動的スキャン：
 * 基準行（2行目など）にプルダウンが設定されているすべての列（診療科目、交渉条件、NGエリア、NG曜日1〜3、勤務不可時間など）を
 * 動的に見つけ出し、対象行にカプセル型プルダウン体裁を丸ごとコピーする
 * （※医師名列・医籍番号列・NG拠点列は個別ロジックで動的制御するため除く）
 */
function neg_ensureRowDropdownsFromTemplate_(sheet, targetRow, colMap) {
  const lastCol = colMap.lastCol || sheet.getLastColumn();
  const skipCols = new Set([
    colMap.doctorNameCol,
    colMap.licenseNoCol,
    ...(colMap.ngLocationCols || [])
  ]);

  for (let col = 1; col <= lastCol; col++) {
    if (skipCols.has(col)) continue;

    const targetCell = sheet.getRange(targetRow, col);
    if (targetCell.getDataValidation() !== null) continue;

    const templateRow = neg_findTemplateRowWithValidation_(sheet, col, targetRow);
    if (templateRow > 0) {
      const templateCell = sheet.getRange(templateRow, col);
      templateCell.copyTo(targetCell, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    }
  }
}

function neg_findTemplateRowWithValidation_(sheet, col, excludeRow) {
  const maxCheckRow = Math.min(sheet.getLastRow(), 20);
  for (let r = 2; r <= maxCheckRow; r++) {
    if (r === excludeRow) continue;
    if (sheet.getRange(r, col).getDataValidation() !== null) return r;
  }
  return -1;
}

function neg_findTemplateRowWithMultipleAreaOptions_(sheet, col, excludeRow) {
  const maxCheckRow = Math.min(sheet.getLastRow(), 20);
  for (let r = 2; r <= maxCheckRow; r++) {
    if (r === excludeRow) continue;
    const rule = sheet.getRange(r, col).getDataValidation();
    if (rule && rule.getCriteriaValues()[0] && rule.getCriteriaValues()[0].length >= 2) {
      return r;
    }
  }
  return -1;
}

function neg_setDropdownKeepChipStyle_(cell, options, styleTemplateCell) {
  if (cell.getDataValidation() === null && styleTemplateCell && styleTemplateCell.getDataValidation() !== null) {
    styleTemplateCell.copyTo(cell, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  }

  const existingRule = cell.getDataValidation();
  const builder = existingRule ? existingRule.copy() : SpreadsheetApp.newDataValidation();
  const newRule = builder
    .requireValueInList(options, true)
    .setAllowInvalid(false)
    .build();
  cell.setDataValidation(newRule);
}
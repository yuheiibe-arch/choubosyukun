/**
 * ============================================================================
 * 02b_3_Merge_SheetWriter.gs（マージ管理 3分割の3）
 * ・確認済(チェックON)行の完全保護
 * ・未送信残存行の「同一プラン上書き禁止（既存行＆担当者を完全維持）」
 * ・別医師へプラン変更された場合も、既存の担当者名をP列＆送信文案へ自動引継
 * ============================================================================
 */

function extractNormalizedDocSetFromRowCell_(docCellText) {
  const set = new Set();
  String(docCellText || "")
    .split(/[\r\n、,，・\/／]+/)
    .forEach(function(part) {
      const cleaned = part.replace(/先生$/g, "").replace(/[\(（\[【].*$/g, "").trim();
      const norm = toStandardKanji_(cleaned);
      if (norm && norm !== "該当候補なし" && norm !== "対象未定" && norm !== "{医師名}") {
        set.add(norm);
      }
    });
  return set;
}

function parseRowSubSlotsWithDoctors_(row) {
  const dList = String(row[3] || "").split(/\r?\n/).map(normalizeDateStr_).filter(Boolean);
  const cList = String(row[4] || "").split(/\r?\n/).map(function(s) { return s.trim(); }).filter(Boolean);
  const tList = String(row[5] || "").split(/\r?\n/).map(normalizeTimeRange_).filter(Boolean);
  const cnt = Math.max(dList.length, cList.length, 1);
  const docSet = extractNormalizedDocSetFromRowCell_(row[8]);
  const staff = extractCleanStaffName_(row[15]);
  const slots = [];

  for (let k = 0; k < cnt; k++) {
    const d = dList[k] || dList[0] || "";
    const c = cList[k] || cList[0] || "";
    const t = tList[k] || tList[0] || "";
    if (!d || !c) continue;
    slots.push({
      slotIndex: k,
      slotCount: cnt,
      dateStr: d,
      d6ClinicKey: c,
      absenceTime: t,
      dept: row[6],
      idObj: makeSlotIdentityForMerge_(d, c, row[6], t),
      docSet: docSet,
      staff: staff
    });
  }
  return slots;
}

function hasAnyCommonDoctorInSets_(setA, setB) {
  if (!setA || !setB) return false;
  if (setA.size === 0 && setB.size === 0) return true;
  for (const d of setA) {
    if (setB.has(d)) return true;
  }
  return false;
}

function mergeAndWriteReviewSheetSmart_(params) {
  const {
    revSheet, incomingRows, d6Data, d3Data, wageSs, holidaySet, staffList
  } = params;

  const todayYmd = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");

  // 1. 新着行の枠・対象医師一覧を事前解析
  const incomingParsedList = (incomingRows || []).map(function(r) {
    return {
      row: r,
      subSlots: parseRowSubSlotsWithDoctors_(r)
    };
  });

  const survivingItems = [];
  const protectedExistingSubSlots = [];
  const existingAssignedStaffList = [];

  let keptCheckedCount = 0;
  let keptUncheckedCount = 0;
  let rebuiltBundleCount = 0;
  let removedFilledOrReplacedCount = 0;

  // 2. シート上の既存行を読み取って仕分け
  const lastRow = revSheet.getLastRow();
  if (lastRow >= 2) {
    const existingRange = revSheet.getRange(2, 1, lastRow - 1, 24);
    const rawVals = existingRange.getValues();
    const dispVals = existingRange.getDisplayValues();
    const bgs = existingRange.getBackgrounds();
    const fontColors = existingRange.getFontColors();

    for (let i = 0; i < dispVals.length; i++) {
      const disp = dispVals[i];
      const raw = rawVals[i];
      if (!String(disp[3] || "").trim() && !String(disp[4] || "").trim()) continue;

      const safeRow = disp.map(function(v, colIdx) {
        if (colIdx === 0) return (raw[0] === true || v === "TRUE");
        if ((colIdx === 12 || colIdx === 13 || colIdx === 14) && typeof raw[colIdx] === "number") {
          return raw[colIdx];
        }
        return v;
      });

      const existingCleanStaff = extractCleanStaffName_(safeRow[15]);
      safeRow[15] = existingCleanStaff;
      if (safeRow[16]) {
        safeRow[16] = normalizeDraftMessageFormat_(safeRow[16], existingCleanStaff);
      }

      const isChecked = (safeRow[0] === true);
      const exSubSlots = parseRowSubSlotsWithDoctors_(safeRow);

      if (existingCleanStaff) {
        exSubSlots.forEach(function(sl) {
          existingAssignedStaffList.push({
            idObj: sl.idObj,
            staff: existingCleanStaff
          });
        });
      }

      // A) 確認済（チェックON）の行 ➔ 完全保護
      if (isChecked) {
        keptCheckedCount++;
        survivingItems.push({
          row: safeRow,
          isChecked: true,
          bgRow: bgs[i],
          fontRow: fontColors[i]
        });
        exSubSlots.forEach(function(sl) {
          protectedExistingSubSlots.push({ idObj: sl.idObj, docSet: sl.docSet, isChecked: true });
        });
        continue;
      }

      // B) 未チェック（未送信）の行 ➔ 同一医師の再送は上書きせず既存行＆担当者を保護
      const survivingSubSlots = [];
      for (let k = 0; k < exSubSlots.length; k++) {
        const sl = exSubSlots[k];

        const isChangedToDifferentDoctor = incomingParsedList.some(function(inc) {
          return inc.subSlots.some(function(incSl) {
            if (!isSameOrOverlappingSlotIdentity_(sl.idObj, incSl.idObj)) return false;
            if (hasAnyCommonDoctorInSets_(sl.docSet, incSl.docSet)) return false;
            return true;
          });
        });

        if (isChangedToDifferentDoctor) continue;

        const d6Check = checkSlotStillUnfilledInD6_(sl.dateStr, sl.d6ClinicKey, sl.absenceTime, safeRow[6], d6Data, todayYmd);
        if (!d6Check.unfilled) continue;

        survivingSubSlots.push({
          dateStr: sl.dateStr,
          d6ClinicKey: d6Check.d6Row ? d6Check.d6Row.clinicRaw : sl.d6ClinicKey,
          absenceTime: d6Check.d6Row ? d6Check.d6Row.absenceTime : sl.absenceTime,
          dept: safeRow[6],
          d6Row: d6Check.d6Row,
          idObj: sl.idObj,
          docSet: sl.docSet,
          isSpecial: detectSlotSpecialFromExistingRow_(safeRow, k, sl.slotCount, d6Check.d6Row)
        });
      }

      if (survivingSubSlots.length === 0) {
        removedFilledOrReplacedCount++;
        continue;
      }

      if (survivingSubSlots.length === exSubSlots.length) {
        keptUncheckedCount++;
        survivingItems.push({
          row: safeRow,
          isChecked: false,
          bgRow: null,
          fontRow: null
        });
        survivingSubSlots.forEach(function(sl) {
          protectedExistingSubSlots.push({ idObj: sl.idObj, docSet: sl.docSet, isChecked: false });
        });
      } else {
        const rebuiltRow = rebuildSurvivingSubsetRow_(safeRow, survivingSubSlots, {
          d3Data: d3Data,
          wageSs: wageSs,
          holidaySet: holidaySet
        });
        rebuiltBundleCount++;
        survivingItems.push({
          row: rebuiltRow,
          isChecked: false,
          bgRow: null,
          fontRow: null
        });
        survivingSubSlots.forEach(function(sl) {
          protectedExistingSubSlots.push({ idObj: sl.idObj, docSet: sl.docSet, isChecked: false });
        });
      }
    }
  }

  // 3. 新規到着行の差し込み（既存保護枠はスキップ ＆ プラン変更枠には既存の担当者名を自動引継）
  let addedNewCount = 0;
  incomingParsedList.forEach(function(inc) {
    const newRow = inc.row;
    const incSlots = inc.subSlots;

    if (incSlots.length > 0) {
      const allAlreadyKeptOnSheet = incSlots.every(function(incSl) {
        return protectedExistingSubSlots.some(function(prot) {
          return isSameOrOverlappingSlotIdentity_(incSl.idObj, prot.idObj) &&
                 hasAnyCommonDoctorInSets_(incSl.docSet, prot.docSet);
        });
      });
      if (allAlreadyKeptOnSheet) return;

      const trulyNewSubSlots = incSlots.filter(function(incSl) {
        return !protectedExistingSubSlots.some(function(prot) {
          return isSameOrOverlappingSlotIdentity_(incSl.idObj, prot.idObj) &&
                 hasAnyCommonDoctorInSets_(incSl.docSet, prot.docSet);
        });
      });

      if (trulyNewSubSlots.length > 0 && trulyNewSubSlots.length < incSlots.length) {
        const subsetForRebuild = trulyNewSubSlots.map(function(sl) {
          const d6Check = checkSlotStillUnfilledInD6_(sl.dateStr, sl.d6ClinicKey, sl.absenceTime, newRow[6], d6Data, todayYmd);
          return {
            dateStr: sl.dateStr,
            d6ClinicKey: d6Check.d6Row ? d6Check.d6Row.clinicRaw : sl.d6ClinicKey,
            absenceTime: d6Check.d6Row ? d6Check.d6Row.absenceTime : sl.absenceTime,
            dept: newRow[6],
            d6Row: d6Check.d6Row,
            isSpecial: detectSlotSpecialFromExistingRow_(newRow, sl.slotIndex, sl.slotCount, d6Check.d6Row)
          };
        });
        const rebuiltIncoming = rebuildSurvivingSubsetRow_(newRow, subsetForRebuild, {
          d3Data: d3Data,
          wageSs: wageSs,
          holidaySet: holidaySet
        });
        for (let c = 0; c < 24; c++) newRow[c] = rebuiltIncoming[c];
      }
    }

    let finalStaff = extractCleanStaffName_(newRow[15]);
    if (!finalStaff && existingAssignedStaffList.length > 0) {
      for (let sIdx = 0; sIdx < incSlots.length; sIdx++) {
        const hitStaffObj = existingAssignedStaffList.find(function(exSt) {
          return isSameOrOverlappingSlotIdentity_(incSlots[sIdx].idObj, exSt.idObj);
        });
        if (hitStaffObj && hitStaffObj.staff) {
          finalStaff = hitStaffObj.staff;
          break;
        }
      }
    }

    newRow[15] = finalStaff;
    if (newRow[16]) {
      newRow[16] = normalizeDraftMessageFormat_(newRow[16], finalStaff);
    }

    addedNewCount++;
    survivingItems.push({
      row: newRow,
      isChecked: false,
      bgRow: null,
      fontRow: null
    });
  });

  // 4. 日付順 ➔ 開始時間順 ➔ 拠点名順でソート
  survivingItems.sort(function(a, b) {
    const dA = normalizeDateStr_(String(a.row[3] || "").split(/\r?\n/)[0]);
    const dB = normalizeDateStr_(String(b.row[3] || "").split(/\r?\n/)[0]);
    if (dA !== dB) return dA.localeCompare(dB);

    const tA = toMinutes_(String(a.row[5] || "").split(/\r?\n/)[0].split("-")[0]);
    const tB = toMinutes_(String(b.row[5] || "").split(/\r?\n/)[0].split("-")[0]);
    const minA = isNaN(tA) ? 9999 : tA;
    const minB = isNaN(tB) ? 9999 : tB;
    if (minA !== minB) return minA - minB;

    return String(a.row[4] || "").localeCompare(String(b.row[4] || ""), "ja");
  });

  const finalRows = survivingItems.map(function(it) { return it.row; });

  // 5. シートへ書き込み ＆ チェックボックス・入力規則・確認済グレー背景の復元
  writeDataExactRows_(revSheet, finalRows, 24);

  if (finalRows.length > 0) {
    const outBgs = survivingItems.map(function(it) {
      if (it.isChecked) {
        const hasGray = it.bgRow && it.bgRow[0] && it.bgRow[0].toLowerCase() !== "#ffffff";
        return hasGray ? it.bgRow : new Array(24).fill("#f3f4f6");
      }
      return new Array(24).fill("#ffffff");
    });
    const outFonts = survivingItems.map(function(it) {
      if (it.isChecked) {
        const hasCustomFont = it.fontRow && it.fontRow[0] && it.fontRow[0].toLowerCase() !== "#000000";
        return hasCustomFont ? it.fontRow : new Array(24).fill("#6b7280");
      }
      return new Array(24).fill("#000000");
    });

    revSheet.getRange(2, 1, finalRows.length, 24).setBackgrounds(outBgs).setFontColors(outFonts);

    const checkRange = revSheet.getRange(2, 1, finalRows.length, 1);
    checkRange.insertCheckboxes();
    checkRange.setValues(survivingItems.map(function(it) { return [Boolean(it.isChecked)]; }));

    const specialRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["特別時給", "通常時給"], true)
      .setAllowInvalid(false)
      .build();
    revSheet.getRange(2, 12, finalRows.length, 1).setDataValidation(specialRule);

    if (staffList && staffList.length > 0) {
      const staffRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(staffList, true)
        .setAllowInvalid(true)
        .build();
      revSheet.getRange(2, 16, finalRows.length, 1).setDataValidation(staffRule);
    }

    const postTypeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["通常", "引用して更新"], true)
      .setAllowInvalid(false)
      .build();
    revSheet.getRange(2, 20, finalRows.length, 1).setDataValidation(postTypeRule);
  }

  applyLeftMiddleToSheet_(revSheet);

  return {
    totalCount: finalRows.length,
    addedNewCount: addedNewCount,
    keptCheckedCount: keptCheckedCount,
    keptUncheckedCount: keptUncheckedCount,
    rebuiltBundleCount: rebuiltBundleCount,
    removedFilledOrReplacedCount: removedFilledOrReplacedCount
  };
}
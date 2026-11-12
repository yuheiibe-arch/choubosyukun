/**
 * ============================================================================
 * 02b_3_Merge_SheetWriter.gs（マージ管理 3分割の3）
 * ・既存シートの残存行（確認済＝保護／未チェック＝D6充足整理＆再構成）と新規行の統合・描画
 * ============================================================================
 */

function mergeAndWriteReviewSheetSmart_(params) {
  const {
    revSheet, incomingRows, d6Data, d3Data, wageSs, holidaySet, staffList
  } = params;

  const todayYmd = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");

  // 1. 新規到着行が対象としている枠リスト（未チェック既存行の上書き判定用）
  const incomingSlotIdentities = [];
  (incomingRows || []).forEach(function(r) {
    const dList = String(r[3] || "").split(/\r?\n/).map(normalizeDateStr_).filter(Boolean);
    const cList = String(r[4] || "").split(/\r?\n/).map(function(s) { return s.trim(); }).filter(Boolean);
    const tList = String(r[5] || "").split(/\r?\n/).map(normalizeTimeRange_).filter(Boolean);
    const cnt = Math.max(dList.length, cList.length, 1);
    for (let k = 0; k < cnt; k++) {
      const d = dList[k] || dList[0] || "";
      const c = cList[k] || cList[0] || "";
      const t = tList[k] || tList[0] || "";
      if (d && c) {
        incomingSlotIdentities.push(makeSlotIdentityForMerge_(d, c, r[6], t));
      }
    }
  });

  const survivingItems = [];
  const checkedSlotDocSignatures = new Set();
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

      const isChecked = (safeRow[0] === true);
      const dList = String(safeRow[3] || "").split(/\r?\n/).map(normalizeDateStr_).filter(Boolean);
      const cList = String(safeRow[4] || "").split(/\r?\n/).map(function(s) { return s.trim(); }).filter(Boolean);
      const tList = String(safeRow[5] || "").split(/\r?\n/).map(normalizeTimeRange_).filter(Boolean);
      const slotCount = Math.max(dList.length, cList.length, 1);

      // A) 確認済（チェックON）の行 ➔ Slack転記待ちのため一切上書き・削除せず完全保護
      if (isChecked) {
        keptCheckedCount++;
        survivingItems.push({
          row: safeRow,
          isChecked: true,
          bgRow: bgs[i],
          fontRow: fontColors[i]
        });

        const docNorm = toStandardKanji_(String(safeRow[8] || "").split(/\r?\n/)[0]);
        for (let k = 0; k < slotCount; k++) {
          const idObj = makeSlotIdentityForMerge_(dList[k] || dList[0], cList[k] || cList[0], safeRow[6], tList[k] || tList[0]);
          if (docNorm) {
            checkedSlotDocSignatures.add([idObj.dateStr, idObj.normClinic, idObj.timeRange, docNorm].join("|"));
          }
        }
        continue;
      }

      // B) 未チェック（未送信）の行 ➔ 各枠が「未充足かつ今回上書き変更されていないか」を判定
      const survivingSubSlots = [];
      for (let k = 0; k < slotCount; k++) {
        const sDate = dList[k] || dList[0] || "";
        const sClinicKey = cList[k] || cList[0] || "";
        const sTime = tList[k] || tList[0] || "";
        const idObj = makeSlotIdentityForMerge_(sDate, sClinicKey, safeRow[6], sTime);

        const isSuperseded = incomingSlotIdentities.some(function(incId) {
          return isSameOrOverlappingSlotIdentity_(idObj, incId);
        });
        if (isSuperseded) continue;

        const d6Check = checkSlotStillUnfilledInD6_(sDate, sClinicKey, sTime, safeRow[6], d6Data, todayYmd);
        if (!d6Check.unfilled) continue;

        survivingSubSlots.push({
          dateStr: sDate,
          d6ClinicKey: d6Check.d6Row ? d6Check.d6Row.clinicRaw : sClinicKey,
          absenceTime: d6Check.d6Row ? d6Check.d6Row.absenceTime : sTime,
          dept: safeRow[6],
          d6Row: d6Check.d6Row,
          isSpecial: detectSlotSpecialFromExistingRow_(safeRow, k, slotCount, d6Check.d6Row)
        });
      }

      if (survivingSubSlots.length === 0) {
        removedFilledOrReplacedCount++;
        continue;
      }

      if (survivingSubSlots.length === slotCount) {
        keptUncheckedCount++;
        survivingItems.push({
          row: safeRow,
          isChecked: false,
          bgRow: null,
          fontRow: null
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
      }
    }
  }

  // 3. 新規到着行を追加
  let addedNewCount = 0;
  (incomingRows || []).forEach(function(newRow) {
    const firstDate = normalizeDateStr_(String(newRow[3] || "").split(/\r?\n/)[0]);
    const firstClinic = String(newRow[4] || "").split(/\r?\n/)[0].trim();
    const firstTime = normalizeTimeRange_(String(newRow[5] || "").split(/\r?\n/)[0]);
    const docNorm = toStandardKanji_(String(newRow[8] || "").split(/\r?\n/)[0]);
    const idObj = makeSlotIdentityForMerge_(firstDate, firstClinic, newRow[6], firstTime);
    const sig = [idObj.dateStr, idObj.normClinic, idObj.timeRange, docNorm].join("|");

    if (docNorm && !String(newRow[3] || "").includes("\n") && checkedSlotDocSignatures.has(sig)) {
      return;
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
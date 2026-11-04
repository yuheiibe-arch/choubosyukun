/**
 * ============================================================================
 * 03b_3_Individual_DocResolver.gs
 * 個別アプローチ指名医師の厳格特定（当日D3最優先・医師マスタ6時間キャッシュ・ハルシネーション遮断）
 * ============================================================================
 */

var INDIVIDUAL_DOC_MASTER_MEM_CACHE_ = null;
var INDIVIDUAL_DOC_MASTER_CACHE_KEY_ = "CHOUBOSYU_DOC_MASTER_V1";

/**
 * 個別アプローチ枠の医師特定エントリーポイント
 */
function resolveIndividualApproachSlot_(cleanTarget, cleanLic, dateStr, clinicName, targetDept, d6Row, d3Data, effectiveAbsenceTime, absStart, absEnd) {
  const targetClinicDisp = d6Row ? d6Row.clinicRaw : (targetDept === "ワクチン専任" ? `${clinicName}(ワクチン専任)` : clinicName);

  // 1. 指名なし（後続の 03c_Individual_Candidate.gs で勤務くん実績から候補を自動抽出するケース）
  if (!cleanTarget && !cleanLic) {
    return {
      status: "OK",
      resolvedApproachType: "個別アプローチ",
      resolvedDept: targetDept,
      doctorName: "",
      currentShiftLabel: "-",
      consultStart: absStart || "",
      consultEnd: absEnd || "",
      consultShiftLabel: `個別アプローチ(${effectiveAbsenceTime})`
    };
  }

  // 2. 指名あり：当日D3 ➔ 医師マスタ(高速キャッシュ) の順で厳格に1名特定
  const strictRes = resolveIndividualDoctorStrictFast_(
    cleanTarget, cleanLic, dateStr, clinicName, targetDept, effectiveAbsenceTime, d3Data
  );

  if (strictRes.status === "OK") {
    return {
      status: "OK",
      resolvedApproachType: "個別アプローチ",
      resolvedDept: targetDept,
      doctorName: strictRes.doctorName,
      fromClinic: "",
      isCurrentVaccine: false,
      currentClinicLabel: "-",
      currentStart: "",
      currentEnd: "",
      consultStart: absStart || "",
      consultEnd: absEnd || "",
      hasTimeExpansion: false,
      consultNoteSuffix: "",
      currentShiftLabel: strictRes.currentShiftLabel || "-",
      consultShiftLabel: `${targetClinicDisp} ${effectiveAbsenceTime}`,
      isFullTime: Boolean(strictRes.isFullTime),
      alreadyHasAllowance: false,
      evidenceNote: strictRes.evidenceNote
    };
  }

  // 3. 1名に特定できなかった場合（該当なし・音声誤変換・同姓複数など）は確実に『確認リスト』へ送る
  return {
    status: "要確認",
    resolvedApproachType: "個別アプローチ",
    resolvedDept: targetDept,
    reason: strictRes.reason,
    doctorName: cleanTarget,
    fromClinic: "",
    currentShiftLabel: "要確認",
    consultShiftLabel: `${targetClinicDisp} ${effectiveAbsenceTime}`
  };
}

/**
 * 【高速＆厳格】個別アプローチ指名医師の段階的特定ロジック
 */
function resolveIndividualDoctorStrictFast_(cleanTarget, cleanLic, dateStr, clinicName, targetDept, effectiveAbsenceTime, d3Data) {
  const normTarget = toStandardKanji_(cleanTarget);
  const deptKey = (targetDept === "内科") ? "内科" : "小児科";

  // Step 1: まず「その日（dateStr）にD3シフトにいる医師」を第一候補としてチェック（0.001秒）
  const sameDayD3 = d3Data.filter(s => s.dateStr === dateStr);
  let dayHits = [];
  if (cleanLic) {
    dayHits = sameDayD3.filter(s => s.docId === cleanLic);
  }
  if (dayHits.length === 0 && normTarget.length >= 2) {
    dayHits = sameDayD3.filter(s => s.normName === normTarget || s.normName.startsWith(normTarget));
    if (dayHits.length === 0) {
      dayHits = matchSameFamilyPhonetic_(normTarget, sameDayD3.map(s => ({
        fullName: s.docName,
        normFull: s.normName,
        normLast: s.normLast || s.normName.slice(0, 2),
        firstKana: extractKatakanaPart_(s.firstKana),
        lastKana: extractKatakanaPart_(s.lastKana),
        kanaName: extractKatakanaPart_(s.kanaName),
        rawShift: s
      }))).map(x => x.rawShift);
    }
  }

  const uniqueDayNames = Array.from(new Set(dayHits.map(s => s.docName)));
  if (uniqueDayNames.length === 1) {
    const s = dayHits[0];
    const shiftTime = `${s.startTime}-${s.endTime}`;
    if (s.clinicName !== clinicName && (!effectiveAbsenceTime || isTimeOverlapping_(shiftTime, effectiveAbsenceTime))) {
      return {
        status: "要確認",
        reason: `当日【${s.clinicName}】(${shiftTime})で勤務中のため個別アプローチ不可・移動か要確認 (${s.docName}先生)`
      };
    }
    return {
      status: "OK",
      doctorName: s.docName,
      isFullTime: s.isFullTime,
      currentShiftLabel: s.clinicName === clinicName ? `${s.clinicName} ${shiftTime}` : "-",
      evidenceNote: `個別アプローチ指名医師(当日D3シフト一致:${s.docName})`
    };
  } else if (uniqueDayNames.length > 1) {
    return {
      status: "要確認",
      reason: `当日シフトに候補が複数名あり (${uniqueDayNames.join(" / ")})`
    };
  }

  // Step 2: 統合医師インデックス（医師マスタ＋D3全日程、キャッシュ付き）で厳格特定
  const masterPool = getUnifiedDoctorIndexFast_(d3Data);
  const deptPool = masterPool.filter(d => !d.dept || d.dept.includes(deptKey));

  // Tier 1: 医籍番号完全一致
  if (cleanLic) {
    const licHit = masterPool.find(d => d.licenseNo === cleanLic);
    if (licHit) {
      return {
        status: "OK",
        doctorName: licHit.fullName,
        isFullTime: licHit.isFullTime,
        evidenceNote: `個別アプローチ指名医師(医籍番号一致:${cleanLic})`
      };
    }
  }

  // Tier 2: フルネーム完全一致
  const exactFull = deptPool.filter(d => d.normFull === normTarget);
  if (exactFull.length === 1) {
    return {
      status: "OK",
      doctorName: exactFull[0].fullName,
      isFullTime: exactFull[0].isFullTime,
      evidenceNote: `個別アプローチ指名医師(${exactFull[0].source}完全一致)`
    };
  }

  // Tier 3: 前方一致（「竹内正」→「竹内正宣」など）※2文字以上かつ1名に絞れた時のみOK
  if (normTarget.length >= 2) {
    const prefixMatches = deptPool.filter(d => d.normFull.startsWith(normTarget));
    if (prefixMatches.length === 1) {
      return {
        status: "OK",
        doctorName: prefixMatches[0].fullName,
        isFullTime: prefixMatches[0].isFullTime,
        evidenceNote: `個別アプローチ指名医師(前方一致特定:「${cleanTarget}」→「${prefixMatches[0].fullName}」)`
      };
    }
    if (prefixMatches.length > 1) {
      return {
        status: "要確認",
        reason: `同姓・前方一致候補が複数名あり (${prefixMatches.slice(0, 6).map(d => d.fullName).join(" / ")})`
      };
    }
  }

  // Tier 4: 姓が一致し、下の名前の1文字目の読みカナが一致（「宮本義」→「宮本吉輝」など）※1名に絞れた時のみOK
  const phoneticMatches = matchSameFamilyPhonetic_(normTarget, deptPool);
  if (phoneticMatches.length === 1) {
    return {
      status: "OK",
      doctorName: phoneticMatches[0].fullName,
      isFullTime: phoneticMatches[0].isFullTime,
      evidenceNote: `個別アプローチ指名医師(姓名カナ特定:「${cleanTarget}」→「${phoneticMatches[0].fullName}」)`
    };
  }
  if (phoneticMatches.length > 1) {
    return {
      status: "要確認",
      reason: `同姓・同音候補が複数名あり (${phoneticMatches.slice(0, 6).map(d => d.fullName).join(" / ")})`
    };
  }

  // Tier 5: 該当なし・曖昧な場合は類似候補を添えてすべて『確認リスト』へ送る
  const scoredCandidates = deptPool
    .map(d => ({
      doc: d,
      score: calcDoctorSimilarityScore_(cleanTarget, {
        normName: d.normFull,
        normLast: d.normLast,
        normFirst: d.normFirst,
        kanaName: d.kanaName,
        lastKana: d.lastKana
      })
    }))
    .filter(x => x.score >= 70)
    .sort((a, b) => b.score - a.score);

  if (scoredCandidates.length > 0) {
    const candStr = scoredCandidates.slice(0, 5).map(x => x.doc.fullName).join(" / ");
    return {
      status: "要確認",
      reason: `「${cleanTarget}」は完全一致なし（候補: ${candStr}）`
    };
  }

  return {
    status: "要確認",
    reason: `医師マスタ・シフトに「${cleanTarget}」先生が見つかりません`
  };
}

/**
 * 姓が一致し、下の名前の1文字目の読み（同音異字：義⇔吉＝ヨシ等）が一致する医師を抽出する
 */
function matchSameFamilyPhonetic_(normTarget, pool) {
  if (!normTarget || normTarget.length < 2) return [];
  const givenNamePhoneticMap = {
    "義": ["ヨシ", "ギ"], "吉": ["ヨシ", "キチ"], "良": ["ヨシ", "リョウ"], "芳": ["ヨシ"], "由": ["ヨシ", "ユ"],
    "正": ["マサ", "ショウ", "タダ"], "政": ["マサ"], "雅": ["マサ"], "昌": ["マサ"], "真": ["マサ", "シン", "マ"],
    "智": ["トモ", "サト"], "知": ["トモ", "チ"], "友": ["トモ"],
    "裕": ["ユウ", "ヒロ"], "祐": ["ユウ"], "雄": ["ユウ"],
    "博": ["ヒロ"], "広": ["ヒロ"], "宏": ["ヒロ"], "浩": ["ヒロ"],
    "隆": ["タカ"], "貴": ["タカ"], "孝": ["タカ"],
    "健": ["ケン", "タケ"], "賢": ["ケン"], "憲": ["ケン"]
  };

  const sameFamilyDocs = pool.filter(d =>
    d.normLast && normTarget.startsWith(d.normLast) && normTarget.length > d.normLast.length
  );
  if (sameFamilyDocs.length === 0) return [];

  const restChar = normTarget.slice(sameFamilyDocs[0].normLast.length, sameFamilyDocs[0].normLast.length + 1);
  const possibleReadings = givenNamePhoneticMap[restChar] || [];
  if (possibleReadings.length === 0) return [];

  return sameFamilyDocs.filter(d =>
    possibleReadings.some(r =>
      (d.firstKana && d.firstKana.startsWith(r)) ||
      (d.lastKana && d.kanaName && d.kanaName.slice(d.lastKana.length).startsWith(r))
    )
  );
}

/**
 * 医師マスタ＋D3シフトの統合インデックスを爆速取得（ScriptCache 6時間 ＋ メモリキャッシュ）
 */
function getUnifiedDoctorIndexFast_(d3Data) {
  if (INDIVIDUAL_DOC_MASTER_MEM_CACHE_) return INDIVIDUAL_DOC_MASTER_MEM_CACHE_;

  const cache = CacheService.getScriptCache();
  let masterList = null;

  try {
    const cachedJson = cache.get(INDIVIDUAL_DOC_MASTER_CACHE_KEY_);
    if (cachedJson) {
      masterList = JSON.parse(cachedJson);
    }
  } catch (e) {}

  if (!masterList) {
    masterList = [];
    try {
      const masterSs = SpreadsheetApp.openById(NEGOTIATION_CONFIG.DOCTOR_MASTER_SS_ID);
      const masterSh = masterSs.getSheetByName(NEGOTIATION_CONFIG.DOCTOR_MASTER_SHEET_NAME) || masterSs.getSheets()[0];
      const mData = masterSh.getDataRange().getDisplayValues();

      if (mData.length > 1) {
        const h = mData[0].map(s => String(s || "").trim());
        const iFull = h.indexOf("姓名");
        const iLast = h.indexOf("姓");
        const iFirst = h.indexOf("名");
        const iKana = h.indexOf("セイメイ");
        const iLastKana = h.indexOf("セイ");
        const iFirstKana = h.indexOf("メイ");
        const iLic = h.indexOf("医籍登録番号");
        const iDept = h.indexOf("診療科目");

        for (let r = 1; r < mData.length; r++) {
          const row = mData[r];
          const rawLast = iLast !== -1 ? String(row[iLast] || "").trim() : "";
          const rawFirst = iFirst !== -1 ? String(row[iFirst] || "").trim() : "";
          let full = iFull !== -1 ? String(row[iFull] || "").replace(/[\s ]+/g, "").trim() : "";
          if (!full && (rawLast || rawFirst)) full = (rawLast + rawFirst).replace(/[\s ]+/g, "");
          if (!full) continue;

          const deptCleaned = typeof neg_cleanDoctorDepartments_ === "function"
            ? neg_cleanDoctorDepartments_(row[iDept])
            : String(row[iDept] || "");
          if (!deptCleaned) continue;

          let last = rawLast;
          let first = rawFirst;
          if (!last && iFull !== -1 && String(row[iFull]).match(/[\s ]/)) {
            const parts = String(row[iFull]).trim().split(/[\s ]+/);
            last = parts[0] || "";
            first = parts.slice(1).join("");
          }

          const rawKanaFull = iKana !== -1 ? String(row[iKana] || "").trim() : "";
          let lastKana = iLastKana !== -1 ? String(row[iLastKana] || "").trim() : "";
          let firstKana = iFirstKana !== -1 ? String(row[iFirstKana] || "").trim() : "";
          if (!lastKana && rawKanaFull.match(/[\s ]/)) {
            const kParts = rawKanaFull.split(/[\s ]+/);
            lastKana = kParts[0] || "";
            firstKana = kParts.slice(1).join("");
          }

          masterList.push({
            fullName: full,
            normFull: toStandardKanji_(full),
            normLast: toStandardKanji_(last),
            normFirst: toStandardKanji_(first),
            kanaName: extractKatakanaPart_(rawKanaFull),
            lastKana: extractKatakanaPart_(lastKana),
            firstKana: extractKatakanaPart_(firstKana),
            licenseNo: iLic !== -1 ? String(row[iLic] || "").trim() : "",
            dept: deptCleaned,
            isFullTime: false,
            source: "医師マスタ"
          });
        }
      }

      const jsonStr = JSON.stringify(masterList);
      if (jsonStr.length < 95000) {
        cache.put(INDIVIDUAL_DOC_MASTER_CACHE_KEY_, jsonStr, 21600);
      }
    } catch (e) {
      console.log("医師マスタ読込スキップ: " + e.message);
    }
  }

  const map = new Map();
  for (const m of masterList) {
    map.set(m.normFull, { ...m });
  }
  if (Array.isArray(d3Data)) {
    for (const s of d3Data) {
      if (!s.normName) continue;
      if (map.has(s.normName)) {
        const ex = map.get(s.normName);
        if (s.isFullTime) ex.isFullTime = true;
        if (!ex.normLast && s.normLast) ex.normLast = s.normLast;
        if (!ex.normFirst && s.normFirst) ex.normFirst = s.normFirst;
        if (!ex.lastKana && s.lastKana) ex.lastKana = extractKatakanaPart_(s.lastKana);
        if (!ex.firstKana && s.firstKana) ex.firstKana = extractKatakanaPart_(s.firstKana);
      } else {
        map.set(s.normName, {
          fullName: s.docName,
          normFull: s.normName,
          normLast: s.normLast || "",
          normFirst: s.normFirst || "",
          kanaName: extractKatakanaPart_(s.kanaName),
          lastKana: extractKatakanaPart_(s.lastKana),
          firstKana: extractKatakanaPart_(s.firstKana),
          licenseNo: s.docId || "",
          dept: s.dept || "小児科",
          isFullTime: Boolean(s.isFullTime),
          source: "D3シフト"
        });
      }
    }
  }

  INDIVIDUAL_DOC_MASTER_MEM_CACHE_ = Array.from(map.values());
  return INDIVIDUAL_DOC_MASTER_MEM_CACHE_;
}
/**
 * ============================================================================
 * 05a_Slack_StructParser.gs（Slack＆D6転記 3分割の1）
 * ・日付ヘッダー生成・正規化
 * ・Slack1日分メッセージの構造化分解・拠点別マージ・シリアライズ（引用 `>` や `⇒` 結果行も保護）
 * ============================================================================
 */

function getSlackThreadKeyAndHeader_(dateStr, fallbackTitle) {
  const norm = normalizeDateStr_(dateStr);
  if (!norm) {
    return { halfKey: "", headerText: fallbackTitle || "```超募集MTGログ```" };
  }
  const d = parseDate_(norm);
  const y = d.getFullYear();
  const m = d.getMonth() + 1;
  const day = d.getDate();
  const lastDay = new Date(y, m, 0).getDate();
  const isFirstHalf = day <= 15;
  const halfLabel = isFirstHalf ? "前半" : "後半";
  const rangeStr = isFirstHalf ? `${m}/1~${m}/15` : `${m}/16~${m}/${lastDay}`;
  return {
    halfKey: `${m}月${halfLabel}`,
    headerText: `\`\`\`【${m}月${halfLabel}】${rangeStr} 超募集MTGログ\`\`\``
  };
}

function formatSlackDateHeader_(dateStr) {
  const norm = normalizeDateStr_(dateStr);
  if (!norm) return "";
  const d = parseDate_(norm);
  if (!d || isNaN(d.getTime())) return norm;
  const days = ["日", "月", "火", "水", "木", "金", "土"];
  return `${d.getMonth() + 1}月${d.getDate()}日 (${days[d.getDay()]})`;
}

function getSlackMdKeyFromDateStr_(dateStr) {
  const norm = normalizeDateStr_(dateStr);
  if (!norm) return "";
  const d = parseDate_(norm);
  if (!d || isNaN(d.getTime())) return "";
  return (d.getMonth() + 1) + "/" + d.getDate();
}

function normalizeSlackBodyDateHeader_(slackText, targetDate) {
  const header = formatSlackDateHeader_(targetDate);
  if (!header) return slackText;
  const s = String(slackText || "").trim();
  if (!s) return "";
  if (/^`[^`\r\n]+`/.test(s)) {
    return s.replace(/^`[^`\r\n]+`/, "`" + header + "`");
  }
  return "`" + header + "`\n" + s;
}

function buildSingleSlotSlackBlockText_(targetDate, targetClinic, targetTime, isSpecial, memoToAppend) {
  const header = formatSlackDateHeader_(targetDate);
  const cleanMemo = String(memoToAppend || "").startsWith("・") ? memoToAppend : `・${memoToAppend}`;
  return `\`${header}\`\n\`\`\`\n【${targetClinic}】${targetTime}${isSpecial ? "【特別時給】" : ""}\n${cleanMemo}\n\`\`\``;
}

/**
 * Slackの1日分メッセージ本文（引用 `>` 付きや `⇒` 結果行付きも含む）を構造化データへ分解する
 */
function parseSlackDateMessageStruct_(rawSlackText) {
  const sections = new Map();
  const resultLines = [];
  if (!rawSlackText) return { sections: sections, resultLines: resultLines };

  const lines = String(rawSlackText).replace(/&gt;/g, ">").split(/\r?\n/);
  let currentSec = null;

  for (let i = 0; i < lines.length; i++) {
    const rawLine = lines[i];
    const trimmedRaw = rawLine.trim();
    if (!trimmedRaw) continue;

    // 上流（AI会議_3_3b）が付与した調整結果行（⇒ / ⇨ / →）は保護して末尾へ維持
    if (/^(?:⇒|⇨|→)/.test(trimmedRaw)) {
      if (!resultLines.includes(trimmedRaw)) resultLines.push(trimmedRaw);
      continue;
    }

    const unquoted = rawLine.replace(/^>\s?/, "").trim();
    if (!unquoted || unquoted === "```") continue;
    if (/^`\d{1,2}月\s*\d{1,2}日/.test(unquoted)) continue;
    if (/^(?:⇒|⇨|→)/.test(unquoted)) {
      if (!resultLines.includes(unquoted)) resultLines.push(unquoted);
      continue;
    }

    const headerMatch = unquoted.match(/^【([^】]+)】\s*([0-9]{1,2}:[0-9]{2}\s*-\s*[0-9]{1,2}:[0-9]{2})?(.*)$/);
    if (headerMatch) {
      const clinicDisp = headerMatch[1].trim();
      const timeRange = headerMatch[2] ? normalizeTimeRange_(headerMatch[2].replace(/\s+/g, "")) : "";
      const isSpecial = String(headerMatch[3] || "").includes("特別時給");
      const deptKey = clinicDisp.includes("内科") ? "内科" : "小児科";
      const secKey = [normalizeClinicForMatch_(clinicDisp), deptKey, timeRange].join("|");

      if (!sections.has(secKey)) {
        sections.set(secKey, {
          clinicDisp: clinicDisp,
          timeRange: timeRange,
          isSpecial: isSpecial,
          bullets: []
        });
      } else if (isSpecial) {
        sections.get(secKey).isSpecial = true;
      }
      currentSec = sections.get(secKey);
      continue;
    }

    if (currentSec) {
      const cleanBulletBody = unquoted.replace(/^・\s*/, "").trim();
      if (cleanBulletBody) {
        const bulletLine = "・" + cleanBulletBody;
        const exists = currentSec.bullets.some(function(b) {
          return b.replace(/^・\s*/, "").trim() === cleanBulletBody;
        });
        if (!exists) currentSec.bullets.push(bulletLine);
      }
    }
  }

  return { sections: sections, resultLines: resultLines };
}

function mergeSlackDateMessageStructs_(baseStruct, addStruct) {
  const mergedSections = new Map();

  const appendFrom = function(srcMap) {
    if (!srcMap) return;
    srcMap.forEach(function(sec, key) {
      if (!mergedSections.has(key)) {
        mergedSections.set(key, {
          clinicDisp: sec.clinicDisp,
          timeRange: sec.timeRange,
          isSpecial: Boolean(sec.isSpecial),
          bullets: sec.bullets.slice()
        });
      } else {
        const ex = mergedSections.get(key);
        if (sec.isSpecial) ex.isSpecial = true;
        sec.bullets.forEach(function(b) {
          const cleanB = b.replace(/^・\s*/, "").trim();
          const dup = ex.bullets.some(function(eb) {
            return eb.replace(/^・\s*/, "").trim() === cleanB;
          });
          if (!dup) ex.bullets.push(b);
        });
      }
    });
  };

  appendFrom(baseStruct ? baseStruct.sections : null);
  appendFrom(addStruct ? addStruct.sections : null);

  const mergedResults = [];
  [(baseStruct && baseStruct.resultLines) || [], (addStruct && addStruct.resultLines) || []].forEach(function(arr) {
    arr.forEach(function(r) {
      if (r && !mergedResults.includes(r)) mergedResults.push(r);
    });
  });

  return { sections: mergedSections, resultLines: mergedResults };
}

function serializeSlackDateMessageStruct_(dateStr, structObj) {
  const header = formatSlackDateHeader_(dateStr);
  const blockTexts = [];

  (structObj.sections || new Map()).forEach(function(sec) {
    const spTag = sec.isSpecial ? "【特別時給】" : "";
    const headLine = "【" + sec.clinicDisp + "】" + (sec.timeRange || "") + spTag;
    const lines = [headLine].concat(sec.bullets || []);
    blockTexts.push(lines.join("\n"));
  });

  const coreLines = [
    "`" + header + "`",
    "```",
    blockTexts.join("\n\n"),
    "```"
  ];
  const coreText = coreLines.join("\n");

  if (structObj.resultLines && structObj.resultLines.length > 0) {
    const quotedCore = coreText.split("\n").map(function(l) { return "> " + l; }).join("\n");
    return quotedCore + "\n" + structObj.resultLines.join("\n");
  }
  return coreText;
}
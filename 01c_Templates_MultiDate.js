/**
 * ============================================================================
 * 01c_Templates_MultiDate.gs（01系 3分割の3）
 * 同一医師に対する複数日・応募シフト移動の統合文案テンプレート（buildMultiDateTemplate_）
 * ============================================================================
 */

function buildMultiDateTemplate_(entries) {
  if (!entries || entries.length === 0) return "";
  if (entries.length === 1) return buildMessageTemplate_(entries[0]);

  const first = entries[0];
  const docHonorific = `${first.doctorName}先生`;
  const resolvedStaffEntry = entries.find(function(e) { return Boolean(extractCleanStaffName_(e.staffName)); });
  const staffLabel = resolvedStaffEntry ? extractCleanStaffName_(resolvedStaffEntry.staffName) : "{担当者名}";

  const dateGroupMap = new Map();
  for (const e of entries) {
    if (!dateGroupMap.has(e.dateStr)) {
      dateGroupMap.set(e.dateStr, []);
    }
    dateGroupMap.get(e.dateStr).push(e);
  }
  const dateGroups = Array.from(dateGroupMap.values());
  const uniqueDateLabels = dateGroups.map(function(g) {
    return (g[0].isTomorrow ? "明日" : "") + formatMdDay_(g[0].dateObj);
  }).join("、");

  if (first.isFullTime) {
    const blocks = dateGroups.map(function(group) {
      const gFirst = group[0];
      const ymd = formatYmdDay_(gFirst.dateObj);
      const fromC = gFirst.approachType === "勤務地変更相談" ? gFirst.fromClinicName : (gFirst.currentClinicLabel || gFirst.clinicName);
      const consultLines = group.map(function(e, idx) {
        const prefix = group.length > 1 ? `ご相談${getCircledNumber_(idx + 1)})` : "ご相談)";
        return `${prefix} 【${e.clinicName}】${e.consultStart}~${e.consultEnd}`;
      });
      return [
        `・${ymd}`,
        `現在)【${fromC}】${gFirst.currentStart}~${gFirst.currentEnd}`,
        ...consultLines
      ].join("\n");
    }).join("\n\n");

    return [
      docHonorific,
      "お疲れ様です。",
      `ドクターサポート部の${staffLabel}です。`,
      "",
      `この度は、${uniqueDateLabels}のご勤務につきましてご相談がありご連絡させていただきました。`,
      "医師不在となる可能性がある拠点があり、大変恐縮ではございますがもし先生のご都合がよろしければ、下記のご勤務変更・ご調整をいただくことは可能でしょうか。",
      "",
      blocks,
      "",
      "急なお願いとなり大変恐れ入りますが、ご無理のない範囲でご検討いただけますと幸いです。",
      "お忙しいところ、ご負担をおかけして恐縮ではございますが、何卒よろしくお願い申し上げます。",
      "",
      `ドクターサポート部\u3000${staffLabel}`
    ].join("\n");
  }

  if (first.shiftType === "応募") {
    const m = first.dateObj.getMonth() + 1;
    const countPrefix = dateGroups.length === 1 ? "当日" : (dateGroups.length === 2 ? "両日とも" : "いずれも");
    const blocks = dateGroups.map(function(group) {
      const gFirst = group[0];
      const ymd = formatYmdDay_(gFirst.dateObj);
      const curW = formatComma_(gFirst.currentDailyWage);
      const consultLines = group.map(function(e, idx) {
        const prefix = group.length > 1 ? `ご相談${getCircledNumber_(idx + 1)})` : "ご相談)";
        const conW = formatComma_(e.consultDailyWage);
        return `${prefix} 【${e.clinicName}】${e.consultStart}~${e.consultEnd} 日給${conW}円`;
      });
      return [
        `■${ymd}`,
        `現在)【${gFirst.fromClinicName}】${gFirst.currentStart}~${gFirst.currentEnd} 日給${curW}円`,
        ...consultLines
      ].join("\n");
    }).join("\n\n");

    return [
      docHonorific,
      "平素よりお世話になっております。",
      `ドクターサポート事務局の${staffLabel}でございます。`,
      "",
      `このたびは${m}月に多くのご応募いただき、誠にありがとうございます。`,
      `恐れ入りますが、${uniqueDateLabels}のご応募につきまして、ご相談差し上げたくご連絡いたしました。`,
      `${countPrefix}医師不在の拠点が多く、もし先生のご都合がよろしければ、ご勤務地を変更いただくことは可能でしょうか。`,
      "その際は時給を高時給の拠点に合わせてご案内させていただきます。",
      "",
      blocks,
      "",
      "難しい場合は現在の拠点にて採用させていただきますので、お手数をおかけいたしますが、ご検討いただけましたら幸いです。",
      "何卒よろしくお願い申し上げます。",
      "",
      `ドクターサポート事務局\u3000${staffLabel}`
    ].join("\n");
  }

  const uniqueFrom = Array.from(new Set(entries.map(function(e) { return e.fromClinicName; })));
  const uniqueTo = Array.from(new Set(entries.map(function(e) { return e.clinicName; })));
  let routeSentence = "当日のご勤務拠点を下記の通りご変更いただくことは可能でしょうか。";
  if (uniqueFrom.length === 1 && uniqueTo.length === 1) {
    routeSentence = `当日のご勤務拠点を 【${uniqueFrom[0]}】 より 【${uniqueTo[0]}】 へご変更いただくことは可能でしょうか。`;
  } else if (dateGroups.length === 1 && uniqueFrom.length === 1 && uniqueTo.length > 1) {
    const toListStr = uniqueTo.map(function(c) { return `【${c}】`; }).join(" または ");
    routeSentence = `当日のご勤務拠点を 【${uniqueFrom[0]}】 より ${toListStr} へご変更いただくことは可能でしょうか。`;
  }

  const blocks = dateGroups.map(function(group) {
    const gFirst = group[0];
    const md = formatMdDay_(gFirst.dateObj);
    const curW = formatComma_(gFirst.currentDailyWage);
    const fromC = gFirst.approachType === "勤務地変更相談" ? gFirst.fromClinicName : (gFirst.currentClinicLabel || gFirst.clinicName);
    const consultLines = group.map(function(e, idx) {
      const prefix = group.length > 1 ? `ご相談${getCircledNumber_(idx + 1)})` : "ご相談)";
      const conW = formatComma_(e.consultDailyWage);
      const allowSuffix = e.hasRequestAllow ? "+依頼手当10,000円" : "";
      return `${prefix} 【${e.clinicName}】${e.consultStart}~${e.consultEnd}(日給:${conW}円${allowSuffix})`;
    });
    return [
      `■${md}`,
      `現在)【${fromC}】 ${gFirst.currentStart}~${gFirst.currentEnd} (日給:${curW}円)`,
      "↓",
      ...consultLines
    ].join("\n");
  }).join("\n\n");

  return [
    docHonorific,
    "平素よりお世話になっております。",
    `ドクターサポート事務局の${staffLabel}でございます。`,
    "",
    `ご勤務のご相談が重なり恐縮ですが、${uniqueDateLabels}のご勤務につきましてご相談でございます。`,
    routeSentence,
    "",
    blocks,
    "",
    "※上記募集は公開中のため、ご応募いただいた際に充足している可能性もございます。",
    "恐れ入りますが、充足の際はご容赦くださいませ。",
    "",
    "お忙しいところ、ご負担をおかけして恐縮ではございますが、ご検討いただけましたら幸いです。",
    "何卒よろしくお願い申し上げます。",
    "",
    `ドクターサポート事務局\u3000${staffLabel}`
  ].join("\n");
}
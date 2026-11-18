/**
 * ============================================================================
 * 01b_Templates_Single.gs（01系 3分割の2）
 * 単枠用の8章送信文案テンプレート（buildMessageTemplate_：常勤・非常勤／早出・延長・移動・個別）
 * ============================================================================
 */

function buildMessageTemplate_(params) {
  const {
    isFullTime, doctorName, staffName, approachType, shiftType,
    dateObj, isTomorrow, clinicName, currentClinicLabel, fromClinicName,
    absenceStart, absenceEnd, currentStart, currentEnd,
    consultStart, consultEnd, consultNoteSuffix, diffHoursText,
    isMorningToNightExt, consultPm1DailyWage,
    currentDailyWage, consultDailyWage, hasRequestAllow, isPublicRecruit,
    multiTargets
  } = params;

  const docHonorific = `${doctorName || "{医師名}"}先生`;
  const staffLabel = extractCleanStaffName_(staffName) || "{担当者名}";
  const mdDay = formatMdDay_(dateObj);
  const ymdDay = formatYmdDay_(dateObj);
  const tomorrowPrefix = isTomorrow ? "明日" : "";
  const curWageStr = formatComma_(currentDailyWage);
  const pm1WageStr = formatComma_(consultPm1DailyWage);
  const conWageStr = formatComma_(consultDailyWage);
  const curClinicDisp = currentClinicLabel || clinicName;
  const noteSuffix = consultNoteSuffix ? ` ${consultNoteSuffix}` : "";
  const pm1NoteSuffix = noteSuffix ? noteSuffix.replace(/~\d{2}:\d{2}/, "~18:00") : "";

  if (isFullTime) {
    if (approachType === "延長相談" && isMorningToNightExt) {
      return [
        docHonorific,
        "お疲れ様です。",
        `ドクターサポート部の${staffLabel}です。`,
        "",
        `この度は、${tomorrowPrefix}${mdDay}のご勤務につきましてご相談がありご連絡させていただきました。`,
        `現在、【${clinicName}】${absenceStart}~${absenceEnd}の時間帯に医師不在となる可能性があるため、`,
        `長時間のご勤務となり大変恐縮ではございますが、もし先生のご都合がよろしければ、可能な限りご勤務時間を延長していただくことは可能でしょうか。`,
        `18:00まで、あるいは${consultEnd}までなど、可能な範囲でご検討いただけますと幸いです。`,
        "",
        `・${ymdDay}`,
        `現在)【${curClinicDisp}】${currentStart}~${currentEnd}`,
        `ご相談①) 【${clinicName}】${currentStart}~18:00${pm1NoteSuffix}`,
        `ご相談②) 【${clinicName}】${currentStart}~${consultEnd}${noteSuffix}`,
        "",
        "長時間にわたり大きなご負担をおかけするお願いとなり大変心苦しい限りではございますが、ご無理のない範囲でご検討いただけますと幸いです。",
        "お忙しいところ誠に恐縮ではございますが、何卒よろしくお願い申し上げます。",
        "",
        `ドクターサポート部\u3000${staffLabel}`
      ].join("\n");
    }

    let consultPhrase = "ご勤務いただくことは可能でしょうか。";
    let hasCurrentRow = false;
    let curLabelForFullTime = curClinicDisp;
    let targetStartForFullTime = consultStart || absenceStart;
    let targetEndForFullTime = consultEnd || absenceEnd;

    if (approachType === "延長相談") {
      consultPhrase = `ご勤務時間を${diffHoursText}延長していただくことは可能でしょうか。`;
      hasCurrentRow = true;
    } else if (approachType === "早出相談") {
      consultPhrase = `ご勤務時間を${diffHoursText}前倒ししていただくことは可能でしょうか。`;
      hasCurrentRow = true;
    } else if (approachType === "勤務地変更相談") {
      consultPhrase = `ご勤務拠点を 【${fromClinicName}】 より 【${clinicName}】 へご変更いただくことは可能でしょうか。`;
      hasCurrentRow = true;
      curLabelForFullTime = fromClinicName;
    } else {
      targetStartForFullTime = absenceStart;
      targetEndForFullTime = absenceEnd;
    }

    const lines = [
      docHonorific,
      "お疲れ様です。",
      `ドクターサポート部の${staffLabel}です。`,
      "",
      `この度は、${tomorrowPrefix}${mdDay}のご勤務につきましてご相談がありご連絡させていただきました。`,
      `現在、【${clinicName}】${absenceStart}~${absenceEnd}の時間帯に医師不在となる可能性があるため、`,
      `大変恐縮ではございますがもし先生のご都合がよろしければ、${consultPhrase}`,
      "",
      `・${ymdDay}`
    ];
    if (hasCurrentRow) lines.push(`現在)【${curLabelForFullTime}】${currentStart}~${currentEnd}`);
    lines.push(`ご相談) 【${clinicName}】${targetStartForFullTime}~${targetEndForFullTime}${noteSuffix}`);
    lines.push("");
    lines.push("急なお願いとなり大変恐れ入りますが、ご無理のない範囲でご検討いただけますと幸いです。");
    lines.push("お忙しいところ、ご負担をおかけして恐縮ではございますが、何卒よろしくお願い申し上げます。");
    lines.push("");
    lines.push(`ドクターサポート部\u3000${staffLabel}`);
    return lines.join("\n");
  }

  if (approachType === "早出相談") {
    return [
      docHonorific,
      "平素よりお世話になっております。",
      `ドクターサポート事務局の${staffLabel}でございます。`,
      "",
      `この度は、${tomorrowPrefix}${mdDay}のご勤務につきましてご相談がありご連絡させていただきました。`,
      `現在、【${clinicName}】${absenceStart}~${absenceEnd}の時間帯に医師不在となる可能性があるため、`,
      `大変恐縮ではございますがもし先生のご都合がよろしければ、ご勤務時間を${diffHoursText}前倒ししていただくことは可能でしょうか。`,
      "",
      `・${ymdDay}`,
      `現在)【${curClinicDisp}】${currentStart}~${currentEnd}`,
      `日給:${curWageStr}円`,
      `ご相談) 【${clinicName}】${consultStart}~${currentEnd}${noteSuffix}`,
      `日給:${conWageStr}円`,
      "",
      "急なお願いとなり大変恐れ入りますが、ご無理のない範囲でご検討いただけますと幸いです。",
      "お忙しいところ、ご負担をおかけして恐縮ではございますが、何卒よろしくお願い申し上げます。",
      "",
      `ドクターサポート事務局\u3000${staffLabel}`
    ].join("\n");
  }

  if (approachType === "延長相談") {
    if (isMorningToNightExt) {
      return [
        docHonorific,
        "平素よりお世話になっております。",
        `ドクターサポート事務局の${staffLabel}でございます。`,
        "",
        `この度は、${tomorrowPrefix}${mdDay}のご勤務につきましてご相談がありご連絡させていただきました。`,
        `現在、【${clinicName}】${absenceStart}~${absenceEnd}の時間帯に医師不在となる可能性があるため、`,
        "長時間のご勤務となり大変恐縮ではございますが、もし先生のご都合がよろしければ、可能な限りご勤務時間を延長していただくことは可能でしょうか。",
        `18:00まで、あるいは${consultEnd}までなど、可能な範囲でご検討いただけますと幸いです。`,
        "",
        `・${ymdDay}`,
        `現在)【${curClinicDisp}】${currentStart}~${currentEnd}`,
        `日給:${curWageStr}円`,
        "",
        `ご相談①) 【${clinicName}】${currentStart}~18:00${pm1NoteSuffix}`,
        `日給:${pm1WageStr}円`,
        `ご相談②) 【${clinicName}】${currentStart}~${consultEnd}${noteSuffix}`,
        `日給:${conWageStr}円`,
        "",
        "長時間にわたり大きなご負担をおかけするお願いとなり大変心苦しい限りではございますが、ご無理のない範囲でご検討いただけますと幸いです。",
        "お忙しいところ誠に恐縮ではございますが、何卒よろしくお願い申し上げます。",
        "",
        `ドクターサポート事務局\u3000${staffLabel}`
      ].join("\n");
    }

    return [
      docHonorific,
      "平素よりお世話になっております。",
      `ドクターサポート事務局の${staffLabel}でございます。`,
      "",
      `この度は、${tomorrowPrefix}${mdDay}のご勤務につきましてご相談がありご連絡させていただきました。`,
      `現在、【${clinicName}】${absenceStart}~${absenceEnd}の時間帯に医師不在となる可能性があるため、`,
      `大変恐縮ではございますがもし先生のご都合がよろしければ、ご勤務時間を${diffHoursText}延長していただくことは可能でしょうか。`,
      "",
      `・${ymdDay}`,
      `現在)【${curClinicDisp}】${currentStart}~${currentEnd}`,
      `日給:${curWageStr}円`,
      `ご相談) 【${clinicName}】${currentStart}~${consultEnd}${noteSuffix}`,
      `日給:${conWageStr}円`,
      "",
      "急なお願いとなり大変恐れ入りますが、ご無理のない範囲でご検討いただけますと幸いです。",
      "お忙しいところ、ご負担をおかけして恐縮ではございますが、何卒よろしくお願い申し上げます。",
      "",
      `ドクターサポート事務局\u3000${staffLabel}`
    ].join("\n");
  }

  if (approachType === "勤務地変更相談") {
    if (shiftType === "応募") {
      const m = dateObj.getMonth() + 1;
      return [
        docHonorific,
        "平素よりお世話になっております。",
        `ドクターサポート事務局の${staffLabel}でございます。`,
        "",
        `このたびは${m}月に多くのご応募いただき、誠にありがとうございます。`,
        `恐れ入りますが、${mdDay}のご応募につきまして、ご相談差し上げたくご連絡いたしました。`,
        "医師不在の拠点が多く、もし先生のご都合がよろしければ、ご勤務地を変更いただくことは可能でしょうか。",
        "その際は時給を高時給の拠点に合わせてご案内させていただきます。",
        "",
        `■${ymdDay}`,
        `現在)【${fromClinicName}】${currentStart}~${currentEnd} 日給${curWageStr}円`,
        `ご相談) 【${clinicName}】${consultStart}~${consultEnd} 日給${conWageStr}円`,
        "",
        "難しい場合は現在の拠点にて採用させていただきますので、お手数をおかけいたしますが、ご検討いただけましたら幸いです。",
        "何卒よろしくお願い申し上げます。",
        "",
        `ドクターサポート事務局\u3000${staffLabel}`
      ].join("\n");
    }

    const allowSuffix = hasRequestAllow ? "+依頼手当10,000円" : "";

    if (Array.isArray(multiTargets) && multiTargets.length > 1) {
      const toNames = multiTargets.map(function(t) { return `【${t.clinicName}】`; }).join(" または ");
      const consultLines = multiTargets.map(function(t, idx) {
        const cW = formatComma_(t.consultDailyWage);
        const aSuf = t.hasRequestAllow ? "+依頼手当10,000円" : allowSuffix;
        return `ご相談${getCircledNumber_(idx + 1)}) 【${t.clinicName}】${t.consultStart || consultStart}~${t.consultEnd || consultEnd}(日給:${cW}円${aSuf})`;
      });
      return [
        docHonorific,
        "平素よりお世話になっております。",
        `ドクターサポート事務局の${staffLabel}でございます。`,
        "",
        `ご勤務のご相談が重なり恐縮ですが、${tomorrowPrefix}${mdDay}のご勤務につきましてご相談でございます。`,
        `当日のご勤務拠点を 【${fromClinicName}】 より ${toNames} へご変更いただくことは可能でしょうか。`,
        "",
        `■${mdDay}`,
        `現在)【${fromClinicName}】 ${currentStart}~${currentEnd} (日給:${curWageStr}円)`,
        "↓",
        ...consultLines,
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

    return [
      docHonorific,
      "平素よりお世話になっております。",
      `ドクターサポート事務局の${staffLabel}でございます。`,
      "",
      `ご勤務のご相談が重なり恐縮ですが、${tomorrowPrefix}${mdDay}のご勤務につきましてご相談でございます。`,
      `当日のご勤務拠点を 【${fromClinicName}】 より 【${clinicName}】 へご変更いただくことは可能でしょうか。`,
      "",
      `■${mdDay}`,
      `現在)【${fromClinicName}】 ${currentStart}~${currentEnd} (日給:${curWageStr}円)`,
      "↓",
      `ご相談) 【${clinicName}】${consultStart}~${consultEnd}(日給:${conWageStr}円${allowSuffix})`,
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

  const lines = [
    docHonorific,
    "平素よりお世話になっております。",
    `ドクターサポート事務局の${staffLabel}でございます。`,
    "",
    `この度は、${tomorrowPrefix}${mdDay}のご勤務につきましてご相談がありご連絡させていただきました。`,
    `現在、【${clinicName}】${absenceStart}~${absenceEnd}の時間帯に医師不在となる可能性があるため、`,
    "大変恐縮ではございますがもし先生のご都合がよろしければ、ご勤務いただくことは可能でしょうか。",
    "",
    `・${ymdDay}`,
    `ご相談) 【${clinicName}】${absenceStart}~${absenceEnd}`,
    `日給:${conWageStr}円`
  ];
  if (isPublicRecruit !== false) {
    lines.push("");
    lines.push("※上記募集は公開中のため、ご応募いただいた際に充足している可能性もございます。");
    lines.push("恐れ入りますが、充足の際はご容赦くださいませ。");
  }
  lines.push("");
  lines.push("急なお願いとなり大変恐れ入りますが、ご無理のない範囲でご検討いただけますと幸いです。");
  lines.push("お忙しいところ、ご負担をおかけして恐縮ではございますが、何卒よろしくお願い申し上げます。");
  lines.push("");
  lines.push(`ドクターサポート事務局\u3000${staffLabel}`);
  return lines.join("\n");
}
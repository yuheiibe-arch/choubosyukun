/**
 * ============================================================================
 * 02a_3_Bundle_Templates.gs（旧02b_Setup 4分割の3）
 * おまとめ打診用のシンプル送信文案テンプレート3種
 * 1) 同日複数拠点（AまたはB）への移動相談
 * 2) 複数枠の個別アプローチ（追加勤務）
 * 3) 複数日程の移動・延長・個別が混在する統合おまとめ
 * ============================================================================
 */

function pickValidStaffFromEntries_(entries) {
  for (let i = 0; i < (entries || []).length; i++) {
    const clean = extractCleanStaffName_(entries[i].staffName);
    if (clean) return clean;
  }
  return "{担当者名}";
}

/**
 * パターン1：同日に複数拠点（AまたはB）へ勤務地変更相談を行う際のシンプルテンプレート
 */
function buildSameDateMultiClinicMoverTemplate_(entries) {
  const first = entries[0];
  const docName = String(first.doctorName || "{医師名}").replace(/先生$/g, "");
  const staff = pickValidStaffFromEntries_(entries);
  const mdDay = typeof formatMdDay_ === "function" ? formatMdDay_(first.dateObj) : first.dateStr;
  const fromClinic = first.fromClinicName || "現在拠点";
  const curTime = `${first.currentStart || ""}~${first.currentEnd || ""}`;
  const curWageNum = Number(first.currentDailyWage || 0);
  const curWageStr = curWageNum > 0 ? ` (日給:${formatComma_(curWageNum)}円)` : "";

  const toHeaderStr = entries.map(function(e) { return `【${e.clinicName}】`; }).join(" または ");
  const circledNums = ["①", "②", "③", "④", "⑤", "⑥"];

  const consultLines = entries.map(function(e, idx) {
    const numLabel = circledNums[idx] || (idx + 1);
    const conTime = `${e.consultStart || e.absenceStart || ""}~${e.consultEnd || e.absenceEnd || ""}`;
    const conWageNum = Number(e.consultDailyWage || 0);
    const allowSuffix = e.hasRequestAllow ? "+依頼手当10,000円" : "";
    const wagePart = conWageNum > 0
      ? ` (日給:${formatComma_(conWageNum)}円${allowSuffix})`
      : (allowSuffix ? ` (${allowSuffix.replace(/^\+/, "")})` : "");
    return `ご相談${numLabel}) 【${e.clinicName}】${conTime}${wagePart}`;
  });

  return [
    docName + "先生",
    "平素よりお世話になっております。",
    "ドクターサポート事務局の" + staff + "でございます。",
    "",
    "この度は、" + mdDay + "のご勤務につきましてご相談がありご連絡させていただきました。",
    "ご相談の案が多く大変申し訳ございません。",
    "当日のご勤務拠点を 【" + fromClinic + "】 より " + toHeaderStr + " へご変更いただくことは可能でしょうか。",
    "",
    "■" + mdDay,
    "現在)【" + fromClinic + "】 " + curTime + curWageStr,
    "↓",
    ...consultLines,
    "",
    "※上記募集は公開中のため、ご応募いただいた際に充足している可能性もございます。",
    "恐れ入りますが、充足の際はご容赦くださいませ。",
    "",
    "お忙しいところ、ご負担をおかけして恐縮ではございますが、ご検討いただけましたら幸いです。",
    "何卒よろしくお願い申し上げます。",
    "",
    "ドクターサポート事務局\u3000" + staff
  ].join("\n");
}

/**
 * パターン2：個別アプローチ（追加勤務）で複数枠をまとめて打診する際のシンプルテンプレート
 */
function buildMultiDateIndividualTemplate_(entries) {
  const first = entries[0];
  const docName = String(first.doctorName || "{医師名}").replace(/先生$/g, "");
  const staff = pickValidStaffFromEntries_(entries);

  const slotLines = entries.map(function(e) {
    const ymdLabel = typeof formatYmdDay_ === "function" ? formatYmdDay_(e.dateObj) : e.dateStr;
    const absRange = (e.absenceStart || "") + "~" + (e.absenceEnd || "");
    const conWageNum = Number(e.consultDailyWage || 0);
    const wageText = conWageNum > 0 ? " (日給:" + formatComma_(conWageNum) + "円)" : "";
    return "・" + ymdLabel + " 【" + e.clinicName + "】" + absRange + wageText;
  });

  return [
    docName + "先生",
    "平素より大変お世話になっております。",
    "ドクターサポート事務局の" + staff + "でございます。",
    "",
    "この度は、下記日程のご勤務につきましてご相談がありご連絡させていただきました。",
    "ご相談の案が多く大変申し訳ございません。",
    "現在、下記枠にて医師不在となる可能性がある時間帯がございますため、",
    "もし先生のご都合がよろしければ、いずれかの日程でもご勤務いただくことは可能でしょうか。",
    "",
    "■ご相談日程",
    ...slotLines,
    "",
    "※上記募集は公開中のため、ご応募いただいた際に充足している可能性もございます。",
    "恐れ入りますが、充足の際はご容赦くださいませ。",
    "",
    "急なお願いとなり大変恐れ入りますが、ご無理のない範囲でご検討いただけますと幸いです。",
    "お忙しいところ、ご負担をおかけして恐縮ではございますが、何卒よろしくお願い申し上げます。",
    "",
    "ドクターサポート事務局\u3000" + staff
  ].join("\n");
}

/**
 * パターン3：複数日程の移動、または「移動＋個別」「延長＋別枠」など複数枠が混在する場合のシンプル統合テンプレート
 */
function buildUnifiedMultiSlotDraftTemplate_(entries) {
  const first = entries[0];
  const docName = String(first.doctorName || "{医師名}").replace(/先生$/g, "");
  const staff = pickValidStaffFromEntries_(entries);
  const circledNums = ["①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧"];

  const blocks = entries.map(function(e, idx) {
    const numLabel = circledNums[idx] || (idx + 1);
    const ymdLabel = typeof formatYmdDay_ === "function" ? formatYmdDay_(e.dateObj) : e.dateStr;
    const curTime = `${e.currentStart || ""}~${e.currentEnd || ""}`;
    const conTime = `${e.consultStart || e.absenceStart || ""}~${e.consultEnd || e.absenceEnd || ""}`;
    const curWageNum = Number(e.currentDailyWage || 0);
    const conWageNum = Number(e.consultDailyWage || 0);
    const curWageStr = curWageNum > 0 ? ` (日給:${formatComma_(curWageNum)}円)` : "";
    const allowSuffix = (e.approachType === "勤務地変更相談" && e.hasRequestAllow) ? "+依頼手当10,000円" : "";
    const conWageStr = conWageNum > 0
      ? ` (日給:${formatComma_(conWageNum)}円${allowSuffix})`
      : (allowSuffix ? ` (${allowSuffix.replace(/^\+/, "")})` : "");

    if (e.approachType === "勤務地変更相談") {
      const fromC = e.fromClinicName || "現在拠点";
      return [
        `■ご相談${numLabel}：${ymdLabel}`,
        `現在)【${fromC}】${curTime}${curWageStr}`,
        `↓`,
        `ご相談)【${e.clinicName}】${conTime}${conWageStr}`
      ].join("\n");
    } else if (e.approachType === "延長相談" || e.approachType === "早出相談") {
      return [
        `■ご相談${numLabel}：${ymdLabel}`,
        `現在)【${e.clinicName}】${curTime}${curWageStr}`,
        `↓`,
        `ご相談)【${e.clinicName}】${conTime}${conWageStr}`
      ].join("\n");
    } else {
      const absTime = `${e.absenceStart || ""}~${e.absenceEnd || ""}`;
      return [
        `■ご相談${numLabel}：${ymdLabel}`,
        `ご相談)【${e.clinicName}】${absTime}${conWageStr}`
      ].join("\n");
    }
  });

  return [
    docName + "先生",
    "平素より大変お世話になっております。",
    "ドクターサポート事務局の" + staff + "でございます。",
    "",
    "この度は、下記のご勤務につきましてご相談がありご連絡させていただきました。",
    "ご相談の案が多く大変申し訳ございません。",
    "もし先生のご都合がよろしければ、下記のいずれかの枠にてご協力をいただくことは可能でしょうか。",
    "",
    blocks.join("\n\n"),
    "",
    "※上記募集は公開中のため、ご応募いただいた際に充足している可能性もございます。",
    "恐れ入りますが、充足の際はご容赦くださいませ。",
    "",
    "急なお願いとなり大変恐れ入りますが、ご無理のない範囲でご検討いただけますと幸いです。",
    "お忙しいところ、ご負担をおかけして恐縮ではございますが、何卒よろしくお願い申し上げます。",
    "",
    "ドクターサポート事務局\u3000" + staff
  ].join("\n");
}
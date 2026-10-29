/**
 * ============================================================================
 * 07_WebApp_Bunankun.gs
 * 🌐 【文案くん連携】ワンクリックで「D6作業メモ転記 ＋ Slack親スレッド返信 ＋ 履歴記録」を行うWebアプリ
 * ============================================================================
 */

/**
 * 文案くんのリンクをクリックした時に開くWebアプリ窓口
 */
function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  const targetDate = normalizeDateStr_(p.date || "");
  const targetClinic = String(p.clinic || "").trim();
  const targetDept = String(p.dept || "小児科").trim();
  const targetTime = normalizeTimeRange_(p.time || "");
  const isSpecial = (p.sp === "1" || p.sp === "true" || p.sp === "特別時給");
  const approachType = String(p.type || "移動").trim();
  const doctorName = String(p.doc || "").trim();
  const staffName = String(p.staff || "").trim();
  const rawMemo = String(p.memo || "").trim();
  const isExecute = (p.exec === "1");

  if (!targetDate || !targetClinic || !rawMemo) {
    return HtmlService.createHtmlOutput(
      `<div style="font-family:sans-serif;padding:24px;"><h3>⚠️ パラメータが不足しています</h3><p>文案くんのリンクから再度開いてください。</p></div>`
    ).setTitle("文案くん連携エラー");
  }

  const memoToAppend = rawMemo.startsWith("・") ? rawMemo : `・${rawMemo}`;
  const spTag = isSpecial ? "【特別時給】" : "";
  const dateHeader = formatSlackDateHeader_(targetDate);
  const slackBody = `\`${dateHeader}\`\n【${targetClinic}】${targetTime}${spTag}\n${memoToAppend}`;

  // 1. 「実行する」ボタンが押された後の処理（D6転記 ＋ Slackスレッド返信 ＋ D7履歴記録）
  if (isExecute) {
    try {
      const result = executeDirectSlackAndD6Post_({
        targetDate,
        targetClinic,
        targetDept,
        targetTime,
        approachType,
        doctorName,
        staffName,
        memoToAppend,
        slackBody
      });

      return HtmlService.createHtmlOutput(`
        <div style="font-family:sans-serif;max-width:560px;margin:30px auto;padding:24px;border:2px solid #2e7d32;border-radius:12px;background:#f1f8e9;">
          <h2 style="color:#2e7d32;margin-top:0;">✅ Slack投稿 ＆ D6作業メモ転記が完了しました！</h2>
          <p><strong>対象枠：</strong> ${targetDate} 【${targetClinic}】 ${targetTime} ${spTag}</p>
          <p><strong>D6作業メモ：</strong> ${result.d6Status}</p>
          <p><strong>Slack投稿先：</strong> ${result.slackThreadInfo} (${result.postId})</p>
          <pre style="background:#fff;padding:12px;border-radius:8px;border:1px solid #ccc;">${slackBody}</pre>
          <p style="color:#555;font-size:13px;">※このタブは閉じていただいて大丈夫です。</p>
        </div>
      `).setTitle("✅ Slack投稿・D6転記完了");
    } catch (err) {
      return HtmlService.createHtmlOutput(`
        <div style="font-family:sans-serif;max-width:560px;margin:30px auto;padding:24px;border:2px solid #c62828;border-radius:12px;background:#ffebee;">
          <h2 style="color:#c62828;margin-top:0;">❌ エラーが発生しました</h2>
          <p>${err.message}</p>
        </div>
      `).setTitle("❌ エラー");
    }
  }

  // 2. 初回クリック時：確認画面を表示（「医師への送信なども済みですか？」の確認ワンクッション）
  const baseUrl = ScriptApp.getService().getUrl();
  const execUrl = `${baseUrl}?exec=1&date=${encodeURIComponent(targetDate)}&clinic=${encodeURIComponent(targetClinic)}&dept=${encodeURIComponent(targetDept)}&time=${encodeURIComponent(targetTime)}&sp=${isSpecial ? "1" : "0"}&type=${encodeURIComponent(approachType)}&doc=${encodeURIComponent(doctorName)}&staff=${encodeURIComponent(staffName)}&memo=${encodeURIComponent(memoToAppend)}`;

  return HtmlService.createHtmlOutput(`
    <div style="font-family:sans-serif;max-width:560px;margin:30px auto;padding:24px;border:1px solid #ddd;border-radius:12px;box-shadow:0 4px 12px rgba(0,0,0,0.08);">
      <h3 style="margin-top:0;color:#1565c0;">📢 Slack投稿・D6作業メモ転記の確認</h3>
      <p style="color:#d32f2f;font-weight:bold;">医師へのメッセージ送信は完了していますか？</p>
      <div style="background:#f5f5f5;padding:14px;border-radius:8px;margin-bottom:20px;">
        <div style="font-size:13px;color:#666;margin-bottom:6px;">▼ Slack投稿プレビュー（超募集MTGログ 親スレッドへ自動返信）</div>
        <pre style="margin:0;font-size:14px;white-space:pre-wrap;">${slackBody}</pre>
      </div>
      <ul style="font-size:14px;color:#333;line-height:1.7;">
        <li><strong>D6アプローチシート転記：</strong> ${targetDate} 【${targetClinic}】 ${targetTime} の作業メモ欄へ追記</li>
        <li><strong>打診種別 / 医師 / 担当：</strong> ${approachType} / ${doctorName}先生 / 担当：${staffName || "未指定"}</li>
      </ul>
      <div style="margin-top:24px;text-align:center;">
        <a href="${execUrl}" onclick="this.style.pointerEvents='none';this.innerText='⏳ 送信・転記処理中...';" style="display:inline-block;background:#1976d2;color:#fff;text-decoration:none;font-weight:bold;padding:14px 28px;border-radius:8px;font-size:16px;">
          🚀 送信済み・Slack投稿＆D6転記を実行する
        </a>
      </div>
    </div>
  `).setTitle("Slack投稿・D6転記の確認");
}

/**
 * 既存の関数群を使って D6転記・Slack投稿・D7履歴記録を実行する内部関数
 */
function executeDirectSlackAndD6Post_(params) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const {
      targetDate,
      targetClinic,
      targetDept,
      targetTime,
      approachType,
      doctorName,
      staffName,
      memoToAppend,
      slackBody
    } = params;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const histSheet = ss.getSheetByName(SHEET_NAMES.D7_HISTORY);
    const logSheet = ss.getSheetByName(SHEET_NAMES.D7_LOG);

    // 1. D6『アプローチシート』の該当行を検索して作業メモを追記
    const unfilledSs = SpreadsheetApp.openById(EXTERNAL_IDS.D6_D3_UNFILLED_SS);
    const d6Sheet = unfilledSs.getSheetByName(SHEET_NAMES.D6_APPROACH);
    const d6Headers = d6Sheet.getRange(1, 1, 1, d6Sheet.getLastColumn()).getDisplayValues()[0];
    const colDate = d6Headers.indexOf("日付") + 1;
    const colClinic = d6Headers.indexOf("拠点名") + 1;
    const colTime = d6Headers.indexOf("不在時間") + 1;
    const colMemo = d6Headers.indexOf("作業メモ") + 1;

    const d6LastRow = d6Sheet.getLastRow();
    const d6Keys = d6LastRow >= 3
      ? d6Sheet.getRange(1, 1, d6LastRow, Math.max(colDate, colClinic, colTime)).getDisplayValues()
      : [];

    const targetNormClinic = normalizeClinicForMatch_(targetClinic);
    const exactTimeD6Rows = [];
    const sameClinicDateD6Rows = [];

    for (let r = 2; r < d6Keys.length; r++) {
      const rDate = normalizeDateStr_(d6Keys[r][colDate - 1]);
      if (rDate !== targetDate) continue;

      const rClinic = String(d6Keys[r][colClinic - 1] || "").trim();
      const rNormClinic = normalizeClinicForMatch_(rClinic);
      const rTime = normalizeTimeRange_(d6Keys[r][colTime - 1]);

      const isKitaOrKame = (targetNormClinic === "北葛西" || targetNormClinic === "亀有");
      const isClinicMatch = isKitaOrKame
        ? (rClinic === targetClinic || (rNormClinic === targetNormClinic && (!targetDept || rClinic.includes(targetDept))))
        : (rClinic === targetClinic || rNormClinic === targetNormClinic);

      if (isClinicMatch) {
        sameClinicDateD6Rows.push(r + 1);
        if (!targetTime || !rTime || rTime === targetTime || isTimeOverlapping_(rTime, targetTime)) {
          exactTimeD6Rows.push(r + 1);
        }
      }
    }

    const matchedD6Rows = exactTimeD6Rows.length > 0 ? exactTimeD6Rows : sameClinicDateD6Rows;
    let d6Status = "該当枠なし（Slack投稿のみ完了）";

    if (matchedD6Rows.length > 0 && memoToAppend) {
      for (const matchedD6Row of matchedD6Rows) {
        const memoCell = d6Sheet.getRange(matchedD6Row, colMemo);
        const currentMemo = String(memoCell.getDisplayValue() || "").trim();
        const cleanTargetMemo = memoToAppend.replace(/^・/, "").trim();
        const lines = currentMemo ? currentMemo.split(/\r?\n/).map(l => l.trim()).filter(Boolean) : [];
        const alreadyExists = lines.some(l => l.replace(/^・/, "").trim() === cleanTargetMemo || l.includes(cleanTargetMemo));
        if (!alreadyExists) {
          lines.push(memoToAppend);
          memoCell.setValue(lines.join("\n"));
        }
      }
      SpreadsheetApp.flush();
      d6Status = `${matchedD6Rows.length}件の枠（行:${matchedD6Rows.join(",")}）に「${memoToAppend}」を追記完了`;
    }

    // 2. Slackの「超募集MTGログ」親スレッドを検索して返信投稿
    const token = PropertiesService.getScriptProperties().getProperty("SLACK_BOT_TOKEN");
    const parentThreadMap = fetchSlackParentThreadMap_(token, SLACK_CHANNEL_ID);
    const threadTitle = buildSlackThreadTitle_(targetDate);
    const postId = postToSlackThread_(threadTitle, slackBody, targetDate, parentThreadMap);
    const { halfKey } = getSlackThreadKeyAndHeader_(targetDate, threadTitle);

    // 3. D7『打診履歴』と『ログ』に証跡を記録
    const nowStr = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");
    const userEmail = Session.getActiveUser().getEmail();
    const actor = staffName || userEmail;

    if (histSheet && doctorName) {
      appendRowAndTrimBlank_(
        histSheet,
        [nowStr, targetDate, targetClinic, targetTime, targetDept, approachType, `${doctorName}先生`, actor, postId],
        9
      );
      applyLeftMiddleToSheet_(histSheet);
    }

    if (logSheet) {
      const actionLabel = matchedD6Rows.length > 0 ? "文案くん連携:D6転記・Slack投稿完了" : "文案くん連携:Slack投稿完了(D6該当枠なし)";
      appendRowAndTrimBlank_(
        logSheet,
        [nowStr, actionLabel, postId, `${targetDate} ${targetClinic}${targetTime}`, 1, `${memoToAppend} (${actor})`],
        6
      );
      applyLeftMiddleToSheet_(logSheet);
    }

    return {
      d6Status,
      postId,
      slackThreadInfo: `${halfKey} 超募集MTGログ`
    };
  } finally {
    lock.releaseLock();
  }
}
/**
 * Slack API が JSON ではなく HTML ("<!DOCTYPE ...") を返している原因を特定する診断スクリプト
 * ※メッセージ投稿（chat.postMessage / chat.update）は一切行わない安全な読み取り専用テストです
 */
function debugSlackHtmlError() {
  console.log("==================================================");
  console.log("🔍 1. 定数・スクリプトプロパティの健全性チェック");
  console.log("==================================================");

  const props = PropertiesService.getScriptProperties();
  const rawToken = props.getProperty("SLACK_BOT_TOKEN");
  const chId = (typeof SLACK_CHANNEL_ID !== "undefined") ? SLACK_CHANNEL_ID : "(未定義)";

  console.log("・SLACK_CHANNEL_ID:", JSON.stringify(chId));
  if (!rawToken) {
    console.log("❌ SLACK_BOT_TOKEN がスクリプトプロパティに存在しません！（空です）");
  } else {
    const trimmed = rawToken.trim();
    console.log("・SLACK_BOT_TOKEN 文字数:", rawToken.length, "(trim後:", trimmed.length + ")");
    console.log("・先頭8文字:", trimmed.slice(0, 8) + "...");
    console.log("・改行や空白の混入:", rawToken !== trimmed || /[\r\n\s]/.test(trimmed) ? "⚠️ あり（要注意！）" : "なし");
    console.log("・非ASCII文字(全角等)の混入:", /[^\x20-\x7E]/.test(trimmed) ? "⚠️ あり（要注意！）" : "なし");
  }

  console.log("\n==================================================");
  console.log("🔍 2. cleanUrl_ 生成後の実際のURLとコード定義チェック");
  console.log("==================================================");

  const rawHistBase = "https://" + "slack.com/api/conversations.history";
  const cleanedHistBase = (typeof cleanUrl_ === "function") ? cleanUrl_(rawHistBase) : rawHistBase;
  const fullHistUrl = cleanedHistBase + "?channel=" + encodeURIComponent(chId) + "&limit=5";
  console.log("・テスト対象URL:", fullHistUrl);

  if (typeof fetchSlackParentThreadMap_ === "function") {
    console.log("・現在デプロイされている fetchSlackParentThreadMap_ の先頭12行:");
    console.log(fetchSlackParentThreadMap_.toString().split("\n").slice(0, 12).join("\n"));
  } else {
    console.log("⚠️ fetchSlackParentThreadMap_ 関数が見つかりません！");
  }

  console.log("\n==================================================");
  console.log("🔍 3. Slack API 通信テスト（HTTPステータス＆返却HTMLの中身確認）");
  console.log("==================================================");

  if (rawToken && chId !== "(未定義)") {
    const endpoints = [
      { name: "auth.test", url: "https://slack.com/api/auth.test" },
      { name: "conversations.history (cleanUrl経由)", url: fullHistUrl },
      { name: "conversations.history (POST形式テスト)", url: "https://slack.com/api/conversations.history", isPost: true }
    ];

    endpoints.forEach(function(ep) {
      try {
        const options = ep.isPost
          ? {
              method: "post",
              contentType: "application/x-www-form-urlencoded",
              headers: { Authorization: "Bearer " + rawToken.trim() },
              payload: { channel: chId, limit: "5" },
              muteHttpExceptions: true,
              followRedirects: false
            }
          : {
              method: "get",
              headers: { Authorization: "Bearer " + rawToken.trim() },
              muteHttpExceptions: true,
              followRedirects: false
            };

        const res = UrlFetchApp.fetch(ep.url, options);
        const code = res.getResponseCode();
        const headers = res.getAllHeaders();
        const body = res.getContentText();

        console.log(`\n📡 [${ep.name}]`);
        console.log(`  ・HTTP Status: ${code}`);
        if (headers["Location"] || headers["location"]) {
          console.log(`  ・Redirect Location: ${headers["Location"] || headers["location"]}`);
        }
        console.log(`  ・Content-Type: ${headers["Content-Type"] || headers["content-type"] || "不明"}`);
        console.log(`  ・Response先頭300文字:\n${body.slice(0, 300)}`);
      } catch (e) {
        console.log(`❌ [${ep.name}] 例外エラー: ${e.message}`);
      }
    });
  }

  console.log("\n==================================================");
  console.log("🔍 4. シート上の「エラー」行（2列目・24列目）の確認");
  console.log("==================================================");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ["オフライン修正シート", "文案・Slack投稿レビュー"].forEach(function(sName) {
    const sh = ss.getSheetByName(sName);
    if (!sh || sh.getLastRow() < 2) return;
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 24).getDisplayValues();
    let errCount = 0;
    vals.forEach(function(r, idx) {
      if (r[0] === "TRUE" || r[1] === "エラー" || r[23]) {
        errCount++;
        console.log(`  📄 [${sName} ${idx + 2}行目] Check=${r[0]} | 状態=${r[1]} | 日付=${r[3].replace(/\n/g, ",")} | 拠点=${r[4].replace(/\n/g, ",")} | X列エラー=${r[23]}`);
      }
    });
    if (errCount === 0) {
      console.log(`  📄 [${sName}] チェックONまたはエラー状態の行はありません`);
    }
  });
}
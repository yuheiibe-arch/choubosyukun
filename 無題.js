/**
 * Gmail受信JSONの生データと、インポート本体（runJsonImportCore_）のコード全文を確認するデバッグスクリプト
 */
function debugInspectGmailBatchAndImporter() {
  console.log("==================================================");
  console.log("🔍 1. Gmailから取得される最新JSONバッチの生データ確認");
  console.log("==================================================");

  [true, false].forEach(isOffline => {
    const label = isOffline ? "オフラインモード(true)" : "通常音声モード(false)";
    try {
      const batch = collectLatestJsonBatchFromGmail_(isOffline);
      if (!batch || !batch.items || batch.items.length === 0) {
        console.log(`\n📭 [${label}] 該当メールなし`);
        return;
      }
      console.log(`\n📬 [${label}] 件数: ${batch.items.length}件 / latestMsgId: ${batch.latestMsgId}`);
      batch.items.forEach((item, idx) => {
        console.log(`  [${idx + 1}] ${JSON.stringify(item)}`);
      });
    } catch (e) {
      console.log(`⚠️ [${label}] 取得エラー: ${e.message}`);
    }
  });

  console.log("\n==================================================");
  console.log("🔍 2. インポート・パース系関数のソースコード全文出力");
  console.log("==================================================");

  const targetFuncs = [
    "collectLatestJsonBatchFromGmail_",
    "runJsonImportCore_",
    "consolidateMultiDateDoctorDrafts_"
  ];

  const globalObj = this;
  targetFuncs.forEach(fnName => {
    if (typeof globalObj[fnName] !== "function") {
      console.log(`❌ 関数未定義: ${fnName}`);
      return;
    }
    const lines = globalObj[fnName].toString().split(/\r?\n/);
    const chunkSize = 70; // GASログ省略を防ぐため70行ずつ分割出力
    for (let i = 0; i < lines.length; i += chunkSize) {
      const partNum = Math.floor(i / chunkSize) + 1;
      const totalParts = Math.ceil(lines.length / chunkSize);
      console.log(`\n--- 📄 ${fnName} (${partNum}/${totalParts}) ---\n` + lines.slice(i, i + chunkSize).join("\n"));
    }
  });
}
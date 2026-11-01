// ==========================================
// 【一高祭・文化部門】チケット発行＆チェックイン管理システム（完全版・テスト機能付）
// ==========================================

const GAT_PROP_LAST = "GAT_lastProcessed";
const GAT_PROP_RUNNING = "GAT_running";
const GAT_PROP_BATCH_SIZE = "GAT_batchSize";
const GAT_PROP_TRIGGER_FN = "continueGenerateAllTickets";

// ------------------------------------------
//  入場パート定義（A/B/C の3パート制）
//  終了時刻は設定せず、開始時刻を過ぎたパートのチケットは
//  以降ずっと入場可能（＝Cパートの時間になれば A・B・C 全てのチケットが入場可能）
// ------------------------------------------
const TICKET_PARTS = {
  a: { label: "Aパート（9:15）", startHour: 9, startMinute: 15 },
  b: { label: "Bパート（11:00）", startHour: 10, startMinute: 45 },
  c: { label: "Cパート（13:00）", startHour: 12, startMinute: 55 }
};

/**
 * 時間帯コード（a/b/c/any等）から表示用ラベルを返す
 */
function getTimeLabel(time) {
  const part = TICKET_PARTS[String(time || "").toLowerCase()];
  return part ? part.label : "制限なし";
}

/**
 * 指定パートのチケットが「今」入場可能かどうかを判定する。
 * 開始時刻を過ぎていれば true（終了時刻の概念はない）。
 * a/b/c 以外（"any" 等、管理者即席発券分）は常に true。
 */
function isEntryAllowedNow(time) {
  const part = TICKET_PARTS[String(time || "").toLowerCase()];
  if (!part) return true;

  const now = new Date();
  const hour = parseInt(Utilities.formatDate(now, "Asia/Tokyo", "H"), 10);
  const minute = parseInt(Utilities.formatDate(now, "Asia/Tokyo", "m"), 10);
  const nowMinutes = hour * 60 + minute;
  const startMinutes = part.startHour * 60 + part.startMinute;

  return nowMinutes >= startMinutes;
}

/**
 * 「フォームの回答」シートのD列（時間帯）・E列（枚数）を集計し、
 * パート別（a/b/c）の申請済み枚数合計を返す
 */
function sumCountsByPart(sheet, lastRow) {
  const totals = { a: 0, b: 0, c: 0 };
  if (lastRow < 2) return totals;

  const values = sheet.getRange(2, 4, lastRow - 1, 2).getValues();
  for (let i = 0; i < values.length; i++) {
    const t = String(values[i][0] || "").toLowerCase();
    const c = parseInt(values[i][1]) || 0;
    if (totals.hasOwnProperty(t)) {
      totals[t] += c;
    }
  }
  return totals;
}

// パートごとの定員（申請枚数の上限）。必要に応じて調整してください。
const PART_CAPACITY = {
  a: 500,
  b: 800,
  c: 800
};

// ------------------------------------------
//  一括チケット発行・バッチ処理関連関数
// ------------------------------------------

/**
 * 管理者ページからのバッチ発行開始
 * @param {number} batchSize 1回あたり処理する件数
 * @param {boolean} autoContinue 残りがある場合1分毎に自動継続するか
 */
function startGenerateAllTickets(batchSize, autoContinue) {
  batchSize = Number(batchSize) || 50;
  if (autoContinue === undefined) autoContinue = true;
  
  const props = PropertiesService.getScriptProperties();
  props.setProperty(GAT_PROP_BATCH_SIZE, String(batchSize));
  
  if (!props.getProperty(GAT_PROP_LAST)) {
    props.setProperty(GAT_PROP_LAST, "1");
  }
  
  return generateAllTicketsBatch(batchSize, autoContinue);
}

/**
 * トリガー等から呼ばれる自動継続用エントリポイント
 */
function continueGenerateAllTickets() {
  const props = PropertiesService.getScriptProperties();
  const batchSize = Number(props.getProperty(GAT_PROP_BATCH_SIZE) || 50);
  generateAllTicketsBatch(batchSize, true);
}

/**
 * 実際のバッチ処理本体
 */
function generateAllTicketsBatch(batchSize, autoContinue) {
  if (autoContinue === undefined) autoContinue = true;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    return { success: false, message: "別の処理が実行中です。" };
  }

  const props = PropertiesService.getScriptProperties();
  try {
    props.setProperty(GAT_PROP_RUNNING, "1");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("フォームの回答");
    if (!sheet) {
      writeLog("error", "generateAllTicketsBatch: シート「フォームの回答」が見つかりません");
      props.deleteProperty(GAT_PROP_RUNNING);
      return { success: false, message: "シートが見つかりません" };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      props.deleteProperty(GAT_PROP_RUNNING);
      props.deleteProperty(GAT_PROP_LAST);
      removeContinueTriggerIfExists();
      return { success: true, message: "対象データがありません", processedCount: 0 };
    }

    let lastProcessed = Number(props.getProperty(GAT_PROP_LAST) || 1);
    let nextRow = lastProcessed + 1;
    if (nextRow > lastRow) {
      props.deleteProperty(GAT_PROP_RUNNING);
      props.deleteProperty(GAT_PROP_LAST);
      removeContinueTriggerIfExists();
      return { success: true, message: "全件処理完了済みです", processedCount: 0 };
    }

    const endRow = Math.min(nextRow + batchSize - 1, lastRow);
    const numRows = endRow - nextRow + 1;
    const range = sheet.getRange(nextRow, 1, numRows, 7);
    const values = range.getValues();

    let processedCount = 0;

    for (let i = 0; i < values.length; i++) {
      const rowIndex = nextRow + i;
      try {
        const row = values[i];
        const mailSent = row[6]; // G列: ステータス
        if (mailSent === "sent" || mailSent === "instant_sent") {
          continue;
        }

        const email = row[1];
        const name = row[2];
        const time = row[3];
        const count = Number(row[4]) || 0;
        if (!email || !name || !count) {
          continue;
        }

        // チケットURLの生成（未生成の場合のみ作成）
        let ticketUrls = [];
        const existingUrlsStr = String(row[5] || "").trim();
        if (existingUrlsStr !== "") {
          ticketUrls = existingUrlsStr.split("\n");
        } else {
          for (let j = 0; j < count; j++) {
            const ticketId = "TKT-" + Utilities.getUuid();
            const ticketUrl =
              "https://ticket.dai1-sc.online/?id=" +
              ticketId;
            ticketUrls.push(ticketUrl);
          }
          sheet.getRange(rowIndex, 6).setValue(ticketUrls.join("\n"));
        }

        // メール送信処理
        const subject = "【一高祭・文化部門】入場電子チケットのご案内（" + count + "枚）";
        const timeLabel = getTimeLabel(time);

        let body = name + "さん\n\n";
        body += "一高祭・文化部門のチケットを " + count + " 枚発行しました。\n";
        body += "入場時間帯：" + timeLabel + "\n";
        body += "以下のリンクからチケットを確認できます。\n\n";
        ticketUrls.forEach((url, index) => {
          body += "【チケット " + (index + 1) + "】\n" + url + "\n\n";
        });

        const htmlBody = `
<div style="font-family: Arial, sans-serif; max-width: 600px; margin: auto; padding: 20px;">
  <div style="background-color: #002366; padding: 15px; text-align: center; border-radius: 8px 8px 0 0;">
    <h1 style="color: white; font-size: 20px; margin: 0;">一高祭・文化部門 入場チケット</h1>
  </div>
  <div style="background-color: #f9f9f9; padding: 20px; border: 1px solid #ddd; border-top: none;">
    <p style="font-size: 16px;">${name}さん</p>
    <p>一高祭・文化部門のチケットを <strong>${count}枚</strong> 発行しました。<br>
    入場時間帯：<strong>${timeLabel}</strong><br>
    以下のボタンからチケットを確認してください。</p>
    ${ticketUrls.map((url, index) => `
    <div style="text-align: center; margin: 20px 0;">
      <a href="${url}" style="background-color: #002366; color: white; padding: 12px 30px; text-decoration: none; border-radius: 6px; font-size: 16px; display: inline-block; font-weight: bold;">
        チケット ${index + 1} を開く
      </a>
    </div>
    `).join("")}
    <p style="font-size: 12px; color: #888; margin-top: 30px;">
      ※このメールは自動送信です。<br>
      ※お問い合わせは tt03dai1-seitokai@g.bears.ed.jp までご連絡ください。
    </p>
  </div>
  <div style="background-color: #002366; padding: 10px; text-align: center; border-radius: 0 0 8px 8px;">
    <p style="color: white; font-size: 12px; margin: 0;">熊本県立第一高等学校 生徒会</p>
  </div>
</div>
`;
        try {
          if (!email || email === "dai1-sc-instant@g.bears.ed.jp") {
            sheet.getRange(rowIndex, 7).setValue("instant_sent");
          } else {
            GmailApp.sendEmail(email, subject, body, {
              htmlBody: htmlBody,
              name: "第一高校 生徒会",
              from: "tt03-dai1seitokai@g.bears.ed.jp"
            });
            sheet.getRange(rowIndex, 7).setValue("sent");
          }
          processedCount++;
        } catch (err) {
          writeLog("error", "generateAllTicketsBatch メール送信失敗: " + err.message);
        }
      } catch (rowErr) {
        writeLog("error", "generateAllTicketsBatch 行処理失敗（行 " + (nextRow + i) + "）: " + rowErr.message);
      }
    }

    props.setProperty(GAT_PROP_LAST, String(endRow));

    if (endRow < lastRow) {
      if (autoContinue) {
        ensureContinueTriggerExists();
      }
    } else {
      props.deleteProperty(GAT_PROP_LAST);
      removeContinueTriggerIfExists();
    }

    return {
      success: true,
      message: `${numRows} 行中 ${processedCount} 件の送信処理を行いました`,
      processedCount: processedCount,
      isFinished: endRow >= lastRow
    };

  } finally {
    props.deleteProperty(GAT_PROP_RUNNING);
    lock.releaseLock();
  }
}

/**
 * リアルタイム進捗確認関数（管理者ページ用）
 */
function getBatchProgress() {
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("フォームの回答");

  if (!sheet) {
    return { error: "シートが見つかりません" };
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return {
      total: 0,
      urlGenerated: 0,
      mailSent: 0,
      lastProcessedRow: 1,
      isRunning: false,
      hasTrigger: false,
      progressPercent: 100
    };
  }

  const total = lastRow - 1;
  const values = sheet.getRange(2, 6, total, 2).getValues(); // F列:URL, G列:ステータス

  let urlGenerated = 0;
  let mailSent = 0;

  for (let i = 0; i < values.length; i++) {
    const urlVal = String(values[i][0] || "").trim();
    const statusVal = String(values[i][1] || "").trim();

    if (urlVal !== "") urlGenerated++;
    if (statusVal === "sent" || statusVal === "instant_sent") mailSent++;
  }

  const isRunning = props.getProperty(GAT_PROP_RUNNING) === "1";
  const lastProcessedRow = Number(props.getProperty(GAT_PROP_LAST) || 1);
  const hasTrigger = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === GAT_PROP_TRIGGER_FN);
  const progressPercent = total > 0 ? Math.floor((mailSent / total) * 100) : 0;

  return {
    total: total,
    urlGenerated: urlGenerated,
    mailSent: mailSent,
    lastProcessedRow: lastProcessedRow,
    isRunning: isRunning,
    hasTrigger: hasTrigger,
    progressPercent: progressPercent
  };
}

/**
 * バッチ処理の停止
 */
function stopBatchGeneration() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(GAT_PROP_RUNNING);
  removeContinueTriggerIfExists();
  return { success: true, message: "自動トリガーを停止しました" };
}

/**
 * 進捗位置のリセット
 */
function clearGenerateAllTicketsProgress() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(GAT_PROP_LAST);
  props.deleteProperty(GAT_PROP_RUNNING);
  removeContinueTriggerIfExists();
  return { success: true, message: "進捗をリセットしました" };
}

function ensureContinueTriggerExists() {
  removeContinueTriggerIfExists();
  ScriptApp.newTrigger(GAT_PROP_TRIGGER_FN).timeBased().after(60 * 1000).create();
  writeLog("info", "次バッチ用の自動継続トリガーを作成しました（60秒後に実行予定）");
}

function removeContinueTriggerIfExists() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === GAT_PROP_TRIGGER_FN) {
      ScriptApp.deleteTrigger(t);
    }
  });
}

// ------------------------------------------
//  管理者用即席チケット発行関数
// ------------------------------------------
function generateTickets(name, email, count) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("フォームの回答");
  if (!sheet) throw new Error("シートが見つかりません。");
  const insertRow = sheet.getLastRow() + 1;

  let ticketUrls = [];
  for (let i = 0; i < count; i++) {
    const ticketId = "TKT-" + Utilities.getUuid();
    const ticketUrl =
      "https://script.google.com/a/macros/g.bears.ed.jp/s/AKfycbxdMLfj6qHITIEInB52DC1DJqojKlb-dWb43AoqMchlkCFyuqRpZtqB_0u6vcFydGw/exec?id=" +
      ticketId;
    ticketUrls.push(ticketUrl);
  }
  const urlsString = ticketUrls.join("\n");

  sheet.getRange(insertRow, 1).setValue(new Date());
  sheet.getRange(insertRow, 2).setValue(email || "dai1-sc-instant@g.bears.ed.jp");
  sheet.getRange(insertRow, 3).setValue(name);
  sheet.getRange(insertRow, 4).setValue("any");
  sheet.getRange(insertRow, 5).setValue(count);
  sheet.getRange(insertRow, 6).setValue(urlsString);
  sheet.getRange(insertRow, 7).setValue(!email || email === "dai1-sc-instant@g.bears.ed.jp" ? "instant_sent" : "sent");

  // 【重要】doPost内のロックが解放される前に書き込みを確定させる
  SpreadsheetApp.flush();

  if (!email || email === "dai1-sc-instant@g.bears.ed.jp") {
    return urlsString;
  } else {
    const subject = "一高祭・文化部門チケットのご案内（" + count + "枚）";
    let body = name + "さん\n\n管理者から文化部門のチケットが " + count + " 枚発行されました。\n\n";
    ticketUrls.forEach((url, index) => {
      body += "【チケット " + (index + 1) + "】\n" + url + "\n\n";
    });

    const htmlBody = `
<div style="font-family: Arial, sans-serif; max-width: 600px; margin: auto; padding: 20px;">
  <div style="background-color: #002366; padding: 15px; text-align: center; border-radius: 8px 8px 0 0;">
    <h1 style="color: white; font-size: 20px; margin: 0;">一高祭・文化部門 入場チケット</h1>
  </div>
  <div style="background-color: #f9f9f9; padding: 20px; border: 1px solid #ddd; border-top: none;">
    <p style="font-size: 16px;">${name}さん</p>
    <p>管理者から文化部門のチケットが <strong>${count}枚</strong> 発行されました。</p>
    ${ticketUrls.map((url, index) => `
    <div style="text-align: center; margin: 20px 0;">
      <a href="${url}" style="background-color: #002366; color: white; padding: 12px 30px; text-decoration: none; border-radius: 6px; font-size: 16px; display: inline-block; font-weight: bold;">
        チケット ${index + 1} を開く
      </a>
    </div>
    `).join("")}
  </div>
</div>`;

    try {
      GmailApp.sendEmail(email, subject, body, {
        htmlBody: htmlBody,
        name: "第一高校 生徒会",
        from: "tt03-dai1seitokai@g.bears.ed.jp"
      });
    } catch (err) {
      writeLog("error", "メール送信失敗（generateTickets）: " + err.message);
    }
    return urlsString;
  }
}

// ------------------------------------------
//  GETリクエスト処理（表示・通信API全般）
// ------------------------------------------
function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const mode = params.mode || "";

  // 1. 管理者画面の表示
  if (mode === "admin") {
    return HtmlService.createHtmlOutputFromFile("admin")
      .setTitle("管理者ページ")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  // 2. バッチ進捗状況の取得
  if (mode === "get_progress") {
    return ContentService.createTextOutput(JSON.stringify(getBatchProgress()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 3. バッチ発行の開始
  if (mode === "start_batch") {
    const batchSize = Number(params.batchSize) || 50;
    const autoContinue = params.autoContinue !== "false";
    const res = startGenerateAllTickets(batchSize, autoContinue);
    return ContentService.createTextOutput(JSON.stringify(res))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 4. バッチ発行の停止
  if (mode === "stop_batch") {
    return ContentService.createTextOutput(JSON.stringify(stopBatchGeneration()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 5. 進捗位置のリセット
  if (mode === "clear_progress") {
    return ContentService.createTextOutput(JSON.stringify(clearGenerateAllTicketsProgress()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 6. 管理者照会用チケット検索
  if (mode === "search_ticket") {
    const keyword = params.keyword || "";
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("フォームの回答");
    const last = sheet.getLastRow();
    let result = null;

    for (let row = 2; row <= last; row++) {
      const email = String(sheet.getRange(row, 2).getValue() || "");
      const name = String(sheet.getRange(row, 3).getValue() || "");

      if (email.includes(keyword) || name.includes(keyword)) {
        result = {
          name: name,
          email: email,
          count: sheet.getRange(row, 5).getValue(),
          tickets: (sheet.getRange(row, 6).getValue() || "").toString().split("\n")
        };
        break;
      }
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  }

  // 6.5. 改ざん検知ログの記録
  if (mode === "log_tampering") {
    const studentName = params.studentName || "不明";
    logTampering(studentName);
    return ContentService.createTextOutput(JSON.stringify({ status: "logged" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 7. 【復元】申請フォーム用：残り枚数集計
  if (mode === "remaining") {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("フォームの回答");
    const lastRow = sheet.getLastRow();
    const totals = sumCountsByPart(sheet, lastRow);

    return ContentService.createTextOutput(JSON.stringify(totals))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 8. 【復元・修正】申請フォーム用：登録存在チェック
  // 【重要・修正】以前はセルを1件ずつ getValue() で読んでおり、
  // 台帳の行数（最大1000行想定）に比例して非常に遅くなっていた。
  // 1000人規模の同時アクセス時にこの関数が長時間 実行中 のままになり、
  // 同時実行数の枯渇（"too many scripts running simultaneously"）の一因になっていたため、
  // 1回のbulk読み込み（getRange().getValues()）にまとめて高速化した。
  if (mode === "check") {
    const email = params.email || "";
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("フォームの回答");
    const lastRow = sheet.getLastRow();

    if (lastRow >= 2) {
      // B列（email）〜G列（ステータス）を一括取得
      const rows = sheet.getRange(2, 2, lastRow - 1, 6).getValues();
      for (let i = 0; i < rows.length; i++) {
        if (rows[i][0] === email) {
          return ContentService.createTextOutput(JSON.stringify({
            exists: true,
            time: rows[i][2],   // D列（フォームの回答のB列基準で+2列目）
            count: rows[i][3],  // E列
            mailSent: rows[i][5] // G列
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ exists: false }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 9. 申請フォーム用：新規申し込み処理
  if (mode === "submit") {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {
      return ContentService.createTextOutput(JSON.stringify({ status: "busy" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 【重要・修正】以前はメール送信（GmailApp.sendEmail）と endSession() の
    // 呼び出しまで、この申請書き込み用のグローバルロックを握ったまま実行していた。
    // メール送信はGoogleのメールサーバーとの通信が入るため数百ms〜1秒以上
    // かかることがあり、その間ずっと他の申請者（最大200人が同時にここへ来る想定）が
    // 足止めされてしまう。ロックが必要なのは「スプレッドシートへの書き込み確定まで」
    // だけなので、ロック解放後にメール送信・セッション終了を行うよう分離し、
    // ロック保持時間を最小化した。
    const email = params.email || "";
    const name = params.name || "";
    const time = params.time || "";
    const count = parseInt(params.count) || 0;
    const sessionId = params.sessionId || "";
    let writeResult = "error"; // "error" | "full" | "duplicate" | "ok"

    try {
      if (email && name && time && count && TICKET_PARTS.hasOwnProperty(time)) {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("フォームの回答");
        const lastRow = sheet.getLastRow();

        const partTotals = sumCountsByPart(sheet, lastRow);
        const current = partTotals[time];

        if (current + count > PART_CAPACITY[time]) {
          writeResult = "full";
        } else {
          // 【追加】同一メールでの二重登録防止（クライアント側の自動リトライで
          // 同じ申請が複数回届いた場合の保険）
          let isDuplicate = false;
          if (lastRow >= 2) {
            const existingEmails = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
            for (let i = 0; i < existingEmails.length; i++) {
              if (existingEmails[i][0] === email) {
                isDuplicate = true;
                break;
              }
            }
          }

          if (isDuplicate) {
            // 既に登録済みなので、新規追加はせず「成功」として返す
            // （重複送信によるメール連投を避けるため、ここではメールは再送しない）
            writeResult = "duplicate";
          } else {
            const insertRow = lastRow + 1;
            sheet.getRange(insertRow, 1, 1, 5).setValues([[new Date(), email, name, time, count]]);

            // 【重要】ロックを解放する前に、書き込みを必ずスプレッドシートへ確定させる。
            // これが無いと、書き込みが内部的にまだ確定していない状態で
            // 次のリクエストがロックを取得してしまい、getLastRow() が古い値を返して
            // 同じ行に上書きしてしまう（＝データ消失）ことがある。
            SpreadsheetApp.flush();
            writeResult = "ok";
          }
        }
      }
    } finally {
      lock.releaseLock();
    }

    // ここから先はロックを離した後の処理（他の申請者を待たせない）
    if (writeResult === "ok") {
      try {
        sendReceiptEmail(email, name, time, count);
      } catch (err) {
        writeLog("error", "受付メール送信失敗: " + err.message);
      }
    }
    if (writeResult === "ok" || writeResult === "duplicate") {
      if (sessionId) endSession(sessionId);
      return ContentService.createTextOutput(JSON.stringify({ status: "ok" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (writeResult === "full") {
      return ContentService.createTextOutput(JSON.stringify({ status: "full" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({ status: "error" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 10. 【排他制御対応】申請フォーム用：申し込み内容編集処理
  if (mode === "edit") {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {
      return ContentService.createTextOutput(JSON.stringify({ status: "busy" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const email = params.email || "";
    const name = params.name || "";
    const time = params.time || "";
    const count = parseInt(params.count) || 0;
    let writeResult = "full"; // "full" | "ok"

    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("フォームの回答");
      let lastRow = sheet.getLastRow();

      if (lastRow >= 2) {
        const emails = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
        for (let i = emails.length - 1; i >= 0; i--) {
          if (emails[i][0] === email) {
            sheet.deleteRow(i + 2);
            break;
          }
        }
      }

      lastRow = sheet.getLastRow();
      const partTotals = sumCountsByPart(sheet, lastRow);
      const current = partTotals[time] || 0;

      if (TICKET_PARTS.hasOwnProperty(time) && current + count <= PART_CAPACITY[time]) {
        const insertRow = lastRow + 1;
        sheet.getRange(insertRow, 1, 1, 5).setValues([[new Date(), email, name, time, count]]);

        // 【重要】申請受付処理と同様、ロック解放前に書き込みを確定させる
        SpreadsheetApp.flush();
        writeResult = "ok";
      }
    } finally {
      lock.releaseLock();
    }

    // メール送信はロック解放後に行い、他の編集・申請者を待たせないようにする
    if (writeResult === "ok") {
      try {
        sendReceiptEmail(email, name, time, count);
      } catch (err) {
        writeLog("error", "編集完了メール送信失敗: " + err.message);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "ok" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: "full" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 13. セッション作成・取得
  if (mode === "get_or_create_session") {
    const email = params.email || "";
    const result = getOrCreateSession(email);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 14. セッション状態確認（待機列チェック）
  if (mode === "check_session_status") {
    const sessionId = params.sessionId || "";
    const result = checkSessionStatus(sessionId);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 15. セッション明示終了 API
  if (mode === "end_session") {
    const sessionId = params.sessionId || "";
    const result = endSession(sessionId);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // 11. チェックイン画面の表示
  if (mode === "checkin") {
    return HtmlService.createHtmlOutputFromFile("checkin")
      .setTitle("チェックイン")
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  // 12. チケット表示処理（API応答・HTML表示 両対応）
  const ticketId = params.id || "";

  // ★【テスト用分岐】TEST-9999 または TEST- から始まるIDはシート更新無しの仮想チケットとして応答 ★
  if (ticketId === "TEST-9999" || ticketId.indexOf("TEST-") === 0) {
    const testResult = {
      ticketId: ticketId,
      studentName: "【テスト用】動作確認",
      partName: "制限なし（動作確認用）",
      status: "未入場",
      logoBase64: "",
      ticketUrl: "https://ticket.dai1-sc.online/?id=" + ticketId
    };
    if (mode === "json" || params.id) {
      return ContentService.createTextOutput(JSON.stringify(testResult))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("フォームの回答");
  const data = sheet.getDataRange().getValues();

  let studentName = "不明";
  let status = "未入場";
  let partName = "－";
  let found = false;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const ticketsStr = String(row[5] || "");
    if (!ticketsStr) continue;

    const ticketList = ticketsStr.split("\n");
    for (const tUrl of ticketList) {
      if (tUrl.includes(ticketId)) {
        studentName = String(row[2] || "不明");
        const timeSlot = String(row[3] || "").toLowerCase();
        partName = getTimeLabel(timeSlot);

        if (tUrl.includes("[入場済み]")) status = "入場済み";
        found = true;
        break;
      }
    }
    if (found) break;
  }

  const logoFileId = "1Q3vJAGHXNaqf1Zt7MYl9H-ptGMDVn8je";
  let logoBase64 = "";
  try {
    const logoBlob = DriveApp.getFileById(logoFileId).getBlob();
    const resized = logoBlob.getAs("image/png");
    logoBase64 = Utilities.base64Encode(resized.getBytes());
  } catch (err) {
    writeLog("error", "doGet ロゴ取得失敗: " + err.message);
  }

  const result = {
    ticketId: ticketId,
    studentName: studentName,
    partName: partName,
    status: status,
    logoBase64: logoBase64,
    ticketUrl: "https://ticket.dai1-sc.online/?id=" + ticketId
  };

  if (mode === "json" || params.id) {
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const template = HtmlService.createTemplateFromFile("index");
  template.ticketId = ticketId;
  template.ticketUrl = result.ticketUrl;
  template.studentName = studentName;
  template.status = status;
  template.partName = partName;
  template.logoBase64 = logoBase64;

  return template.evaluate().setTitle("電子チケット").addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// POSTリクエスト処理（QR読取・即席発券・ログ記録等）
function doPost(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return ContentService.createTextOutput("BUSY");
  }

  try {
    if (e && e.postData && e.postData.contents) {
      let checkData = JSON.parse(e.postData.contents);
      if (checkData.action === "health_check") {
        let health = checkServerHealth();
        return ContentService.createTextOutput(JSON.stringify(health)).setMimeType(ContentService.MimeType.JSON);
      }
    }

    let params = {};
    let ticketId = "";
    let mode = "";

    if (e && e.postData && e.postData.contents) {
      params = JSON.parse(e.postData.contents);
      ticketId = params.ticketId || "";
      mode = params.mode || "";
    } else {
      params = e && e.parameter ? e.parameter : {};
      ticketId = params.id || "";
      mode = params.mode || "";
    }

    if (mode === "log_invalid") {
      writeLog("invalid", params.data || "unknown");
      return ContentService.createTextOutput("logged");
    }

    if (mode === "log_tampering") {
      const studentName = params.studentName || "不明";
      logTampering(studentName);
      return ContentService.createTextOutput("logged");
    }

    // 【追加】撮影許可証交付の記録処理
    if (mode === "photo_permit") {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let permitSheet = ss.getSheetByName("撮影許可");
      if (!permitSheet) {
        permitSheet = ss.insertSheet("撮影許可");
        permitSheet.appendRow(["タイムスタンプ", "チケットID", "名前", "メールアドレス", "チケットURL"]);
      }

      if (ticketId === "TEST-9999" || ticketId.indexOf("TEST-") === 0) {
        permitSheet.appendRow([new Date(), ticketId, "【テスト用】動作確認", "test@example.com", "https://ticket.dai1-sc.online/?id=" + ticketId]);
        SpreadsheetApp.flush();
        return ContentService.createTextOutput("PERMIT_OK");
      }

      const sheet = ss.getSheetByName("フォームの回答");
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
        for (let i = 0; i < data.length; i++) {
          const ticketsRaw = String(data[i][5] || "");
          if (ticketsRaw.includes(ticketId)) {
            const email = data[i][1];
            const studentName = data[i][2];
            const ticketUrls = ticketsRaw.split("\n");
            let targetUrl = "";
            for (const url of ticketUrls) {
              if (url.includes(ticketId)) {
                targetUrl = url;
                break;
              }
            }
            permitSheet.appendRow([new Date(), ticketId, studentName, email, targetUrl]);
            SpreadsheetApp.flush();
            writeLog("photo_permit", `招待生徒 ${studentName} (${ticketId}) の撮影許可証交付を記録しました`);
            return ContentService.createTextOutput("PERMIT_OK");
          }
        }
      }
      return ContentService.createTextOutput("NOT_FOUND");
    }
    
    if (mode === "admin_issue") {
      try {
        const urls = generateTickets(params.name, params.email || "dai1-sc-instant@g.bears.ed.jp", Number(params.count));
        writeLog("admin_issue", `管理者が ${params.name} さんに ${params.count} 枚発行(即席)`);
        return ContentService.createTextOutput(urls);
      } catch (err) {
        return ContentService.createTextOutput("ERROR: " + err);
      }
    }

    if (!ticketId) return ContentService.createTextOutput("notfound");

    // ★【テスト用分岐】 TEST-9999 または TEST- から始まるチケットIDの場合はスプレッドシートの書き換えを行わずに成功を返す ★
    if (ticketId === "TEST-9999" || ticketId.indexOf("TEST-") === 0) {
      let responseText = "OK";
      let testLog = "【テスト実行】テスト用チケットが初回入場しました";

      if (mode === "exit") {
        responseText = "EXIT_OK";
        testLog = "【テスト実行】テスト用チケットが一時退場しました";
      } else if (mode === "re_entry") {
        responseText = "RE_ENTRY_OK";
        testLog = "【テスト実行】テスト用チケットが再入場しました";
      }

      writeLog("checkin", testLog);
      // "レスポンス文字|入場区分" の形式で返却
      return ContentService.createTextOutput(responseText + "|any");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("フォームの回答");
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return ContentService.createTextOutput("notfound");

    const dataRange = sheet.getRange(2, 1, lastRow - 1, 6);
    const data = dataRange.getValues();

    let found = false;
    let studentName = "";
    let rowNum = -1;
    let tickets = [];
    let ticketTime = "";

    for (let i = 0; i < data.length; i++) {
      const ticketsRaw = data[i][5];
      if (ticketsRaw && ticketsRaw.includes(ticketId)) {
        found = true;
        studentName = data[i][2];
        ticketTime = data[i][3];
        rowNum = i + 2;
        tickets = ticketsRaw.toString().split("\n");
        break;
      }
    }

    if (!found) {
      writeLog("invalid", `無効なチケットが読み込まれました（内容：${ticketId}）`);
      return ContentService.createTextOutput("notfound");
    }

    // 【A/B/C 3パート制】終了時刻は設定せず、各パートの開始時刻を過ぎていれば
    // そのパートのチケットは入場可能（＝Cパートの時間になればA・Bのチケットも入場可）。
    // まだ開始前のパートのチケットのみ弾く。
    if (!isEntryAllowedNow(ticketTime)) {
      return ContentService.createTextOutput("WRONG_TIME_" + String(ticketTime).toUpperCase());
    }

    let currentStatus = "未入場"; 
    for (let j = 0; j < tickets.length; j++) {
      if (tickets[j].includes(ticketId)) {
        if (tickets[j].includes("[入場済み]")) currentStatus = "入場済み";
        else if (tickets[j].includes("[退場済み]")) currentStatus = "退場済み";
        break;
      }
    }

    let nextStatus = "";
    let responseText = "";
    let logMessage = "";

    if (mode === "entry" || mode === "") {
      if (currentStatus === "入場済み" || currentStatus === "退場済み") return ContentService.createTextOutput("already");
      nextStatus = "[入場済み]";
      responseText = "OK";
      logMessage = `招待生徒 ${studentName} が初回入場しました`;

    } else if (mode === "exit") {
      if (currentStatus === "未入場") return ContentService.createTextOutput("ERROR_NOT_IN_YET");
      if (currentStatus === "退場済み") return ContentService.createTextOutput("ALREADY_OUT");
      nextStatus = "[退場済み]";
      responseText = "EXIT_OK";
      logMessage = `招待生徒 ${studentName} が一時退場しました`;

    } else if (mode === "re_entry") {
      if (currentStatus === "未入場") return ContentService.createTextOutput("ERROR_NOT_IN_YET");
      if (currentStatus === "入場済み") return ContentService.createTextOutput("ALREADY_IN");
      nextStatus = "[入場済み]";
      responseText = "RE_ENTRY_OK";
      logMessage = `招待生徒 ${studentName} が再入場しました`;
    }

    const updatedTickets = tickets.map(url => {
      if (url.includes(ticketId)) {
        let baseUrl = url.replace(" [入場済み]", "").replace(" [退場済み]", "");
        return baseUrl + " " + nextStatus;
      }
      return url;
    });

    sheet.getRange(rowNum, 6).setValue(updatedTickets.join("\n"));
    // 【重要】複数端末でほぼ同時にチェックインが行われた場合の
    // 上書き事故を防ぐため、ロック解放前に書き込みを確定させる
    SpreadsheetApp.flush();
    writeLog("checkin", logMessage);

    return ContentService.createTextOutput(responseText + "|" + ticketTime);

  } catch (error) {
    writeLog("error", "doPostエラー: " + (error && error.message ? error.message : String(error)));
    return ContentService.createTextOutput("SERVER_ERROR");
  } finally {
    lock.releaseLock();
  }
}

// ------------------------------------------
//  共通ヘルパー関数（ログ・通信等）
// ------------------------------------------
function writeLog(type, message) {
  try {
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ログ");
    if (!logSheet) return;
    const last = logSheet.getLastRow() + 1;
    const now = new Date();
    logSheet.getRange(last, 1).setValue(Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss"));
    logSheet.getRange(last, 2).setValue(type);
    logSheet.getRange(last, 3).setValue(message);
  } catch (err) {
    console.error("writeLog failed:", err);
  }
}

function checkServerHealth() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("フォームの回答");
    if (!sheet) return { status: "error", message: "シート名「フォームの回答」が見つかりません。" };
    return { status: "ok", message: "サーバー（GAS・台帳）は正常に稼働しています！" };
  } catch (error) {
    return { status: "error", message: "システムエラー: " + error.toString() };
  }
}

// ------------------------------------------
//  改ざん検知ログ記録関数
// ------------------------------------------
function logTampering(studentName) {
  const name = studentName || "不明";
  const message = "生徒名" + name + "のチケットの改ざんを検知しました。";
  writeLog("tampering", message);
  return { success: true, message: message };
}

/**
 * 【復元】独自申請フォーム用：受付完了メール送信関数（HTML形式）
 */
function sendReceiptEmail(email, name, time, count) {
  const timeLabel = getTimeLabel(time);

  const subject = "【一高祭・文化部門】お申し込み受付のご確認";

  let body = `${name} 様\n\n`;
  body += "一高祭・文化部門のチケットお申し込みを受け付けました。\n";
  body += "以下の内容でお申し込みを承っております。\n\n";
  body += "----------------------------------------\n";
  body += `お名前 ：${name} 様\n`;
  body += `お申込枚数：${count} 枚\n`;
  body += `ご入場区分：${timeLabel}\n`;
  body += "----------------------------------------\n\n";
  body += "※このメールはアドレスの確認およびお申し込み内容のご確認のために自動送信されています。\n";
  body += "※電子チケットの発行・送付は、全体の申請状況を確認した後に改めて行います。\n";
  body += "※チケット発券までは、お申し込み内容の変更・修正が可能です。\n\n";
  body += "お問い合わせ：tt03-dai1seitokai@g.bears.ed.jp\n";

  const htmlBody = `
  <div style="font-family: 'Helvetica Neue', Helvetica, 'Arial', 'Hiragino Sans', 'メイリオ', sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; color: #1a202c; background-color: #f7fafc;">
    <div style="background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
      <div style="background-color: #0f2942; color: #ffffff; padding: 24px; text-align: center;">
        <h1 style="margin: 0; font-size: 20px; font-weight: bold; letter-spacing: 0.05em;">一高祭・文化部門</h1>
        <p style="margin: 6px 0 0 0; font-size: 13px; color: #cbd5e0;">お申し込み受付のご確認</p>
      </div>
      <div style="padding: 24px;">
        <p style="margin: 0 0 16px 0; font-size: 15px; font-weight: bold;">${name} 様</p>
        <p style="margin: 0 0 20px 0; font-size: 14px; line-height: 1.6; color: #4a5568;">
          一高祭・文化部門のチケットお申し込みを受け付けました。<br>
          以下の内容でお申し込みを承っております。
        </p>
        <div style="background-color: #edf2f7; border-left: 4px solid #1a365d; padding: 16px; margin-bottom: 24px; border-radius: 0 4px 4px 0;">
          <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
            <tr>
              <td style="padding: 6px 0; color: #4a5568; width: 100px; font-weight: bold;">お名前</td>
              <td style="padding: 6px 0; color: #1a202c; font-weight: bold;">：${name} 様</td>
            </tr>
            <tr>
              <td style="padding: 6px 0; color: #4a5568; font-weight: bold;">お申込枚数</td>
              <td style="padding: 6px 0; color: #1a202c; font-weight: bold;">：${count} 枚</td>
            </tr>
            <tr>
              <td style="padding: 6px 0; color: #4a5568; font-weight: bold;">ご入場区分</td>
              <td style="padding: 6px 0; color: #1a202c; font-weight: bold;">：${timeLabel}</td>
            </tr>
          </table>
        </div>
        <div style="font-size: 12px; color: #718096; line-height: 1.6; border-top: 1px solid #e2e8f0; padding-top: 16px;">
          <p style="margin: 0 0 6px 0;">※このメールはアドレスの確認およびお申し込み内容のご確認のために自動送信されています。</p>
          <p style="margin: 0 0 6px 0;">※電子チケットの発行・送付は、全体の申請状況を確認した後に改めて行います。</p>
          <p style="margin: 0 0 6px 0;">※チケット発券までは、お申し込み内容の変更・修正が可能です。</p>
        </div>
      </div>
      <div style="background-color: #edf2f7; padding: 16px; text-align: center; font-size: 12px; color: #718096; border-top: 1px solid #e2e8f0;">
        お問い合わせ：<a href="mailto:tt03-dai1seitokai@g.bears.ed.jp" style="color: #2b4c7e; text-decoration: underline;">tt03-dai1seitokai@g.bears.ed.jp</a>
      </div>
    </div>
  </div>
  `;

  try {
    GmailApp.sendEmail(email, subject, body, {
      htmlBody: htmlBody,
      name: "第一高校 生徒会",
      from: "tt03-dai1seitokai@g.bears.ed.jp"
    });
  } catch (err) {
    writeLog("error", "sendReceiptEmail メール送信失敗: " + err.message);
    throw err;
  }
}

// ==========================================
// 【修正版】セッション管理システム（スプレッドシート方式）
// ==========================================
// 【重要・修正】従来はセッションIDの一覧を SESSION_LIST / WAITING_QUEUE という
// カンマ区切りの単一プロパティに格納していたが、PropertiesServiceの
// 1プロパティあたりの値サイズ上限は 9KB（公式Quotas: Properties value size = 9 KB/val）。
// セッションID（"SID-"+UUID、約41バイト/件）を1つの文字列にまとめると、
// active/waiting合計で約224件を超えた時点で setProperty() が例外を投げ、
// doGet ごとクラッシュしていた（1000人規模のアクセス集中で必ず発生する）。
// 「フォームの回答」シートと同じ発想で、待機キュー専用シートに1行1セッションで
// 保持する方式に変更し、サイズ上限を気にしなくてよい設計にした。

const QUEUE_SHEET_NAME = "待機キュー";
const SESSION_EXPIRY_MINUTES = 15;
const MAX_ACTIVE_SESSIONS = 200;
const MINUTES_PER_GROUP = 3;

/**
 * 待機キューシートの取得（無ければ作成）
 */
function getQueueSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(QUEUE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(QUEUE_SHEET_NAME);
    sheet.appendRow(["sessionId", "email", "status", "createdAt", "expiryAt"]);
  }
  return sheet;
}

/**
 * 待機キューシートを1回のbulk読み込みで配列化
 * （行番号 row を保持しておくことで、後続の更新・削除をピンポイントで行える）
 */
function readQueue_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 5).getValues().map((r, i) => ({
    row: i + 2,
    sessionId: r[0],
    email: r[1],
    status: r[2],
    createdAt: r[3],
    expiryAt: r[4]
  }));
}

/**
 * 期限切れセッションの一括削除（下の行から削除しないと行番号がズレる点に注意）
 */
function purgeExpired_(sheet, rows, now) {
  let changed = false;
  for (let i = rows.length - 1; i >= 0; i--) {
    if (now > Number(rows[i].expiryAt || 0)) {
      sheet.deleteRow(rows[i].row);
      changed = true;
    }
  }
  return changed;
}

/**
 * セッションの作成・取得
 */
function getOrCreateSession(email) {
  if (!email) {
    return { canProceed: false, error: "メールアドレスが不正です" };
  }

  const lock = LockService.getScriptLock();
  // ★ タイムアウトを10秒から30秒に延長して一斉アクセス時の脱落を防止
  if (!lock.tryLock(30000)) {
    // ログイン画面に戻させないよう、一時混雑フラグ(isBusy)を返して再試行させる
    return {
      canProceed: false,
      isBusy: true,
      message: "アクセスが大変混雑しています。自動的に待機列への接続を再試行中..."
    };
  }

  try {
    const sheet = getQueueSheet_();
    let rows = readQueue_(sheet);
    const now = new Date().getTime();

    if (purgeExpired_(sheet, rows, now)) {
      rows = readQueue_(sheet);
    }

    // 同じメールアドレスの既存セッションがあれば使い回す
    const existing = rows.find(r => r.email === email);
    if (existing) {
      if (existing.status === "active") {
        return { canProceed: true, sessionId: existing.sessionId, position: 0 };
      }
      const position = rows.filter(r => r.status === "waiting" && r.row < existing.row).length + 1;
      const estimatedWaitTime = calculateEstimatedWaitTime(position);
      return {
        canProceed: false,
        sessionId: existing.sessionId,
        position: position,
        estimatedWaitTime: estimatedWaitTime,
        message: `現在申請が混雑しています。あと${position}組お待ちください。（予想待機時間：約${estimatedWaitTime}分）`
      };
    }

    // 新規セッション作成
    const sessionId = "SID-" + Utilities.getUuid();
    const expiryAt = now + (SESSION_EXPIRY_MINUTES * 60 * 1000);
    const activeCount = rows.filter(r => r.status === "active").length;
    const status = activeCount < MAX_ACTIVE_SESSIONS ? "active" : "waiting";

    sheet.appendRow([sessionId, email, status, now, expiryAt]);

    if (status === "active") {
      writeLog("session", `新規セッション許可: ${email} (${sessionId})`);
      return { canProceed: true, sessionId: sessionId, position: 0 };
    } else {
      const position = rows.filter(r => r.status === "waiting").length + 1;
      const estimatedWaitTime = calculateEstimatedWaitTime(position);
      writeLog("session", `新規セッション待機: ${email} (${sessionId}, 順位:${position})`);

      return {
        canProceed: false,
        sessionId: sessionId,
        position: position,
        estimatedWaitTime: estimatedWaitTime,
        message: `現在申請が混雑しています。あと${position}組お待ちください。（予想待機時間：約${estimatedWaitTime}分）`
      };
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * 待機中セッションの状態チェック（自動繰り上げ対応）
 * タイムアウト時は必ず isBusy（再試行可能）を返し、ログイン画面へ強制的に
 * 戻さないようにしている（以前この扱いの不一致が原因の不具合があったため）。
 */
function checkSessionStatus(sessionId) {
  if (!sessionId) {
    return { canProceed: false, error: "セッションIDが無効です" };
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    return { canProceed: false, isBusy: true, message: "混雑のため確認を再試行中..." };
  }

  try {
    const sheet = getQueueSheet_();
    let rows = readQueue_(sheet);
    const now = new Date().getTime();

    if (purgeExpired_(sheet, rows, now)) {
      rows = readQueue_(sheet);
    }

    const idx = rows.findIndex(r => r.sessionId === sessionId);
    if (idx === -1) {
      return { canProceed: false, error: "セッションが有効期限切れか存在しません。再読み込みしてください。" };
    }

    // 空きスロットがあれば待機列の先頭から繰り上げ
    const activeCount = rows.filter(r => r.status === "active").length;
    const freeSlots = MAX_ACTIVE_SESSIONS - activeCount;
    if (freeSlots > 0) {
      const waiting = rows
        .filter(r => r.status === "waiting")
        .sort((a, b) => a.row - b.row)
        .slice(0, freeSlots);
      waiting.forEach(r => {
        sheet.getRange(r.row, 3).setValue("active");
        writeLog("session", `待機列から繰り上げ許可: ${r.email} (${r.sessionId})`);
      });
      if (waiting.length > 0) rows = readQueue_(sheet);
    }

    const target = rows.find(r => r.sessionId === sessionId);
    if (target.status === "active") {
      return { canProceed: true, sessionId: sessionId, position: 0, message: "申請フォームをご利用いただけます" };
    }

    const position = rows.filter(r => r.status === "waiting" && r.row < target.row).length + 1;
    const estimatedWaitTime = calculateEstimatedWaitTime(position);

    return {
      canProceed: false,
      sessionId: sessionId,
      position: position,
      estimatedWaitTime: estimatedWaitTime,
      message: `現在あと${position}組お待ちください。（予想待機時間：約${estimatedWaitTime}分）`
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * セッション終了（申請完了時などに呼ぶ）
 */
function endSession(sessionId) {
  if (!sessionId) return { success: false };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { success: false };

  try {
    const sheet = getQueueSheet_();
    const rows = readQueue_(sheet);
    const target = rows.find(r => r.sessionId === sessionId);
    if (target) sheet.deleteRow(target.row);
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 管理者用セッション統計取得
 */
function getSessionStats() {
  const sheet = getQueueSheet_();
  const rows = readQueue_(sheet);
  const now = new Date().getTime();
  const alive = rows.filter(r => now <= Number(r.expiryAt || 0));

  return {
    activeCount: alive.filter(r => r.status === "active").length,
    waitingCount: alive.filter(r => r.status === "waiting").length,
    maxActive: MAX_ACTIVE_SESSIONS
  };
}

function calculateEstimatedWaitTime(position) {
  const pos = Math.max(1, Number(position) || 1);
  const groupNumber = Math.ceil(pos / MAX_ACTIVE_SESSIONS);
  return groupNumber * MINUTES_PER_GROUP;
}

function getUnrecordedEmails(startDate) {
  try {
    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("フォームの回答");

    if (!sheet) {
      return {
        success: false,
        count: 0,
        unrecorded: [],
        message: "シート「フォームの回答」が見つかりません。"
      };
    }

    const timezone = Session.getScriptTimeZone() || "Asia/Tokyo";

    let searchDate;
    if (startDate) {
      const parts = String(startDate).split("-");
      if (parts.length !== 3) {
        throw new Error("日付はYYYY-MM-DD形式で指定してください。");
      }

      searchDate = new Date(
        Number(parts[0]),
        Number(parts[1]) - 1,
        Number(parts[2])
      );

      if (isNaN(searchDate.getTime())) {
        throw new Error("正しい日付を指定してください。");
      }
    } else {
      searchDate = new Date();
      searchDate.setDate(searchDate.getDate() - 7);
    }

    const gmailAfterDate = new Date(searchDate.getTime());
    gmailAfterDate.setDate(gmailAfterDate.getDate() - 1);

    const afterText = Utilities.formatDate(
      gmailAfterDate,
      timezone,
      "yyyy/MM/dd"
    );

    const queries = [
      'from:tt03-dai1seitokai@g.bears.ed.jp after:' + afterText + ' subject:"お申し込み受付のご確認"',
      'from:tt03dai1seitokai@g.bears.ed.jp after:' + afterText + ' subject:"お申し込み受付のご確認"'
    ];

    const registeredEmails = {};

    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const emailValues = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
      emailValues.forEach(function(row) {
        const email = normalizeEmailForCompare(row[0]);
        if (email) {
          registeredEmails[email] = true;
        }
      });
    }

    const unrecordedMap = {};

    queries.forEach(function(query) {
      const threads = GmailApp.search(query);

      threads.forEach(function(thread) {
        thread.getMessages().forEach(function(message) {
          const subject = message.getSubject() || "";
          if (subject.indexOf("お申し込み受付のご確認") === -1) {
            return;
          }

          const from = normalizeEmailForCompare(message.getFrom());
          if (
            from !== "tt03-dai1seitokai@g.bears.ed.jp" &&
            from !== "tt03dai1seitokai@g.bears.ed.jp"
          ) {
            return;
          }

          const recipientText =
            (message.getTo() || "") + "," + (message.getCc() || "");

          const recipients = recipientText.match(
            /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi
          ) || [];

          recipients.forEach(function(recipient) {
            const email = normalizeEmailForCompare(recipient);
            if (!email || registeredEmails[email]) {
              return;
            }

            if (!unrecordedMap[email]) {
              unrecordedMap[email] = {
                email: email,
                date: Utilities.formatDate(
                  message.getDate(),
                  timezone,
                  "yyyy/MM/dd HH:mm:ss"
                ),
                subject: subject
              };
            }
          });
        });
      });
    });

    const unrecorded = Object.keys(unrecordedMap)
      .map(function(email) {
        return unrecordedMap[email];
      })
      .sort(function(a, b) {
        return a.date.localeCompare(b.date);
      });

    return {
      success: true,
      count: unrecorded.length,
      unrecorded: unrecorded,
      startDate: Utilities.formatDate(searchDate, timezone, "yyyy-MM-dd")
    };

  } catch (error) {
    writeLog(
      "error",
      "getUnrecordedEmailsエラー: " + (error && error.message ? error.message : String(error))
    );

    return {
      success: false,
      count: 0,
      unrecorded: [],
      message: error && error.message ? error.message : String(error)
    };
  }
}

function normalizeEmailForCompare(value) {
  if (!value) return "";

  const match = String(value).match(
    /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i
  );

  return match ? match[0].toLowerCase().trim() : "";
}

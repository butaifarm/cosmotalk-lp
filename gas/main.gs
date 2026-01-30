// ============================================================
// コスモトーク 問い合わせフォーム受信
// スプレッドシートへの書き込み + Slack通知
// ============================================================

// Slack Webhook URLはスクリプトプロパティ「SLACK_WEBHOOK_URL」から取得
// GASエディタ → プロジェクトの設定 → スクリプトプロパティ で設定してください
function getSlackWebhookUrl() {
  return PropertiesService.getScriptProperties().getProperty("SLACK_WEBHOOK_URL") || "";
}

// シート名
var SHEET_NAME = "問い合わせ";

/**
 * 初回セットアップ: シートとヘッダーを作成
 * 1回手動実行してください
 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  var headers = ["受信日時", "ご希望内容", "氏名", "会社名", "メールアドレス", "電話番号"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 280);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 220);
  sheet.setColumnWidth(6, 140);
}

/**
 * POSTリクエストを受信
 */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    var inquiryType = data.inquiry_type || "お問い合わせ";
    var name = data.name || "";
    var company = data.company || "";
    var email = data.email || "";
    var phone = data.phone || "";
    var message = data.message || "";
    var timestamp = new Date();

    writeToSheet(timestamp, inquiryType, name, company, email, phone, message);
    sendSlackNotification(timestamp, inquiryType, name, company, email, phone, message);

    return ContentService
      .createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * スプレッドシートにデータを書き込む
 */
function writeToSheet(timestamp, inquiryType, name, company, email, phone, message) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    setup();
    sheet = ss.getSheetByName(SHEET_NAME);
  }

  var formattedDate = Utilities.formatDate(timestamp, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
  sheet.appendRow([formattedDate, inquiryType, name, company, email, phone]);
}

/**
 * Slack通知を送信する
 */
function sendSlackNotification(timestamp, inquiryType, name, company, email, phone, message) {
  var SLACK_WEBHOOK_URL = getSlackWebhookUrl();
  if (!SLACK_WEBHOOK_URL) {
    Logger.log("Slack Webhook URLが未設定です");
    return;
  }

  var formattedDate = Utilities.formatDate(timestamp, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");

  var payload = {
    text: "コスモトーク 新規お問い合わせ【" + inquiryType + "】",
    blocks: [
      {
        type: "header",
        text: {
          type: "plain_text",
          text: "【" + inquiryType + "】新規お問い合わせ"
        }
      },
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: "*ご希望内容:*\n" + inquiryType
        }
      },
      {
        type: "section",
        fields: [
          { type: "mrkdwn", text: "*氏名:*\n" + name },
          { type: "mrkdwn", text: "*会社名:*\n" + company },
          { type: "mrkdwn", text: "*メール:*\n" + email },
          { type: "mrkdwn", text: "*電話番号:*\n" + phone }
        ]
      },
      {
        type: "context",
        elements: [
          { type: "mrkdwn", text: "受信日時: " + formattedDate }
        ]
      }
    ]
  };

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };

  UrlFetchApp.fetch(SLACK_WEBHOOK_URL, options);
}

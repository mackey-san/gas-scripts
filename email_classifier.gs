/**
 * Gmailの「要処理」ラベル付き未読メールをClaude APIで分類し、
 * スプレッドシートへの記録とSlack通知、ラベルの付け替えを行う。
 */

// ラベル・シート名・Claude API関連の設定値
const LABEL_TODO = '要処理';
const LABEL_DONE = '処理済み';
const SHEET_LOG = 'メールログ';
const SHEET_ERROR = 'エラーログ';
const CLAUDE_MODEL = 'claude-haiku-4-5'; // コストを抑えるためHaikuの最新版を使用
const CLAUDE_API_URL = 'https://api.anthropic.com/v1/messages';

// 「要処理」ラベル付き未読メールを分類・記録・通知するメイン処理
function processUnreadEmails() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const claudeApiKey = scriptProperties.getProperty('CLAUDE_API_KEY');
    const slackWebhookUrl = scriptProperties.getProperty('SLACK_WEBHOOK_URL');

    if (!claudeApiKey || !slackWebhookUrl) {
      logError_('スクリプトプロパティに CLAUDE_API_KEY または SLACK_WEBHOOK_URL が設定されていません。');
      return;
    }

    const todoLabel = GmailApp.getUserLabelByName(LABEL_TODO);
    if (!todoLabel) {
      logError_('「' + LABEL_TODO + '」ラベルが見つかりません。Gmail側でラベルを作成してください。');
      return;
    }
    const doneLabel = getOrCreateLabel_(LABEL_DONE);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = getOrCreateSheet_(ss, SHEET_LOG, ['受信日時', '送信者', '件名', '分類', '要約']);

    const threads = todoLabel.getThreads();

    threads.forEach(function (thread) {
      const unreadMessages = thread.getMessages().filter(function (message) {
        return message.isUnread();
      });

      // 未読メールが無いスレッドはスキップ
      if (unreadMessages.length === 0) {
        return;
      }

      let hasError = false;

      unreadMessages.forEach(function (message) {
        try {
          processMessage_(message, claudeApiKey, slackWebhookUrl, logSheet);
          message.markRead();
        } catch (error) {
          hasError = true;
          logError_('メール処理中にエラーが発生しました（件名: ' + message.getSubject() + '）: ' + error);
        }
      });

      // スレッド内の全メールが正常に処理できた場合のみラベルを付け替える
      if (!hasError) {
        thread.addLabel(doneLabel);
        thread.removeLabel(todoLabel);
      }
    });
  } catch (error) {
    logError_('processUnreadEmails全体でエラーが発生しました: ' + error);
  }
}

// 1件のメールをClaude APIで分類し、スプレッドシートへの記録とSlack通知を行う
function processMessage_(message, claudeApiKey, slackWebhookUrl, logSheet) {
  const subject = message.getSubject();
  const from = message.getFrom();
  const receivedAt = message.getDate();
  const body = message.getPlainBody();

  const classification = classifyEmail_(claudeApiKey, subject, body);

  // 受信日時・送信者・件名・分類・要約の順で「メールログ」シートに記録
  logSheet.appendRow([receivedAt, from, subject, classification.category, classification.summary]);

  // Slackへ通知
  notifySlack_(slackWebhookUrl, subject, classification.category, classification.summary);
}

// Claude APIにメール本文を送り、分類結果と要約をJSONで受け取る
function classifyEmail_(apiKey, subject, body) {
  const prompt =
    '以下のメールを読み、内容を「クレーム」「質問」「注文」「その他」のいずれか1つに分類してください。\n' +
    'また、メール内容を1〜2文の日本語で要約してください。\n' +
    '出力は説明文を含めず、次のJSON形式のみを返してください。\n' +
    '{"category": "分類結果", "summary": "要約文"}\n\n' +
    '件名: ' + subject + '\n' +
    '本文:\n' + body;

  const payload = {
    model: CLAUDE_MODEL,
    max_tokens: 1024,
    messages: [
      { role: 'user', content: prompt }
    ]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(CLAUDE_API_URL, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode !== 200) {
    throw new Error('Claude APIの呼び出しに失敗しました（ステータス' + responseCode + '）: ' + responseText);
  }

  const responseJson = JSON.parse(responseText);
  const textBlock = responseJson.content && responseJson.content[0] && responseJson.content[0].text;
  if (!textBlock) {
    throw new Error('Claude APIのレスポンスにテキストが含まれていません: ' + responseText);
  }

  let result;
  try {
    result = JSON.parse(textBlock);
  } catch (parseError) {
    throw new Error('Claude APIのレスポンスをJSONとして解析できませんでした: ' + textBlock);
  }

  // 想定外の分類が返ってきた場合は「その他」として扱う
  const validCategories = ['クレーム', '質問', '注文', 'その他'];
  const category = validCategories.indexOf(result.category) !== -1 ? result.category : 'その他';

  return {
    category: category,
    summary: result.summary || ''
  };
}

// Slack Incoming Webhookへ分類結果を通知する
function notifySlack_(webhookUrl, subject, category, summary) {
  const payload = {
    text:
      '新着メール分類通知\n' +
      '件名: ' + subject + '\n' +
      '分類: ' + category + '\n' +
      '要約: ' + summary
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(webhookUrl, options);
  const responseCode = response.getResponseCode();
  if (responseCode !== 200) {
    throw new Error('Slack通知に失敗しました（ステータス' + responseCode + '）: ' + response.getContentText());
  }
}

// エラー内容をスプレッドシートの「エラーログ」シートに記録する
function logError_(errorMessage) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const errorSheet = getOrCreateSheet_(ss, SHEET_ERROR, ['発生日時', 'エラー内容']);
    errorSheet.appendRow([new Date(), errorMessage]);
  } catch (error) {
    // エラーログへの書き込み自体に失敗した場合は実行ログに出力する
    Logger.log('エラーログの書き込みに失敗しました: ' + error);
  }
}

// 指定したラベルを取得し、存在しない場合は新規作成する
function getOrCreateLabel_(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
  }
  return label;
}

// 指定したシートを取得し、存在しない場合はヘッダー行付きで新規作成する
function getOrCreateSheet_(spreadsheet, sheetName, headerRow) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    sheet.appendRow(headerRow);
  }
  return sheet;
}

// 5分おきにprocessUnreadEmailsを自動実行するトリガーを登録する
function setProcessingTrigger() {
  // 重複登録を防ぐため、既存のprocessUnreadEmails用トリガーを一旦削除する
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'processUnreadEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 5分おきにprocessUnreadEmailsを実行するトリガーを新規作成
  ScriptApp.newTrigger('processUnreadEmails')
    .timeBased()
    .everyMinutes(5)
    .create();
}

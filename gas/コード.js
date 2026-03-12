/**
 * ============================================
 * 訪問看護記録システム - 統合版（リハビリ別添対応＋管理画面API）
 * ============================================
 * 
 * 【機能一覧】
 * 1. 音声入力テキストをClaude AIで記録生成
 * 2. 記録をスプレッドシートに保存
 * 3. 従業員リストの取得
 * 4. 訪問看護報告書「病状の経過」自動生成（管理画面から実行）
 *    ★ リハビリ別添の自動生成
 * 5. 毎朝の申し送り自動通知（Google Chat連携）
 * 6. 日別集計
 * 7. ★ 管理画面API（ダッシュボード・記録・利用者・スタッフ・報告書・営業時間分析）
 *    ★ targetMonthパラメータ対応（ダッシュボード・利用者・スタッフ）
 * 
 * ※ 報告書生成はadmin.htmlから実行（スプレッドシートメニューからの生成は廃止）
 */

// ============================================
// 基本設定
// ============================================
var CONFIG = {
  ANTHROPIC_API_KEY: PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY'),
  SPREADSHEET_ID: '1c12YLCD47AsgEhuuU-nxdLNywgCXgMkN7ALFs7LGtK4',
  MODEL: 'claude-haiku-4-5-20251001'
};

var REPORT_CONFIG = {
  MODEL_SONNET: 'claude-sonnet-4-20250514',
  MAX_TOKENS_REPORT: 1024,
  BATCH_SIZE: 30,
  TRIGGER_DELAY_MINUTES: 1,
  TEMP_DATA_SHEET: '_処理用データ'
};

var CHAT_CONFIG = {
  WEBHOOK_URL: 'https://chat.googleapis.com/v1/spaces/AAQAD3oLouM/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=8LfnOXNbLggeE6MHE94mq392fUlvmTK9bUQaDHAc0qQ',
  TRIGGER_HOUR: 8,
  TRIGGER_MINUTE: 40
};

// ============================================
// プロンプト定義
// ============================================

var SYSTEM_PROMPT = 'あなたは訪問看護・リハビリテーションの記録作成を支援するアシスタントです。\n' +
  '看護師やリハビリ職から伝えられた訪問記録の内容を、以下のフォーマットに沿って整理してください。\n\n' +
  '【出力フォーマット】\n' +
  '＜本日の訪問での変化、共有事項＞\n' +
  '・（重要な変化や共有事項を3点以内で記載）\n' +
  '＜補足＞\n' +
  '・（5行以内、250文字以内で補足情報を記載）\n\n' +
  '【ルール】\n' +
  '1. ＜本日の訪問での変化、共有事項＞は前回と比べた大きな変化や重要な点を「3点以内」で記入\n' +
  '2. 1つのトピックが2行になる場合は、トピックを分けて記入\n' +
  '3. ＜補足＞は5行以内（250文字以内）にまとめる\n' +
  '4. iPadで経過記録を確認した際にスクロールせずに見れる範囲が目安\n' +
  '5. 重要な共有事項は5行を超えても可\n' +
  '6. 医療用語は正確に使用し、簡潔で明瞭な文章にする\n' +
  '7. 客観的事実を中心に記載する\n' +
  '8. フォーマット部分のみを出力し、余計な説明は不要\n' +
  '9. アスタリスク(*)や太字記号は使用しない';

var SYSTEM_PROMPT_REHAB = 'あなたは訪問看護・リハビリテーションの記録作成を支援するアシスタントです。\n' +
  'リハビリ職から伝えられた訪問記録の内容を、以下のフォーマットに沿って整理してください。\n\n' +
  '【出力フォーマット】\n' +
  'リハビリ記録\n' +
  '＜本日の訪問での変化、共有事項＞\n' +
  '・（重要な変化や共有事項を3点以内で記載）\n' +
  '＜補足＞\n' +
  '・（5行以内、250文字以内で補足情報を記載）\n\n' +
  '【ルール】\n' +
  '1. 必ず冒頭に「リハビリ記録」と記載する\n' +
  '2. ＜本日の訪問での変化、共有事項＞は前回と比べた大きな変化や重要な点を「3点以内」で記入\n' +
  '3. 1つのトピックが2行になる場合は、トピックを分けて記入\n' +
  '4. ＜補足＞は5行以内（250文字以内）にまとめる\n' +
  '5. iPadで経過記録を確認した際にスクロールせずに見れる範囲が目安\n' +
  '6. 重要な共有事項は5行を超えても可\n' +
  '7. 医療用語は正確に使用し、簡潔で明瞭な文章にする\n' +
  '8. 客観的事実を中心に記載する\n' +
  '9. フォーマット部分のみを出力し、余計な説明は不要\n' +
  '10. アスタリスク(*)や太字記号は使用しない';

var PROGRESS_REPORT_PROMPT = 'あなたは訪問看護の記録作成を支援するアシスタントです。\n' +
  '以下の訪問記録データをもとに、この利用者の「病状の経過」を作成してください。\n\n' +
  '【対象読者】\n' +
  '- 主治医（医療的判断の参考にする）\n' +
  '- ケアマネジャー（ケアプラン作成・サービス調整の参考にする）\n\n' +
  '【記載ルール】\n\n' +
  '1. 文章のわかりやすさ\n' +
  '   - 医療専門用語は必要最小限にし、使用する場合は括弧書きで簡単な説明を添える\n' +
  '   - 略語は初出時に正式名称を併記する（例：SpO2（酸素飽和度））\n' +
  '   - 一文は短めに、主語と述語を明確にする\n\n' +
  '2. 構成\n' +
  '   - 冒頭に全体的な状態（安定・変化あり等）を一文で示す\n' +
  '   - 主要な疾患・健康課題ごとに整理する\n' +
  '   - 観察した事実 → 看護対応 → 結果・反応 の流れで記載する\n' +
  '   - ADL（日常生活動作）や生活面の変化も含める\n\n' +
  '3. 数値・データの記載\n' +
  '   - バイタルサインは範囲で示す（例：血圧120〜140/70〜85mmHg）\n' +
  '   - 変化があった場合は時期を明記する（例：月後半より〜）\n\n' +
  '4. 文字数\n' +
  '   - 300〜500文字\n\n' +
  '5. 出力形式\n' +
  '   - 本文のみを出力（利用者名や見出しは不要）\n' +
  '   - 改行は使わず、1段落の文章で出力\n' +
  '   - 前置きや説明は不要\n' +
  '   - アスタリスク(*)や太字記号は使用しない';

var REHAB_ADDENDUM_PROMPT = 'あなたは訪問看護の記録作成を支援するアシスタントです。\n' +
  '以下のリハビリテーション訪問記録データをもとに、訪問看護報告書の「リハビリテーション別添」を作成してください。\n\n' +
  '【対象読者】\n' +
  '- 主治医（リハビリテーションの進捗状況を正確に把握するため）\n' +
  '- ケアマネジャー（ケアプラン作成・サービス調整の参考にする）\n\n' +
  '【記載項目と構成】\n' +
  '以下の4項目を見出し付きで記載すること。\n\n' +
  '1. 実施したリハビリテーションの内容\n' +
  '   - 具体的な訓練名と内容を記載する\n' +
  '   - 実施頻度や回数が読み取れる場合は含める\n\n' +
  '2. 家族等への指導内容\n' +
  '   - 介助方法の指導内容\n' +
  '   - 環境調整の提案内容\n' +
  '   - 記録に該当する内容がない場合は「特記事項なし」と記載\n\n' +
  '3. リスク管理\n' +
  '   - 転倒予防対策\n' +
  '   - 拘縮予防の取り組み\n' +
  '   - バイタルサインの確認と対応\n' +
  '   - 記録に該当する内容がない場合は「特記事項なし」と記載\n\n' +
  '4. 訓練の進捗状況\n' +
  '   - 利用者の反応や改善状況\n' +
  '   - 実用性の評価\n' +
  '   - 今後の課題\n\n' +
  '【記載ルール】\n' +
  '1. 具体的かつ簡潔に記載する\n' +
  '2. 各項目は箇条書き1〜2行で簡潔にまとめる\n' +
  '3. 全体の文字数は300〜350文字に収める（厳守）\n' +
  '4. 語尾は常体（だ・である調）で統一する\n\n' +
  '【出力形式】\n' +
  '以下の形式で出力すること。前置きや説明は不要。アスタリスク(*)や太字記号は使用しない。\n' +
  '項目間に空行は入れず、詰めて記載すること。\n\n' +
  '＜実施したリハビリテーションの内容＞\n' +
  '・（内容）\n' +
  '＜家族等への指導内容＞\n' +
  '・（内容）\n' +
  '＜リスク管理＞\n' +
  '・（内容）\n' +
  '＜訓練の進捗状況＞\n' +
  '・（内容）';

var PICKUP_PROMPT = 'あなたは訪問看護ステーションの管理者を支援するアシスタントです。\n' +
  '以下の前日の訪問記録データを確認し、翌日の朝に全スタッフへ共有すべき重要な情報をピックアップしてください。\n\n' +
  '【ピックアップ基準】（該当するものをすべて抽出）\n' +
  '1. バイタルサインの異常値（血圧160以上または90以下、発熱37.5度以上、SpO2 93%以下、脈拍100以上または50以下など）\n' +
  '2. 状態の悪化（疼痛増強、浮腫増強、食欲低下、ADL低下、意識レベル変化など）\n' +
  '3. 転倒・転落などのインシデント\n' +
  '4. 主治医への報告・連絡が必要な内容、または主治医に報告した内容\n' +
  '5. 新しい症状の出現\n' +
  '6. 服薬に関する問題（飲み忘れ、副作用疑い、残薬調整など）\n' +
  '7. 臨時訪問（予定外の訪問）があった場合の内容\n' +
  '8. 特別訪問看護指示書による訪問\n\n' +
  '【出力ルール】\n' +
  '1. 該当する利用者がいない場合は「該当なし」とだけ出力する\n' +
  '2. 該当する利用者がいる場合、以下のフォーマットで出力する\n' +
  '3. 利用者ごとにまとめ、簡潔に記載する\n' +
  '4. 緊急度が高い順に並べる\n' +
  '5. アスタリスク(*)や太字記号は使用しない\n' +
  '6. 前置きや説明は不要、フォーマット部分のみ出力\n\n' +
  '【出力フォーマット】（該当者がいる場合）\n' +
  '■ [利用者名]（[職種]：[記録者]）\n' +
  '[ピックアップ理由を1〜2行で簡潔に記載]\n' +
  '→ [推奨される対応やフォロー事項]';


// ============================================
// メニュー（朝の申し送り・日別集計のみ）
// ============================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('朝の申し送り')
    .addItem('今すぐ送信（テスト）', 'sendMorningReportNow')
    .addItem('自動配信を設定（毎朝8:40）', 'setupDailyTrigger')
    .addItem('自動配信を停止', 'removeDailyTrigger')
    .addSeparator()
    .addItem('配信状況を確認', 'checkTriggerStatus')
    .addToUi();

  ui.createMenu('📊 集計')
    .addItem('日別集計を今すぐ更新', 'updateDailySummary')
    .addSeparator()
    .addItem('自動集計を設定（8:00/12:00/18:00）', 'setupSummaryTriggers')
    .addItem('自動集計を停止', 'removeSummaryTriggers')
    .addItem('自動集計の状況を確認', 'checkSummaryTriggerStatus')
    .addToUi();
}


// ============================================
// ウェブアプリ（POST/GET）
// ============================================

function doPost(e) {
  var result;
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action || 'save';
    
    if (action === 'generate') {
      result = generateRecordInternal(data);
    } else if (action === 'save') {
      result = saveRecordInternal(data);
    } else if (action === 'generate_and_save') {
      result = generateAndSaveInternal(data);
    } else if (action === 'getStaff') {
      result = getStaffList();
    } else if (action === 'admin_dashboard') {
      result = getAdminDashboard(data);
    } else if (action === 'admin_records') {
      result = getAdminRecords(data);
    } else if (action === 'admin_users') {
      result = getAdminUsers(data);
    } else if (action === 'admin_staff') {
      result = getAdminStaff(data);
    } else if (action === 'admin_reports') {
      result = getAdminReports(data);
    } else if (action === 'admin_record_detail') {
      result = getAdminRecordDetail(data);
    } else if (action === 'admin_business_hours') {
      result = getAdminBusinessHours(data);
    } else if (action === 'admin_generate_reports') {
      result = adminGenerateReports(data);
    } else if (action === 'admin_generate_status') {
      result = adminGenerateStatus(data);
    } else if (action === 'admin_update_report') {
      result = adminUpdateReport(data);
    } else if (action === 'getPatients') {
      result = getPatientsList();
    } else if (action === 'addPatient') {
      result = addPatient(data);
    } else if (action === 'importPatients') {
      result = importPatientsFromCsv(data);
    } else if (action === 'deletePatient') {
      result = deletePatient(data);
    } else if (action === 'getUsers') {
      result = getUserList();
    } else if (action === 'staff_login') {
      result = staffLogin(data.staffName, data.password);
    } else if (action === 'updateStaffPassword') {
      result = updateStaffPassword(data.staffName, data.password);
    } else if (action === 'addStaff') {
      result = addStaff(data.name, data.jobType, data.password);
    } else {
      result = { success: false, message: 'Unknown action' };
    }
  } catch (error) {
    result = { success: false, message: error.toString() };
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var result = { success: true, message: 'API is running' };
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}


// ============================================
// 利用者リスト取得（記録入力画面のサジェスト用）
// ============================================

/**
 * 利用者リストを返す（利用者シートがあればマスタから、なければ記録シートから）
 */
function getUserList() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

    // 利用者シート（マスタ）があればそこから取得
    var patientSheet = ss.getSheetByName('利用者');
    if (patientSheet && patientSheet.getLastRow() >= 2) {
      var patientData = patientSheet.getRange(2, 1, patientSheet.getLastRow() - 1, 2).getValues();
      var users = [];
      for (var p = 0; p < patientData.length; p++) {
        var kana = String(patientData[p][0]).trim();
        if (kana) users.push(kana);
      }
      users.sort();
      return { success: true, data: { users: users } };
    }

    // マスタがなければ記録シートから取得（従来動作）
    var sheet = ss.getSheetByName('記録');
    if (!sheet) return { success: true, data: { users: [] } };
    var allData = sheet.getDataRange().getValues();
    if (allData.length <= 1) return { success: true, data: { users: [] } };
    var headers = allData[0], patientIdx = headers.indexOf('利用者名');
    var nameSet = {};
    for (var i = 1; i < allData.length; i++) {
      var name = patientIdx >= 0 ? String(allData[i][patientIdx]).trim() : '';
      if (name) nameSet[name] = true;
    }
    return { success: true, data: { users: Object.keys(nameSet).sort() } };
  } catch (error) {
    return { success: false, message: '利用者リスト取得エラー: ' + error.toString() };
  }
}

// ============================================
// スタッフログイン
// ============================================

function staffLogin(staffName, password) {
  if (!staffName || !password) {
    return { success: false, message: 'スタッフ名とパスワードを入力してください' };
  }
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('従業員');
  if (!sheet) return { success: false, message: '従業員シートが見つかりません' };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var name = String(data[i][0]).trim();
    var jobType = String(data[i][1]).trim();
    var pw = String(data[i][2]).trim();
    if (name === staffName && pw === password) {
      return { success: true, data: { role: 'staff', staffName: name, jobType: jobType } };
    }
  }
  return { success: false, message: 'スタッフ名またはパスワードが正しくありません' };
}

// ============================================
// スタッフパスワード更新
// ============================================

function updateStaffPassword(staffName, password) {
  if (!staffName || !password) {
    return { success: false, message: 'スタッフ名とパスワードを入力してください' };
  }
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('従業員');
  if (!sheet) return { success: false, message: '従業員シートが見つかりません' };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === staffName) {
      sheet.getRange(i + 1, 3).setValue(password);
      return { success: true, data: { message: 'パスワードを更新しました' } };
    }
  }
  return { success: false, message: 'スタッフが見つかりません' };
}

// ============================================
// 従業員リスト取得
// ============================================

function getStaffList() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('従業員');
    
    if (!sheet) {
      return { success: false, message: '従業員シートが見つかりません' };
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, data: { staff: [] } };
    }
    
    var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    var staff = data
      .filter(function(row) { return row[0] && row[1]; })
      .map(function(row) {
        return {
          name: row[0].toString().trim(),
          jobType: row[1].toString().trim()
        };
      });
    
    return { success: true, data: { staff: staff } };
  } catch (error) {
    return { success: false, message: '従業員リスト取得エラー: ' + error.toString() };
  }
}


// ============================================
// スタッフ追加（パスワード列対応）
// ============================================

function addStaff(name, jobType, password) {
  try {
    if (!name || !jobType) {
      return { success: false, message: '名前と職種を入力してください' };
    }
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('従業員');
    if (!sheet) return { success: false, message: '従業員シートが見つかりません' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === name) {
        return { success: false, message: name + 'は既に登録されています' };
      }
    }

    sheet.appendRow([name, jobType, password || '']);
    return { success: true, data: { message: name + 'を追加しました' } };
  } catch (error) {
    return { success: false, message: 'スタッフ追加エラー: ' + error.toString() };
  }
}

// ============================================
// AI記録生成
// ============================================

function generateRecordInternal(data) {
  var transcript = data.transcript || '';
  var jobType = data.jobType || '';
  
  if (!transcript.trim()) {
    return { success: false, message: '音声入力の内容がありません' };
  }
  
  var prompt = (jobType === 'リハビリ職') ? SYSTEM_PROMPT_REHAB : SYSTEM_PROMPT;
  
  try {
    var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': CONFIG.ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify({
        model: CONFIG.MODEL,
        max_tokens: 1000,
        system: prompt,
        messages: [{
          role: 'user',
          content: '以下の訪問記録の音声入力内容を、指定されたフォーマットに整理してください。\n\n【音声入力内容】\n' + transcript
        }]
      }),
      muteHttpExceptions: true
    });
    
    var result = JSON.parse(response.getContentText());
    
    if (result.content && result.content[0]) {
      var record = result.content[0].text;
      record = record.replace(/\*\*/g, '').replace(/\*/g, '');
      return { success: true, message: 'Generated', data: { record: record } };
    } else if (result.error) {
      return { success: false, message: 'API Error: ' + result.error.message };
    } else {
      return { success: false, message: '記録の生成に失敗しました' };
    }
  } catch (error) {
    return { success: false, message: 'AI生成エラー: ' + error.toString() };
  }
}


// ============================================
// スプレッドシート保存
// ============================================

function saveRecordInternal(data) {
  var visitDate = data.date || '';
  var patientName = data.patientName || '';
  var jobType = data.jobType || '';
  var recorderName = data.recorderName || '';
  var record = data.record || '';
  
  if (!visitDate || !patientName || !record) {
    return { success: false, message: '必要な情報が不足しています' };
  }
  
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('記録');
    
    if (!sheet) {
      sheet = ss.getSheets()[0];
    }
    
    if (sheet.getLastRow() === 0) {
      setupSheetHeader(sheet);
    }
    
    var now = new Date();
    var timestamp = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    var formattedDate = formatVisitDate(visitDate);
    
    sheet.appendRow([formattedDate, patientName, jobType, recorderName, record, timestamp]);
    
    return { success: true, message: '保存しました', data: { timestamp: timestamp } };
  } catch (error) {
    return { success: false, message: '保存エラー: ' + error.toString() };
  }
}

function generateAndSaveInternal(data) {
  var generateResult = generateRecordInternal(data);
  if (!generateResult.success) return generateResult;
  
  data.record = generateResult.data.record;
  var saveResult = saveRecordInternal(data);
  if (!saveResult.success) return saveResult;
  
  return {
    success: true,
    message: '生成して保存しました',
    data: {
      record: generateResult.data.record,
      timestamp: saveResult.data.timestamp
    }
  };
}

function formatVisitDate(dateStr) {
  if (!dateStr) return '';
  try {
    var date = new Date(dateStr);
    return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd');
  } catch (e) {
    return dateStr;
  }
}


// ============================================
// シートセットアップ
// ============================================

function setupSheetHeader(sheet) {
  var headers = ['訪問日', '利用者名', '職種', '記録者', '記録内容', '登録日時'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#00a67d');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 500);
  sheet.setColumnWidth(6, 150);
  sheet.setFrozenRows(1);
}

function setupStaffSheet() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('従業員');
  if (!sheet) {
    sheet = ss.insertSheet('従業員');
  }
  var headers = ['名前', '職種'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#00a67d');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 100);
  sheet.setFrozenRows(1);
}

function initialSetup() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('記録');
  if (!sheet) {
    sheet = ss.insertSheet('記録');
  }
  setupSheetHeader(sheet);
  setupStaffSheet();
  Logger.log('セットアップ完了');
}


// ============================================
// 一時データ管理
// ============================================

function saveTempData(ss, userNames, userRecords, headers) {
  var existing = ss.getSheetByName(REPORT_CONFIG.TEMP_DATA_SHEET);
  if (existing) ss.deleteSheet(existing);
  var tempSheet = ss.insertSheet(REPORT_CONFIG.TEMP_DATA_SHEET);
  tempSheet.getRange(1, 1).setValue(JSON.stringify(headers));
  var rows = [];
  for (var i = 0; i < userNames.length; i++) {
    rows.push([userNames[i], JSON.stringify(userRecords[userNames[i]])]);
  }
  if (rows.length > 0) tempSheet.getRange(2, 1, rows.length, 2).setValues(rows);
  tempSheet.hideSheet();
}

function loadTempData(ss) {
  var tempSheet = ss.getSheetByName(REPORT_CONFIG.TEMP_DATA_SHEET);
  if (!tempSheet) return null;
  var headers = JSON.parse(tempSheet.getRange(1, 1).getValue());
  var lastRow = tempSheet.getLastRow();
  if (lastRow < 2) return null;
  var data = tempSheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var userNames = [];
  var userRecords = {};
  for (var i = 0; i < data.length; i++) {
    userNames.push(data[i][0]);
    userRecords[data[i][0]] = JSON.parse(data[i][1]);
  }
  return { headers: headers, userNames: userNames, userRecords: userRecords };
}

function cleanupTempData(ss) {
  var tempSheet = ss.getSheetByName(REPORT_CONFIG.TEMP_DATA_SHEET);
  if (tempSheet) ss.deleteSheet(tempSheet);
}


// ============================================
// バッチ処理（リハビリ別添対応）
// ============================================

function processBatchV2() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var props = PropertiesService.getScriptProperties();
  var currentIndex = parseInt(props.getProperty('BATCH_CURRENT_INDEX') || '0');
  var totalCount = parseInt(props.getProperty('BATCH_TOTAL_COUNT') || '0');
  var outputSheetName = props.getProperty('BATCH_OUTPUT_SHEET') || '';
  
  if (!outputSheetName || totalCount === 0) { cleanupAll(ss); return; }
  if (currentIndex >= totalCount) { finishProcessingV2(ss); return; }
  
  var tempData = loadTempData(ss);
  if (!tempData) { cleanupAll(ss); return; }
  
  var endIndex = Math.min(currentIndex + REPORT_CONFIG.BATCH_SIZE, totalCount);
  var batchUserNames = tempData.userNames.slice(currentIndex, endIndex);
  var outputSheet = ss.getSheetByName(outputSheetName);
  if (!outputSheet) { cleanupAll(ss); return; }
  
  var generatedAt = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  var jobTypeIndex = tempData.headers.indexOf('職種');
  var results = [];
  
  for (var i = 0; i < batchUserNames.length; i++) {
    var userName = batchUserNames[i];
    try {
      var allRecords = tempData.userRecords[userName];
      var visitCount = allRecords.length;
      var periodStr = getPeriodString(allRecords, tempData.headers);
      
      var nursingRecords = [];
      var rehabRecords = [];
      for (var j = 0; j < allRecords.length; j++) {
        var jt = jobTypeIndex >= 0 ? String(allRecords[j][jobTypeIndex]).trim() : '';
        if (jt === 'リハビリ職') { rehabRecords.push(allRecords[j]); }
        else { nursingRecords.push(allRecords[j]); }
      }
      
      var progressText = '';
      if (nursingRecords.length > 0) {
        progressText = callClaudeForProgressReport(formatUserRecordsForReport(nursingRecords, tempData.headers));
        Utilities.sleep(1200);
      }
      
      var rehabAddendum = '';
      if (rehabRecords.length > 0) {
        rehabAddendum = callClaudeForRehabAddendum(formatUserRecordsForReport(rehabRecords, tempData.headers));
        Utilities.sleep(1200);
      }
      
      results.push([userName, visitCount, periodStr, progressText, rehabAddendum, generatedAt]);
    } catch (error) {
      results.push([userName, tempData.userRecords[userName].length, '', 'エラー: ' + error.message, '', generatedAt]);
    }
  }
  
  if (results.length > 0) {
    var lastRow = outputSheet.getLastRow();
    outputSheet.getRange(lastRow + 1, 1, results.length, 6).setValues(results);
    outputSheet.getRange(lastRow + 1, 4, results.length, 1).setWrap(true);
    outputSheet.getRange(lastRow + 1, 5, results.length, 1).setWrap(true);
  }
  
  props.setProperty('BATCH_CURRENT_INDEX', String(endIndex));
  
  if (endIndex < totalCount) {
    ScriptApp.newTrigger('processBatchV2').timeBased().after(REPORT_CONFIG.TRIGGER_DELAY_MINUTES * 60 * 1000).create();
  } else {
    finishProcessingV2(ss);
  }
}

function callClaudeForRehabAddendum(rehabDataText) {
  var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'x-api-key': CONFIG.ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: REPORT_CONFIG.MODEL_SONNET, max_tokens: REPORT_CONFIG.MAX_TOKENS_REPORT,
      system: REHAB_ADDENDUM_PROMPT,
      messages: [{ role: 'user', content: '以下のリハビリテーション訪問記録データから「リハビリテーション別添」を作成してください。\n\n【リハビリ訪問記録データ】\n' + rehabDataText }]
    }),
    muteHttpExceptions: true
  });
  var responseCode = response.getResponseCode();
  var result = JSON.parse(response.getContentText());
  if (responseCode !== 200) throw new Error('API Error (' + responseCode + ')');
  if (result.content && result.content[0]) { return result.content[0].text.replace(/\*\*/g, '').replace(/\*/g, '').trim(); }
  throw new Error('リハビリ別添の生成に失敗しました');
}

function finishProcessingV2(ss) {
  var props = PropertiesService.getScriptProperties();
  var outputSheetName = props.getProperty('BATCH_OUTPUT_SHEET') || '';
  var startTimeStr = props.getProperty('BATCH_START_TIME') || '';
  var totalCount = props.getProperty('BATCH_TOTAL_COUNT') || '0';
  var periodText = props.getProperty('BATCH_PERIOD_TEXT') || '';
  
  cleanupTriggers(); cleanupTempData(ss); clearBatchProperties();
  
  var outputSheet = ss.getSheetByName(outputSheetName);
  if (outputSheet) {
    var lastRow = outputSheet.getLastRow();
    for (var i = 2; i <= lastRow; i++) outputSheet.setRowHeight(i, 80);
  }
  
  var startTime = startTimeStr ? new Date(startTimeStr) : new Date();
  var durationMinutes = Math.round((new Date() - startTime) / 1000 / 60);
  
  try {
    var email = Session.getActiveUser().getEmail();
    if (email) {
      MailApp.sendEmail({
        to: email,
        subject: '【完了】病状の経過 生成処理（' + periodText + '）',
        body: '処理件数：' + totalCount + '名\n処理時間：約' + durationMinutes + '分\n対象期間：' + periodText + '\n出力シート：' + outputSheetName + '\n\nスプレッドシート：' + ss.getUrl()
      });
    }
  } catch (e) { Logger.log('メール送信エラー: ' + e.toString()); }
}


// ============================================
// 処理管理
// ============================================

function cancelProcessingInternal() { var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID); cleanupTriggers(); cleanupTempData(ss); clearBatchProperties(); }
function cleanupAll(ss) { cleanupTriggers(); cleanupTempData(ss); clearBatchProperties(); }

function cleanupTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === 'processBatchV2' || fn === 'processBatch') ScriptApp.deleteTrigger(triggers[i]);
  }
}

function clearBatchProperties() {
  var props = PropertiesService.getScriptProperties();
  ['BATCH_CURRENT_INDEX','BATCH_TOTAL_COUNT','BATCH_START_TIME','BATCH_PERIOD_TEXT','BATCH_OUTPUT_SHEET','PROCESSING_FLAG','PROGRESS_INFO','PROGRESS_USER_RECORDS'].forEach(function(key) { props.deleteProperty(key); });
}

function isProcessing() { return PropertiesService.getScriptProperties().getProperty('PROCESSING_FLAG') === 'true'; }


// ============================================
// 朝の申し送り
// ============================================

function sendMorningReport() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('記録');
    if (!sheet) { Logger.log('記録シートが見つかりません'); return; }
    
    var yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    var yesterdayDisplay = Utilities.formatDate(yesterday, 'Asia/Tokyo', 'M月d日（E）');
    var records = getRecordsByDate(sheet, yesterday);
    
    if (records.length === 0) {
      sendToGoogleChat('📋 *朝の申し送り*（' + yesterdayDisplay + '分）\n\n前日の訪問記録はありませんでした。');
      return;
    }
    
    var recordsText = formatRecordsForPickup(records);
    var pickupResult = callClaudeForPickup(recordsText);
    sendToGoogleChat(buildChatMessage(yesterdayDisplay, records.length, pickupResult));
  } catch (error) {
    Logger.log('朝の申し送りエラー: ' + error.toString());
    try { sendToGoogleChat('⚠️ *朝の申し送り - エラー*\n\nエラー: ' + error.message); } catch (e) {}
  }
}

function sendMorningReportNow() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('テスト送信', '前日分の申し送りを送信します。', ui.ButtonSet.YES_NO) === ui.Button.YES) {
    sendMorningReport();
    ui.alert('送信完了', 'Google Chatに送信しました。', ui.ButtonSet.OK);
  }
}

function getRecordsByDate(sheet, targetDate) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  var dateIndex = headers.indexOf('訪問日');
  var patientIndex = headers.indexOf('利用者名');
  var jobTypeIndex = headers.indexOf('職種');
  var recorderIndex = headers.indexOf('記録者');
  var recordIndex = headers.indexOf('記録内容');
  if (dateIndex === -1 || recordIndex === -1) return [];
  
  var targetDateStr = Utilities.formatDate(targetDate, 'Asia/Tokyo', 'yyyy/MM/dd');
  var records = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowDate = '';
    if (row[dateIndex] instanceof Date) rowDate = Utilities.formatDate(row[dateIndex], 'Asia/Tokyo', 'yyyy/MM/dd');
    else if (typeof row[dateIndex] === 'string') rowDate = row[dateIndex].trim();
    
    if (rowDate === targetDateStr) {
      records.push({
        date: rowDate,
        patientName: patientIndex >= 0 ? String(row[patientIndex]).trim() : '',
        jobType: jobTypeIndex >= 0 ? String(row[jobTypeIndex]).trim() : '',
        recorder: recorderIndex >= 0 ? String(row[recorderIndex]).trim() : '',
        record: recordIndex >= 0 ? String(row[recordIndex]).trim() : ''
      });
    }
  }
  return records;
}

function formatRecordsForPickup(records) {
  var lines = ['訪問日\t利用者名\t職種\t記録者\t記録内容'];
  for (var i = 0; i < records.length; i++) {
    var r = records[i];
    lines.push(r.date + '\t' + r.patientName + '\t' + r.jobType + '\t' + r.recorder + '\t' + r.record.replace(/\n/g, ' ').replace(/\t/g, ' '));
  }
  return lines.join('\n');
}

function callClaudeForPickup(recordsText) {
  var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'x-api-key': CONFIG.ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: 'claude-haiku-4-5-20251001', max_tokens: 1500, system: PICKUP_PROMPT,
      messages: [{ role: 'user', content: '以下の前日の訪問記録から、全スタッフに共有すべき重要な情報をピックアップしてください。\n\n【前日の訪問記録】\n' + recordsText }]
    }),
    muteHttpExceptions: true
  });
  var result = JSON.parse(response.getContentText());
  if (response.getResponseCode() !== 200) throw new Error('API Error');
  if (result.content && result.content[0]) return result.content[0].text.replace(/\*\*/g, '').replace(/\*/g, '').trim();
  throw new Error('ピックアップ生成に失敗しました');
}

function buildChatMessage(dateDisplay, recordCount, pickupResult) {
  var message = '📋 *朝の申し送り*（' + dateDisplay + '分）\n━━━━━━━━━━━━━━━━━━\n前日の訪問件数：' + recordCount + '件\n\n';
  if (pickupResult === '該当なし') message += '✅ 特に共有が必要な変化はありませんでした。\n';
  else message += '🔔 *共有事項*\n\n' + pickupResult + '\n';
  message += '\n━━━━━━━━━━━━━━━━━━\n※ AIによる自動ピックアップです。詳細はスプレッドシートの記録をご確認ください。';
  return message;
}

function sendToGoogleChat(message) {
  var response = UrlFetchApp.fetch(CHAT_CONFIG.WEBHOOK_URL, {
    method: 'POST', contentType: 'application/json; charset=UTF-8',
    payload: JSON.stringify({ text: message }), muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) throw new Error('Google Chat送信エラー');
  return true;
}


// ============================================
// トリガー管理
// ============================================

function setupDailyTrigger() {
  removeMorningTriggers();
  ScriptApp.newTrigger('sendMorningReport').timeBased().atHour(CHAT_CONFIG.TRIGGER_HOUR).nearMinute(CHAT_CONFIG.TRIGGER_MINUTE).everyDays(1).create();
  SpreadsheetApp.getUi().alert('設定完了', '毎朝' + CHAT_CONFIG.TRIGGER_HOUR + ':' + CHAT_CONFIG.TRIGGER_MINUTE + 'に自動配信されます。', SpreadsheetApp.getUi().ButtonSet.OK);
}

function removeDailyTrigger() { removeMorningTriggers(); SpreadsheetApp.getUi().alert('停止完了', '自動配信を停止しました。', SpreadsheetApp.getUi().ButtonSet.OK); }

function removeMorningTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) { if (triggers[i].getHandlerFunction() === 'sendMorningReport') ScriptApp.deleteTrigger(triggers[i]); }
}

function checkTriggerStatus() {
  var ui = SpreadsheetApp.getUi();
  var triggers = ScriptApp.getProjectTriggers();
  var found = false;
  for (var i = 0; i < triggers.length; i++) { if (triggers[i].getHandlerFunction() === 'sendMorningReport') { found = true; break; } }
  if (found) ui.alert('配信状況', '✅ 自動配信：有効\n⏰ 毎朝' + CHAT_CONFIG.TRIGGER_HOUR + ':' + String(CHAT_CONFIG.TRIGGER_MINUTE).padStart(2, '0'), ui.ButtonSet.OK);
  else ui.alert('配信状況', '❌ 自動配信：無効', ui.ButtonSet.OK);
}


// ============================================
// 共通ヘルパー関数
// ============================================

function parseDate(dateValue) {
  if (dateValue instanceof Date) return dateValue;
  if (typeof dateValue === 'string') { var date = new Date(dateValue); return isNaN(date.getTime()) ? null : date; }
  return null;
}

function groupRecordsByUser(records, headers) {
  var userIndex = headers.indexOf('利用者名');
  if (userIndex === -1) throw new Error('「利用者名」列が見つかりません');
  var grouped = {};
  for (var i = 0; i < records.length; i++) {
    var userName = records[i][userIndex];
    if (!userName) continue;
    var userNameStr = userName.toString().trim();
    if (!userNameStr) continue;
    if (!grouped[userNameStr]) grouped[userNameStr] = [];
    grouped[userNameStr].push(records[i]);
  }
  return grouped;
}

function formatUserRecordsForReport(records, headers) {
  var text = headers.join('\t') + '\n';
  for (var i = 0; i < records.length; i++) {
    var formattedRecord = records[i].map(function(cell) {
      if (cell instanceof Date) return Utilities.formatDate(cell, 'Asia/Tokyo', 'yyyy/MM/dd');
      return String(cell).replace(/\n/g, ' ').replace(/\t/g, ' ');
    });
    text += formattedRecord.join('\t') + '\n';
  }
  return text;
}

function getPeriodString(records, headers) {
  var dateIndex = headers.indexOf('訪問日');
  if (dateIndex === -1) return '';
  var dates = records.map(function(r) { return parseDate(r[dateIndex]); }).filter(function(d) { return d !== null; }).sort(function(a, b) { return a - b; });
  if (dates.length === 0) return '';
  return Utilities.formatDate(dates[0], 'Asia/Tokyo', 'yyyy/MM/dd') + '〜' + Utilities.formatDate(dates[dates.length - 1], 'Asia/Tokyo', 'yyyy/MM/dd');
}

function callClaudeForProgressReport(userDataText) {
  var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'x-api-key': CONFIG.ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: REPORT_CONFIG.MODEL_SONNET, max_tokens: REPORT_CONFIG.MAX_TOKENS_REPORT, system: PROGRESS_REPORT_PROMPT,
      messages: [{ role: 'user', content: '以下の訪問記録データから「病状の経過」を作成してください。\n\n【訪問記録データ】\n' + userDataText }]
    }),
    muteHttpExceptions: true
  });
  var result = JSON.parse(response.getContentText());
  if (response.getResponseCode() !== 200) throw new Error('API Error');
  if (result.content && result.content[0]) return result.content[0].text.replace(/\*\*/g, '').replace(/\*/g, '').trim();
  throw new Error('記録の生成に失敗しました');
}


// ============================================
// 共通ヘルパー：targetMonth正規化
// ============================================

/**
 * targetMonth文字列を "yyyy/MM" 形式に正規化する
 * 入力: "2026/01", "2026-01", "2026/1" など
 * 出力: "2026/01"
 * 空文字の場合は現在月を返す
 */
function normalizeTargetMonth(targetMonth) {
  if (!targetMonth) {
    var now = new Date();
    return Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM');
  }
  var m = String(targetMonth).match(/(\d{4})[\/\-](\d{1,2})/);
  if (!m) {
    var now2 = new Date();
    return Utilities.formatDate(now2, 'Asia/Tokyo', 'yyyy/MM');
  }
  return m[1] + '/' + String(parseInt(m[2])).padStart(2, '0');
}

/**
 * 日付文字列 "yyyy/MM/dd" または "yyyy-MM-dd" から "yyyy/MM" を抽出
 */
function extractMonth(dateStr) {
  if (!dateStr) return '';
  var s = String(dateStr).trim();
  s = s.replace(/-/g, '/');
  return s.substring(0, 7);
}

/**
 * 前月の "yyyy/MM" を返す
 */
function getPrevMonth(targetMonth) {
  var m = targetMonth.match(/(\d{4})\/(\d{2})/);
  if (!m) return '';
  var y = parseInt(m[1]), mo = parseInt(m[2]);
  if (mo === 1) { y--; mo = 12; }
  else { mo--; }
  return y + '/' + String(mo).padStart(2, '0');
}


// ============================================
// 日別集計
// ============================================

function updateDailySummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName('記録');
  if (!sourceSheet) { SpreadsheetApp.getUi().alert('「記録」シートが見つかりません。'); return; }
  
  var data = sourceSheet.getDataRange().getValues();
  if (data.length <= 1) { SpreadsheetApp.getUi().alert('データがありません。'); return; }
  
  var headers = data[0];
  var dateIndex = headers.indexOf('訪問日');
  var jobTypeIndex = headers.indexOf('職種');
  if (dateIndex === -1) { SpreadsheetApp.getUi().alert('「訪問日」列が見つかりません。'); return; }
  
  var dailyCount = {};
  var staffTypes = {};
  
  for (var i = 1; i < data.length; i++) {
    var rawDate = data[i][dateIndex];
    var jobType = jobTypeIndex >= 0 ? String(data[i][jobTypeIndex]).trim() : '';
    if (!rawDate) continue;
    var dateStr;
    if (rawDate instanceof Date) dateStr = Utilities.formatDate(rawDate, 'Asia/Tokyo', 'yyyy/MM/dd');
    else dateStr = String(rawDate).trim().substring(0, 10);
    if (!dateStr) continue;
    if (!dailyCount[dateStr]) dailyCount[dateStr] = { total: 0 };
    dailyCount[dateStr].total++;
    if (jobType) { staffTypes[jobType] = true; dailyCount[dateStr][jobType] = (dailyCount[dateStr][jobType] || 0) + 1; }
  }
  
  var sortedDates = Object.keys(dailyCount).sort();
  var typeList = Object.keys(staffTypes).sort();
  
  var summarySheet = ss.getSheetByName('日別集計');
  if (!summarySheet) summarySheet = ss.insertSheet('日別集計');
  else { summarySheet.clearContents(); summarySheet.clearFormats(); }
  
  var outputHeaders = ['訪問日', '合計件数'].concat(typeList);
  summarySheet.getRange(1, 1, 1, outputHeaders.length).setValues([outputHeaders]);
  summarySheet.getRange(1, 1, 1, outputHeaders.length).setBackground('#00a67d').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');
  
  var outputRows = [];
  for (var d = 0; d < sortedDates.length; d++) {
    var rec = dailyCount[sortedDates[d]];
    var row = [sortedDates[d], rec.total];
    for (var t = 0; t < typeList.length; t++) row.push(rec[typeList[t]] || 0);
    outputRows.push(row);
  }
  
  if (outputRows.length > 0) {
    summarySheet.getRange(2, 1, outputRows.length, outputHeaders.length).setValues(outputRows);
    summarySheet.getRange(2, 2, outputRows.length, outputHeaders.length - 1).setHorizontalAlignment('center');
  }
  
  var totalRowIdx = outputRows.length + 2;
  var totalRow = ['合計'];
  for (var c = 2; c <= outputHeaders.length; c++) {
    var col = columnNumberToLetter(c);
    totalRow.push('=SUM(' + col + '2:' + col + (outputRows.length + 1) + ')');
  }
  summarySheet.getRange(totalRowIdx, 1, 1, outputHeaders.length).setValues([totalRow]);
  summarySheet.getRange(totalRowIdx, 1, 1, outputHeaders.length).setBackground('#e8f0fe').setFontWeight('bold').setHorizontalAlignment('center');
  summarySheet.autoResizeColumns(1, outputHeaders.length);
  summarySheet.getRange(totalRowIdx + 2, 1).setValue('最終更新：' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')).setFontColor('#999999').setFontSize(9);
  
  ss.setActiveSheet(summarySheet);
  SpreadsheetApp.getUi().alert('集計完了！\n' + sortedDates.length + '日分 / 合計' + outputRows.reduce(function(s, r) { return s + r[1]; }, 0) + '件');
}

function columnNumberToLetter(col) {
  var letter = '';
  while (col > 0) { var rem = (col - 1) % 26; letter = String.fromCharCode(65 + rem) + letter; col = Math.floor((col - 1) / 26); }
  return letter;
}

function setupSummaryTriggers() {
  removeSummaryTriggers();
  var hours = [8, 12, 18];
  for (var i = 0; i < hours.length; i++) {
    ScriptApp.newTrigger('updateDailySummary').timeBased().atHour(hours[i]).nearMinute(0).everyDays(1).create();
  }
  SpreadsheetApp.getUi().alert('トリガー設定完了', '毎日 8:00 / 12:00 / 18:00 に自動更新されます。', SpreadsheetApp.getUi().ButtonSet.OK);
}

function removeSummaryTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) { if (triggers[i].getHandlerFunction() === 'updateDailySummary') ScriptApp.deleteTrigger(triggers[i]); }
}

function checkSummaryTriggerStatus() {
  var triggers = ScriptApp.getProjectTriggers();
  var count = 0;
  for (var i = 0; i < triggers.length; i++) { if (triggers[i].getHandlerFunction() === 'updateDailySummary') count++; }
  var ui = SpreadsheetApp.getUi();
  if (count > 0) ui.alert('集計トリガー状況', '✅ 自動集計：有効（' + count + '件）\n⏰ 毎日 8:00 / 12:00 / 18:00', ui.ButtonSet.OK);
  else ui.alert('集計トリガー状況', '❌ 自動集計：無効', ui.ButtonSet.OK);
}


// ============================================
// テスト用関数
// ============================================

function testGenerate() {
  var testData = { postData: { contents: JSON.stringify({ action: 'generate', transcript: 'バイタル測定しました。血圧120の80、体温36.5度、脈拍72回で安定しています。', jobType: '看護師' }) } };
  Logger.log(doPost(testData).getContent());
}

function testGetStaff() {
  var testData = { postData: { contents: JSON.stringify({ action: 'getStaff' }) } };
  Logger.log(doPost(testData).getContent());
}

function testChatWebhook() {
  sendToGoogleChat('🔧 テストメッセージ\n\n訪問看護記録アプリからの通知テストです。');
}


// ============================================
// 管理画面API（targetMonth対応版）
// ============================================

/**
 * ダッシュボード統計データ
 * data.targetMonth: "2026/01" 等 → その月の統計を返す（未指定は今月）
 */
function getAdminDashboard(data) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var recordSheet = ss.getSheetByName('記録');
    var staffSheet = ss.getSheetByName('従業員');
    
    if (!recordSheet) return { success: false, message: '記録シートが見つかりません' };
    
    var allData = recordSheet.getDataRange().getValues();
    if (allData.length <= 1) {
      return { success: true, data: { monthlyVisits: 0, todayVisits: 0, activeUsers: 0, staffCount: 0, staffBreakdown: '', recentRecords: [], monthlyChange: 0, targetMonth: '' }};
    }
    
    var headers = allData[0];
    var records = allData.slice(1);
    var dateIdx = headers.indexOf('訪問日');
    var patientIdx = headers.indexOf('利用者名');
    var jobIdx = headers.indexOf('職種');
    var recorderIdx = headers.indexOf('記録者');
    var recordIdx = headers.indexOf('記録内容');
    var tsIdx = headers.indexOf('登録日時');
    
    var targetMonth = normalizeTargetMonth(data.targetMonth || '');
    var prevMonth = getPrevMonth(targetMonth);
    var recorder = data.recorder || '';

    // recorderフィルタ
    if (recorder && recorderIdx >= 0) {
      records = records.filter(function(r) {
        return String(r[recorderIdx]).trim() === recorder;
      });
    }

    var now = new Date();
    var todayStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd');
    var isCurrentMonth = (targetMonth === Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM'));

    var monthlyVisits = 0, prevMonthVisits = 0, todayVisits = 0;
    var activeUsers = {};
    var recentRecords = [];

    for (var i = 0; i < records.length; i++) {
      var row = records[i];
      var dateVal = row[dateIdx];
      var dateStr = '';
      if (dateVal instanceof Date) dateStr = Utilities.formatDate(dateVal, 'Asia/Tokyo', 'yyyy/MM/dd');
      else if (dateVal) dateStr = String(dateVal).trim();
      if (!dateStr) continue;
      
      var monthStr = extractMonth(dateStr);
      if (monthStr === targetMonth) {
        monthlyVisits++;
        var pName = patientIdx >= 0 ? String(row[patientIdx]).trim() : '';
        if (pName) activeUsers[pName] = true;
        
        recentRecords.push({
          rowIndex: i + 2, date: dateStr,
          patient: patientIdx >= 0 ? String(row[patientIdx]).trim() : '',
          jobType: jobIdx >= 0 ? String(row[jobIdx]).trim() : '',
          recorder: recorderIdx >= 0 ? String(row[recorderIdx]).trim() : '',
          record: recordIdx >= 0 ? String(row[recordIdx]).trim().substring(0, 100) : '',
          timestamp: tsIdx >= 0 ? String(row[tsIdx]).trim() : ''
        });
      }
      if (monthStr === prevMonth) prevMonthVisits++;
      if (isCurrentMonth && dateStr === todayStr) todayVisits++;
    }
    
    recentRecords.sort(function(a, b) { return b.date.localeCompare(a.date) || b.timestamp.localeCompare(a.timestamp); });
    recentRecords = recentRecords.slice(0, 10);
    
    var monthlyChange = prevMonthVisits > 0 ? Math.round((monthlyVisits - prevMonthVisits) / prevMonthVisits * 100) : 0;
    
    var staffCount = 0, nurseCount = 0, rehabCount = 0;
    if (staffSheet) {
      var staffData = staffSheet.getDataRange().getValues();
      for (var s = 1; s < staffData.length; s++) {
        if (staffData[s][0]) { staffCount++; var jt = String(staffData[s][1]).trim(); if (jt === '看護師') nurseCount++; else if (jt === 'リハビリ職') rehabCount++; }
      }
    }
    
    return { success: true, data: {
      monthlyVisits: monthlyVisits,
      todayVisits: isCurrentMonth ? todayVisits : '-',
      activeUsers: Object.keys(activeUsers).length,
      staffCount: staffCount,
      staffBreakdown: '看護師' + nurseCount + ' / リハ' + rehabCount,
      recentRecords: recentRecords,
      monthlyChange: monthlyChange,
      targetMonth: targetMonth,
      prevMonth: prevMonth
    }};
  } catch (error) { return { success: false, message: 'ダッシュボードエラー: ' + error.toString() }; }
}

/**
 * 記録一覧（フィルター対応）
 */
function getAdminRecords(data) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('記録');
    if (!sheet) return { success: false, message: '記録シートが見つかりません' };
    
    var allData = sheet.getDataRange().getValues();
    if (allData.length <= 1) return { success: true, data: { records: [], total: 0 } };
    
    var headers = allData[0];
    var records = allData.slice(1);
    var dateIdx = headers.indexOf('訪問日'), patientIdx = headers.indexOf('利用者名'), jobIdx = headers.indexOf('職種');
    var recorderIdx = headers.indexOf('記録者'), recordIdx = headers.indexOf('記録内容'), tsIdx = headers.indexOf('登録日時');
    
    var filterStartDate = data.startDate || '', filterEndDate = data.endDate || '';
    var filterPatient = data.patient || '', filterJobType = data.jobType || '', filterRecorder = data.recorder || '';
    
    var filtered = [];
    for (var i = 0; i < records.length; i++) {
      var row = records[i];
      var dateVal = row[dateIdx], dateStr = '';
      if (dateVal instanceof Date) dateStr = Utilities.formatDate(dateVal, 'Asia/Tokyo', 'yyyy/MM/dd');
      else if (dateVal) dateStr = String(dateVal).trim();
      
      if (filterStartDate && dateStr < filterStartDate) continue;
      if (filterEndDate && dateStr > filterEndDate) continue;
      var patient = patientIdx >= 0 ? String(row[patientIdx]).trim() : '';
      if (filterPatient && patient !== filterPatient) continue;
      var jobType = jobIdx >= 0 ? String(row[jobIdx]).trim() : '';
      if (filterJobType && jobType !== filterJobType) continue;
      var recorder = recorderIdx >= 0 ? String(row[recorderIdx]).trim() : '';
      if (filterRecorder && recorder !== filterRecorder) continue;
      
      filtered.push({ rowIndex: i + 2, date: dateStr, patient: patient, jobType: jobType, recorder: recorder,
        recordPreview: recordIdx >= 0 ? String(row[recordIdx]).trim().substring(0, 80) : '',
        timestamp: tsIdx >= 0 ? String(row[tsIdx]).trim() : '' });
    }
    
    filtered.sort(function(a, b) { return b.date.localeCompare(a.date) || b.timestamp.localeCompare(a.timestamp); });
    
    var patients = {}, recorders = {};
    for (var j = 0; j < records.length; j++) {
      var p = patientIdx >= 0 ? String(records[j][patientIdx]).trim() : '';
      var r = recorderIdx >= 0 ? String(records[j][recorderIdx]).trim() : '';
      if (p) patients[p] = true; if (r) recorders[r] = true;
    }
    
    return { success: true, data: { records: filtered, total: filtered.length, patients: Object.keys(patients).sort(), recorders: Object.keys(recorders).sort() } };
  } catch (error) { return { success: false, message: '記録取得エラー: ' + error.toString() }; }
}

/**
 * 記録詳細（行番号指定）
 */
function getAdminRecordDetail(data) {
  try {
    var rowIndex = data.rowIndex;
    if (!rowIndex) return { success: false, message: '行番号が指定されていません' };
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('記録');
    if (!sheet) return { success: false, message: '記録シートが見つかりません' };
    
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var row = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    var dateIdx = headers.indexOf('訪問日'), patientIdx = headers.indexOf('利用者名'), jobIdx = headers.indexOf('職種');
    var recorderIdx = headers.indexOf('記録者'), recordIdx = headers.indexOf('記録内容'), tsIdx = headers.indexOf('登録日時');
    
    var dateVal = row[dateIdx], dateStr = '';
    if (dateVal instanceof Date) dateStr = Utilities.formatDate(dateVal, 'Asia/Tokyo', 'yyyy/MM/dd');
    else if (dateVal) dateStr = String(dateVal).trim();
    
    return { success: true, data: { date: dateStr, patient: patientIdx >= 0 ? String(row[patientIdx]).trim() : '', jobType: jobIdx >= 0 ? String(row[jobIdx]).trim() : '', recorder: recorderIdx >= 0 ? String(row[recorderIdx]).trim() : '', record: recordIdx >= 0 ? String(row[recordIdx]).trim() : '', timestamp: tsIdx >= 0 ? String(row[tsIdx]).trim() : '' } };
  } catch (error) { return { success: false, message: '記録詳細取得エラー: ' + error.toString() }; }
}

/**
 * 利用者一覧（マスタベース）
 * 利用者シートをマスタとして使用。マスタにない誤入力は表示しない。
 * ※ 記録シートは一切変更しません
 */
function getAdminUsers(data) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('記録');
    if (!sheet) return { success: false, message: '記録シートが見つかりません' };

    // 利用者シート（マスタ）を読み込み
    var patientSheet = ss.getSheetByName('利用者');
    var masterList = [];   // [{kanji, kana}]
    var kanaToKanji = {};  // かな名→漢字名
    if (patientSheet && patientSheet.getLastRow() >= 2) {
      var patientData = patientSheet.getRange(2, 1, patientSheet.getLastRow() - 1, 2).getValues();
      for (var p = 0; p < patientData.length; p++) {
        var kana = String(patientData[p][0]).trim();
        var kanji = String(patientData[p][1]).trim();
        if (kanji) {
          masterList.push({ kanji: kanji, kana: kana });
          if (kana) kanaToKanji[kana] = kanji;
        }
      }
    }
    var hasMaster = masterList.length > 0;

    // 記録シートから訪問データを集計
    var allData = sheet.getDataRange().getValues();
    var headers = allData.length > 0 ? allData[0] : [];
    var records = allData.length > 1 ? allData.slice(1) : [];
    var dateIdx = headers.indexOf('訪問日'), patientIdx = headers.indexOf('利用者名'), recorderIdx = headers.indexOf('記録者');
    var now = new Date();
    var thisMonth = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM');
    var userMap = {};  // key=記録上の名前（カタカナ）

    for (var i = 0; i < records.length; i++) {
      var row = records[i];
      var patient = patientIdx >= 0 ? String(row[patientIdx]).trim() : '';
      if (!patient) continue;
      var dateVal = row[dateIdx], dateStr = '';
      if (dateVal instanceof Date) dateStr = Utilities.formatDate(dateVal, 'Asia/Tokyo', 'yyyy/MM/dd');
      else if (dateVal) dateStr = String(dateVal).trim();
      var recorder = recorderIdx >= 0 ? String(row[recorderIdx]).trim() : '';

      if (!userMap[patient]) userMap[patient] = { monthlyCount: 0, lastVisit: '', recorders: {} };
      if (dateStr > userMap[patient].lastVisit) userMap[patient].lastVisit = dateStr;
      if (dateStr.substring(0, 7) === thisMonth) userMap[patient].monthlyCount++;
      if (recorder) userMap[patient].recorders[recorder] = (userMap[patient].recorders[recorder] || 0) + 1;
    }

    var users = [];

    if (hasMaster) {
      // マスタがある場合：マスタの利用者だけを表示（誤入力は除外）
      for (var m = 0; m < masterList.length; m++) {
        var kana = masterList[m].kana;
        var kanji = masterList[m].kanji;
        var info = userMap[kana] || { monthlyCount: 0, lastVisit: '', recorders: {} };
        var mainRecorder = '', maxCount = 0;
        var recorderNames = Object.keys(info.recorders);
        for (var r = 0; r < recorderNames.length; r++) { if (info.recorders[recorderNames[r]] > maxCount) { maxCount = info.recorders[recorderNames[r]]; mainRecorder = recorderNames[r]; } }
        users.push({ name: kanji, kanaName: kana, monthlyCount: info.monthlyCount, lastVisit: info.lastVisit, mainRecorder: mainRecorder });
      }
      // 漢字名の五十音順でソート
      users.sort(function(a, b) { return (a.kanaName || a.name).localeCompare(b.kanaName || b.name, 'ja'); });
    } else {
      // マスタがない場合：従来通り記録シートから取得
      var names = Object.keys(userMap).sort();
      for (var u = 0; u < names.length; u++) {
        var info2 = userMap[names[u]];
        var mainRecorder2 = '', maxCount2 = 0;
        var recorderNames2 = Object.keys(info2.recorders);
        for (var r2 = 0; r2 < recorderNames2.length; r2++) { if (info2.recorders[recorderNames2[r2]] > maxCount2) { maxCount2 = info2.recorders[recorderNames2[r2]]; mainRecorder2 = recorderNames2[r2]; } }
        users.push({ name: names[u], kanaName: '', monthlyCount: info2.monthlyCount, lastVisit: info2.lastVisit, mainRecorder: mainRecorder2 });
      }
    }

    return { success: true, data: { users: users } };
  } catch (error) { return { success: false, message: '利用者取得エラー: ' + error.toString() }; }
}

/**
 * スタッフ一覧
 * data.targetMonth: "2026/01" 等 → その月の記録件数を集計（未指定は今月）
 */
function getAdminStaff(data) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var staffSheet = ss.getSheetByName('従業員');
    var recordSheet = ss.getSheetByName('記録');
    if (!staffSheet) return { success: false, message: '従業員シートが見つかりません' };
    
    var targetMonth = normalizeTargetMonth(data.targetMonth || '');
    
    var staffData = staffSheet.getDataRange().getValues();
    var monthlyCounts = {};
    if (recordSheet) {
      var recData = recordSheet.getDataRange().getValues();
      if (recData.length > 1) {
        var recHeaders = recData[0], recDateIdx = recHeaders.indexOf('訪問日'), recRecorderIdx = recHeaders.indexOf('記録者');
        for (var i = 1; i < recData.length; i++) {
          var dateVal = recData[i][recDateIdx], dateStr = '';
          if (dateVal instanceof Date) dateStr = Utilities.formatDate(dateVal, 'Asia/Tokyo', 'yyyy/MM/dd');
          else if (dateVal) dateStr = String(dateVal).trim();
          var monthStr = extractMonth(dateStr);
          if (monthStr === targetMonth) {
            var rec = recRecorderIdx >= 0 ? String(recData[i][recRecorderIdx]).trim() : '';
            if (rec) monthlyCounts[rec] = (monthlyCounts[rec] || 0) + 1;
          }
        }
      }
    }
    
    var staffList = [];
    for (var s = 1; s < staffData.length; s++) {
      var name = String(staffData[s][0]).trim(), jobType = String(staffData[s][1]).trim();
      if (!name) continue;
      staffList.push({ name: name, jobType: jobType, monthlyCount: monthlyCounts[name] || 0 });
    }
    
    return { success: true, data: { staff: staffList, targetMonth: targetMonth } };
  } catch (error) { return { success: false, message: 'スタッフ取得エラー: ' + error.toString() }; }
}

/**
 * 報告書一覧（記録者情報付き）
 */
function getAdminReports(data) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheets = ss.getSheets();
    
    var patientRecorderMap = {};
    var recordSheet = ss.getSheetByName('記録');
    if (recordSheet) {
      var recData = recordSheet.getDataRange().getValues();
      if (recData.length > 1) {
        var recHeaders = recData[0], recPatientIdx = recHeaders.indexOf('利用者名'), recRecorderIdx = recHeaders.indexOf('記録者');
        for (var ri = 1; ri < recData.length; ri++) {
          var patient = recPatientIdx >= 0 ? String(recData[ri][recPatientIdx]).trim() : '';
          var recorder = recRecorderIdx >= 0 ? String(recData[ri][recRecorderIdx]).trim() : '';
          if (patient && recorder) {
            if (!patientRecorderMap[patient]) patientRecorderMap[patient] = {};
            patientRecorderMap[patient][recorder] = (patientRecorderMap[patient][recorder] || 0) + 1;
          }
        }
      }
    }
    
    function getMainRecorder(patientName) {
      var recorders = patientRecorderMap[patientName];
      if (!recorders) return '';
      var maxCount = 0, mainRec = '', names = Object.keys(recorders);
      for (var k = 0; k < names.length; k++) { if (recorders[names[k]] > maxCount) { maxCount = recorders[names[k]]; mainRec = names[k]; } }
      return mainRec;
    }
    
    var reports = [];
    for (var i = 0; i < sheets.length; i++) {
      var name = sheets[i].getName();
      if (name.indexOf('病状の経過_') === 0) {
        var period = name.replace('病状の経過_', '');
        var sheetData = sheets[i].getDataRange().getValues();
        var recordCount = sheetData.length - 1;
        var reportEntries = [];
        if (sheetData.length > 1) {
          var reportHeaders = sheetData[0];
          var nameIdx = reportHeaders.indexOf('利用者名'), countIdx = reportHeaders.indexOf('訪問回数');
          var progressIdx = reportHeaders.indexOf('病状の経過'), rehabIdx = reportHeaders.indexOf('リハビリ別添'), genDateIdx = reportHeaders.indexOf('生成日時');
          for (var r = 1; r < sheetData.length; r++) {
            var patientName = nameIdx >= 0 ? String(sheetData[r][nameIdx]).trim() : '';
            reportEntries.push({
              patient: patientName, recorder: getMainRecorder(patientName),
              visitCount: countIdx >= 0 ? sheetData[r][countIdx] : 0,
              hasProgress: progressIdx >= 0 ? !!String(sheetData[r][progressIdx]).trim() : false,
              hasRehab: rehabIdx >= 0 ? !!String(sheetData[r][rehabIdx]).trim() : false,
              progress: progressIdx >= 0 ? String(sheetData[r][progressIdx]).trim() : '',
              rehabAddendum: rehabIdx >= 0 ? String(sheetData[r][rehabIdx]).trim() : '',
              generatedAt: genDateIdx >= 0 ? String(sheetData[r][genDateIdx]).trim() : ''
            });
          }
        }
        reports.push({ period: period, sheetName: name, userCount: recordCount, entries: reportEntries });
      }
    }
    
    reports.sort(function(a, b) { return b.period.localeCompare(a.period); });
    return { success: true, data: { reports: reports } };
  } catch (error) { return { success: false, message: '報告書取得エラー: ' + error.toString() }; }
}

/**
 * 営業時間内/外の記録割合
 */
function getAdminBusinessHours(data) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('記録');
    if (!sheet) return { success: false, message: '記録シートが見つかりません' };
    
    var allData = sheet.getDataRange().getValues();
    if (allData.length <= 1) return { success: true, data: { staffHours: [] } };
    
    var headers = allData[0], records = allData.slice(1);
    var recorderIdx = headers.indexOf('記録者'), tsIdx = headers.indexOf('登録日時'), dateIdx = headers.indexOf('訪問日');
    if (tsIdx === -1) return { success: false, message: '登録日時列が見つかりません' };
    
    var bizStartHour = parseInt(data.startHour || '9'), bizStartMin = parseInt(data.startMinute || '0');
    var bizEndHour = parseInt(data.endHour || '18'), bizEndMin = parseInt(data.endMinute || '0');
    var bizStartTotal = bizStartHour * 60 + bizStartMin, bizEndTotal = bizEndHour * 60 + bizEndMin;
    
    var targetMonth = normalizeTargetMonth(data.targetMonth || '');
    
    var staffMap = {};
    
    for (var i = 0; i < records.length; i++) {
      var row = records[i];
      var recorder = recorderIdx >= 0 ? String(row[recorderIdx]).trim() : '';
      if (!recorder) continue;
      
      var dateVal = row[dateIdx], visitDateStr = '';
      if (dateVal instanceof Date) visitDateStr = Utilities.formatDate(dateVal, 'Asia/Tokyo', 'yyyy/MM/dd');
      else if (dateVal) visitDateStr = String(dateVal).trim();
      if (!visitDateStr) continue;
      
      var recordMonth = extractMonth(visitDateStr);
      if (recordMonth !== targetMonth) continue;
      
      var tsVal = row[tsIdx];
      var tsDateStr = '';
      var hour = -1, minute = 0;
      
      if (tsVal instanceof Date) {
        tsDateStr = Utilities.formatDate(tsVal, 'Asia/Tokyo', 'yyyy/MM/dd');
        hour = tsVal.getHours();
        minute = tsVal.getMinutes();
      } else if (tsVal) {
        var tsStr = String(tsVal).trim();
        var dateMatch = tsStr.match(/(\d{4}[\/-]\d{1,2}[\/-]\d{1,2})/);
        if (dateMatch) tsDateStr = dateMatch[1].replace(/-/g, '/');
        var timeMatch = tsStr.match(/(\d{1,2}):(\d{2})/);
        if (timeMatch) { hour = parseInt(timeMatch[1]); minute = parseInt(timeMatch[2]); }
      }
      
      if (hour < 0) continue;
      
      if (!staffMap[recorder]) staffMap[recorder] = { inHours: 0, outHours: 0, total: 0 };
      staffMap[recorder].total++;
      
      var timeTotal = hour * 60 + minute;
      var isSameDay = (visitDateStr === tsDateStr);
      var isWithinBizHours = (timeTotal >= bizStartTotal && timeTotal < bizEndTotal);
      
      if (isSameDay && isWithinBizHours) {
        staffMap[recorder].inHours++;
      } else {
        staffMap[recorder].outHours++;
      }
    }
    
    var staffHours = [];
    var names = Object.keys(staffMap).sort();
    for (var s = 0; s < names.length; s++) {
      var info = staffMap[names[s]];
      staffHours.push({ name: names[s], inHours: info.inHours, outHours: info.outHours, total: info.total, inRate: info.total > 0 ? Math.round(info.inHours / info.total * 100) : 0 });
    }
    
    return { success: true, data: {
      staffHours: staffHours,
      businessHours: { start: String(bizStartHour).padStart(2, '0') + ':' + String(bizStartMin).padStart(2, '0'), end: String(bizEndHour).padStart(2, '0') + ':' + String(bizEndMin).padStart(2, '0') },
      targetMonth: targetMonth
    }};
  } catch (error) { return { success: false, message: '営業時間分析エラー: ' + error.toString() }; }
}


// ============================================
// 管理画面から報告書生成・進捗確認・編集保存
// ============================================

/**
 * 管理画面から報告書生成を開始
 * data.targetMonth: "2026/01" 形式
 */
function adminGenerateReports(data) {
  try {
    if (isProcessing()) {
      return { success: false, message: '現在別の処理が進行中です。完了後に再実行してください。' };
    }
    
    var inputMonth = data.targetMonth || '';
    if (!inputMonth) return { success: false, message: '対象月を指定してください' };
    
    var dateMatch = inputMonth.match(/(\d{4})[\/\-](\d{1,2})/);
    if (!dateMatch) return { success: false, message: '年月の形式が正しくありません（例：2026/01）' };
    
    var year = parseInt(dateMatch[1]);
    var month = parseInt(dateMatch[2]);
    var startDate = new Date(year, month - 1, 1);
    var endDate = new Date(year, month, 0);
    var sheetLabel = year + '年' + String(month).padStart(2, '0') + '月';
    
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var inputSheet = ss.getSheetByName('記録');
    if (!inputSheet) return { success: false, message: '記録シートが見つかりません' };
    
    var allData = inputSheet.getDataRange().getValues();
    if (allData.length <= 1) return { success: false, message: 'データがありません' };
    
    var headers = allData[0];
    var records = allData.slice(1);
    var dateIndex = headers.indexOf('訪問日');
    if (dateIndex === -1) return { success: false, message: '訪問日列が見つかりません' };
    
    records = records.filter(function(row) {
      var visitDate = parseDate(row[dateIndex]);
      if (!visitDate) return false;
      return visitDate >= startDate && visitDate <= endDate;
    });
    
    if (records.length === 0) return { success: false, message: '指定月のデータがありません' };
    
    var userRecords = groupRecordsByUser(records, headers);
    var userNames = Object.keys(userRecords).sort();
    var outputSheetName = '病状の経過_' + sheetLabel;
    
    var existingSheet = ss.getSheetByName(outputSheetName);
    var outputSheet;
    if (existingSheet) { outputSheet = existingSheet; outputSheet.clear(); }
    else { outputSheet = ss.insertSheet(outputSheetName); }
    
    var outputHeaders = ['利用者名', '訪問回数', '対象期間', '病状の経過', 'リハビリ別添', '生成日時'];
    outputSheet.getRange(1, 1, 1, outputHeaders.length).setValues([outputHeaders]);
    outputSheet.getRange(1, 1, 1, outputHeaders.length).setFontWeight('bold').setBackground('#00a67d').setFontColor('white').setHorizontalAlignment('center');
    outputSheet.setColumnWidth(1, 120); outputSheet.setColumnWidth(2, 80); outputSheet.setColumnWidth(3, 120);
    outputSheet.setColumnWidth(4, 600); outputSheet.setColumnWidth(5, 600); outputSheet.setColumnWidth(6, 140);
    
    saveTempData(ss, userNames, userRecords, headers);
    
    var props = PropertiesService.getScriptProperties();
    props.setProperties({
      'BATCH_CURRENT_INDEX': '0', 'BATCH_TOTAL_COUNT': String(userNames.length),
      'BATCH_START_TIME': new Date().toISOString(), 'BATCH_PERIOD_TEXT': sheetLabel,
      'BATCH_OUTPUT_SHEET': outputSheetName, 'PROCESSING_FLAG': 'true'
    });
    
    ScriptApp.newTrigger('processBatchV2').timeBased().after(1000).create();
    
    return { 
      success: true, 
      data: { 
        message: '生成を開始しました',
        totalCount: userNames.length,
        recordCount: records.length,
        period: sheetLabel,
        outputSheet: outputSheetName
      }
    };
  } catch (error) {
    return { success: false, message: '報告書生成エラー: ' + error.toString() };
  }
}

/**
 * 報告書生成の進捗を返す
 */
function adminGenerateStatus(data) {
  try {
    var props = PropertiesService.getScriptProperties();
    var processing = props.getProperty('PROCESSING_FLAG') === 'true';
    var currentIndex = parseInt(props.getProperty('BATCH_CURRENT_INDEX') || '0');
    var totalCount = parseInt(props.getProperty('BATCH_TOTAL_COUNT') || '0');
    var periodText = props.getProperty('BATCH_PERIOD_TEXT') || '';
    var startTimeStr = props.getProperty('BATCH_START_TIME') || '';
    
    var elapsed = 0;
    if (startTimeStr) {
      elapsed = Math.round((new Date() - new Date(startTimeStr)) / 1000 / 60);
    }
    
    return {
      success: true,
      data: {
        processing: processing,
        currentIndex: currentIndex,
        totalCount: totalCount,
        percent: totalCount > 0 ? Math.round(currentIndex / totalCount * 100) : 0,
        period: periodText,
        elapsedMinutes: elapsed
      }
    };
  } catch (error) {
    return { success: false, message: '進捗確認エラー: ' + error.toString() };
  }
}

/**
 * 報告書の内容を更新（シートに書き戻す）
 * data.sheetName: シート名（例: "病状の経過_2026年01月"）
 * data.patient: 利用者名
 * data.field: "progress" or "rehabAddendum"
 * data.value: 新しいテキスト
 */
function adminUpdateReport(data) {
  try {
    var sheetName = data.sheetName || '';
    var patient = data.patient || '';
    var field = data.field || '';
    var value = data.value || '';
    
    if (!sheetName || !patient || !field) {
      return { success: false, message: '必要な情報が不足しています' };
    }
    
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, message: 'シート「' + sheetName + '」が見つかりません' };
    
    var allData = sheet.getDataRange().getValues();
    if (allData.length <= 1) return { success: false, message: 'データがありません' };
    
    var headers = allData[0];
    var nameIdx = headers.indexOf('利用者名');
    var progressIdx = headers.indexOf('病状の経過');
    var rehabIdx = headers.indexOf('リハビリ別添');
    
    if (nameIdx === -1) return { success: false, message: '利用者名列が見つかりません' };
    
    var targetColIdx = -1;
    if (field === 'progress') targetColIdx = progressIdx;
    else if (field === 'rehabAddendum') targetColIdx = rehabIdx;
    
    if (targetColIdx === -1) return { success: false, message: '対象列が見つかりません' };
    
    for (var i = 1; i < allData.length; i++) {
      if (String(allData[i][nameIdx]).trim() === patient) {
        sheet.getRange(i + 1, targetColIdx + 1).setValue(value);
        return { success: true, data: { message: patient + 'の報告書を更新しました' } };
      }
    }
    
    return { success: false, message: '利用者「' + patient + '」が見つかりません' };
  } catch (error) {
    return { success: false, message: '報告書更新エラー: ' + error.toString() };
  }
}

// ============================================
// 利用者マスタ管理
// ============================================

/**
 * 利用者マスタのシートセットアップ
 */
function setupPatientsSheet() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('利用者');
  if (!sheet) {
    sheet = ss.insertSheet('利用者');
  }
  var headers = ['利用者名（カナ）', '利用者名（漢字）'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#00a67d');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 180);
  sheet.setFrozenRows(1);
}

/**
 * 利用者リスト取得（入力画面・管理画面共通）
 */
function getPatientsList() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('利用者');

    if (!sheet) {
      return { success: true, data: { patients: [] } };
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, data: { patients: [] } };
    }

    var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    var patients = data
      .filter(function(row) { return row[0] || row[1]; })
      .map(function(row) {
        return {
          kana: String(row[0] || '').trim(),
          kanji: String(row[1] || '').trim()
        };
      });

    // カナ順でソート
    patients.sort(function(a, b) { return a.kana.localeCompare(b.kana, 'ja'); });

    return { success: true, data: { patients: patients } };
  } catch (error) {
    return { success: false, message: '利用者リスト取得エラー: ' + error.toString() };
  }
}

/**
 * 利用者を1名追加
 */
function addPatient(data) {
  try {
    var kana = (data.kana || '').trim();
    var kanji = (data.kanji || '').trim();

    if (!kana) {
      return { success: false, message: 'カナ名を入力してください' };
    }

    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('利用者');

    if (!sheet) {
      setupPatientsSheet();
      sheet = ss.getSheetByName('利用者');
    }

    // 重複チェック
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < existing.length; i++) {
        if (String(existing[i][0]).trim() === kana) {
          return { success: false, message: '同じカナ名の利用者が既に登録されています: ' + kana };
        }
      }
    }

    sheet.appendRow([kana, kanji]);

    return { success: true, data: { message: kana + ' を登録しました' } };
  } catch (error) {
    return { success: false, message: '利用者追加エラー: ' + error.toString() };
  }
}

/**
 * 利用者をCSVから一括インポート
 * data.csvData: "カナ,漢字\nヤマダタロウ,山田太郎\n..."
 * data.mode: "append"（追加）or "replace"（全置換）
 */
function importPatients(data) {
  try {
    var csvData = (data.csvData || '').trim();
    var mode = data.mode || 'append';

    if (!csvData) {
      return { success: false, message: 'CSVデータがありません' };
    }

    var lines = csvData.split('\n');
    var patients = [];
    var startLine = 0;

    // ヘッダー行をスキップ（カナ or 利用者 を含む場合）
    if (lines.length > 0 && (lines[0].indexOf('カナ') >= 0 || lines[0].indexOf('利用者') >= 0)) {
      startLine = 1;
    }

    for (var i = startLine; i < lines.length; i++) {
      var line = lines[i].trim();
      if (!line) continue;

      // CSV解析（簡易：カンマ区切り、ダブルクォート対応）
      var cols = parseCSVLine(line);
      var kana = (cols[0] || '').trim();
      var kanji = (cols[1] || '').trim();

      if (kana) {
        patients.push([kana, kanji]);
      }
    }

    if (patients.length === 0) {
      return { success: false, message: 'インポートするデータがありません' };
    }

    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('利用者');

    if (!sheet) {
      setupPatientsSheet();
      sheet = ss.getSheetByName('利用者');
    }

    if (mode === 'replace') {
      // ヘッダー以外を全削除
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, 2).clearContent();
      }
      // 新データを書き込み
      sheet.getRange(2, 1, patients.length, 2).setValues(patients);
    } else {
      // 追加モード：既存データとの重複チェック
      var lastRow = sheet.getLastRow();
      var existingKana = {};
      if (lastRow >= 2) {
        var existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (var e = 0; e < existing.length; e++) {
          existingKana[String(existing[e][0]).trim()] = true;
        }
      }

      var newPatients = patients.filter(function(p) { return !existingKana[p[0]]; });
      var skipped = patients.length - newPatients.length;

      if (newPatients.length > 0) {
        var insertRow = sheet.getLastRow() + 1;
        sheet.getRange(insertRow, 1, newPatients.length, 2).setValues(newPatients);
      }

      return {
        success: true,
        data: {
          message: newPatients.length + '名を追加しました' + (skipped > 0 ? '（重複' + skipped + '名スキップ）' : ''),
          imported: newPatients.length,
          skipped: skipped
        }
      };
    }

    return {
      success: true,
      data: {
        message: patients.length + '名をインポートしました',
        imported: patients.length,
        skipped: 0
      }
    };
  } catch (error) {
    return { success: false, message: 'インポートエラー: ' + error.toString() };
  }
}

/**
 * CSVから利用者をインポート（利用者シートに書き込み、記録シートは変更しない）
 * data.csvData: CSV文字列（1列目=漢字名, 2列目=かな名）
 * data.mode: "append"（重複スキップ）or "replace"（全置換）
 */
function importPatientsFromCsv(data) {
  try {
    var csvData = data.csvData || '';
    var mode = data.mode || 'append';

    if (!csvData.trim()) return { success: false, message: 'CSVデータが空です' };

    var lines = csvData.split(/\r?\n/).filter(function(l) { return l.trim(); });
    if (lines.length === 0) return { success: false, message: 'CSVデータが空です' };

    // ヘッダー行の判定（漢字名 or かな名 を含む行はスキップ）
    var startIdx = 0;
    if (lines[0].indexOf('漢字名') >= 0 || lines[0].indexOf('かな名') >= 0) {
      startIdx = 1;
    }

    var patients = [];
    for (var i = startIdx; i < lines.length; i++) {
      var cols = lines[i].split(',');
      var kanjiName = (cols[0] || '').trim();
      var kanaName = (cols[1] || '').trim();
      if (kanjiName) {
        patients.push([kanjiName, kanaName]);
      }
    }

    if (patients.length === 0) return { success: false, message: 'インポートするデータがありません' };

    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('利用者');

    // 利用者シートがなければ作成
    if (!sheet) {
      sheet = ss.insertSheet('利用者');
      sheet.getRange(1, 1, 1, 2).setValues([['漢字名', 'かな名']]);
      sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#00a67d').setFontColor('white').setHorizontalAlignment('center');
      sheet.setColumnWidth(1, 150);
      sheet.setColumnWidth(2, 150);
      sheet.setFrozenRows(1);
    }

    var imported = 0, skipped = 0;

    if (mode === 'replace') {
      // 全置換: ヘッダー以外を削除して全件インポート
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
      if (patients.length > 0) {
        sheet.getRange(2, 1, patients.length, 2).setValues(patients);
      }
      imported = patients.length;
    } else {
      // 追加モード: かな名が既存と一致する場合はスキップ
      var existing = {};
      var lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        var existingData = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
        for (var e = 0; e < existingData.length; e++) {
          existing[String(existingData[e][1]).trim()] = true; // かな名でチェック
        }
      }

      var newPatients = [];
      for (var p = 0; p < patients.length; p++) {
        var kana = patients[p][1]; // かな名
        if (existing[kana]) {
          skipped++;
        } else {
          newPatients.push(patients[p]);
          imported++;
        }
      }

      if (newPatients.length > 0) {
        var insertRow = sheet.getLastRow() + 1;
        sheet.getRange(insertRow, 1, newPatients.length, 2).setValues(newPatients);
      }
    }

    return { success: true, data: { imported: imported, skipped: skipped } };
  } catch (error) {
    return { success: false, message: 'インポートエラー: ' + error.toString() };
  }
}

/**
 * 利用者を1名削除
 */
function deletePatient(data) {
  try {
    var kana = (data.kana || '').trim();
    if (!kana) return { success: false, message: '削除対象が指定されていません' };

    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName('利用者');
    if (!sheet) return { success: false, message: '利用者シートが見つかりません' };

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: '利用者が登録されていません' };

    var allData = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (var i = 0; i < allData.length; i++) {
      if (String(allData[i][0]).trim() === kana) {
        sheet.deleteRow(i + 2);
        return { success: true, data: { message: kana + ' を削除しました' } };
      }
    }

    return { success: false, message: '利用者が見つかりません: ' + kana };
  } catch (error) {
    return { success: false, message: '利用者削除エラー: ' + error.toString() };
  }
}

/**
 * 簡易CSVパーサー（ダブルクォート対応）
 */
function parseCSVLine(line) {
  var result = [];
  var current = '';
  var inQuotes = false;

  for (var i = 0; i < line.length; i++) {
    var ch = line[i];
    if (inQuotes) {
      if (ch === '"') {
        if (i + 1 < line.length && line[i + 1] === '"') {
          current += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        current += ch;
      }
    } else {
      if (ch === '"') {
        inQuotes = true;
      } else if (ch === ',') {
        result.push(current);
        current = '';
      } else {
        current += ch;
      }
    }
  }
  result.push(current);
  return result;
}
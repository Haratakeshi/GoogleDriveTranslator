// Config.gs - 設定管理ファイル

/**
 * アプリケーション設定
 * スクリプトプロパティから認証情報を取得
 */
const CONFIG = {
  // OpenAI API設定（スクリプトプロパティから取得）
  get OPENAI_API_KEY() {
    return PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  },
  get OPENAI_ORGANIZATION_ID() {
    return PropertiesService.getScriptProperties().getProperty('OPENAI_ORGANIZATION_ID');
  },
  get OPENAI_PROJECT_ID() {
    return PropertiesService.getScriptProperties().getProperty('OPENAI_PROJECT_ID');
  },
  OPENAI_MODEL: 'gpt-4o-mini', // モデル名を修正
  OPENAI_API_URL: 'https://api.openai.com/v1/chat/completions',
  
  // スプレッドシートID（初回実行時に自動作成）
  get HISTORY_SHEET_ID() {
    let id = PropertiesService.getScriptProperties().getProperty('HISTORY_SHEET_ID');
    if (!id) {
      id = createHistorySpreadsheet();
      PropertiesService.getScriptProperties().setProperty('HISTORY_SHEET_ID', id);
    }
    return id;
  },
  
  get DICTIONARY_SHEET_ID() {
    let id = PropertiesService.getScriptProperties().getProperty('DICTIONARY_SHEET_ID');
    if (!id) {
      id = createDictionarySpreadsheet();
      PropertiesService.getScriptProperties().setProperty('DICTIONARY_SHEET_ID', id);
    }
    return id;
  },
  
  // 翻訳設定
  MAX_CHARS_PER_REQUEST: 3000, // 1回のリクエストの最大文字数
  MAX_RETRY_COUNT: 3, // 再試行回数
  LENGTH_THRESHOLD: 1.5, // 文字数閾値（150%）
  BATCH_SIZE: 50, // バッチ処理のサイズ
  
  // API制限
  API_RATE_LIMIT_DELAY: 1000, // APIコール間の遅延（ミリ秒）
  MAX_CONCURRENT_REQUESTS: 3, // 同時リクエスト数
  
  // サポート言語
  SUPPORTED_LANGUAGES: {
    'ja': '日本語',
    'en': '英語',
    'zh-CN': '中国語（簡体字）',
    'zh-TW': '中国語（繁体字）',
    'ko': '韓国語',
    'es': 'スペイン語',
    'fr': 'フランス語',
    'de': 'ドイツ語',
    'it': 'イタリア語',
    'pt': 'ポルトガル語',
    'ru': 'ロシア語',
    'ar': 'アラビア語',
    'hi': 'ヒンディー語',
    'th': 'タイ語',
    'vi': 'ベトナム語'
  },
  
  // ファイルタイプ設定
  SUPPORTED_MIME_TYPES: {
    'application/vnd.google-apps.spreadsheet': {
      type: 'spreadsheet',
      name: 'スプレッドシート'
    },
    'application/vnd.google-apps.document': {
      type: 'document', 
      name: 'ドキュメント'
    },
    'application/vnd.google-apps.presentation': {
      type: 'presentation',
      name: 'プレゼンテーション'
    }
  },
  
  // エラーメッセージ
  ERROR_MESSAGES: {
    INVALID_URL: '無効なGoogle DriveのURLです',
    FILE_NOT_FOUND: 'ファイルが見つかりません',
    NO_PERMISSION: 'ファイルへのアクセス権限がありません',
    UNSUPPORTED_FILE: 'サポートされていないファイル形式です',
    API_ERROR: 'APIエラーが発生しました',
    TRANSLATION_FAILED: '翻訳に失敗しました',
    NETWORK_ERROR: 'ネットワークエラーが発生しました'
  },
  
  // デフォルト設定
  DEFAULT_DICT_NAME: '一般辞書',
  DEFAULT_SOURCE_LANG: 'auto', // 自動検出
  
  // ログ設定
  ENABLE_LOGGING: true,
  LOG_LEVEL: 'INFO' // DEBUG, INFO, WARN, ERROR
};

/**
 * 履歴管理用スプレッドシートを作成
 * @return {string} 作成したスプレッドシートのID
 */
function createHistorySpreadsheet() {
  const spreadsheet = SpreadsheetApp.create('翻訳ツール_履歴データ');
  const sheet = spreadsheet.getActiveSheet();
  sheet.setName('翻訳履歴');
  
  // ヘッダー設定
  const headers = [
    'タイムスタンプ', '翻訳元URL', '翻訳先URL', '原文言語', 
    '翻訳先言語', '使用辞書', '原文文字数', '翻訳後文字数', 
    '処理時間（秒）', 'ステータス', 'エラーメッセージ'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // スタイル設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#34a853');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // 列幅の調整
  sheet.setColumnWidth(1, 150); // タイムスタンプ
  sheet.setColumnWidth(2, 300); // 翻訳元URL
  sheet.setColumnWidth(3, 300); // 翻訳先URL
  
  // フリーズ設定
  sheet.setFrozenRows(1);
  
  Logger.log('履歴スプレッドシートを作成しました: ' + spreadsheet.getId());
  return spreadsheet.getId();
}

/**
 * 辞書管理用スプレッドシートを作成
 * @return {string} 作成したスプレッドシートのID
 */
function createDictionarySpreadsheet() {
  const spreadsheet = SpreadsheetApp.create('翻訳ツール_辞書データ');
  
  // デフォルト辞書シートを作成
  const sheet = spreadsheet.getActiveSheet();
  sheet.setName(CONFIG.DEFAULT_DICT_NAME);
  
  // ヘッダー設定
  const headers = ['原語', '訳語', '品詞', '備考', '登録日時', '使用回数'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // スタイル設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4a86e8');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // 列幅の調整
  sheet.setColumnWidth(1, 200); // 原語
  sheet.setColumnWidth(2, 200); // 訳語
  sheet.setColumnWidth(3, 100); // 品詞
  sheet.setColumnWidth(4, 200); // 備考
  sheet.setColumnWidth(5, 150); // 登録日時
  sheet.setColumnWidth(6, 100); // 使用回数
  
  // フリーズ設定
  sheet.setFrozenRows(1);
  
  // サンプルデータを追加（オプション）
  const sampleData = [
    ['Google Drive', 'グーグルドライブ', '名詞', 'サービス名', new Date(), 0],
    ['Spreadsheet', 'スプレッドシート', '名詞', '表計算ソフト', new Date(), 0],
    ['Document', 'ドキュメント', '名詞', '文書', new Date(), 0]
  ];
  
  if (sampleData.length > 0) {
    sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);
  }
  
  Logger.log('辞書スプレッドシートを作成しました: ' + spreadsheet.getId());
  return spreadsheet.getId();
}

/**
 * 設定の初期化（初回実行時に呼び出し）
 */
function initializeConfig() {
  // 必要な設定が揃っているかチェック
  const requiredProperties = ['OPENAI_API_KEY', 'OPENAI_ORGANIZATION_ID', 'OPENAI_PROJECT_ID'];
  const scriptProperties = PropertiesService.getScriptProperties();
  const properties = scriptProperties.getProperties();
  
  const missingProperties = requiredProperties.filter(prop => !properties[prop]);
  
  if (missingProperties.length > 0) {
    throw new Error(`以下のスクリプトプロパティが設定されていません: ${missingProperties.join(', ')}`);
  }
  
  // スプレッドシートIDが設定されていない場合は作成
  if (!properties['HISTORY_SHEET_ID']) {
    CONFIG.HISTORY_SHEET_ID; // getterが自動的に作成
  }
  
  if (!properties['DICTIONARY_SHEET_ID']) {
    CONFIG.DICTIONARY_SHEET_ID; // getterが自動的に作成
  }
  
  Logger.log('設定の初期化が完了しました');
}

/**
 * ログ出力関数
 * @param {string} level - ログレベル（DEBUG, INFO, WARN, ERROR）
 * @param {string} message - ログメッセージ
 * @param {Object} data - 追加データ（オプション）
 */
function log(level, message, data = null) {
  if (!CONFIG.ENABLE_LOGGING) return;
  
  const levels = ['DEBUG', 'INFO', 'WARN', 'ERROR'];
  const currentLevelIndex = levels.indexOf(CONFIG.LOG_LEVEL);
  const messageLevelIndex = levels.indexOf(level);
  
  if (messageLevelIndex >= currentLevelIndex) {
    const timestamp = new Date().toISOString();
    const logMessage = `[${timestamp}] [${level}] ${message}`;
    
    try {
      if (data) {
        // エラーオブジェクトなどを安全に文字列化
        const dataString = (data instanceof Error) ? 
          `Error: ${data.message} (Stack: ${data.stack})` : 
          JSON.stringify(data, null, 2);
        
        console.log(logMessage, dataString);
        if (level === 'ERROR') {
          Logger.log(logMessage + '\n' + dataString);
        }
      } else {
        console.log(logMessage);
        if (level === 'ERROR') {
          Logger.log(logMessage);
        }
      }
    } catch (e) {
      // ログ出力自体でエラーが発生した場合のフォールバック
      console.log(`[${timestamp}] [FATAL] Failed to log message: ${message}`);
      Logger.log(`[${timestamp}] [FATAL] Failed to log message: ${message}`);
    }
  }
}
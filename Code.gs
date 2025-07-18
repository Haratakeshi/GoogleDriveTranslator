// Code.gs - メインのエントリーポイント

/**
 * WebアプリケーションのGETリクエストハンドラ
 * @return {HtmlOutput} HTMLページ
 */
function doGet() {
  try {
    // 初期化処理
    initializeConfig();
    
    return HtmlService.createTemplateFromFile('UI')
      .evaluate()
      .setTitle('Google Drive翻訳ツール')
      .setWidth(1000)
      .setHeight(800)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (error) {
    log('ERROR', 'doGet error', error);
    return HtmlService.createHtmlOutput('エラーが発生しました: ' + error.message);
  }
}

/**
 * HTMLファイルに他のHTMLファイルを含める
 * @param {string} filename - インクルードするファイル名
 * @return {string} HTMLコンテンツ
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 翻訳ジョブのキューをセットアップする（クライアントから呼び出される）
 * @param {string} fileUrl - Google DriveファイルのURL
 * @param {string} targetLang - 翻訳先言語コード
 * @param {string} dictName - 使用する辞書名（オプション）
 * @return {Object} タスク情報（taskId, totalJobs, targetFileUrl）
 */
function setupTranslationQueue(fileUrl, targetLang, dictName) {
  log('INFO', `翻訳キューセットアップ開始: URL=${fileUrl}, 言語=${targetLang}, 辞書=${dictName}`);
  
  try {
    if (!fileUrl || !isValidGoogleDriveUrl(fileUrl)) {
      throw new Error(CONFIG.ERROR_MESSAGES.INVALID_URL);
    }
    
    const fileId = extractFileId(fileUrl);
    if (!fileId) {
      throw new Error(CONFIG.ERROR_MESSAGES.INVALID_URL);
    }

    const fileHandler = new FileHandler();
    const fileInfo = fileHandler.getFileInfo(fileId);
    log('INFO', `ファイル情報取得: ${fileInfo.name} (${fileInfo.type})`);

    // 翻訳先のファイル（空）を先に作成
    const targetFileInfo = fileHandler.createEmptyTranslatedFile(fileInfo, targetLang);
    log('INFO', `翻訳先ファイル作成: ${targetFileInfo.name}`);

    // 翻訳ジョブのリストを作成
    const jobs = fileHandler.createTranslationJobs(fileInfo);
    log('INFO', `翻訳ジョブを${jobs.length}件作成しました`);

    if (jobs.length === 0) {
      return { taskId: null, totalJobs: 0, targetFileUrl: targetFileInfo.url, message: '翻訳対象のテキストが見つかりませんでした。' };
    }

    // タスクIDを生成し、ジョブと設定をキャッシュに保存
    const taskId = `task_${new Date().getTime()}`;
    const cache = CacheService.getUserCache();
    const taskData = {
      jobs: jobs,
      totalJobs: jobs.length, // totalJobsを追加
      remainingJobs: jobs.length,
      targetLang: targetLang,
      dictName: dictName || '',
      targetFileId: targetFileInfo.id,
      originalFileType: fileInfo.type
    };
    cache.put(taskId, JSON.stringify(taskData), 21600); // 6時間有効

    return {
      taskId: taskId,
      totalJobs: jobs.length,
      targetFileUrl: targetFileInfo.url
    };

  } catch (error) {
    log('ERROR', '翻訳キューセットアップエラー', error);
    return { error: error.message };
  }
}

/**
 * キューから次のジョブを処理する（クライアントから繰り返し呼び出される）
 * @param {string} taskId - タスクID
 * @return {Object} 処理結果
 */
function processNextInQueue(taskId) {
  const cache = CacheService.getUserCache();
  const cachedData = cache.get(taskId);

  if (!cachedData) {
    log('ERROR', `タスクが見つかりません: ${taskId}`);
    return { status: 'error', message: 'タスクの有効期限が切れました。' };
  }

  try {
    const taskData = JSON.parse(cachedData);
    const { jobs, targetLang, dictName, targetFileId, originalFileType } = taskData;

    if (!jobs || jobs.length === 0) {
      cache.remove(taskId);
      return { status: 'complete', message: 'すべての翻訳が完了しました。' };
    }

    // ジョブを1つ取り出す
    const job = jobs.shift();
    
    if (!job) {
      log('WARN', `ジョブが不正です: ${taskId}`);
      cache.remove(taskId); // 不正な状態なのでキャッシュをクリア
      return { status: 'error', message: '翻訳ジョブが不正な状態です。' };
    }

    // テキストを翻訳
    const translator = new Translator(dictName);
    const translatedText = translator.translateText(job.text, targetLang);

    // 翻訳結果をファイルに書き込む
    const fileHandler = new FileHandler();
    fileHandler.writeTranslatedJob(targetFileId, originalFileType, job, translatedText);

    // キャッシュを更新
    taskData.remainingJobs = jobs.length;
    
    // 最後のジョブを処理した場合
    if (jobs.length === 0) {
      cache.remove(taskId); // タスク完了なのでキャッシュを削除
      log('INFO', `タスク完了: ${taskId}`);
      return {
        status: 'complete',
        completedJobs: taskData.totalJobs,
        totalJobs: taskData.totalJobs
      };
    }

    // まだジョブが残っている場合
    cache.put(taskId, JSON.stringify(taskData), 21600);
    log('INFO', `ジョブ処理完了: ${taskId} -残り${jobs.length}件`);

    return {
      status: 'processing',
      completedJobs: taskData.totalJobs - jobs.length,
      totalJobs: taskData.totalJobs
    };

  } catch (error) {
    log('ERROR', `ジョブ処理エラー: ${taskId}`, { error: error.message, stack: error.stack });
    // エラーが発生してもキュー処理を止めないために、エラー情報を返し、次の処理を促す
    return { 
      status: 'error', 
      message: `エラーが発生しました: ${error.message}`,
      completedJobs: taskData.totalJobs - jobs.length, // エラー発生時点での完了数
      totalJobs: taskData.totalJobs
    };
  }
}

/**
 * 辞書リストを取得（クライアントから呼び出される）
 * @return {Array} 辞書情報の配列
 */
function getDictionaryList() {
  try {
    const dictionary = new Dictionary();
    return dictionary.getList();
  } catch (error) {
    log('ERROR', '辞書リスト取得エラー', error);
    return [];
  }
}

/**
 * 新規辞書を作成（クライアントから呼び出される）
 * @param {string} name - 辞書名
 * @return {Object} 処理結果
 */
function createDictionary(name) {
  try {
    if (!name || name.trim() === '') {
      throw new Error('辞書名を入力してください');
    }
    
    const dictionary = new Dictionary();
    return dictionary.create(name.trim());
  } catch (error) {
    log('ERROR', '辞書作成エラー', error);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * 翻訳履歴を取得（クライアントから呼び出される）
 * @param {number} limit - 取得する履歴の最大数
 * @return {Array} 履歴データの配列
 */
function getTranslationHistory(limit = 50) {
  try {
    const history = new History();
    return history.getRecent(limit);
  } catch (error) {
    log('ERROR', '履歴取得エラー', error);
    return [];
  }
}

/**
 * 翻訳の進捗状況を取得（クライアントから呼び出される）
 * @param {string} taskId - タスクID
 * @return {Object} 進捗情報
 */
function getTranslationProgress(taskId) {
  // 将来的な実装用（現在は同期処理）
  return {
    status: 'completed',
    progress: 100
  };
}

/**
 * Google DriveのURLかどうかを検証
 * @param {string} url - 検証するURL
 * @return {boolean} 有効なGoogle DriveのURLかどうか
 */
function isValidGoogleDriveUrl(url) {
  const patterns = [
    /^https:\/\/docs\.google\.com\/spreadsheets\/d\//,
    /^https:\/\/docs\.google\.com\/document\/d\//,
    /^https:\/\/docs\.google\.com\/presentation\/d\//,
    /^https:\/\/drive\.google\.com\/file\/d\//,
    /^https:\/\/drive\.google\.com\/open\?id=/
  ];
  
  return patterns.some(pattern => pattern.test(url));
}

/**
 * URLからファイルIDを抽出
 * @param {string} url - Google DriveのURL
 * @return {string|null} ファイルID
 */
function extractFileId(url) {
  const patterns = [
    /\/d\/([a-zA-Z0-9-_]+)/,
    /id=([a-zA-Z0-9-_]+)/,
    /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/,
    /\/document\/d\/([a-zA-Z0-9-_]+)/,
    /\/presentation\/d\/([a-zA-Z0-9-_]+)/
  ];
  
  for (const pattern of patterns) {
    const match = url.match(pattern);
    if (match && match[1]) {
      return match[1];
    }
  }
  return null;
}

/**
 * 辞書データをエクスポート（クライアントから呼び出される）
 * @param {string} dictName - 辞書名
 * @return {Object} エクスポートデータ
 */
function exportDictionary(dictName) {
  try {
    const dictionary = new Dictionary();
    return dictionary.export(dictName);
  } catch (error) {
    log('ERROR', '辞書エクスポートエラー', error);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * 辞書データをインポート（クライアントから呼び出される）
 * @param {string} dictName - 辞書名
 * @param {Array} data - インポートするデータ
 * @return {Object} 処理結果
 */
function importDictionary(dictName, data) {
  try {
    const dictionary = new Dictionary();
    return dictionary.import(dictName, data);
  } catch (error) {
    log('ERROR', '辞書インポートエラー', error);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * システム情報を取得（デバッグ用）
 * @return {Object} システム情報
 */
function getSystemInfo() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties().getProperties();
    
    // スプレッドシートのURLを取得
    let historySheetUrl = '';
    let dictionarySheetUrl = '';
    const historySheetId = CONFIG.HISTORY_SHEET_ID; // getter経由で取得
    const dictionarySheetId = CONFIG.DICTIONARY_SHEET_ID; // getter経由で取得
    
    try {
      if (historySheetId) {
        historySheetUrl = SpreadsheetApp.openById(historySheetId).getUrl();
      }
    } catch (e) {
      log('WARN', '履歴シートURLの取得に失敗', e);
    }
    
    try {
      if (dictionarySheetId) {
        dictionarySheetUrl = SpreadsheetApp.openById(dictionarySheetId).getUrl();
      }
    } catch (e) {
      log('WARN', '辞書シートURLの取得に失敗', e);
    }
    
    return {
      hasApiKey: !!scriptProperties.OPENAI_API_KEY,
      hasOrgId: !!scriptProperties.OPENAI_ORGANIZATION_ID,
      hasProjectId: !!scriptProperties.OPENAI_PROJECT_ID,
      historySheetId: historySheetId || 'Not set',
      historySheetUrl: historySheetUrl,
      dictionarySheetId: dictionarySheetId || 'Not set',
      dictionarySheetUrl: dictionarySheetUrl,
      supportedLanguages: Object.keys(CONFIG.SUPPORTED_LANGUAGES).length,
      model: CONFIG.OPENAI_MODEL
    };
  } catch (error) {
    log('ERROR', 'システム情報取得エラー', error);
    return null;
  }
}

/**
 * テスト用関数（開発時のみ使用）
 */
function testTranslation() {
  // テスト用のダミーURL
  const testUrl = 'https://docs.google.com/spreadsheets/d/YOUR_TEST_FILE_ID/edit';
  const result = translateFile(testUrl, 'en', '');
  console.log(result);
}

/**
 * 初期セットアップ（管理者用）
 */
function setupApplication() {
  try {
    // 設定の初期化
    initializeConfig();
    
    // 必要な権限の事前承認
    // Drive API
    DriveApp.getRootFolder();
    
    // Sheets API
    SpreadsheetApp.create('temp').getId();
    
    // Docs API
    DocumentApp.create('temp').getId();
    
    // Slides API  
    SlidesApp.create('temp').getId();
    
    console.log('セットアップが完了しました');
    
    // システム情報を表示
    console.log(getSystemInfo());
    
  } catch (error) {
    console.error('セットアップエラー:', error);
  }
}
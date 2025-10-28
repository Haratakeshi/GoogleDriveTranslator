// Code.gs - メインのエントリーポイント

/**
 * WebアプリケーションのGETリクエストハンドラ
 * @return {HtmlOutput} HTMLページ
 */
function doGet() {
  try {
    // 初期化処理
    initializeConfig();
    
    const htmlTemplate = HtmlService.createTemplateFromFile('UI');
    htmlTemplate.faviconUrl = getFaviconDataUri(CONFIG.FAVICON_FILE_ID);

    return htmlTemplate.evaluate()
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

    const taskId = `task_${new Date().getTime()}`;
    const cache = CacheService.getUserCache();
    log('INFO', `タスクID生成: ${taskId}`);

    let targetFileId;
    let targetFileUrl = '';

    // 翻訳ジョブのリストを作成
    const jobs = fileHandler.createTranslationJobs(fileInfo);
    log('INFO', `[${taskId}] 翻訳ジョブを${jobs.length}件作成しました`);

    if (jobs.length === 0) {
      return { taskId: null, totalJobs: 0, message: '翻訳対象のテキストが見つかりませんでした。' };
    }

    // 処理方法の決定
    if (fileInfo.type === 'pdf') {
      // PDFの場合は、一時ドキュメントが翻訳先となる
      targetFileId = jobs[0].tempDocId; // どのジョブも同じIDを持っている
      // PDFの場合、最終的なURLは後処理で生成されるため、ここでは一時的なドキュメントのURLを返す
      targetFileUrl = `https://docs.google.com/document/d/${targetFileId}/edit`;
    } else {
      // 他形式の場合は、先に空のコピーファイルを作成する
      const targetFileInfo = fileHandler.createEmptyTranslatedFile(fileInfo, targetLang);
      targetFileId = targetFileInfo.id;
      targetFileUrl = targetFileInfo.url;
      log('INFO', `翻訳先ファイル作成: ${targetFileInfo.name}`);
    }

    // タスクデータをキャッシュに保存
    const taskData = {
      jobs: jobs,
      totalJobs: jobs.length,
      remainingJobs: jobs.length,
      targetLang: targetLang,
      dictName: dictName || '',
      targetFileId: targetFileId, // 動的に設定したIDを使用
      originalFileType: fileInfo.type,
      // 履歴記録用データ
      sourceUrl: fileUrl,
      targetUrl: targetFileUrl, // 動的に設定したURLを使用
      sourceLang: CONFIG.DEFAULT_SOURCE_LANG,
      totalSourceChars: 0,
      totalTargetChars: 0,
      totalDuration: 0,
      startTime: new Date().getTime()
    };
    cache.put(taskId, JSON.stringify(taskData), 21600); // 6時間有効
    log('INFO', `[${taskId}] タスクデータをキャッシュに保存しました。`);

    return {
      taskId: taskId,
      totalJobs: jobs.length,
      targetFileUrl: targetFileUrl
    };

  } catch (error) {
    log('ERROR', '翻訳キューセットアップエラー', { message: error.message, stack: error.stack });
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

  let taskData; // catchブロックでも参照できるようスコープを広げる

  try {
    taskData = JSON.parse(cachedData);
    const { jobs, targetLang, dictName, targetFileId, originalFileType } = taskData;
    const dictionaryCacheKey = `${taskId}_dictionary`;

    if (!jobs || jobs.length === 0) {
      cache.remove(taskId);
      cache.remove(dictionaryCacheKey); // 念のため辞書キャッシュも削除
      return { status: 'complete', message: 'すべての翻訳が完了しました。' };
    }

    // ジョブを1つ取り出す
    const job = jobs.shift();
    
    if (!job) {
      log('WARN', `ジョブが不正です: ${taskId}`);
      cache.remove(taskId); // 不正な状態なのでキャッシュをクリア
      cache.remove(dictionaryCacheKey);
      return { status: 'error', message: '翻訳ジョブが不正な状態です。' };
    }

    // 処理時間測定開始
    const jobStartTime = new Date().getTime();

    // 1件ずつの用語抽出・辞書照合・翻訳処理
    let confirmedPairs = [];
    let newTermCandidates = [];
    let translatedText = '';
    
    try {
      // 1. 現在のジョブのテキストから用語抽出
      if (job.text && job.text.trim().length > 0) {
        const termExtractor = new TermExtractor();
        const extractionResult = termExtractor.extract(job.text, 'auto', targetLang, '一般', {});
        const extractedTerms = extractionResult.extracted_terms || [];
        log('INFO', `[${taskId}] ジョブ用語抽出完了。抽出数: ${extractedTerms.length}`);

        if (extractedTerms.length > 0 && dictName) {
          // 2. 用語を現在の辞書と照合
          const termMatcher = new TermMatcher();
          const sourceTerms = extractedTerms.map(t => t.term);
          const matchResult = termMatcher.match(sourceTerms, dictName);
          confirmedPairs = matchResult.confirmedPairs || [];
          log('INFO', `[${taskId}] ジョブ用語照合完了。確定ペア数: ${confirmedPairs.length}`);
        }
      }

      // 3. 翻訳処理（確定ペアを使用）
      const translator = new Translator(dictName);
      const translationResult = translator.translateText(job.text, targetLang, confirmedPairs);
      translatedText = translationResult.translatedText;
      newTermCandidates = translationResult.newTermCandidates;
      log('INFO', `[${taskId}] 翻訳完了。新しい用語候補数: ${newTermCandidates.length}`);

    } catch (translationError) {
      log('ERROR', `[${taskId}] 翻訳処理でエラーが発生しました`, translationError);
      translatedText = job.text; // エラー時は原文をそのまま使用
    }

    // 翻訳結果をファイルに書き込む
    const fileHandler = new FileHandler();
    // PDF翻訳の場合は、一時的なGoogleドキュメントに書き込むため、documentハンドラを指定する
    const writerFileType = (originalFileType === 'pdf') ? 'document' : originalFileType;
    fileHandler.writeTranslatedJob(targetFileId, writerFileType, job, translatedText);

    // 処理時間測定終了と統計情報の更新
    const jobEndTime = new Date().getTime();
    const jobDuration = (jobEndTime - jobStartTime) / 1000; // 秒単位
    const sourceChars = job.text ? job.text.length : 0;
    const targetChars = translatedText ? translatedText.length : 0;
    
    // 統計情報を累積
    taskData.totalSourceChars += sourceChars;
    taskData.totalTargetChars += targetChars;
    taskData.totalDuration += jobDuration;

    // 新しい用語ペアの品質管理
    if (dictName && newTermCandidates && newTermCandidates.length > 0) {
      try {
        log('INFO', `[${taskId}] 新しい用語ペアの品質管理を開始します。`, { dictName: dictName, count: newTermCandidates.length });
        const qualityManager = new DictionaryQualityManager();
        const dictionary = new Dictionary();
        qualityManager.evaluateAndRegister(newTermCandidates, dictionary, dictName);
      } catch (e) {
        // 品質管理でエラーが発生しても処理は続行する
        log('ERROR', `[${taskId}] 用語の品質管理プロセスでエラーが発生しました。処理は続行されます。`, { message: e.message, stack: e.stack });
      }
    }

    // キャッシュを更新
    taskData.remainingJobs = jobs.length;
    
    // 最後のジョブを処理した場合
    if (jobs.length === 0) {
      let finalUrl = taskData.targetUrl; // デフォルトは元のURL

      // PDF翻訳の場合、最終処理を実行し、最終URLを取得
      if (job.handlerType === 'PdfHandler') {
        const pdfHandler = new PdfHandler();
        finalUrl = pdfHandler.finalizeTranslation(job.tempDocId, job.originalPdfId) || finalUrl;
      }

      // 翻訳履歴を記録
      try {
        const history = new History();
        const totalDuration = taskData.totalDuration + ((new Date().getTime() - taskData.startTime) / 1000);
        const historyData = {
          sourceUrl: taskData.sourceUrl || '',
          targetUrl: finalUrl, // 最終的なURLを履歴に記録
          sourceLang: taskData.sourceLang || CONFIG.DEFAULT_SOURCE_LANG,
          targetLang: targetLang,
          dictName: dictName,
          charCountSource: taskData.totalSourceChars || 0,
          charCountTarget: taskData.totalTargetChars || 0,
          duration: Math.round(totalDuration),
          status: 'success'
        };
        history.record(historyData);
        log('INFO', `[${taskId}] 翻訳履歴を記録しました`);
      } catch (e) {
        log('ERROR', `[${taskId}] 翻訳履歴の記録に失敗しました`, { message: e.message, stack: e.stack });
      }
      
      cache.remove(taskId); // タスク完了なのでキャッシュを削除
      cache.remove(dictionaryCacheKey); // 辞書キャッシュも削除
      log('INFO', `タスク完了: ${taskId}`);
      return {
        status: 'complete',
        completedJobs: taskData.totalJobs,
        totalJobs: taskData.totalJobs,
        targetFileUrl: finalUrl // UIに最終的なURLを返す
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
    
    // エラーが発生した場合も履歴を記録
    if (taskData) {
      try {
        const history = new History();
        const totalDuration = taskData.totalDuration + ((new Date().getTime() - taskData.startTime) / 1000);
        const historyData = {
          sourceUrl: taskData.sourceUrl || '',
          targetUrl: taskData.targetUrl || '',
          sourceLang: taskData.sourceLang || CONFIG.DEFAULT_SOURCE_LANG,
          targetLang: taskData.targetLang || '',
          dictName: taskData.dictName || '',
          charCountSource: taskData.totalSourceChars || 0,
          charCountTarget: taskData.totalTargetChars || 0,
          duration: Math.round(totalDuration),
          status: 'error',
          errorMessage: error.message
        };
        history.record(historyData);
        log('INFO', `[${taskId}] エラー履歴を記録しました`);
      } catch (e) {
        log('ERROR', `[${taskId}] エラー履歴の記録に失敗しました`, { message: e.message, stack: e.stack });
      }
      
      // キャッシュをクリア
      cache.remove(taskId);
      cache.remove(`${taskId}_dictionary`);
    }
    
    // エラーが発生してもキュー処理を止めないために、エラー情報を返し、次の処理を促す
    const completedJobs = taskData ? (taskData.totalJobs - taskData.jobs.length) : 'N/A';
    const totalJobs = taskData ? taskData.totalJobs : 'N/A';
    return {
      status: 'error',
      message: `エラーが発生しました: ${error.message}`,
      completedJobs: completedJobs,
      totalJobs: totalJobs
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
 * 単一ファイル翻訳のテスト用関数（開発時のみ使用）
 */
function testSingleFileTranslation() {
  // テスト用のダミーURL
  const testUrl = 'https://docs.google.com/spreadsheets/d/YOUR_TEST_FILE_ID/edit';
  const result = setupTranslationQueue(testUrl, 'en', '');
  console.log('単一ファイル翻訳テスト結果:', result);
  return result;
}

/**
 * バッチ処理のテスト用関数（開発時のみ使用）
 */
function testBatchTranslation() {
  // テスト用のダミーURL配列
  const testUrls = [
    'https://docs.google.com/spreadsheets/d/YOUR_TEST_FILE_ID_1/edit',
    'https://docs.google.com/document/d/YOUR_TEST_FILE_ID_2/edit'
  ];
  
  const setupResult = setupBatchTranslation(testUrls, 'en', '', 'テストバッチ');
  console.log('バッチ翻訳セットアップテスト結果:', setupResult);
  
  if (setupResult.batchId) {
    const startResult = startBatchTranslation(setupResult.batchId);
    console.log('バッチ翻訳開始テスト結果:', startResult);
    
    const statusResult = getBatchProgress(setupResult.batchId);
    console.log('バッチ進行状況テスト結果:', statusResult);
  }
  
  return setupResult;
}

/**
 * 既存機能との互換性テスト用関数（開発時のみ使用）
 */
function testCompatibility() {
  console.log('=== 既存機能との互換性テスト開始 ===');
  
  const results = {
    singleFileTranslation: null,
    batchTranslation: null,
    dictionaryFunctions: null,
    historyFunctions: null,
    systemInfo: null
  };
  
  try {
    // 1. 単一ファイル翻訳機能のテスト
    console.log('1. 単一ファイル翻訳機能のテスト');
    results.singleFileTranslation = {
      setupQueue: typeof setupTranslationQueue === 'function',
      processNext: typeof processNextInQueue === 'function',
      getProgress: typeof getTranslationProgress === 'function'
    };
    console.log('単一ファイル翻訳機能:', results.singleFileTranslation);
    
    // 2. バッチ翻訳機能のテスト
    console.log('2. バッチ翻訳機能のテスト');
    results.batchTranslation = {
      setupBatch: typeof setupBatchTranslation === 'function',
      startBatch: typeof startBatchTranslation === 'function',
      processNextInBatch: typeof processNextInBatch === 'function',
      getBatchProgress: typeof getBatchProgress === 'function',
      pauseBatch: typeof pauseBatchTranslation === 'function',
      resumeBatch: typeof resumeBatchTranslation === 'function'
    };
    console.log('バッチ翻訳機能:', results.batchTranslation);
    
    // 3. 辞書機能のテスト
    console.log('3. 辞書機能のテスト');
    results.dictionaryFunctions = {
      getDictionaryList: typeof getDictionaryList === 'function',
      createDictionary: typeof createDictionary === 'function',
      exportDictionary: typeof exportDictionary === 'function',
      importDictionary: typeof importDictionary === 'function'
    };
    console.log('辞書機能:', results.dictionaryFunctions);
    
    // 4. 履歴機能のテスト
    console.log('4. 履歴機能のテスト');
    results.historyFunctions = {
      getTranslationHistory: typeof getTranslationHistory === 'function',
      getBatchHistory: typeof getBatchHistory === 'function',
      getBatchFileHistory: typeof getBatchFileHistory === 'function',
      getBatchStatistics: typeof getBatchStatistics === 'function'
    };
    console.log('履歴機能:', results.historyFunctions);
    
    // 5. システム情報のテスト
    console.log('5. システム情報のテスト');
    const systemInfo = getSystemInfo();
    results.systemInfo = {
      hasApiKey: systemInfo ? systemInfo.hasApiKey : false,
      hasHistorySheet: systemInfo ? !!systemInfo.historySheetId : false,
      hasDictionarySheet: systemInfo ? !!systemInfo.dictionarySheetId : false,
      supportedLanguages: systemInfo ? systemInfo.supportedLanguages : 0
    };
    console.log('システム情報:', results.systemInfo);
    
    // 6. クラスのインスタンス化テスト
    console.log('6. クラスのインスタンス化テスト');
    try {
      const batchManager = new BatchTranslationManager();
      const queueManager = new QueueManager();
      const batchHistory = new BatchHistory();
      console.log('クラスインスタンス化: 成功');
      results.classInstantiation = true;
    } catch (error) {
      console.error('クラスインスタンス化エラー:', error);
      results.classInstantiation = false;
    }
    
    console.log('=== 互換性テスト完了 ===');
    console.log('テスト結果サマリー:', results);
    
    return results;
    
  } catch (error) {
    console.error('互換性テスト中にエラーが発生:', error);
    return { error: error.message, results: results };
  }
}

/**
 * バッチ処理システムの総合テスト（開発時のみ使用）
 */
function testBatchSystem() {
  console.log('=== バッチ処理システム総合テスト開始 ===');
  
  try {
    // 1. 互換性テスト
    console.log('1. 互換性テスト実行中...');
    const compatibilityResult = testCompatibility();
    
    // 2. バッチマネージャーのテスト
    console.log('2. バッチマネージャーテスト実行中...');
    const batchManager = new BatchTranslationManager();
    console.log('バッチマネージャー作成: 成功');
    
    // 3. キューマネージャーのテスト
    console.log('3. キューマネージャーテスト実行中...');
    const queueManager = new QueueManager();
    const queueInit = queueManager.initializeQueue();
    console.log('キューマネージャー初期化:', queueInit.status);
    
    // 4. バッチ履歴のテスト
    console.log('4. バッチ履歴テスト実行中...');
    const batchHistory = new BatchHistory();
    const historyStats = batchHistory.getBatchStatistics();
    console.log('バッチ履歴統計取得: 成功');
    
    // 5. 設定値の確認
    console.log('5. 設定値確認中...');
    const configTest = {
      maxBatchSize: CONFIG.BATCH_PROCESSING.MAX_BATCH_SIZE,
      maxRetryAttempts: CONFIG.BATCH_PROCESSING.MAX_RETRY_ATTEMPTS,
      parallelLimit: CONFIG.BATCH_PROCESSING.PARALLEL_PROCESSING_LIMIT,
      resumeTokenExpiry: CONFIG.BATCH_PROCESSING.RESUME_TOKEN_EXPIRY
    };
    console.log('バッチ処理設定:', configTest);
    
    console.log('=== バッチ処理システム総合テスト完了 ===');
    
    return {
      status: 'success',
      compatibility: compatibilityResult,
      configTest: configTest,
      message: 'すべてのテストが正常に完了しました'
    };
    
  } catch (error) {
    console.error('バッチ処理システムテスト中にエラーが発生:', error);
    return {
      status: 'error',
      error: error.message,
      stack: error.stack
    };
  }
}

// === バッチ処理API関数 ===

/**
 * バッチ翻訳をセットアップする（クライアントから呼び出される）
 * @param {Array} fileUrls - ファイルURLの配列
 * @param {string} targetLang - 翻訳先言語コード
 * @param {string} dictName - 使用する辞書名（オプション）
 * @param {string} batchName - バッチ名（オプション）
 * @return {Object} バッチ情報
 */
function setupBatchTranslation(fileUrls, targetLang, dictName = '', batchName = '') {
  log('INFO', `バッチ翻訳セットアップ開始: ${fileUrls.length}ファイル, 言語=${targetLang}, 辞書=${dictName}`);

  try {
    if (!fileUrls || !Array.isArray(fileUrls) || fileUrls.length === 0) {
      throw new Error('ファイルURLの配列が必要です');
    }

    if (fileUrls.length > CONFIG.BATCH_PROCESSING.MAX_BATCH_SIZE) {
      throw new Error(`バッチサイズが上限を超えています（最大${CONFIG.BATCH_PROCESSING.MAX_BATCH_SIZE}ファイル）`);
    }

    if (!targetLang || !CONFIG.SUPPORTED_LANGUAGES[targetLang]) {
      throw new Error('有効な翻訳先言語を指定してください');
    }

    const batchManager = new BatchTranslationManager();
    const result = batchManager.createBatch(fileUrls, targetLang, dictName, batchName);

    if (result.error) {
      throw new Error(result.error);
    }

    log('INFO', `バッチ翻訳セットアップ完了: バッチID=${result.batchId}, 有効ファイル=${result.validFiles}件`);

    return result;

  } catch (error) {
    log('ERROR', 'バッチ翻訳セットアップエラー', error);
    return { error: error.message };
  }
}

/**
 * バッチ翻訳を開始する（クライアントから呼び出される）
 * @param {string} batchId - バッチID
 * @return {Object} 開始結果
 */
function startBatchTranslation(batchId) {
  log('INFO', `バッチ翻訳開始: ${batchId}`);

  try {
    if (!batchId) {
      throw new Error('バッチIDが必要です');
    }

    const batchManager = new BatchTranslationManager();
    const result = batchManager.startBatch(batchId);

    if (result.status === 'error') {
      throw new Error(result.message);
    }

    log('INFO', `バッチ翻訳開始完了: ${batchId}`);

    return result;

  } catch (error) {
    log('ERROR', 'バッチ翻訳開始エラー', error);
    return { error: error.message };
  }
}

/**
 * バッチ内の次のファイルを処理する（クライアントから繰り返し呼び出される）
 * @param {string} batchId - バッチID
 * @return {Object} 処理結果
 */
function processNextInBatch(batchId) {
  try {
    if (!batchId) {
      return { status: 'error', message: 'バッチIDが必要です' };
    }

    const batchManager = new BatchTranslationManager();
    const result = batchManager.processNextFile(batchId);

    // ファイル処理が開始された場合、実際の翻訳処理を行う
    if (result.status === 'file_processing' && result.taskId) {
      // 既存の processNextInQueue を使用して実際の翻訳を行う
      const translationResult = processNextInQueue(result.taskId);
      
      if (translationResult.status === 'complete' || translationResult.status === 'error') {
        // ファイル処理完了をバッチマネージャーに通知
        const completionResult = batchManager.onFileCompleted(batchId, result.taskId, translationResult);
        
        // バッチ全体が完了した場合
        if (completionResult.status === 'completed') {
          return completionResult;
        }
        
        // 個別ファイル完了の場合、次のファイル処理情報も含める
        return {
          ...completionResult,
          continueProcessing: completionResult.remainingFiles > 0
        };
      }
      
      // 翻訳処理中の場合
      return {
        status: 'processing',
        batchId: batchId,
        fileName: result.fileName,
        taskId: result.taskId,
        completedJobs: translationResult.completedJobs || 0,
        totalJobs: translationResult.totalJobs || result.totalJobs,
        processedFiles: result.processedFiles,
        totalFiles: result.totalFiles
      };
    }

    return result;

  } catch (error) {
    log('ERROR', 'バッチファイル処理エラー', error);
    return { status: 'error', message: error.message };
  }
}

/**
 * バッチ処理の進行状況を取得する（クライアントから呼び出される）
 * @param {string} batchId - バッチID
 * @return {Object} 進行状況情報
 */
function getBatchProgress(batchId) {
  try {
    if (!batchId) {
      return { status: 'error', message: 'バッチIDが必要です' };
    }

    const batchManager = new BatchTranslationManager();
    return batchManager.getBatchStatus(batchId);

  } catch (error) {
    log('ERROR', 'バッチ進行状況取得エラー', error);
    return { status: 'error', message: error.message };
  }
}

/**
 * バッチ処理を一時停止する（クライアントから呼び出される）
 * @param {string} batchId - バッチID
 * @return {Object} 一時停止結果
 */
function pauseBatchTranslation(batchId) {
  log('INFO', `バッチ翻訳一時停止: ${batchId}`);

  try {
    if (!batchId) {
      throw new Error('バッチIDが必要です');
    }

    const batchManager = new BatchTranslationManager();
    const result = batchManager.pauseBatch(batchId);

    if (result.status === 'error') {
      throw new Error(result.message);
    }

    log('INFO', `バッチ翻訳一時停止完了: ${batchId}`);

    return result;

  } catch (error) {
    log('ERROR', 'バッチ翻訳一時停止エラー', error);
    return { error: error.message };
  }
}

/**
 * バッチ処理を再開する（クライアントから呼び出される）
 * @param {string} batchId - バッチID
 * @return {Object} 再開結果
 */
function resumeBatchTranslation(batchId) {
  log('INFO', `バッチ翻訳再開: ${batchId}`);

  try {
    if (!batchId) {
      throw new Error('バッチIDが必要です');
    }

    const batchManager = new BatchTranslationManager();
    const result = batchManager.resumeBatch(batchId);

    if (result.status === 'error') {
      throw new Error(result.message);
    }

    log('INFO', `バッチ翻訳再開完了: ${batchId}`);

    return result;

  } catch (error) {
    log('ERROR', 'バッチ翻訳再開エラー', error);
    return { error: error.message };
  }
}

/**
 * バッチ履歴を取得する（クライアントから呼び出される）
 * @param {Object} options - 取得オプション
 * @return {Array} バッチ履歴の配列
 */
function getBatchHistory(options = {}) {
  try {
    const batchHistory = new BatchHistory();
    return batchHistory.getBatchHistory(options);

  } catch (error) {
    log('ERROR', 'バッチ履歴取得エラー', error);
    return [];
  }
}

/**
 * バッチのファイル履歴を取得する（クライアントから呼び出される）
 * @param {string} batchId - バッチID
 * @return {Array} ファイル履歴の配列
 */
function getBatchFileHistory(batchId) {
  try {
    if (!batchId) {
      return [];
    }

    const batchHistory = new BatchHistory();
    return batchHistory.getBatchFileHistory(batchId);

  } catch (error) {
    log('ERROR', 'バッチファイル履歴取得エラー', error);
    return [];
  }
}

/**
 * バッチ統計情報を取得する（クライアントから呼び出される）
 * @param {string} batchId - バッチID（省略時は全体統計）
 * @return {Object} 統計情報
 */
function getBatchStatistics(batchId = null) {
  try {
    const batchHistory = new BatchHistory();
    return batchHistory.getBatchStatistics(batchId);

  } catch (error) {
    log('ERROR', 'バッチ統計取得エラー', error);
    return {
      totalBatches: 0,
      completedBatches: 0,
      failedBatches: 0,
      processingBatches: 0,
      totalFiles: 0,
      completedFiles: 0,
      failedFiles: 0
    };
  }
}

/**
 * キュー統計情報を取得する（クライアントから呼び出される）
 * @return {Object} キュー統計情報
 */
function getQueueStatistics() {
  try {
    const queueManager = new QueueManager();
    return queueManager.getQueueStatistics();

  } catch (error) {
    log('ERROR', 'キュー統計取得エラー', error);
    return {
      totalTasks: 0,
      queuedTasks: 0,
      activeTasks: 0,
      completedTasks: 0,
      failedTasks: 0,
      concurrencyUtilization: 0
    };
  }
}

/**
 * 古いタスクとバッチデータをクリーンアップする（管理者用）
 * @param {number} maxAgeHours - 最大保持時間（時間）
 * @return {Object} クリーンアップ結果
 */
function cleanupOldBatchData(maxAgeHours = 24) {
  log('INFO', `古いバッチデータのクリーンアップを開始: ${maxAgeHours}時間以上経過したデータ`);

  try {
    const results = {
      queueCleaned: 0,
      resumeInfoCleaned: 0,
      totalCleaned: 0
    };

    // キューのクリーンアップ
    try {
      const queueManager = new QueueManager();
      const queueResult = queueManager.cleanupOldTasks(maxAgeHours);
      results.queueCleaned = queueResult.cleaned || 0;
    } catch (error) {
      log('WARN', 'キュークリーンアップでエラー', error);
    }

    // バッチ履歴の再開情報クリーンアップ
    try {
      const batchHistory = new BatchHistory();
      const resumeResult = batchHistory.cleanupExpiredResumeInfo();
      results.resumeInfoCleaned = resumeResult || 0;
    } catch (error) {
      log('WARN', '再開情報クリーンアップでエラー', error);
    }

    results.totalCleaned = results.queueCleaned + results.resumeInfoCleaned;

    log('INFO', `バッチデータクリーンアップ完了: 合計${results.totalCleaned}件削除`);

    return {
      status: 'completed',
      ...results,
      message: `${results.totalCleaned}件のデータをクリーンアップしました`
    };

  } catch (error) {
    log('ERROR', 'バッチデータクリーンアップエラー', error);
    return {
      status: 'error',
      message: error.message
    };
  }
}

/**
 * バッチ処理のヘルスチェックを実行する（クライアントから呼び出される）
 * @param {string} batchId - バッチID
 * @return {Object} ヘルスチェック結果
 */
function performBatchHealthCheck(batchId) {
  try {
    if (!batchId) {
      return { status: 'error', message: 'バッチIDが必要です' };
    }

    const batchManager = new BatchTranslationManager();
    return batchManager.performHealthCheck(batchId);

  } catch (error) {
    log('ERROR', 'バッチヘルスチェックエラー', error);
    return {
      status: 'error',
      message: error.message
    };
  }
}

/**
 * バッチ処理の自動回復を試行する（クライアントから呼び出される）
 * @param {string} batchId - バッチID
 * @return {Object} 回復試行結果
 */
function attemptBatchAutoRecovery(batchId) {
  log('INFO', `バッチ自動回復試行: ${batchId}`);

  try {
    if (!batchId) {
      return { status: 'error', message: 'バッチIDが必要です' };
    }

    const batchManager = new BatchTranslationManager();
    const result = batchManager.attemptAutoRecovery(batchId);

    log('INFO', `バッチ自動回復完了: ${batchId}, ステータス=${result.status}`);

    return result;

  } catch (error) {
    log('ERROR', 'バッチ自動回復エラー', error);
    return {
      status: 'error',
      message: error.message
    };
  }
}

/**
 * 失敗したファイルを手動で再試行する（クライアントから呼び出される）
 * @param {string} batchId - バッチID
 * @param {Array} fileIndices - 再試行するファイルのインデックス配列
 * @return {Object} 再試行結果
 */
function retryFailedFiles(batchId, fileIndices) {
  log('INFO', `失敗ファイル手動再試行: ${batchId}, ファイル数=${fileIndices.length}`);

  try {
    if (!batchId) {
      return { status: 'error', message: 'バッチIDが必要です' };
    }

    if (!fileIndices || !Array.isArray(fileIndices) || fileIndices.length === 0) {
      return { status: 'error', message: 'ファイルインデックスの配列が必要です' };
    }

    const batchManager = new BatchTranslationManager();
    const batchData = batchManager.getBatchData(batchId);
    
    if (!batchData) {
      return { status: 'error', message: 'バッチが見つかりません' };
    }

    let retriedCount = 0;
    const errors = [];

    fileIndices.forEach(index => {
      if (index >= 0 && index < batchData.files.length) {
        const file = batchData.files[index];
        
        if (file.status === 'failed' && file.retryCount < CONFIG.BATCH_PROCESSING.MAX_RETRY_ATTEMPTS) {
          file.status = 'retrying';
          file.retryCount = (file.retryCount || 0) + 1;
          file.errorMessage = '手動再試行';
          
          // 失敗数を調整
          if (batchData.failedFiles > 0) {
            batchData.failedFiles--;
          }
          
          retriedCount++;
          log('INFO', `[${batchId}] ファイル手動再試行: ${file.fileName} (インデックス=${index})`);
        } else {
          errors.push(`ファイル${index}は再試行できません (ステータス=${file.status}, 再試行回数=${file.retryCount})`);
        }
      } else {
        errors.push(`無効なファイルインデックス: ${index}`);
      }
    });

    if (retriedCount > 0) {
      batchData.lastUpdated = new Date().getTime();
      batchManager.saveBatchData(batchId, batchData);
    }

    log('INFO', `失敗ファイル手動再試行完了: ${batchId}, 再試行=${retriedCount}件, エラー=${errors.length}件`);

    return {
      status: retriedCount > 0 ? 'success' : 'no_files_retried',
      batchId: batchId,
      retriedCount: retriedCount,
      requestedCount: fileIndices.length,
      errors: errors,
      message: `${retriedCount}件のファイルを再試行に設定しました`
    };

  } catch (error) {
    log('ERROR', '失敗ファイル手動再試行エラー', error);
    return {
      status: 'error',
      message: error.message
    };
  }
}

/**
 * バッチ処理をキャンセルする（クライアントから呼び出される）
 * @param {string} batchId - バッチID
 * @param {string} reason - キャンセル理由（オプション）
 * @return {Object} キャンセル結果
 */
function cancelBatchTranslation(batchId, reason = '') {
  log('INFO', `バッチ翻訳キャンセル: ${batchId}, 理由=${reason}`);

  try {
    if (!batchId) {
      return { status: 'error', message: 'バッチIDが必要です' };
    }

    const batchManager = new BatchTranslationManager();
    const batchData = batchManager.getBatchData(batchId);
    
    if (!batchData) {
      return { status: 'error', message: 'バッチが見つかりません' };
    }

    if (batchData.status === 'completed') {
      return { status: 'error', message: 'バッチは既に完了しています' };
    }

    if (batchData.status === 'cancelled') {
      return { status: 'already_cancelled', message: 'バッチは既にキャンセルされています' };
    }

    // 処理中のファイルをキャンセル状態に変更
    let cancelledFiles = 0;
    batchData.files.forEach(file => {
      if (file.status === 'pending' || file.status === 'processing' || file.status === 'retrying') {
        file.status = 'cancelled';
        file.errorMessage = reason || 'ユーザーによるキャンセル';
        file.completedTime = new Date().getTime();
        cancelledFiles++;
      }
    });

    batchData.status = 'cancelled';
    batchData.cancelledAt = new Date().getTime();
    batchData.cancelReason = reason || 'ユーザーによるキャンセル';
    batchData.lastUpdated = new Date().getTime();

    // バッチ履歴を更新
    const batchHistory = new BatchHistory();
    batchHistory.updateBatch(batchId, {
      status: 'cancelled',
      errorMessage: reason || 'ユーザーによるキャンセル'
    });

    batchManager.saveBatchData(batchId, batchData);

    log('INFO', `バッチ翻訳キャンセル完了: ${batchId}, キャンセルファイル=${cancelledFiles}件`);

    return {
      status: 'cancelled',
      batchId: batchId,
      cancelledFiles: cancelledFiles,
      completedFiles: batchData.completedFiles,
      failedFiles: batchData.failedFiles,
      reason: reason || 'ユーザーによるキャンセル',
      message: `バッチ処理をキャンセルしました。${cancelledFiles}件のファイルがキャンセルされました。`
    };

  } catch (error) {
    log('ERROR', 'バッチキャンセルエラー', error);
    return {
      status: 'error',
      message: error.message
    };
  }
}

// --- バッチ処理用API関数（UIから呼び出される） ---

/**
 * バッチ処理を作成する（UIから呼び出される）
 * @param {Array} fileUrls - ファイルURLの配列
 * @param {string} targetLang - 翻訳先言語コード
 * @param {string} dictName - 使用する辞書名（オプション）
 * @param {string} batchName - バッチ名（オプション）
 * @return {Object} バッチ情報
 */
function createBatch(fileUrls, targetLang, dictName, batchName) {
  try {
    const batchManager = new BatchTranslationManager();
    return batchManager.createBatch(fileUrls, targetLang, dictName, batchName);
  } catch (error) {
    log('ERROR', 'createBatch API error', error);
    return { error: error.message };
  }
}

/**
 * バッチ処理を開始する（UIから呼び出される）
 * @param {string} batchId - バッチID
 * @return {Object} 処理結果
 */
function startBatch(batchId) {
  try {
    const batchManager = new BatchTranslationManager();
    return batchManager.startBatch(batchId);
  } catch (error) {
    log('ERROR', 'startBatch API error', error);
    return { status: 'error', message: error.message };
  }
}

/**
 * バッチ処理のステータスを取得する（UIから呼び出される）
 * @param {string} batchId - バッチID
 * @return {Object} ステータス情報
 */
function getBatchStatus(batchId) {
  try {
    const batchManager = new BatchTranslationManager();
    return batchManager.getBatchStatus(batchId);
  } catch (error) {
    log('ERROR', 'getBatchStatus API error', error);
    return { status: 'error', message: error.message };
  }
}

/**
 * バッチ処理を一時停止する（UIから呼び出される）
 * @param {string} batchId - バッチID
 * @return {Object} 処理結果
 */
function pauseBatch(batchId) {
  try {
    const batchManager = new BatchTranslationManager();
    return batchManager.pauseBatch(batchId);
  } catch (error) {
    log('ERROR', 'pauseBatch API error', error);
    return { status: 'error', message: error.message };
  }
}

/**
 * バッチ処理を再開する（UIから呼び出される）
 * @param {string} batchId - バッチID
 * @return {Object} 処理結果
 */
function resumeBatch(batchId) {
  try {
    const batchManager = new BatchTranslationManager();
    return batchManager.resumeBatch(batchId);
  } catch (error) {
    log('ERROR', 'resumeBatch API error', error);
    return { status: 'error', message: error.message };
  }
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
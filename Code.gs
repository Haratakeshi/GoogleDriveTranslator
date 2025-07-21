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
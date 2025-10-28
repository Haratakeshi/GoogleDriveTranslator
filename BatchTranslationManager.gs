// BatchTranslationManager.gs - バッチ処理の中核クラス

/**
 * 複数ファイルの翻訳を管理するバッチ処理システム
 * 既存の単一ファイル処理との互換性を保ちながら、複数ファイルの処理を統括
 */
class BatchTranslationManager {
  constructor() {
    this.batchHistory = new BatchHistory();
    this.cache = CacheService.getUserCache();
    this.maxConcurrentFiles = CONFIG.BATCH_PROCESSING.PARALLEL_PROCESSING_LIMIT;
    this.statusUpdateInterval = CONFIG.BATCH_PROCESSING.STATUS_UPDATE_INTERVAL;
  }

  /**
   * バッチ処理を作成する
   * @param {Array} fileUrls - ファイルURLの配列
   * @param {string} targetLang - 翻訳先言語コード
   * @param {string} dictName - 使用する辞書名（オプション）
   * @param {string} batchName - バッチ名（オプション）
   * @return {Object} バッチ情報
   */
  createBatch(fileUrls, targetLang, dictName = '', batchName = '') {
    log('INFO', `バッチ処理作成開始: ${fileUrls.length}ファイル, 言語=${targetLang}, 辞書=${dictName}`);

    try {
      // バッチデータの初期化
      const batchData = {
        totalFiles: fileUrls.length,
        targetLang: targetLang,
        dictName: dictName,
        settings: {
          targetLang: targetLang,
          dictName: dictName,
          createdAt: new Date().getTime()
        },
        batchName: batchName || `Batch_${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd_HH-mm')}`
      };

      // バッチ履歴にバッチを作成
      const batchId = this.batchHistory.createBatch(batchData);

      // バッチファイル情報の配列を作成
      const batchFiles = [];
      const fileErrors = [];

      fileUrls.forEach((fileUrl, index) => {
        try {
          if (!fileUrl || !this.isValidGoogleDriveUrl(fileUrl)) {
            throw new Error('無効なGoogle DriveのURLです');
          }

          const fileId = this.extractFileId(fileUrl);
          if (!fileId) {
            throw new Error('ファイルIDを抽出できませんでした');
          }

          const fileHandler = new FileHandler();
          const fileInfo = fileHandler.getFileInfo(fileId);

          batchFiles.push({
            index: index,
            sourceUrl: fileUrl,
            fileId: fileId,
            fileName: fileInfo.name,
            fileType: fileInfo.type,
            status: 'pending',
            errorMessage: '',
            retryCount: 0
          });

          log('INFO', `[${batchId}] ファイル${index + 1}/${fileUrls.length}: ${fileInfo.name} (${fileInfo.type})`);

        } catch (error) {
          log('ERROR', `[${batchId}] ファイル${index + 1}の情報取得エラー: ${fileUrl}`, error);
          fileErrors.push({
            index: index,
            sourceUrl: fileUrl,
            error: error.message
          });
        }
      });

      // バッチ処理データをキャッシュに保存
      const batchProcessData = {
        batchId: batchId,
        batchName: batchData.batchName,
        targetLang: targetLang,
        dictName: dictName,
        status: 'pending',
        totalFiles: fileUrls.length,
        validFiles: batchFiles.length,
        invalidFiles: fileErrors.length,
        processedFiles: 0,
        completedFiles: 0,
        failedFiles: 0,
        currentFileIndex: 0,
        files: batchFiles,
        errors: fileErrors,
        startTime: new Date().getTime(),
        lastUpdated: new Date().getTime(),
        pausedAt: null,
        resumeData: null
      };

      this.cache.put(`batch_${batchId}`, JSON.stringify(batchProcessData), 21600); // 6時間有効
      log('INFO', `[${batchId}] バッチ処理データをキャッシュに保存: 有効ファイル${batchFiles.length}件, 無効ファイル${fileErrors.length}件`);

      return {
        batchId: batchId,
        totalFiles: fileUrls.length,
        validFiles: batchFiles.length,
        invalidFiles: fileErrors.length,
        fileErrors: fileErrors,
        status: 'created'
      };

    } catch (error) {
      log('ERROR', 'バッチ処理作成エラー', error);
      return { error: error.message };
    }
  }

  /**
   * バッチ処理を開始する
   * @param {string} batchId - バッチID
   * @return {Object} 処理結果
   */
  startBatch(batchId) {
    log('INFO', `バッチ処理開始: ${batchId}`);

    try {
      const batchData = this.getBatchData(batchId);
      if (!batchData) {
        return { status: 'error', message: 'バッチが見つかりません' };
      }

      if (batchData.status === 'processing') {
        return { status: 'already_running', message: 'バッチは既に処理中です' };
      }

      // バッチステータスを処理中に更新
      batchData.status = 'processing';
      batchData.currentFileIndex = 0;
      batchData.lastUpdated = new Date().getTime();

      // バッチ履歴を更新
      this.batchHistory.updateBatch(batchId, {
        status: 'processing'
      });

      this.saveBatchData(batchId, batchData);

      log('INFO', `[${batchId}] バッチ処理を開始しました: ${batchData.validFiles}ファイル`);

      return {
        status: 'started',
        batchId: batchId,
        totalFiles: batchData.validFiles,
        message: 'バッチ処理を開始しました'
      };

    } catch (error) {
      log('ERROR', `バッチ処理開始エラー: ${batchId}`, error);
      return { status: 'error', message: error.message };
    }
  }

  /**
   * バッチ内の次のファイルを処理する
   * @param {string} batchId - バッチID
   * @return {Object} 処理結果
   */
  processNextFile(batchId) {
    try {
      const batchData = this.getBatchData(batchId);
      if (!batchData) {
        return { status: 'error', message: 'バッチが見つかりません' };
      }

      if (batchData.status === 'paused') {
        return { status: 'paused', message: 'バッチ処理は一時停止中です' };
      }

      if (batchData.status === 'completed') {
        return { status: 'already_completed', message: 'バッチ処理は既に完了しています' };
      }

      // 処理すべきファイルを検索
      const pendingFiles = batchData.files.filter(file => file.status === 'pending' || file.status === 'retrying');
      
      if (pendingFiles.length === 0) {
        // 全ファイル処理完了
        return this.completeBatch(batchId);
      }

      // 次のファイルを取得
      const nextFile = pendingFiles[0];
      nextFile.status = 'processing';
      nextFile.startTime = new Date().getTime();

      // バッチ履歴にファイル処理記録を作成
      const jobId = this.batchHistory.recordFileJob({
        batchId: batchId,
        sourceUrl: nextFile.sourceUrl,
        fileName: nextFile.fileName,
        fileType: nextFile.fileType,
        status: 'processing',
        startTime: new Date()
      });

      nextFile.jobId = jobId;

      this.saveBatchData(batchId, batchData);

      log('INFO', `[${batchId}] ファイル処理開始: ${nextFile.fileName} (${nextFile.index + 1}/${batchData.totalFiles})`);

      // 既存の単一ファイル処理システムを使用
      const setupResult = this.setupSingleFileTranslation(nextFile.sourceUrl, batchData.targetLang, batchData.dictName);
      
      if (setupResult.error) {
        // ファイル処理失敗 - 再試行判定
        const shouldRetry = this.shouldRetryFile(nextFile, setupResult.error);
        
        if (shouldRetry) {
          nextFile.status = 'retrying';
          nextFile.retryCount++;
          nextFile.errorMessage = setupResult.error;
          nextFile.lastRetryTime = new Date().getTime();
          
          log('WARN', `[${batchId}] ファイル処理失敗、再試行予定: ${nextFile.fileName} (再試行回数: ${nextFile.retryCount}/${CONFIG.BATCH_PROCESSING.MAX_RETRY_ATTEMPTS})`);
          
          // バッチ履歴を更新
          this.batchHistory.updateFileJob(jobId, {
            status: 'retrying',
            errorMessage: setupResult.error,
            retryCount: nextFile.retryCount
          });

          this.saveBatchData(batchId, batchData);

          return {
            status: 'file_retry_scheduled',
            batchId: batchId,
            fileName: nextFile.fileName,
            error: setupResult.error,
            retryCount: nextFile.retryCount,
            maxRetries: CONFIG.BATCH_PROCESSING.MAX_RETRY_ATTEMPTS,
            processedFiles: batchData.processedFiles,
            totalFiles: batchData.validFiles
          };
        } else {
          // 最終失敗
          nextFile.status = 'failed';
          nextFile.errorMessage = setupResult.error;
          nextFile.completedTime = new Date().getTime();
          batchData.failedFiles++;
          batchData.processedFiles++;

          // バッチ履歴を更新
          this.batchHistory.updateFileJob(jobId, {
            status: 'failed',
            errorMessage: setupResult.error,
            completedTime: new Date(),
            retryCount: nextFile.retryCount
          });

          this.saveBatchData(batchId, batchData);

          log('ERROR', `[${batchId}] ファイル処理最終失敗: ${nextFile.fileName}, エラー: ${setupResult.error}`);

          return {
            status: 'file_failed',
            batchId: batchId,
            fileName: nextFile.fileName,
            error: setupResult.error,
            retryCount: nextFile.retryCount,
            processedFiles: batchData.processedFiles,
            totalFiles: batchData.validFiles
          };
        }
      }

      // 単一ファイル処理の情報を保存
      nextFile.taskId = setupResult.taskId;
      nextFile.targetFileUrl = setupResult.targetFileUrl;
      nextFile.totalJobs = setupResult.totalJobs;

      this.saveBatchData(batchId, batchData);

      return {
        status: 'file_processing',
        batchId: batchId,
        fileName: nextFile.fileName,
        taskId: setupResult.taskId,
        targetFileUrl: setupResult.targetFileUrl,
        totalJobs: setupResult.totalJobs,
        processedFiles: batchData.processedFiles,
        totalFiles: batchData.validFiles
      };

    } catch (error) {
      log('ERROR', `バッチファイル処理エラー: ${batchId}`, error);
      return { status: 'error', message: error.message };
    }
  }

  /**
   * バッチ内のファイル処理が完了した際の処理
   * @param {string} batchId - バッチID
   * @param {string} taskId - 完了したタスクID
   * @param {Object} result - 処理結果
   * @return {Object} 処理結果
   */
  onFileCompleted(batchId, taskId, result) {
    try {
      const batchData = this.getBatchData(batchId);
      if (!batchData) {
        return { status: 'error', message: 'バッチが見つかりません' };
      }

      // 該当ファイルを検索
      const file = batchData.files.find(f => f.taskId === taskId);
      if (!file) {
        log('WARN', `[${batchId}] 該当ファイルが見つかりません: taskId=${taskId}`);
        return { status: 'error', message: '該当ファイルが見つかりません' };
      }

      // ファイル処理完了の処理
      file.status = result.status === 'complete' ? 'completed' : 'failed';
      file.completedTime = new Date().getTime();
      file.duration = (file.completedTime - file.startTime) / 1000;

      if (result.status === 'complete') {
        file.targetFileUrl = result.targetFileUrl || file.targetFileUrl;
        batchData.completedFiles++;
        log('INFO', `[${batchId}] ファイル処理完了: ${file.fileName}`);
      } else {
        file.errorMessage = result.message || 'ファイル処理に失敗しました';
        batchData.failedFiles++;
        log('ERROR', `[${batchId}] ファイル処理失敗: ${file.fileName}`, file.errorMessage);
      }

      batchData.processedFiles++;
      batchData.lastUpdated = new Date().getTime();

      // バッチ履歴を更新
      this.batchHistory.updateFileJob(file.jobId, {
        status: file.status,
        targetUrl: file.targetFileUrl || '',
        errorMessage: file.errorMessage || '',
        completedTime: new Date(),
        duration: file.duration
      });

      this.saveBatchData(batchId, batchData);

      // 進行状況をチェック
      const remainingFiles = batchData.files.filter(f => f.status === 'pending' || f.status === 'retrying').length;
      
      if (remainingFiles === 0) {
        // 全ファイル処理完了
        return this.completeBatch(batchId);
      }

      return {
        status: 'file_completed',
        batchId: batchId,
        fileName: file.fileName,
        fileStatus: file.status,
        processedFiles: batchData.processedFiles,
        completedFiles: batchData.completedFiles,
        failedFiles: batchData.failedFiles,
        totalFiles: batchData.validFiles,
        remainingFiles: remainingFiles
      };

    } catch (error) {
      log('ERROR', `ファイル完了処理エラー: ${batchId}`, error);
      return { status: 'error', message: error.message };
    }
  }

  /**
   * バッチ処理を一時停止する
   * @param {string} batchId - バッチID
   * @return {Object} 処理結果
   */
  pauseBatch(batchId) {
    log('INFO', `バッチ処理一時停止: ${batchId}`);

    try {
      const batchData = this.getBatchData(batchId);
      if (!batchData) {
        return { status: 'error', message: 'バッチが見つかりません' };
      }

      if (batchData.status === 'paused') {
        return { status: 'already_paused', message: 'バッチは既に一時停止中です' };
      }

      if (batchData.status === 'completed') {
        return { status: 'error', message: 'バッチは既に完了しています' };
      }

      batchData.status = 'paused';
      batchData.pausedAt = new Date().getTime();
      batchData.lastUpdated = new Date().getTime();

      // 再開用データを保存
      const resumeData = {
        pausedAt: batchData.pausedAt,
        processedFiles: batchData.processedFiles,
        completedFiles: batchData.completedFiles,
        failedFiles: batchData.failedFiles,
        resumeCount: 0
      };

      this.batchHistory.saveResumeInfo(batchId, resumeData);

      // バッチ履歴を更新
      this.batchHistory.updateBatch(batchId, {
        status: 'paused'
      });

      this.saveBatchData(batchId, batchData);

      log('INFO', `[${batchId}] バッチ処理を一時停止しました`);

      return {
        status: 'paused',
        batchId: batchId,
        processedFiles: batchData.processedFiles,
        totalFiles: batchData.validFiles,
        message: 'バッチ処理を一時停止しました'
      };

    } catch (error) {
      log('ERROR', `バッチ一時停止エラー: ${batchId}`, error);
      return { status: 'error', message: error.message };
    }
  }

  /**
   * バッチ処理を再開する
   * @param {string} batchId - バッチID
   * @return {Object} 処理結果
   */
  resumeBatch(batchId) {
    log('INFO', `バッチ処理再開: ${batchId}`);

    try {
      const batchData = this.getBatchData(batchId);
      if (!batchData) {
        return { status: 'error', message: 'バッチが見つかりません' };
      }

      if (batchData.status !== 'paused') {
        return { status: 'error', message: 'バッチは一時停止中ではありません' };
      }

      // 再開情報を取得
      const resumeInfo = this.batchHistory.getResumeInfo(batchId);
      
      batchData.status = 'processing';
      batchData.pausedAt = null;
      batchData.lastUpdated = new Date().getTime();

      if (resumeInfo) {
        resumeInfo.resumeData.resumeCount = (resumeInfo.resumeData.resumeCount || 0) + 1;
        this.batchHistory.saveResumeInfo(batchId, resumeInfo.resumeData);
      }

      // バッチ履歴を更新
      this.batchHistory.updateBatch(batchId, {
        status: 'processing'
      });

      this.saveBatchData(batchId, batchData);

      log('INFO', `[${batchId}] バッチ処理を再開しました`);

      return {
        status: 'resumed',
        batchId: batchId,
        processedFiles: batchData.processedFiles,
        totalFiles: batchData.validFiles,
        message: 'バッチ処理を再開しました'
      };

    } catch (error) {
      log('ERROR', `バッチ再開エラー: ${batchId}`, error);
      return { status: 'error', message: error.message };
    }
  }

  /**
   * バッチの進行状況を取得する
   * @param {string} batchId - バッチID
   * @return {Object} 進行状況情報
   */
  getBatchStatus(batchId) {
    try {
      const batchData = this.getBatchData(batchId);
      if (!batchData) {
        return { status: 'error', message: 'バッチが見つかりません' };
      }

      const processingFiles = batchData.files.filter(f => f.status === 'processing').length;
      const pendingFiles = batchData.files.filter(f => f.status === 'pending').length;
      const retryingFiles = batchData.files.filter(f => f.status === 'retrying').length;

      const progress = batchData.validFiles > 0 ? 
        Math.round((batchData.processedFiles / batchData.validFiles) * 100) : 0;

      const elapsedTime = batchData.startTime ? 
        Math.round((new Date().getTime() - batchData.startTime) / 1000) : 0;

      return {
        status: batchData.status,
        batchId: batchId,
        batchName: batchData.batchName,
        progress: progress,
        totalFiles: batchData.validFiles,
        processedFiles: batchData.processedFiles,
        completedFiles: batchData.completedFiles,
        failedFiles: batchData.failedFiles,
        processingFiles: processingFiles,
        pendingFiles: pendingFiles,
        retryingFiles: retryingFiles,
        invalidFiles: batchData.invalidFiles,
        elapsedTime: elapsedTime,
        targetLang: batchData.targetLang,
        dictName: batchData.dictName,
        pausedAt: batchData.pausedAt,
        lastUpdated: batchData.lastUpdated,
        files: batchData.files.map(file => ({
          index: file.index,
          fileName: file.fileName,
          fileType : file.fileType,
          status: file.status,
          errorMessage: file.errorMessage,
          targetFileUrl: file.targetFileUrl,
          duration: file.duration
        }))
      };

    } catch (error) {
      log('ERROR', `バッチステータス取得エラー: ${batchId}`, error);
      return { status: 'error', message: error.message };
    }
  }

  /**
   * バッチ処理を完了する
   * @param {string} batchId - バッチID
   * @return {Object} 完了結果
   */
  completeBatch(batchId) {
    try {
      const batchData = this.getBatchData(batchId);
      if (!batchData) {
        return { status: 'error', message: 'バッチが見つかりません' };
      }

      batchData.status = 'completed';
      batchData.completedAt = new Date().getTime();
      batchData.lastUpdated = new Date().getTime();

      const totalDuration = (batchData.completedAt - batchData.startTime) / 1000;

      // バッチ履歴を更新
      this.batchHistory.updateBatch(batchId, {
        status: 'completed',
        completedFiles: batchData.completedFiles,
        failedFiles: batchData.failedFiles,
        totalDuration: totalDuration
      });

      this.saveBatchData(batchId, batchData);

      log('INFO', `[${batchId}] バッチ処理完了: 成功${batchData.completedFiles}件, 失敗${batchData.failedFiles}件, 処理時間${Math.round(totalDuration)}秒`);

      // 成功したファイルのURLを収集
      const completedFiles = batchData.files
        .filter(file => file.status === 'completed')
        .map(file => ({
          fileName: file.fileName,
          sourceUrl: file.sourceUrl,
          targetUrl: file.targetFileUrl
        }));

      return {
        status: 'completed',
        batchId: batchId,
        totalFiles: batchData.validFiles,
        completedFiles: batchData.completedFiles,
        failedFiles: batchData.failedFiles,
        duration: Math.round(totalDuration),
        completedFileList: completedFiles,
        message: `バッチ処理が完了しました。成功: ${batchData.completedFiles}件, 失敗: ${batchData.failedFiles}件`
      };

    } catch (error) {
      log('ERROR', `バッチ完了処理エラー: ${batchId}`, error);
      return { status: 'error', message: error.message };
    }
  }

  // === プライベートメソッド ===

  /**
   * バッチデータをキャッシュから取得
   * @param {string} batchId - バッチID
   * @return {Object|null} バッチデータ
   */
  getBatchData(batchId) {
    const cached = this.cache.get(`batch_${batchId}`);
    return cached ? JSON.parse(cached) : null;
  }

  /**
   * バッチデータをキャッシュに保存
   * @param {string} batchId - バッチID
   * @param {Object} batchData - バッチデータ
   */
  saveBatchData(batchId, batchData) {
    this.cache.put(`batch_${batchId}`, JSON.stringify(batchData), 21600); // 6時間有効
  }

  /**
   * 単一ファイル翻訳をセットアップ（既存関数を利用）
   * @param {string} fileUrl - ファイルURL
   * @param {string} targetLang - 翻訳先言語
   * @param {string} dictName - 辞書名
   * @return {Object} セットアップ結果
   */
  setupSingleFileTranslation(fileUrl, targetLang, dictName) {
    return setupTranslationQueue(fileUrl, targetLang, dictName);
  }

  /**
   * Google DriveのURLかどうかを検証
   * @param {string} url - 検証するURL
   * @return {boolean} 有効なGoogle DriveのURLかどうか
   */
  isValidGoogleDriveUrl(url) {
    return isValidGoogleDriveUrl(url);
  }

  /**
   * URLからファイルIDを抽出
   * @param {string} url - Google DriveのURL
   * @return {string|null} ファイルID
   */
  extractFileId(url) {
    return extractFileId(url);
  }

  /**
   * ファイルの再試行を行うべきかどうかを判定
   * @param {Object} file - ファイル情報
   * @param {string} errorMessage - エラーメッセージ
   * @return {boolean} 再試行すべきかどうか
   */
  shouldRetryFile(file, errorMessage) {
    // 再試行回数の上限チェック
    const currentRetryCount = file.retryCount || 0;
    if (currentRetryCount >= CONFIG.BATCH_PROCESSING.MAX_RETRY_ATTEMPTS) {
      return false;
    }

    // 再试行不适用的错误类型
    const nonRetryableErrors = [
      '無効なGoogle DriveのURL',
      'ファイルが見つかりません',
      'ファイルへのアクセス権限がありません',
      'サポートされていないファイル形式',
      'ファイルIDを抽出できませんでした'
    ];

    // エラーメッセージに基づく再試行判定
    if (nonRetryableErrors.some(error => errorMessage.includes(error))) {
      log('DEBUG', `再試行不可エラー: ${errorMessage}`);
      return false;
    }

    // API関連エラーやネットワークエラーは再試行可能
    const retryableErrorPatterns = [
      'APIエラー',
      'ネットワークエラー',
      'タイムアウト',
      'サービス利用不可',
      '一時的なエラー',
      '接続エラー'
    ];

    if (retryableErrorPatterns.some(pattern => errorMessage.includes(pattern))) {
      log('DEBUG', `再試行可能エラー: ${errorMessage}`);
      return true;
    }

    // デフォルトでは一度だけ再試行を許可
    return currentRetryCount === 0;
  }

  /**
   * エラーの重要度を判定
   * @param {string} errorMessage - エラーメッセージ
   * @return {string} エラーレベル ('critical', 'warning', 'info')
   */
  getErrorSeverity(errorMessage) {
    const criticalErrors = [
      'ファイルへのアクセス権限がありません',
      'ファイルが見つかりません',
      'OpenAI APIキーが設定されていません'
    ];

    const warningErrors = [
      'APIエラー',
      'ネットワークエラー',
      'タイムアウト'
    ];

    if (criticalErrors.some(error => errorMessage.includes(error))) {
      return 'critical';
    } else if (warningErrors.some(error => errorMessage.includes(error))) {
      return 'warning';
    } else {
      return 'info';
    }
  }

  /**
   * バッチ処理のヘルスチェック
   * @param {string} batchId - バッチID
   * @return {Object} ヘルスチェック結果
   */
  performHealthCheck(batchId) {
    try {
      const batchData = this.getBatchData(batchId);
      if (!batchData) {
        return {
          status: 'unhealthy',
          issues: ['バッチデータが見つかりません'],
          recommendations: ['バッチを再作成してください']
        };
      }

      const issues = [];
      const recommendations = [];

      // 停滞チェック
      const lastUpdated = batchData.lastUpdated || batchData.startTime;
      const staleThreshold = 30 * 60 * 1000; // 30分
      const timeSinceUpdate = new Date().getTime() - lastUpdated;

      if (timeSinceUpdate > staleThreshold && batchData.status === 'processing') {
        issues.push('処理が30分以上停滞しています');
        recommendations.push('バッチを一時停止して再開を試してください');
      }

      // エラー率チェック
      const errorRate = batchData.failedFiles / (batchData.processedFiles || 1);
      if (errorRate > 0.5 && batchData.processedFiles > 3) {
        issues.push(`エラー率が高すぎます: ${Math.round(errorRate * 100)}%`);
        recommendations.push('ファイルの権限設定や辞書設定を確認してください');
      }

      // 再試行の頻度チェック
      const retryingFiles = batchData.files.filter(f => f.status === 'retrying').length;
      if (retryingFiles > 3) {
        issues.push(`多数のファイルが再試行待ちです: ${retryingFiles}件`);
        recommendations.push('システムリソースやAPI制限を確認してください');
      }

      const healthStatus = issues.length === 0 ? 'healthy' : 
                          issues.length <= 2 ? 'warning' : 'unhealthy';

      return {
        status: healthStatus,
        batchId: batchId,
        issues: issues,
        recommendations: recommendations,
        metrics: {
          errorRate: Math.round(errorRate * 100),
          timeSinceUpdate: Math.round(timeSinceUpdate / 1000 / 60), // 分単位
          retryingFiles: retryingFiles,
          processingTime: Math.round((new Date().getTime() - batchData.startTime) / 1000 / 60) // 分単位
        }
      };

    } catch (error) {
      log('ERROR', `ヘルスチェックエラー: ${batchId}`, error);
      return {
        status: 'error',
        issues: ['ヘルスチェックの実行に失敗しました'],
        error: error.message
      };
    }
  }

  /**
   * バッチ処理の自動回復を試行
   * @param {string} batchId - バッチID
   * @return {Object} 回復試行結果
   */
  attemptAutoRecovery(batchId) {
    try {
      const healthCheck = this.performHealthCheck(batchId);
      
      if (healthCheck.status === 'healthy') {
        return { status: 'no_action_needed', message: 'バッチは正常に動作しています' };
      }

      const batchData = this.getBatchData(batchId);
      if (!batchData) {
        return { status: 'failed', message: 'バッチデータが見つかりません' };
      }

      let recoveryActions = [];

      // 停滞している場合の回復アクション
      const lastUpdated = batchData.lastUpdated || batchData.startTime;
      const timeSinceUpdate = new Date().getTime() - lastUpdated;
      const staleThreshold = 30 * 60 * 1000; // 30分

      if (timeSinceUpdate > staleThreshold && batchData.status === 'processing') {
        // 処理中だが更新されていないファイルを再試行対象にする
        const stuckFiles = batchData.files.filter(f => 
          f.status === 'processing' && 
          (new Date().getTime() - (f.startTime || 0)) > staleThreshold
        );

        stuckFiles.forEach(file => {
          file.status = 'retrying';
          file.retryCount = (file.retryCount || 0) + 1;
          file.errorMessage = '処理タイムアウトのため再試行';
          recoveryActions.push(`ファイル再試行: ${file.fileName}`);
        });

        batchData.lastUpdated = new Date().getTime();
        this.saveBatchData(batchId, batchData);
      }

      // 多数のファイルが失敗している場合の回復アクション
      const failedFiles = batchData.files.filter(f => f.status === 'failed');
      if (failedFiles.length > 0) {
        // 一部の失敗ファイルを再試行対象にする（権限関連以外）
        const recoverableFiles = failedFiles.filter(f => 
          f.retryCount < CONFIG.BATCH_PROCESSING.MAX_RETRY_ATTEMPTS &&
          !f.errorMessage.includes('権限') &&
          !f.errorMessage.includes('見つかりません')
        ).slice(0, 3); // 最大3ファイルまで

        recoverableFiles.forEach(file => {
          file.status = 'retrying';
          file.retryCount = (file.retryCount || 0) + 1;
          batchData.failedFiles--;
          recoveryActions.push(`失敗ファイル再試行: ${file.fileName}`);
        });

        if (recoverableFiles.length > 0) {
          batchData.lastUpdated = new Date().getTime();
          this.saveBatchData(batchId, batchData);
        }
      }

      log('INFO', `[${batchId}] 自動回復実行: ${recoveryActions.length}件のアクション`);

      return {
        status: recoveryActions.length > 0 ? 'recovery_attempted' : 'no_recovery_possible',
        batchId: batchId,
        actions: recoveryActions,
        healthStatus: healthCheck.status,
        message: recoveryActions.length > 0 ? 
          `${recoveryActions.length}件の回復アクションを実行しました` :
          '自動回復可能な問題は見つかりませんでした'
      };

    } catch (error) {
      log('ERROR', `自動回復エラー: ${batchId}`, error);
      return {
        status: 'error',
        message: error.message
      };
    }
  }
}
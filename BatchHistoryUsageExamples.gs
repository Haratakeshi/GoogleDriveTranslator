// BatchHistoryUsageExamples.gs - バッチ履歴管理システムの使用例

/**
 * バッチ履歴管理システムの基本的な使用例
 * 
 * このファイルは実装例を示すためのもので、実際のバッチ処理実装時の参考として使用してください。
 */

/**
 * 使用例1: 基本的なバッチ処理の履歴管理
 */
function exampleBasicBatchProcessing() {
  // 履歴管理インスタンスを作成
  const history = new History();
  
  // 1. バッチを作成
  const batchData = {
    batchName: '月次レポート翻訳_2024年1月',
    totalFiles: 15,
    targetLang: 'en',
    dictName: 'ビジネス辞書',
    settings: {
      domain: 'ビジネス',
      priority: 'high',
      notifyOnComplete: true
    }
  };
  
  const batchId = history.createBatch(batchData);
  log('INFO', `バッチを作成しました: ${batchId}`);
  
  // 2. ファイル処理をシミュレート
  const filesToProcess = [
    { url: 'https://docs.google.com/document/d/file1', name: 'report1.docx', type: 'document' },
    { url: 'https://docs.google.com/spreadsheets/d/file2', name: 'data1.xlsx', type: 'spreadsheet' },
    { url: 'https://docs.google.com/presentation/d/file3', name: 'slides1.pptx', type: 'presentation' }
  ];
  
  let completedFiles = 0;
  let failedFiles = 0;
  let totalChars = 0;
  let totalCost = 0;
  
  for (let i = 0; i < filesToProcess.length; i++) {
    const file = filesToProcess[i];
    
    // ファイル処理開始を記録
    const jobId = history.recordBatchFileJob({
      batchId: batchId,
      sourceUrl: file.url,
      fileName: file.name,
      fileType: file.type,
      status: 'processing',
      charCount: 0,
      startTime: new Date()
    });
    
    try {
      // ここで実際の翻訳処理を行う（シミュレート）
      const translationResult = simulateTranslation(file);
      
      // 成功した場合の記録更新
      history.updateBatchFileJob(jobId, {
        status: 'completed',
        targetUrl: translationResult.targetUrl,
        charCount: translationResult.charCount,
        completedTime: new Date(),
        duration: translationResult.duration,
        apiCost: translationResult.cost
      });
      
      completedFiles++;
      totalChars += translationResult.charCount;
      totalCost += translationResult.cost;
      
      // 従来の履歴システムにも記録（互換性維持）
      history.record({
        batchId: batchId,
        sourceUrl: file.url,
        targetUrl: translationResult.targetUrl,
        sourceLang: 'ja',
        targetLang: 'en',
        dictName: 'ビジネス辞書',
        charCountSource: translationResult.charCount,
        charCountTarget: translationResult.translatedCharCount,
        duration: translationResult.duration,
        status: 'success'
      });
      
    } catch (error) {
      // 失敗した場合の記録更新
      history.updateBatchFileJob(jobId, {
        status: 'failed',
        completedTime: new Date(),
        errorMessage: error.message
      });
      
      failedFiles++;
      
      // 従来の履歴システムにも記録
      history.record({
        batchId: batchId,
        sourceUrl: file.url,
        sourceLang: 'ja',
        targetLang: 'en',
        dictName: 'ビジネス辞書',
        status: 'error',
        errorMessage: error.message
      });
    }
    
    // 進行状況を更新
    history.updateBatchProgress(batchId, completedFiles, failedFiles, {
      totalSourceChars: totalChars,
      totalApiCost: totalCost
    });
    
    log('INFO', `ファイル ${i + 1}/${filesToProcess.length} 処理完了`);
  }
  
  log('INFO', `バッチ処理完了: 成功=${completedFiles}, 失敗=${failedFiles}`);
  
  return {
    batchId: batchId,
    completedFiles: completedFiles,
    failedFiles: failedFiles,
    totalCost: totalCost
  };
}

/**
 * 使用例2: バッチ処理の中断と再開
 */
function exampleBatchPauseResume() {
  const history = new History();
  
  // バッチを作成
  const batchId = history.createBatch({
    batchName: '大量文書翻訳バッチ',
    totalFiles: 100,
    targetLang: 'en'
  });
  
  // 処理中にエラーが発生したとシミュレート
  const processedFiles = 25;
  const remainingFiles = [/* 残りのファイルリスト */];
  
  // 再開情報を保存してバッチを一時停止
  const resumeData = {
    processedFiles: processedFiles,
    remainingFileUrls: remainingFiles,
    lastProcessedIndex: processedFiles - 1,
    accumulatedStats: {
      totalChars: 50000,
      totalCost: 15.50,
      totalDuration: 1800
    },
    processingSettings: {
      batchSize: 5,
      retryCount: 3
    },
    resumeCount: 0
  };
  
  const pauseSuccess = history.pauseBatch(batchId, resumeData);
  log('INFO', `バッチ一時停止: ${pauseSuccess ? '成功' : '失敗'}`);
  
  // 後で再開
  const retrievedResumeData = history.resumeBatch(batchId);
  if (retrievedResumeData) {
    log('INFO', `バッチ再開: ${retrievedResumeData.processedFiles}ファイルから継続`);
    
    // 再開データを使用して処理を継続
    continueProcessingFromResume(batchId, retrievedResumeData);
  }
  
  return { batchId: batchId, resumeData: retrievedResumeData };
}

/**
 * 使用例3: バッチ統計とレポート生成
 */
function exampleBatchReporting() {
  const history = new History();
  
  // バッチ統計を取得
  const batchStats = history.getBatchStatistics();
  
  // 複合統計を取得（個別 + バッチ）
  const combinedStats = history.getCombinedStatistics('month');
  
  // レポートを生成
  const report = {
    reportDate: new Date(),
    batchSummary: {
      totalBatches: batchStats.totalBatches,
      completedBatches: batchStats.completedBatches,
      failedBatches: batchStats.failedBatches,
      successRate: batchStats.successRate + '%'
    },
    fileSummary: {
      totalFiles: batchStats.totalFiles,
      completedFiles: batchStats.completedFiles,
      failedFiles: batchStats.failedFiles,
      averageFilesPerBatch: batchStats.averageFilesPerBatch
    },
    costAnalysis: {
      totalApiCost: '$' + batchStats.totalApiCost.toFixed(4),
      averageCostPerFile: batchStats.totalFiles > 0 ? 
        '$' + (batchStats.totalApiCost / batchStats.totalFiles).toFixed(4) : '$0',
      totalProcessingTime: Math.round(batchStats.totalDuration / 60) + ' minutes'
    },
    languageDistribution: batchStats.languageDistribution,
    combinedPerformance: {
      totalTranslations: combinedStats.combined.totalTranslations,
      overallSuccessRate: combinedStats.combined.successRate + '%',
      totalCharactersProcessed: combinedStats.combined.totalSourceChars
    }
  };
  
  log('INFO', 'バッチレポートを生成しました', report);
  
  return report;
}

/**
 * 使用例4: 履歴のクリーンアップとメンテナンス
 */
function exampleHistoryMaintenance() {
  const history = new History();
  
  // 期限切れの再開情報をクリーンアップ
  const cleanupCount = history.cleanupExpiredResumeInfo();
  log('INFO', `期限切れ再開情報をクリーンアップ: ${cleanupCount}件`);
  
  // 古いバッチ履歴のアーカイブ（90日以上前）
  const oldBatches = history.getBatchHistoryList({ limit: 1000 })
    .filter(batch => {
      const batchDate = new Date(batch.createdAt);
      const daysAgo = (new Date() - batchDate) / (1000 * 60 * 60 * 24);
      return daysAgo > CONFIG.BATCH_HISTORY.CLEANUP_AFTER_DAYS;
    });
  
  log('INFO', `アーカイブ対象バッチ: ${oldBatches.length}件`);
  
  // メンテナンス統計
  const maintenanceStats = {
    cleanupCount: cleanupCount,
    archiveCandidates: oldBatches.length,
    totalBatchHistory: history.getBatchHistoryList({ limit: 10000 }).length,
    lastCleanupDate: new Date()
  };
  
  return maintenanceStats;
}

// ==============================================
// ヘルパー関数（シミュレーション用）
// ==============================================

/**
 * 翻訳処理のシミュレーション（実際の実装では置き換える）
 */
function simulateTranslation(file) {
  // ランダムな処理時間をシミュレート
  const processingTime = Math.random() * 30 + 10; // 10-40秒
  const charCount = Math.floor(Math.random() * 5000) + 500; // 500-5500文字
  const translatedCharCount = Math.floor(charCount * (0.8 + Math.random() * 0.4)); // 80-120%
  
  // 5%の確率で失敗をシミュレート
  if (Math.random() < 0.05) {
    throw new Error('翻訳APIエラー: レート制限に達しました');
  }
  
  return {
    targetUrl: file.url.replace('/edit', '_translated/edit'),
    charCount: charCount,
    translatedCharCount: translatedCharCount,
    duration: processingTime,
    cost: charCount * 0.000025 // 仮の料金計算
  };
}

/**
 * 再開データからの処理継続（実際の実装では置き換える）
 */
function continueProcessingFromResume(batchId, resumeData) {
  log('INFO', `バッチ ${batchId} の処理を再開します`);
  log('INFO', `処理済み: ${resumeData.processedFiles}ファイル`);
  log('INFO', `残り: ${resumeData.remainingFileUrls?.length || 0}ファイル`);
  
  // 実際の再開処理はここに実装
  // resumeData.remainingFileUrls を使用して残りのファイルを処理
  
  return true;
}
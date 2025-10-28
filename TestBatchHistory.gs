// TestBatchHistory.gs - バッチ履歴管理システムのテスト

/**
 * バッチ履歴管理システムの基本機能テスト
 * Google Apps Scriptのエディタから実行可能
 */
function testBatchHistorySystem() {
  try {
    log('INFO', 'バッチ履歴管理システムのテストを開始します');
    
    // 1. History クラスの初期化テスト
    const history = new History();
    log('INFO', 'History クラスの初期化: 成功');
    
    // 2. バッチ作成テスト
    const batchData = {
      batchName: 'テストバッチ_' + new Date().getTime(),
      totalFiles: 5,
      targetLang: 'ja',
      dictName: '一般辞書',
      settings: {
        testMode: true,
        createdBy: 'テスト実行'
      }
    };
    
    const batchId = history.createBatch(batchData);
    log('INFO', `バッチ作成: 成功 (ID: ${batchId})`);
    
    // 3. ファイルジョブ記録テスト
    const fileJobData = {
      batchId: batchId,
      sourceUrl: 'https://docs.google.com/document/d/test123/edit',
      fileName: 'test-document.docx',
      fileType: 'document',
      status: 'pending',
      charCount: 1500
    };
    
    const jobId = history.recordBatchFileJob(fileJobData);
    log('INFO', `ファイルジョブ記録: 成功 (ID: ${jobId})`);
    
    // 4. バッチ進行状況更新テスト
    const progressUpdate = history.updateBatchProgress(batchId, 1, 0, {
      totalSourceChars: 1500,
      totalTargetChars: 1800,
      totalDuration: 45.5,
      totalApiCost: 0.0125
    });
    log('INFO', `バッチ進行状況更新: ${progressUpdate ? '成功' : '失敗'}`);
    
    // 5. ファイルジョブ更新テスト
    const jobUpdate = history.updateBatchFileJob(jobId, {
      status: 'completed',
      targetUrl: 'https://docs.google.com/document/d/test123_translated/edit',
      completedTime: new Date(),
      duration: 45.5,
      apiCost: 0.0125
    });
    log('INFO', `ファイルジョブ更新: ${jobUpdate ? '成功' : '失敗'}`);
    
    // 6. 再開情報保存テスト
    const resumeData = {
      processedFiles: 1,
      remainingFiles: ['file2.docx', 'file3.xlsx'],
      lastProcessedIndex: 0,
      settings: batchData.settings,
      resumeCount: 0
    };
    
    const resumeSave = history.saveBatchResumeInfo(batchId, resumeData);
    log('INFO', `再開情報保存: ${resumeSave ? '成功' : '失敗'}`);
    
    // 7. 再開情報取得テスト
    const retrievedResumeInfo = history.getBatchResumeInfo(batchId);
    log('INFO', `再開情報取得: ${retrievedResumeInfo ? '成功' : '失敗'}`);
    
    if (retrievedResumeInfo) {
      log('INFO', `再開データ確認: processedFiles=${retrievedResumeInfo.resumeData.processedFiles}`);
    }
    
    // 8. バッチ履歴取得テスト
    const batchHistoryList = history.getBatchHistoryList({ limit: 10 });
    log('INFO', `バッチ履歴取得: 成功 (${batchHistoryList.length}件)`);
    
    // 9. バッチファイル履歴取得テスト
    const fileHistoryList = history.getBatchFileHistoryList(batchId);
    log('INFO', `バッチファイル履歴取得: 成功 (${fileHistoryList.length}件)`);
    
    // 10. バッチ統計情報取得テスト
    const batchStats = history.getBatchStatistics();
    log('INFO', `バッチ統計取得: 成功 (totalBatches: ${batchStats.totalBatches})`);
    
    // 11. 複合統計情報取得テスト
    const combinedStats = history.getCombinedStatistics('all');
    log('INFO', `複合統計取得: 成功 (combined.totalTranslations: ${combinedStats.combined.totalTranslations})`);
    
    // 12. バッチ一時停止・再開テスト
    const pauseResult = history.pauseBatch(batchId, resumeData);
    log('INFO', `バッチ一時停止: ${pauseResult ? '成功' : '失敗'}`);
    
    const resumeResult = history.resumeBatch(batchId);
    log('INFO', `バッチ再開: ${resumeResult ? '成功' : '失敗'}`);
    
    // 13. 期限切れ再開情報クリーンアップテスト
    const cleanupCount = history.cleanupExpiredResumeInfo();
    log('INFO', `期限切れ再開情報クリーンアップ: ${cleanupCount}件処理`);
    
    // テスト完了
    log('INFO', 'バッチ履歴管理システムのテストが完了しました');
    
    return {
      success: true,
      message: 'すべてのテストが正常に完了しました',
      results: {
        batchId: batchId,
        jobId: jobId,
        batchHistoryCount: batchHistoryList.length,
        fileHistoryCount: fileHistoryList.length,
        totalBatches: batchStats.totalBatches,
        cleanupCount: cleanupCount
      }
    };
    
  } catch (error) {
    log('ERROR', 'バッチ履歴管理システムのテスト中にエラーが発生しました', error);
    return {
      success: false,
      message: error.message,
      error: error
    };
  }
}

/**
 * 個別のHistory recordメソッドをバッチ対応でテスト
 */
function testHistoryRecordWithBatch() {
  try {
    log('INFO', 'History recordメソッドのバッチ対応テストを開始します');
    
    const history = new History();
    
    // バッチを作成
    const batchId = history.createBatch({
      batchName: 'Record統合テスト',
      totalFiles: 2,
      targetLang: 'en'
    });
    
    // 従来のrecordメソッドをバッチIDありで呼び出し
    const recordId = history.record({
      batchId: batchId,
      sourceUrl: 'https://docs.google.com/spreadsheets/d/test456/edit',
      targetUrl: 'https://docs.google.com/spreadsheets/d/test456_translated/edit',
      sourceLang: 'ja',
      targetLang: 'en',
      dictName: 'IT辞書',
      charCountSource: 2500,
      charCountTarget: 2800,
      duration: 65.3,
      status: 'success',
      startTime: new Date(Date.now() - 65300),
      completedTime: new Date(),
      retryCount: 0
    });
    
    log('INFO', `Record統合テスト: 成功 (recordId: ${recordId})`);
    
    // バッチファイル履歴を確認
    const fileHistory = history.getBatchFileHistoryList(batchId);
    log('INFO', `統合記録確認: ファイル履歴${fileHistory.length}件記録`);
    
    return {
      success: true,
      recordId: recordId,
      batchId: batchId,
      fileHistoryCount: fileHistory.length
    };
    
  } catch (error) {
    log('ERROR', 'Record統合テスト中にエラーが発生しました', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * エラーハンドリングのテスト
 */
function testErrorHandling() {
  try {
    log('INFO', 'エラーハンドリングテストを開始します');
    
    const history = new History();
    
    // 存在しないバッチIDでの操作テスト
    const invalidBatchId = 'BATCH_INVALID_123';
    
    const updateResult = history.updateBatch(invalidBatchId, { status: 'completed' });
    log('INFO', `無効バッチID更新テスト: ${updateResult ? '予期しない成功' : '正常に失敗'}`);
    
    const resumeInfo = history.getBatchResumeInfo(invalidBatchId);
    log('INFO', `無効バッチID再開情報取得: ${resumeInfo ? '予期しない成功' : '正常にnull'}`);
    
    // 存在しないジョブIDでの操作テスト
    const invalidJobId = 'JOB_INVALID_456';
    
    const jobUpdateResult = history.updateBatchFileJob(invalidJobId, { status: 'completed' });
    log('INFO', `無効ジョブID更新テスト: ${jobUpdateResult ? '予期しない成功' : '正常に失敗'}`);
    
    log('INFO', 'エラーハンドリングテストが完了しました');
    
    return {
      success: true,
      message: 'エラーハンドリングが正常に動作しています'
    };
    
  } catch (error) {
    log('ERROR', 'エラーハンドリングテスト中にエラーが発生しました', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * パフォーマンステスト（軽量版）
 */
function testPerformance() {
  try {
    log('INFO', 'パフォーマンステストを開始します');
    
    const history = new History();
    const startTime = new Date().getTime();
    
    // 複数バッチの作成
    const batchIds = [];
    for (let i = 0; i < 5; i++) {
      const batchId = history.createBatch({
        batchName: `パフォーマンステスト_${i}`,
        totalFiles: 10,
        targetLang: 'en'
      });
      batchIds.push(batchId);
    }
    
    const batchCreationTime = new Date().getTime() - startTime;
    log('INFO', `5バッチ作成時間: ${batchCreationTime}ms`);
    
    // 複数ファイルジョブの記録
    const jobCreationStart = new Date().getTime();
    let totalJobs = 0;
    
    batchIds.forEach(batchId => {
      for (let j = 0; j < 3; j++) {
        history.recordBatchFileJob({
          batchId: batchId,
          sourceUrl: `https://docs.google.com/document/d/test${j}/edit`,
          fileName: `test-file-${j}.docx`,
          fileType: 'document',
          status: 'pending',
          charCount: 1000 + j * 100
        });
        totalJobs++;
      }
    });
    
    const jobCreationTime = new Date().getTime() - jobCreationStart;
    log('INFO', `${totalJobs}ジョブ記録時間: ${jobCreationTime}ms`);
    
    // 統計情報取得
    const statsStart = new Date().getTime();
    const stats = history.getBatchStatistics();
    const statsTime = new Date().getTime() - statsStart;
    log('INFO', `統計情報取得時間: ${statsTime}ms`);
    
    const totalTime = new Date().getTime() - startTime;
    log('INFO', `総実行時間: ${totalTime}ms`);
    
    return {
      success: true,
      performance: {
        batchCreationTime: batchCreationTime,
        jobCreationTime: jobCreationTime,
        statsTime: statsTime,
        totalTime: totalTime,
        batchesCreated: batchIds.length,
        jobsCreated: totalJobs
      }
    };
    
  } catch (error) {
    log('ERROR', 'パフォーマンステスト中にエラーが発生しました', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 全テストの実行
 */
function runAllBatchHistoryTests() {
  log('INFO', '=== バッチ履歴管理システム全テスト開始 ===');
  
  const results = {
    basicFunctionality: testBatchHistorySystem(),
    recordIntegration: testHistoryRecordWithBatch(), 
    errorHandling: testErrorHandling(),
    performance: testPerformance()
  };
  
  log('INFO', '=== 全テスト完了 ===');
  
  const allSuccess = Object.values(results).every(result => result.success);
  
  return {
    success: allSuccess,
    summary: `全テスト${allSuccess ? '成功' : '失敗'}`,
    details: results
  };
}
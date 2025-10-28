// BatchHistory.gs - バッチ履歴管理クラス

/**
 * バッチ処理の履歴記録と管理を行うクラス
 * 複数ファイルの一括翻訳における進行状況、結果、再開情報を管理
 */
class BatchHistory {
  constructor() {
    this.spreadsheetId = CONFIG.HISTORY_SHEET_ID;
    this.spreadsheet = null;
    this.batchHistorySheet = null;
    this.fileHistorySheet = null;
    this.resumeInfoSheet = null;
    this.initSpreadsheet();
  }
  
  /**
   * スプレッドシートを初期化
   */
  initSpreadsheet() {
    try {
      this.spreadsheet = SpreadsheetApp.openById(this.spreadsheetId);
      this.batchHistorySheet = this.getOrCreateBatchHistorySheet();
      this.fileHistorySheet = this.getOrCreateFileHistorySheet();
      this.resumeInfoSheet = this.getOrCreateResumeInfoSheet();
    } catch (error) {
      log('ERROR', 'バッチ履歴スプレッドシートの初期化エラー', error);
      throw new Error('バッチ履歴スプレッドシートにアクセスできません');
    }
  }
  
  /**
   * バッチ履歴シートを取得または作成
   * @return {Sheet} バッチ履歴シート
   */
  getOrCreateBatchHistorySheet() {
    let sheet = this.spreadsheet.getSheetByName('バッチ履歴');
    
    if (!sheet) {
      sheet = this.spreadsheet.insertSheet('バッチ履歴');
      
      // ヘッダー設定
      const headers = [
        'バッチID',
        'バッチ名',
        'ステータス',
        '作成日時',
        '更新日時',
        '総ファイル数',
        '完了ファイル数',
        '失敗ファイル数',
        '対象言語',
        '使用辞書',
        '総原文文字数',
        '総翻訳文字数',
        '総処理時間（秒）',
        '総APIコスト（USD）',
        '作成者',
        'エラーメッセージ',
        '設定情報'
      ];
      
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);
      
      // スタイル設定
      headerRange.setBackground('#1f4e79');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
      
      // 列幅の調整
      sheet.setColumnWidth(1, 150);  // バッチID
      sheet.setColumnWidth(2, 200);  // バッチ名
      sheet.setColumnWidth(3, 100);  // ステータス
      sheet.setColumnWidth(4, 150);  // 作成日時
      sheet.setColumnWidth(5, 150);  // 更新日時
      sheet.setColumnWidth(6, 100);  // 総ファイル数
      sheet.setColumnWidth(7, 100);  // 完了ファイル数
      sheet.setColumnWidth(8, 100);  // 失敗ファイル数
      sheet.setColumnWidth(9, 120);  // 対象言語
      sheet.setColumnWidth(10, 150); // 使用辞書
      sheet.setColumnWidth(11, 120); // 総原文文字数
      sheet.setColumnWidth(12, 120); // 総翻訳文字数
      sheet.setColumnWidth(13, 120); // 総処理時間
      sheet.setColumnWidth(14, 120); // 総APIコスト
      sheet.setColumnWidth(15, 150); // 作成者
      sheet.setColumnWidth(16, 200); // エラーメッセージ
      sheet.setColumnWidth(17, 300); // 設定情報
      
      // フリーズ設定
      sheet.setFrozenRows(1);
      
      // データ検証を設定
      this.setupBatchDataValidation(sheet);
    }
    
    return sheet;
  }
  
  /**
   * ファイル履歴シートを取得または作成
   * @return {Sheet} ファイル履歴シート
   */
  getOrCreateFileHistorySheet() {
    let sheet = this.spreadsheet.getSheetByName('ファイル履歴');
    
    if (!sheet) {
      sheet = this.spreadsheet.insertSheet('ファイル履歴');
      
      // ヘッダー設定
      const headers = [
        'ジョブID',
        'バッチID',
        '翻訳元URL',
        '翻訳先URL',
        'ファイル名',
        'ファイルタイプ',
        'ステータス',
        '文字数',
        '開始時刻',
        '完了時刻',
        'エラーメッセージ',
        '再開トークン',
        '処理時間（秒）',
        'APIコスト（USD）',
        '再試行回数'
      ];
      
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);
      
      // スタイル設定
      headerRange.setBackground('#0f5132');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
      
      // 列幅の調整
      sheet.setColumnWidth(1, 150);  // ジョブID
      sheet.setColumnWidth(2, 150);  // バッチID
      sheet.setColumnWidth(3, 300);  // 翻訳元URL
      sheet.setColumnWidth(4, 300);  // 翻訳先URL
      sheet.setColumnWidth(5, 200);  // ファイル名
      sheet.setColumnWidth(6, 120);  // ファイルタイプ
      sheet.setColumnWidth(7, 100);  // ステータス
      sheet.setColumnWidth(8, 100);  // 文字数
      sheet.setColumnWidth(9, 150);  // 開始時刻
      sheet.setColumnWidth(10, 150); // 完了時刻
      sheet.setColumnWidth(11, 200); // エラーメッセージ
      sheet.setColumnWidth(12, 300); // 再開トークン
      sheet.setColumnWidth(13, 120); // 処理時間
      sheet.setColumnWidth(14, 120); // APIコスト
      sheet.setColumnWidth(15, 100); // 再試行回数
      
      // フリーズ設定
      sheet.setFrozenRows(1);
      
      // データ検証を設定
      this.setupFileDataValidation(sheet);
    }
    
    return sheet;
  }
  
  /**
   * 再開情報シートを取得または作成
   * @return {Sheet} 再開情報シート
   */
  getOrCreateResumeInfoSheet() {
    let sheet = this.spreadsheet.getSheetByName('再開情報');
    
    if (!sheet) {
      sheet = this.spreadsheet.insertSheet('再開情報');
      
      // ヘッダー設定
      const headers = [
        'バッチID',
        '再開データ（JSON）',
        '作成日時',
        '最終更新日時',
        'アクティブ',
        '有効期限',
        'メタデータ'
      ];
      
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);
      
      // スタイル設定
      headerRange.setBackground('#7b2d26');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
      
      // 列幅の調整
      sheet.setColumnWidth(1, 150);  // バッチID
      sheet.setColumnWidth(2, 400);  // 再開データ
      sheet.setColumnWidth(3, 150);  // 作成日時
      sheet.setColumnWidth(4, 150);  // 最終更新日時
      sheet.setColumnWidth(5, 100);  // アクティブ
      sheet.setColumnWidth(6, 150);  // 有効期限
      sheet.setColumnWidth(7, 300);  // メタデータ
      
      // フリーズ設定
      sheet.setFrozenRows(1);
    }
    
    return sheet;
  }
  
  /**
   * バッチデータ検証を設定
   * @param {Sheet} sheet - 対象シート
   */
  setupBatchDataValidation(sheet) {
    // ステータスの検証
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['pending', 'processing', 'completed', 'failed', 'cancelled', 'paused'])
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 3, 1000, 1).setDataValidation(statusRule);
  }
  
  /**
   * ファイルデータ検証を設定
   * @param {Sheet} sheet - 対象シート
   */
  setupFileDataValidation(sheet) {
    // ステータスの検証
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['pending', 'processing', 'completed', 'failed', 'skipped', 'retrying'])
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 7, 1000, 1).setDataValidation(statusRule);
    
    // ファイルタイプの検証
    const fileTypeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['spreadsheet', 'document', 'presentation', 'pdf'])
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 6, 1000, 1).setDataValidation(fileTypeRule);
  }
  
  /**
   * 新しいバッチを作成
   * @param {Object} batchData - バッチデータ
   * @return {string} 作成したバッチID
   */
  createBatch(batchData) {
    try {
      const batchId = this.generateBatchId();
      const currentUser = Session.getActiveUser().getEmail();
      const currentTime = new Date();
      
      const row = [
        batchId,
        batchData.batchName || `Batch_${Utilities.formatDate(currentTime, 'Asia/Tokyo', 'yyyy-MM-dd_HH-mm')}`,
        'pending',
        currentTime,
        currentTime,
        batchData.totalFiles || 0,
        0, // completedFiles
        0, // failedFiles
        batchData.targetLang || '',
        batchData.dictName || 'なし',
        0, // totalSourceChars
        0, // totalTargetChars
        0, // totalDuration
        0, // totalApiCost
        currentUser,
        '', // errorMessage
        JSON.stringify(batchData.settings || {})
      ];
      
      this.batchHistorySheet.appendRow(row);
      
      // 条件付き書式を適用
      this.applyBatchConditionalFormatting();
      
      log('INFO', `バッチを作成しました: ID=${batchId}`);
      
      return batchId;
      
    } catch (error) {
      log('ERROR', 'バッチ作成エラー', error);
      throw error;
    }
  }
  
  /**
   * バッチ情報を更新
   * @param {string} batchId - バッチID
   * @param {Object} updateData - 更新データ
   * @return {boolean} 更新成功かどうか
   */
  updateBatch(batchId, updateData) {
    try {
      const rowIndex = this.findBatchRowIndex(batchId);
      if (rowIndex === -1) {
        throw new Error(`バッチが見つかりません: ${batchId}`);
      }
      
      const currentValues = this.batchHistorySheet.getRange(rowIndex, 1, 1, 17).getValues()[0];
      
      // 更新可能なフィールドのマッピング
      const fieldMap = {
        batchName: 1,
        status: 2,
        totalFiles: 5,
        completedFiles: 6,
        failedFiles: 7,
        totalSourceChars: 10,
        totalTargetChars: 11,
        totalDuration: 12,
        totalApiCost: 13,
        errorMessage: 15
      };
      
      // 更新日時は常に更新
      currentValues[4] = new Date();
      
      // 指定されたフィールドを更新
      for (const [field, columnIndex] of Object.entries(fieldMap)) {
        if (updateData.hasOwnProperty(field)) {
          currentValues[columnIndex] = updateData[field];
        }
      }
      
      this.batchHistorySheet.getRange(rowIndex, 1, 1, 17).setValues([currentValues]);
      
      // 条件付き書式を適用
      this.applyBatchConditionalFormatting();
      
      log('INFO', `バッチを更新しました: ID=${batchId}`);
      
      return true;
      
    } catch (error) {
      log('ERROR', 'バッチ更新エラー', error);
      return false;
    }
  }
  
  /**
   * ファイル処理記録を追加
   * @param {Object} fileData - ファイルデータ
   * @return {string} 作成したジョブID
   */
  recordFileJob(fileData) {
    try {
      const jobId = this.generateJobId();
      const currentTime = new Date();
      
      const row = [
        jobId,
        fileData.batchId,
        fileData.sourceUrl || '',
        fileData.targetUrl || '',
        fileData.fileName || '',
        fileData.fileType || '',
        fileData.status || 'pending',
        fileData.charCount || 0,
        fileData.startTime || currentTime,
        fileData.completedTime || '',
        fileData.errorMessage || '',
        fileData.resumeToken || '',
        fileData.duration || 0,
        fileData.apiCost || 0,
        fileData.retryCount || 0
      ];
      
      this.fileHistorySheet.appendRow(row);
      
      // 条件付き書式を適用
      this.applyFileConditionalFormatting();
      
      log('INFO', `ファイルジョブを記録しました: ID=${jobId}`);
      
      return jobId;
      
    } catch (error) {
      log('ERROR', 'ファイルジョブ記録エラー', error);
      throw error;
    }
  }
  
  /**
   * ファイル処理状況を更新
   * @param {string} jobId - ジョブID
   * @param {Object} updateData - 更新データ
   * @return {boolean} 更新成功かどうか
   */
  updateFileJob(jobId, updateData) {
    try {
      const rowIndex = this.findFileJobRowIndex(jobId);
      if (rowIndex === -1) {
        throw new Error(`ファイルジョブが見つかりません: ${jobId}`);
      }
      
      const currentValues = this.fileHistorySheet.getRange(rowIndex, 1, 1, 15).getValues()[0];
      
      // 更新可能なフィールドのマッピング
      const fieldMap = {
        targetUrl: 3,
        status: 6,
        charCount: 7,
        completedTime: 9,
        errorMessage: 10,
        resumeToken: 11,
        duration: 12,
        apiCost: 13,
        retryCount: 14
      };
      
      // 指定されたフィールドを更新
      for (const [field, columnIndex] of Object.entries(fieldMap)) {
        if (updateData.hasOwnProperty(field)) {
          currentValues[columnIndex] = updateData[field];
        }
      }
      
      this.fileHistorySheet.getRange(rowIndex, 1, 1, 15).setValues([currentValues]);
      
      // 条件付き書式を適用
      this.applyFileConditionalFormatting();
      
      log('INFO', `ファイルジョブを更新しました: ID=${jobId}`);
      
      return true;
      
    } catch (error) {
      log('ERROR', 'ファイルジョブ更新エラー', error);
      return false;
    }
  }
  
  /**
   * 再開情報を保存
   * @param {string} batchId - バッチID
   * @param {Object} resumeData - 再開データ
   * @return {boolean} 保存成功かどうか
   */
  saveResumeInfo(batchId, resumeData) {
    try {
      const currentTime = new Date();
      const expiryTime = new Date(currentTime.getTime() + CONFIG.BATCH_PROCESSING.RESUME_TOKEN_EXPIRY);
      
      // 既存の再開情報を検索
      const rowIndex = this.findResumeInfoRowIndex(batchId);
      
      const row = [
        batchId,
        JSON.stringify(resumeData),
        rowIndex === -1 ? currentTime : this.resumeInfoSheet.getRange(rowIndex, 3).getValue(),
        currentTime,
        true,
        expiryTime,
        JSON.stringify({
          createdBy: Session.getActiveUser().getEmail(),
          version: '1.0',
          resumeCount: resumeData.resumeCount || 0
        })
      ];
      
      if (rowIndex === -1) {
        this.resumeInfoSheet.appendRow(row);
      } else {
        this.resumeInfoSheet.getRange(rowIndex, 1, 1, 7).setValues([row]);
      }
      
      log('INFO', `再開情報を保存しました: バッチID=${batchId}`);
      
      return true;
      
    } catch (error) {
      log('ERROR', '再開情報保存エラー', error);
      return false;
    }
  }
  
  /**
   * 再開情報を取得
   * @param {string} batchId - バッチID
   * @return {Object|null} 再開情報
   */
  getResumeInfo(batchId) {
    try {
      const rowIndex = this.findResumeInfoRowIndex(batchId);
      if (rowIndex === -1) {
        return null;
      }
      
      const row = this.resumeInfoSheet.getRange(rowIndex, 1, 1, 7).getValues()[0];
      
      // 有効期限チェック
      const expiryTime = new Date(row[5]);
      if (new Date() > expiryTime) {
        // 期限切れの場合は無効化
        this.resumeInfoSheet.getRange(rowIndex, 5).setValue(false);
        return null;
      }
      
      return {
        batchId: row[0],
        resumeData: JSON.parse(row[1]),
        createdAt: row[2],
        updatedAt: row[3],
        isActive: row[4],
        expiryTime: row[5],
        metadata: JSON.parse(row[6] || '{}')
      };
      
    } catch (error) {
      log('ERROR', '再開情報取得エラー', error);
      return null;
    }
  }
  
  /**
   * バッチ履歴を取得
   * @param {Object} options - 取得オプション
   * @return {Array} バッチ履歴の配列
   */
  getBatchHistory(options = {}) {
    try {
      const lastRow = this.batchHistorySheet.getLastRow();
      if (lastRow <= 1) return [];
      
      const allData = this.batchHistorySheet.getRange(2, 1, lastRow - 1, 17).getValues();
      
      let records = allData.map(row => ({
        batchId: row[0],
        batchName: row[1],
        status: row[2],
        createdAt: row[3],
        updatedAt: row[4],
        totalFiles: row[5],
        completedFiles: row[6],
        failedFiles: row[7],
        targetLang: row[8],
        dictName: row[9],
        totalSourceChars: row[10],
        totalTargetChars: row[11],
        totalDuration: row[12],
        totalApiCost: row[13],
        createdBy: row[14],
        errorMessage: row[15],
        settings: JSON.parse(row[16] || '{}')
      }));
      
      // フィルタリング
      if (options.status) {
        records = records.filter(r => r.status === options.status);
      }
      if (options.createdBy) {
        records = records.filter(r => r.createdBy === options.createdBy);
      }
      if (options.targetLang) {
        records = records.filter(r => r.targetLang === options.targetLang);
      }
      
      // ソート（最新順）
      records.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
      
      // 件数制限
      return records.slice(0, options.limit || 50);
      
    } catch (error) {
      log('ERROR', 'バッチ履歴取得エラー', error);
      return [];
    }
  }
  
  /**
   * バッチのファイル履歴を取得
   * @param {string} batchId - バッチID
   * @return {Array} ファイル履歴の配列
   */
  getBatchFileHistory(batchId) {
    try {
      const lastRow = this.fileHistorySheet.getLastRow();
      if (lastRow <= 1) return [];
      
      const allData = this.fileHistorySheet.getRange(2, 1, lastRow - 1, 15).getValues();
      
      const records = allData
        .filter(row => row[1] === batchId) // バッチIDでフィルタ
        .map(row => ({
          jobId: row[0],
          batchId: row[1],
          sourceUrl: row[2],
          targetUrl: row[3],
          fileName: row[4],
          fileType: row[5],
          status: row[6],
          charCount: row[7],
          startTime: row[8],
          completedTime: row[9],
          errorMessage: row[10],
          resumeToken: row[11],
          duration: row[12],
          apiCost: row[13],
          retryCount: row[14]
        }));
      
      // 開始時刻順でソート
      records.sort((a, b) => new Date(a.startTime) - new Date(b.startTime));
      
      return records;
      
    } catch (error) {
      log('ERROR', 'バッチファイル履歴取得エラー', error);
      return [];
    }
  }
  
  /**
   * バッチ統計情報を取得
   * @param {string} batchId - バッチID（省略時は全体統計）
   * @return {Object} 統計情報
   */
  getBatchStatistics(batchId = null) {
    try {
      const batchRecords = batchId ? 
        this.getBatchHistory({ limit: 1 }).filter(r => r.batchId === batchId) :
        this.getBatchHistory({ limit: 1000 });
      
      if (batchRecords.length === 0) {
        return this.getEmptyBatchStatistics();
      }
      
      const stats = {
        totalBatches: batchRecords.length,
        completedBatches: 0,
        failedBatches: 0,
        processingBatches: 0,
        totalFiles: 0,
        completedFiles: 0,
        failedFiles: 0,
        totalSourceChars: 0,
        totalTargetChars: 0,
        totalDuration: 0,
        totalApiCost: 0,
        averageFilesPerBatch: 0,
        successRate: 0,
        averageProcessingTime: 0,
        languageDistribution: {},
        statusDistribution: {}
      };
      
      batchRecords.forEach(batch => {
        // ステータス統計
        switch (batch.status) {
          case 'completed':
            stats.completedBatches++;
            break;
          case 'failed':
            stats.failedBatches++;
            break;
          case 'processing':
            stats.processingBatches++;
            break;
        }
        
        // 基本統計
        stats.totalFiles += parseInt(batch.totalFiles) || 0;
        stats.completedFiles += parseInt(batch.completedFiles) || 0;
        stats.failedFiles += parseInt(batch.failedFiles) || 0;
        stats.totalSourceChars += parseInt(batch.totalSourceChars) || 0;
        stats.totalTargetChars += parseInt(batch.totalTargetChars) || 0;
        stats.totalDuration += parseFloat(batch.totalDuration) || 0;
        stats.totalApiCost += parseFloat(batch.totalApiCost) || 0;
        
        // 言語分布
        if (batch.targetLang) {
          stats.languageDistribution[batch.targetLang] = 
            (stats.languageDistribution[batch.targetLang] || 0) + 1;
        }
        
        // ステータス分布
        stats.statusDistribution[batch.status] = 
          (stats.statusDistribution[batch.status] || 0) + 1;
      });
      
      // 計算統計
      if (stats.totalBatches > 0) {
        stats.averageFilesPerBatch = Math.round(stats.totalFiles / stats.totalBatches);
        stats.successRate = Math.round((stats.completedBatches / stats.totalBatches) * 100);
        stats.averageProcessingTime = Math.round(stats.totalDuration / stats.totalBatches);
      }
      
      return stats;
      
    } catch (error) {
      log('ERROR', 'バッチ統計取得エラー', error);
      return this.getEmptyBatchStatistics();
    }
  }
  
  /**
   * 空のバッチ統計情報を返す
   * @return {Object} 空の統計情報
   */
  getEmptyBatchStatistics() {
    return {
      totalBatches: 0,
      completedBatches: 0,
      failedBatches: 0,
      processingBatches: 0,
      totalFiles: 0,
      completedFiles: 0,
      failedFiles: 0,
      totalSourceChars: 0,
      totalTargetChars: 0,
      totalDuration: 0,
      totalApiCost: 0,
      averageFilesPerBatch: 0,
      successRate: 0,
      averageProcessingTime: 0,
      languageDistribution: {},
      statusDistribution: {}
    };
  }
  
  /**
   * バッチIDを生成
   * @return {string} バッチID
   */
  generateBatchId() {
    const timestamp = new Date().getTime();
    const random = Math.floor(Math.random() * 1000);
    return `BATCH_${timestamp}_${random}`;
  }
  
  /**
   * ジョブIDを生成
   * @return {string} ジョブID
   */
  generateJobId() {
    const timestamp = new Date().getTime();
    const random = Math.floor(Math.random() * 1000);
    return `JOB_${timestamp}_${random}`;
  }
  
  /**
   * バッチの行インデックスを検索
   * @param {string} batchId - バッチID
   * @return {number} 行インデックス（見つからない場合は-1）
   */
  findBatchRowIndex(batchId) {
    try {
      const lastRow = this.batchHistorySheet.getLastRow();
      if (lastRow <= 1) return -1;
      
      const batchIds = this.batchHistorySheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < batchIds.length; i++) {
        if (batchIds[i][0] === batchId) {
          return i + 2; // 1-indexed + header row
        }
      }
      return -1;
    } catch (error) {
      log('ERROR', 'バッチ行検索エラー', error);
      return -1;
    }
  }
  
  /**
   * ファイルジョブの行インデックスを検索
   * @param {string} jobId - ジョブID
   * @return {number} 行インデックス（見つからない場合は-1）
   */
  findFileJobRowIndex(jobId) {
    try {
      const lastRow = this.fileHistorySheet.getLastRow();
      if (lastRow <= 1) return -1;
      
      const jobIds = this.fileHistorySheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < jobIds.length; i++) {
        if (jobIds[i][0] === jobId) {
          return i + 2; // 1-indexed + header row
        }
      }
      return -1;
    } catch (error) {
      log('ERROR', 'ファイルジョブ行検索エラー', error);
      return -1;
    }
  }
  
  /**
   * 再開情報の行インデックスを検索
   * @param {string} batchId - バッチID
   * @return {number} 行インデックス（見つからない場合は-1）
   */
  findResumeInfoRowIndex(batchId) {
    try {
      const lastRow = this.resumeInfoSheet.getLastRow();
      if (lastRow <= 1) return -1;
      
      const batchIds = this.resumeInfoSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < batchIds.length; i++) {
        if (batchIds[i][0] === batchId) {
          return i + 2; // 1-indexed + header row
        }
      }
      return -1;
    } catch (error) {
      log('ERROR', '再開情報行検索エラー', error);
      return -1;
    }
  }
  
  /**
   * バッチ履歴の条件付き書式を適用
   */
  applyBatchConditionalFormatting() {
    try {
      const lastRow = this.batchHistorySheet.getLastRow();
      if (lastRow <= 1) return;
      
      const statusRange = this.batchHistorySheet.getRange(2, 3, lastRow - 1, 1);
      
      // 既存のルールをクリア
      statusRange.clearFormat();
      
      // 完了は緑
      const completedRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('completed')
        .setBackground('#d9ead3')
        .setRanges([statusRange])
        .build();
      
      // 失敗は赤
      const failedRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('failed')
        .setBackground('#f4cccc')
        .setRanges([statusRange])
        .build();
      
      // 処理中は黄色
      const processingRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('processing')
        .setBackground('#fff2cc')
        .setRanges([statusRange])
        .build();
      
      // 一時停止は灰色
      const pausedRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('paused')
        .setBackground('#efefef')
        .setRanges([statusRange])
        .build();
      
      // ルールを適用
      const rules = this.batchHistorySheet.getConditionalFormatRules();
      rules.push(completedRule, failedRule, processingRule, pausedRule);
      this.batchHistorySheet.setConditionalFormatRules(rules);
      
    } catch (error) {
      log('WARN', 'バッチ条件付き書式適用エラー', error);
    }
  }
  
  /**
   * ファイル履歴の条件付き書式を適用
   */
  applyFileConditionalFormatting() {
    try {
      const lastRow = this.fileHistorySheet.getLastRow();
      if (lastRow <= 1) return;
      
      const statusRange = this.fileHistorySheet.getRange(2, 7, lastRow - 1, 1);
      
      // 既存のルールをクリア
      statusRange.clearFormat();
      
      // 完了は緑
      const completedRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('completed')
        .setBackground('#d9ead3')
        .setRanges([statusRange])
        .build();
      
      // 失敗は赤
      const failedRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('failed')
        .setBackground('#f4cccc')
        .setRanges([statusRange])
        .build();
      
      // 処理中は黄色
      const processingRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('processing')
        .setBackground('#fff2cc')
        .setRanges([statusRange])
        .build();
      
      // 再試行は薄い黄色
      const retryingRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('retrying')
        .setBackground('#fce5cd')
        .setRanges([statusRange])
        .build();
      
      // スキップは薄い灰色
      const skippedRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('skipped')
        .setBackground('#f3f3f3')
        .setRanges([statusRange])
        .build();
      
      // ルールを適用
      const rules = this.fileHistorySheet.getConditionalFormatRules();
      rules.push(completedRule, failedRule, processingRule, retryingRule, skippedRule);
      this.fileHistorySheet.setConditionalFormatRules(rules);
      
    } catch (error) {
      log('WARN', 'ファイル条件付き書式適用エラー', error);
    }
  }
  
  /**
   * 古い再開情報をクリーンアップ
   * @return {number} クリーンアップした件数
   */
  cleanupExpiredResumeInfo() {
    try {
      const lastRow = this.resumeInfoSheet.getLastRow();
      if (lastRow <= 1) return 0;
      
      const currentTime = new Date();
      const allData = this.resumeInfoSheet.getRange(2, 1, lastRow - 1, 7).getValues();
      
      let cleanupCount = 0;
      const rowsToDelete = [];
      
      allData.forEach((row, index) => {
        const expiryTime = new Date(row[5]);
        if (currentTime > expiryTime) {
          rowsToDelete.push(index + 2); // 行番号（1-indexed）
          cleanupCount++;
        }
      });
      
      // 後ろから削除（行番号がずれないように）
      rowsToDelete.reverse().forEach(rowIndex => {
        this.resumeInfoSheet.deleteRows(rowIndex, 1);
      });
      
      if (cleanupCount > 0) {
        log('INFO', `期限切れの再開情報をクリーンアップしました: ${cleanupCount}件`);
      }
      
      return cleanupCount;
      
    } catch (error) {
      log('ERROR', '再開情報クリーンアップエラー', error);
      return 0;
    }
  }
}
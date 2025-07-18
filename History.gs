// History.gs - 履歴管理クラス

/**
 * 翻訳履歴の記録と管理を行うクラス
 */
class History {
  constructor() {
    this.spreadsheetId = CONFIG.HISTORY_SHEET_ID;
    this.spreadsheet = null;
    this.sheet = null;
    this.initSpreadsheet();
  }
  
  /**
   * スプレッドシートを初期化
   */
  initSpreadsheet() {
    try {
      this.spreadsheet = SpreadsheetApp.openById(this.spreadsheetId);
      this.sheet = this.getOrCreateHistorySheet();
    } catch (error) {
      log('ERROR', '履歴スプレッドシートの初期化エラー', error);
      throw new Error('履歴スプレッドシートにアクセスできません');
    }
  }
  
  /**
   * 履歴シートを取得または作成
   * @return {Sheet} 履歴シート
   */
  getOrCreateHistorySheet() {
    let sheet = this.spreadsheet.getSheetByName('翻訳履歴');
    
    if (!sheet) {
      sheet = this.spreadsheet.insertSheet('翻訳履歴');
      
      // ヘッダー設定
      const headers = [
        'ID',
        'タイムスタンプ',
        '翻訳元URL',
        '翻訳先URL',
        'ファイル名',
        'ファイルタイプ',
        '原文言語',
        '翻訳先言語',
        '使用辞書',
        '原文文字数',
        '翻訳後文字数',
        '文字数比率',
        '処理時間（秒）',
        'APIコスト（推定）',
        'ステータス',
        'エラーメッセージ',
        'ユーザー'
      ];
      
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);
      
      // スタイル設定
      headerRange.setBackground('#34a853');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
      
      // 列幅の調整
      sheet.setColumnWidth(1, 100);  // ID
      sheet.setColumnWidth(2, 150);  // タイムスタンプ
      sheet.setColumnWidth(3, 300);  // 翻訳元URL
      sheet.setColumnWidth(4, 300);  // 翻訳先URL
      sheet.setColumnWidth(5, 200);  // ファイル名
      sheet.setColumnWidth(6, 120);  // ファイルタイプ
      sheet.setColumnWidth(7, 100);  // 原文言語
      sheet.setColumnWidth(8, 100);  // 翻訳先言語
      sheet.setColumnWidth(9, 150);  // 使用辞書
      sheet.setColumnWidth(10, 100); // 原文文字数
      sheet.setColumnWidth(11, 100); // 翻訳後文字数
      sheet.setColumnWidth(12, 100); // 文字数比率
      sheet.setColumnWidth(13, 100); // 処理時間
      sheet.setColumnWidth(14, 100); // APIコスト
      sheet.setColumnWidth(15, 100); // ステータス
      sheet.setColumnWidth(16, 200); // エラーメッセージ
      sheet.setColumnWidth(17, 150); // ユーザー
      
      // フリーズ設定
      sheet.setFrozenRows(1);
      
      // データ検証を設定
      this.setupDataValidation(sheet);
    }
    
    return sheet;
  }
  
  /**
   * データ検証を設定
   * @param {Sheet} sheet - 対象シート
   */
  setupDataValidation(sheet) {
    // ステータスの検証
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['success', 'error', 'cancelled', 'processing'])
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 15, 1000, 1).setDataValidation(statusRule);
    
    // ファイルタイプの検証
    const fileTypeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['spreadsheet', 'document', 'presentation'])
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 6, 1000, 1).setDataValidation(fileTypeRule);
  }
  
  /**
   * 履歴を記録
   * @param {Object} data - 記録するデータ
   * @return {string} 記録したレコードのID
   */
  record(data) {
    try {
      const recordId = this.generateRecordId();
      const fileName = this.extractFileName(data.sourceUrl);
      const fileType = this.detectFileType(data.sourceUrl);
      const charRatio = data.charCountSource > 0 ? 
        Math.round((data.charCountTarget / data.charCountSource) * 100) : 0;
      const estimatedCost = this.estimateApiCost(data.charCountSource);
      const currentUser = Session.getActiveUser().getEmail();
      
      const row = [
        recordId,
        new Date(),
        data.sourceUrl || '',
        data.targetUrl || '',
        fileName,
        fileType,
        data.sourceLang || CONFIG.DEFAULT_SOURCE_LANG,
        data.targetLang || '',
        data.dictName || 'なし',
        data.charCountSource || 0,
        data.charCountTarget || 0,
        `${charRatio}%`,
        data.duration || 0,
        estimatedCost,
        data.status || 'unknown',
        data.errorMessage || '',
        currentUser
      ];
      
      this.sheet.appendRow(row);
      
      // 条件付き書式を適用
      this.applyConditionalFormatting();
      
      // 統計情報を更新
      this.updateStatistics(data);
      
      log('INFO', `履歴を記録しました: ID=${recordId}`);
      
      return recordId;
      
    } catch (error) {
      log('ERROR', '履歴記録エラー', error);
      throw error;
    }
  }
  
  /**
   * 最近の履歴を取得
   * @param {number} limit - 取得する件数
   * @param {Object} filter - フィルター条件（オプション）
   * @return {Array} 履歴データの配列
   */
  getRecent(limit = 50, filter = {}) {
    try {
      const lastRow = this.sheet.getLastRow();
      if (lastRow <= 1) return []; // ヘッダーのみ
      
      // 全データを取得（フィルタリングのため）
      const allData = this.sheet.getRange(2, 1, lastRow - 1, 17).getValues();
      
      // データをオブジェクトに変換
      let records = allData.map(row => ({
        id: row[0],
        timestamp: row[1] ? new Date(row[1]).toISOString() : null, // DateをISO文字列に変換
        sourceUrl: row[2],
        targetUrl: row[3],
        fileName: row[4],
        fileType: row[5],
        sourceLang: row[6],
        targetLang: row[7],
        dictName: row[8],
        charCountSource: row[9],
        charCountTarget: row[10],
        charRatio: row[11],
        duration: row[12],
        apiCost: row[13],
        status: row[14],
        errorMessage: row[15],
        user: row[16]
      }));
      
      // フィルタリング
      if (filter.status) {
        records = records.filter(r => r.status === filter.status);
      }
      if (filter.targetLang) {
        records = records.filter(r => r.targetLang === filter.targetLang);
      }
      if (filter.fileType) {
        records = records.filter(r => r.fileType === filter.fileType);
      }
      if (filter.user) {
        records = records.filter(r => r.user === filter.user);
      }
      if (filter.dateFrom) {
        records = records.filter(r => new Date(r.timestamp) >= new Date(filter.dateFrom));
      }
      if (filter.dateTo) {
        records = records.filter(r => new Date(r.timestamp) <= new Date(filter.dateTo));
      }
      
      // ソート（最新順）
      records.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
      
      // 件数制限
      return records.slice(0, limit);
      
    } catch (error) {
      log('ERROR', '履歴取得エラー', error);
      return [];
    }
  }
  
  /**
   * 履歴の統計情報を取得
   * @param {string} period - 期間（today, week, month, all）
   * @return {Object} 統計情報
   */
  getStatistics(period = 'all') {
    try {
      // パフォーマンスのため、最大取得件数を制限
      const maxRecords = period === 'all' ? 1000 : 500;
      const records = this.getRecent(maxRecords);
      
      if (!records || records.length === 0) {
        return this.getEmptyStatistics(period);
      }
      
      // 期間でフィルタリング
      const now = new Date();
      let filteredRecords = records;
      
      switch (period) {
        case 'today':
          const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
          filteredRecords = records.filter(r => new Date(r.timestamp) >= today);
          break;
          
        case 'week':
          const weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
          filteredRecords = records.filter(r => new Date(r.timestamp) >= weekAgo);
          break;
          
        case 'month':
          const monthAgo = new Date(now.getFullYear(), now.getMonth() - 1, now.getDate());
          filteredRecords = records.filter(r => new Date(r.timestamp) >= monthAgo);
          break;
      }
      
      // 統計計算
      const stats = {
        period: period,
        totalTranslations: filteredRecords.length,
        successCount: 0,
        errorCount: 0,
        totalSourceChars: 0,
        totalTargetChars: 0,
        totalDuration: 0,
        totalApiCost: 0,
        languageStats: {},
        fileTypeStats: {},
        dictionaryStats: {},
        userStats: {},
        hourlyDistribution: new Array(24).fill(0),
        averageCharRatio: 0,
        mostTranslatedFiles: []
      };
      
      // 詳細統計の計算（パフォーマンスを考慮）
      const fileCount = {};
      
      filteredRecords.forEach(record => {
        // 基本統計
        if (record.status === 'success') {
          stats.successCount++;
        } else if (record.status === 'error') {
          stats.errorCount++;
        }
        
        // 文字数（数値として扱う）
        const sourceChars = parseInt(record.charCountSource) || 0;
        const targetChars = parseInt(record.charCountTarget) || 0;
        stats.totalSourceChars += sourceChars;
        stats.totalTargetChars += targetChars;
        
        // 処理時間（数値として扱う）
        const duration = parseFloat(record.duration) || 0;
        stats.totalDuration += duration;
        
        // APIコスト（文字列の場合も考慮）
        const apiCost = parseFloat(record.apiCost) || 0;
        stats.totalApiCost += apiCost;
        
        // 言語統計（上位のみ）
        if (Object.keys(stats.languageStats).length < 10) {
          const langKey = `${record.sourceLang}_to_${record.targetLang}`;
          stats.languageStats[langKey] = (stats.languageStats[langKey] || 0) + 1;
        }
        
        // ファイルタイプ統計
        if (record.fileType) {
          stats.fileTypeStats[record.fileType] = (stats.fileTypeStats[record.fileType] || 0) + 1;
        }
        
        // 辞書使用統計（上位のみ）
        if (Object.keys(stats.dictionaryStats).length < 10 && record.dictName) {
          stats.dictionaryStats[record.dictName] = (stats.dictionaryStats[record.dictName] || 0) + 1;
        }
        
        // ユーザー統計（上位のみ）
        if (Object.keys(stats.userStats).length < 10 && record.user) {
          stats.userStats[record.user] = (stats.userStats[record.user] || 0) + 1;
        }
        
        // 時間帯分布
        try {
          const hour = new Date(record.timestamp).getHours();
          if (hour >= 0 && hour < 24) {
            stats.hourlyDistribution[hour]++;
          }
        } catch (e) {
          // 日付パースエラーは無視
        }
        
        // ファイル別カウント（成功のみ）
        if (record.status === 'success' && record.fileName) {
          fileCount[record.fileName] = (fileCount[record.fileName] || 0) + 1;
        }
      });
      
      // 平均文字数比率
      if (stats.totalSourceChars > 0) {
        stats.averageCharRatio = Math.round((stats.totalTargetChars / stats.totalSourceChars) * 100);
      }
      
      // 最も翻訳されたファイル（上位10件のみ）
      stats.mostTranslatedFiles = Object.entries(fileCount)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10)
        .map(([fileName, count]) => ({ fileName, count }));
      
      // 数値を適切にフォーマット
      stats.totalApiCost = Math.round(stats.totalApiCost * 10000) / 10000; // 小数点4桁
      stats.totalDuration = Math.round(stats.totalDuration);
      
      return stats;
      
    } catch (error) {
      log('ERROR', '統計情報取得エラー', error);
      return this.getEmptyStatistics(period);
    }
  }
  
  /**
   * 空の統計情報を返す
   * @param {string} period - 期間
   * @return {Object} 空の統計情報
   */
  getEmptyStatistics(period) {
    return {
      period: period,
      totalTranslations: 0,
      successCount: 0,
      errorCount: 0,
      totalSourceChars: 0,
      totalTargetChars: 0,
      totalDuration: 0,
      totalApiCost: 0,
      languageStats: {},
      fileTypeStats: {},
      dictionaryStats: {},
      userStats: {},
      hourlyDistribution: new Array(24).fill(0),
      averageCharRatio: 0,
      mostTranslatedFiles: []
    };
  }
  
  /**
   * 履歴をエクスポート
   * @param {Object} options - エクスポートオプション
   * @return {Object} エクスポートデータ
   */
  exportHistory(options = {}) {
    try {
      const records = this.getRecent(options.limit || 10000, options.filter || {});
      
      return {
        success: true,
        exportDate: new Date(),
        recordCount: records.length,
        records: records,
        version: '1.0'
      };
      
    } catch (error) {
      log('ERROR', '履歴エクスポートエラー', error);
      return {
        success: false,
        message: error.message
      };
    }
  }
  
  /**
   * 履歴をクリア
   * @param {Object} options - クリアオプション
   * @return {Object} 処理結果
   */
  clearHistory(options = {}) {
    try {
      const lastRow = this.sheet.getLastRow();
      if (lastRow <= 1) {
        return {
          success: true,
          message: '履歴はすでに空です',
          deletedCount: 0
        };
      }
      
      let rowsToDelete = lastRow - 1;
      
      // 期間指定がある場合
      if (options.beforeDate) {
        const beforeDate = new Date(options.beforeDate);
        const allData = this.sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // タイムスタンプ列
        
        const rowsToKeep = [];
        allData.forEach((row, index) => {
          if (new Date(row[0]) >= beforeDate) {
            rowsToKeep.push(index + 2); // 行番号（1-indexed）
          }
        });
        
        if (rowsToKeep.length === allData.length) {
          return {
            success: true,
            message: '削除対象のレコードがありません',
            deletedCount: 0
          };
        }
        
        // 保持する行を一時的に保存
        const keepData = rowsToKeep.map(rowNum => 
          this.sheet.getRange(rowNum, 1, 1, 17).getValues()[0]
        );
        
        // すべて削除してから保持するデータを再挿入
        this.sheet.deleteRows(2, lastRow - 1);
        
        if (keepData.length > 0) {
          const startRow = 2;
          this.sheet.getRange(startRow, 1, keepData.length, 17).setValues(keepData);
        }
        
        rowsToDelete = allData.length - keepData.length;
      } else {
        // 全削除
        this.sheet.deleteRows(2, rowsToDelete);
      }
      
      log('INFO', `${rowsToDelete}件の履歴を削除しました`);
      
      return {
        success: true,
        message: `${rowsToDelete}件の履歴を削除しました`,
        deletedCount: rowsToDelete
      };
      
    } catch (error) {
      log('ERROR', '履歴クリアエラー', error);
      return {
        success: false,
        message: error.message
      };
    }
  }
  
  /**
   * レコードIDを生成
   * @return {string} レコードID
   */
  generateRecordId() {
    const timestamp = new Date().getTime();
    const random = Math.floor(Math.random() * 1000);
    return `TR${timestamp}${random}`;
  }
  
  /**
   * URLからファイル名を抽出
   * @param {string} url - Google DriveのURL
   * @return {string} ファイル名
   */
  extractFileName(url) {
    try {
      const fileId = extractFileId(url);
      if (fileId) {
        const file = DriveApp.getFileById(fileId);
        return file.getName();
      }
    } catch (error) {
      // エラーの場合はURLの一部を返す
    }
    
    return 'Unknown File';
  }
  
  /**
   * URLからファイルタイプを検出
   * @param {string} url - Google DriveのURL
   * @return {string} ファイルタイプ
   */
  detectFileType(url) {
    if (url.includes('/spreadsheets/')) return 'spreadsheet';
    if (url.includes('/document/')) return 'document';
    if (url.includes('/presentation/')) return 'presentation';
    return 'unknown';
  }
  
  /**
   * APIコストを推定
   * @param {number} charCount - 文字数
   * @return {string} 推定コスト（ドル）
   */
  estimateApiCost(charCount) {
    // GPT-4o-miniの料金（仮定値）
    // 入力: $0.15 / 1M tokens
    // 出力: $0.60 / 1M tokens
    // 1 token ≈ 0.75文字（日本語の場合）
    
    const tokensIn = Math.ceil(charCount / 0.75);
    const tokensOut = Math.ceil(charCount / 0.75); // 同じ長さと仮定
    
    const costIn = (tokensIn / 1000000) * 0.15;
    const costOut = (tokensOut / 1000000) * 0.60;
    
    const totalCost = costIn + costOut;
    
    return totalCost.toFixed(4);
  }
  
  /**
   * 条件付き書式を適用
   */
  applyConditionalFormatting() {
    try {
      const lastRow = this.sheet.getLastRow();
      if (lastRow <= 1) return;
      
      // ステータス列の条件付き書式
      const statusRange = this.sheet.getRange(2, 15, lastRow - 1, 1);
      
      // 既存のルールをクリア
      statusRange.clearFormat();
      
      // 成功の場合は緑
      const successRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('success')
        .setBackground('#d9ead3')
        .setRanges([statusRange])
        .build();
      
      // エラーの場合は赤
      const errorRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('error')
        .setBackground('#f4cccc')
        .setRanges([statusRange])
        .build();
      
      // ルールを適用
      const rules = this.sheet.getConditionalFormatRules();
      rules.push(successRule, errorRule);
      this.sheet.setConditionalFormatRules(rules);
      
    } catch (error) {
      log('WARN', '条件付き書式の適用エラー', error);
    }
  }
  
  /**
   * 統計情報を更新（サマリーシートに記録）
   * @param {Object} data - 翻訳データ
   */
  updateStatistics(data) {
    try {
      // サマリーシートを取得または作成
      let summarySheet = this.spreadsheet.getSheetByName('統計サマリー');
      if (!summarySheet) {
        summarySheet = this.createSummarySheet();
      }
      
      // 今日の日付
      const today = new Date();
      const dateStr = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy-MM-dd');
      
      // 既存のレコードを検索
      const lastRow = summarySheet.getLastRow();
      let targetRow = -1;
      
      if (lastRow > 1) {
        const dates = summarySheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (let i = 0; i < dates.length; i++) {
          const cellDate = Utilities.formatDate(new Date(dates[i][0]), 'Asia/Tokyo', 'yyyy-MM-dd');
          if (cellDate === dateStr) {
            targetRow = i + 2;
            break;
          }
        }
      }
      
      // レコードの更新または追加
      if (targetRow > 0) {
        // 既存レコードを更新
        const currentValues = summarySheet.getRange(targetRow, 2, 1, 5).getValues()[0];
        const newValues = [
          (currentValues[0] || 0) + 1, // 翻訳回数
          (currentValues[1] || 0) + (data.charCountSource || 0), // 総文字数（原文）
          (currentValues[2] || 0) + (data.charCountTarget || 0), // 総文字数（翻訳）
          (currentValues[3] || 0) + (data.duration || 0), // 総処理時間
          (currentValues[4] || 0) + parseFloat(this.estimateApiCost(data.charCountSource || 0)) // 総コスト
        ];
        summarySheet.getRange(targetRow, 2, 1, 5).setValues([newValues]);
      } else {
        // 新規レコードを追加
        const newRow = [
          today,
          1, // 翻訳回数
          data.charCountSource || 0, // 総文字数（原文）
          data.charCountTarget || 0, // 総文字数（翻訳）
          data.duration || 0, // 総処理時間
          parseFloat(this.estimateApiCost(data.charCountSource || 0)) // 総コスト
        ];
        summarySheet.appendRow(newRow);
      }
      
    } catch (error) {
      log('WARN', '統計更新エラー', error);
    }
  }
  
  /**
   * サマリーシートを作成
   * @return {Sheet} サマリーシート
   */
  createSummarySheet() {
    const sheet = this.spreadsheet.insertSheet('統計サマリー');
    
    // ヘッダー設定
    const headers = ['日付', '翻訳回数', '総文字数（原文）', '総文字数（翻訳）', '総処理時間（秒）', '総コスト（USD）'];
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    
    // 列幅設定
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(4, 150);
    sheet.setColumnWidth(5, 150);
    sheet.setColumnWidth(6, 120);
    
    // フリーズ設定
    sheet.setFrozenRows(1);
    
    return sheet;
  }
}

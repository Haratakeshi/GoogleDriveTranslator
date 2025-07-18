// Dictionary.gs - 辞書管理クラス

/**
 * 翻訳辞書の管理を行うクラス
 */
class Dictionary {
  constructor() {
    this.spreadsheetId = CONFIG.DICTIONARY_SHEET_ID;
    this.spreadsheet = null;
    this.cache = new Map(); // パフォーマンス向上のためのキャッシュ
    this.initSpreadsheet();
  }
  
  /**
   * スプレッドシートを初期化
   */
  initSpreadsheet() {
    try {
      this.spreadsheet = SpreadsheetApp.openById(this.spreadsheetId);
    } catch (error) {
      log('ERROR', '辞書スプレッドシートの初期化に失敗しました。IDまたは権限を確認してください。', {
        spreadsheetId: this.spreadsheetId,
        errorMessage: error.message
      });
      this.spreadsheet = null; // エラーが発生した場合はnullを設定
    }
  }
  
  /**
   * 辞書一覧を取得
   * @return {Array} 辞書情報の配列
   */
  getList() {
    try {
      // スプレッドシートの存在確認
      if (!this.spreadsheet) {
        log('ERROR', 'スプレッドシートが初期化されていません');
        return [];
      }
      
      const sheets = this.spreadsheet.getSheets();
      if (!sheets || sheets.length === 0) {
        log('WARN', 'シートが見つかりません');
        return [];
      }
      
      const dictList = [];
      
      sheets.forEach(sheet => {
        try {
          const name = sheet.getName();
          const lastRow = sheet.getLastRow();
          const termCount = lastRow > 1 ? lastRow - 1 : 0; // ヘッダー行を除く
          
          // 辞書の統計情報を取得（エラーを避けるため簡略化）
          const stats = {
            lastUpdated: null,
            mostUsedTerms: []
          };
          
          // 統計情報の取得を試みる（エラーが発生しても続行）
          try {
            const detailedStats = this.getDictionaryStats(sheet);
            stats.lastUpdated = detailedStats.lastUpdated;
            stats.mostUsedTerms = detailedStats.mostUsedTerms;
          } catch (e) {
            log('WARN', `統計情報の取得に失敗: ${name}`, e);
          }
          
          dictList.push({
            name: name,
            termCount: termCount,
            // DateオブジェクトをISO文字列に変換してシリアライズ問題を回避
            lastUpdated: stats.lastUpdated ? stats.lastUpdated.toISOString() : null,
            mostUsedTerms: stats.mostUsedTerms
          });
        } catch (sheetError) {
          log('WARN', 'シート情報の取得に失敗', sheetError);
        }
      });
      
      // 名前順でソート
      dictList.sort((a, b) => a.name.localeCompare(b.name));
      
      log('INFO', `${dictList.length}個の辞書を取得しました`);
      return dictList;
      
    } catch (error) {
      log('ERROR', '辞書リスト取得エラー', error);
      // エラーが発生しても空配列を返す（nullではなく）
      return [];
    }
  }
  
  /**
   * 新規辞書を作成
   * @param {string} name - 辞書名
   * @return {Object} 処理結果
   */
  create(name) {
    try {
      // 名前の検証
      if (!this.validateDictionaryName(name)) {
        throw new Error('辞書名に使用できない文字が含まれています');
      }
      
      // 既存チェック
      const existingSheet = this.spreadsheet.getSheetByName(name);
      if (existingSheet) {
        throw new Error(`辞書「${name}」は既に存在します`);
      }
      
      // 新しいシートを作成
      const newSheet = this.spreadsheet.insertSheet(name);
      
      // ヘッダー行を設定
      const headers = ['原語', '訳語', '品詞', '備考', '登録日時', '使用回数'];
      const headerRange = newSheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);
      
      // ヘッダーのスタイル設定
      headerRange.setBackground('#4a86e8');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
      
      // 列幅の設定
      newSheet.setColumnWidth(1, 200); // 原語
      newSheet.setColumnWidth(2, 200); // 訳語
      newSheet.setColumnWidth(3, 100); // 品詞
      newSheet.setColumnWidth(4, 250); // 備考
      newSheet.setColumnWidth(5, 150); // 登録日時
      newSheet.setColumnWidth(6, 100); // 使用回数
      
      // フリーズ設定
      newSheet.setFrozenRows(1);
      
      // データ検証を設定（品詞）
      const posRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['名詞', '動詞', '形容詞', '副詞', '前置詞', '接続詞', 'その他'])
        .setAllowInvalid(true)
        .build();
      newSheet.getRange(2, 3, 1000, 1).setDataValidation(posRule);
      
      log('INFO', `辞書「${name}」を作成しました`);
      
      // キャッシュをクリア
      this.cache.clear();
      
      return {
        success: true,
        message: `辞書「${name}」を作成しました`
      };
      
    } catch (error) {
      log('ERROR', '辞書作成エラー', error);
      return {
        success: false,
        message: error.message
      };
    }
  }
  
  /**
   * 辞書から用語を検索（単一テキスト用）
   * @param {string} dictName - 辞書名
   * @param {string} text - 検索対象テキスト
   * @return {Array} マッチした用語の配列
   */
  getTerms(dictName, text) {
    if (!dictName || !text) return [];
    
    try {
      // キャッシュチェック
      const cacheKey = `${dictName}_terms`;
      let allTerms = this.cache.get(cacheKey);
      
      if (!allTerms) {
        // キャッシュミスの場合はスプレッドシートから取得
        const sheet = this.spreadsheet.getSheetByName(dictName);
        if (!sheet) {
          log('WARN', `辞書「${dictName}」が見つかりません`);
          return [];
        }
        
        const lastRow = sheet.getLastRow();
        if (lastRow <= 1) return []; // ヘッダーのみ
        
        const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
        allTerms = data.map((row, index) => ({
          source: row[0],
          target: row[1],
          pos: row[2],
          notes: row[3],
          rowIndex: index + 2 // スプレッドシートの行番号
        })).filter(term => term.source && term.target);
        
        // キャッシュに保存（5分間）
        this.cache.set(cacheKey, allTerms);
        Utilities.sleep(10); // API制限対策
      }
      
      // テキストに含まれる用語を検索
      const foundTerms = [];
      const lowerText = text.toLowerCase();
      
      allTerms.forEach(term => {
        // 大文字小文字を無視して検索
        if (lowerText.includes(term.source.toLowerCase())) {
          foundTerms.push({
            source: term.source,
            target: term.target,
            pos: term.pos,
            notes: term.notes,
            rowIndex: term.rowIndex
          });
        }
      });
      
      // 長い用語を優先（より具体的な翻訳を適用）
      foundTerms.sort((a, b) => b.source.length - a.source.length);
      
      // 使用回数を更新
      if (foundTerms.length > 0) {
        this.updateUsageCount(dictName, foundTerms);
      }
      
      return foundTerms;
      
    } catch (error) {
      log('ERROR', '用語検索エラー', error);
      return [];
    }
  }
  
  /**
   * 辞書から用語を検索（バッチ処理用）
   * @param {string} dictName - 辞書名
   * @param {Array} texts - 検索対象テキストの配列
   * @return {Array} マッチした用語の配列（重複除去済み）
   */
  getBatchTerms(dictName, texts) {
    if (!dictName || !texts || texts.length === 0) return [];
    
    try {
      // すべてのテキストを結合して検索
      const combinedText = texts.join(' ');
      const terms = this.getTerms(dictName, combinedText);
      
      // 重複を除去
      const uniqueTerms = new Map();
      terms.forEach(term => {
        const key = `${term.source}_${term.target}`;
        if (!uniqueTerms.has(key)) {
          uniqueTerms.set(key, term);
        }
      });
      
      return Array.from(uniqueTerms.values());
      
    } catch (error) {
      log('ERROR', 'バッチ用語検索エラー', error);
      return [];
    }
  }
  
  /**
   * 辞書に用語を追加
   * @param {string} dictName - 辞書名
   * @param {Array} terms - 追加する用語の配列
   * @return {Object} 処理結果
   */
  addTerms(dictName, terms) {
    if (!dictName || !terms || terms.length === 0) {
      return { success: true, added: 0 };
    }
    
    try {
      const sheet = this.spreadsheet.getSheetByName(dictName);
      if (!sheet) {
        log('WARN', `辞書「${dictName}」が見つかりません`);
        return { success: false, message: '辞書が見つかりません' };
      }
      
      // 既存の用語を取得
      const lastRow = sheet.getLastRow();
      const existingTerms = new Set();
      
      if (lastRow > 1) {
        const existingData = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
        existingData.forEach(row => {
          if (row[0] && row[1]) {
            existingTerms.add(`${row[0]}_${row[1]}`);
          }
        });
      }
      
      // 新規用語のみを抽出
      const newTerms = terms.filter(term => {
        const key = `${term.source}_${term.target}`;
        return !existingTerms.has(key);
      });
      
      if (newTerms.length === 0) {
        return { success: true, added: 0 };
      }
      
      // 新規用語を追加
      const newData = newTerms.map(term => [
        term.source || '',
        term.target || '',
        term.pos || '',
        term.notes || '',
        new Date(),
        0
      ]);
      
      const startRow = lastRow + 1;
      sheet.getRange(startRow, 1, newData.length, 6).setValues(newData);
      
      // キャッシュをクリア
      const cacheKey = `${dictName}_terms`;
      this.cache.delete(cacheKey);
      
      log('INFO', `${newTerms.length}個の新規用語を辞書「${dictName}」に追加しました`);
      
      return {
        success: true,
        added: newTerms.length
      };
      
    } catch (error) {
      log('ERROR', '用語追加エラー', error);
      return {
        success: false,
        message: error.message
      };
    }
  }
  
  /**
   * 使用回数を更新
   * @param {string} dictName - 辞書名
   * @param {Array} terms - 使用された用語の配列
   */
  updateUsageCount(dictName, terms) {
    try {
      const sheet = this.spreadsheet.getSheetByName(dictName);
      if (!sheet) return;
      
      // バッチ更新のためのデータを準備
      const updates = [];
      
      terms.forEach(term => {
        if (term.rowIndex) {
          const currentCountRange = sheet.getRange(term.rowIndex, 6);
          const currentCount = currentCountRange.getValue() || 0;
          updates.push({
            range: currentCountRange,
            value: currentCount + 1
          });
        }
      });
      
      // バッチ更新
      updates.forEach(update => {
        update.range.setValue(update.value);
      });
      
    } catch (error) {
      log('WARN', '使用回数更新エラー', error);
    }
  }
  
  /**
   * 辞書をエクスポート
   * @param {string} dictName - 辞書名
   * @return {Object} エクスポートデータ
   */
  export(dictName) {
    try {
      const sheet = this.spreadsheet.getSheetByName(dictName);
      if (!sheet) {
        throw new Error(`辞書「${dictName}」が見つかりません`);
      }
      
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) {
        return {
          success: true,
          dictName: dictName,
          terms: [],
          exportDate: new Date()
        };
      }
      
      const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
      const terms = data.map(row => ({
        source: row[0],
        target: row[1],
        pos: row[2],
        notes: row[3],
        createdAt: row[4] ? new Date(row[4]).toISOString() : null, // DateをISO文字列に変換
        usageCount: row[5]
      })).filter(term => term.source && term.target);
      
      log('INFO', `辞書「${dictName}」をエクスポートしました（${terms.length}語）`);
      
      return {
        success: true,
        dictName: dictName,
        terms: terms,
        exportDate: new Date().toISOString(), // DateをISO文字列に変換
        version: '1.0'
      };
      
    } catch (error) {
      log('ERROR', '辞書エクスポートエラー', error);
      return {
        success: false,
        message: error.message
      };
    }
  }
  
  /**
   * 辞書をインポート
   * @param {string} dictName - 辞書名
   * @param {Object} data - インポートデータ
   * @return {Object} 処理結果
   */
  import(dictName, data) {
    try {
      // データ検証
      if (!data || !data.terms || !Array.isArray(data.terms)) {
        throw new Error('無効なインポートデータです');
      }
      
      // 辞書が存在しない場合は作成
      let sheet = this.spreadsheet.getSheetByName(dictName);
      if (!sheet) {
        const createResult = this.create(dictName);
        if (!createResult.success) {
          throw new Error(createResult.message);
        }
        sheet = this.spreadsheet.getSheetByName(dictName);
      }
      
      // インポートオプション（既存データの扱い）
      const importMode = data.importMode || 'merge'; // merge, replace, append
      
      if (importMode === 'replace') {
        // 既存データを削除
        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          sheet.deleteRows(2, lastRow - 1);
        }
      }
      
      // 用語を追加
      const result = this.addTerms(dictName, data.terms);
      
      log('INFO', `辞書「${dictName}」にデータをインポートしました`);
      
      return {
        success: true,
        message: `${result.added}個の用語をインポートしました`,
        imported: result.added,
        skipped: data.terms.length - result.added
      };
      
    } catch (error) {
      log('ERROR', '辞書インポートエラー', error);
      return {
        success: false,
        message: error.message
      };
    }
  }
  
  /**
   * 辞書の統計情報を取得
   * @param {Sheet} sheet - シートオブジェクト
   * @return {Object} 統計情報
   */
  getDictionaryStats(sheet) {
    try {
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) {
        return {
          lastUpdated: null,
          mostUsedTerms: []
        };
      }
      
      // 最終更新日時を取得
      const dates = sheet.getRange(2, 5, lastRow - 1, 1).getValues();
      const validDates = dates.filter(row => row[0] instanceof Date).map(row => row[0]);
      const lastUpdated = validDates.length > 0 ? 
        new Date(Math.max(...validDates.map(d => d.getTime()))) : null;
      
      // 使用頻度の高い用語を取得
      const usageData = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
      const termsWithUsage = usageData
        .map((row, index) => ({
          source: row[0],
          target: row[1],
          usageCount: row[5] || 0
        }))
        .filter(term => term.source && term.usageCount > 0)
        .sort((a, b) => b.usageCount - a.usageCount)
        .slice(0, 5); // トップ5
      
      return {
        lastUpdated: lastUpdated,
        mostUsedTerms: termsWithUsage
      };
      
    } catch (error) {
      log('WARN', '統計情報取得エラー', error);
      return {
        lastUpdated: null,
        mostUsedTerms: []
      };
    }
  }
  
  /**
   * 辞書名の検証
   * @param {string} name - 辞書名
   * @return {boolean} 有効かどうか
   */
  validateDictionaryName(name) {
    if (!name || name.trim() === '') return false;
    
    // シート名として使用できない文字をチェック
    const invalidChars = /[:\\/\[\]\*\?]/;
    if (invalidChars.test(name)) return false;
    
    // 長さチェック（Googleスプレッドシートの制限）
    if (name.length > 100) return false;
    
    return true;
  }
  
  /**
   * 辞書を削除
   * @param {string} dictName - 辞書名
   * @return {Object} 処理結果
   */
  delete(dictName) {
    try {
      // デフォルト辞書は削除不可
      if (dictName === CONFIG.DEFAULT_DICT_NAME) {
        throw new Error('デフォルト辞書は削除できません');
      }
      
      const sheet = this.spreadsheet.getSheetByName(dictName);
      if (!sheet) {
        throw new Error(`辞書「${dictName}」が見つかりません`);
      }
      
      // 削除確認（重要な操作なので）
      this.spreadsheet.deleteSheet(sheet);
      
      // キャッシュをクリア
      this.cache.clear();
      
      log('INFO', `辞書「${dictName}」を削除しました`);
      
      return {
        success: true,
        message: `辞書「${dictName}」を削除しました`
      };
      
    } catch (error) {
      log('ERROR', '辞書削除エラー', error);
      return {
        success: false,
        message: error.message
      };
    }
  }
  
  /**
   * 辞書を複製
   * @param {string} sourceName - コピー元の辞書名
   * @param {string} targetName - コピー先の辞書名
   * @return {Object} 処理結果
   */
  duplicate(sourceName, targetName) {
    try {
      // 名前の検証
      if (!this.validateDictionaryName(targetName)) {
        throw new Error('辞書名に使用できない文字が含まれています');
      }
      
      const sourceSheet = this.spreadsheet.getSheetByName(sourceName);
      if (!sourceSheet) {
        throw new Error(`辞書「${sourceName}」が見つかりません`);
      }
      
      // 既存チェック
      if (this.spreadsheet.getSheetByName(targetName)) {
        throw new Error(`辞書「${targetName}」は既に存在します`);
      }
      
      // シートを複製
      const newSheet = sourceSheet.copyTo(this.spreadsheet);
      newSheet.setName(targetName);
      
      // キャッシュをクリア
      this.cache.clear();
      
      log('INFO', `辞書「${sourceName}」を「${targetName}」として複製しました`);
      
      return {
        success: true,
        message: `辞書を複製しました`
      };
      
    } catch (error) {
      log('ERROR', '辞書複製エラー', error);
      return {
        success: false,
        message: error.message
      };
    }
  }
}
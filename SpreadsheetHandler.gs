// SpreadsheetHandler.gs - Googleスプレッドシートの処理を担当するクラス

class SpreadsheetHandler {
  constructor() {
    this.spreadsheetApp = SpreadsheetApp;
  }

  /**
   * スプレッドシートから翻訳対象のジョブリストを作成する
   * @param {string} fileId - ファイルID
   * @return {Array} ジョブの配列
   */
  createTranslationJobs(fileId) {
    const spreadsheet = this.spreadsheetApp.openById(fileId);
    const sheets = spreadsheet.getSheets();
    const jobs = [];

    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      // 空のシートや非表示のシートはスキップ
      if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0 || sheet.isSheetHidden()) return;

      const range = sheet.getDataRange();
      const values = range.getValues();
      const formulas = range.getFormulas();

      for (let row = 0; row < values.length; row++) {
        for (let col = 0; col < values[row].length; col++) {
          const value = values[row][col];
          // テキストで、かつ数式でなく、空白でないセルのみを対象
          if (value && typeof value === 'string' && value.trim() !== '' && !formulas[row][col]) {
            jobs.push({
              type: 'cell',
              text: value,
              location: {
                sheetName: sheetName,
                range: `R${row + 1}C${col + 1}` // A1形式ではなくR1C1形式で指定
              }
            });
          }
        }
      }
    });
    log('INFO', `スプレッドシートから ${jobs.length} 件のジョブを作成しました`, { fileId });
    return jobs;
  }

  /**
   * 翻訳されたジョブをスプレッドシートに書き込む
   * @param {string} targetFileId - 書き込み先のファイルID
   * @param {Object} job - 書き込むジョブ情報
   * @param {string} translatedText - 翻訳されたテキスト
   */
  writeTranslatedJob(targetFileId, job, translatedText) {
    try {
      const sheet = this.spreadsheetApp.openById(targetFileId).getSheetByName(job.location.sheetName);
      if (sheet) {
        sheet.getRange(job.location.range).setValue(translatedText);
      }
    } catch(e) {
        log('ERROR', `スプレッドシートへのジョブの書き込みに失敗: ${targetFileId}`, {job, error: e.message});
    }
  }

  /**
   * スプレッドシートから全てのテキストコンテンツを抽出する
   * @param {string} fileId - ファイルID
   * @return {string} 抽出された全テキスト
   */
  extractAllText(fileId) {
    const spreadsheet = this.spreadsheetApp.openById(fileId);
    const sheets = spreadsheet.getSheets();
    const texts = [];

    sheets.forEach(sheet => {
       if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0 || sheet.isSheetHidden()) return;
      const range = sheet.getDataRange();
      const values = range.getValues();
      const formulas = range.getFormulas();

      for (let row = 0; row < values.length; row++) {
        for (let col = 0; col < values[row].length; col++) {
          const value = values[row][col];
          if (value && typeof value === 'string' && value.trim() !== '' && !formulas[row][col]) {
            texts.push(value);
          }
        }
      }
    });
    return texts.join('\n');
  }
}
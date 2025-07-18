// FileHandler.gs - ファイル操作クラス

/**
 * Google Driveファイルの読み書きを管理するクラス
 */
class FileHandler {
  constructor() {
    this.drive = DriveApp;
    this.supportedTypes = CONFIG.SUPPORTED_MIME_TYPES;
  }

  /**
   * ファイル情報を取得（コンテンツも含む）
   * @param {string} fileId - ファイルID
   * @return {Object} ファイル情報
   */
  getFileInfo(fileId) {
    try {
      const file = this.drive.getFileById(fileId);
      const mimeType = file.getMimeType();

      if (!this.supportedTypes[mimeType]) {
        throw new Error(CONFIG.ERROR_MESSAGES.UNSUPPORTED_FILE);
      }

      return {
        id: fileId,
        name: file.getName(),
        mimeType: mimeType,
        type: this.supportedTypes[mimeType].type,
        fileObject: file
      };

    } catch (error) {
      if (error.message.includes('not found')) {
        throw new Error(CONFIG.ERROR_MESSAGES.FILE_NOT_FOUND);
      } else if (error.message.includes('permission')) {
        throw new Error(CONFIG.ERROR_MESSAGES.NO_PERMISSION);
      }
      throw error;
    }
  }

  /**
   * 翻訳対象のジョブリストを作成する
   * @param {Object} fileInfo - 元のファイル情報
   * @return {Array} ジョブの配列
   */
  createTranslationJobs(fileInfo) {
    log('INFO', `ジョブ作成開始: ${fileInfo.name}`);
    switch (fileInfo.type) {
      case 'spreadsheet':
        return this.createSpreadsheetJobs(fileInfo.id);
      case 'document':
        return this.createDocumentJobs(fileInfo.id);
      case 'presentation':
        return this.createPresentationJobs(fileInfo.id);
      default:
        throw new Error('サポートされていないファイルタイプです');
    }
  }

  /**
   * ファイルから全てのテキストコンテンツを抽出する
   * @param {Object} fileInfo - ファイル情報
   * @return {string} 抽出された全テキスト
   */
  extractAllText(fileInfo) {
    log('INFO', `全テキスト抽出開始: ${fileInfo.name}`);
    try {
      switch (fileInfo.type) {
        case 'spreadsheet':
          return this._extractTextFromSpreadsheet(fileInfo.id);
        case 'document':
          return this._extractTextFromDocument(fileInfo.id);
        case 'presentation':
          return this._extractTextFromPresentation(fileInfo.id);
        default:
          throw new Error('サポートされていないファイルタイプです');
      }
    } catch (e) {
      log('ERROR', `テキスト抽出に失敗: ${fileInfo.name}`, { error: e.message });
      throw new Error(`ファイルからのテキスト抽出に失敗しました: ${e.message}`);
    }
  }

  /**
   * 翻訳先となる空のファイルを作成する
   * @param {Object} originalFileInfo - 元のファイル情報
   * @param {string} targetLang - 翻訳先言語
   * @return {Object} 作成されたファイルの情報 {id, url, name}
   */
  createEmptyTranslatedFile(originalFileInfo, targetLang) {
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
    const newFileName = `${originalFileInfo.name}_${CONFIG.SUPPORTED_LANGUAGES[targetLang]}_${timestamp}`;
    log('INFO', `翻訳ファイルのコピーを作成: ${newFileName}`);

    const originalFile = DriveApp.getFileById(originalFileInfo.id);
    const newFile = originalFile.makeCopy(newFileName);
    
    return {
      id: newFile.getId(),
      url: newFile.getUrl(),
      name: newFileName
    };
  }

  /**
   * 翻訳されたジョブをファイルに書き込む
   * @param {string} targetFileId - 書き込み先のファイルID
   * @param {string} fileType - ファイルの種類 ('spreadsheet', 'document', etc.)
   * @param {Object} job - 書き込むジョブ情報
   * @param {string} translatedText - 翻訳されたテキスト
   */
  writeTranslatedJob(targetFileId, fileType, job, translatedText) {
    try {
      switch (fileType) {
        case 'spreadsheet':
          const sheet = SpreadsheetApp.openById(targetFileId).getSheetByName(job.location.sheetName);
          if (sheet) {
            sheet.getRange(job.location.range).setValue(translatedText);
          }
          break;
        // TODO: Add cases for document and presentation
      }
    } catch(e) {
        log('ERROR', `ジョブの書き込みに失敗: ${targetFileId}`, {job, error: e.message});
        // Continue to next job
    }
  }

  // --- Private Job Creation Methods ---

  createSpreadsheetJobs(fileId) {
    const spreadsheet = SpreadsheetApp.openById(fileId);
    const sheets = spreadsheet.getSheets();
    const jobs = [];

    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      const range = sheet.getDataRange();
      const values = range.getValues();
      const formulas = range.getFormulas();

      for (let row = 0; row < values.length; row++) {
        for (let col = 0; col < values[row].length; col++) {
          const value = values[row][col];
          // テキストで、かつ数式でないセルのみを対象
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
    return jobs;
  }

  createDocumentJobs(fileId) {
    // TODO: Implement document job creation
    log('WARN', 'Document job creation is not yet implemented.');
    return [];
  }

  createPresentationJobs(fileId) {
    // TODO: Implement presentation job creation
    log('WARN', 'Presentation job creation is not yet implemented.');
    return [];
  }

  // --- Private Text Extraction Methods ---

  /**
   * スプレッドシートからテキストを抽出する
   * @private
   * @param {string} fileId - ファイルID
   * @return {string} 抽出されたテキスト
   */
  _extractTextFromSpreadsheet(fileId) {
    const spreadsheet = SpreadsheetApp.openById(fileId);
    const sheets = spreadsheet.getSheets();
    const texts = [];

    sheets.forEach(sheet => {
      const range = sheet.getDataRange();
      const values = range.getValues();
      const formulas = range.getFormulas();

      for (let row = 0; row < values.length; row++) {
        for (let col = 0; col < values[row].length; col++) {
          const value = values[row][col];
          // テキストで、かつ数式でないセルのみを対象
          if (value && typeof value === 'string' && value.trim() !== '' && !formulas[row][col]) {
            texts.push(value);
          }
        }
      }
    });
    return texts.join('\n');
  }

  /**
   * ドキュメントからテキストを抽出する
   * @private
   * @param {string} fileId - ファイルID
   * @return {string} 抽出されたテキスト
   */
  _extractTextFromDocument(fileId) {
    const doc = DocumentApp.openById(fileId);
    return doc.getBody().getText();
  }

  /**
   * プレゼンテーションからテキストを抽出する
   * @private
   * @param {string} fileId - ファイルID
   * @return {string} 抽出されたテキスト
   */
  _extractTextFromPresentation(fileId) {
    const presentation = SlidesApp.openById(fileId);
    const slides = presentation.getSlides();
    const texts = [];

    slides.forEach(slide => {
      // スライド上の全シェイプからテキストを抽出
      slide.getShapes().forEach(shape => {
        if (shape.hasText()) {
          texts.push(shape.getText().asString());
        }
      });
      // スピーカーノートからテキストを抽出
      const notesShape = slide.getNotesPage().getSpeakerNotesShape();
      if (notesShape && notesShape.hasText()) {
        texts.push(notesShape.getText().asString());
      }
    });
    return texts.join('\n');
  }
}

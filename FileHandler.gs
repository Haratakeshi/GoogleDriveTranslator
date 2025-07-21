// FileHandler.gs - ファイル操作の司令塔クラス

/**
 * Google Driveファイルの操作を統括し、ファイル種別に応じて
 * 専用のハンドラに処理を委譲するクラス
 */
class FileHandler {
  constructor() {
    this.drive = DriveApp;
    this.supportedTypes = CONFIG.SUPPORTED_MIME_TYPES;
    // 各ファイルタイプに対応するハンドラーをマッピング
    this.handlers = {
      spreadsheet: new SpreadsheetHandler(),
      document: new DocumentHandler(),
      presentation: new PresentationHandler(),
      pdf: new PdfHandler(), // PDFハンドラを追加
    };
  }

  /**
   * ファイル情報を取得する
   * @param {string} fileId - ファイルID
   * @return {Object} ファイル情報
   */
  getFileInfo(fileId) {
    try {
      const file = this.drive.getFileById(fileId);
      const mimeType = file.getMimeType();
      const typeInfo = this.supportedTypes[mimeType] || (mimeType === 'application/pdf' ? { type: 'pdf' } : null);

      if (!typeInfo) {
        throw new Error(CONFIG.ERROR_MESSAGES.UNSUPPORTED_FILE);
      }

      return {
        id: fileId,
        name: file.getName(),
        mimeType: mimeType,
        type: typeInfo.type,
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
   * 適切なハンドラを呼び出して翻訳対象のジョブリストを作成する
   * @param {Object} fileInfo - 元のファイル情報
   * @return {Array} ジョブの配列
   */
  createTranslationJobs(fileInfo) {
    const handler = this.handlers[fileInfo.type];
    if (handler) {
      log('INFO', `[${fileInfo.type}] ジョブ作成を開始: ${fileInfo.name}`);
      return handler.createTranslationJobs(fileInfo.id);
    }
    throw new Error(`サポートされていないファイルタイプです: ${fileInfo.type}`);
  }

  /**
   * 適切なハンドラを呼び出してファイルから全てのテキストコンテンツを抽出する
   * @param {Object} fileInfo - ファイル情報
   * @return {string} 抽出された全テキスト
   */
  extractAllText(fileInfo) {
    const handler = this.handlers[fileInfo.type];
    if (handler) {
      log('INFO', `[${fileInfo.type}] 全テキスト抽出を開始: ${fileInfo.name}`);
      return handler.extractAllText(fileInfo.id);
    }
    throw new Error(`サポートされていないファイルタイプです: ${fileInfo.type}`);
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
   * 適切なハンドラを呼び出して翻訳されたジョブをファイルに書き込む
   * @param {string} targetFileId - 書き込み先のファイルID
   * @param {string} fileType - ファイルの種類 ('spreadsheet', 'document', etc.)
   * @param {Object} job - 書き込むジョブ情報
   * @param {string} translatedText - 翻訳されたテキスト
   */
  writeTranslatedJob(targetFileId, fileType, job, translatedText) {
    const handler = this.handlers[fileType];
    if (handler) {
      handler.writeTranslatedJob(targetFileId, job, translatedText);
    } else {
      log('WARN', `書き込み処理がサポートされていないファイルタイプです: ${fileType}`);
    }
  }
}
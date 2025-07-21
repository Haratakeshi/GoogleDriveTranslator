// PdfHandler.gs - PDF翻訳のワークフローを管理するクラス

class PdfHandler {
  constructor() {
    // Drive API v3 を利用 (マニフェストで有効化が必要)
    this.driveApi = Drive;
    this.driveApp = DriveApp;
  }

  /**
   * PDF翻訳のメインプロセスを開始し、翻訳ジョブリストを生成する
   * @param {string} fileId - PDFファイルのID
   * @return {Array} DocumentHandlerが生成した翻訳ジョブ
   */
  createTranslationJobs(fileId) {
    log('INFO', `PDF翻訳プロセスを開始します: ${fileId}`);

    // 1. PDFをGoogleドキュメントに変換
    const tempDocId = this.convertPdfToDoc(fileId);
    if (!tempDocId) {
      throw new Error('PDFからGoogleドキュメントへの変換に失敗しました。');
    }
    log('INFO', `一時的なGoogleドキュメントを作成しました: ${tempDocId}`);

    // 2. DocumentHandlerにジョブ作成を委譲
    const documentHandler = new DocumentHandler();
    const jobs = documentHandler.createTranslationJobs(tempDocId);

    // 3. 後処理のために、元のPDF IDと一時ドキュメントIDを各ジョブに紐付ける
    jobs.forEach(job => {
      job.originalPdfId = fileId;
      job.tempDocId = tempDocId;
      job.handlerType = 'PdfHandler'; // 完了処理の際にどのハンドラを使うか示す
    });

    log('INFO', `PDFから ${jobs.length} 件の翻訳ジョブを作成しました。`);
    return jobs;
  }

  /**
   * 翻訳が完了したジョブの後処理（PDFエクスポートとクリーンアップ）を実行する
   * @param {string} tempDocId - 一時GoogleドキュメントのID
   * @param {string} originalPdfId - 元のPDFファイルのID
   */
  finalizeTranslation(tempDocId, originalPdfId) {
    try {
      log('INFO', `翻訳後処理を開始します: tempDocId=${tempDocId}, originalPdfId=${originalPdfId}`);
      // 1. 翻訳済みドキュメントをPDFとしてエクスポート
      this.exportDocAsPdf(tempDocId, originalPdfId);

      // 2. 一時ドキュメントを削除
      this.cleanup(tempDocId);
      log('INFO', `翻訳後処理が正常に完了しました。`);
    } catch (e) {
      log('ERROR', `PDF翻訳の最終処理に失敗しました: ${e.message}`, { tempDocId, originalPdfId });
    }
  }

  /**
   * PDFをGoogleドキュメントに変換する (OCRを利用)
   * @private
   * @param {string} pdfFileId - PDFファイルのID
   * @return {string|null} 新規作成されたGoogleドキュメントのID
   */
  convertPdfToDoc(pdfFileId) {
    try {
      const pdfFile = this.driveApi.Files.get(pdfFileId, { supportsAllDrives: true });
      const newFileName = pdfFile.name.replace(/\.pdf$/i, '');
      
      const resource = {
        name: newFileName,
        mimeType: 'application/vnd.google-apps.document'
      };

      const newDoc = this.driveApi.Files.copy(resource, pdfFileId, { supportsAllDrives: true });
      return newDoc.id;
    } catch (e) {
      log('ERROR', `Drive APIによるPDF変換に失敗: ${e.message}`, { pdfFileId });
      return null;
    }
  }

  /**
   * GoogleドキュメントをPDFとしてエクスポートする
   * @private
   * @param {string} docId - GoogleドキュメントのID
   * @param {string} originalPdfId - 元のPDFファイルのID
   */
  exportDocAsPdf(docId, originalPdfId) {
    try {
      const originalPdfFile = this.driveApp.getFileById(originalPdfId);
      const originalFolder = originalPdfFile.getParents().next();
      const translatedDocFile = this.driveApp.getFileById(docId);
      
      const pdfBlob = translatedDocFile.getAs('application/pdf');
      const newPdfName = `${originalPdfFile.getName().replace(/\.pdf$/i, '')}_translated.pdf`;
      
      const newPdfFile = originalFolder.createFile(pdfBlob).setName(newPdfName);
      log('INFO', `翻訳済みPDFを作成しました: ${newPdfFile.getId()}`, { name: newPdfName });
    } catch (e) {
      log('ERROR', `PDFエクスポートに失敗: ${e.message}`, { docId, originalPdfId });
    }
  }

  /**
   * 一時ファイルを削除する
   * @private
   * @param {string} fileId - 削除するファイルのID
   */
  cleanup(fileId) {
    try {
      // このgetはファイル名を取得するためだけなので、エラーが出ても処理を続行させる
      let fileName = 'unknown';
      try {
        const file = this.driveApi.Files.get(fileId, { supportsAllDrives: true });
        fileName = file.name;
      } catch (e) {
        log('WARN', `クリーンアップ中のファイル名取得に失敗: ${e.message}`, { fileId });
      }

      this.driveApi.Files.remove(fileId, { supportsAllDrives: true });
      log('INFO', `一時ファイル (${fileName}) を削除しました。`);
    } catch (e) {
      // ファイルが存在しない場合などのエラーは無視する
      log('WARN', `一時ファイルの削除に失敗しました（無視）: ${e.message}`, { fileId });
    }
  }
}

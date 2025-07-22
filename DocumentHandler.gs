// DocumentHandler.gs - Googleドキュメントの処理を担当するクラス

class DocumentHandler {
  constructor() {
    this.documentApp = DocumentApp;
  }

  /**
   * ドキュメントから翻訳対象のジョブリストを作成する（段落単位）
   * @param {string} fileId - ファイルID
   * @return {Array} ジョブの配列
   */
  createTranslationJobs(fileId) {
    const doc = this.documentApp.openById(fileId);
    const body = doc.getBody();
    const paragraphs = body.getParagraphs();
    const jobs = [];

    paragraphs.forEach((paragraph, index) => {
      const text = paragraph.getText();
      // 空でない段落のみをジョブの対象とする
      if (text.trim() !== '') {
        jobs.push({
          type: 'paragraph',
          text: text,
          location: {
            index: index // 段落のインデックスを保存
          }
        });
      }
    });

    log('INFO', `ドキュメントから ${jobs.length} 件のジョブを作成しました`, { fileId });
    return jobs;
  }

  /**
   * 翻訳されたジョブをドキュメントに書き込む
   * @param {string} targetFileId - 書き込み先のファイルID
   * @param {Object} job - 書き込むジョブ情報
   * @param {string} translatedText - 翻訳されたテキスト
   */
  writeTranslatedJob(targetFileId, job, translatedText) {
    try {
      const doc = this.documentApp.openById(targetFileId);
      const body = doc.getBody();
      const paragraphs = body.getParagraphs();
      const paragraphIndex = job.location.index;

      if (paragraphIndex < paragraphs.length) {
        const originalParagraph = paragraphs[paragraphIndex];
        // 画像などのオブジェクトを保持しながらテキストのみを置換
        this.replaceTextPreservingImages(originalParagraph, job.text, translatedText);
      } else {
        log('WARN', `段落インデックスが範囲外です: ${paragraphIndex}`, { targetFileId });
      }
    } catch (e) {
      log('ERROR', `ドキュメントへのジョブの書き込みに失敗: ${targetFileId}`, { job, error: e.message });
    }
  }

  /**
   * ドキュメントから全てのテキストコンテンツを抽出する
   * @param {string} fileId - ファイルID
   * @return {string} 抽出された全テキスト
   */
  extractAllText(fileId) {
    const doc = this.documentApp.openById(fileId);
    return doc.getBody().getText();
  }

  /**
   * 段落内の画像やオブジェクトを保持しながらテキストのみを置換する
   * @private
   * @param {Paragraph} paragraph - 置換対象の段落
   * @param {string} originalText - 元のテキスト
   * @param {string} translatedText - 翻訳されたテキスト
   */
  replaceTextPreservingImages(paragraph, originalText, translatedText) {
    const numChildren = paragraph.getNumChildren();
    
    // 段落内の要素を調べて、画像があるかチェック
    let hasImages = false;
    for (let i = 0; i < numChildren; i++) {
      const child = paragraph.getChild(i);
      if (child.getType() === this.documentApp.ElementType.INLINE_IMAGE || 
          child.getType() === this.documentApp.ElementType.INLINE_DRAWING) {
        hasImages = true;
        break;
      }
    }

    if (!hasImages) {
      // 画像がない場合は従来通りの処理
      paragraph.clear();
      paragraph.appendText(translatedText);
    } else {
      // 画像がある場合はテキスト要素のみを置換
      for (let i = 0; i < numChildren; i++) {
        const child = paragraph.getChild(i);
        if (child.getType() === this.documentApp.ElementType.TEXT) {
          const textElement = child.asText();
          const text = textElement.getText();
          // 元のテキストと一致する部分を翻訳されたテキストで置換
          if (text.trim() === originalText.trim()) {
            textElement.setText(translatedText);
          }
        }
      }
    }
  }
}
// PresentationHandler.gs - Googleスライドの処理を担当するクラス

class PresentationHandler {
  constructor() {
    this.slidesApp = SlidesApp;
  }

  /**
   * プレゼンテーションから翻訳対象のジョブリストを作成する（テキスト要素単位）
   * @param {string} fileId - ファイルID
   * @return {Array} ジョブの配列
   */
  createTranslationJobs(fileId) {
    const presentation = this.slidesApp.openById(fileId);
    const slides = presentation.getSlides();
    const jobs = [];

    slides.forEach((slide, slideIndex) => {
      // スライド上の全シェイプからテキストを持つものをジョブ化
      slide.getShapes().forEach(shape => {
        // Check if the shape supports text and has content
        if (typeof shape.getText === 'function') {
          const originalText = shape.getText().asString();
          if (originalText.trim() !== '') {
            jobs.push({
              type: 'shape',
              text: originalText,
              location: {
                slideIndex: slideIndex,
                shapeId: shape.getObjectId()
              }
            });
          }
        }
      });

      // スピーカーノートをジョブ化
      const notesPage = slide.getNotesPage();
      const notesShape = notesPage.getSpeakerNotesShape();
      if (notesShape && typeof notesShape.getText === 'function') {
        const originalText = notesShape.getText().asString();
        if (originalText.trim() !== '') {
          jobs.push({
            type: 'notes',
            text: originalText,
            location: {
              slideIndex: slideIndex
            }
          });
        }
      }
    });

    log('INFO', `プレゼンテーションから ${jobs.length} 件のジョブを作成しました`, { fileId });
    return jobs;
  }

  /**
   * 翻訳されたジョブをプレゼンテーションに書き込む
   * @param {string} targetFileId - 書き込み先のファイルID
   * @param {Object} job - 書き込むジョブ情報
   * @param {string} translatedText - 翻訳されたテキスト
   */
  writeTranslatedJob(targetFileId, job, translatedText) {
    try {
      const presentation = this.slidesApp.openById(targetFileId);
      const slides = presentation.getSlides();
      const slideIndex = job.location.slideIndex;

      if (slideIndex >= slides.length) {
        log('WARN', `スライドインデックスが範囲外です: ${slideIndex}`, { targetFileId });
        return;
      }
      const slide = slides[slideIndex];

      if (job.type === 'shape') {
        const element = slide.getPageElementById(job.location.shapeId);
        if (element && element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
          const shape = element.asShape();
          if (typeof shape.getText === 'function') {
            shape.getText().setText(translatedText);
          }
        } else {
          log('WARN', `シェイプが見つからないか、テキストをサポートしていません: ${job.location.shapeId}`, { targetFileId });
        }
      } else if (job.type === 'notes') {
        const notesPage = slide.getNotesPage();
        const notesShape = notesPage.getSpeakerNotesShape();
        if (notesShape) {
          notesShape.getText().setText(translatedText);
        }
      }
    } catch (e) {
      log('ERROR', `プレゼンテーションへのジョブの書き込みに失敗: ${targetFileId}`, { job, error: e.message });
    }
  }

  /**
   * プレゼンテーションから全てのテキストコンテンツを抽出する
   * @param {string} fileId - ファイルID
   * @return {string} 抽出された全テキスト
   */
  extractAllText(fileId) {
    const presentation = this.slidesApp.openById(fileId);
    const slides = presentation.getSlides();
    const texts = [];

    slides.forEach(slide => {
      slide.getShapes().forEach(shape => {
        if (typeof shape.getText === 'function') {
          texts.push(shape.getText().asString());
        }
      });
      const notesPage = slide.getNotesPage();
      const notesShape = notesPage.getSpeakerNotesShape();
      if (notesShape && typeof notesShape.getText === 'function') {
        texts.push(notesShape.getText().asString());
      }
    });
    return texts.join('\n');
  }
}
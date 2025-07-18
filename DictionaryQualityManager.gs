// DictionaryQualityManager.gs - 辞書品質管理機能

/**
 * 辞書品質管理クラス
 * 辞書の品質管理と用語の妥当性評価を行う機能を提供
 */
class DictionaryQualityManager {
  
  /**
   * コンストラクタ
   */
  constructor() {
    this.config = CONFIG;
    this.qualityThreshold = CONFIG.QUALITY_GATE_THRESHOLD || 0.7;
  }
  
  /**
   * 用語の品質を評価する
   * @param {Object} term - 評価対象の用語オブジェクト
   * @param {string} context - 文脈情報
   * @return {Object} 品質評価結果
   */
  evaluateTermQuality(term, context = '') {
    // TODO: 実装予定
    // 1. 信頼度スコア評価
    // 2. 文脈適合性チェック
    // 3. 承認・保留・却下判定
    
    return {
      quality_score: 0.0,
      decision: 'pending', // 'approved', 'pending', 'rejected'
      reasons: [],
      recommendations: []
    };
  }
  
  /**
   * 用語リストの品質を一括評価する
   * @param {Array} terms - 評価対象の用語リスト
   * @param {string} context - 文脈情報
   * @return {Object} 一括評価結果
   */
  evaluateBatchQuality(terms, context = '') {
    // TODO: 実装予定
    const results = terms.map(term => this.evaluateTermQuality(term, context));
    
    return {
      total_terms: terms.length,
      approved_terms: results.filter(r => r.decision === 'approved').length,
      pending_terms: results.filter(r => r.decision === 'pending').length,
      rejected_terms: results.filter(r => r.decision === 'rejected').length,
      average_quality_score: 0.0,
      evaluation_results: results
    };
  }
  
  /**
   * 信頼度スコアを評価する
   * @param {Object} term - 用語オブジェクト
   * @return {number} 信頼度スコア（0.0-1.0）
   * @private
   */
  _evaluateConfidenceScore(term) {
    // 類似度スコア、信頼度スコア、用語の出現頻度などを考慮
    // 現状は、termオブジェクトに存在する最も代表的なスコアを返す
    return term.similarity || term.confidence || 0.0;
  }
  
  /**
   * 文脈適合性をチェックする
   * @param {Object} term - 用語オブジェクト
   * @param {string} context - 文脈情報
   * @return {number} 適合性スコア（0.0-1.0）
   * @private
   */
  _checkContextRelevance(term, context) {
    // TODO: 実装予定
    // 用語が文脈に適合しているかを評価
    return 0.0;
  }
  
  /**
   * 品質判定を行う
   * @param {number} qualityScore - 品質スコア
   * @return {string} 判定結果（'approved', 'pending', 'rejected'）
   * @private
   */
  _makeQualityDecision(qualityScore) {
    if (qualityScore >= this.qualityThreshold) {
      return 'approved';
    } else if (qualityScore >= this.qualityThreshold * 0.5) {
      return 'pending';
    } else {
      return 'rejected';
    }
  }
  
 /**
  * 用語ペアを評価し、品質に応じて自動登録またはレビュー待ちに分類する
  * @param {Array<Object>} termPairs - 評価対象の用語ペア配列。各要素は {source, target, similarity, ...} を想定
  * @param {Dictionary} dictionary - 辞書オブジェクト
  * @param {string} dictName - 登録対象の辞書名（シート名）
  * @return {Object} 処理結果
  */
 evaluateAndRegister(termPairs, dictionary, dictName) {
   if (!termPairs || !Array.isArray(termPairs) || !dictionary || !dictName) {
     log('ERROR', 'evaluateAndRegister: 無効な引数です。辞書名が含まれているか確認してください。');
     return { success: false, message: '無効な引数です。辞書名が含まれているか確認してください。' };
   }

   const results = {
     success: true,
     processed: termPairs.length,
     approved: 0,
     pending: 0,
     rejected: 0,
     approvedTerms: [],
     pendingTerms: [],
     rejectedTerms: []
   };

   for (const termPair of termPairs) {
     try {
       // 1. 品質評価ロジック
       const qualityScore = this._evaluateConfidenceScore(termPair);
       // TODO: _checkContextRelevanceも将来的に組み込み、スコアを統合する
       termPair.qualityScore = qualityScore; // 後で参照できるようにスコアをペアに追加

       // 2. ステータス分類ロジック
       const decision = this._makeQualityDecision(qualityScore);
       termPair.decision = decision;

       switch (decision) {
         case 'approved':
           results.approved++;
           results.approvedTerms.push(termPair);
           break;
         case 'pending':
           results.pending++;
           results.pendingTerms.push(termPair);
           break;
         case 'rejected':
           results.rejected++;
           results.rejectedTerms.push(termPair);
           break;
       }
     } catch (e) {
       log('ERROR', `用語ペアの評価中にエラーが発生しました: ${termPair.source}`, e);
     }
   }

   // 3. 自動登録ロジック
   if (results.approvedTerms.length > 0) {
     const termsToAdd = results.approvedTerms.map(p => ({
       source: p.source,
       target: p.target,
       pos: p.pos || '名詞', // 品詞情報がない場合はデフォルト値
       notes: `自動承認 (スコア: ${p.qualityScore.toFixed(2)})`
     }));

     const addResult = dictionary.addTerms(dictName, termsToAdd);
     if (addResult && !addResult.success) {
       log('WARN', '辞書への自動登録に一部失敗しました。', addResult);
     } else if (addResult) {
       log('INFO', `${addResult.added}件の用語が辞書に自動登録されました。`);
     }
   }

   // 4. レビュー待ち処理
   if (results.pendingTerms.length > 0) {
     this._logPendingReviews(results.pendingTerms);
   }

   if (results.rejectedTerms.length > 0) {
     log('INFO', `${results.rejectedTerms.length}件の用語が自動的に却下されました。`);
   }

   return results;
 }

 /**
  * レビュー待ちの用語を記録する
  * @param {Array<Object>} pendingTerms - レビュー待ちの用語ペア配列
  * @private
  */
 _logPendingReviews(pendingTerms) {
   // TODO: 将来的に専用のスプレッドシートやデータベースに記録する機能を実装
   // 現状はログ出力のみ
   const termsToReview = pendingTerms.map(p =>
     `{ source: "${p.source}", target: "${p.target}", score: ${p.qualityScore.toFixed(2)} }`
   );
   log('INFO', `レビュー待ちの用語が ${pendingTerms.length} 件あります。`, termsToReview);

   // ここで専用シートへの書き込み処理を呼び出す
   // 例: this._writeToReviewSheet(pendingTerms);
 }
  
  /**
   * 辞書全体の品質統計を計算する
   * @param {string} dictionaryName - 辞書名
   * @return {Object} 品質統計
   */
  calculateDictionaryStatistics(dictionaryName = CONFIG.DEFAULT_DICT_NAME) {
    // TODO: 実装予定
    return {
      total_terms: 0,
      quality_score_distribution: {
        high: 0,    // 0.8以上
        medium: 0,  // 0.5-0.8
        low: 0      // 0.5未満
      },
      coverage_rate: 0.0,
      last_updated: new Date()
    };
  }
  
  /**
   * 品質改善の推奨事項を生成する
   * @param {Object} statistics - 品質統計
   * @return {Array} 推奨事項リスト
   */
  generateQualityRecommendations(statistics) {
    // TODO: 実装予定
    const recommendations = [];
    
    // 品質スコアが低い用語の改善提案
    // カバレッジ率向上の提案
    // 重複用語の統合提案など
    
    return recommendations;
  }
}
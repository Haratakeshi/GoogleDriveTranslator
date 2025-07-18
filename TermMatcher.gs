// TermMatcher.gs - 用語照合機能

/**
 * 用語照合クラス
 * 抽出された用語と既存辞書の照合を行う機能を提供します。
 * 処理は以下の順序で段階的に行われます：
 * 1. 完全一致: 抽出された用語が辞書の原語と完全に一致するかを評価します。
 * 2. 正規化一致: 表記揺れ（大文字・小文字、記号など）を吸収して一致するかを評価します。
 * 3. 部分一致: 抽出された用語が辞書の原語に含まれるかを評価します。
 * 4. あいまい一致: レーベンシュタイン距離に基づき、用語間の類似度を評価します。
 */
class TermMatcher {

  /**
   * コンストラクタ
   * @param {Object} [options={}] - 設定オプション
   * @param {number} [options.similarityThreshold=0.8] - あいまい一致と判断する類似度の閾値
   * @param {number} [options.partialMatchMinLength=3] - 部分一致と判断する最小文字数
   */
  constructor(options = {}) {
    this.config = CONFIG;
    this.dictionary = new Dictionary();
    this.similarityThreshold = options.similarityThreshold || 0.8;
    this.partialMatchMinLength = options.partialMatchMinLength || 3;

    log('INFO', 'TermMatcher initialized', {
      similarityThreshold: this.similarityThreshold,
      partialMatchMinLength: this.partialMatchMinLength
    });
  }

  /**
   * 抽出された用語と既存辞書を照合します。
   * 結果は「翻訳適用が確定したペア」と「新規登録候補の用語」に分類されます。
   * @param {Array<string>} extractedTerms - 抽出された用語の文字列配列
   * @param {string} [dictionaryName=CONFIG.DEFAULT_DICT_NAME] - 使用する辞書名
   * @return {{confirmedPairs: Array<Object>, newCandidates: Array<Object>}} 照合結果
   */
  match(extractedTerms, dictionaryName = CONFIG.DEFAULT_DICT_NAME) {
    log('INFO', `Term matching started for ${extractedTerms.length} terms in dictionary "${dictionaryName}"`);

    if (!extractedTerms || extractedTerms.length === 0) {
      log('INFO', 'No terms to match.');
      return { confirmedPairs: [], newCandidates: [] };
    }

    const uniqueExtractedTerms = [...new Set(extractedTerms)];
    const dictionaryTerms = this.dictionary._getOrCacheAllTerms(dictionaryName);

    if (!dictionaryTerms || dictionaryTerms.length === 0) {
      log('WARN', `Dictionary "${dictionaryName}" is empty or could not be loaded.`);
      const newCandidates = uniqueExtractedTerms.map(term => ({
        source: term,
        reason: 'new_term',
        details: '辞書が空のため、新規用語として扱います。'
      }));
      return { confirmedPairs: [], newCandidates };
    }

    let remainingTerms = [...uniqueExtractedTerms];
    
    // 1. 完全一致
    const exactMatchResult = this._exactMatch(remainingTerms, dictionaryTerms);
    const confirmedPairs = exactMatchResult.matched.map(m => ({
      source: m.extractedTerm,
      target: m.dictTerm.target,
      matchType: 'exact',
      details: '辞書と完全に一致しました。',
      dictionaryTerm: m.dictTerm
    }));
    remainingTerms = exactMatchResult.unmatched;
    log('INFO', `Exact match found: ${confirmedPairs.length}`);

    // 2. 正規化一致
    const normalizedMatchResult = this._normalizedMatch(remainingTerms, dictionaryTerms);
    normalizedMatchResult.matched.forEach(m => {
      confirmedPairs.push({
        source: m.extractedTerm,
        target: m.dictTerm.target,
        matchType: 'normalized',
        details: '表記揺れを正規化した結果、辞書と一致しました。',
        dictionaryTerm: m.dictTerm
      });
    });
    remainingTerms = normalizedMatchResult.unmatched;
    log('INFO', `Normalized match found: ${normalizedMatchResult.matched.length}`);

    // 3. 部分一致 & あいまい一致
    const partialFuzzyResult = this._partialAndFuzzyMatch(remainingTerms, dictionaryTerms);
    partialFuzzyResult.partial.forEach(m => {
      confirmedPairs.push({
        source: m.extractedTerm,
        target: m.dictTerm.target,
        matchType: 'partial',
        details: '辞書用語の一部として一致しました。',
        dictionaryTerm: m.dictTerm
      });
    });
    
    const newCandidates = partialFuzzyResult.fuzzy.map(m => ({
      source: m.extractedTerm,
      reason: 'fuzzy_match',
      details: `類似の用語が見つかりました (類似度: ${m.similarity.toFixed(2)})。`,
      similarTerm: m.dictTerm
    }));
    remainingTerms = partialFuzzyResult.unmatched;
    log('INFO', `Partial match found: ${partialFuzzyResult.partial.length}`);
    log('INFO', `Fuzzy match (new candidates) found: ${partialFuzzyResult.fuzzy.length}`);

    // 4. 最終的に残ったものを新規用語として追加
    remainingTerms.forEach(term => {
      newCandidates.push({
        source: term,
        reason: 'new_term',
        details: '辞書に一致する用語が見つかりませんでした。'
      });
    });
    log('INFO', `Completely new terms found: ${remainingTerms.length}`);

    log('INFO', 'Term matching finished.', {
      confirmedPairsCount: confirmedPairs.length,
      newCandidatesCount: newCandidates.length
    });

    return { confirmedPairs, newCandidates };
  }

  /**
   * 完全一致のマッチングを実行します。
   * @param {Array<string>} extractedTerms - 抽出された用語リスト
   * @param {Array<Object>} dictionaryTerms - 辞書の用語リスト
   * @return {{matched: Array<Object>, unmatched: Array<string>}} マッチング結果
   * @private
   */
  _exactMatch(extractedTerms, dictionaryTerms) {
    const matched = [];
    const unmatched = [];
    const dictMap = new Map(dictionaryTerms.map(t => [t.source, t]));

    for (const term of extractedTerms) {
      if (dictMap.has(term)) {
        matched.push({ extractedTerm: term, dictTerm: dictMap.get(term) });
      } else {
        unmatched.push(term);
      }
    }
    return { matched, unmatched };
  }

  /**
   * 正規化後のマッチングを実行します。
   * @param {Array<string>} extractedTerms - マッチしていない用語リスト
   * @param {Array<Object>} dictionaryTerms - 辞書の用語リスト
   * @return {{matched: Array<Object>, unmatched: Array<string>}} マッチング結果
   * @private
   */
  _normalizedMatch(extractedTerms, dictionaryTerms) {
    const matched = [];
    const unmatched = [];
    const normalizedDictMap = new Map();
    // 重複する正規化キーの場合は、最初に見つかったものを採用
    for (const dictTerm of dictionaryTerms) {
      const normalized = this._normalizeTerm(dictTerm.source);
      if (!normalizedDictMap.has(normalized)) {
        normalizedDictMap.set(normalized, dictTerm);
      }
    }

    for (const term of extractedTerms) {
      const normalizedTerm = this._normalizeTerm(term);
      if (normalizedDictMap.has(normalizedTerm)) {
        matched.push({ extractedTerm: term, dictTerm: normalizedDictMap.get(normalizedTerm) });
      } else {
        unmatched.push(term);
      }
    }
    return { matched, unmatched };
  }

  /**
   * 部分一致とあいまい一致のマッチングを同時に実行します。
   * 部分一致が見つかった用語は、あいまい一致の対象から外れます。
   * @param {Array<string>} extractedTerms - マッチしていない用語リスト
   * @param {Array<Object>} dictionaryTerms - 辞書の用語リスト
   * @return {{partial: Array, fuzzy: Array, unmatched: Array}} マッチング結果
   * @private
   */
  _partialAndFuzzyMatch(extractedTerms, dictionaryTerms) {
    const partial = [];
    const fuzzy = [];
    const stillUnmatched = [];
    
    for (const term of extractedTerms) {
      let foundPartial = false;
      // 部分一致チェック
      if (term.length >= this.partialMatchMinLength) {
        for (const dictTerm of dictionaryTerms) {
          // 抽出用語が辞書用語に含まれるケース
          if (dictTerm.source.includes(term)) {
            partial.push({ extractedTerm: term, dictTerm });
            foundPartial = true;
            break; 
          }
        }
      }

      if (foundPartial) {
        continue; // 部分一致が見つかったら、この用語の処理は終了
      }

      // あいまい一致チェック
      let bestMatch = { dictTerm: null, similarity: 0 };
      const normalizedTerm = this._normalizeTerm(term);

      for (const dictTerm of dictionaryTerms) {
        const normalizedDictTerm = this._normalizeTerm(dictTerm.source);
        const similarity = this._calculateSimilarity(normalizedTerm, normalizedDictTerm);
        if (similarity > bestMatch.similarity) {
          bestMatch = { dictTerm, similarity };
        }
      }

      if (bestMatch.similarity >= this.similarityThreshold) {
        fuzzy.push({ extractedTerm: term, ...bestMatch });
      } else {
        stillUnmatched.push(term);
      }
    }

    return { partial, fuzzy, unmatched: stillUnmatched };
  }

  /**
   * 用語を正規化します。
   * @param {string} term - 正規化対象の用語
   * @return {string} 正規化された用語
   * @private
   */
  _normalizeTerm(term) {
    return Utils.normalizeTerm(term);
  }

  /**
   * 2つの用語の類似度を計算します。
   * @param {string} term1 - 比較対象の用語1
   * @param {string} term2 - 比較対象の用語2
   * @return {number} 類似度スコア（0.0 - 1.0）
   * @private
   */
  _calculateSimilarity(term1, term2) {
    // このメソッド内では既に正規化されていることを前提とする
    return Utils.calculateSimilarityScore(term1, term2);
  }
}
// Utils.gs - ユーティリティ機能（用語抽出関連の拡張）

/**
 * ユーティリティクラス（用語抽出機能拡張）
 * 用語抽出関連のユーティリティ機能を提供
 */
class Utils {
  
  /**
   * テキストを正規化する
   * @param {string} text - 正規化対象テキスト
   * @return {string} 正規化されたテキスト
   */
  static normalizeText(text) {
    if (!text || typeof text !== 'string') {
      return '';
    }
    
    // TODO: 実装予定
    // 1. 全角・半角の統一
    // 2. 不要な空白の除去
    // 3. 改行文字の正規化
    // 4. 特殊文字の処理
    
    return text.trim()
               .replace(/\s+/g, ' ')  // 複数の空白を単一の空白に
               .replace(/[\r\n]+/g, '\n');  // 改行の正規化
  }
  
  /**
   * 用語を正規化する
   * @param {string} term - 正規化対象の用語
   * @return {string} 正規化された用語
   */
  static normalizeTerm(term) {
    if (!term || typeof term !== 'string') {
      return '';
    }
    
    return term
      // 1. 基本的なトリミング
      .trim()
      // 2. 全角英数字を半角に統一
      .replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
        return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
      })
      // 3. 長音記号の統一（ー、−、‐、―を統一）
      .replace(/[ー−‐―]/g, 'ー')
      // 4. 大文字小文字の統一
      .toLowerCase()
      // 5. 記号の除去（日本語文字、アルファベット、数字、基本的な記号は保持）
      .replace(/[^\w\s\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF\uFF10-\uFF19\uFF21-\uFF3A\uFF41-\uFF5Aー]/g, '')
      // 6. 空白の正規化
      .replace(/\s+/g, ' ')
      // 7. 最終トリミング
      .trim();
  }
  
  /**
   * 信頼度スコアを調整する
   * @param {number} baseScore - 基本信頼度スコア
   * @param {Object} factors - 調整要因
   * @return {number} 調整後の信頼度スコア（0.0-1.0）
   */
  static adjustConfidenceScore(baseScore, factors = {}) {
    if (typeof baseScore !== 'number' || baseScore < 0 || baseScore > 1) {
      return 0.0;
    }
    
    // TODO: 実装予定
    // 調整要因:
    // - contextRelevance: 文脈適合性
    // - termFrequency: 用語出現頻度
    // - definitionClarity: 定義の明確さ
    // - domainSpecificity: 専門分野特異性
    
    let adjustedScore = baseScore;
    
    // 文脈適合性による調整
    if (factors.contextRelevance !== undefined) {
      adjustedScore *= (0.7 + 0.3 * factors.contextRelevance);
    }
    
    // 用語出現頻度による調整
    if (factors.termFrequency !== undefined) {
      adjustedScore *= (0.8 + 0.2 * Math.min(factors.termFrequency / 10, 1));
    }
    
    return Math.min(Math.max(adjustedScore, 0.0), 1.0);
  }
  
  /**
   * レーベンシュタイン距離を計算する
   * @param {string} str1 - 比較対象文字列1
   * @param {string} str2 - 比較対象文字列2
   * @return {number} レーベンシュタイン距離
   */
  static calculateLevenshteinDistance(str1, str2) {
    if (!str1 || !str2) {
      return Math.max(str1?.length || 0, str2?.length || 0);
    }
    
    const matrix = [];
    const len1 = str1.length;
    const len2 = str2.length;
    
    // 初期化
    for (let i = 0; i <= len1; i++) {
      matrix[i] = [i];
    }
    for (let j = 0; j <= len2; j++) {
      matrix[0][j] = j;
    }
    
    // 動的プログラミング
    for (let i = 1; i <= len1; i++) {
      for (let j = 1; j <= len2; j++) {
        const cost = str1[i - 1] === str2[j - 1] ? 0 : 1;
        matrix[i][j] = Math.min(
          matrix[i - 1][j] + 1,      // 削除
          matrix[i][j - 1] + 1,      // 挿入
          matrix[i - 1][j - 1] + cost // 置換
        );
      }
    }
    
    return matrix[len1][len2];
  }
  
  /**
   * 類似度スコアを計算する（0.0-1.0）
   * @param {string} str1 - 比較対象文字列1
   * @param {string} str2 - 比較対象文字列2
   * @return {number} 類似度スコア
   */
  static calculateSimilarityScore(str1, str2) {
    if (!str1 || !str2) {
      return 0.0;
    }
    
    if (str1 === str2) {
      return 1.0;
    }
    
    const distance = Utils.calculateLevenshteinDistance(str1, str2);
    const maxLength = Math.max(str1.length, str2.length);
    
    return maxLength === 0 ? 1.0 : (maxLength - distance) / maxLength;
  }
  
  /**
   * Jaccard係数を計算する（0.0-1.0）
   * 2つの集合の類似度を測る指標。積集合のサイズを和集合のサイズで割ることで計算される。
   * @param {string} str1 - 比較対象文字列1
   * @param {string} str2 - 比較対象文字列2
   * @return {number} Jaccard係数
   */
  static calculateJaccardCoefficient(str1, str2) {
    if (typeof str1 !== 'string' || typeof str2 !== 'string') {
      return 0.0;
    }
    if (str1 === str2) {
      return 1.0;
    }
    if (str1.length === 0 && str2.length === 0) {
      return 1.0;
    }
    if (str1.length === 0 || str2.length === 0) {
      return 0.0;
    }

    const set1 = new Set(str1);
    const set2 = new Set(str2);

    const intersectionSize = [...set1].filter(char => set2.has(char)).length;
    const unionSize = set1.size + set2.size - intersectionSize;

    if (unionSize === 0) {
      return 1.0; // 両方の文字列が空の場合
    }

    return intersectionSize / unionSize;
  }
  
  /**
   * 配列から重複を除去する
   * @param {Array} array - 重複除去対象の配列
   * @param {Function} keyExtractor - キー抽出関数（オプション）
   * @return {Array} 重複除去された配列
   */
  static removeDuplicates(array, keyExtractor = null) {
    if (!Array.isArray(array)) {
      return [];
    }
    
    if (keyExtractor && typeof keyExtractor === 'function') {
      const seen = new Set();
      return array.filter(item => {
        const key = keyExtractor(item);
        if (seen.has(key)) {
          return false;
        }
        seen.add(key);
        return true;
      });
    }
    
    return [...new Set(array)];
  }
  
  /**
   * 安全なJSON解析
   * @param {string} jsonString - JSON文字列
   * @param {*} defaultValue - デフォルト値
   * @return {*} 解析結果またはデフォルト値
   */
  static safeJsonParse(jsonString, defaultValue = null) {
    try {
      return JSON.parse(jsonString);
    } catch (error) {
      log('WARN', 'JSON解析に失敗しました', { jsonString, error: error.message });
      return defaultValue;
    }
  }
}
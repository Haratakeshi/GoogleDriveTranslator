// TermExtractor.gs - 用語抽出機能

/**
 * 用語抽出クラス
 * テキストから固有名詞や専門用語を抽出する機能を提供
 */
class TermExtractor {
  
  /**
   * コンストラクタ
   */
  constructor() {
    this.config = CONFIG;
    this.cache = CacheService.getUserCache();
  }
  
  /**
   * テキストから用語を抽出する
   * @param {string} text - 抽出対象テキスト
   * @param {string} sourceLang - 原文言語
   * @param {string} targetLang - 翻訳先言語
   * @param {string} domain - 専門分野
   * @param {Object} options - オプション設定
   * @return {Object} 抽出された用語リストと統計情報
   */
  extract(text, sourceLang = 'auto', targetLang = 'ja', domain = '一般', options = {}) {
    try {
      // 1. 入力検証
      if (!text || typeof text !== 'string' || text.trim().length === 0) {
        return {
          extracted_terms: [],
          extraction_metadata: {
            total_terms_found: 0,
            domain_confidence: 0.0,
            error: 'テキストが空または無効です'
          }
        };
      }

      // 2. 専門分野の検証
      if (!this.config.SUPPORTED_DOMAINS[domain]) {
        throw new Error(this.config.ERROR_MESSAGES.INVALID_DOMAIN);
      }

      // 3. テキスト正規化
      const normalizedText = this._normalizeText(text);
      
      // 4. テキストサイズチェック
      if (normalizedText.length > this.config.TERM_BATCH_SIZE) {
        log('WARN', `テキストサイズが制限を超えています: ${normalizedText.length} > ${this.config.TERM_BATCH_SIZE}`);
        // 必要に応じてテキストを分割する処理を追加
      }

      // 5. OpenAI API呼び出し
      const apiResponse = this._callOpenAIForExtraction(normalizedText, sourceLang, targetLang, domain);
      
      if (!apiResponse || !apiResponse.extracted_terms) {
        throw new Error('API応答が無効です');
      }

      // 6. 品質フィルタリング
      const filteredTerms = this._filterQuality(apiResponse.extracted_terms, normalizedText);

      // 7. 重複除去
      const uniqueTerms = this._removeDuplicates(filteredTerms);

      // 8. 統計生成
      const statistics = this._generateStatistics(uniqueTerms);

      const extractionResult = {
        extracted_terms: uniqueTerms,
        extraction_metadata: {
          ...statistics,
          domain: domain,
          source_language: sourceLang,
          target_language: targetLang,
          original_text_length: text.length,
          normalized_text_length: normalizedText.length,
          processing_timestamp: new Date().toISOString()
        }
      };

      // デバッグ用: 最終的な用語抽出結果をJSON出力
      log('INFO', '最終用語抽出結果:', JSON.stringify(extractionResult, null, 2));

      return extractionResult;

    } catch (error) {
      log('ERROR', '用語抽出に失敗しました', error);
      return {
        extracted_terms: [],
        extraction_metadata: {
          total_terms_found: 0,
          domain_confidence: 0.0,
          error: error.message || this.config.ERROR_MESSAGES.TERM_EXTRACTION_FAILED
        }
      };
    }
  }
  
  /**
   * テキストを正規化する
   * @param {string} text - 正規化対象テキスト
   * @return {string} 正規化されたテキスト
   * @private
   */
  _normalizeText(text) {
    if (!text || typeof text !== 'string') {
      return '';
    }
    
    // Utilsクラスの正規化機能を使用
    return Utils.normalizeText(text);
  }
  
  /**
   * OpenAI APIを呼び出して用語を抽出する
   * @param {string} normalizedText - 正規化されたテキスト
   * @param {string} sourceLang - 原文言語
   * @param {string} targetLang - 翻訳先言語
   * @param {string} domain - 専門分野
   * @return {Object} API応答
   * @private
   */
  _callOpenAIForExtraction(normalizedText, sourceLang, targetLang, domain) {
    const systemPrompt = `You are a specialized term extraction system that analyzes text and extracts important SOURCE LANGUAGE terms only.

Extract terms from the provided text according to these guidelines:

# Term Types to Extract:
1. **Proper Nouns**: Names of people, places, organizations, brands
2. **Technical Terms**: Domain-specific technical concepts
3. **Compound Terms**: Multi-word concepts that function as a unit
4. **Acronyms**: Abbreviations and shortened forms
5. **Product Names**: Specific product or service names
6. **General Noun Phrases**: Contextually important noun phrases
7. **Content Words**: Meaningful nouns, verbs, and adjectives (2+ characters)

# Domain Context: ${this.config.SUPPORTED_DOMAINS[domain]}
# Source Language: ${sourceLang === 'auto' ? 'Auto-detect' : this.config.SUPPORTED_LANGUAGES[sourceLang] || sourceLang}

# IMPORTANT INSTRUCTIONS:
1. Extract SOURCE LANGUAGE terms only - DO NOT provide translations
2. The normalized_term field should be the normalized form of the SOURCE term (not translation)
3. For each term, provide confidence score (0.0-1.0) based on importance and clarity
4. Include context where the term appears in the source text
5. Focus on terms that would be important for creating translation dictionaries
6. **EXTRACT AGGRESSIVELY**: Even from short text like "【依頼内容】", extract meaningful parts like "依頼" and "内容"
7. **IGNORE SYMBOLS**: Focus on the meaningful content words, not symbols like 【】・※
8. **MINIMUM 2 CHARACTERS**: Extract terms that are at least 2 characters long

# Examples:
- From "【依頼内容】" → extract "依頼" (request) and "内容" (content)
- From "・プラットフォーム用" → extract "プラットフォーム" (platform)
- From "※注意事項" → extract "注意" (attention) and "事項" (matter)

Return results in the specified JSON structure.`;

    const userPrompt = `Please extract important terms from the following text:

${normalizedText}`;

    // JSONスキーマを定義（OpenAI Structured Outputs用）
    const responseSchema = {
      type: "object",
      properties: {
        extracted_terms: {
          type: "array",
          items: {
            type: "object",
            properties: {
              term: { type: "string", description: "The extracted term in source language" },
              normalized_term: { type: "string", description: "Normalized version of the source term (NOT translation)" },
              term_type: {
                type: "string",
                enum: ["Proper Noun", "Technical Term", "Compound Term", "Acronym", "Product Name", "General Noun Phrase", "Content Word"],
                description: "Type of the extracted term"
              },
              context: { type: "string", description: "Context where the term appears in source text" },
              confidence: { type: "number", minimum: 0.0, maximum: 1.0, description: "Confidence score for term importance" },
              source_language: { type: "string", description: "Detected source language" }
            },
            required: ["term", "normalized_term", "term_type", "context", "confidence", "source_language"]
          }
        },
        extraction_metadata: {
          type: "object",
          properties: {
            total_terms_found: { type: "integer", description: "Total number of terms extracted" },
            domain_confidence: { type: "number", minimum: 0.0, maximum: 1.0, description: "Confidence that text matches specified domain" }
          },
          required: ["total_terms_found", "domain_confidence"]
        }
      },
      required: ["extracted_terms", "extraction_metadata"]
    };

    const payload = {
      model: this.config.TERM_EXTRACTION_MODEL,
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userPrompt }
      ],
      temperature: this.config.TERM_EXTRACTION_TEMPERATURE,
      max_tokens: this.config.TERM_EXTRACTION_MAX_TOKENS,
      response_format: {
        type: "json_schema",
        json_schema: {
          name: "term_extraction_result",
          schema: responseSchema
        }
      }
    };

    const options = {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${this.config.OPENAI_API_KEY}`,
        'Content-Type': 'application/json',
        'OpenAI-Organization': this.config.OPENAI_ORGANIZATION_ID,
        'OpenAI-Project': this.config.OPENAI_PROJECT_ID
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(this.config.OPENAI_API_URL, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();

      if (responseCode !== 200) {
        log('ERROR', `Term Extraction API Error: ${responseCode}`, responseText);
        throw new Error(`API Error: ${responseCode}`);
      }

      const result = JSON.parse(responseText);
      if (result.error) {
        throw new Error(result.error.message);
      }

      const content = JSON.parse(result.choices[0].message.content);
      
      // デバッグ用: 用語抽出結果の詳細JSONログ出力
      log('INFO', '用語抽出API生レスポンス:', JSON.stringify(result, null, 2));
      log('INFO', '用語抽出結果詳細:', JSON.stringify(content, null, 2));
      
      // 応答の検証
      if (!content.extracted_terms || !Array.isArray(content.extracted_terms)) {
        throw new Error('API応答の形式が不正です');
      }

      return content;

    } catch (error) {
      log('ERROR', 'OpenAI Term Extraction API call failed', error);
      throw error;
    }
  }
  
  /**
   * 抽出された用語の品質をフィルタリングする
   * @param {Array} terms - 抽出された用語リスト
   * @param {string} originalText - 元のテキスト（信頼度調整用）
   * @return {Array} フィルタリングされた用語リスト
   * @private
   */
  _filterQuality(terms, originalText = '') {
    if (!Array.isArray(terms)) {
      return [];
    }

    return terms.filter(term => {
      // 基本的な検証
      if (!term || typeof term !== 'object') {
        return false;
      }

      // 必須フィールドの検証
      if (!term.term || !term.normalized_term || !term.term_type) {
        return false;
      }

      // 用語の長さ検証（あまりに短い、または長すぎる用語を除外）
      const termLength = term.term.trim().length;
      if (termLength < 2 || termLength > 100) {
        return false;
      }

      // 信頼度スコアの検証
      if (typeof term.confidence !== 'number' || term.confidence < 0 || term.confidence > 1) {
        return false;
      }

      // 最低信頼度閾値のチェック（短いテキストの場合は閾値を下げる）
      const minConfidence = originalText.length <= 10 ? 0.1 : this.config.QUALITY_GATE_THRESHOLD * 0.3;
      if (term.confidence < minConfidence) {
        return false;
      }

      // 一般的すぎる用語の除外（例：a, the, is, など）
      const commonWords = ['a', 'an', 'the', 'is', 'are', 'was', 'were', 'be', 'been', 'being', 'have', 'has', 'had', 'do', 'does', 'did', 'will', 'would', 'could', 'should', 'may', 'might', 'can', 'must'];
      if (commonWords.includes(term.normalized_term.toLowerCase())) {
        return false;
      }

      return true;
    }).map(term => {
      return {
        ...term,
        quality_filtered: true
      };
    });
  }
  
  /**
   * 重複する用語を除去する
   * @param {Array} terms - 用語リスト
   * @return {Array} 重複除去された用語リスト
   * @private
   */
  _removeDuplicates(terms) {
    if (!Array.isArray(terms)) {
      return [];
    }

    // Utilsクラスの重複除去機能を使用
    // 正規化された用語をキーとして重複を除去
    return Utils.removeDuplicates(terms, term => term.normalized_term.toLowerCase());
  }
  
  /**
   * 抽出統計を生成する
   * @param {Array} terms - 用語リスト
   * @return {Object} 統計情報
   * @private
   */
  _generateStatistics(terms) {
    if (!Array.isArray(terms) || terms.length === 0) {
      return {
        total_terms_found: 0,
        domain_confidence: 0.0,
        average_confidence: 0.0,
        term_type_distribution: {},
        high_confidence_terms: 0,
        medium_confidence_terms: 0,
        low_confidence_terms: 0
      };
    }

    // 用語タイプ別の分布を計算
    const termTypeDistribution = {};
    let totalConfidence = 0;
    let highConfidenceCount = 0;
    let mediumConfidenceCount = 0;
    let lowConfidenceCount = 0;

    terms.forEach(term => {
      // 用語タイプ別カウント
      if (term.term_type) {
        termTypeDistribution[term.term_type] = (termTypeDistribution[term.term_type] || 0) + 1;
      }

      // 信頼度の集計
      if (typeof term.confidence === 'number') {
        totalConfidence += term.confidence;
        
        if (term.confidence >= 0.8) {
          highConfidenceCount++;
        } else if (term.confidence >= 0.5) {
          mediumConfidenceCount++;
        } else {
          lowConfidenceCount++;
        }
      }
    });

    // 平均信頼度を計算
    const averageConfidence = terms.length > 0 ? totalConfidence / terms.length : 0.0;

    // ドメイン信頼度を計算（高信頼度用語の割合に基づく）
    const domainConfidence = terms.length > 0 ? highConfidenceCount / terms.length : 0.0;

    return {
      total_terms_found: terms.length,
      domain_confidence: Math.min(Math.max(domainConfidence, 0.0), 1.0),
      average_confidence: Math.min(Math.max(averageConfidence, 0.0), 1.0),
      term_type_distribution: termTypeDistribution,
      high_confidence_terms: highConfidenceCount,
      medium_confidence_terms: mediumConfidenceCount,
      low_confidence_terms: lowConfidenceCount
    };
  }
}
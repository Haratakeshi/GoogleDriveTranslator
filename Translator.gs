// Translator.gs - 翻訳処理クラス

/**
 * OpenAI APIを使用した翻訳処理を管理するクラス
 */
class Translator {
  constructor(dictName) {
    this.apiKey = CONFIG.OPENAI_API_KEY;
    this.orgId = CONFIG.OPENAI_ORGANIZATION_ID;
    this.projectId = CONFIG.OPENAI_PROJECT_ID;
    this.model = CONFIG.OPENAI_MODEL;
    this.apiUrl = CONFIG.OPENAI_API_URL;
    this.dictionary = new Dictionary();
    this.termExtractor = new TermExtractor();
    this.termMatcher = new TermMatcher();
    this.qualityManager = new DictionaryQualityManager();
    this.dictName = dictName || '';
  }

  /**
   * 単一のテキストを翻訳する
   * @param {string} text - 翻訳するテキスト
   * @param {string} targetLang - 翻訳先言語
   * @param {Array<Object>|null} preConfirmedPairs - (オプション) 事前に確定した用語ペア
   * @return {string} 翻訳されたテキスト
   */
  translateText(text, targetLang, preConfirmedPairs = null) {
    if (!text || !text.trim()) {
      return '';
    }

    let confirmedPairs;

    if (preConfirmedPairs) {
      // 事前確定辞書が提供された場合、用語抽出・照合をスキップ
      confirmedPairs = preConfirmedPairs;
      log('INFO', '事前確定辞書を使用して翻訳します。', { count: confirmedPairs.length });
    } else {
      // 既存の処理：用語抽出と照合
      // 1. 用語抽出
      const extractionResult = this.termExtractor.extract(text, 'auto', targetLang);
      const extractedTerms = extractionResult.extracted_terms.map(t => t.term);
      log('INFO', '用語抽出完了', { count: extractedTerms.length });

      // 2. 用語照合
      const matchResult = this.dictName ? this.termMatcher.match(extractedTerms, this.dictName) : { confirmedPairs: [], newCandidates: [] };
      confirmedPairs = matchResult.confirmedPairs;
      log('INFO', '用語照合完了', { confirmed: confirmedPairs.length, candidates: matchResult.newCandidates.length });
    }

    // 3. 翻訳実行
    const result = this.callOpenAI([text], targetLang, confirmedPairs);

    // 4. 新規用語の自動登録
    if (result.usedTermPairs && result.usedTermPairs.length > 0 && this.dictName) {
      this.registerNewTerms(result.usedTermPairs);
    }

    // 翻訳結果と新しい用語候補を返す
    return {
      translatedText: result.translations[0] || text, // 失敗時は原文を返す
      newTermCandidates: result.usedTermPairs || []
    };
  }

  /**
   * OpenAI APIを呼び出して翻訳
   * @param {Array<string>} texts - 翻訳するテキストの配列
   * @param {string} targetLang - 翻訳先言語
   * @param {Array<Object>} confirmedPairs - 適用が確定した辞書データ
   * @return {Object} 翻訳結果 {translations: ["..."], usedTermPairs: [...]}
   */
  callOpenAI(texts, targetLang, confirmedPairs) {
    const systemPrompt = `You are a high-precision, AI-powered translation engine that STRICTLY enforces dictionary mappings and generates new term pairs for dictionary learning.

# Primary Goal
Translate texts from a source language to a target language with ABSOLUTE adherence to the provided dictionary and generate new term pairs for future translations.

# Input JSON Structure
{
  "dictionary": [
    { "source": "source_term_1", "target": "target_term_1" },
    ...
  ],
  "textsToTranslate": [
    { "id": 0, "text": "The text to be translated." },
    ...
  ]
}

# Output JSON Structure
{
  "translations": [
    {
      "id": 0,
      "sourceText": "The original text.",
      "translatedText": "The translated text."
    },
    ...
  ],
  "usedTermPairs": [
    { "source": "original_term", "target": "translated_term" },
    ...
  ]
}

# MANDATORY DICTIONARY ENFORCEMENT RULES
1. **ABSOLUTE DICTIONARY PRIORITY**: For ANY dictionary source term found in the text, you MUST use the corresponding target term.
2. **CONSISTENCY GUARANTEE**: The same source term MUST always translate to the same target term throughout ALL texts.
3. **COMPREHENSIVE MATCHING**: Apply dictionary mappings for exact matches and case variations.

# Core Instructions
1. **Translate**: Translate every object in the \`textsToTranslate\` array into ${CONFIG.SUPPORTED_LANGUAGES[targetLang]}.
2. **Dictionary Enforcement**: Apply all provided dictionary mappings where source terms appear in the text.
3. **Output \`translations\` Array**:
   - The \`translations\` array MUST have the exact same number of elements as the input \`textsToTranslate\` array.
   - Each object must contain \`id\`, \`sourceText\` (copy of original), and \`translatedText\`.
4. **CRITICAL: Output \`usedTermPairs\` Array**:
   - **ALWAYS include new term pairs** that you create during translation
   - Include important nouns, technical terms, proper nouns, and key phrases you translate
   - Each object MUST have \`source\` (original term) and \`target\` (your translation) keys
   - Focus on terms that would be useful for future translations
   - Example: If you translate "プラットフォーム" to "平台", include {"source": "プラットフォーム", "target": "平台"}
   - **DO NOT leave usedTermPairs empty** - always include meaningful term pairs you created
5. **Quality Assurance**: 
   - Preserve all formatting, symbols, numbers, and special characters exactly as they appear in the source.
6. **Strict JSON Format**: Output must be a single, valid JSON object matching the specified structure exactly.`;

    const textsToTranslate = texts.map((text, index) => ({ id: index, text: text }));
    const inputJson = {
      dictionary: confirmedPairs.map(term => ({ source: term.source, target: term.target })),
      textsToTranslate: textsToTranslate
    };

    const userPrompt = `Please process the following JSON data according to the instructions in the system prompt.\n\n${JSON.stringify(inputJson, null, 2)}`;

    const payload = {
      model: this.model,
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userPrompt }
      ],
      temperature: 0.1,
      max_tokens: 16384,
      response_format: { type: "json_object" }
    };

    const options = {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${this.apiKey}`,
        'Content-Type': 'application/json',
        'OpenAI-Organization': this.orgId,
        'OpenAI-Project': this.projectId
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(this.apiUrl, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();

      if (responseCode !== 200) {
        log('ERROR', `API Error: ${responseCode}`, responseText);
        throw new Error(`API Error: ${responseCode}`);
      }

      const result = JSON.parse(responseText);
      if (result.error) {
        throw new Error(result.error.message);
      }

      const content = JSON.parse(result.choices[0].message.content);

      // デバッグ用: 翻訳API生レスポンスと翻訳結果の詳細JSONログ出力
      log('INFO', '翻訳API生レスポンス:', JSON.stringify(result, null, 2));
      log('INFO', '翻訳結果Schema詳細:', JSON.stringify(content, null, 2));

      if (!content.translations || !Array.isArray(content.translations) || content.translations.length !== texts.length) {
        log('ERROR', 'API response validation failed', content);
        throw new Error('翻訳結果の形式または件数が不正です。');
      }

      const orderedTranslations = new Array(texts.length);
      for (const item of content.translations) {
        if (item.id !== undefined && item.id < texts.length) {
          orderedTranslations[item.id] = item.translatedText;
        }
      }
      
      for (let i = 0; i < texts.length; i++) {
        if (orderedTranslations[i] === undefined) {
          orderedTranslations[i] = texts[i];
        }
      }

      const finalResult = {
        translations: orderedTranslations,
        usedTermPairs: content.usedTermPairs || []
      };

      // デバッグ用: 最終翻訳結果をJSON出力
      log('INFO', '最終翻訳結果詳細:', JSON.stringify(finalResult, null, 2));

      return finalResult;

    } catch (error) {
      log('ERROR', 'OpenAI API call failed', error);
      return {
        translations: texts,
        usedTermPairs: []
      };
    }
  }

  /**
   * 翻訳で使用された新規用語を辞書に自動登録する
   * @param {Array<Object>} usedTermPairs - 翻訳で使用された用語ペア
   */
  registerNewTerms(usedTermPairs) {
    if (!usedTermPairs || !Array.isArray(usedTermPairs) || usedTermPairs.length === 0) {
      return;
    }

    try {
      // 既存の辞書データを取得して重複チェック
      const existingTerms = this.dictionary._getOrCacheAllTerms(this.dictName);
      const existingKeys = new Set(existingTerms.map(term => `${term.source}_${term.target}`));

      // 重複していない新規用語のみをフィルタリング
      const newTerms = usedTermPairs.filter(pair => {
        const key = `${pair.source}_${pair.target}`;
        const isDuplicate = existingKeys.has(key);
        if (isDuplicate) {
          log('INFO', `重複により新規登録をスキップ: "${pair.source}" -> "${pair.target}"`);
        }
        return !isDuplicate;
      });

      if (newTerms.length === 0) {
        log('INFO', '新規登録対象の用語がありませんでした');
        return;
      }

      // 辞書に新規用語を追加
      const addResult = this.dictionary.addTerms(this.dictName, newTerms);
      
      if (addResult.success) {
        log('INFO', `${addResult.added}個の新規用語を辞書「${this.dictName}」に自動登録しました`);
        
        // 辞書キャッシュをクリアして次回翻訳で最新データを使用
        this.dictionary.cache.delete(`${this.dictName}_terms_all_data`);
        
        // デバッグ用: 登録された用語をJSON出力
        log('INFO', '登録された新規用語詳細:', JSON.stringify(newTerms, null, 2));
      } else {
        log('ERROR', '新規用語の自動登録に失敗しました', addResult.message);
      }

    } catch (error) {
      log('ERROR', '新規用語自動登録処理でエラーが発生しました', error);
    }
  }
}
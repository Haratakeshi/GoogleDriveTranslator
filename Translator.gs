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
    const systemPrompt = `You are a high-precision, AI-powered translation engine that strictly follows JSON-based input and output formats.

# Primary Goal
Translate texts from a source language to a target language, meticulously applying a provided dictionary and identifying new term pairs.

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

# Core Instructions
1.  **Translate**: Translate every object in the \`textsToTranslate\` array into ${CONFIG.SUPPORTED_LANGUAGES[targetLang]}.
2.  **Mandatory Dictionary Application**: You MUST use the provided \`dictionary\`. If a \`source\` term from the dictionary appears in the text, its corresponding \`target\` term MUST be used in the translation.
3.  **Output \`translations\` Array**:
    *   The \`translations\` array in the output MUST have the exact same number of elements as the input \`textsToTranslate\` array.
    *   Each object in the array must contain \`id\`, \`sourceText\` (copy of original), and \`translatedText\`.
4.  **Output \`usedTermPairs\` Array**:
    *   This array must contain all term pairs that were actually used or identified during translation.
    *   This includes:
        *   Pairs applied directly from the provided \`dictionary\`.
        *   New technical terms, proper nouns, or domain-specific phrases you identified and translated.
    *   Each object MUST have \`source\` and \`target\` keys.
    *   Do NOT include generic words. Focus on terms that are important for maintaining consistency.
5.  **Strict JSON Format**: The final output must be a single, valid JSON object matching the specified structure. Do not add any text outside the JSON structure.`;

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

      return {
        translations: orderedTranslations,
        usedTermPairs: content.usedTermPairs || []
      };

    } catch (error) {
      log('ERROR', 'OpenAI API call failed', error);
      return {
        translations: texts,
        usedTermPairs: []
      };
    }
  }
}
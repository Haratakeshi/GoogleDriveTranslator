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
    this.dictName = dictName || '';
  }

  /**
   * 単一のテキストを翻訳する
   * @param {string} text - 翻訳するテキスト
   * @param {string} targetLang - 翻訳先言語
   * @return {string} 翻訳されたテキスト
   */
  translateText(text, targetLang) {
    if (!text || !text.trim()) {
      return '';
    }
    // 辞書データを取得
    const dictTerms = this.dictName ? this.dictionary.getBatchTerms(this.dictName, [text]) : [];

    // 翻訳実行（callOpenAIはテキストの配列を期待するため、配列で渡す）
    const result = this.callOpenAI([text], targetLang, dictTerms);

    // 新しい用語を辞書に登録
    if (this.dictName && result.newTerms && result.newTerms.length > 0) {
      try {
        this.dictionary.addTerms(this.dictName, result.newTerms);
      } catch (e) {
        log('WARN', '辞書への用語追加に失敗しました', e);
      }
    }

    return result.translations[0] || text; // 失敗時は原文を返す
  }

  /**
   * OpenAI APIを呼び出して翻訳
   * @param {Array} texts - 翻訳するテキストの配列（現在は要素1つのみ）
   * @param {string} targetLang - 翻訳先言語
   * @param {Array} dictTerms - 辞書用語
   * @return {Object} 翻訳結果 {translations: ["..."], newTerms: [...]}
   */
  callOpenAI(texts, targetLang, dictTerms) {
    const systemPrompt = `You will function as a high-precision translation API that receives JSON and returns JSON.
Translate each text object within the 'textsToTranslate' array from the user-provided JSON object into the specified language: ${CONFIG.SUPPORTED_LANGUAGES[targetLang]}.

# Instructions
1.  Translate every element in the input JSON's 'textsToTranslate' array.
2.  The root of the output JSON must have a key named 'translations'. Its value must be an array of translation result objects.
3.  [IMPORTANT] The output key name is 'translations'. Do not confuse it with the input key 'textsToTranslate'.
4.  The number of elements in the 'translations' array must be EXACTLY the same as the number of elements in the input 'textsToTranslate' array.
5.  Each translation result object MUST include the following three keys: the original 'id', the original 'text' copied into 'sourceText', and the translation result in 'translatedText'.
6.  If you find any new pairs of technical terms during the translation process, add them to the 'newTerms' array.`;

    const textsToTranslate = texts.map((text, index) => ({ id: index, text: text }));
    const inputJson = {
      dictionary: dictTerms.map(term => ({ source: term.source, target: term.target })),
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
      max_tokens: 16384, // トークン上限を引き上げ
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
        newTerms: content.newTerms || []
      };

    } catch (error) {
      log('ERROR', 'OpenAI API call failed', error);
      return {
        translations: texts,
        newTerms: []
      };
    }
  }
}
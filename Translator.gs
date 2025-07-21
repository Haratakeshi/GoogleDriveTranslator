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
    const systemPrompt = `You are an expert translator who processes structured XML input. Your mission is to:
1. Accurately translate the text within the <text> tags into the specified target language.
2. Strictly follow the translation rules provided in the <dictionary> tags.
3. Extract all meaningful noun and proper noun pairs from your translation and list them in the output.

# Rules:
- You MUST use the target term from the dictionary for any matching source term.
- If a dictionary entry has identical source and target, that term MUST remain untranslated.

# Output Format:
Return a SINGLE, VALID JSON object with two keys:
1. "translations": An array of translated strings, in the same order as the input <text> tags.
2. "term_pairs": An array of all extracted {source, target} pairs. If no relevant terms are found, return an empty array [].`;

    // XML形式で辞書を作成
    const dictionaryXml = confirmedPairs.length > 0
      ? `<dictionary>\n${confirmedPairs.map(p => `  <term source="${Utils.escapeXml(p.source)}">${Utils.escapeXml(p.target)}</term>`).join('\n')}\n</dictionary>`
      : '<dictionary></dictionary>';

    // XML形式で翻訳対象テキストを作成
    const textsXml = `<texts>\n${texts.map((text, index) => `  <text id="${index}">${Utils.escapeXml(text)}</text>`).join('\n')}\n</texts>`;

    const userPrompt = `Please translate the content within the <text> tags from the source language to ${CONFIG.SUPPORTED_LANGUAGES[targetLang]}.

${dictionaryXml}

${textsXml}`;

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
        log('ERROR', 'API response validation failed for translations', content);
        throw new Error('翻訳結果の形式または件数が不正です。');
      }

      const finalResult = {
        translations: content.translations,
        usedTermPairs: content.term_pairs || [] // 全ての用語ペアを取得
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
      // 1. 無意味な用語ペアをフィルタリング
      const validTermPairs = usedTermPairs.filter(pair => {
        // 原文と訳文が同じペア（翻訳されていない）を除外
        if (pair.source === pair.target) {
          log('INFO', `翻訳されていないペアをスキップ: "${pair.source}" -> "${pair.target}"`);
          return false;
        }
        
        // 空文字や無効なペアを除外
        if (!pair.source || !pair.target || 
            typeof pair.source !== 'string' || typeof pair.target !== 'string' ||
            pair.source.trim().length === 0 || pair.target.trim().length === 0) {
          log('INFO', `無効なペアをスキップ: "${pair.source}" -> "${pair.target}"`);
          return false;
        }
        
        // 一般的すぎる単語や記号のみのペアを除外
        const commonSymbols = ['【', '】', '・', '※', '■', '□', '●', '○'];
        if (commonSymbols.some(symbol => pair.source.includes(symbol) && pair.source.length <= 3)) {
          log('INFO', `記号のみのペアをスキップ: "${pair.source}" -> "${pair.target}"`);
          return false;
        }
        
        return true;
      });

      if (validTermPairs.length === 0) {
        log('INFO', '有効な新規用語ペアがありませんでした');
        return;
      }

      // 2. 既存の辞書データを取得して重複チェック
      const existingTerms = this.dictionary._getOrCacheAllTerms(this.dictName);
      const existingKeys = new Set(existingTerms.map(term => `${term.source}_${term.target}`));

      // 3. 重複していない新規用語のみをフィルタリング
      const newTerms = validTermPairs.filter(pair => {
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

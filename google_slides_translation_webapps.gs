/**
 * Web対応版 Google Slides翻訳スクリプト（完全版）
 * Webアプリケーションとして公開して利用できます
 */

// ===============================================
// Webアプリケーション用エンドポイント
// ===============================================

/**
 * Webアプリケーションのエントリーポイント（GET）
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Google Slides 翻訳・校正ツール')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * プレゼンテーション情報を取得
 */
function getPresentationInfo(presentationId) {
  try {
    var presentation = SlidesApp.openById(presentationId);
    var slides = presentation.getSlides();
    
    return {
      success: true,
      title: presentation.getName(),
      slideCount: slides.length,
      message: 'プレゼンテーション情報を取得しました'
    };
  } catch (error) {
    return {
      success: false,
      message: 'エラー: ' + error.toString()
    };
  }
}

/**
 * 翻訳を実行
 */
function executeTranslation(presentationId, sourceLang, targetLang, scope, enableProofread) {
  try {
    var presentation = SlidesApp.openById(presentationId);
    var slides = presentation.getSlides();
    
    // autoの場合はenに変換
    if (sourceLang === 'auto') {
      sourceLang = 'en';
    }
    
    if (scope === 'all') {
      // 全スライドを翻訳
      for (var i = 0; i < slides.length; i++) {
        translateSlide(slides[i], sourceLang, targetLang, enableProofread);
      }
      
      return {
        success: true,
        message: '全' + slides.length + 'スライドの翻訳が完了しました'
      };
    } else {
      // 特定スライドを翻訳（将来の拡張用）
      var slideIndex = parseInt(scope);
      if (slideIndex >= 0 && slideIndex < slides.length) {
        translateSlide(slides[slideIndex], sourceLang, targetLang, enableProofread);
        return {
          success: true,
          message: 'スライド ' + (slideIndex + 1) + ' の翻訳が完了しました'
        };
      } else {
        return {
          success: false,
          message: 'スライド番号が範囲外です'
        };
      }
    }
  } catch (error) {
    return {
      success: false,
      message: 'エラー: ' + error.toString()
    };
  }
}

/**
 * 校正を実行
 */
function executeProofread(presentationId, proofreadType) {
  try {
    var presentation = SlidesApp.openById(presentationId);
    var slides = presentation.getSlides();
    
    if (proofreadType === 'databricks') {
      // Databricks校正
      for (var i = 0; i < slides.length; i++) {
        databricksProofreadSlideWithFormatPreservation(slides[i]);
      }
      return {
        success: true,
        message: '全' + slides.length + 'スライドの高度な校正が完了しました'
      };
    } else {
      // 基本校正
      for (var i = 0; i < slides.length; i++) {
        proofreadSlideWithFormatPreservation(slides[i]);
      }
      return {
        success: true,
        message: '全' + slides.length + 'スライドの基本校正が完了しました'
      };
    }
  } catch (error) {
    return {
      success: false,
      message: 'エラー: ' + error.toString()
    };
  }
}

// ===============================================
// 既存の翻訳機能（元のコード）
// ===============================================

/**
 * 新しいリスト項目判定関数（完全書き直し）
 * @param {string} line 1行のテキスト
 * @returns {boolean} リスト項目かどうか
 */
function isListItemNew(line) {
  if (!line || typeof line !== 'string' || line.trim() === '') {
    return false;
  }
  
  // シンプルなリストマーカーのパターン（制約なし）
  const patterns = [
    /^\s*•\s+/,         // • マーカー
    /^\s*\*\s+/,        // * マーカー  
    /^\s*-\s+/,         // - マーカー
    /^\s*\d+\.\s+/,     // 1. マーカー
    /^\s*\d+\)\s+/,     // 1) マーカー
    /^\s*[a-zA-Z]\.\s+/, // a. マーカー
    /^\s*[a-zA-Z]\)\s+/  // a) マーカー
  ];
  
  for (let i = 0; i < patterns.length; i++) {
    if (patterns[i].test(line)) {
      return true;
    }
  }
  
  return false;
}

/**
 * 除外パターンかどうかを判定する（改良版）
 * @param {string} content コンテンツ
 * @returns {boolean} 処理をスキップすべきかどうか
 */
function shouldSkipProcessing(content) {
  if (!content || typeof content !== 'string') {
    return false;
  }
  
  const skipPatterns = [
    /^VaR\s+OUTPUT$/i,                    // VaR OUTPUT（完全一致）
    /VaR\s+OUTPUT/i,                      // VaR OUTPUTを含む（部分一致）
    /^[A-Z]{2,}\s+[A-Z]{2,}$/,           // 連続する大文字単語のみ
    /^[A-Z]+\s+[A-Z]+$/,                 // 略語の組み合わせのみ
    /^\d+\.\d+$/,                        // 小数点を含む数値のみ
    /^Version\s+\d+\.\d+$/i,             // バージョン番号のみ
  ];
  
  return skipPatterns.some(function(pattern) {
    return pattern.test(content.trim());
  });
}

/**
 * 行からリスト項目のマーカーとコンテンツを正確に抽出
 * @param {string} line 1行のテキスト
 * @returns {Object|null} {marker, content} または null
 */
function findListItemMatch(line) {
  if (!line || typeof line !== 'string') {
    return null;
  }
  
  const patterns = [
    /^(\s*[•\-\*▪▫]\s+)(.+)$/,
    /^(\s*\d{1,2}[\.\)]\s+)(.+)$/,
    /^(\s*[a-zA-Z][\.\)]\s+)(.+)$/,
    /^(\s*[ivxlcdmIVXLCDM]{1,4}[\.\)]\s+)(.+)$/i
  ];
  
  for (let i = 0; i < patterns.length; i++) {
    const pattern = patterns[i];
    const match = line.match(pattern);
    if (match) {
      return {
        marker: match[1],
        content: match[2]
      };
    }
  }
  
  return null;
}

/**
 * リスト項目の改行パターンを識別する
 * @param {string} text テキスト
 * @returns {Object} リスト情報オブジェクト
 */
function analyzeListStructure(text) {
  if (!text || typeof text !== 'string') {
    return { isListText: false, listType: null, items: [], totalItems: 0 };
  }

  const contentExcludePatterns = [
    /^VaR\s+OUTPUT$/i,
    /^[A-Z]{2,}\s+[A-Z]{2,}$/,
    /^[A-Z]+\s+[A-Z]+$/,
  ];

  const listPatterns = {
    bullet: /^(\s*[•\-\*▪▫]\s+)(.+)$/gm,
    numbered: /^(\s*\d{1,2}[\.\)]\s+)(.+)$/gm,
    alpha: /^(\s*[a-zA-Z][\.\)]\s+)(.+)$/gm,
    roman: /^(\s*[ivxlcdmIVXLCDM]{1,4}[\.\)]\s+)(.+)$/gm
  };

  let listType = null;
  let items = [];
  let isListText = false;

  for (const typeName in listPatterns) {
    const pattern = listPatterns[typeName];
    const matches = [...text.matchAll(pattern)];
    if (matches.length > 0) {
      
      const validMatches = matches.filter(function(match) {
        const markerPart = match[1];
        const contentPart = match[2];
        
        if (!contentPart || contentPart.trim().length === 0) return false;
        if (contentPart.trim().length < 2) return false;
        if (/^\d+\.?\d*$/.test(contentPart.trim())) return false;
        
        return true;
      });
      
      if (validMatches.length >= 1) {
        isListText = true;
        listType = typeName;
        items = validMatches.map(function(match) {
          return {
            fullMatch: match[0],
            marker: match[1].trim(),
            content: match[2],
            index: match.index,
            isExcludedContent: contentExcludePatterns.some(function(pattern) {
              return pattern.test(match[2].trim());
            })
          };
        });
        break;
      }
    }
  }

  if (isListText && items.length > 0) {
    items = items.map(function(item) {
      return {
        fullMatch: item.fullMatch,
        marker: item.marker,
        content: item.content,
        index: item.index,
        isExcludedContent: item.isExcludedContent,
        indentLevel: (item.marker.match(/^\s*/) ? item.marker.match(/^\s*/)[0] : '').length,
        hasSubItems: false
      };
    });

    for (let i = 0; i < items.length - 1; i++) {
      if (items[i + 1] && items[i + 1].indentLevel > items[i].indentLevel) {
        items[i].hasSubItems = true;
      }
    }
  }

  return {
    isListText: isListText,
    listType: listType,
    items: items,
    hasMultipleLevels: items.some(function(item) { return item.indentLevel > 0; }),
    totalItems: items.length
  };
}

/**
 * 新しい処理関数（isListItemNew使用版・改行修正）
 */
function processTextWithListPreservationNew(text, processingFunction) {
  if (!text || typeof text !== 'string' || text.trim() === '') {
    return text;
  }

  const listInfo = analyzeListStructure(text);
  
  if (!listInfo.isListText) {
    return processingFunction(text);
  }

  console.log('リスト処理開始: ' + listInfo.totalItems + '項目');

  const lines = text.split('\n');
  const processedLines = [];
  
  for (let lineIndex = 0; lineIndex < lines.length; lineIndex++) {
    const line = lines[lineIndex];
    
    if (line.trim() === '') {
      processedLines.push(line);
      continue;
    }
    
    // 新しい判定関数を使用
    const lineIsListItem = isListItemNew(line);
    console.log('行判定: "' + line + '" → ' + lineIsListItem);
    
    if (lineIsListItem) {
      const itemMatch = findListItemMatch(line);
      
      if (itemMatch) {
        const marker = itemMatch.marker;
        const content = itemMatch.content;
        
        console.log('処理中の項目: マーカー="' + marker + '", コンテンツ="' + content + '"');
        
        if (content.trim()) {
          if (shouldSkipProcessing(content)) {
            processedLines.push(line);  // 行全体をそのまま保持
            console.log('  -> 処理をスキップしました');
          } else {
            const processedContent = processingFunction(content);
            // ★修正：マーカーとコンテンツを確実に1行で結合
            processedLines.push(marker + processedContent);
            console.log('  -> 処理しました: "' + processedContent + '"');
          }
        } else {
          processedLines.push(line);
        }
      } else {
        console.log('マーカーパース失敗: "' + line + '"');
        processedLines.push(processingFunction(line));
      }
    } else {
      // 非リスト項目として処理
      if (line.trim()) {
        processedLines.push(processingFunction(line));
      } else {
        processedLines.push(line);
      }
    }
  }
  
  let result = processedLines.join('\n');
  
  // ★追加：後処理で余分な改行を削除
  result = cleanupBulletFormatting(result);
  
  console.log('リスト処理完了');
  return result;
}


/**
 * 箇条書きの改行を整形する（新規追加）
 */
function cleanupBulletFormatting(text) {
  if (!text || typeof text !== 'string') return text;
  
  // パターン1: マーカーの直後に改行がある場合を修正
  // "• \n本文" → "• 本文"
  text = text.replace(/([•\-\*▪▫])\s*\n+\s*/g, '$1 ');
  
  // パターン2: 番号付きリストマーカーの直後に改行がある場合
  // "1. \n本文" → "1. 本文"
  text = text.replace(/(\d+[\.\)])\s*\n+\s*/g, '$1 ');
  
  // パターン3: アルファベットリストマーカーの直後に改行がある場合
  // "a. \n本文" → "a. 本文"
  text = text.replace(/([a-zA-Z][\.\)])\s*\n+\s*/g, '$1 ');
  
  return text;
}


/**
 * リスト対応版：保護機能付きGoogle翻訳（改行修正版）
 */
function translateWithGoogleListAwareNew(text, src, tgt, enableProofread = true) {
  if (!text || text.trim() === '' || text.match(/^[\n\r\s•\-\*\d\.\)\(\s]*$/)) {
    return text;
  }

  // === 前処理：英語の箇条書き行にマーカーを付与 ===
  const BULLET_MARKER = '\u2060';
  let preprocessedText = text;
  
  if (src === 'en' && tgt === 'ja') {
    const lines = text.split('\n');
    const markedLines = [];
    
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      
      if (line.trim() === '') {
        markedLines.push(line);
        continue;
      }
      
      const isBulletLine = isListItemNew(line);
      if (!isBulletLine) {
        markedLines.push(line);
        continue;
      }
      
      const itemMatch = findListItemMatch(line);
      if (!itemMatch) {
        markedLines.push(line);
        continue;
      }
      
      const content = itemMatch.content.trim();
      
      if (shouldSkipProcessing(content)) {
        markedLines.push(line);
        continue;
      }
      
      const isBulletish = (
        content.length >= 3 &&
        content.length <= 80 &&
        !/[.!?]{2,}/.test(content) &&
        (content.match(/[.!?]/g) || []).length <= 1 &&
        !/^(Version|VaR OUTPUT|[A-Z]{2,}\s+[A-Z]{2,})$/i.test(content)
      );
      
      if (isBulletish) {
        console.log(`[前処理] マーカー付与: "${content.substring(0, 40)}..."`);
        // ★修正：改行を含まないようにマーカーを付与
        markedLines.push(line.trimEnd() + BULLET_MARKER);
      } else {
        markedLines.push(line);
      }
    }
    
    preprocessedText = markedLines.join('\n');
  }

  // === 翻訳処理 ===
  const translatedText = processTextWithListPreservationNew(preprocessedText, function(textToTranslate) {
    return protectAndTranslateWithReplacement(textToTranslate, function(protectedText) {
      return LanguageApp.translate(protectedText, src, tgt);
    });
  });

  // === 後処理：マーカーがある行の句点を削除 + 改行整形 ===
  let cleanedText = translatedText;
  if (tgt === 'ja') {
    const lines = translatedText.split('\n');
    const cleanedLines = [];
    
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      
      if (line.includes(BULLET_MARKER)) {
        let cleaned = line.replace(BULLET_MARKER, '');
        
        // 行末の句点を削除
        const bulletMatch = cleaned.match(/^(\s*[•\-\*\d\.\)]+\s*)(.+)$/);
        if (bulletMatch) {
          const marker = bulletMatch[1];
          const content = bulletMatch[2].replace(/。\s*$/, '').trim();
          // ★修正：マーカーとコンテンツを確実に1行で結合
          cleaned = marker + content;
          console.log(`[後処理] 句点削除: "${content.substring(0, 40)}..."`);
        }
        
        cleanedLines.push(cleaned);
      } else {
        cleanedLines.push(line);
      }
    }
    
    cleanedText = cleanedLines.join('\n');
    
    // ★追加：最終的な改行整形
    cleanedText = cleanupBulletFormatting(cleanedText);
  } else {
    cleanedText = translatedText.replace(new RegExp(BULLET_MARKER, 'g'), '');
  }

  // === 校正処理 ===
  let finalResult = cleanedText;
  if (tgt === 'ja' && enableProofread) {
    finalResult = basicJapaneseProofread(cleanedText);
  }
  
  Utilities.sleep(25);
  return finalResult;
}

/**
 * 新版：書式保持版シェイプ翻訳
 */
function translateShapeWithListAwareFormatPreservationNew(shape, slide, src, tgt, enableProofread = true) {
  const textRange = shape.getText();
  const fullText = textRange.asString();
  
  if (fullText.trim() === '') return;
  
  try {
    const paragraphs = textRange.getParagraphs();
    const listInfo = analyzeListStructure(fullText);
    
    if (listInfo.isListText) {
      console.log('リスト構造を検出: ' + listInfo.listType + ', ' + listInfo.totalItems + '項目');
      
      const translatedText = translateWithGoogleListAwareNew(fullText, src, tgt, enableProofread);
      
      const originalLines = fullText.split('\n');
      const translatedLines = translatedText.split('\n');
      
      console.log('元の行数: ' + originalLines.length + ', 翻訳後の行数: ' + translatedLines.length);
      
      textRange.setText(translatedText);
      
      // 基本的な書式復元
      try {
        if (paragraphs.length > 0) {
          const defaultFormat = collectParagraphFormat(paragraphs[0]);
          const newParagraphs = textRange.getParagraphs();
          for (let i = 0; i < newParagraphs.length; i++) {
            restoreParagraphFormat(newParagraphs[i], defaultFormat);
          }
        }
      } catch (error) {
        console.log('リスト書式復元エラー:', error);
      }
      
      // ★追加：リスト翻訳後のフォントサイズ調整
      if (tgt === 'ja') {
        const newParagraphs = textRange.getParagraphs();
        for (let i = 0; i < newParagraphs.length; i++) {
          const paragraphRange = newParagraphs[i].getRange();
          if (paragraphRange.asString().trim() !== '') {
            reduceTextSize(paragraphRange);
          }
        }
      }
      
    } else {
      // 通常のテキスト処理
      for (let i = 0; i < paragraphs.length; i++) {
        const paragraph = paragraphs[i];
        const paragraphRange = paragraph.getRange();
        const originalText = paragraphRange.asString();
        
        if (originalText.trim() === '') continue;
        
        const paragraphFormat = collectParagraphFormat(paragraph);
        const characterFormats = collectCharacterFormats(paragraphRange);
        
        const translatedText = translateWithGoogleListAwareNew(originalText, src, tgt, enableProofread);
        
        paragraphRange.setText(translatedText);
        
        restoreParagraphFormat(paragraph, paragraphFormat);
        restoreCharacterFormats(paragraphRange, characterFormats, originalText, translatedText);
        
        // ★追加：通常テキスト翻訳後のフォントサイズ調整
        if (tgt === 'ja' && paragraphRange.asString().trim() !== '') {
          reduceTextSize(paragraphRange);
        }
      }
    }
    
  } catch (error) {
    console.log('新版シェイプ翻訳エラー:', error);
    const translatedText = translateWithGoogleListAwareNew(fullText, src, tgt, enableProofread);
    textRange.setText(translatedText);
    
    // ★追加：エラー時のフォントサイズ調整
    if (tgt === 'ja') {
      reduceTextSize(textRange);
    }
  }
}
/**
 * 元の関数名との互換性のためのエイリアス
 */
function translateShapeWithListAwareFormatPreservation(shape, slide, src, tgt, enableProofread = true) {
  return translateShapeWithListAwareFormatPreservationNew(shape, slide, src, tgt, enableProofread);
}

/**
 * その他の互換性エイリアス
 */
function translateWithGoogleListAware(text, src, tgt) {
  return translateWithGoogleListAwareNew(text, src, tgt);
}

function processTextWithListPreservation(text, processingFunction) {
  return processTextWithListPreservationNew(text, processingFunction);
}

function isListItem(line) {
  return isListItemNew(line);
}

// ---------

/**
 * カスタム除外辞書を取得する
 * @returns {Object} カスタム除外辞書
 */
function getCustomExclusions() {
  try {
    const customExclusionsJson = PropertiesService.getScriptProperties().getProperty('CUSTOM_EXCLUSIONS');
    return customExclusionsJson ? JSON.parse(customExclusionsJson) : {};
  } catch (error) {
    console.log('カスタム除外辞書の読み込みエラー:', error);
    return {};
  }
}

/**
 * カスタム除外辞書を保存する
 * @param {Object} exclusions カスタム除外辞書
 */
function saveCustomExclusions(exclusions) {
  try {
    PropertiesService.getScriptProperties().setProperty('CUSTOM_EXCLUSIONS', JSON.stringify(exclusions));
  } catch (error) {
    console.log('カスタム除外辞書の保存エラー:', error);
  }
}

/**
 * 翻訳除外辞書を管理する
 */
function manageTranslationExclusions() {
  const ui = SlidesApp.getUi();
  
  const response = ui.alert(
    '翻訳除外辞書管理',
    '翻訳から除外する単語を管理します。何を行いますか？',
    ui.ButtonSet.YES_NO_CANCEL
  );
  
  if (response === ui.Button.YES) {
    addCustomExclusion();
  } else if (response === ui.Button.NO) {
    showCurrentExclusions();
  }
}

/**
 * カスタム除外単語を追加する
 */
function addCustomExclusion() {
  const ui = SlidesApp.getUi();
  
  const response = ui.prompt(
    '新しい除外単語を追加',
    '翻訳から除外したい単語を入力してください:\n例: MyService, CustomTool, SpecialAPI',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const newTerm = response.getResponseText().trim();
    if (newTerm) {
      const customExclusions = getCustomExclusions();
      customExclusions[newTerm] = newTerm;
      saveCustomExclusions(customExclusions);
      
      ui.alert(`「${newTerm}」を翻訳除外辞書に追加しました！`);
    } else {
      ui.alert('単語が入力されていません。');
    }
  }
}

/**
 * 現在の除外辞書を表示する
 */
function showCurrentExclusions() {
  const ui = SlidesApp.getUi();
  const customExclusions = getCustomExclusions();
  const allExclusions = { ...TRANSLATION_EXCLUSIONS, ...customExclusions };
  
  const defaultTerms = Object.keys(TRANSLATION_EXCLUSIONS);
  const customTerms = Object.keys(customExclusions);
  
  let message = '【翻訳除外辞書】\n\n';
  
  message += `■ デフォルト除外用語 (${defaultTerms.length}個):\n`;
  message += defaultTerms.slice(0, 20).join(', ');
  if (defaultTerms.length > 20) {
    message += `\n...他 ${defaultTerms.length - 20}個`;
  }
  
  if (customTerms.length > 0) {
    message += `\n\n■ カスタム除外用語 (${customTerms.length}個):\n`;
    message += customTerms.join(', ');
    message += '\n\n※カスタム用語を削除する場合は、スクリプトエディタで手動削除してください。';
  } else {
    message += '\n\n■ カスタム除外用語: なし';
  }
  
  ui.alert('翻訳除外辞書', message, ui.ButtonSet.OK);
}

// Databricks API設定
const DATABRICKS_TOKEN = PropertiesService.getScriptProperties().getProperty('DATABRICKS_TOKEN');
const DATABRICKS_WORKSPACE = PropertiesService.getScriptProperties().getProperty('DATABRICKS_WORKSPACE');
const DATABRICKS_MODEL_ENDPOINT = 'databricks-meta-llama-3-1-70b-instruct';

/**
 * 翻訳除外辞書 - 翻訳しない単語・用語（置換機能付き）
 * キーと値が異なる場合は置換、同じ場合は保護のみ
 */
const TRANSLATION_EXCLUSIONS = {
// 置換機能例
'permission':'制限',  // permission → 制限 に置換
'Authentication':'認証',  // Authentication → 認証 に置換
'Authorization':'認可',  // 認可に置換

// 追加
'production':'プロダクション環境', 
'maintenance tax':'維持費', 
'Reverse-ETL':'Reverse-ETL',
'Reverse ETL':'Reverse ETL',
'Lock-in':'Lock-in', 
'lock-in':'lock-in', 
'Lakebase':'Lakebase',
'lakebase':'lakebase',
'Tick Data':'Tick Data',
'Refined':'Refined',
'Raw': 'Raw',
'Aggregated': 'Aggregated',
'Query and Process': 'Query and Process',
'Transform': 'Transform',
'Ingest': 'Ingest',
'Delta Lake': 'Delta Lake',
// 'Serve': 'Serve',
'Platinum layer': 'Platinum layer',
'Analysis/Output': 'Analysis/Output',
// Databricks 公式
'4X-Large': '4X-Large',
'ACL': 'ACL',
'ACID': 'ACID',
'2X-Small': '2X-Small',
'2X-Large': '2X-Large',
'ADLS': 'ADLS',
'3X-Large': '3X-Large',
'ACR': 'ACR',
'Apache Kudu': 'Apache Kudu',
'Apache Spark™': 'Apache Spark™',
'Apache Software Foundation': 'Apache Software Foundation',
'Apache': 'Apache',
'ARNs': 'ARNs',
'APIs': 'APIs',
'ASF': 'ASF',
'Apache Hive': 'Apache Hive',
'Auto Loader': 'Auto Loader',
'AWS Marketplace': 'AWS Marketplace',
'Azure Data Lake Storage Gen2': 'Azure Data Lake Storage Gen2',
'Apache Kylin': 'Apache Kylin',
'Apache Spark': 'Apache Spark',
'Avro': 'Avro',
'AVRO': 'AVRO',
'Azure': 'Azure',
'AWS': 'AWS',
'Azure Synapse Analytics': 'Azure Synapse Analytics',
'Azure SQL Data Warehouse': 'Azure SQL Data Warehouse',
'Azure Databricks': 'Azure Databricks',
'Bayesian': 'Bayesian',
'BINARYFILE': 'BINARYFILE',
'Azure Key Vault': 'Azure Key Vault',
'Azure Blob': 'Azure Blob',
'Azure Event Hub': 'Azure Event Hub',
'BigQuery': 'BigQuery',
'Bokeh': 'Bokeh',
'Cassandra': 'Cassandra',
'CDC': 'CDC',
'Boxplot': 'Boxplot',
'Bitbucket': 'Bitbucket',
'CIDR': 'CIDR',
'Azure Cosmos DB': 'Azure Cosmos DB',
'ARN': 'ARN',
'Azure Data Lake Storage Gen1': 'Azure Data Lake Storage Gen1',
'Azure Active Directory': 'Azure Active Directory',
'API': 'API',
'Azure Synapse': 'Azure Synapse',
'Berkeley Packet Filter': 'Berkeley Packet Filter',
'Conda': 'Conda',
'BPF': 'BPF',
'CI/CD': 'CI/CD',
'CNN': 'CNN',
'Azure Data Lake Store': 'Azure Data Lake Store',
'CRAN': 'CRAN',
'Community Edition': 'Community Edition',
'Couchbase': 'Couchbase',
'Data Explorer': 'Data Explorer',
'data source': 'data source',
'CSV': 'CSV',
'Databricks Community Edition': 'Databricks Community Edition',
'Databricks Delta': 'Databricks Delta',
'Databricks Filesystem': 'Databricks Filesystem',
'Databricks DataInsight': 'Databricks DataInsight',
'Databricks': 'Databricks',
'Databricks Machine Learning': 'Databricks Machine Learning',
'Databricks Data Analytics Platform': 'Databricks Data Analytics Platform',
'DataFrames': 'DataFrames',
'DBU': 'DBU',
'Databricks Runtime': 'Databricks Runtime',
'Databricks Platform Free Trial': 'Databricks Platform Free Trial',
'Databricks Container Services': 'Databricks Container Services',
'DataFrame': 'DataFrame',
'data scientists': 'data scientists',
'Databricks SQL': 'Databricks SQL',
'Delta': 'Delta',
'Delta Lake': 'Delta Lake',
'Delta Sharing': 'Delta Sharing',
'Databricks Workflows': 'Databricks Workflows',
'DBFS': 'DBFS',
'DevOps': 'DevOps',
'DLT': 'DLT',
'Donut': 'Donut',
'Docker': 'Docker',
'Docker Container Services': 'Docker Container Services',
'eBay': 'eBay',
'Delta Engine': 'Delta Engine',
'EC2': 'EC2',
'Amazon ECR': 'Amazon ECR',
'Amazon Aurora': 'Amazon Aurora',
'Aurora': 'Aurora',
'Databricks Workspace': 'Databricks Workspace',
'EMR': 'EMR',
'Elasticsearch': 'Elasticsearch',
'ETL': 'ETL',
'GCS': 'GCS',
'Excel': 'Excel',
'GATK': 'GATK',
'Google Cloud Storage': 'Google Cloud Storage',
'GitHub': 'GitHub',
'Hadoop': 'Hadoop',
'HIPAA': 'HIPAA',
'Hive': 'Hive',
'eQTL': 'eQTL',
'GovCloud': 'GovCloud',
'Git': 'Git',
'Google Spreadsheets': 'Google Spreadsheets',
'Hive SerDes': 'Hive SerDes',
'HMS': 'HMS',
'Horovod': 'Horovod',
'Hugging Face': 'Hugging Face',
'IAM': 'IAM',
'HP': 'HP',
'IoT': 'IoT',
'Java': 'Java',
'JAR': 'JAR',
'iPass': 'iPass',
'JVM': 'JVM',
'JSON': 'JSON',
'Kafka': 'Kafka',
'GraphX': 'GraphX',
'HDFS': 'HDFS',
'Fivetran': 'Fivetran',
'GCP': 'GCP',
'Lambda': 'Lambda',
'Kinesis': 'Kinesis',
'LZO': 'LZO',
'MapReduce': 'MapReduce',
'HomeAway': 'HomeAway',
'Keras': 'Keras',
'Hive metastore': 'Hive metastore',
// 'IT': 'IT',
'Looker': 'Looker',
'MLlib': 'MLlib',
'Maven': 'Maven',
'Koalas': 'Koalas',
'KMS': 'KMS',
'MongoDB': 'MongoDB',
'MongoDB Connector for Spark': 'MongoDB Connector for Spark',
'Microsoft': 'Microsoft',
'MLOps': 'MLOps',
'Netflix': 'Netflix',
'NoSQL': 'NoSQL',
'Nephos': 'Nephos',
'NIST': 'NIST',
'MongoDB Atlas': 'MongoDB Atlas',
'NumPy': 'NumPy',
'MySQL': 'MySQL',
'ODBC': 'ODBC',
'Matplotlib': 'Matplotlib',
'MLflow': 'MLflow',
'ORC': 'ORC',
'Pandas': 'Pandas',
'Partner Connect': 'Partner Connect',
'Partner Hub': 'Partner Hub',
'Parquet': 'Parquet',
'PARQUET': 'PARQUET',
'Pandas DataFrame': 'Pandas DataFrame',
'Photon': 'Photon',
'PHI': 'PHI',
'PR': 'PR',
'PostgreSQL': 'PostgreSQL',
'PySpark': 'PySpark',
'PyPI': 'PyPI',
'PyCharm': 'PyCharm',
'Qlik Sense': 'Qlik Sense',
'Python': 'Python',
'R': 'R',
'PyTorch': 'PyTorch',
'Regeneron': 'Regeneron',
'Redshift': 'Redshift',
'ROC': 'ROC',
'RStudio': 'RStudio',
'repo': 'repo',
'SAML': 'SAML',
'SCD': 'SCD',
'Plotly': 'Plotly',
'RDBMS': 'RDBMS',
'RDD': 'RDD',
'SerDes': 'SerDes',
'Scala': 'Scala',
'S3': 'S3',
'Redash': 'Redash',
'Runtime': 'Runtime',
'REST': 'REST',
'SCIM': 'SCIM',
'Spark UI': 'Spark UI',
'Spark SQL': 'Spark SQL',
'SparkHub': 'SparkHub',
'Spark': 'Spark',
'Scikit-Learn': 'Scikit-Learn',
'SLA': 'SLA',
'SparkR': 'SparkR',
'SQL': 'SQL',
'Sparklyr': 'Sparklyr',
'SSD': 'SSD',
'Repos': 'Repos',
'SaaS': 'SaaS',
'Tableau': 'Tableau',
'r4.4large': 'r4.4large',
'Synapse': 'Synapse',
'Terraform': 'Terraform',
'Snowflake': 'Snowflake',
'Spark API': 'Spark API',
'TensorBoard': 'TensorBoard',
'TensorFlow': 'TensorFlow',
'TEXT': 'TEXT',
'Tungsten': 'Tungsten',
'unified analytics': 'unified analytics',
'Unity Catalog': 'Unity Catalog',
'Unified Data Analytics': 'Unified Data Analytics',
'X-Small': 'X-Small',
'XGBoost': 'XGBoost',
'VPC': 'VPC',
'Yahoo': 'Yahoo',
'X-Large': 'X-Large',
'Kudu': 'Kudu',
'Kylin': 'Kylin',
'Model Registry': 'Model Registry',
'azure': 'azure',
'Workspace Model Registry': 'Workspace Model Registry',
'MLflow Model Registry': 'MLflow Model Registry',
'Databricks Marketplace': 'Databricks Marketplace',
'SQL Server': 'SQL Server',
'CLI': 'CLI',
'Resilient Distributed Datasets': 'Resilient Distributed Datasets',
'write.parquet': 'write.parquet',
'write.orc': 'write.orc',
'write.text': 'write.text',
'read.json': 'read.json',
'read.parquet': 'read.parquet',
'read.orc': 'read.orc',
'read.text': 'read.text',
'read.jdbc': 'read.jdbc',
'change data feed': 'change data feed',
'external location': 'external location',
'Low-Med': 'Low-Med',
'Med-High': 'Med-High',
'CAN ATTACH TO': 'CAN ATTACH TO',
'CAN RESTART': 'CAN RESTART',
// 'production': 'production',
'warehouse': 'warehouse',
'Amazon': 'Amazon',
'Airflow': 'Airflow',
'Databricks Repo': 'Databricks Repo',
'SDK': 'SDK',
'SBT': 'SBT',
'DHCP': 'DHCP',
'FSCK': 'FSCK',
'PowerShell': 'PowerShell',
'zsh': 'zsh',
'IaC': 'IaC',
'Msck': 'Msck',
'TSV': 'TSV',
'SSL': 'SSL',
'UID': 'UID',
'PWD': 'PWD',
'UDF': 'UDF',
'rh': 'rh',
'ML': 'ML',
'LTS': 'LTS',
'write.json': 'write.json',
'COPY INTO': 'COPY INTO',
'Lakehouse Monitoring': 'Lakehouse Monitoring',
'Petastorm': 'Petastorm',
'Hyperopt': 'Hyperopt',
'Prophet': 'Prophet',
'LightGBM': 'LightGBM',
'ARIMA': 'ARIMA',
'Auto-ARIMA': 'Auto-ARIMA',
'Shapley': 'Shapley',
'sklearn': 'sklearn',
'scikit-learn': 'scikit-learn',
'PyTorch Lightning': 'PyTorch Lightning',
'HorovodRunner': 'HorovodRunner',
'TorchDistributor': 'TorchDistributor',
'John Snow Labs': 'John Snow Labs',
'BERT': 'BERT',
'LangChain': 'LangChain',
'ALTER CATALOG': 'ALTER CATALOG',
'ALTER CREDENTIAL': 'ALTER CREDENTIAL',
'ALTER DATABASE': 'ALTER DATABASE',
'ALTER LOCATION': 'ALTER LOCATION',
'ALTER PROVIDER': 'ALTER PROVIDER',
'ALTER RECIPIENT': 'ALTER RECIPIENT',
'ALTER TABLE': 'ALTER TABLE',
'lineage': 'lineage',
'customer-managed VPC': 'customer-managed VPC',
'billable usage log': 'billable usage log',
'billable usage': 'billable usage',
'Boto': 'Boto',
'Z-order': 'Z-order',
'Z-ordering': 'Z-ordering',
'source': 'source',
'ALTER SCHEMA': 'ALTER SCHEMA',
'ALTER SHARE': 'ALTER SHARE',
'ALTER VIEW': 'ALTER VIEW',
'Well-Architected': 'Well-Architected',
'Google Cloud Architecture Framework': 'Google Cloud Architecture Framework',
'COMMENT ON': 'COMMENT ON',
'COMMENT ON PROVIDER': 'COMMENT ON PROVIDER',
'COMMENT ON SHARE': 'COMMENT ON SHARE',
'COMMENT ON RECIPIENT': 'COMMENT ON RECIPIENT',
'COMMENT ON VOLUME': 'COMMENT ON VOLUME',
'COMMENT ON CONNECTION': 'COMMENT ON CONNECTION',
'CREATE BLOOMFILTER INDEX': 'CREATE BLOOMFILTER INDEX',
'CREATE CATALOG': 'CREATE CATALOG',
'CREATE FUNCTION': 'CREATE FUNCTION',
'CREATE LOCATION': 'CREATE LOCATION',
'CREATE RECIPIENT': 'CREATE RECIPIENT',
'CREATE SCHEMA': 'CREATE SCHEMA',
'CREATE TABLE CLONE': 'CREATE TABLE CLONE',
'liquid clustering': 'liquid clustering',
'CREATE SHARE': 'CREATE SHARE',
'Iceberg': 'Iceberg',
'disaster recovery': 'disaster recovery',
'CREATE TABLE': 'CREATE TABLE',
'CREATE TABLE LIKE': 'CREATE TABLE LIKE',
'SHOW CREATE TABLE': 'SHOW CREATE TABLE',
'REPLACE TABLE': 'REPLACE TABLE',
'DESC': 'DESC',
'Lakeview': 'Lakeview',
'Network Connectivity Configs': 'Network Connectivity Configs',
'SparkSession': 'SparkSession',
'GitHub Actions': 'GitHub Actions',
'Snowplow': 'Snowplow',
'Arcion': 'Arcion',
'Hevo': 'Hevo',
'Hevo Data': 'Hevo Data',
'Infoworks': 'Infoworks',
'DataFoundry': 'DataFoundry',
'Qlik': 'Qlik',
'Qlik Replicate': 'Qlik Replicate',
'Qlik Compose': 'Qlik Compose',
'Rivery': 'Rivery',
'Stitch': 'Stitch',
'StreamSets': 'StreamSets',
'Syncsort': 'Syncsort',
'dbt': 'dbt',
'dbt Cloud': 'dbt Cloud',
'dbt Core': 'dbt Core',
'Matillion': 'Matillion',
'Prophecy': 'Prophecy',
'Labelbox': 'Labelbox',
'Hex': 'Hex',
'MicroStrategy': 'MicroStrategy',
'Mode': 'Mode',
'Sigma': 'Sigma',
'SQL Workbench/J': 'SQL Workbench/J',
'ThoughtSpot': 'ThoughtSpot',
'TIBCO Spotfire': 'TIBCO Spotfire',
'TIBCO Spotfire Analyst': 'TIBCO Spotfire Analyst',
'Census': 'Census',
'Hightouch': 'Hightouch',
'AtScale': 'AtScale',
'Stardog': 'Stardog',
'Alation': 'Alation',
'Anomalo': 'Anomalo',
'erwin Data Modeler': 'erwin Data Modeler',
'erwin Data Modeler by Quest': 'erwin Data Modeler by Quest',
'Lightup': 'Lightup',
'Hunters': 'Hunters',
'Privacera': 'Privacera',
'CREATE VIEW': 'CREATE VIEW',
'Microsoft Entra ID': 'Microsoft Entra ID',
'Databricks Connect': 'Databricks Connect',
'Windows': 'Windows',
'RudderStack': 'RudderStack',
'WindowSpecDefinition': 'WindowSpecDefinition',
'2.0.61.Final-windows-x86_64': '2.0.61.Final-windows-x86_64',
'feature table': 'feature table',
'deletion vector': 'deletion vector',
'Vector Search': 'Vector Search',
'Z-ordered': 'Z-ordered',
'Poetry': 'Poetry',
'Generative AI': 'Generative AI',
'large language model': 'large language model',
'LLM': 'LLM',
'foundation model': 'foundation model',
'fine tuning': 'fine tuning',
'AI Functions': 'AI Functions',
'Transformer': 'Transformer',
'DetrForSegmentation': 'DetrForSegmentation',
'pyodbc': 'pyodbc',
'TABLESAMPLE': 'TABLESAMPLE',
'Node.js': 'Node.js',
'Node Version Manager': 'Node Version Manager',
'CDKTF CLI': 'CDKTF CLI',
'Graviton': 'Graviton',
'UDAFs': 'UDAFs',
'globber': 'globber',
'AI Playground': 'AI Playground',
'Genie': 'Genie',
'AI Functions on Databricks': 'AI Functions on Databricks',
'Eclipse': 'Eclipse',
'IntelliJ IDEA': 'IntelliJ IDEA',
'IDEs': 'IDEs',
'RocksDB': 'RocksDB',
'Guava': 'Guava',
'Protobuf': 'Protobuf',
'AvroDeserializer': 'AvroDeserializer',
'Catalyst': 'Catalyst',
'TreeNode': 'TreeNode',
'DBeaver': 'DBeaver',
'DataGrip': 'DataGrip',
'OAuth': 'OAuth',
'Tableau Cloud': 'Tableau Cloud',
'Tableau Server': 'Tableau Server',
'Infrastructure-as-Code': 'Infrastructure-as-Code',
'bamboolib': 'bamboolib',
'Silver layer': 'Silver layer',
'Gold layer': 'Gold layer',
'Bronze layer': 'Bronze layer',
'medallion architecture': 'medallion architecture',
'vacuum': 'vacuum',
'ALTER CONNECTION': 'ALTER CONNECTION',
'ALTER STREAMING TABLE': 'ALTER STREAMING TABLE',
'ALTER VOLUME': 'ALTER VOLUME',
'CREATE CONNECTION': 'CREATE CONNECTION',
'CREATE DATABASE': 'CREATE DATABASE',
'CREATE MATERIALIZED VIEW': 'CREATE MATERIALIZED VIEW',
'CREATE SERVER': 'CREATE SERVER',
'CREATE STREAMING TABLE': 'CREATE STREAMING TABLE',
'CREATE VOLUME': 'CREATE VOLUME',
'DECLARE VARIABLE': 'DECLARE VARIABLE',
'DROP BLOOMFILTER INDEX': 'DROP BLOOMFILTER INDEX',
'DROP CATALOG': 'DROP CATALOG',
'DROP CONNECTION': 'DROP CONNECTION',
'DROP DATABASE': 'DROP DATABASE',
'DROP CREDENTIAL': 'DROP CREDENTIAL',
'DROP FUNCTION': 'DROP FUNCTION',
'DROP LOCATION': 'DROP LOCATION',
'DROP PROVIDER': 'DROP PROVIDER',
'DROP RECIPIENT': 'DROP RECIPIENT',
'DROP SCHEMA': 'DROP SCHEMA',
'DROP SHARE': 'DROP SHARE',
'DROP TABLE': 'DROP TABLE',
'DROP VARIABLE': 'DROP VARIABLE',
'DROP VIEW': 'DROP VIEW',
'DROP VOLUME': 'DROP VOLUME',
'MSCK REPAIR TABLE': 'MSCK REPAIR TABLE',
'REFRESH FOREIGN': 'REFRESH FOREIGN',
'REFRESH': 'REFRESH',
'MATERIALIZED VIEW': 'MATERIALIZED VIEW',
'STREAMING TABLE': 'STREAMING TABLE',
'SYNC': 'SYNC',
'TRUNCATE TABLE': 'TRUNCATE TABLE',
'UNDROP TABLE': 'UNDROP TABLE',
'DELETE FROM': 'DELETE FROM',
'INSERT INTO': 'INSERT INTO',
'INSERT OVERWRITE DIRECTORY': 'INSERT OVERWRITE DIRECTORY',
'LOAD DATA': 'LOAD DATA',
'MERGE INTO': 'MERGE INTO',
'UPDATE': 'UPDATE',
'CACHE SELECT': 'CACHE SELECT',
'CONVERT TO DELTA': 'CONVERT TO DELTA',
'DESCRIBE HISTORY': 'DESCRIBE HISTORY',
'FSCK REPAIR TABLE': 'FSCK REPAIR TABLE',
'GENERATE': 'GENERATE',
'OPTIMIZE': 'OPTIMIZE',
'REORG TABLE': 'REORG TABLE',
'RESTORE': 'RESTORE',
'ANALYZE TABLE': 'ANALYZE TABLE',
'CACHE TABLE': 'CACHE TABLE',
'CLEAR CACHE': 'CLEAR CACHE',
'REFRESH CACHE': 'REFRESH CACHE',
'REFRESH FUNCTION': 'REFRESH FUNCTION',
'REFRESH TABLE': 'REFRESH TABLE',
'UNCACHE TABLE': 'UNCACHE TABLE',
'DESCRIBE CATALOG': 'DESCRIBE CATALOG',
'DESCRIBE CONNECTION': 'DESCRIBE CONNECTION',
'DESCRIBE CREDENTIAL': 'DESCRIBE CREDENTIAL',
'DESCRIBE DATABASE': 'DESCRIBE DATABASE',
'DESCRIBE FUNCTION': 'DESCRIBE FUNCTION',
'DESCRIBE LOCATION': 'DESCRIBE LOCATION',
'DESCRIBE PROVIDER': 'DESCRIBE PROVIDER',
'DESCRIBE QUERY': 'DESCRIBE QUERY',
'DESCRIBE RECIPIENT': 'DESCRIBE RECIPIENT',
'DESCRIBE SCHEMA': 'DESCRIBE SCHEMA',
'DESCRIBE SHARE': 'DESCRIBE SHARE',
'DESCRIBE TABLE': 'DESCRIBE TABLE',
'DESCRIBE VOLUME': 'DESCRIBE VOLUME',
'LIST': 'LIST',
'SHOW ALL IN SHARE': 'SHOW ALL IN SHARE',
'SHOW CATALOGS': 'SHOW CATALOGS',
'SHOW COLUMNS': 'SHOW COLUMNS',
'SHOW CONNECTIONS': 'SHOW CONNECTIONS',
'SHOW CREDENTIALS': 'SHOW CREDENTIALS',
'SHOW DATABASES': 'SHOW DATABASES',
'SHOW FUNCTIONS': 'SHOW FUNCTIONS',
'SHOW GROUPS': 'SHOW GROUPS',
'SHOW LOCATIONS': 'SHOW LOCATIONS',
'SHOW PARTITIONS': 'SHOW PARTITIONS',
'SHOW PROVIDERS': 'SHOW PROVIDERS',
'SHOW RECIPIENTS': 'SHOW RECIPIENTS',
'SHOW SCHEMAS': 'SHOW SCHEMAS',
'SHOW SHARES': 'SHOW SHARES',
'SHOW SHARES IN PROVIDER': 'SHOW SHARES IN PROVIDER',
'SHOW TABLE': 'SHOW TABLE',
'SHOW TABLES': 'SHOW TABLES',
'SHOW TABLES DROPPED': 'SHOW TABLES DROPPED',
'SHOW TBLPROPERTIES': 'SHOW TBLPROPERTIES',
'SHOW USERS': 'SHOW USERS',
'SHOW VIEWS': 'SHOW VIEWS',
'SHOW VOLUMES': 'SHOW VOLUMES',
'RESET': 'RESET',
'SET': 'SET',
'SET TIMEZONE': 'SET TIMEZONE',
'SET VARIABLE': 'SET VARIABLE',
'USE CATALOG': 'USE CATALOG',
'USE DATABASE': 'USE DATABASE',
'USE SCHEMA': 'USE SCHEMA',
'ADD ARCHIVE': 'ADD ARCHIVE',
'ADD FILE': 'ADD FILE',
'ADD JAR': 'ADD JAR',
'LIST ARCHIVE': 'LIST ARCHIVE',
'LIST FILE': 'LIST FILE',
'LIST JAR': 'LIST JAR',
'ALTER GROUP': 'ALTER GROUP',
'CREATE GROUP': 'CREATE GROUP',
'DENY': 'DENY',
'DROP GROUP': 'DROP GROUP',
'GRANT': 'GRANT',
'GRANT SHARE': 'GRANT SHARE',
'REPAIR PRIVILEGES': 'REPAIR PRIVILEGES',
'REVOKE': 'REVOKE',
'REVOKE SHARE': 'REVOKE SHARE',
'SHOW GRANTS': 'SHOW GRANTS',
'SHOW GRANTS ON SHARE': 'SHOW GRANTS ON SHARE',
'SHOW GRANTS TO RECIPIENT': 'SHOW GRANTS TO RECIPIENT',
'spark-tensorflow-connector': 'spark-tensorflow-connector',
'GraphFrames': 'GraphFrames',
'sparkdl.xgboost': 'sparkdl.xgboost',
'sparkdl': 'sparkdl',
'predictive optimization': 'predictive optimization',
'EDA': 'EDA',
'GitLab': 'GitLab',
'ai_query': 'ai_query',
'Feature Engineering in Unity Catalog': 'Feature Engineering in Unity Catalog',
'system table': 'system table',
'Dataiku': 'Dataiku',
'Mosaic AI': 'Mosaic AI',
'Databricks Mosaic AI': 'Databricks Mosaic AI',
'Mosaic AI Model Serving': 'Mosaic AI Model Serving',
'Mosaic AI Vector Search': 'Mosaic AI Vector Search',
'Mosaic AI Playground': 'Mosaic AI Playground',
'Mosaic AI Foundational Models API': 'Mosaic AI Foundational Models API',
'Mosaic AI Pretraining': 'Mosaic AI Pretraining',
'Mosaic AI AutoML': 'Mosaic AI AutoML',
'Mosaic AI Functions': 'Mosaic AI Functions',
'DatabricksIQ': 'DatabricksIQ',
'CAN VIEW': 'CAN VIEW',
'NO PERMISSIONS': 'NO PERMISSIONS',
'CAN RUN': 'CAN RUN',
'CAN EDIT': 'CAN EDIT',
'CAN MANAGE': 'CAN MANAGE',
'CAN MANAGE RUN': 'CAN MANAGE RUN',
'IS OWNER': 'IS OWNER',
'CAN QUERY': 'CAN QUERY',
'CAN VIEW METADATA': 'CAN VIEW METADATA',
'CAN EDIT METADATA': 'CAN EDIT METADATA',
'CAN READ': 'CAN READ',
'CAN MANAGE STAGING VERSIONS': 'CAN MANAGE STAGING VERSIONS',
'CAN MANAGE PRODUCTION VERSIONS': 'CAN MANAGE PRODUCTION VERSIONS',
'ai_analyze_sentiment': 'ai_analyze_sentiment',
'ai_classify': 'ai_classify',
'ai_extract': 'ai_extract',
'ai_fix_grammar': 'ai_fix_grammar',
'ai_gen': 'ai_gen',
'ai_mask': 'ai_mask',
'ai_similarity': 'ai_similarity',
'ai_summarize': 'ai_summarize',
'ai_translate': 'ai_translate',
'MPT': 'MPT',
'Llama': 'Llama',
'Mixtral': 'Mixtral',
'Mistral': 'Mistral',
'Feature Serving': 'Feature Serving',
'Amazon Bedrock': 'Amazon Bedrock',
'Anthropic': 'Anthropic',
'Cohere': 'Cohere',
'AI21 Labs': 'AI21 Labs',
'Vertex AI': 'Vertex AI',
'Databricks-to-Databricks': 'Databricks-to-Databricks',
'web terminal': 'web terminal',
'Query Watchdog': 'Query Watchdog',
'ipynb': 'ipynb',
'VPC Service Controls': 'VPC Service Controls',
'Service Perimeter': 'Service Perimeter',
'emergency access': 'emergency access',
'FedRAMP Moderate': 'FedRAMP Moderate',
'docstring-to-markdown': 'docstring-to-markdown',
'rmarkdown': 'rmarkdown',
'read-evaluate-print loop': 'read-evaluate-print loop',
'Databricks Assistant': 'Databricks Assistant',
'Genie space': 'Genie space',
'pay-per-token': 'pay-per-token',
'Hyperparameter tuning': 'Hyperparameter tuning',
'LakeFlow': 'LakeFlow',
'LakeFlow Connect': 'LakeFlow Connect',
'LakeFlow Pipelines': 'LakeFlow Pipelines',
'LakeFlow Job': 'LakeFlow Job',
'decision tree': 'decision tree',
'Databricks Autologging': 'Databricks Autologging',
'AI/BI': 'AI/BI',
'shallow clone': 'shallow clone',
'deep clone': 'deep clone',
'UniForm': 'UniForm',
'managed table': 'managed table',
'unmanaged table': 'unmanaged table',
'external table': 'external table',
'foreign table': 'foreign table',
'Databricks Geos': 'Databricks Geos',
'Geo': 'Geo',
'Geos': 'Geos',
'Databricks Geo': 'Databricks Geo',
'capsule8-alerts-dataplane': 'capsule8-alerts-dataplane',
'Clean Room': 'Clean Room',
'Artifact': 'Artifact',
'AI/BI dashboard': 'AI/BI dashboard',
'Genie spaces': 'Genie spaces',
'Mosaic AI Gateway': 'Mosaic AI Gateway',
'Glue': 'Glue',
'Budget policy': 'Budget policy',
'materialized view': 'materialized view',
'streaming table': 'streaming table',
'SuperAnnotate': 'SuperAnnotate',
'Groq': 'Groq',
'LiteLLM': 'LiteLLM',
'CrewAI': 'CrewAI',
'DSPy': 'DSPy',
'MLflow Tracing': 'MLflow Tracing',
'SME': 'SME',
'Serverless budget policy': 'Serverless budget policy',

  // Databricks プラットフォーム・製品
  'Databricks': 'Databricks',
  'BI': 'BI',
  'Databricks Platform': 'Databricks Platform',
  'Databricks Lakehouse Platform': 'Databricks Lakehouse Platform',
  'Lakehouse': 'Lakehouse',
  'Data Lakehouse': 'Data Lakehouse',
  'Lakehouse Architecture': 'Lakehouse Architecture',
  
  // Databricks Unity Catalog関連
  'Unity Catalog': 'Unity Catalog',
  'UC': 'UC',
  'Unity Catalog Metastore': 'Unity Catalog Metastore',
  'Catalog': 'Catalog',
  'Schema': 'Schema',
  'External Location': 'External Location',
  'Storage Credential': 'Storage Credential',
  'Data Lineage': 'Data Lineage',
  'Data Governance': 'Data Governance',
  
  // Databricks Delta関連
  'Delta Lake': 'Delta Lake',
  'Delta': 'Delta',
  'Delta Table': 'Delta Table',
  'Delta Tables': 'Delta Tables',
  'Delta Sharing': 'Delta Sharing',
  'Delta Live Tables': 'Delta Live Tables',
  'DLT': 'DLT',
  'Delta Engine': 'Delta Engine',
  'Delta Cache': 'Delta Cache',
  'Delta Log': 'Delta Log',
  'Delta Merge': 'Delta Merge',
  'Delta Clone': 'Delta Clone',
  'Delta Time Travel': 'Delta Time Travel',
  'OPTIMIZE': 'OPTIMIZE',
  'VACUUM': 'VACUUM',
  'Z-ORDER': 'Z-ORDER',
  
  // Databricks MLflow関連
  'MLflow': 'MLflow',
  'MLflow Tracking': 'MLflow Tracking',
  'MLflow Projects': 'MLflow Projects',
  'MLflow Models': 'MLflow Models',
  'MLflow Registry': 'MLflow Registry',
  'Model Registry': 'Model Registry',
  'MLflow Serving': 'MLflow Serving',
  'MLflow Experiments': 'MLflow Experiments',
  'MLflow Runs': 'MLflow Runs',
  'MLflow Artifacts': 'MLflow Artifacts',
  
  // Databricks Compute関連
  'Compute': 'Compute',
  'Cluster': 'Cluster',
  'Clusters': 'Clusters',
  'All-Purpose Cluster': 'All-Purpose Cluster',
  'Job Cluster': 'Job Cluster',
  'SQL Warehouse': 'SQL Warehouse',
  'SQL Warehouses': 'SQL Warehouses',
  'Serverless': 'Serverless',
  'Serverless SQL Warehouse': 'Serverless SQL Warehouse',
  'Serverless Compute': 'Serverless Compute',
  'Driver Node': 'Driver Node',
  'Worker Node': 'Worker Node',
  'Autoscaling': 'Autoscaling',
  'Auto Termination': 'Auto Termination',
  'Spot Instance': 'Spot Instance',
  'Preemptible Instance': 'Preemptible Instance',
  'Pool': 'Pool',
  'Instance Pool': 'Instance Pool',
  
  // Databricks Photon関連
  'Photon': 'Photon',
  'Photon Engine': 'Photon Engine',
  'Photon-enabled': 'Photon-enabled',
  
  // Databricks Workflows関連
  'Workflows': 'Workflows',
  'Jobs': 'Jobs',
  'Job': 'Job',
  'Task': 'Task',
  'Multi-task Job': 'Multi-task Job',
  'Workflow Job': 'Workflow Job',
  'Scheduled Job': 'Scheduled Job',
  'Pipeline': 'Pipeline',
  'Data Pipeline': 'Data Pipeline',
  'ETL Pipeline': 'ETL Pipeline',
  
  // Databricks Notebook・開発環境関連
  'Notebook': 'Notebook',
  'Notebooks': 'Notebooks',
  'Databricks Notebook': 'Databricks Notebook',
  'Collaborative Notebook': 'Collaborative Notebook',
  'Repo': 'Repo',
  'Repos': 'Repos',
  'Git Integration': 'Git Integration',
  'Workspace': 'Workspace',
  'Workspace Files': 'Workspace Files',
  'DBFS': 'DBFS',
  'Databricks File System': 'Databricks File System',
  'Magic Command': 'Magic Command',
  'Cell': 'Cell',
  'Command': 'Command',
  
  // Databricks Auto Loader関連
  'Auto Loader': 'Auto Loader',
  'Autoloader': 'Autoloader',
  'Structured Streaming': 'Structured Streaming',
  'Streaming': 'Streaming',
  'Stream Processing': 'Stream Processing',
  'Incremental Processing': 'Incremental Processing',
  'Cloud Files': 'Cloud Files',
  'Trigger': 'Trigger',
  
  // Databricks Feature Store関連
  'Feature Store': 'Feature Store',
  'Feature Engineering': 'Feature Engineering',
  'Feature Table': 'Feature Table',
  'Feature Lookup': 'Feature Lookup',
  'Feature Serving': 'Feature Serving',
  'Feature Pipeline': 'Feature Pipeline',
  'Online Store': 'Online Store',
  'Offline Store': 'Offline Store',
  
  // Databricks SQL関連
  'Databricks SQL': 'Databricks SQL',
  'SQL Analytics': 'SQL Analytics',
  'SQL Editor': 'SQL Editor',
  'Query': 'Query',
  'Dashboard': 'Dashboard',
  'Dashboards': 'Dashboards',
  'Alert': 'Alert',
  'Alerts': 'Alerts',
  'Query Profile': 'Query Profile',
  'Query History': 'Query History',
  'SQL Endpoint': 'SQL Endpoint',
  'SQL Persona': 'SQL Persona',
  
  // Databricks Machine Learning関連
  'Databricks Machine Learning': 'Databricks Machine Learning',
  'Databricks ML': 'Databricks ML',
  'ML Runtime': 'ML Runtime',
  'MLR': 'MLR',
  'AutoML': 'AutoML',
  'Automated ML': 'Automated ML',
  'Hyperopt': 'Hyperopt',
  'Hyperparameter Tuning': 'Hyperparameter Tuning',
  'Distributed Training': 'Distributed Training',
  'Model Serving': 'Model Serving',
  'Model Endpoint': 'Model Endpoint',
  'Inference': 'Inference',
  'Real-time Inference': 'Real-time Inference',
  'Batch Inference': 'Batch Inference',
  
  // Databricks Partner Connect関連
  'Partner Connect': 'Partner Connect',
  'Partner Solutions': 'Partner Solutions',
  'Marketplace': 'Marketplace',
  'Solution Accelerator': 'Solution Accelerator',
  'Solution Accelerators': 'Solution Accelerators',
  
  // Databricks セキュリティ関連
  'Access Control': 'Access Control',
  'RBAC': 'RBAC',
  'Table Access Control': 'Table Access Control',
  'Column-level Security': 'Column-level Security',
  'Row-level Security': 'Row-level Security',
  'Audit Log': 'Audit Log',
  'Audit Logs': 'Audit Logs',
  'Secret Scope': 'Secret Scope',
  'Secrets': 'Secrets',
  'Personal Access Token': 'Personal Access Token',
  'PAT': 'PAT',
  'Service Principal': 'Service Principal',
  'SCIM': 'SCIM',
  'SSO': 'SSO',
  'SAML': 'SAML',
  'Identity Provider': 'Identity Provider',
  'VPC': 'VPC',
  'Private Link': 'Private Link',
  'Customer-managed Keys': 'Customer-managed Keys',
  'CMK': 'CMK',
  
  // Databricks API・SDK関連
  'Databricks API': 'Databricks API',
  'REST API': 'REST API',
  'Databricks SDK': 'Databricks SDK',
  'Databricks CLI': 'Databricks CLI',
  'Terraform Provider': 'Terraform Provider',
  'Databricks Connect': 'Databricks Connect',
  'ODBC': 'ODBC',
  'JDBC': 'JDBC',
  'Connector': 'Connector',
  'Driver': 'Driver',
  
  // Databricks Runtime関連
  'Databricks Runtime': 'Databricks Runtime',
  'DBR': 'DBR',
  'Runtime Version': 'Runtime Version',
  'LTS Runtime': 'LTS Runtime',
  'Databricks Runtime for ML': 'Databricks Runtime for ML',
  'Photon Runtime': 'Photon Runtime',
  'GPU Runtime': 'GPU Runtime',
  'Genomics Runtime': 'Genomics Runtime',
  
  // Databricks クラウド・リージョン関連
  'Databricks on AWS': 'Databricks on AWS',
  'Databricks on Azure': 'Databricks on Azure',
  'Databricks on GCP': 'Databricks on GCP',
  'Multi-cloud': 'Multi-cloud',
  'Cross-cloud': 'Cross-cloud',
  'Region': 'Region',
  'Availability Zone': 'Availability Zone',
  
  // Databricks 料金・プラン関連
  'Standard Tier': 'Standard Tier',
  'Premium Tier': 'Premium Tier',
  'Enterprise Tier': 'Enterprise Tier',
  'DBU': 'DBU',
  'Databricks Unit': 'Databricks Unit',
  'Consumption-based': 'Consumption-based',
  
  // Apache Spark関連（Databricksで重要）
  'Apache Spark': 'Apache Spark',
  'Spark': 'Spark',
  'Spark SQL': 'Spark SQL',
  'Spark Streaming': 'Spark Streaming',
  'Structured Streaming': 'Structured Streaming',
  'MLlib': 'MLlib',
  'GraphX': 'GraphX',
  'PySpark': 'PySpark',
  'Spark DataFrame': 'Spark DataFrame',
  'RDD': 'RDD',
  'Catalyst': 'Catalyst',
  'Tungsten': 'Tungsten',
  'Adaptive Query Execution': 'Adaptive Query Execution',
  'AQE': 'AQE',
  'Dynamic Partition Pruning': 'Dynamic Partition Pruning',
  'Broadcast Join': 'Broadcast Join',
  'Shuffle': 'Shuffle',
  'Partitioning': 'Partitioning',
  'Bucketing': 'Bucketing',
  
  // 一般IT用語
  'API': 'API',
  'SDK': 'SDK',
  'REST API': 'REST API',
  'GraphQL': 'GraphQL',
  'JSON': 'JSON',
  'XML': 'XML',
  'YAML': 'YAML',
  'SQL': 'SQL',
  'NoSQL': 'NoSQL',
  'ETL': 'ETL',
  'ELT': 'ELT',
  'CI/CD': 'CI/CD',
  'DevOps': 'DevOps',
  'MLOps': 'MLOps',
  'DataOps': 'DataOps',
  'Docker': 'Docker',
  'Kubernetes': 'Kubernetes',
  'Terraform': 'Terraform',
  'GitHub': 'GitHub',
  'GitLab': 'GitLab',
  'Azure': 'Azure',
  'AWS': 'AWS',
  'GCP': 'GCP',
  'Google Cloud': 'Google Cloud',
  
  // データ関連
  'Big Data': 'Big Data',
  'Data Lake': 'Data Lake',
  'Data Warehouse': 'Data Warehouse',
  'Data Pipeline': 'Data Pipeline',
  'ETL Pipeline': 'ETL Pipeline',
  'Real-time': 'Real-time',
  'Streaming': 'Streaming',
  'Batch Processing': 'Batch Processing',
  'Stream Processing': 'Stream Processing',
  'OLTP': 'OLTP',
  'OLAP': 'OLAP',
  
  // AI/ML用語
  'Machine Learning': 'Machine Learning',
  'Deep Learning': 'Deep Learning',
  'Neural Network': 'Neural Network',
  'Transformer': 'Transformer',
  'LLM': 'LLM',
  'Large Language Model': 'Large Language Model',
  'NLP': 'NLP',
  'Computer Vision': 'Computer Vision',
  'Feature Engineering': 'Feature Engineering',
  'Model Training': 'Model Training',
  'Inference': 'Inference',
  'AutoML': 'AutoML',
  'Hyperparameter': 'Hyperparameter',
  'Cross-validation': 'Cross-validation',
  
  // プログラミング言語・フレームワーク
  'Python': 'Python',
  'Scala': 'Scala',
  'Java': 'Java',
  'R': 'R',
  'JavaScript': 'JavaScript',
  'TypeScript': 'TypeScript',
  'React': 'React',
  'Vue.js': 'Vue.js',
  'Angular': 'Angular',
  'Node.js': 'Node.js',
  'Express': 'Express',
  'Flask': 'Flask',
  'Django': 'Django',
  'Spring Boot': 'Spring Boot',
  'TensorFlow': 'TensorFlow',
  'PyTorch': 'PyTorch',
  'Scikit-learn': 'Scikit-learn',
  'Pandas': 'Pandas',
  'NumPy': 'NumPy',
  
  // その他のサービス・ツール
  'Slack': 'Slack',
  'Microsoft Teams': 'Microsoft Teams',
  'Zoom': 'Zoom',
  'Jira': 'Jira',
  'Confluence': 'Confluence',
  'Tableau': 'Tableau',
  'Power BI': 'Power BI',
  'Power BI + Copilot': 'Power BI + Copilot',
  'Looker': 'Looker',
  'Snowflake': 'Snowflake',
  'Redshift': 'Redshift',
  'BigQuery': 'BigQuery',
  'MongoDB': 'MongoDB',
  'PostgreSQL': 'PostgreSQL',
  'MySQL': 'MySQL',
  'Redis': 'Redis',
  'Elasticsearch': 'Elasticsearch',
  
  // 追加のDatabricksプラットフォーム関連用語
  'Databricks Data Intelligence Platform': 'Databricks Data Intelligence Platform',
  'Data Intelligence Platform': 'Data Intelligence Platform',
  'Mosaic AI': 'Mosaic AI',
  'Databricks IQ': 'Databricks IQ',
  'Predictive IQ': 'Predictive IQ',
  'Predictive optimization': 'Predictive optimization',
  'Predictive Optimization': 'Predictive Optimization',
  'Assistant': 'Assistant',
  'AI Assistant': 'AI Assistant',
  
  // データレイヤー・階層
  'bronze': 'bronze',
  'silver': 'silver',
  'gold': 'gold',
  'Bronze': 'Bronze',
  'Silver': 'Silver',
  'Gold': 'Gold',
  'bronze layer': 'bronze layer',
  'silver layer': 'silver layer',
  'gold layer': 'gold layer',
  'Bronze Layer': 'Bronze Layer',
  'Silver Layer': 'Silver Layer',
  'Gold Layer': 'Gold Layer',
  
  // AI/BI関連
  'AI/BI': 'AI/BI',
  'AI/BI Dashboards': 'AI/BI Dashboards',
  'AI/BI Genie': 'AI/BI Genie',
  'Gen AI': 'Gen AI',
  'Generative AI': 'Generative AI',
  'Classic ML': 'Classic ML',
  'Vector Search': 'Vector Search',
  'AI Gateway': 'AI Gateway',
  'AI Functions': 'AI Functions',
  'AI Services': 'AI Services',
  
  // データ処理・エンジニアリング
  'Batch & Streaming': 'Batch & Streaming',
  'Data Engineering & Processing': 'Data Engineering & Processing',
  'Automation & Orchestration': 'Automation & Orchestration',
  'Workflows': 'Workflows',
  'CI/CD tools': 'CI/CD tools',
  'DataOps': 'DataOps',
  'MLOps': 'MLOps',
  
  // データ管理・ガバナンス
  'Data and AI Governance': 'Data and AI Governance',
  'Federation': 'Federation',
  'Catalog & Lineage': 'Catalog & Lineage',
  'Access Control': 'Access Control',
  'Data & AI Assets': 'Data & AI Assets',
  'Lakehouse Monitoring': 'Lakehouse Monitoring',
  'Data Management': 'Data Management',
  'Collaboration': 'Collaboration',
  'Clean Rooms': 'Clean Rooms',
  
  // ストレージ・データソース
  'ADLS Gen 2': 'ADLS Gen 2',
  'Azure Data Lake Storage Gen 2': 'Azure Data Lake Storage Gen 2',
  'Iceberg': 'Iceberg',
  'Apache Iceberg': 'Apache Iceberg',
  'HMS': 'HMS',
  'Hive Metastore': 'Hive Metastore',
  'IoT Hub': 'IoT Hub',
  'Event Hub': 'Event Hub',
  'Azure Data Factory': 'Azure Data Factory',
  'Fabric Data Factory': 'Fabric Data Factory',
  'Synapse SQL Server': 'Synapse SQL Server',
  'SQL DB': 'SQL DB',
  'Cosmos DB': 'Cosmos DB',
  'Operational DBs': 'Operational DBs',
  
  // 分析・可視化・アプリケーション
  'Query / Process': 'Query / Process',
  'Query/Process': 'Query/Process',
  'Data Analysis': 'Data Analysis',
  'Data Consumer': 'Data Consumer',
  'Business/AI App': 'Business/AI App',
  'Business AI App': 'Business AI App',
  'Sharing Partner': 'Sharing Partner',
  'Applications': 'Applications',
  'Databricks Apps': 'Databricks Apps',
  
  // Azure・Microsoft関連
  'Entra ID': 'Entra ID',
  'Azure Entra ID': 'Azure Entra ID',
  'ID Provider': 'ID Provider',
  'Purview': 'Purview',
  'Azure Purview': 'Azure Purview',
  'Copilot': 'Copilot',
  'Microsoft Copilot': 'Microsoft Copilot',
  
  // 外部サービス・プラットフォーム
  'Hugging Face': 'Hugging Face',
  'OpenAI': 'OpenAI',
  'External Orchestrator': 'External Orchestrator',
  'Connectors & APIs': 'Connectors & APIs',
  '3rd party': '3rd party',
  'Third party': 'Third party',
  
  // データ形式・構造
  'semi-structured': 'semi-structured',
  'unstructured': 'unstructured',
  'structured': 'structured',
  'Semi-structured': 'Semi-structured',
  'Unstructured': 'Unstructured',
  'Structured': 'Structured',
  'Files/Logs': 'Files/Logs',
  'Sensors and IoT': 'Sensors and IoT',
  'Business Apps': 'Business Apps',
  'Market places': 'Market places',
  'Data Shares': 'Data Shares'
};

/**
 * グーグルスライドの翻訳機能を追加するメニューを作成する
 */
function onOpen() {
  SlidesApp.getUi()
    .createMenu('翻訳機能')
    .addItem('[全スライド] Google 翻訳（英語→日本語）', 'translateAllSlidesEnToJa')
    .addItem('[特定ページ] Google 翻訳（英語→日本語）', 'translateSpecificPageEnToJa')
    .addItem('[全スライド] Google 翻訳（英語→日本語・校正なし）', 'translateAllSlidesEnToJaNoProofread')
    .addItem('[特定ページ] Google 翻訳（英語→日本語・校正なし）', 'translateSpecificPageEnToJaNoProofread')
    .addSeparator()
    .addItem('[全スライド] Google 翻訳（日本語→英語）', 'translateAllSlidesJaToEn')
    .addItem('[特定ページ] Google 翻訳（日本語→英語）', 'translateSpecificPageJaToEn')
    .addSeparator()
    .addItem('[全スライド] 日本語を校正（基本）', 'proofreadAllSlidesJapanese')
    .addItem('[特定ページ] 日本語を校正（基本）', 'proofreadSpecificPageJapanese')
    .addSeparator()
    .addItem('[全スライド] Databricks AI校正', 'databricksProofreadAllSlides')
    .addItem('[特定ページ] Databricks AI校正', 'databricksProofreadSpecificPage')
    .addSeparator()
    .addItem('翻訳除外追加 Once', 'manageTranslationExclusions')
    .addItem('Databricks API設定', 'setupDatabricksAPI')
    .addToUi();
}

// 翻訳関数のエントリーポイント
function translateAllSlidesEnToJa() { translateAllSlides('en', 'ja'); }
function translateSpecificPageEnToJa() { translateSpecificPage('en', 'ja'); }
function translateAllSlidesEnToJaNoProofread() { translateAllSlides('en', 'ja', false); }
function translateSpecificPageEnToJaNoProofread() { translateSpecificPage('en', 'ja', false); }
function translateAllSlidesJaToEn() { translateAllSlides('ja', 'en'); }
function translateSpecificPageJaToEn() { translateSpecificPage('ja', 'en'); }
function proofreadAllSlidesJapanese() { proofreadAllSlides(); }
function proofreadSpecificPageJapanese() { proofreadSpecificPage(); }
/**
 * Databricks API設定を行う
 */
function setupDatabricksAPI() {
  const ui = SlidesApp.getUi();
  
  const workspaceResponse = ui.prompt(
    'Databricks API設定 (1/2)',
    'Databricksワークスペースの名前を入力してください:\n例: your-workspace\n※ https://your-workspace.cloud.databricks.com のyour-workspace部分',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (workspaceResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const workspace = workspaceResponse.getResponseText().trim();
  if (!workspace) {
    ui.alert('ワークスペース名が入力されていません。');
    return;
  }
  
  const tokenResponse = ui.prompt(
    'Databricks API設定 (2/2)',
    'Databricksアクセストークンを入力してください:\n\n※トークンは安全に暗号化されて保存されます',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (tokenResponse.getSelectedButton() === ui.Button.OK) {
    const token = tokenResponse.getResponseText().trim();
    if (token) {
      PropertiesService.getScriptProperties().setProperties({
        'DATABRICKS_WORKSPACE': workspace,
        'DATABRICKS_TOKEN': token
      });
      ui.alert('Databricks API設定が完了しました！');
    } else {
      ui.alert('アクセストークンが入力されていません。');
    }
  }
}

/**
 * 元の文字列の大文字・小文字パターンを新しい文字列に適用する
 * @param {string} original 元の文字列
 * @param {string} replacement 置換後の文字列
 * @returns {string} 大文字・小文字パターンを適用した文字列
 */
function preserveCase(original, replacement) {
  // 元の文字列が全て大文字の場合
  if (original === original.toUpperCase()) {
    return replacement.toUpperCase();
  }
  
  // 元の文字列が全て小文字の場合
  if (original === original.toLowerCase()) {
    return replacement.toLowerCase();
  }
  
  // 最初の文字のみ大文字の場合
  if (original[0] === original[0].toUpperCase() && original.slice(1) === original.slice(1).toLowerCase()) {
    return replacement[0].toUpperCase() + replacement.slice(1).toLowerCase();
  }
  
  // その他の場合は置換後の文字列をそのまま返す
  return replacement;
}

/**
 * テキスト内の翻訳除外用語を保護し、置換機能も提供する（書式保持版）
 * @param {string} text 処理するテキスト
 * @param {function} translationFunction 翻訳処理を行う関数
 * @returns {string} 除外用語を保護・置換した翻訳結果
 */
function protectAndTranslateWithReplacement(text, translationFunction) {
  if (!text || typeof text !== 'string' || text.trim() === '') {
    return text;
  }

  let protectedText = text;
  const placeholders = {};
  let placeholderIndex = 0;

  // カスタム除外辞書を取得
  const customExclusions = getCustomExclusions();
  const allExclusions = { ...TRANSLATION_EXCLUSIONS, ...customExclusions };

  // 除外用語をプレースホルダーに置換
  Object.entries(allExclusions).forEach(([key, value]) => {
    const escapedTerm = key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const regex = new RegExp(`\\b${escapedTerm}\\b`, 'gi');
    
    protectedText = protectedText.replace(regex, (match) => {
      const placeholder = `※${placeholderIndex}※`;
      
      // キーと値が異なる場合は置換、同じ場合は保護のみ
      if (key.toLowerCase() !== value.toLowerCase()) {
        // 置換機能：元の文字列の大文字・小文字パターンを保持
        placeholders[placeholder] = preserveCase(match, value);
      } else {
        // 保護機能：元の文字列をそのまま保持
        placeholders[placeholder] = match;
      }
      // 修正前: preserveCase関数で大文字小文字パターンを保持
// if (key.toLowerCase() !== value.toLowerCase()) {
//   placeholders[placeholder] = preserveCase(match, value);
// } else {
//  placeholders[placeholder] = match;
// }

// 修正後: 辞書の値をそのまま使用
// placeholders[placeholder] = value;
      
      placeholderIndex++;
      return placeholder;
    });
  });

  // プレースホルダーで保護されたテキストを翻訳
  let translatedText = translationFunction(protectedText);

  // プレースホルダーを元の用語または置換後の用語に戻す
  Object.keys(placeholders).forEach(placeholder => {
    const escapedPlaceholder = placeholder.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const regex = new RegExp(escapedPlaceholder, 'g');
    translatedText = translatedText.replace(regex, placeholders[placeholder]);
  });

  return translatedText;
}

/**
 * 文字レベルの書式情報を収集する
 * @param {TextRange} textRange テキスト範囲
 * @returns {Array} 書式情報の配列
 */
function collectCharacterFormats(textRange) {
  const formatInfo = [];
  const text = textRange.asString();
  
  for (let i = 0; i < text.length; i++) {
    try {
      const charRange = textRange.getRange(i, i + 1);
      const textStyle = charRange.getTextStyle();
      
      formatInfo.push({
        index: i,
        char: text[i],
        fontSize: textStyle.getFontSize(),
        fontFamily: textStyle.getFontFamily(),
        bold: textStyle.isBold(),
        italic: textStyle.isItalic(),
        underline: textStyle.hasUnderline(),
        strikethrough: textStyle.isStrikethrough(),
        foregroundColor: textStyle.getForegroundColor(),
        backgroundColor: textStyle.getBackgroundColor(),
        link: textStyle.hasLink() ? textStyle.getLink() : null
      });
    } catch (error) {
      formatInfo.push({
        index: i,
        char: text[i],
        fontSize: null,
        fontFamily: null,
        bold: null,
        italic: null,
        underline: null,
        strikethrough: null,
        foregroundColor: null,
        backgroundColor: null,
        link: null
      });
    }
  }
  
  return formatInfo;
}

/**
 * 段落レベルの書式情報を収集する
 * @param {Paragraph} paragraph 段落
 * @returns {Object} 段落の書式情報
 */
function collectParagraphFormat(paragraph) {
  try {
    const paragraphStyle = paragraph.getParagraphStyle();
    const range = paragraph.getRange();
    const textStyle = range.getTextStyle();
    
    return {
      // 段落スタイル
      alignment: paragraphStyle.getParagraphAlignment(),
      indentStart: paragraphStyle.getIndentStart(),
      indentEnd: paragraphStyle.getIndentEnd(),
      indentFirstLine: paragraphStyle.getIndentFirstLine(),
      spaceAbove: paragraphStyle.getSpaceAbove(),
      spaceBelow: paragraphStyle.getSpaceBelow(),
      lineSpacing: paragraphStyle.getLineSpacing(),
      direction: paragraphStyle.getParagraphDirection(),
      spacingMode: paragraphStyle.getSpacingMode(),
      // リスト情報
      listStyle: paragraph.getList() ? {
        listId: paragraph.getList().getListId(),
        nestingLevel: paragraph.getList().getNestingLevel()
      } : null,
      // デフォルトテキストスタイル
      defaultTextStyle: {
        fontSize: textStyle.getFontSize(),
        fontFamily: textStyle.getFontFamily(),
        bold: textStyle.isBold(),
        italic: textStyle.isItalic(),
        underline: textStyle.hasUnderline(),
        strikethrough: textStyle.isStrikethrough(),
        foregroundColor: textStyle.getForegroundColor(),
        backgroundColor: textStyle.getBackgroundColor()
      }
    };
  } catch (error) {
    console.log('段落書式収集エラー:', error);
    return {};
  }
}

/**
 * 段落の書式を復元する
 * @param {Paragraph} paragraph 段落
 * @param {Object} formatInfo 書式情報
 */
function restoreParagraphFormat(paragraph, formatInfo) {
  try {
    const paragraphStyle = paragraph.getParagraphStyle();
    const range = paragraph.getRange();
    const textStyle = range.getTextStyle();
    
    // 段落スタイルの復元
    if (formatInfo.alignment !== null && formatInfo.alignment !== undefined) {
      paragraphStyle.setParagraphAlignment(formatInfo.alignment);
    }
    if (formatInfo.indentStart !== null && formatInfo.indentStart !== undefined) {
      paragraphStyle.setIndentStart(formatInfo.indentStart);
    }
    if (formatInfo.indentEnd !== null && formatInfo.indentEnd !== undefined) {
      paragraphStyle.setIndentEnd(formatInfo.indentEnd);
    }
    if (formatInfo.indentFirstLine !== null && formatInfo.indentFirstLine !== undefined) {
      paragraphStyle.setIndentFirstLine(formatInfo.indentFirstLine);
    }
    if (formatInfo.spaceAbove !== null && formatInfo.spaceAbove !== undefined) {
      paragraphStyle.setSpaceAbove(formatInfo.spaceAbove);
    }
    if (formatInfo.spaceBelow !== null && formatInfo.spaceBelow !== undefined) {
      paragraphStyle.setSpaceBelow(formatInfo.spaceBelow);
    }
    if (formatInfo.lineSpacing !== null && formatInfo.lineSpacing !== undefined) {
      paragraphStyle.setLineSpacing(formatInfo.lineSpacing);
    }
    if (formatInfo.direction !== null && formatInfo.direction !== undefined) {
      paragraphStyle.setParagraphDirection(formatInfo.direction);
    }
    if (formatInfo.spacingMode !== null && formatInfo.spacingMode !== undefined) {
      paragraphStyle.setSpacingMode(formatInfo.spacingMode);
    }
    
    // デフォルトテキストスタイルの復元
    if (formatInfo.defaultTextStyle) {
      const defStyle = formatInfo.defaultTextStyle;
      if (defStyle.fontSize !== null && defStyle.fontSize !== undefined) {
        textStyle.setFontSize(defStyle.fontSize);
      }
      if (defStyle.fontFamily !== null && defStyle.fontFamily !== undefined) {
        textStyle.setFontFamily(defStyle.fontFamily);
      }
      if (defStyle.bold !== null && defStyle.bold !== undefined) {
        textStyle.setBold(defStyle.bold);
      }
      if (defStyle.italic !== null && defStyle.italic !== undefined) {
        textStyle.setItalic(defStyle.italic);
      }
      if (defStyle.underline !== null && defStyle.underline !== undefined) {
        textStyle.setUnderline(defStyle.underline);
      }
      if (defStyle.strikethrough !== null && defStyle.strikethrough !== undefined) {
        textStyle.setStrikethrough(defStyle.strikethrough);
      }
      if (defStyle.foregroundColor !== null && defStyle.foregroundColor !== undefined) {
        textStyle.setForegroundColor(defStyle.foregroundColor);
      }
      if (defStyle.backgroundColor !== null && defStyle.backgroundColor !== undefined) {
        textStyle.setBackgroundColor(defStyle.backgroundColor);
      }
    }
    
  } catch (error) {
    console.log('段落書式復元エラー:', error);
  }
}

/**
 * 文字レベルの書式を復元する（高度な復元）
 * @param {TextRange} textRange 新しいテキスト範囲
 * @param {Array} originalFormats 元の書式情報
 * @param {string} originalText 元のテキスト
 * @param {string} newText 新しいテキスト
 */
function restoreCharacterFormats(textRange, originalFormats, originalText, newText) {
  try {
    // 単語レベルでのマッピングを試行
    const originalWords = originalText.split(/(\s+)/);
    const newWords = newText.split(/(\s+)/);
    
    let originalCharIndex = 0;
    let newCharIndex = 0;
    
    for (let i = 0; i < Math.min(originalWords.length, newWords.length); i++) {
      const originalWord = originalWords[i];
      const newWord = newWords[i];
      
      if (originalWord.trim() === '' || newWord.trim() === '') {
        // スペースの場合はそのまま進む
        originalCharIndex += originalWord.length;
        newCharIndex += newWord.length;
        continue;
      }
      
      // 単語の書式を適用
      try {
        if (newCharIndex < newText.length && newCharIndex + newWord.length <= newText.length) {
          const wordRange = textRange.getRange(newCharIndex, newCharIndex + newWord.length);
          
          // 元の単語の最初の文字の書式を取得
          if (originalCharIndex < originalFormats.length) {
            const format = originalFormats[originalCharIndex];
            const wordStyle = wordRange.getTextStyle();
            
            if (format.fontSize !== null) wordStyle.setFontSize(format.fontSize);
            if (format.fontFamily !== null) wordStyle.setFontFamily(format.fontFamily);
            if (format.bold !== null) wordStyle.setBold(format.bold);
            if (format.italic !== null) wordStyle.setItalic(format.italic);
            if (format.underline !== null) wordStyle.setUnderline(format.underline);
            if (format.strikethrough !== null) wordStyle.setStrikethrough(format.strikethrough);
            if (format.foregroundColor !== null) wordStyle.setForegroundColor(format.foregroundColor);
            if (format.backgroundColor !== null) wordStyle.setBackgroundColor(format.backgroundColor);
            if (format.link !== null) wordStyle.setLinkUrl(format.link.getUrl());
          }
        }
      } catch (error) {
        console.log(`単語 "${newWord}" の書式復元エラー:`, error);
      }
      
      originalCharIndex += originalWord.length;
      newCharIndex += newWord.length;
    }
    
  } catch (error) {
    console.log('文字書式復元エラー:', error);
  }
}

/**
 * 保護機能付きGoogle翻訳（置換機能対応版）
 */
function translateWithGoogleListAware(text, src, tgt) {
  if (!text || text.trim() === '' || text.match(/^[\n\r\s•\-\*\d\.\)\(\s]*$/)) {
    return text;
  }
/**
 * リスト項目の改行パターンを識別する
 * @param {string} text テキスト
 * @returns {boolean} リスト関連の改行を含む場合true
 */
function containsListPatterns(text) {
  // リストマーカーのパターン
  const listPatterns = [
    /^[\s]*[•\-\*]\s+/,           // 箇条書き (•, -, *)
    /^[\s]*\d+[\.\)]\s+/,        // 番号付きリスト (1., 1), 2., 2))
    /^[\s]*[a-zA-Z][\.\)]\s+/,   // アルファベットリスト (a., a), A., A))
    /^[\s]*[ivxIVX]+[\.\)]\s+/,  // ローマ数字リスト (i., ii., IV., V))
    /[\n\r][\s]*[•\-\*]\s+/,     // 改行後の箇条書き
    /[\n\r][\s]*\d+[\.\)]\s+/,   // 改行後の番号付きリスト
    /[\n\r][\s]*[a-zA-Z][\.\)]\s+/ // 改行後のアルファベットリスト
  ];
  
  return listPatterns.some(pattern => pattern.test(text));
}




  // 除外用語を保護・置換して翻訳
  const translatedText = protectAndTranslateWithReplacement(text, (protectedText) => {
    return LanguageApp.translate(protectedText, src, tgt);
  });

  let finalResult = translatedText;
  if (tgt === 'ja') {
    finalResult = basicJapaneseProofread(translatedText);
  }
  
  Utilities.sleep(20);
  return finalResult;
}

/**
 * 基本的な日本語校正を行う
 */
function basicJapaneseProofread(text) {
  if (!text || typeof text !== 'string') return text;
  
  let corrected = text;
  
  // マーカーが残っている場合は削除（念のため）
  corrected = corrected.replace(/\u2060/g, '');
  
  // 句点処理（既存のロジックは残す）
  // ただし、マーカー方式を使う場合はこの処理は不要かもしれない
  corrected = corrected.replace(/^(?:\s*[-・•\d\)．\.]+?\s*)?(.+?)。$/gm, (m, body) => {
    return /[、。！？]/.test(body) || body.length > 25 ? m : body;
  });
  
  // ... 残りの処理は同じ
  
  return corrected;
}



/**
 * 箇条書き行かどうかを判定（厳密版）
 */
function isBulletLineStrict(line) {
  if (!isListItemNew(line)) return false;
  
  // マーカーを除去したコンテンツ部分を取得
  const itemMatch = findListItemMatch(line);
  if (!itemMatch) return false;
  
  const content = itemMatch.content.trim();
  
  // 除外パターン
  if (shouldSkipProcessing(content)) return false;
  
  // 箇条書きっぽさの判定
  const checks = {
    shortLength: content.length <= 30,           // 短い
    fewPunctuation: (content.match(/[、。]/g) || []).length <= 1, // 句読点が少ない
    notSentence: !/です。?$|ます。?$|した。?$/.test(content), // 文末表現がない
    hasContent: content.length >= 2              // 最低限の内容がある
  };
  
  // 3つ以上の条件を満たせば箇条書きとみなす
  const passCount = Object.values(checks).filter(v => v).length;
  return passCount >= 3;
}

/**
 * 高度な日本語校正を行う
 */
function advancedJapaneseProofread(text) {
  if (!text || typeof text !== 'string') return text;
  
  let corrected = basicJapaneseProofread(text);
  
  const presentationReplacements = [
    ['このスライドでは', 'ここでは'],
    ['以下に示します', '以下の通りです'],
    ['図表を示しています', '図表をご覧ください'],
    ['データを表示しています', 'データは以下の通りです'],
    ['結果を示しています', '結果は以下の通りです'],
    ['まとめると', 'まとめ'],
    ['結論として', '結論'],
    ['について考えてみましょう', 'について'],
    ['検討してみましょう', '検討します']
  ];
  
  presentationReplacements.forEach(([from, to]) => {
    const regex = new RegExp(from, 'g');
    corrected = corrected.replace(regex, to);
  });
  
  const businessTermReplacements = [
    ['活用', '利用'],
    ['実施', '実行'],
    ['推進', '促進'],
    ['課題解決', '問題解決'],
    ['最適化', '改善']
  ];
  
  businessTermReplacements.forEach(([from, to]) => {
    const regex = new RegExp(from, 'g');
    corrected = corrected.replace(regex, to);
  });
  
  return corrected;
}

/**
 * 書式保持版：シェイプのテキストを翻訳する
 */
 function translateShapeWithFormatPreservation(shape, slide, src, tgt) {
  const textRange = shape.getText();
  const fullText = textRange.asString();
  
  if (fullText.trim() === '') return;
  
  try {
    const paragraphs = textRange.getParagraphs();
    
    // 各段落を個別に処理
    for (let i = 0; i < paragraphs.length; i++) {
      const paragraph = paragraphs[i];
      const paragraphRange = paragraph.getRange();
      const originalText = paragraphRange.asString();
      
      if (originalText.trim() === '') continue;
      
      // 段落の書式情報を保存
      const paragraphFormat = collectParagraphFormat(paragraph);
      const characterFormats = collectCharacterFormats(paragraphRange);
      
      // 翻訳
      const translatedText = translateWithGoogle(originalText, src, tgt);
      
      // テキストを置き換え
      paragraphRange.setText(translatedText);
      
      // 書式を復元
      restoreParagraphFormat(paragraph, paragraphFormat);
      restoreCharacterFormats(paragraphRange, characterFormats, originalText, translatedText);
    }
    
  } catch (error) {
    console.log('シェイプ翻訳エラー:', error);
    // フォールバック：通常の翻訳
    const translatedText = translateWithGoogle(fullText, src, tgt);
    textRange.setText(translatedText);
  }
}

/**
 * 書式保持版：シェイプのテキストを校正する
 */
function proofreadShapeWithFormatPreservation(shape, isAdvanced = false) {
  const textRange = shape.getText();
  const fullText = textRange.asString();
  
  if (fullText.trim() === '') return;
  
  try {
    const paragraphs = textRange.getParagraphs();
    
    for (let i = 0; i < paragraphs.length; i++) {
      const paragraph = paragraphs[i];
      const paragraphRange = paragraph.getRange();
      const originalText = paragraphRange.asString();
      
      if (originalText.trim() === '') continue;
      
      const paragraphFormat = collectParagraphFormat(paragraph);
      const characterFormats = collectCharacterFormats(paragraphRange);
      
      // 校正
      const proofreadText = isAdvanced ? 
        advancedJapaneseProofread(originalText) : 
        basicJapaneseProofread(originalText);
      
      paragraphRange.setText(proofreadText);
      
      // 書式を復元
      restoreParagraphFormat(paragraph, paragraphFormat);
      restoreCharacterFormats(paragraphRange, characterFormats, originalText, proofreadText);
    }
    
  } catch (error) {
    console.log('シェイプ校正エラー:', error);
    const proofreadText = isAdvanced ? 
      advancedJapaneseProofread(fullText) : 
      basicJapaneseProofread(fullText);
    textRange.setText(proofreadText);
  }
}

/**
 * Databricks APIを使用した校正（書式保持版）
 */
function databricksProofreadShapeWithFormatPreservation(shape) {
  const textRange = shape.getText();
  const fullText = textRange.asString();
  
  if (fullText.trim() === '') return;
  
  try {
    const paragraphs = textRange.getParagraphs();
    
    for (let i = 0; i < paragraphs.length; i++) {
      const paragraph = paragraphs[i];
      const paragraphRange = paragraph.getRange();
      const originalText = paragraphRange.asString();
      
      if (originalText.trim() === '') continue;
      
      const paragraphFormat = collectParagraphFormat(paragraph);
      const characterFormats = collectCharacterFormats(paragraphRange);
      
      // Databricks校正
      const proofreadText = proofreadWithDatabricks(originalText);
      
      paragraphRange.setText(proofreadText);
      
      // 書式を復元
      restoreParagraphFormat(paragraph, paragraphFormat);
      restoreCharacterFormats(paragraphRange, characterFormats, originalText, proofreadText);
    }
    
  } catch (error) {
    console.log('Databricks校正エラー:', error);
    const proofreadText = proofreadWithDatabricks(fullText);
    textRange.setText(proofreadText);
  }
}

/**
 * スライドのテキストを翻訳する（テーブル対応版）
 */
function translateSlide(slide, src, tgt, enableProofread = true) {
  console.log('=== スライド翻訳開始 ===');
  
  // スピーカーノートを翻訳
  try {
    translateSpeakerNotesWithFormatPreservation(slide, src, tgt, enableProofread);
  } catch (e) {
    console.log('スピーカーノート翻訳エラー:', e.message);
  }

  const pageElements = slide.getPageElements();
  console.log(`ページ要素数: ${pageElements.length}`);

  // シェイプとグループを翻訳
  for (let i = 0; i < pageElements.length; i++) {
    const pageElement = pageElements[i];
    const elementType = pageElement.getPageElementType();
    
    try {
      if (elementType === SlidesApp.PageElementType.SHAPE) {
        translateShapeWithListAwareFormatPreservationNew(pageElement.asShape(), slide, src, tgt, enableProofread);
      } else if (elementType === SlidesApp.PageElementType.GROUP) {
        translateGroupWithFormatPreservation(pageElement.asGroup(), slide, src, tgt, enableProofread);
      } else if (elementType === SlidesApp.PageElementType.TABLE) {
        console.log(`テーブル ${i} を検出（後で処理）`);
      }
    } catch (error) {
      console.log(`要素 ${i} (${elementType}) の翻訳エラー: ${error.message}`);
    }
  }
  
  // テーブルを翻訳（新版）
  try {
    translateAllTablesInSlide(slide, src, tgt, enableProofread);
  } catch (e) {
    console.log('テーブル翻訳エラー:', e.message);
  }
  
  console.log('=== スライド翻訳完了 ===');
  Utilities.sleep(20);
}

/**
 * グループ内のシェイプのテキストを翻訳する（書式保持版）
 */
function translateGroupWithFormatPreservation(group, slide, src, tgt, enableProofread = true) {
  const childElements = group.getChildren();

  for (let childElement of childElements) {
    if (childElement.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      translateShapeWithListAwareFormatPreservation(childElement.asShape(), slide, src, tgt, enableProofread);
    } else if (childElement.getPageElementType() === SlidesApp.PageElementType.GROUP) {
      translateGroupWithFormatPreservation(childElement.asGroup(), slide, src, tgt, enableProofread);
    }
  }
}

/**
 * グループ内のシェイプのテキストを校正する（書式保持版）
 */
function proofreadGroupWithFormatPreservation(group, isAdvanced = false) {
  const childElements = group.getChildren();

  for (let childElement of childElements) {
    if (childElement.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      proofreadShapeWithFormatPreservation(childElement.asShape(), isAdvanced);
    } else if (childElement.getPageElementType() === SlidesApp.PageElementType.GROUP) {
      proofreadGroupWithFormatPreservation(childElement.asGroup(), isAdvanced);
    }
  }
}

/**
 * Databricks APIを使用してグループ内のシェイプのテキストを校正する（書式保持版）
 */
function databricksProofreadGroupWithFormatPreservation(group) {
  const childElements = group.getChildren();

  for (let childElement of childElements) {
    if (childElement.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      databricksProofreadShapeWithListAwareFormatPreservation(childElement.asShape());
    } else if (childElement.getPageElementType() === SlidesApp.PageElementType.GROUP) {
      databricksProofreadGroupWithFormatPreservation(childElement.asGroup());
    }
  }
}

/**
 * スライド内の全テーブルを翻訳（新関数）
 */
function translateAllTablesInSlide(slide, src, tgt, enableProofread) {
  console.log('--- テーブル翻訳開始 ---');
  
  const pageElements = slide.getPageElements();
  let tableCount = 0;
  let totalCells = 0;
  
  for (let i = 0; i < pageElements.length; i++) {
    const element = pageElements[i];
    
    try {
      if (element.getPageElementType() === SlidesApp.PageElementType.TABLE) {
        tableCount++;
        console.log(`\nテーブル ${tableCount} (要素 ${i})`);
        
        const table = element.asTable();
        const cellCount = translateSingleTable(table, src, tgt, enableProofread);
        totalCells += cellCount;
        
        console.log(`✓ テーブル ${tableCount} 完了: ${cellCount}セル翻訳`);
      }
    } catch (error) {
      console.log(`テーブル処理エラー: ${error.message}`);
    }
  }
  
  if (tableCount === 0) {
    console.log('表が見つかりませんでした');
  } else {
    console.log(`--- テーブル翻訳完了: ${tableCount}個、${totalCells}セル ---`);
  }
}



/**
 * 1つのテーブルを翻訳
 * @returns {number} 翻訳したセル数
 */
function translateSingleTable(table, src, tgt, enableProofread) {
  const numRows = table.getNumRows();
  const numColumns = table.getNumColumns();
  console.log(`  サイズ: ${numRows}行 × ${numColumns}列`);
  
  const processedCells = new Set();
  let translatedCount = 0;
  
  for (let row = 0; row < numRows; row++) {
    for (let col = 0; col < numColumns; col++) {
      const cellKey = `${row},${col}`;
      
      // すでに処理済みならスキップ
      if (processedCells.has(cellKey)) {
        console.log(`  [${row},${col}] スキップ（処理済み）`);
        continue;
      }
      
      try {
        // ★重要：セルを取得する前にnullチェック
        const cell = table.getCell(row, col);
        
        if (!cell) {
          console.log(`  [${row},${col}] セルが存在しません（マージセルの一部）`);
          processedCells.add(cellKey);
          continue;
        }
        
        // このセルを処理済みとしてマーク
        processedCells.add(cellKey);
        
        // マージセル情報を取得して、マージ範囲を記録
        try {
          const rowSpan = cell.getRowSpan();
          const colSpan = cell.getColumnSpan();
          
          if (rowSpan && colSpan) {
            console.log(`  [${row},${col}] マージセル: ${rowSpan}行 × ${colSpan}列`);
            
            // マージ範囲内のすべてのセルを処理済みとしてマーク
            for (let r = row; r < row + rowSpan && r < numRows; r++) {
              for (let c = col; c < col + colSpan && c < numColumns; c++) {
                if (r !== row || c !== col) {  // 元のセルは既に追加済み
                  processedCells.add(`${r},${c}`);
                }
              }
            }
          }
        } catch (spanError) {
          // getRowSpan/getColumnSpanが使えない場合は無視
          console.log(`  [${row},${col}] スパン情報取得エラー: ${spanError.message}`);
        }
        
        // セルの翻訳
        const wasTranslated = translateTableCell(cell, src, tgt, enableProofread, row, col);
        if (wasTranslated) {
          translatedCount++;
        }
        
      } catch (error) {
        console.log(`  [${row},${col}] エラー: ${error.message}`);
        processedCells.add(cellKey);  // エラーでも処理済みとしてマーク
      }
    }
  }
  
  console.log(`  翻訳完了: ${translatedCount}/${numRows * numColumns}セル`);
  return translatedCount;
}

/**
 * 全てのスライドを翻訳する（テーブル統合版）
 */
function translateAllSlides(src, tgt, enableProofread = true) {
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  
  console.log(`=== 全スライド翻訳開始 (${slides.length}枚) ===`);

  for (let i = 0; i < slides.length; i++) {
    console.log(`\n--- スライド ${i + 1}/${slides.length} ---`);
    try {
      translateSlide(slides[i], src, tgt, enableProofread);
    } catch (error) {
      console.log(`スライド ${i + 1} でエラー: ${error.message}`);
    }
  }
  
  console.log('\n=== 全スライド翻訳完了 ===');
  SlidesApp.getUi().alert(`${slides.length}枚のスライドの翻訳が完了しました。`);
  Utilities.sleep(20);
}

/**
 * 全てのスライドを校正する（書式保持版）
 */
function proofreadAllSlides() {
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();

  for (let slide of slides) {
    proofreadSlideWithFormatPreservation(slide);
  }
}

/**
 * スライドのテキストを校正する（書式保持版）
 */
function proofreadSlideWithFormatPreservation(slide) {
  proofreadSpeakerNotesWithFormatPreservation(slide);

  const pageElements = slide.getPageElements();

  for (let pageElement of pageElements) {
    if (pageElement.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      proofreadShapeWithFormatPreservation(pageElement.asShape(), true);
    } else if (pageElement.getPageElementType() === SlidesApp.PageElementType.GROUP) {
      proofreadGroupWithFormatPreservation(pageElement.asGroup(), true);
    }
  }
}

/**
 * 特定のスライドページを翻訳する（テーブル統合版）
 */
function translateSpecificPage(src, tgt, enableProofread = true) {
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  const pageNumber = getPageNumberFromUser();

  if (pageNumber === null) {
    return;
  }

  if (pageNumber < 1 || pageNumber > slides.length) {
    SlidesApp.getUi().alert(`ページ番号が無効です。1から${slides.length}の間で指定してください。`);
    return;
  }

  console.log(`\n--- スライド ${pageNumber} の翻訳 ---`);
  const slide = slides[pageNumber - 1];
  translateSlide(slide, src, tgt, enableProofread);
  
  SlidesApp.getUi().alert(`スライド${pageNumber}の翻訳が完了しました。`);
  Utilities.sleep(20);
}



/**
 * 特定のスライドページを校正する（書式保持版）
 */
function proofreadSpecificPage() {
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  const pageNumber = getPageNumberFromUser();

  if (pageNumber === null) {
    return;
  }

  const slide = slides[pageNumber - 1];
  proofreadSlideWithFormatPreservation(slide);
}

/**
 * スピーカーノートを翻訳する（書式保持版）
 */
function translateSpeakerNotesWithFormatPreservation(slide, src, tgt, enableProofread = true) {
  try {
    const speakerNotesShape = slide.getNotesPage().getSpeakerNotesShape();
    const textRange = speakerNotesShape.getText();
    const fullText = textRange.asString();
    
    if (fullText.trim() === '') return;
    
    const paragraphs = textRange.getParagraphs();
    
    for (let i = 0; i < paragraphs.length; i++) {
      const paragraph = paragraphs[i];
      const paragraphRange = paragraph.getRange();
      const originalText = paragraphRange.asString();
      
      if (originalText.trim() === '') continue;
      
      const paragraphFormat = collectParagraphFormat(paragraph);
      const characterFormats = collectCharacterFormats(paragraphRange);
      
      const translatedText = translateWithGoogleListAwareNew(originalText, src, tgt, enableProofread);
      
      paragraphRange.setText(translatedText);
      
      restoreParagraphFormat(paragraph, paragraphFormat);
      restoreCharacterFormats(paragraphRange, characterFormats, originalText, translatedText);
    }
    
  } catch (error) {
    console.log('スピーカーノート翻訳エラー:', error);
  }
}

/**
 * スピーカーノートを校正する（書式保持版）
 */
function proofreadSpeakerNotesWithFormatPreservation(slide, isAdvanced = true) {
  try {
    const speakerNotesShape = slide.getNotesPage().getSpeakerNotesShape();
    const textRange = speakerNotesShape.getText();
    const fullText = textRange.asString();
    
    if (fullText.trim() === '') return;
    
    const paragraphs = textRange.getParagraphs();
    
    for (let i = 0; i < paragraphs.length; i++) {
      const paragraph = paragraphs[i];
      const paragraphRange = paragraph.getRange();
      const originalText = paragraphRange.asString();
      
      if (originalText.trim() === '') continue;
      
      const paragraphFormat = collectParagraphFormat(paragraph);
      const characterFormats = collectCharacterFormats(paragraphRange);
      
      const proofreadText = isAdvanced ? 
        advancedJapaneseProofread(originalText) : 
        basicJapaneseProofread(originalText);
      
      paragraphRange.setText(proofreadText);
      
      restoreParagraphFormat(paragraph, paragraphFormat);
      restoreCharacterFormats(paragraphRange, characterFormats, originalText, proofreadText);
    }
    
  } catch (error) {
    console.log('スピーカーノート校正エラー:', error);
  }
}

/**
 * ユーザーからスライドのページ番号を取得する
 */
function getPageNumberFromUser() {
  const ui = SlidesApp.getUi();
  const response = ui.prompt(
    '翻訳こんにゃく（書式保持・置換版）',
    '翻訳したいスライドのページ番号を入力してください。',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.CANCEL) {
    console.log('キャンセルが押されたため、処理を中断しました。');
    return null;
  } else if (response.getSelectedButton() === ui.Button.CLOSE) {
    console.log('閉じるボタンが押されました。');
    return null;
  }

  const pageNumber = parseInt(response.getResponseText(), 10);
  console.log(`ページ番号は、${pageNumber} です。`);
  return pageNumber;
}

/**
 * Databricks APIを使用して全てのスライドを校正する（書式保持版）
 */
function databricksProofreadAllSlides() {
  if (!DATABRICKS_TOKEN || !DATABRICKS_WORKSPACE) {
    SlidesApp.getUi().alert('Databricks API設定が不完全です。「Databricks API設定」から設定してください。');
    return;
  }
  
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();

  for (let i = 0; i < slides.length; i++) {
    const slide = slides[i];
    databricksProofreadSlideWithFormatPreservation(slide);
    
    // API制限を考慮してスリープ
    Utilities.sleep(20);
  }
  
  SlidesApp.getUi().alert(`${slides.length}枚のスライドの校正が完了しました。`);
}

/**
 * Databricks APIを使用して特定のスライドページを校正する（書式保持版）
 */
function databricksProofreadSpecificPage() {
  if (!DATABRICKS_TOKEN || !DATABRICKS_WORKSPACE) {
    SlidesApp.getUi().alert('Databricks API設定が不完全です。「Databricks API設定」から設定してください。');
    return;
  }
  
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  const pageNumber = getPageNumberFromUser();

  if (pageNumber === null) {
    return;
  }

  const slide = slides[pageNumber - 1];
  databricksProofreadSlideWithFormatPreservation(slide);
  
  SlidesApp.getUi().alert(`スライド${pageNumber}の校正が完了しました。`);
}

/**
 * Databricks APIを使用してテキストを校正する（保護機能付き）
 */
/**
 * Databricks APIを使用してテキストを校正する（保護機能付き・リスト対応版）
 */
function proofreadWithDatabricks(text) {
  if (!text || typeof text !== 'string' || text.trim() === '') {
    return text;
  }

  // リスト構造を保持しながら処理
  return processTextWithListPreservationNew(text, function(textToProofread) {
    return protectAndTranslateWithReplacement(textToProofread, function(protectedText) {
      const customExclusions = getCustomExclusions();
      const allExclusions = { ...TRANSLATION_EXCLUSIONS, ...customExclusions };
      
      const exclusionList = Object.keys(allExclusions).slice(0, 20).join(', ');

      const prompt = `以下の日本語テキストをプレゼンテーション用として自然で読みやすく校正してください。

重要な注意事項：
以下の技術用語・サービス名は絶対に変更せず、そのまま保持してください：
${exclusionList}

校正ガイドライン：
1. 不自然な敬語や冗長な表現を簡潔にする
2. プレゼンテーション向けの分かりやすい表現にする
3. 句読点を適切に使用する（、。を基本）
4. カタカナの長音符は「ー」に統一
5. 英語と日本語の間に適切なスペースを入れる
6. 文章の意味は変えずに、より自然な日本語にする
7. 上記の技術用語・サービス名は変更厳禁
8. 校正後のテキストのみを返してください（説明文は不要）

元のテキスト：${protectedText}

校正後：`;

      const apiUrl = `https://${DATABRICKS_WORKSPACE}.cloud.databricks.com/serving-endpoints/${DATABRICKS_MODEL_ENDPOINT}/invocations`;
      
      const payload = {
        inputs: {
          messages: [
            {
              role: "system",
              content: "あなたは日本語の校正専門家です。技術用語やサービス名は絶対に変更せず、プレゼンテーション用の文書を自然で読みやすい日本語に校正してください。校正結果のみを返し、説明は不要です。"
            },
            {
              role: "user",
              content: prompt
            }
          ],
          max_tokens: 1000,
          temperature: 0.3,
          top_p: 0.9
        }
      };

      const options = {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${DATABRICKS_TOKEN}`,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      try {
        const response = UrlFetchApp.fetch(apiUrl, options);
        const responseCode = response.getResponseCode();
        
        if (responseCode !== 200) {
          console.log('Databricks APIエラー (ステータス:', responseCode, ')');
          return advancedJapaneseProofread(protectedText);
        }
        
        const responseText = response.getContentText();
        const responseJson = JSON.parse(responseText);
        
        if (responseJson.choices && responseJson.choices.length > 0) {
          let correctedText = responseJson.choices[0].message.content;
          
          // 余分な説明文を除去
          correctedText = correctedText.replace(/^.*?校正後[：:]\s*/m, '');
          correctedText = correctedText.replace(/^.*?結果[：:]\s*/m, '');
          correctedText = correctedText.replace(/^校正された文章[：:]\s*/m, '');
          correctedText = correctedText.trim();
          
          // 複数行ある場合は最初の行のみを使用
          const lines = correctedText.split('\n');
          if (lines.length > 1 && lines[0].length > 10) {
            correctedText = lines[0];
          }
          
          return correctedText || protectedText;
        } else if (responseJson.predictions && responseJson.predictions.length > 0) {
          let correctedText = responseJson.predictions[0];
          correctedText = correctedText.replace(/^.*?校正後[：:]\s*/m, '');
          correctedText = correctedText.trim();
          return correctedText || protectedText;
        } else {
          console.log('Databricks API からの応答が不正です:', responseJson);
          return advancedJapaneseProofread(protectedText);
        }
      } catch (error) {
        console.log('Databricks API エラー:', error);
        return advancedJapaneseProofread(protectedText);
      }
    });
  });
}

/**
 * Databricks APIを使用してスライドのテキストを校正する（書式保持版）
 */
function databricksProofreadSlideWithFormatPreservation(slide) {
  databricksProofreadSpeakerNotesWithFormatPreservation(slide);

  const pageElements = slide.getPageElements();

  for (let pageElement of pageElements) {
    if (pageElement.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      databricksProofreadShapeWithListAwareFormatPreservation(pageElement.asShape());
    } else if (pageElement.getPageElementType() === SlidesApp.PageElementType.GROUP) {
      databricksProofreadGroupWithFormatPreservation(pageElement.asGroup());
    }
  }
}

/**
 * Databricks APIを使用してスピーカーノートを校正する（書式保持版）
 */
function databricksProofreadSpeakerNotesWithFormatPreservation(slide) {
  try {
    const speakerNotesShape = slide.getNotesPage().getSpeakerNotesShape();
    const textRange = speakerNotesShape.getText();
    const fullText = textRange.asString();
    
    if (fullText.trim() === '') return;
    
    const paragraphs = textRange.getParagraphs();
    
    for (let i = 0; i < paragraphs.length; i++) {
      const paragraph = paragraphs[i];
      const paragraphRange = paragraph.getRange();
      const originalText = paragraphRange.asString();
      
      if (originalText.trim() === '') continue;
      
      const paragraphFormat = collectParagraphFormat(paragraph);
      const characterFormats = collectCharacterFormats(paragraphRange);
      
      const proofreadText = proofreadWithDatabricks(originalText);
      
      paragraphRange.setText(proofreadText);
      
      restoreParagraphFormat(paragraph, paragraphFormat);
      restoreCharacterFormats(paragraphRange, characterFormats, originalText, proofreadText);
    }
    
  } catch (error) {
    console.log('Databricksスピーカーノート校正エラー:', error);
  }
}

/**
 * Databricks APIを使用してシェイプのテキストを校正する（リスト対応・書式保持版）
 */
function databricksProofreadShapeWithListAwareFormatPreservation(shape) {
  const textRange = shape.getText();
  const fullText = textRange.asString();
  
  if (fullText.trim() === '') return;
  
  try {
    const paragraphs = textRange.getParagraphs();
    const listInfo = analyzeListStructure(fullText);
    
    if (listInfo.isListText) {
      console.log('リスト構造を検出（校正）: ' + listInfo.listType + ', ' + listInfo.totalItems + '項目');
      
      const proofreadText = proofreadWithDatabricks(fullText);
      
      textRange.setText(proofreadText);
      
      // 基本的な書式復元
      try {
        if (paragraphs.length > 0) {
          const defaultFormat = collectParagraphFormat(paragraphs[0]);
          const newParagraphs = textRange.getParagraphs();
          for (let i = 0; i < newParagraphs.length; i++) {
            restoreParagraphFormat(newParagraphs[i], defaultFormat);
          }
        }
      } catch (error) {
        console.log('リスト書式復元エラー（校正）:', error);
      }
      
    } else {
      // 通常のテキスト処理
      for (let i = 0; i < paragraphs.length; i++) {
        const paragraph = paragraphs[i];
        const paragraphRange = paragraph.getRange();
        const originalText = paragraphRange.asString();
        
        if (originalText.trim() === '') continue;
        
        const paragraphFormat = collectParagraphFormat(paragraph);
        const characterFormats = collectCharacterFormats(paragraphRange);
        
        const proofreadText = proofreadWithDatabricks(originalText);
        
        paragraphRange.setText(proofreadText);
        
        restoreParagraphFormat(paragraph, paragraphFormat);
        restoreCharacterFormats(paragraphRange, characterFormats, originalText, proofreadText);
      }
    }
    
  } catch (error) {
    console.log('Databricksシェイプ校正エラー:', error);
    const proofreadText = proofreadWithDatabricks(fullText);
    textRange.setText(proofreadText);
  }
}
/**
 * テキストのサイズを減らす
 * @param {TextRange} textRange テキスト範囲
 */
function reduceTextSize(textRange) {
  if (textRange.asString().trim() !== "") {
    const textStyle = textRange.getTextStyle();
    if (textStyle.getFontSize() !== null) {
      // フォントサイズを90%に縮小
      textStyle.setFontSize(Math.round(textStyle.getFontSize() * 0.9));
    }
  }
}

/**
 * 1つのセルを翻訳
 * @returns {boolean} 翻訳したかどうか
 */
function translateTableCell(cell, src, tgt, enableProofread, row, col) {
  try {
    const textRange = cell.getText();
    const originalText = textRange.asString();
    
    // 空セルはスキップ
    if (!originalText || originalText.trim() === '') {
      return false;
    }
    
    console.log(`  [${row},${col}]: "${originalText.substring(0, 30)}..."`);
    
    // ★重要：セルの塗りつぶし色を保存
    let cellFill = null;
    try {
      cellFill = cell.getFill();
      if (cellFill && cellFill.getSolidFill()) {
        console.log(`    セル背景色を保存: ${cellFill.getSolidFill().getColor().asRgbColor().asHexString()}`);
      }
    } catch (e) {
      console.log(`    セル背景色の取得エラー: ${e.message}`);
    }
    
    // テキストスタイルを保存（エラー処理を強化）
    const textStyle = textRange.getTextStyle();
    const savedFormat = {};
    
    try { savedFormat.fontSize = textStyle.getFontSize(); } catch (e) {}
    try { savedFormat.fontFamily = textStyle.getFontFamily(); } catch (e) {}
    try { savedFormat.bold = textStyle.isBold(); } catch (e) {}
    try { savedFormat.italic = textStyle.isItalic(); } catch (e) {}
    try { savedFormat.foregroundColor = textStyle.getForegroundColor(); } catch (e) {}
    
    // 翻訳
    let translatedText;
    try {
      translatedText = translateWithGoogleListAwareNew(originalText, src, tgt, enableProofread);
    } catch (error) {
      console.log(`    翻訳エラー、フォールバック使用: ${error.message}`);
      translatedText = LanguageApp.translate(originalText, src, tgt);
    }
    
    console.log(`    → "${translatedText.substring(0, 30)}..."`);
    
    // テキスト設定
    textRange.setText(translatedText);
    
    // ★重要：セルの塗りつぶし色を復元（テキスト書式より先に）
    if (cellFill) {
      try {
        cell.setFill(cellFill);
        console.log(`    セル背景色を復元`);
      } catch (e) {
        console.log(`    セル背景色の復元エラー: ${e.message}`);
      }
    }
    
    // テキスト書式を復元（エラー処理を強化）
    try {
      const newStyle = textRange.getTextStyle();
      
      if (savedFormat.fontSize) {
        try { newStyle.setFontSize(savedFormat.fontSize); } catch (e) {}
      }
      if (savedFormat.fontFamily) {
        try { newStyle.setFontFamily(savedFormat.fontFamily); } catch (e) {}
      }
      if (savedFormat.bold !== null && savedFormat.bold !== undefined) {
        try { newStyle.setBold(savedFormat.bold); } catch (e) {}
      }
      if (savedFormat.italic !== null && savedFormat.italic !== undefined) {
        try { newStyle.setItalic(savedFormat.italic); } catch (e) {}
      }
      if (savedFormat.foregroundColor) {
        try { newStyle.setForegroundColor(savedFormat.foregroundColor); } catch (e) {}
      }
      
      // 日本語の場合、フォントサイズ調整
      if (tgt === 'ja' && savedFormat.fontSize) {
        try {
          newStyle.setFontSize(Math.round(savedFormat.fontSize * 0.9));
        } catch (e) {}
      }
    } catch (error) {
      console.log(`    テキスト書式復元エラー（続行）: ${error.message}`);
    }
    
    return true;
    
  } catch (error) {
    console.log(`    [${row},${col}] エラー: ${error.message}`);
    return false;
  }
}



/**
 * 1つのテーブルを翻訳
 * @returns {number} 翻訳したセル数
 */
function translateSingleTable(table, src, tgt, enableProofread) {
  const numRows = table.getNumRows();
  const numColumns = table.getNumColumns();
  console.log(`  サイズ: ${numRows}行 × ${numColumns}列`);
  
  const processedCells = new Set();
  let translatedCount = 0;
  
  for (let row = 0; row < numRows; row++) {
    for (let col = 0; col < numColumns; col++) {
      const cellKey = `${row},${col}`;
      
      if (processedCells.has(cellKey)) {
        continue;
      }
      
      try {
        const cell = table.getCell(row, col);
        
        // マージセル処理
        try {
          const rowSpan = cell.getRowSpan() || 1;
          const colSpan = cell.getColumnSpan() || 1;
          
          for (let r = row; r < row + rowSpan; r++) {
            for (let c = col; c < col + colSpan; c++) {
              processedCells.add(`${r},${c}`);
            }
          }
        } catch (e) {
          processedCells.add(cellKey);
        }
        
        // セルの翻訳
        const wasTranslated = translateTableCell(cell, src, tgt, enableProofread, row, col);
        if (wasTranslated) {
          translatedCount++;
        }
        
      } catch (error) {
        console.log(`  [${row},${col}] エラー: ${error.message}`);
      }
    }
  }
  
  return translatedCount;
}
/**
 * スライド内の全テーブルを翻訳（新関数）
 */
function translateAllTablesInSlide(slide, src, tgt, enableProofread) {
  console.log('--- テーブル翻訳開始 ---');
  
  const pageElements = slide.getPageElements();
  let tableCount = 0;
  let totalCells = 0;
  
  for (let i = 0; i < pageElements.length; i++) {
    const element = pageElements[i];
    
    try {
      if (element.getPageElementType() === SlidesApp.PageElementType.TABLE) {
        tableCount++;
        console.log(`\nテーブル ${tableCount} (要素 ${i})`);
        
        const table = element.asTable();
        const cellCount = translateSingleTable(table, src, tgt, enableProofread);
        totalCells += cellCount;
        
        console.log(`✓ テーブル ${tableCount} 完了: ${cellCount}セル翻訳`);
      }
    } catch (error) {
      console.log(`テーブル処理エラー: ${error.message}`);
    }
  }
  
  if (tableCount === 0) {
    console.log('表が見つかりませんでした');
  } else {
    console.log(`--- テーブル翻訳完了: ${tableCount}個、${totalCells}セル ---`);
  }
}

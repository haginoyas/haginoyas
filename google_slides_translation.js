/**
 * グーグルスライドの翻訳機能を追加するメニューを作成する
 */
function onOpen() {
  SlidesApp.getUi()
    .createMenu('翻訳機能')
    .addItem('[全スライド] Google 翻訳（英語→日本語）', 'translateAllSlidesEnToJa')
    .addItem('[特定ページ] Google 翻訳（英語→日本語）', 'translateSpecificPageEnToJa')
    .addItem('[全スライド] Google 翻訳（日本語→英語）', 'translateAllSlidesJaToEn')
    .addItem('[特定ページ] Google 翻訳（日本語→英語）', 'translateSpecificPageJaToEn')
    .addToUi();
}

/**
 * 全てのスライドをGoogle翻訳APIで翻訳する（英語→日本語）
 */
function translateAllSlidesEnToJa() {
  translateAllSlides('en', 'ja');
}

/**
 * 特定のスライドページをGoogle翻訳APIで翻訳する（英語→日本語）
 */
function translateSpecificPageEnToJa() {
  translateSpecificPage('en', 'ja');
}

/**
 * 全てのスライドをGoogle翻訳APIで翻訳する（日本語→英語）
 */
function translateAllSlidesJaToEn() {
  translateAllSlides('ja', 'en');
}

/**
 * 特定のスライドページをGoogle翻訳APIで翻訳する（日本語→英語）
 */
function translateSpecificPageJaToEn() {
  translateSpecificPage('ja', 'en');
}

/**
 * Google翻訳APIを使ってテキストを翻訳する
 * @param {string} text 翻訳する元のテキスト
 * @param {string} src 元の言語コード
 * @param {string} tgt 翻訳後の言語コード
 * @returns {string} 翻訳後のテキスト
 */
function translateWithGoogle(text, src, tgt) {
  return LanguageApp.translate(text, src, tgt);
}

/**
 * スライドのテキストを翻訳する
 * @param {Slide} slide 翻訳対象のスライド
 * @param {string} src 元の言語コード
 * @param {string} tgt 翻訳後の言語コード
 */
function translateSlide(slide, src, tgt) {
  translateSpeakerNotes(slide, src, tgt);

  const pageElements = slide.getPageElements();

  for (let pageElement of pageElements) {
    if (pageElement.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      translateShape(pageElement.asShape(), slide, src, tgt);
    } else if (pageElement.getPageElementType() === SlidesApp.PageElementType.GROUP) {
      translateGroup(pageElement.asGroup(), slide, src, tgt);
    }
  }
}

/**
 * シェイプのテキストを翻訳する
 * @param {Shape} shape 翻訳対象のシェイプ
 * @param {Slide} slide シェイプが含まれるスライド
 * @param {string} src 元の言語コード
 * @param {string} tgt 翻訳後の言語コード
 */
function translateShape(shape, slide, src, tgt) {
  /*スタイル維持のために Paragraph ごとに翻訳する*/
  for (let paragraphText of shape.getText().getParagraphs()) {
    paragraphText = paragraphText.getRange();
    const translatedText = translateWithGoogle(paragraphText.asString(), src, tgt);
    paragraphText.setText(translatedText);
    reduceTextSize(paragraphText);
  }
}

/**
 * グループ内のシェイプのテキストを翻訳する
 * @param {Group} group 翻訳対象のグループ
 * @param {Slide} slide グループが含まれるスライド
 * @param {string} src 元の言語コード
 * @param {string} tgt 翻訳後の言語コード
 */
function translateGroup(group, slide, src, tgt) {
  const childElements = group.getChildren();

  for (let childElement of childElements) {
    if (childElement.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
      translateShape(childElement.asShape(), slide, src, tgt);
    } else if (childElement.getPageElementType() === SlidesApp.PageElementType.GROUP) {
      translateGroup(childElement.asGroup(), slide, src, tgt);
    }
  }
}

/**
 * スライド内のテーブルを翻訳する
 * @param {Slide} slide 翻訳対象のスライド
 * @param {string} src 元の言語コード
 * @param {string} tgt 翻訳後の言語コード
 */
function translateTableInSlide(slide, src, tgt) {
  const tables = slide.getTables();

  for (let table of tables) {
    const numRows = table.getNumRows();
    const numColumns = table.getNumColumns();

    for (let row = 0; row < numRows; row++) {
      for (let col = 0; col < numColumns; col++) {
        try {
          const cell = table.getCell(row, col);
          const originalText = cell.getText().asString();
          if (originalText === '') {
            continue;
          }

          const translatedText = translateWithGoogle(originalText, src, tgt);
          cell.getText().setText(translatedText);
          reduceTextSize(cell.getText());
        } catch (error) {
          // 結合セルの先頭以外のセルにアクセスしようとした場合はスキップ
          console.log(`セル (${row}, ${col}) をスキップしました: ${error.message}`);
          continue;
        }
      }
    }
  }
}

/**
 * 全てのスライドを翻訳する
 * @param {string} src 元の言語コード
 * @param {string} tgt 翻訳後の言語コード
 */
function translateAllSlides(src, tgt) {
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();

  for (let slide of slides) {
    translateSlide(slide, src, tgt);
    translateTableInSlide(slide, src, tgt);
  }

  /* Google 翻訳は叩きすぎると怒られるのでスリープする */
  Utilities.sleep(1000);
}

/**
 * 特定のスライドページを翻訳する
 * @param {string} src 元の言語コード
 * @param {string} tgt 翻訳後の言語コード
 */
function translateSpecificPage(src, tgt) {
  const presentation = SlidesApp.getActivePresentation();
  const slides = presentation.getSlides();
  const pageNumber = getPageNumberFromUser();

  if (pageNumber === null) {
    return;
  }

  const slide = slides[pageNumber - 1];
  translateSlide(slide, src, tgt);
  translateTableInSlide(slide, src, tgt);
}

/**
 * スピーカーノートを翻訳する
 * @param {Slide} slide テキストが含まれるスライド
 * @param {string} src 元の言語コード
 * @param {string} tgt 翻訳後の言語コード
 */
function translateSpeakerNotes(slide, src, tgt) {
  const notesPageTexts = slide.getNotesPage().getSpeakerNotesShape().getText().getParagraphs();
  for(let paragraphText of notesPageTexts) {
    paragraphText = paragraphText.getRange()
    translatedText = translateWithGoogle(paragraphText.asString(), src, tgt)
    paragraphText.setText(translatedText);
  }
}

/**
 * ユーザーからスライドのページ番号を取得する
 * @returns {?number} ページ番号 (キャンセルされた場合はnull)
 */
function getPageNumberFromUser() {
  const ui = SlidesApp.getUi();
  const response = ui.prompt(
    '翻訳こんにゃく',
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
 * テキストのサイズを減らす
 * @param {text} object
 */
function reduceTextSize(text) {
  if (text.asString().trim() !== "") {
    const textStyle = text.getTextStyle();
    if (textStyle.getFontSize() !== null) {
      // フォントサイズを減らす
      textStyle.setFontSize((textStyle.getFontSize() * 0.9).toFixed());
    }
  }
}

/**
 * Googleスプレッドシートの翻訳機能を追加するメニューを作成する
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('翻訳機能')
    .addItem('[選択範囲] Google 翻訳（英語→日本語）', 'translateSelectionEnToJa')
    .addItem('[選択範囲] Google 翻訳（日本語→英語）', 'translateSelectionJaToEn')
    .addItem('[指定した列] Google 翻訳（英語→日本語）', 'translateSpecificColumnEnToJa')
    .addItem('[指定した列] Google 翻訳（日本語→英語）', 'translateSpecificColumnJaToEn')
    .addToUi();
}

/**
 * 選択範囲をGoogle翻訳APIで翻訳する（英語→日本語）
 */
function translateSelectionEnToJa() {
  translateSelection('en', 'ja');
}

/**
 * 選択範囲をGoogle翻訳APIで翻訳する（日本語→英語）
 */
function translateSelectionJaToEn() {
  translateSelection('ja', 'en');
}



/**
 * 指定した列をGoogle翻訳APIで翻訳する（英語→日本語）
 */
function translateSpecificColumnEnToJa() {
  translateSpecificColumn('en', 'ja');
}

/**
 * 指定した列をGoogle翻訳APIで翻訳する（日本語→英語）
 */
function translateSpecificColumnJaToEn() {
  translateSpecificColumn('ja', 'en');
}

/**
 * Google翻訳APIを使ってテキストを翻訳する
 * @param {string} text 翻訳する元のテキスト
 * @param {string} src 元の言語コード
 * @param {string} tgt 翻訳後の言語コード
 * @returns {string} 翻訳後のテキスト
 */
function translateWithGoogle(text, src, tgt) {
  // 空文字やnullの場合はそのまま返す
  if (!text || text.toString().trim() === '') {
    return text;
  }
  
  try {
    const result = LanguageApp.translate(text.toString(), src, tgt);
    // API制限回避のための短いスリープ
    Utilities.sleep(100);
    return result;
  } catch (error) {
    console.log(`翻訳エラー: ${error.message}`);
    return text; // エラーの場合は元のテキストを返す
  }
}

/**
 * 選択範囲のセルを翻訳する
 * @param {string} src 元の言語コード
 * @param {string} tgt 翻訳後の言語コード
 */
function translateSelection(src, tgt) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  
  if (!range) {
    showAlert('範囲が選択されていません', '翻訳したい範囲を選択してから実行してください。');
    return;
  }
  
  const values = range.getValues();
  const numRows = values.length;
  const numCols = values[0].length;
  
  // 進行状況を表示するためのトースト
  showToast('翻訳中...', `${numRows}行 × ${numCols}列の範囲を翻訳しています`);
  
  // 各セルを翻訳
  for (let row = 0; row < numRows; row++) {
    for (let col = 0; col < numCols; col++) {
      if (values[row][col]) {
        values[row][col] = translateWithGoogle(values[row][col], src, tgt);
      }
    }
    
    // 行ごとに進行状況を更新
    if (row % 5 === 0 || row === numRows - 1) {
      showToast('翻訳中...', `${row + 1}/${numRows}行 完了`);
    }
  }
  
  // 翻訳結果を設定
  range.setValues(values);
  showToast('翻訳完了', '選択範囲の翻訳が完了しました！');
}



/**
 * 指定した列を翻訳する
 * @param {string} src 元の言語コード
 * @param {string} tgt 翻訳後の言語コード
 */
function translateSpecificColumn(src, tgt) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    '列指定翻訳',
    '翻訳したい列を指定してください（例：A、B、C など）',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.CANCEL) {
    return;
  }

  const columnLetter = response.getResponseText().trim().toUpperCase();
  
  if (!columnLetter.match(/^[A-Z]+$/)) {
    showAlert('入力エラー', '正しい列名を入力してください（例：A、B、C など）');
    return;
  }

  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  
  if (lastRow === 0) {
    showAlert('データがありません', '翻訳するデータがシートに見つかりません。');
    return;
  }

  try {
    const range = sheet.getRange(`${columnLetter}1:${columnLetter}${lastRow}`);
    const values = range.getValues();
    
    showToast('翻訳中...', `${columnLetter}列（${lastRow}行）を翻訳しています`);
    
    // 各セルを翻訳
    for (let row = 0; row < values.length; row++) {
      if (values[row][0]) {
        values[row][0] = translateWithGoogle(values[row][0], src, tgt);
      }
      
      // 進行状況を更新（10行ごと）
      if (row % 10 === 0 || row === values.length - 1) {
        showToast('翻訳中...', `${row + 1}/${values.length}行 完了`);
      }
    }
    
    // 翻訳結果を設定
    range.setValues(values);
    showToast('翻訳完了', `${columnLetter}列の翻訳が完了しました！`);
    
  } catch (error) {
    showAlert('エラー', `指定された列が見つかりません: ${columnLetter}`);
  }
}

/**
 * トーストメッセージを表示する
 * @param {string} title タイトル
 * @param {string} message メッセージ
 */
function showToast(title, message) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title, 3);
}

/**
 * アラートダイアログを表示する
 * @param {string} title タイトル
 * @param {string} message メッセージ
 */
function showAlert(title, message) {
  const ui = SpreadsheetApp.getUi();
  ui.alert(title, message, ui.ButtonSet.OK);
}

/**
 * カスタム関数：セル内のテキストを翻訳する
 * スプレッドシートで =TRANSLATE_TEXT(A1, "en", "ja") のように使用可能
 * @param {string} text 翻訳するテキスト
 * @param {string} src 元の言語コード
 * @param {string} tgt 翻訳後の言語コード
 * @returns {string} 翻訳後のテキスト
 */
function TRANSLATE_TEXT(text, src, tgt) {
  return translateWithGoogle(text, src, tgt);
}

/**
 * カスタム関数：英語から日本語に翻訳する
 * スプレッドシートで =EN_TO_JA(A1) のように使用可能
 * @param {string} text 翻訳するテキスト
 * @returns {string} 翻訳後のテキスト
 */
function EN_TO_JA(text) {
  return translateWithGoogle(text, 'en', 'ja');
}

/**
 * カスタム関数：日本語から英語に翻訳する
 * スプレッドシートで =JA_TO_EN(A1) のように使用可能
 * @param {string} text 翻訳するテキスト
 * @returns {string} 翻訳後のテキスト
 */
function JA_TO_EN(text) {
  return translateWithGoogle(text, 'ja', 'en');
}

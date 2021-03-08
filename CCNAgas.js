function run() {

  const runSheetName = "テスト作成";
  const tableSheetName = "テーブル";

  var form = call_createForm(runSheetName, tableSheetName, 2, 4);

  Logger.log('Published URL: ' + form.getPublishedUrl());
  Logger.log('Editor URL: ' + form.getEditUrl());
  
  const folderId = '1-C23Mbz4Q7IvRpocpmL39kRRROiTLJvv';
  moveForm(form, folderId);
  
}

function doGet() {
  return HtmlService.createTemplateFromFile("CCNAtest").evaluate();
}

function doPost() {
  return HtmlService.createTemplateFromFile("CCNAtest").evaluate();
}

/**
 * call_createForm
 * Googleフォームを生成するためのラッパー関数
 *
 * @param {sheetName:text} フォームの問題文と選択肢を取得するシートの名前
 * @param {startRow:int} 問題文と選択肢が格納されている先頭の行番号
 * @param {startCol:int} 問題文と選択肢が格納されている先頭の列番号
 * @return {form} 生成されたGoogleフォーム(オブジェクト)
 */
function call_createForm(runSheetName, tableSheetName, startRow, startCol) {

  return createForm(getTitle(runSheetName), getDescription(runSheetName), getData(tableSheetName, startRow, startCol));

}

/**
 * getTitle
 * 指定されたシートからフォームのタイトルを取得する
 *
 * @param {sheetName:text} フォームのタイトルを取得するシートの名前
 * @return {text} フォームのタイトル
 */
function getTitle(sheetName) {

  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange("B1").getValue();

}

/**
 * getDescription
 * 指定されたシートからフォームの説明を取得する
 *
 * @param {sheetName:text} フォームの説明を取得するシートの名前
 * @return {text} フォームの説明
 */
function getDescription(sheetName) {

  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getRange("B2").getValue();
  
}

/**
 * getData
 * 指定されたシートからフォームの問題文と選択肢を取得する
 *
 * @param {sheetName:text} フォームの問題文と選択肢を取得するシートの名前
 * @param {startRow:int} 問題文と選択肢が格納されている先頭の行番号
 * @param {startCol:int} 問題文と選択肢が格納されている先頭の列番号
 * @return {array} 問題文と選択肢の２次元配列 (指定されたセルから最終行、最終列までが2次元配列として返される)
 */
function getData(sheetName, startRow, startCol) {
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  var rows = sheet.getLastRow();
  var cols = sheet.getLastColumn();
  
  return sheet.getRange(startRow, startCol, rows - startRow + 1, cols - startCol + 1).getValues();
  
}

/**
 * createForm
 * 指定されたデータでGoogleフォームを生成する
 *
 * @param {title:text} フォームの問題文と選択肢を取得するシートの名前
 * @param {description:text} 問題文と選択肢が格納されている先頭の行番号
 * @param {data:array} 問題文と選択肢が格納されている先頭の列番号
 * @return {form} 生成されたGoogleフォーム(オブジェクト)
 */
function createForm(title, description, data) {
  
  var form = FormApp.create(title);
  
  form.setDescription(description)
      .setIsQuiz(true)
      .setShowLinkToRespondAgain(false);
      
  const validationEmail = FormApp.createTextValidation().requireTextIsEmail().build();
  form.addTextItem().setTitle('回答者(メールアドレス)').setRequired(true).setValidation(validationEmail);

  for (var i = 0 ; i < data.length ; i++) {
    
    var qa = data[i];
    
    // 複数選択の場合、checkboxに条件分岐が必要
    var item = form.addMultipleChoiceItem()
    //var item = form.addCheckboxItem();
    
    var numCols = qa.length;
    var answer = 0;
    var comment = 0;
    
    numCols = numCols - 2;
    answer = qa[qa.length - 2];  
    comment = qa[qa.length - 1];
    
    item.setTitle(qa[0]);
    
    var choices = [];
    
    for (var j = 1 ; j < numCols ; j++) {
      choices.push(item.createChoice(qa[j], j == answer));
    }
    
    item.setChoices(choices);
    item.setPoints(1);
    item.setFeedbackForCorrect(FormApp.createFeedback().setText(comment).build());
    item.setFeedbackForIncorrect(FormApp.createFeedback().setText(comment).build());
    
  }
  
  return form;
}

/**
 * moveForm
 * 生成したフォームを指定のフォルダに移動する
 * 
 * @param form 生成されたGoogleフォーム(オブジェクト)
 * @param folderId 移動先のフォルダID
 */
function moveForm(form, folderId) {

   const file = DriveApp.getFileById(form.getId());
   const folder = DriveApp.getFolderById(folderId);
   file.moveTo(folder);

}
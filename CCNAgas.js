function run(title, description) {

  const sheetId = '1Z081NXbESs3ScnEbHABV2n4msii26mDqW2KaSbDuE3g';
  const sheetName = "テーブル";
  
  var data = getData(sheetId, sheetName, 2, 4)
  var form = createForm(title, description, data);

  Logger.log('Published URL: ' + form.getPublishedUrl());
  Logger.log('Editor URL: ' + form.getEditUrl());
  
  const folderId = '1-C23Mbz4Q7IvRpocpmL39kRRROiTLJvv';
  moveForm(form, folderId);
  
}

function doGet() {
  return HtmlService.createTemplateFromFile("CCNAtest").evaluate();
}

function doPost(e) {
  var title = e.parameter.title;
  var description = e.parameter.description;
  run(title, description);

  return HtmlService.createTemplateFromFile("result").evaluate();
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
function getData(sheetId, sheetName, startRow, startCol) {
  
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  
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
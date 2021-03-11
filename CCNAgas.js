const sheetId = '1Z081NXbESs3ScnEbHABV2n4msii26mDqW2KaSbDuE3g'; // スプレッドシートID
const inputSheet = 'テーブル'; // データテーブルのシート名
const outputSheet = 'テスト作成';　//　出力先のシート名
const folderId = '1-C23Mbz4Q7IvRpocpmL39kRRROiTLJvv'; // 移動先のフォルダID
var title;
var description;
var section;

function doGet() {
  return HtmlService.createTemplateFromFile("CCNAtest").evaluate();
}

function doPost(e) {
  
  title = e.parameter.title;
  description = e.parameter.description;
  section = e.parameter.section;
  
  var data = getData(1, 2);
  var form = createForm(title, description, data);
  
  moveForm(form);

  return ContentService.createTextOutput('Published URL: ' + form.getPublishedUrl() + '、Editor URL: ' + form.getEditUrl());
}

/**
 * getData
 * 指定されたシートからフォームの問題文と選択肢を取得する
 *
 * @param {outputSheet:text} フォームの問題文と選択肢を取得するシートの名前
 * @param {startRow:int} 問題文と選択肢が格納されている先頭の行番号
 * @param {startCol:int} 問題文と選択肢が格納されている先頭の列番号
 * @return {array} 問題文と選択肢の２次元配列 (指定されたセルから最終行、最終列までが2次元配列として返される)
 */
function getData(startRow, startCol) {
  
  var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(outputSheet);
  var colNameRow = '=index(\'' + inputSheet + '\'!A1:L1)';
  var query = '=query(\'' + inputSheet + '\'!A:L, "select * where B = \'' + section + '\'")';

  sheet.getRange(1,1).setValue(colNameRow);
  sheet.getRange(2,1).setValue(query);

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
  
  var colName = data[0];
  for (var i = 1 ; i < data.length ; i++) {
    
    var qa = data[i];
    var titleNum = qa[0] + '-' + qa[1] + '：';    
    // 複数選択の場合、checkboxに条件分岐が必要
    var item = form.addMultipleChoiceItem();
    //var item = form.addCheckboxItem();
    
    var numCols = qa.length;
    var answer = 0;
    var comment = 0;
    
    numCols = numCols - 2;
    answer = qa[qa.length - 2];  
    comment = qa[qa.length - 1];
    
    item.setTitle(titleNum + qa[2]);
    
    var choices = [];
    
    for (var j = 3 ; j < numCols ; j++) {
      if(qa[j] != '') {
        var choice = colName[j] + '：' + qa[j];
      } else {
        break;
      }
      choices.push(item.createChoice(choice, j == answer));
    }
    
    item.setChoices(choices)
        .setPoints(1)
        .setFeedbackForCorrect(FormApp.createFeedback().setText(comment).build())
        .setFeedbackForIncorrect(FormApp.createFeedback().setText(comment).build());
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
function moveForm(form) {

   const file = DriveApp.getFileById(form.getId());
   const folder = DriveApp.getFolderById(folderId);
   file.moveTo(folder);

}
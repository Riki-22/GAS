const formId = '1DBtgB_OBGMczLGXKKfnD1v3m0DxnyZormL8Vp1TAbqQ'; // コピー元のフォームID
const sheetId = '1Z081NXbESs3ScnEbHABV2n4msii26mDqW2KaSbDuE3g'; // スプレッドシートのID
const inputSheet = 'テーブル'; // データテーブルのシート名
const outputSheet = 'テスト作成';　//　出力先のシート名
const imageFolderId = '1Tbd4EPXxSxyIirGY8S7mMmPCEyg_GUdN'; // 画像が保存されているフォルダID
const formFolderId = '1-C23Mbz4Q7IvRpocpmL39kRRROiTLJvv'; // 移動先のフォルダID
var mail;
var title;
var description;
var section;
var random;

function setResponses(e) {

  mail = e.response.getRespondentEmail();
  var itemResponses = e.response.getItemResponses();
  title = itemResponses[0].getResponse();
  description = itemResponses[1].getResponse();
  section = itemResponses[2].getResponse();
  random = itemResponses[3].getResponse();

  run();
}

function run() {
  
  var data = getData(1, 2);
  var form = createForm(title, description, data);
  
  moveForm(form);
  sendMail(form);
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
  var colNameRow = '=index(\'' + inputSheet + '\'!A1:N1)';
  var query = '=query(\'' + inputSheet + '\'!A:N, "select * where B = \'' + section + '\'")';

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
 * @param {colName:array} 選択肢a~fを格納
 * @return {form} 生成されたGoogleフォーム(オブジェクト)
 */
function createForm(title, description, data) {
  
  var doc = DriveApp.getFileById(formId);
  var file = doc.makeCopy(title);
  var form = FormApp.openById(file.getId());

  form.setTitle(title)
      .setDescription(description);
    
  var colName = data[0];
  for (var i = 1 ; i < data.length ; i++) {
    
    var qa = data[i];
    var questionNum = qa[0] + '-' + qa[1];
    var imageName = questionNum + '.png';  // ※正しい拡張子をつけること

    var fol = DriveApp.getFolderById(imageFolderId);
    if (fol.getFilesByName(imageName).hasNext()) {
      
      var blob = fol.getFilesByName(imageName).next().getBlob();
      form.addImageItem().setImage(blob);
    }


    var answer = qa[qa.length - 2];
    var comment = qa[qa.length - 1];
    var choices = [];
    var choice = colName.slice(3, 9);
    if (answer.length == 1) {

      var item = form.addMultipleChoiceItem();
        
      item.setTitle(questionNum + '：' + qa[2]);
      
      for (var j = 0 ; j < choice.length ; j++) {
      
        var k = j + 3;
        if(qa[k] != '') {
        
          var question = colName[k] + '：' + qa[k];
          choices.push(item.createChoice(question, choice[j] == answer));
        } else {
          
          break;
        }
      }
    } else {

      var item = form.addCheckboxItem();

      var answers = answer.split(',');
      
      item.setTitle(questionNum + '：' + qa[2]);
      
      for (var j = 0 ; j < choice.length ; j++) {
      
        var k = j + 3;
        if(qa[k] != '') {
        
          var question = colName[k] + '：' + qa[k];
          choices.push(item.createChoice(question, answers.includes(choice[j])));
        } else {
          
          break;
        }
      }
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
 * @param formFolderId 移動先のフォルダID
 */
function moveForm(form) {

   var file = DriveApp.getFileById(form.getId());
   var folder = DriveApp.getFolderById(formFolderId);
   file.moveTo(folder);
}

function sendMail(form) {

  var subject = 'テスト送信';
  var body = '公開用 URL: ' + form.getPublishedUrl() + '\n編集用 URL: ' + form.getEditUrl();
  GmailApp.sendEmail(mail, subject, body);
}
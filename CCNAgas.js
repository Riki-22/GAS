const formId = '1DBtgB_OBGMczLGXKKfnD1v3m0DxnyZormL8Vp1TAbqQ'; // コピー元のフォームID
const sheetId = '1Z081NXbESs3ScnEbHABV2n4msii26mDqW2KaSbDuE3g'; // スプレッドシートのID
const inputSheet = 'テーブル'; // データテーブルのシート名
const outputSheet = 'テスト作成';　//　出力先のシート名
const imageFolderId = '1Tbd4EPXxSxyIirGY8S7mMmPCEyg_GUdN'; // 画像が保存されているフォルダID
const formFolderId = '1-C23Mbz4Q7IvRpocpmL39kRRROiTLJvv'; // 移動先のフォルダID


// 必ずインストーラブルトリガーをrunに指定すること
function run(e) {

  let mailAddress = e.response.getRespondentEmail();
  let itemResponses = e.response.getItemResponses();
  let title = itemResponses[0].getResponse();
  let description = itemResponses[1].getResponse();
  let section = itemResponses[2].getResponse();
  let random = itemResponses[3].getResponse();
  let maxItem = itemResponses[4].getResponse();

  let data = getData(section, 1, 2); // データを取得するセルの開始位置(sectionカラムの最初のレコード)を入力
  let form = createForm(title, description, data, maxItem);
  
  moveForm(form);
  sendMail(mailAddress, form);
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
function getData(section, startRow, startCol) {
  
  let sheet = SpreadsheetApp.openById(sheetId).getSheetByName(outputSheet);
  let colNameRow = '=index(\'' + inputSheet + '\'!A1:N1)';
  let query = '=query(\'' + inputSheet + '\'!A:N, "select * where B = \'' + section.join('\' or B = \'') + '\'")';

  sheet.getRange(1,1).setValue(colNameRow);
  sheet.getRange(2,1).setValue(query);

  let rows = sheet.getLastRow();
  let cols = sheet.getLastColumn();
  
  return sheet.getRange(startRow, startCol, rows - startRow + 1, cols - startCol + 1).getValues();
}

/**
 * createForm
 * 指定されたデータでGoogleフォームを生成する
 *
 * @param {title:text} フォームの問題文と選択肢を取得するシートの名前
 * @param {description:text} 問題文と選択肢が格納されている先頭の行番号
 * @param {data:array} 問題文と選択肢が格納されている先頭の列番号
 * @param {colName:array} 全てのカラム名
 * @return {form} 生成されたGoogleフォーム(オブジェクト)
 */
function createForm(title, description, data, maxItem) {
  
  let doc = DriveApp.getFileById(formId);
  let file = doc.makeCopy(title);
  let form = FormApp.openById(file.getId());

  form.setTitle(title)
    .setDescription(description);
    
  let dataLength;

  if (maxItem == 'なし') {

    dataLength = data.length;
  } else {

    if(data.length < maxItem) {

      dataLength = data.length;
    } else {

      dataLength = Number(maxItem) + 1;
    }
  }
  
  for (let i = 1 ; i < dataLength ; i++) {
    
    let recode = data[i];
    let questionNum = recode[0] + '-' + recode[1];
    let imageName = questionNum + '.png';  // ※正しい拡張子をつけること
    let imageFolder = DriveApp.getFolderById(imageFolderId);

    if (imageFolder.getFilesByName(imageName).hasNext()) {
      
      let blob = imageFolder.getFilesByName(imageName).next().getBlob();
      let imageItem = form.addImageItem();
      
      imageItem.setImage(blob)
        .setTitle(questionNum + '：図を参照して次の設問に回答してください。')
    }

    let answer = recode[recode.length - 2];
    let comment = recode[recode.length - 1];
    let choices = [];
    let colName = data[0];
    let choice = colName.slice(3, 9);
    let item;

    if (answer.length == 1) {

      item = form.addMultipleChoiceItem();
        
      item.setTitle(questionNum + '：' + recode[2]);
      
      for (let j = 0 ; j < choice.length ; j++) {
      
        let k = j + 3;
        if(recode[k] != '') {
        
          let question = colName[k] + '：' + recode[k];
          choices.push(item.createChoice(question, choice[j] == answer));
        } else {
          
          break;
        }
      }
    } else {

      item = form.addCheckboxItem();
      let answers = answer.split(',');
      
      item.setTitle(questionNum + '：' + recode[2]);
      
      for (let j = 0 ; j < choice.length ; j++) {
      
        let k = j + 3;
        if(recode[k] != '') {
        
          let question = colName[k] + '：' + recode[k];
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

  let createdForm = DriveApp.getFileById(form.getId());
  let folder = DriveApp.getFolderById(formFolderId);
  createdForm.moveTo(folder);
}

function sendMail(mailAddress, form) {

  let subject = 'テスト送信';
  let body = '公開用 URL: ' + form.getPublishedUrl() + '\n編集用 URL: ' + form.getEditUrl();
  GmailApp.sendEmail(mailAddress, subject, body);
}
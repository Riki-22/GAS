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
  let maxItem = Number(itemResponses[3].getResponse());
  let random = itemResponses[4].getResponse();

  let data = getData(section, 1, 2); // データを取得するセルの開始位置(sectionカラムの最初のレコード)を入力
  let colName = data.splice(0, 1)[0]; // カラム名が格納されている最初の配列を切り出す

  if (random == 'ランダムにする') {
    
    data = dataShuffle(data);
  }

  let form = createForm(title, description, colName, data, maxItem);
  
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
  let colNameQuery = '=index(\'' + inputSheet + '\'!A1:N1)';
  let recodeQuery = '=query(\'' + inputSheet + '\'!A:N, "select * where B = \'' + section.join('\' or B = \'') + '\'")';

  sheet.getRange(1,1).setValue(colNameQuery);
  sheet.getRange(2,1).setValue(recodeQuery);

  let rows = sheet.getLastRow();
  let cols = sheet.getLastColumn();
  
  return sheet.getRange(startRow, startCol, rows - startRow + 1, cols - startCol + 1).getValues();
}

function dataShuffle(data) {
  
  for(let i = (data.length - 1); 0 < i; i--){

    // 0〜(i+1)の範囲で値を取得
    let r = Math.floor(Math.random() * (i + 1));

    // 要素の並び替えを実行
    let tmp = data[i];
    data[i] = data[r];
    data[r] = tmp;
  }
  return data;
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
function createForm(title, description, colName, data, maxItem) {
  
  let doc = DriveApp.getFileById(formId);
  let file = doc.makeCopy(title);
  let form = FormApp.openById(file.getId());

  form.setTitle(title)
    .setDescription(description);
    
  let numOfQuestions;
  // 出題数が上限を上回る場合、上限値を代入
  if (data.length > maxItem) {

    numOfQuestions = maxItem;
  } else {

    numOfQuestions = data.length;
  }

  // 出題数の数だけ繰り返す
  for (let i = 0 ; i < numOfQuestions ; i++) {
    
    let recode = data[i];
    let section_questionNum = recode[0] + '-' + recode[1];
    let imageName = section_questionNum + '.png';  // ※正しい拡張子をつけること
    let imageFolder = DriveApp.getFolderById(imageFolderId);

    // 画像フォルダを参照し、問題番号と同じファイル名の画像がある場合は画像を挿入
    if (imageFolder.getFilesByName(imageName).hasNext()) {
      
      let blob = imageFolder.getFilesByName(imageName).next().getBlob();
      let imageItem = form.addImageItem();
      
      imageItem.setImage(blob)
        .setTitle(section_questionNum + '：図を参照して次の設問に回答してください。')
    }

    let questionTitle = section_questionNum + '：' + recode[2];
    let answer = recode[recode.length - 2];
    let commentary = recode[recode.length - 1];
    let item;
    let choices = [];

    // 解答が１つの場合はラジオアイテム、複数の場合はチェックボックスアイテムを追加
    if (answer.length == 1) {

      item = form.addMultipleChoiceItem();
      
      item.setTitle(questionTitle);
      
      // 選択肢a~fまでを追加
      for (let j = colName.indexOf('a') ; j <= colName.indexOf('f') ; j++) {
      
        if(recode[j] != '') {
        
          let choiceTitle = colName[j] + '：' + recode[j];
          choices.push(item.createChoice(choiceTitle, colName[j] == answer));
        } else {
          
          break;
        }
      }
    } else {

      item = form.addCheckboxItem();
      let answers = answer.split(',');
      
      item.setTitle(questionTitle);
      
      for (let j = colName.indexOf('a') ; j <= colName.indexOf('f') ; j++) {
      
        if(recode[j] != '') {
        
          let choiceTitle = colName[j] + '：' + recode[j];
          choices.push(item.createChoice(choiceTitle, answers.includes(colName[j])));
        } else {
          
          break;
        }
      }
    }

    item.setChoices(choices)
      .setPoints(1)
      .setFeedbackForCorrect(FormApp.createFeedback().setText(commentary).build())
      .setFeedbackForIncorrect(FormApp.createFeedback().setText(commentary).build());
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
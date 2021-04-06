
/**
 * run
 * App Scriptダッシュボードでフォーム送信時のトリガーをrunに指定すること
 * 
 * @param {*} e イベントオブジェクト
 */
function run(e) {

  const scriptProperties = PropertiesService.getScriptProperties();
  const formId = scriptProperties.getProperty('formId');
  const sheetId = scriptProperties.getProperty('sheetId');
  const inputSheet = scriptProperties.getProperty('inputSheet');
  const outputSheet = scriptProperties.getProperty('outputSheet');
  const imageFolderId = scriptProperties.getProperty('imageFolderId');
  const formFolderId = scriptProperties.getProperty('formFolderId');

  const mailAddress = e.response.getRespondentEmail();
  const itemResponses = e.response.getItemResponses();
  const title = itemResponses[0].getResponse();
  const description = itemResponses[1].getResponse();
  const section = itemResponses[2].getResponse();
  const maxItem = Number(itemResponses[3].getResponse());
  const random = itemResponses[4].getResponse();

  let data = getData(sheetId, inputSheet, outputSheet, section);
  let colName = data.splice(0, 1)[0]; // カラム名が格納されている最初の配列を切り出す

  if (random == 'ランダムにする') {
    
    data = dataShuffle(data);
  }

  let form = createForm(formId, imageFolderId, title, description, colName, data, maxItem);
  
  moveForm(formFolderId, form);
  sendMail(mailAddress, form);
}

/**
 * getData
 * データが格納されているシート(inputSheet)を参照するQUERY関数を出力先のシート(outputSheet)で実行し、データを全て取得する
 * 
 * @param {string} sheetId データを取得するスプレッドシートのID
 * @param {string} inputSheet データが格納されているシートの名前
 * @param {string} outputSheet  QUERY関数を実行するシートの名前
 * @param {array} section フォームで入力された章番号が格納された１次元配列
 * @returns {array} 取得したデータの２次元配列
 */
function getData(sheetId, inputSheet, outputSheet, section) {
  
  let sheet = SpreadsheetApp.openById(sheetId).getSheetByName(outputSheet);
  let colNameQuery = '=index(\'' + inputSheet + '\'!B1:N1)';
  let recodeQuery = '=query(\'' + inputSheet + '\'!B:N, "select * where B = \'' + section.join('\' or B = \'') + '\'")';

  sheet.getRange(1,1).setValue(colNameQuery);
  sheet.getRange(2,1).setValue(recodeQuery);

  let lastRow = sheet.getLastRow();
  let lastCol = sheet.getLastColumn();
  
  return sheet.getRange(1, 1, lastRow, lastCol).getValues();
}

/**
 * dataShuffle
 * 配列の要素をランダムに並び替える
 * 
 * @param {*} data 
 * @returns {array}
 */
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
 * コピーしたフォームに指定されたデータを追加する
 * 
 * @param {string} formId コピー元のフォームのID
 * @param {string} imageFolderId 画像が格納されているフォルダのID
 * @param {string} title 入力されたフォームのタイトル
 * @param {string} description 入力されたフォームの説明
 * @param {array} colName カラム名が格納されている配列
 * @param {array} data 取得したデータの２次元配列
 * @param {int} maxItem 入力された出題数の上限
 * @returns {form} Googleフォーム(オブジェクト)
 */
function createForm(formId, imageFolderId, title, description, colName, data, maxItem) {
  
  let doc = DriveApp.getFileById(formId);
  let file = doc.makeCopy(title);
  let copiedForm = FormApp.openById(file.getId());

  copiedForm.setTitle(title)
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
      let imageItem = copiedForm.addImageItem();
      
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

      item = copiedForm.addMultipleChoiceItem();      
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

      item = copiedForm.addCheckboxItem();
      item.setTitle(questionTitle);

      let answers = answer.split(',');
      
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
  
  return copiedForm;
}

/**
 * moveForm
 * フォームを指定のフォルダに移動する
 * 
 * @param form Googleフォーム(オブジェクト)
 * @param formFolderId 移動先のフォルダID
 */
function moveForm(formFolderId, form) {

  let createdForm = DriveApp.getFileById(form.getId());
  let folder = DriveApp.getFolderById(formFolderId);
  createdForm.moveTo(folder);
}

/**
 * sendMail
 * フォームのリンクを指定された宛先にメールで送信する
 * 
 * @param {string} mailAddress 入力されたメールアドレス 
 * @param {*} form Googleフォーム(オブジェクト)
 */
function sendMail(mailAddress, form) {

  let subject = 'テスト送信';
  let body = '公開用 URL: ' + form.getPublishedUrl() + '\n編集用 URL: ' + form.getEditUrl();
  GmailApp.sendEmail(mailAddress, subject, body);
}
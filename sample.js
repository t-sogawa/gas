function main() {
  var templateDocId  = '1s9lEtv-V3n950TZvGFUqTVQolN20Q-7VNp6H96y0toA';
  var resultSpreadId = '1uKZB1hd_G0M8rMwerrhdtrtqHe0vzOax7w-bY5YsV5w';
  var saveFolderName = '189RYydvQYAMmGFfNHES8hoijlJAQDCwE';
  
  var newFileName    = getDate('yyyyMMdd') + '_コピーファイル';
  
  // テンプレートファイルから新規コピーを作成
  var newFile = copyDocument(templateDocId, newFileName, saveFolderName);
  // 文言を差し替え
  replaceDocument(newFile.getId());
}

function check(){
  var resultSpreadId = '1uKZB1hd_G0M8rMwerrhdtrtqHe0vzOax7w-bY5YsV5w';
  var resultSpread = SpreadsheetApp.openById(resultSpreadId);
  var resultSheet = resultSpread.getActiveSheet();
  var rows = resultSheet.getRange(2,1,1,30).getValues();
  Logger.log(rows[0][1]);

//  for(var i=2;i<=rowSheet;i++){
// 
//    var strTimeStamp =resultSpread.getRange(i,1).getValue(); //社名
//    var strYMD       =resultSpread.getRange(i,2).getValue(); //姓
//    var strAddress   =resultSpread.getRange(i,3).getValue();　//名
// 
//    var strBody=strDoc.replace(/Name}/,'soggggaaa').replace(/gggg/,strYMD).replace(/Hogera/,strAddress); //社名、姓名を置換
//    Logger.log(strBody); //ドキュメントの内容をログに表示
//  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('メッセージ表示');
  menu.addItem('Hello world! 実行', 'myFunction');
  menu.addToUi();
}
 
function myFunction() {
  Browser.msgBox('Hello world!');
}

/**
* GoogleDrive上でのファイルコピー
* @param {string} [fieldId] - コピー元ファイルID
* @param {string} [name] - 作成ファイル名
* @param {string} [destination] - 作成フォルダID
* @return {file}
*/
function copyDocument(fileId, name, folderId)
{
  var file = DriveApp.getFileById(fileId);
  var destination = DriveApp.getFolderById(folderId);
  
  if (folderId == undefined) return file.makeCopy(name);
  return file.makeCopy(name, destination);
}

function replaceDocument(docId){
  var doc = DocumentApp.openById(docId);
  var body = doc.getBody();
  
  body.replaceText('ここにテキストを挿入', 'KOKONITEXT');
  body.replaceText('{{Name}}', 'IKEDA');
  
  doc.saveAndClose();
}

function insertName(){
 
  /* スプレッドシートのシートを取得と準備 */
  var resultSpread  = SpreadsheetApp.getActiveSheet(); //シートを取得
  var rowSheet = resultSpread.getDataRange().getLastRow(); //シートの使用範囲のうち最終行を取得
 
  /* ドキュメント「メール本文テスト」を取得する */
  var docTest=DocumentApp.openById("1s9lEtv-V3n950TZvGFUqTVQolN20Q-7VNp6H96y0toA"); //ドキュメントをIDで取得
  var strDoc=docTest.getBody().getText(); //ドキュメントの内容を取得
  docTest.getBody().replaceText('/Name/', 'sogwaaaaaaa');
  docTest.getBody().getText().replace('/Name/', 'sogwaaaaaaa');
//  var strBody=strDoc.replace(/Name/,'soggggaaa');
  Logger.log(strDoc);
 
//  /* シートの全ての行について社名、姓名を差し込みログに表示*/
//  for(var i=2;i<=rowSheet;i++){
// 
    var strTimeStamp =resultSpread.getRange(i,1).getValue(); //社名
    var strYMD       =resultSpread.getRange(i,2).getValue(); //姓
    var strAddress   =resultSpread.getRange(i,3).getValue();　//名
// 
//    var strBody=strDoc.replace(/Name}/,'soggggaaa').replace(/gggg/,strYMD).replace(/Hogera/,strAddress); //社名、姓名を置換
//    Logger.log(strBody); //ドキュメントの内容をログに表示
//  }
}

/**
* 現時刻を指定フォーマットで取得
* @param {string} format - YYYYMMDD等のフォーマット指定
* @return {string}
*/
function getDate(format){
  var date = new Date();
  date.setDate(date.getDate() + 7);
  return Utilities.formatDate(date, 'JST', format);
}

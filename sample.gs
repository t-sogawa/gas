// 初期化 スクリプトプロパティの読み込み
var TEMPLATE_DOC_ID       = PropertiesService.getScriptProperties().getProperty("TEMPLATE_DOC_ID");       // テンプレートドキュメントID
var RESULT_SPREAD_ID      = PropertiesService.getScriptProperties().getProperty("RESULT_SPREAD_ID");      // フォーム回答スプレッドシートID
var SAVE_FOLDER_NAME      = PropertiesService.getScriptProperties().getProperty("SAVE_FOLDER_NAME");      // 保存先フォルダID
var TITLE_WORD_COLUMN_NUM = PropertiesService.getScriptProperties().getProperty("TITLE_WORD_COLUMN_NUM"); // タイトル名に使用するカラム列数

// 初期化 スプレッドシートオブジェクトの読み込み(スプレッドオブジェクトはスクリプト起動時に全て読み込む)
var resultSpread = SpreadsheetApp.openById(RESULT_SPREAD_ID); // フォーム回答スプレッドオブジェクト変数
var resultSheet  = resultSpread.getActiveSheet();             // フォーム回答スプレッドシートオブジェト変数

/**
* フォーム送信トリガー用メイン関数
* @param {int} [line]            - スプレッドシートの指定行
* @return {string} [newFileName] - 作成ドキュメントのファイル名
*/
function main() {
  Logger.log('replaceDocument:LINE = '); 
  createDocument();
}

/**
* スプレッドシートメニューから実行
* @param {int} [line]            - スプレッドシートの指定行
* @return {string} [newFileName] - 作成ドキュメントのファイル名
*/
function createDocumentInteractive() {
  if (resultSheet.getLastRow() < 2) {
      Browser.msgBox('データが1件も登録されていません。');
      return;
  }

  var line = Browser.inputBox('データ行数を半角数字で入力して下さい。(2 ～' + resultSheet.getLastRow() + ')');
  
  if (line < 2 || line > resultSheet.getLastRow()) {
      Browser.msgBox('出力できるデータがありません。');
      return;
  }
  
  var newFileName = createDocument(line);
  Browser.msgBox('Document「' + newFileName + '」を生成しました。');
}

/**
* テンプレートドキュメントコピー後、スプレッドシートの指定行内容で置換
* 行数未指定の場合は最終行として処理
* @param {int} [line]            - スプレッドシートの指定行
* @return {string} [newFileName] - 作成ドキュメントのファイル名
*/
function createDocument(line) {
  if (line === undefined) line = resultSheet.getLastRow();
  var newFileName = getDate('yyyyMMdd') + '_' + getWord(TITLE_WORD_COLUMN_NUM, line); // 作成ドキュメントファイル名を生成
  var newFile = copyDocument(TEMPLATE_DOC_ID, newFileName, SAVE_FOLDER_NAME);   // テンプレートファイルからドキュメントコピーファイルを作成
  replaceDocument(newFile.getId(), line);                                       // 文言を差し替え
  
  return newFileName;
}

/**
* Docファイル新規作成(テンプレートコピー)
* @param {string} [fieldId]     - コピー元ファイルID
* @param {string} [name]        - 作成ファイル名
* @param {string} [destination] - 作成フォルダID
* @return {file}
*/
function copyDocument(fileId, newFileName) {
  var tempDoc = DocumentApp.openById(fileId);
  var newFile  = DocumentApp.create(newFileName);
  
  tempDoc.getBody().getParagraphs().forEach(function(value, i){
    newFile.getBody().insertParagraph(i, value.copy());
  });
  return newFile;
}

/**
* Docファイル新規作成(GoogleDriveコピー)
* @param {string} [fieldId]     - コピー元ファイルID
* @param {string} [name]        - 作成ファイル名
* @param {string} [destination] - 作成フォルダID
* @return {file}
*/
function copyDocumentByDrive(fileId, name, folderId)
{
  var doc = DriveApp.getFileById(fileId);
  var destination = DriveApp.getFolderById(folderId);
  
  if (folderId == undefined) return file.makeCopy(name);
  return file.makeCopy(name, destination);
}

/**
* GoogleDrive上でのファイルコピー(置換文字列は{{{target}}}とする)
* @param {string} [fileId] - 置換対象のドキュメントファイルID
* @param {string} [line]  - 置換対象のスプレッドシートの行数
*/
function replaceDocument(fileId, line){
  var doc = DocumentApp.openById(fileId);
  var body = doc.getBody();
  Logger.log('replaceDocument:LINE = ' + line); 
  
  var headerRow  = resultSheet.getRange(1,1,1,resultSheet.getLastColumn()).getValues();
  var replaceRow = resultSheet.getRange(line,1,1,resultSheet.getLastColumn()).getValues();
    
  for(ColNum in headerRow[0]) {
    if (headerRow[0][ColNum] == "") continue;
    
    var headerStr  = '\\{\\{\\{' + headerRow[0][ColNum] + '\\}\\}\\}';
    var replaceStr = formatValue(replaceRow[0][ColNum]);
    
    body.replaceText(headerStr, replaceStr);
    Logger.log(headerStr + ' → ' + replaceStr);  
  }
}

function formatValue(str) {
  // Date型はyyyy/MM/ddに変更
  if (Object.prototype.toString.call(str) === "[object Date]") {
    str = Utilities.formatDate(str, 'JST', 'yyyy/MM/dd');
  }
  
  return str;
}

/**
* フォーム回答スプレッドシートから特定のセル値を取得
* @param {int} column   - 左から何列目か(1列目の場合は1)
* @return {string} line - 上から何行目か(1行目の場合は1)
*/
function getWord(column, line){
  if (line == undefined) line = resultSheet.getLastRow();
 
  return resultSheet.getRange(line,column).getValue();
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

/**
* メニュー追加用関数
*/
function onOpen() {
  var items = [
    {name : "個別で作成"  , functionName : "createDocumentInteractive"},
  ];
  SpreadsheetApp.getActiveSpreadsheet().addMenu("Docを作成",items);
}


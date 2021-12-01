// 日時を取得
function Getnow() {
  var d = new Date();
  var y = d.getFullYear();
  var mon = d.getMonth() + 1;
  var d2 = d.getDate();
  var h = d.getHours();
  var min = d.getMinutes();
  var s = d.getSeconds();
  var now = y+"/"+mon+"/"+d2+" "+h+":"+min+":"+s;
  return now;
}

// PDFに変換
function CreatePDF() {

let date = Getnow();
let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
let sheet = spreadsheet.getActiveSheet();
let ssid = "1H0FuZu1PrA2jfYkCGKFom09MTuaRUtzFDjJiCDAz-ZE";
let sheetName = spreadsheet.getSheetByName("納品書"); 
// let ss = SpreadsheetApp.openById(ssid);
let shid = sheetName.getSheetId();
let folderId = "1eCb4t5CbryQ0MQPgKft2zkUQKM5s_jM5";
let pdfRange = 'A1%3AD33';

  let url = "https://docs.google.com/spreadsheets/d/" + ssid + "/export?gid=" + shid +  "&exportFormat=pdf&format=pdf";


  let opts = {
    exportFormat: "pdf",
    format:       "pdf",
    size:         "A4",
    portrait:     "true",
    fitw:         "true",
    sheetnames:   "false",
    printtitle:   "false",
    pagenumbers:  "false",
    gridlines:    "false", 
    fzr:          "false",
    range:        pdfRange,
    // gid:          ssid
  };

  let url_ext = [];

  for(optName in opts){
    url_ext.push(optName + "=" + opts[optName]);
  }

  let options = url_ext.join("&");
  let token = ScriptApp.getOAuthToken();

  // let fileName = sheet.getRange('A1').getValue() + '_' + sheet.getRange('D1').getValue() + '.pdf';

  let fileName = date;

  

  let pdf = UrlFetchApp.fetch(url + options, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }).getBlob().setName(fileName + '.pdf');

  // Logger.log(pdf);

  let folder = DriveApp.getFolderById(folderId);

  folder.createFile(pdf);

  // GmailにPDFを送信
  let to = "rikuto1999.54@gmail.com ";
  let subject = "今月の納品書";
  let body = "今月の納品書を作成しました  "; 

  GmailApp.sendEmail(to,
                     subject,
                     body,
                     {attachments: pdf})
  
}

// チャットワークに送信
// const token = "6c8f6abc8ed32205ece36078bfdc8807";
// const roomId = "152708713";

// function CWsend() {
//   let pdffile = CreatePDF();
//   let folder = DriveApp.getFolderById("1eCb4t5CbryQ0MQPgKft2zkUQKM5s_jM5");
//   let file = folder.getFilesByName(pdffile);

//   // let sendtext = Createpdf();
//   let token = "6c8f6abc8ed32205ece36078bfdc8807";
//   let roomId = "152708713";

//   // Logger.log(sendtext);

//   let cw = ChatWorkClient.factory({token:token});
//   cw.sendMessage({
//     room_id: roomId,
//     sendfile: file,
//   });
// }
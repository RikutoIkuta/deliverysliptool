const ssid = "1H0FuZu1PrA2jfYkCGKFom09MTuaRUtzFDjJiCDAz-ZE";
const ss = SpreadsheetApp.openById(ssid);
const shid = ss.getActiveSheet().getSheetId();
const folder = DriveApp.getFolderById("1eCb4t5CbryQ0MQPgKft2zkUQKM5s_jM5");
// const fileName = "納品書テストファイル";

// 納品書シートの取得
// function getText(){
//   let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("納品書");
//   let range = sheet.getRange(1,1,33,4);
//   let values = range.getValues();
//   let body = "";
  
//   for(i = 1; i<33; i++){
//     body += values[i] ;
//   // }

//   // Logger.log(values);
// }



// PDFに変換
function CreatePDF() {

  var url = "https://docs.google.com/spreadsheets/d/SSID/export?".replace("SSID", ssid);


  var opts = {
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
    gid:          ssid
  };

  let url_ext = [];

  for(optName in opts){
    url_ext.push(optName + "=" + opts[optName]);
  }

  let options = url_ext.join("&");
  let token = ScriptApp.getOAuthToken();

  let response = UrlFetchApp.fetch(url + options, {
    headers: {
      "Authorization": "Bearer" + token
    }
  });

  let blob = response.getBlob().setName("テストファイル" +".pdf");

  folder.createFile(blob);
}

// チャットワークスに送信
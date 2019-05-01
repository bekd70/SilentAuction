function getBidSheetsNames(){
   var ss   =  SpreadsheetApp.openById("1YktYIZHyah-ZfUObavpQHENJpOU1v1QMfdcZbz4iR1I");
   var sheetsName = [];
   var sheets = ss.getSheets();
   for( var i = 3; i < sheets.length; i++ ){
     sheetsName.push(sheets[i].getName() );
   };
   return sheetsName;
 }


function clearTestingData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("auctionFormInfo");
  var data = sheet.getDataRange().getValues();
  var formURL = [];
  var sheetNames = getBidSheetsNames();
  //Logger.log(formNames);
  
  for (var i=0; i<sheetNames.length; i++){
    var values = sheetNames[i];
    sheet = ss.getSheetByName(values)
    var formSheet = ss.getSheetByName(values).getFormUrl();
    var formURL = [{}];
    formURL = formSheet.split("forms/d/");
    var formID = formURL[1];
    formID = formID.toString().replace("/viewform","");
    FormApp.openById(formID).removeDestination();
    ss.deleteSheet(sheet);
    DriveApp.getFileById(formID).setTrashed(true);
    
  }
  sheet = ss.getSheetByName("auctionFormInfo");
  sheet.getRange(2, 1,data.length-1,6).clear()
  
}

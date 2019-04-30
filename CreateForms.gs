function onOpen() {
  var menu = [{name: 'Set up Silent Auctions', functionName: 'runScript'}, {name: 'Create Auction Doc', functionName: 'createAuctionDoc'},
              {name: 'Sort Auction Bids by Highest Bid', functionName: 'sortBids'} ];
  SpreadsheetApp.getActive().addMenu('Auctions', menu);
}

function sortSheet(studentFormName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(studentFormName);
  var data = sheet.getDataRange().getValues();
  var range = sheet.getRange("A2:D" + data.length+1);

 // Sorts by the values in column 2 (B)
 range.sort({column: 4, ascending: false});
  
}

function sortBids(){
  var sheetsNames = [];
  sheetsNames = getSheetsNames();
  for (i=3; i<sheetsNames.length; i++){
    sortSheet(sheetsNames[i]);
  }
  
}


/**
* Saves new form information to sheet called auctionFormInfo
* saves url to form and name
* @param {str}    studentFormName
* @param {str}    newFormDestID
* @param {str}    newFormURL
**/
function saveFormInfo(photoID,studentName,artworkTitle,newFormURL,sheetURL,period){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("auctionFormInfo");
  sheet.appendRow([photoID,studentName,artworkTitle,newFormURL, sheetURL, period]); 
}

/**
* gets the names of each individual sheet and save it to 
* an array (sheetNames) and returns array
**/
 function getSheetsNames(){
   var ss   =  SpreadsheetApp.openById("1YktYIZHyah-ZfUObavpQHENJpOU1v1QMfdcZbz4iR1I");
   var sheetsName = [];
   var sheets = ss.getSheets();
   for( var i = 0; i < sheets.length; i++ ){
     sheetsName.push(sheets[i].getName() );
   };
   return sheetsName;
 }

/**
 * renames newly created sheet.
 * the sheet must be named 'Form Responses XX'
 * The sheet is then moved to the end of the list.
 * Auction Setup is then made the active sheet again
 *
 * @param {str}  studentFormName     String from the concatenation of Period_StudentName_Preiod
 * 
 */
function renameSheet(studentFormName){
 
  var ss = SpreadsheetApp.openById("1YktYIZHyah-ZfUObavpQHENJpOU1v1QMfdcZbz4iR1I");
  var sheets =ss.getSheets();
  var pos = ss.getNumSheets();
  //Logger.log(pos)
  var sheetNames = getSheetsNames();
 // Logger.log(sheetNames);
  for (var i = 0; i<pos;i++){
    if (sheetNames[i]) {
      if (sheetNames[i].indexOf('Form Responses') > -1) {
        //Logger.log("Present");
        var first = ss.getSheets()[i];
        ss.setActiveSheet(ss.getSheets()[i]);
        first.setName(studentFormName);
        var newSheetID = first.getSheetId();
        ss.moveActiveSheet(pos);
        ss.setActiveSheet(ss.getSheets()[0]);
        //Logger.log(ss.getActiveSheet().getName());
        return newSheetID;
      
      }
    }
  }
  
}


/**
 * Places file for given item into given folder.
 * If the item is an object that does not support the getId() method or
 * the folder is not a Folder object, an error will be thrown.
 * Also removes file from root directory
 * From: http://stackoverflow.com/a/38042090/1677912
 *
 * @param {Object}  item     Any object that has an ID and is also a Drive File.
 * @param {Folder}  folder   Google Drive Folder object.
 */
function saveItemInFolder(item,folder) {
  var id = item.getId();  // Will throw error if getId() not supported.
  folder.addFile(DriveApp.getFileById(id));
  var temp = DriveApp.getFileById(id);
  DriveApp.getRootFolder().removeFile(temp);
}


/**
* Function to create form for each piece of student artwork
* that is in Google Sheet
* @param  {str}  studentFormName     String from the concatenation of Period_StudentName_Preiod
* @param {array}    values   values pulled from row of sheet to populate form
* @param {str}   photoID    ID of the photo stored on google drive
* @param {obj} ss     spreadsheet to store form data on
**/
function createForm(studentFormName, values, photoID, ss) {

  var form = FormApp.create(studentFormName)
  .setAllowResponseEdits(false)
  .setRequireLogin(false)
  .setTitle("Silent Auction of " + values[2] + " by " + values[1])
  .addEditor('twillcott@aes.ac.in');
  //adds new form data to existing spreadsheet
  form.setDestination(FormApp.DestinationType.SPREADSHEET,ss.getId());
  
  var ssID = form.getDestinationId();
  var formUrl = form.getPublishedUrl();
  var formID = form.getId();
 
  
  var formInfo =[];
  formInfo.push(formUrl);
  formInfo.push(ssID);
  formInfo.push(formID);
  
  var img = DriveApp.getFileById(photoID);
  
  var blob = img.getBlob()
  .getAs('image/jpeg');
  
  form.addImageItem()
  .setWidth(100)
  .setImage(blob)
  .setTitle(values[2] + " by " + values[1]); 

  form.addTextItem()
  .setTitle("Bidder's Name")
  .setRequired(true);
  
  form.addTextItem()
  .setTitle("Email Address")
  .setRequired(true);
 
  form.addTextItem()
  .setTitle("Bid Ammount")
  .setRequired(true);
  
  //saves forms in folder specified and removes
  var folder=DriveApp.getFoldersByName('Silent Auction').next();
  saveItemInFolder(form,folder);
  Utilities.sleep(2000);
  SpreadsheetApp.flush();
  
  return formInfo;
}


function runScript(){
  var ss = SpreadsheetApp.openById("1YktYIZHyah-ZfUObavpQHENJpOU1v1QMfdcZbz4iR1I");
  var sheet = ss.getSheetByName("AuctionSetup");
  var data = sheet.getDataRange().getValues();
  var sheets = ss.getSheets();
  
  //getdata from sheet (AuctionSetup) tied to form
  for (var i=1; i<data.length; i++){
  //for (var i=1; i<3; i++){
    var values = data[i];
    var studentFormName = values[3] + "_" + values[1] + "_" + values[2];
    var studentName = values[1];
    var artworkTitle = values[2]; 
    var urlArray = [{}];
    var photoUrl = values[4];
    var urlArray = photoUrl.split("id=");
    var photoID = urlArray[1];
    var period = values[3];
    photoID = photoID.toString().replace("\"","");
    
    //create the form and return id and url of form into formInfo
    //formInfo[0] is URL of form
    //formInfo[1] is ssID
    //formInfo[2] is form destination id
    var formInfo = createForm(studentFormName, values, photoID, ss);
    
    //rename newly created form to the studentFormName
    //return value is id of new sheet
    var newSheetID = renameSheet(studentFormName);
    
    //save information to be added to auctionFormInfo sheet
    var newFormURL = formInfo[0];
    var ssID = formInfo[1];
    var newFormID = formInfo[2];
    var sheetURL = "https://docs.google.com/spreadsheets/d/1YktYIZHyah-ZfUObavpQHENJpOU1v1QMfdcZbz4iR1I/edit#gid=" + newSheetID
    
    saveFormInfo(photoID,studentName,artworkTitle,newFormURL,sheetURL,period);
    
  }
  
  var SORT_DATA_RANGE = "A2:F" + data.length+1;
  var SORT_ORDER = [
    {column: 6, ascending: true},  // 5 = period column, sort by ascending order 
    {column: 2, ascending: true} // 2 = Student Name column number, sort by ascending order 
  ];
  ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet = ss.getSheetByName("auctionFormInfo");
  var range = sheet.getRange(SORT_DATA_RANGE);
  range.sort(SORT_ORDER);
  
  
}
function createAuctionDoc(){
  
  var headerStyle = {};  
  headerStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#336600';  
  headerStyle[DocumentApp.Attribute.BOLD] = true;  
  headerStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#FFFFFF';
  
  var cellStyle = {};
  cellStyle[DocumentApp.Attribute.BOLD] = false;  
  cellStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
  
  var paraStyle = {};
  paraStyle[DocumentApp.Attribute.SPACING_AFTER] = 0;
  paraStyle[DocumentApp.Attribute.LINE_SPACING] = 1;
  
  var folder=DriveApp.getFoldersByName('Silent Auction').next();
  var ss = SpreadsheetApp.openById("1YktYIZHyah-ZfUObavpQHENJpOU1v1QMfdcZbz4iR1I");
  var sheet = ss.getSheetByName("auctionFormInfo");
  var data = sheet.getDataRange().getValues();
  var doc = DocumentApp.create('Silent Auction Links');
  var body = doc.getBody();
  var rowsData = ['Photo', 'Artwork By','Class Period', 'Artwork Title', 'Link to Artwork Auction'];
  body.insertParagraph(0, "Silent Auction Links")
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  table = body.appendTable();
  var tr = table.appendTableRow();
  
  //create header row
  for (var i=0; i<rowsData.length; i++){
    var td = tr.appendTableCell(rowsData[i]);
    td.setAttributes(headerStyle);
  }
  
  //create one row for each peice of artwork in AuctionInfo sheet
  for (var i=1; i<data.length; i++){
    var tr = table.appendTableRow();
    var rowsData = data[i];
    
    //inserts photo of artwork
    var photoBlob   = DriveApp.getFileById(rowsData[0]).getBlob();
    var td = tr.appendTableCell().appendImage(photoBlob).setWidth("100").setHeight("75");
    
    //inserts student name
    td = tr.appendTableCell(rowsData[1]);
    td.setAttributes(cellStyle); 
    
    //insert period information
    td = tr.appendTableCell(rowsData[5]);
    td.setAttributes(cellStyle);
    
    //inserts title of artwork
    td = tr.appendTableCell(rowsData[2]);
    td.setAttributes(cellStyle);
    
    //inserts link to auction form
    td = tr.appendTableCell().editAsText().insertText(0, "Silent Auction link for " + rowsData[2] + " by " + rowsData[1]).setLinkUrl(rowsData[3]);
    td.setAttributes(cellStyle);
      
  }
    
    
  
  
  
  
  saveItemInFolder(doc,folder);
}

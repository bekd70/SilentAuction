function onOpen() {
  var menu = [{name: 'Set up Silent Auctions', functionName: 'runScript'}];
  SpreadsheetApp.getActive().addMenu('Auctions', menu);
}

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
  Logger.log(pos)
  var sheetNames = getSheetsNames();
  Logger.log(sheetNames);
  for (var i = 0; i<pos;i++){
    if (sheetNames[i]) {
      if (sheetNames[i].indexOf('Form Responses') > -1) {
        Logger.log("Present");
        var first = ss.getSheets()[i];
        ss.setActiveSheet(ss.getSheets()[i]);
        first.setName(studentFormName);
        var newSheetID = first.getSheetId();
        ss.moveActiveSheet(pos);
        ss.setActiveSheet(ss.getSheets()[0]);
        Logger.log(ss.getActiveSheet().getName());
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
  
  var formID = form.getId();
  var formUrl = form.getPublishedUrl();
  
  var formInfo =[];
  formInfo.push(formUrl);
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
  
  return formInfo;
}

function runScript(){
  var ss = SpreadsheetApp.openById("1YktYIZHyah-ZfUObavpQHENJpOU1v1QMfdcZbz4iR1I");
  var sheet = ss.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var sheets = ss.getSheets();
  
  for (var i=1; i<data.length; i++){
    var values = data[i];
    Logger.log(values);
    var studentFormName = values[3] + "_" + values[1] + "_" + values[2];
    var urlArray = [{}];
    var url = values[4];
    var urlArray = url.split("id=");
    var photoID = urlArray[1];
    photoID = photoID.toString().replace("\"","")
    Logger.log(studentFormName);
    
    //create the form and return id and url of form into formInfo
    //formInfo[0] is ID of form
    //formInfo[1] is url
    var formInfo = createForm(studentFormName, values, photoID, ss);
    Logger.log(formInfo);
    
    //rename and move new sheet added by above function
    renameSheet(studentFormName);
    
    //rename newly created form to the studentFormName
    //return value is id of new sheet
    var newSheetID = renameSheet(studentFormName);
    Logger.log(newSheetID);
  }
}


// global variables
var timeZone = "EST";
var masterIncomingFolderId = "1sp2QzTccc7wJR6l-CQr-5D2S-MNAb4NU"
var clearlistMainFolderId = "1eGUYdii_6IOE4hCykEjFw1jDfTqha7cw";
var clearlistFilePattern = "^CLEAR.2021" + "[0-9]{4}" + ".csv";
var clearlistTradesImportSheet = "CL Trade Create";
var sharenettMainFolderId = "1mQLc12L--kCPE4ICZ7upfMy4XFzgkVN1";
var sharenettFilePattern = "^SHARE.2021" + "[0-9]{4}" + ".csv";
var sharenettTradesImportSheet = "SN Trade Create";
var rangeInTab = "B:B";
var startColInTab = 2;


// general function which can be reused
function getLastRow(shName, range) {
  Logger.log("Start function: getLastRow of " + shName)
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(shName)
  var Avals = ss.getRange(range).getValues();
  var Alast = Avals.filter(String).length;
  Logger.log("End function: getLastRow")
  return Alast
}


function createFolder(folderID, folderName) {
  var parentFolder = DriveApp.getFolderById(folderID);
  var subFolders = parentFolder.getFolders();
  var doesntExists = true;
  var newFolder = '';
  // Check if folder already exists.
  while (subFolders.hasNext()) {
    var folder = subFolders.next();
    //If the name exists return the id of the folder
    if (folder.getName() === folderName) {
      doesntExists = false;
      newFolder = folder;
      return newFolder.getId();
    };
  };
  //If the name doesn't exists, then create a new folder
  if (doesntExists == true) {
    //If the file doesn't exists
    newFolder = parentFolder.createFolder(folderName);
    return newFolder.getId();
  };
};

function writeDataToSheet(writeToSheetName,rangeInTab, startColInTab, dataToWrite) {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(writeToSheetName);
  var last_rows = getLastRow(writeToSheetName, rangeInTab);
  ss.getRange(last_rows + 1, startColInTab, dataToWrite.length, dataToWrite[0].length).setValues(dataToWrite);
}


function importFromCSV(masterIncomingFolderId, mainAtsFolderID, importFilePattern, writeToSheetName, rangeInTab, startColInTab) {
  var mainFolder = DriveApp.getFolderById(masterIncomingFolderId);
  var f = mainFolder.getFiles();
  var blankfile = [];

  // ATS -> ATS archive -> Empty Files Folder / yyyy-MM -> MM-dd-yyyy [MODIFIED: WHEN WE ADD NEW ATS]
  var atsname = writeToSheetName.substring(0, 2);
  Logger.log(atsname);
  archiveFolderID = createFolder(mainAtsFolderID, atsname+"_Archive_Trades");
  emptyFolderID = createFolder(archiveFolderID, atsname+"_Empty_Files");

  var monthfolder = Utilities.formatDate(new Date(), timeZone, "yyyy-MM");
  var monthfolderid = createFolder(archiveFolderID, monthfolder);
  var todaydatefolder = Utilities.formatDate(new Date(), timeZone, "MM-dd-yyyy");
  var todaydatefolderid = createFolder(monthfolderid, todaydatefolder);
  var destfolder = DriveApp.getFolderById(todaydatefolderid);
  var emptyfolder = DriveApp.getFolderById(emptyFolderID);

  while (f.hasNext()) {
    var file = f.next();
    var regExp = new RegExp(importFilePattern)

    if (file.getName().search(regExp) != -1) {
      name = file.getName();
      try {
        var contents = Utilities.parseCsv(file.getBlob().getDataAsString());
        var header = contents.shift(); // remove header of the files
        writeDataToSheet(writeToSheetName, rangeInTab, startColInTab, contents);
        file.moveTo(destfolder);
      } catch (err) {
        Logger.log(err);
        blankfile.push(name);
        file.moveTo(emptyfolder);
      }
    }
  };
  return blankfile;
}




// use general function code
function importTradesCL(){
  importFromCSV(masterIncomingFolderId, clearlistMainFolderId, clearlistFilePattern, clearlistTradesImportSheet, rangeInTab, startColInTab)
}

function importTradesSN(){
  importFromCSV(masterIncomingFolderId, sharenettMainFolderId, sharenettFilePattern, sharenettTradesImportSheet, rangeInTab, startColInTab)
}


// combine different functions code

function importAllTrade(){
  importTradesCL()
  importTradesSN()
}







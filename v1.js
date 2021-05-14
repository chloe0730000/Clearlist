// global variables
var masterIncomingFolderId = "1sp2QzTccc7wJR6l-CQr-5D2S-MNAb4NU"
var clearlistMainFolderId = "1R6qpWVFjziwHMWfCvJOu9jqJwOrmFhJ1";
var clearlistFilePattern = "^CLEAR.2021" + "[0-9]{4}" + ".csv";
var clearlistTradesImportSheet = "CL Trade Create";
var sharenettMainFolderId = "1Wa2gaF_DAepjRVyX9GIXGNISnPHyEk4N";
var sharenettFilePattern = "^SHARE.2021" + "[0-9]{4}" + ".csv";
var sharenettTradesImportSheet = "SN Trade Create";



// general function which can be reused
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


function importFromCSV(masterIncomingFolderId, mainAtsFolderID, importFilePattern, writeToSheetName) {
  var mainFolder = DriveApp.getFolderById(masterIncomingFolderId);
  var f = mainFolder.getFiles();
  var blankfile = [];

  // ATS -> ATS archive -> Empty Files Folder / yyyy-MM -> MM-dd-yyyy [MODIFIED: WHEN WE ADD NEW ATS]
  if (importFilePattern.search("CLEAR") != -1){
    archiveFolderID = createFolder(mainAtsFolderID, "CL_Archive_Trades");
    emptyFolderID = createFolder(archiveFolderID, "CL_Empty_Files");
  }else if(importFilePattern.search("SHARE")!=-1){
    archiveFolderID = createFolder(mainAtsFolderID, "SN_Archive_Trades");
    emptyFolderID = createFolder(archiveFolderID, "SN_Empty_Files");
  }

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
        writeDataToSheet(writeToSheetName, "B:B", 2, contents);
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
  importFromCSV(masterIncomingFolderId, clearlistMainFolderId, clearlistFilePattern, clearlistTradesImportSheet)
}
function importTradesSN(){
  importFromCSV(masterIncomingFolderId, sharenettMainFolderId, sharenettFilePattern, sharenettTradesImportSheet)
}



// combine different functions code

function importAllTrade(){
  importTradesCL()
  importTradesSN()
}








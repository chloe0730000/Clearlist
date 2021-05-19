
// global variables
var timeZone = "EST";

var masterIncomingFolderId = "1sp2QzTccc7wJR6l-CQr-5D2S-MNAb4NU"

var clearlistMainFolderId = "1eGUYdii_6IOE4hCykEjFw1jDfTqha7cw";
var clearlistTradeArchiveFolderId = "1vvU-7euhVR8kFoaQNUNMpfNqxO9nUrNg";
var clearlistLifecycleArchiveFolderId = "1TPtAX0yeAKpm0F1ald_7QhcVnObrL83f";
var clearlistTacArchiveFolderId = "1CxrW1HlBR-WpPihqDx5zv281G2AI1jtE";
var sharenettMainFolderId = "1mQLc12L--kCPE4ICZ7upfMy4XFzgkVN1";
var sharenettTradeArchiveFolderId = "1Gw_p-h9_jAgiZ_v3uxZZySVbLvtP8LLN";
var sharenettLifecycleArchiveFolderId = "11Oilu8FRPi1C9M7iehswexX0F6oPMfnW";
var sharenettTacArchiveFolderId = "16fZDLgdPQsY8kcwSUr9sqU6yliUg8iJy";

var clearlistFilePattern = "^CLEAR.2021" + "[0-9]{4}" + ".csv";
var clearlistLifecyclePattern =  "^CLEAR.2021" + "[0-9]{4}" + "_LIFECYCLE.csv";
var clearlistTacPattern = "tac_file_clearlist";
var sharenettFilePattern = "^SHARE.2021" + "[0-9]{4}" + ".csv";
var sharenettLifecyclePattern =  "^SHARE.2021" + "[0-9]{4}" + "_LIFECYCLE.csv";
var sharenettTacPattern = "tac_file_sharenett";

var clearlistTradesImportSheet = "CL Trade Create";
var clearlistLifecycleImportSheet = "LIFECYCLE";
var clearlistTacImportSheet = "TAC";
var sharenettTradesImportSheet = "SN Trade Create";
var sharenettLifecycleImportSheet = "LIFECYCLE";
var sharenettTacImportSheet = "TAC";

var rangeInTradeTab = "B:B";
var startColInTradeTab = 2;
var rangeInLifecycleTab = "A:A";
var startColInLifecycleTab = 1;
var rangeInTacTab = "A:A";
var startColInTacTab = 1;


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


function importFromCSV(masterIncomingFolderId, mainAtsFolderID, archiveAtsFolderID, importFilePattern, writeToSheetName, rangeInTab, startColInTab) {
  var mainFolder = DriveApp.getFolderById(masterIncomingFolderId);
  var f = mainFolder.getFiles();
  var blankfile = [];

  // ATS -> ATS archive -> Empty Files Folder / yyyy-MM -> MM-dd-yyyy
  emptyFolderID = createFolder(archiveAtsFolderID, "Empty_Files");

  var monthfolder = Utilities.formatDate(new Date(), timeZone, "yyyy-MM");
  var monthfolderid = createFolder(archiveAtsFolderID, monthfolder);
  var todaydatefolder = Utilities.formatDate(new Date(), timeZone, "MM-dd-yyyy");
  var todaydatefolderid = createFolder(monthfolderid, todaydatefolder);
  var destfolder = DriveApp.getFolderById(todaydatefolderid);
  var emptyfolder = DriveApp.getFolderById(emptyFolderID);

  while (f.hasNext()) {
    var file = f.next();
    var regExp = new RegExp(importFilePattern)

    if (file.getName().search(regExp) != -1) {
      name = file.getName();
      Logger.log(name);
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
  importFromCSV(masterIncomingFolderId, clearlistMainFolderId, clearlistTradeArchiveFolderId ,clearlistFilePattern, clearlistTradesImportSheet, rangeInTradeTab, startColInTradeTab)
}

function importTradesSN(){
  importFromCSV(masterIncomingFolderId, sharenettMainFolderId, sharenettTradeArchiveFolderId,sharenettFilePattern, sharenettTradesImportSheet, rangeInTradeTab, startColInTradeTab)
}

function importLifecycleCL(){
  importFromCSV(masterIncomingFolderId, clearlistMainFolderId, clearlistLifecycleArchiveFolderId, clearlistLifecyclePattern, clearlistLifecycleImportSheet, rangeInLifecycleTab, startColInLifecycleTab)
}

function importLifecycleSN(){
  importFromCSV(masterIncomingFolderId, sharenettMainFolderId, sharenettLifecycleArchiveFolderId, sharenettLifecyclePattern, sharenettLifecycleImportSheet, rangeInLifecycleTab, startColInLifecycleTab)
}

function importTacCL(){
  importFromCSV(masterIncomingFolderId, clearlistMainFolderId, clearlistTacArchiveFolderId, clearlistTacPattern, clearlistTacImportSheet, rangeInTacTab, startColInTacTab)
}

function importTacSN(){
  importFromCSV(masterIncomingFolderId, sharenettMainFolderId, sharenettTacArchiveFolderId, sharenettTacPattern, sharenettTacImportSheet, rangeInTacTab, startColInTacTab)
}

// combine different functions code

function importAllTrade(){
  importTradesCL()
  importTradesSN()
}


function importAllLifecycle(){
  importLifecycleCL()
  importLifecycleSN()
}

function importAllTac(){
  importTacCL()
  importTacSN()
}

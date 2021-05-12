// global variables

var clearlist_main_folder_id = "1sp2QzTccc7wJR6l-CQr-5D2S-MNAb4NU";
var clearlist_trade_archive_folder_id = "18vPPMxZPTIEeXJ-rruva7KAcVy7aXZ37";
var clearlist_trade_empty_folder_id = "1A8C1TJhtjt-hi_4lMAVJTWi2vMYIuCFu";
var clearlist_file_pattern = "^CLEAR.2021" + "[0-9]{4}" + ".csv";
var clearlist_tradesImportSheet = "CL Trade Create";



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

function writeDataToSheet(sheet_to_write, rangeInTab, dataToWrite) {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_to_write);
  var last_rows = getLastRow(sheet_to_write, rangeInTab);
  ss.getRange(last_rows + 1, 2, dataToWrite.length, dataToWrite[0].length).setValues(dataToWrite);
}


function importFromCSV(mainFolderID, archiveFolderID, emptyFolderID, import_file_pattern, sheet_to_write) {

  var mainFolder = DriveApp.getFolderById(mainFolderID);
  var f = mainFolder.getFiles();
  var blankfile = [];

  // create today date folder in Archive_Of_Trade_Files
  var todaydatefolder = Utilities.formatDate(new Date(), timeZone, "MM-dd-yyyy");
  var todaydatefolderid = createFolder(archiveFolderID, todaydatefolder);
  var dest_folder = DriveApp.getFolderById(todaydatefolderid);

  while (f.hasNext()) {
    var file = f.next();
    var regExp = new RegExp(import_file_pattern)

    if (file.getName().search(regExp) != -1) {
      name = file.getName();
      Logger.log(name);
      try {
        var contents = Utilities.parseCsv(file.getBlob().getDataAsString());
        var header = contents.shift(); // remove header of the files
        writeDataToSheet(sheet_to_write, "B:B", contents);
        file.moveTo(dest_folder);
      } catch (err) {
        Logger.log(err);
        blankfile.push(name);
        file.moveTo(emptyFolderID);
      }
    }
  };
}


// use general function code
function importTradesCL(){
  importFromCSV(clearlist_main_folder_id, clearlist_trade_archive_folder_id, clearlist_trade_empty_folder_id, clearlist_file_pattern, clearlist_tradesImportSheet)
}

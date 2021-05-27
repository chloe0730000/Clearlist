
// global variables
var timeZone = "EST";

var bauFolderId = "1s132fsm3mrJX47MLEBzGdVtcCWKLtjbt";
var masterIncomingFolderId = "1sp2QzTccc7wJR6l-CQr-5D2S-MNAb4NU"
var masterOutgoingArchiveFolderId = "12zP8IamGqA5L_Kf-NpfzUXynmhlabvwu";

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


var todaytradeoutputFolderName = "Todays_Trades_Export";
var todayTradeRangePending = "B2:AH";
var todayTradeRangeSettled = "B2:AX";
var todayTradeOutputRange = "B2:R";
var todayTradeOutputPendingFilter = ["PENDING", "SENT"];
var todayTradeOutputSettledFilter = ["SETTLED", "YES",""];
var todayTradeOutputColPendingFilter = [0,33];
var todayTradeOutputColSettledFilter = [0,43,44];
var clearlistTradesLedger = "CL Todays Trades";
var sharenettTradesLedger = "SN Todays Trades";
var todayTradeInsertValueColPending = "AI";
var todayTradeInsertValueColSettled = "AT";



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


// trade create
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


// today trade make as csv


function convertToCSV(ss, totalRows, todayTradeRange, outputTradeRange, todayTradeOutputFilter, todayTradeOutputColFilter, todayTradeInsertValueCol) {
  //var totalRows = ss.getLastRow()
  var totalRows = totalRows + 1; // add first row back 

  var notation = todayTradeRange + totalRows
  var notation2 = outputTradeRange + totalRows
  var data = ss.getRange(notation).getValues()
  var data2 = ss.getRange(notation2).getValues()
  // get available data range in the spreadsheet

  if (typeof todayTradeOutputFilter==="undefined"){
    Logger.log("Without Filter");
    try {
      var csvFile = undefined;
      if (data2.length > 1) {
        var csv = "";
        for (var row = 0; row < data2.length; row++) {
              if (row < data2.length - 1) {
                csv += data2[row].join(",") + "\r\n";
              }
              else {
                csv += data2[row];
                Logger.log("Adding row to CSV")
              }
        }
        csvFile = csv;
      }
      return csvFile;
    }
    catch (err) {
      Logger.log(err);
    }
  }else if (todayTradeOutputFilter.length==2){
    Logger.log("With 2 Filters for pending");
    try {
      //var data = activeRange.getValues();
      var csvFile = undefined;

      // loop through the data in the range and build a string with the csv data
      if (data.length > 1) {
        var csv = "";
        for (var row = 0; row < data.length; row++) {
          //Logger.log("data row "+data[row][0])
          // PROCESSING used to say NEW
          if (data[row][todayTradeOutputColFilter[0]] == todayTradeOutputFilter[0] || data[row][0] == "Transaction Type") {
            if (data[row][todayTradeOutputColFilter[1]] != todayTradeOutputFilter[1]) {
              var change_row_number = row + 2;

              if (row < data2.length - 1) {
                csv += data2[row].join(",") + "\r\n";
              }
              else {
                csv += data2[row];
                Logger.log("Adding row to CSV")
              }
              if (change_row_number != 2) {
                ss.getRange(todayTradeInsertValueCol + change_row_number).setValue("SENT");
              }
            }
          }
        }
        csvFile = csv;
      }
      return csvFile;
    }
    catch (err) {
      Logger.log(err);
      Browser.msgBox(err);
    }
  }else if(todayTradeOutputFilter.length==3){
    Logger.log("With 3 filters for settled");
      try {
      var csvFile = undefined;

      if (data.length > 1) {
        var csv = "";
        for (var row = 0; row < data.length; row++) {
          if (data[row][todayTradeOutputColFilter[0]] == todayTradeOutputFilter[0] || data[row][0] == "Transaction Type") {
            csv += data2[row].join(",") + "\r\n";
            Logger.log("Add title row")
            if (data[row][todayTradeOutputColFilter[1]] == todayTradeOutputFilter[1]) {
              if(data[row][todayTradeOutputColFilter[2]] == todayTradeOutputFilter[2]){
                  var change_row_number = row + 2;

                  if (row < data2.length - 1) {
                    csv += data2[row].join(",") + "\r\n";
                  }
                  else {
                    csv += data2[row];
                    Logger.log("Adding row to CSV")
                  }
                  if (change_row_number != 2) {
                    ss.getRange(todayTradeInsertValueCol + change_row_number).setValue("SENT");
                  }
                }
            }
          }
        }
        csvFile = csv;
      }
      return csvFile;
    }
    catch (err) {
      Logger.log(err);
      Browser.msgBox(err);
    }

  }
}

function convertToCSVandCreateFilesToFolders(fileToConvertCsv, rangeInTab, fileOutputFolderId1, fileOutputFolder1Name, fileOutputFolderId2, ledgerRange, ledgerOutputRange, ledgerOutputFilter,ledgerOutputColFilter, ledgerInsertValueCol) {

  Logger.log("Start function convertTodaysTradeIntoCSVWithNEWTradesOnly")
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(fileToConvertCsv);
  var totalRows = getLastRow(fileToConvertCsv, rangeInTab);

  var timeZone = "EST";
  var monthfolder = Utilities.formatDate(new Date(), timeZone, "yyyy-MM");
  var monthfolderid = createFolder(fileOutputFolderId1, monthfolder);
  var todaydatefolder = Utilities.formatDate(new Date(), timeZone, "MM-dd-yyyy");
  var todaydatefolderid = createFolder(monthfolderid, todaydatefolder);
  var todaysPendingTradesFolderID = createFolder(todaydatefolderid, fileOutputFolder1Name);

  var dest_folder = DriveApp.getFolderById(todaysPendingTradesFolderID);
  var clearlist_outgoing_folder = DriveApp.getFolderById(fileOutputFolderId2);
  
  var dateFormatted = Utilities.formatDate(new Date(), timeZone, "yyyyMMdd");
  var d = new Date();
  var currentTime = d.getHours();
  
  // convert all available sheet data to csv format
  var csvFile = convertToCSV(ss,totalRows, ledgerRange, ledgerOutputRange, ledgerOutputFilter, ledgerOutputColFilter, ledgerInsertValueCol);
  // create a file in the Docs List with the given name and the csv data
  var atsName = ss.getName().split(" ")[0];
  var outputFileName = ss.getName().replace(atsName,'').replace(" ",'').replace(" ",'');
  
  try{
    var fileName = atsName+"_"+outputFileName+ "_" + ledgerOutputFilter[0] + "_" + dateFormatted + "_" + currentTime + ".csv";
  } catch (err) {
    var fileName = atsName+"_"+outputFileName+ "_" + dateFormatted + "_" + currentTime + ".csv";
    Logger.log(err);
  }

  
  var file = dest_folder.createFile(fileName, csvFile);
  var file_output2 = clearlist_outgoing_folder.createFile(fileName, csvFile);
  Logger.log("End function convertTodaysTradeIntoCSVWithNEWTradesOnly")
  return fileName;
}



// use general function code

// trade create part
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


// today trade part

function downloadPendingTradesCSVCL(){
  convertToCSVandCreateFilesToFolders(clearlistTradesLedger, rangeInTradeTab, bauFolderId, todaytradeoutputFolderName, masterOutgoingArchiveFolderId,todayTradeRangePending,todayTradeOutputRange,todayTradeOutputPendingFilter, todayTradeOutputColPendingFilter, todayTradeInsertValueColPending)
}

function downloadPendingTradesCSVSN(){
  convertToCSVandCreateFilesToFolders(sharenettTradesLedger, rangeInTradeTab, bauFolderId, todaytradeoutputFolderName, masterOutgoingArchiveFolderId,todayTradeRangePending,todayTradeOutputRange,todayTradeOutputPendingFilter, todayTradeOutputColPendingFilter, todayTradeInsertValueColPending)
}

function downloadSettledTradesCSVCL(){
  convertToCSVandCreateFilesToFolders(clearlistTradesLedger, rangeInTradeTab, bauFolderId, todaytradeoutputFolderName, masterOutgoingArchiveFolderId,todayTradeRangeSettled,todayTradeOutputRange,todayTradeOutputSettledFilter, todayTradeOutputColSettledFilter, todayTradeInsertValueColSettled)
}

function downloadSettledTradesCSVSN(){
  convertToCSVandCreateFilesToFolders(sharenettTradesLedger, rangeInTradeTab, bauFolderId, todaytradeoutputFolderName, masterOutgoingArchiveFolderId,todayTradeRangeSettled,todayTradeOutputRange,todayTradeOutputSettledFilter, todayTradeOutputColSettledFilter, todayTradeInsertValueColSettled)
}


function nofilter(){
  convertToCSVandCreateFilesToFolders(clearlistTradesLedger, rangeInTradeTab, bauFolderId, todaytradeoutputFolderName, masterOutgoingArchiveFolderId,todayTradeRange,todayTradeOutputRange)
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


var omnibusOnboardingImportSheet="Omnibus Account Onboarding";
var masterBalanceSheet = "MASTER BALANCES";
var omnibusOnboardingFilter = ["YES"];
var omnibusOnboardingColFilter = [12,9, 5];
var omnibusOnboardingInputRange = "B2:P";
var startColInMasterBalance = 2;
var rangeInMasterBalanceTab = "B:B";
var numberColToFillMasterBalance = 3;

function test6(){
  onboardingToMasterBalance(omnibusOnboardingImportSheet, masterBalanceSheet, rangeInMasterBalanceTab, omnibusOnboardingInputRange, omnibusOnboardingColFilter, omnibusOnboardingFilter, startColInMasterBalance, numberColToFillMasterBalance)
}


function onboardingToMasterBalance(importSheet, outputSheet, rangeInTab, onboardingInputRange, onboardingColFilter, onboardingFilter, startColInOutputSheet,numberColToFillOutputSheet) {
  
  // add first row back
  var importSheetTotalRows = getLastRow(importSheet, rangeInTab)+1;
  var outputSheetTotalRows = getLastRow(outputSheet, rangeInTab)+1;

  var notation = onboardingInputRange + importSheetTotalRows;
  var importss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(importSheet);
  var outputss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(outputSheet);
  var data = importss.getRange(notation).getValues();
    try {
        for (var row = 0; row < data.length; row++) {
          if (data[row][onboardingColFilter[0]] == onboardingFilter[0] ) {
            var customerName = data[row][onboardingColFilter[1]];
            var brokerDealerId = data[row][onboardingColFilter[2]];    
            Logger.log("name: "+ customerName + " Broker: "+brokerDealerId);
            outputss.getRange(outputSheetTotalRows, startColInOutputSheet,1, numberColToFillOutputSheet).setValues([["OK", customerName, brokerDealerId]]);
            outputss.getRange(outputSheetTotalRows+1, startColInOutputSheet,1, numberColToFillOutputSheet).setValues([["OK", "Holding_"+customerName, brokerDealerId]]);
            outputSheetTotalRows+=2;
          }
      }
    }
    catch (err) {
      Logger.log(err);
      Browser.msgBox(err);
    }
}






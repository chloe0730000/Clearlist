//////////////////////////////////////////////////////////////////////////////////////////////////////////
//GLOBAL VARIABLES
//////////////////////////////////////////////////////////////////////////////////////////////////////////

//used to limit the # of trades per batch (Google App Script code cannot run for longer than 30 mins)
var maxTradesPerBatch = 50

//These accounts collect fees when trades get settled
var clearlistID = "CLER";
var paxosID = "PAXOS";
var sharenettID = "SHNT"

var tradesImportSheet = "Trade Create";
var tradesFile = "CLEAR_TRD.csv";
var newTrades = "Trade Create";
var tradesLedgerNoLookup = "NoLookup"
var tradesLedger = "Todays Trades";
var rangeToClearInTradeCreate = 'B3:R100';
var ackStatus = "ACK"; //not used 
var timeZone = "EST";
var masterBalancesSheet = "MASTER BALANCES";
var securityOnboardSheet = "Securities Onboarding";
var balancesHistory = "Balances History";
var tradingHistory = "Trading History";
var rangeToClearInTodaysTrades = 'B3:R500';
var cashCreate = "Cash Create";
var secCreate = "Sec Create";
var cashRedeem = "Cash Redeem";
var secRedeem = "Sec Redeem";
var c_secRedeem = "Sec Redeem";
var GTSBalances = "GTS_Balances";
var securitiesOnboarding = "Securities Onboarding";
var clearlistBalancesTab = "Clearlist_Balances";
var customerOnboarding = "Customer Onboarding";
var brokerDealerOnboarding = "Broker Dealer Onboarding";

var privateSecuritiesOpsEmail = "privatesecuritiesops@paxos.com"

//these variables are used for trade processing & settling. The col numbers represent col numbers in Todays Trades (assuming the counting starts at Col B, where B=0)
var tradeStatusColNum = 0
var tradeStatusColumnLetter = "B"

var paxosRowInMBIndexColNum = 19;
var clearlist51424RowInMBIndexColNum = 38;


//buyer pointers
//the col numbers represent the columns in Todays Trades. The contents of the cells indicate the row number or column number of a variable in master balances 
var buyerNetNotionalColNum = 10;
var buyerIDCol = 2;
var buyerRowInMBIndexColNum = 21; //not used anyomore
var buyerHoldingRowInMBIndexColNum = 23; //not used anyomore
var buyerBDIDColNumInTodaysTrades = 3;
var buyerBDFeeColNum = 13;
var buyerBDRowInMBIndexColNum = 26; //not used anyomore
var ATSBuyerFeeColNum = 9;
var priceColNum = 6;
var quantityColNum = 7;


//seller pointers
var sellerNetNotionalColNum = 11;
var sellerIDCol = 4;
var sellerRowInMBIndexColNum = 22; //not used anyomore
var sellerColInMBofSecurityColNum = 27; //still used
var sellerSecurityQuantityColNum = 7;
var sellerHoldingRowInMBIndexColNum = 24; //not used anyomore
var sellerBDRowInMBIndexColNum = 25; //not used anyomore
var sellerBDIDColNumInTodaysTrades = 5;
var sellerBDFeeColNum = 14;
var ATSSellerFeeColNum = 15;

var assetCUSIPColNum = 8;


var tradeIDColNum = 12;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//New variables for SN & CL tabs
var tradeCreateCL = "CL Trade Create"
var todaysTradesCL = "CL Todays Trades"

var tradeCreateSN = "SN Trade Create"
var todaysTradesSN = "SN Todays Trades"

var tradingHistoryCL = "CL Trading History"
var tradingHistorySN = "SN Trading History"



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//TABS
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

//BALANCES

//Master Balances tab
  var ssMB = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(masterBalancesSheet);
  var totalRowsMB = getLastRow(masterBalancesSheet, 'B:B');
  //don't think this is used anywhere. also everywhere else we do B1 not B2
  var notationMB = "B2:Z" + totalRowsMB + 1

//Balances History tab
  var ssBalancesHistory = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(balancesHistory);
  //var totalRowsBalancesHistory = ssBalancesHistory.getLastRow();  

//ONBOARDING TABS

//Customer Onboarding tab details (used to get Customer email)
  var ssCustOnboarding = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(customerOnboarding);
  var notationCO = 'B3:AC';
  var dataCustomerOnboarding = ssCustOnboarding.getRange(notationCO).getValues();
  var customerOnboardingTabCustomerTrellisIDColNum = 3;
  var customerOnboardingTabCustomerEmailColNum = 2;
  var customerOnboardingTabCustomerCashAccountColNum = 25
  var customerOnboardingTabCustomerHoldingCashAccountColNum = 26





//FUNCTIONS
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//GET LAST ROW OF ANY TAB, BASED ON SPECIFIC COLUMN
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//used by multiple functions. We get the last row in order to get the range &/ continue populating the tab after the last row


function getLastRow(shName, range) {
  Logger.log("Start function: getLastRow of " + shName)
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(shName)
  var Avals = ss.getRange(range).getValues();
  var Alast = Avals.filter(String).length;
  Logger.log("End function: getLastRow")
  return Alast
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//DISPLAY ALERT
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

//displays alert
function displayToastAlert(message) {
  SpreadsheetApp.getActive().toast(message, "⚠️ Alert");
}


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//IMPORTING TRADE FILES FROM GOOGLE DRIVE INTO TRADE CREATE TAB 
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

//imports trades from Clearlist GD + moves the trade file to Archive 

var timeZone = "EST";

var bauFolderId = "1s132fsm3mrJX47MLEBzGdVtcCWKLtjbt";
var masterIncomingFolderId = "1sp2QzTccc7wJR6l-CQr-5D2S-MNAb4NU"
var masterOutgoingFolderId = "1Myehii1D3H_sUrvgtuV-yZegp9I1-ru7";

var clearlistMainFolderId = "1eGUYdii_6IOE4hCykEjFw1jDfTqha7cw";
var clearlistTradeArchiveFolderId = "1vvU-7euhVR8kFoaQNUNMpfNqxO9nUrNg";
var clearlistLifecycleArchiveFolderId = "1TPtAX0yeAKpm0F1ald_7QhcVnObrL83f";
var clearlistTacArchiveFolderId = "1CxrW1HlBR-WpPihqDx5zv281G2AI1jtE";
var sharenettMainFolderId = "1mQLc12L--kCPE4ICZ7upfMy4XFzgkVN1";
var sharenettTradeArchiveFolderId = "1Gw_p-h9_jAgiZ_v3uxZZySVbLvtP8LLN";
var sharenettLifecycleArchiveFolderId = "11Oilu8FRPi1C9M7iehswexX0F6oPMfnW";
var sharenettTacArchiveFolderId = "16fZDLgdPQsY8kcwSUr9sqU6yliUg8iJy";

var clearlistFilePattern = "^CLEAR_" + "[0-9]{8}" + ".csv";
var clearlistLifecyclePattern =  "^CLEAR_" + "[0-9]{8}" + "_LIFECYCLE.csv";
var clearlistTacPattern = "CLEAR_TAC.csv";
var sharenettFilePattern = "^SHARE_" + "[0-9]{8}" + ".csv";
var sharenettLifecyclePattern =  "^SHARE_" + "[0-9]{8}" + "_LIFECYCLE.csv";
var sharenettTacPattern = "SHARE_TAC.csv";

var clearlistTradesImportSheet = "CL Trade Create";
var clearlistLifecycleImportSheet = "LIFECYCLE";
var clearlistTacImportSheet = "TAC";
var clearlistTradingHistory = "CL Trading History";
var sharenettTradesImportSheet = "SN Trade Create";
var sharenettLifecycleImportSheet = "LIFECYCLE";
var sharenettTacImportSheet = "TAC";
var sharenettTradingHistory = "SN Trading History";

var rangeInTradeTab = "B:B";
var startColInTradeTab = 2;
var rangeInLifecycleTab = "A:A";
var startColInLifecycleTab = 1;
var rangeInTacTab = "A:A";
var startColInTacTab = 1;

var todaytradeoutputFolderName = "Todays_Trades_Export";
var todayTradeRangePending = "B2:AI";
var todayTradeRangePendingGTS = "B2:AZ";
var todayTradeRangeSettled = "B2:AV";
var todayTradeOutputRange = "B2:R";
var tradingHistoryOutputRange = "A1:Q";
var todayTradeOutputPendingFilter = ["PENDING", "SENT"];
var todayTradeOutputSettledFilter = ["SETTLED", "YES","SENT"];
var todayTradeOutputColPendingFilter = [0,33];
var todayTradeOutputColSettledFilter = [0,45,46];
var clearlistTradesLedger = "CL Todays Trades";
var sharenettTradesLedger = "SN Todays Trades";
var todayTradeInsertValueColPending = "AI";
var todayTradeInsertValueColSettled = "AV";

var tradingHistoryOutputFolderName = "Trading_History_Export"
var tradingHistoryRange = "A2:Q"


//Master Balances information used for Onboarding functions 
var masterBalanceSheet = "MASTER BALANCES";
var startColInMasterBalance = 2;
var rangeInMasterBalanceTab = "B:B";
var numberColToFillMasterBalance = 3;


var omnibusOnboardingImportSheet="Omnibus Account Onboarding";
var omnibusOnboardingFilter = ["YES"];
//index 0  = "ok to onboard" col
//index 1 = 
  //Customer: trellis id col
  //BD: MPID col
  //Omni: omnibus col
//index 2 = broker dealer col
//index 3 = date col
var omnibusOnboardingColFilter = [12,10, 5, 16];
var omnibusOnboardingInputRange = "B2:P";

var customerOnboardingTab = "Customer Onboarding"
var customerOnboardingRange ="B2:Z" 
//index 0  = "ok to onboard" col
//index 1 =
  //Customer: trellis id col
  //BD: MPID col
  //Omni: omnibus col
//index 2 = broker dealer col
//index 3 = date col
var customerOnboardingColFilter = [22,3,4,25]
var customerOnboardingFilter = ["YES"];

var brokerDealerOnboardingTab = "Broker Dealer Onboarding"
var brokerDealerOnboardingRange ="F2:U" 
//index 0  = "ok to onboard" col
//index 1 =
  //Customer: trellis id col
  //BD: MPID col
  //Omni: omnibus col
//index 2 = broker dealer col
//index 3 = date col
var brokerDealerOnboardingColFilter = [14,2,2,21]
var brokerDealerOnboardingFilter = ["YES"];


var todayTradeOutputPendingGTSFilter = ["PENDING","GTSY","50077","SENT"];
var todayTradeOutputSettledGTSFilter = ["SETTLED","GTSY","50077","SENT"];
var todayTradeOutputColPendingGTSFilter = [0,3,5,2,4,49];
var todayTradeOutputColSettledGTSFilter = [0,3,5,2,4,50];
var todayTradeInsertValueColPendingGTS = "AY";
var todayTradeInsertValueColSettledGTS = "AZ";


var tradeHistoryOutputPendingGTSFilter = ["PENDING","GTSY","50077"];
var tradeHistoryOutputSettledGTSFilter = ["SETTLED","GTSY","50077"];
var tradeHistoryOutputColPendingGTSFilter = [0,3,5,2,4];


function omnibusOnboarding(){
  onboardingToMasterBalance(omnibusOnboardingImportSheet, masterBalanceSheet, rangeInMasterBalanceTab, omnibusOnboardingInputRange, omnibusOnboardingColFilter, omnibusOnboardingFilter, startColInMasterBalance, numberColToFillMasterBalance)
}

function customerOnboardingNEW(){
  onboardingToMasterBalance(customerOnboardingTab, masterBalanceSheet, rangeInMasterBalanceTab, customerOnboardingRange, customerOnboardingColFilter, customerOnboardingFilter, startColInMasterBalance, numberColToFillMasterBalance)
  sendOnboardingEmail()

}

function brokerDealerOnboardingNEW(){
 onboardingToMasterBalance(brokerDealerOnboardingTab, masterBalanceSheet, rangeInMasterBalanceTab, brokerDealerOnboardingRange, brokerDealerOnboardingColFilter, brokerDealerOnboardingFilter, startColInMasterBalance, numberColToFillMasterBalance)
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
          Logger.log("data[row][onboardingColFilter[0]] "+ data[row][onboardingColFilter[0]])
          Logger.log("onboardingFilter[0] "+ onboardingFilter[0])
          if (data[row][onboardingColFilter[0]] == onboardingFilter[0] ) {
            var customerName = data[row][onboardingColFilter[1]];
            var brokerDealerId = data[row][onboardingColFilter[2]];    
            Logger.log("name: "+ customerName + " Broker: "+brokerDealerId);
            outputss.getRange(outputSheetTotalRows, startColInOutputSheet,1, numberColToFillOutputSheet).setValues([["OK", customerName, brokerDealerId]]);
            outputss.getRange(outputSheetTotalRows+1, startColInOutputSheet,1, numberColToFillOutputSheet).setValues([["OK", "Holding_"+customerName, brokerDealerId]]);
            importss.getRange(row+2,onboardingColFilter[3]).setValue(new Date)
            outputSheetTotalRows+=2;
            
          }
      }
    }
    catch (err) {
      Logger.log(err);
      Browser.msgBox(err);
    }
}




//creates folders in Archive folders
//used by importFromCSV
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

// reads the contents of the csv and adds them into the sheet line by line 
// since might have multiple files -> should check where is the last row and append it 
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

function importTradesCLNEW(){
  importFromCSV(masterIncomingFolderId, clearlistMainFolderId, clearlistTradeArchiveFolderId ,clearlistFilePattern, clearlistTradesImportSheet, rangeInTradeTab, startColInTradeTab)
}

function importTradesSN(){
  importFromCSV(masterIncomingFolderId, sharenettMainFolderId, sharenettTradeArchiveFolderId,sharenettFilePattern, sharenettTradesImportSheet, rangeInTradeTab, startColInTradeTab)
}

function importAllTrade(){
  importTradesCL()
  importTradesSN()
}

function importLifecycleCL(){
  importFromCSV(masterIncomingFolderId, clearlistMainFolderId, clearlistLifecycleArchiveFolderId, clearlistLifecyclePattern, clearlistLifecycleImportSheet, rangeInLifecycleTab, startColInLifecycleTab)
}

function importLifecycleSN(){
  importFromCSV(masterIncomingFolderId, sharenettMainFolderId, sharenettLifecycleArchiveFolderId, sharenettLifecyclePattern, sharenettLifecycleImportSheet, rangeInLifecycleTab, startColInLifecycleTab)
}

function importAllLifecycle(){
  importLifecycleCL()
  importLifecycleSN()
}

function importTacCL(){
  importFromCSV(masterIncomingFolderId, clearlistMainFolderId, clearlistTacArchiveFolderId, clearlistTacPattern, clearlistTacImportSheet, rangeInTacTab, startColInTacTab)
}

function importTacSN(){
  importFromCSV(masterIncomingFolderId, sharenettMainFolderId, sharenettTacArchiveFolderId, sharenettTacPattern, sharenettTacImportSheet, rangeInTacTab, startColInTacTab)
}

function importAllTac(){
  importTacCL()
  importTacSN()
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//IMPORTING FUNCTIONS THAT USE THE OLD FORMAT - these will be used during Gatsby launch
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function importTacCLOldFormat(){
  var clearlistTacPattern = "tac_file.csv";
  importFromCSV(masterIncomingFolderId, clearlistMainFolderId, clearlistTacArchiveFolderId, clearlistTacPattern, clearlistTacImportSheet, rangeInTacTab, startColInTacTab)
}

function importTradesCLNEWOldFormat(){
  var clearlistFilePattern = "^CLEAR." + "[0-9]{8}" + ".csv";
  importFromCSV(masterIncomingFolderId, clearlistMainFolderId, clearlistTradeArchiveFolderId ,clearlistFilePattern, clearlistTradesImportSheet, rangeInTradeTab, startColInTradeTab)
}










// today trade make as csv

//this function works for all CSVs !!except SETTLED!!, adding a new one for that called convertToCSV (modify and work for all)
function convertToCSVORIGINAL(ss, totalRows, todayTradeRange, outputTradeRange, todayTradeOutputFilter, todayTradeOutputColFilter, todayTradeInsertValueCol) {
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
      var name = "";
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
      return [csvFile,name];
    }
    catch (err) {
      Logger.log(err);
    }
  }else if (todayTradeOutputFilter.length==2){
    Logger.log("With 2 Filters for pending");
    try {
      //var data = activeRange.getValues();
      var csvFile = undefined;
      var name = "";

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
      return [csvFile,name];
    }
    catch (err) {
      Logger.log(err);
      Browser.msgBox(err);
    }
  }else if(todayTradeOutputFilter.length==3 && todayTradeOutputFilter[2]=="50077"){
    Logger.log("With 3 filters for GTS trading history");
      try {
      var csvFile = undefined;
      var name = "GTS";

      if (data.length > 1) {
        var csv = "";
        for (var row = 0; row < data.length; row++) {
          if ( data[row][0] == "Transaction Type") {
            csv += data2[row].join(",") + "\r\n";
            Logger.log("Add title row")
          }

          if (data[row][todayTradeOutputColFilter[0]] == todayTradeOutputFilter[0] 
            && ((data[row][todayTradeOutputColFilter[1]] == todayTradeOutputFilter[1]) | (data[row][todayTradeOutputColFilter[2]] == todayTradeOutputFilter[1]) | (data[row][todayTradeOutputColFilter[3]] == todayTradeOutputFilter[2]) |(data[row][todayTradeOutputColFilter[4]] == todayTradeOutputFilter[2]))) {

                  if (row < data2.length - 1) {
                    csv += data2[row].join(",") + "\r\n";
                  }
                  else {
                    csv += data2[row];
                    Logger.log("Adding row to CSV")
                  }
            }
          }
        }
        csvFile = csv;
      }
      catch (err) {
      Logger.log(err);
      Browser.msgBox(err);
    }
      return [csvFile,name];

  }else if(todayTradeOutputFilter.length==3){
    Logger.log("With 3 filters for settled");
      try {
      var csvFile = undefined;
      var name = "";

      if (data.length > 1) {
        var csv = "";
        for (var row = 0; row < data.length; row++) {
          if ( data[row][0] == "Transaction Type") {
            csv += data2[row].join(",") + "\r\n";
            Logger.log("Add title row")
          }
          if (data[row][todayTradeOutputColFilter[0]] == todayTradeOutputFilter[0] && data[row][todayTradeOutputColFilter[1]] == todayTradeOutputFilter[1] && data[row][todayTradeOutputColFilter[2]] != todayTradeOutputFilter[2]) {

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
      catch (err) {
      Logger.log(err);
      Browser.msgBox(err);
    }
      return [csvFile,name];
    }else if(todayTradeOutputFilter.length==4){

        Logger.log("covert to csv for GTS");
        Logger.log(todayTradeOutputColFilter[5]);

    try {
      var csvFile = undefined;
      var name = "GTS";

      if (data.length > 1) {
        var csv = "";
        for (var row = 0; row < data.length; row++) {
          if ( data[row][0] == "Transaction Type") {
            csv += data2[row].join(",") + "\r\n";
            Logger.log("Add title row")
          }

          if (data[row][todayTradeOutputColFilter[0]] == todayTradeOutputFilter[0] 
            && (data[row][todayTradeOutputColFilter[5]] != todayTradeOutputFilter[3])
            && ((data[row][todayTradeOutputColFilter[1]] == todayTradeOutputFilter[1]) | (data[row][todayTradeOutputColFilter[2]] == todayTradeOutputFilter[1]) | (data[row][todayTradeOutputColFilter[3]] == todayTradeOutputFilter[2]) |(data[row][todayTradeOutputColFilter[4]] == todayTradeOutputFilter[2]))
            ) {
              Logger.log(row);
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
    catch (err) {
      Logger.log(err);
      Browser.msgBox(err);
    }
    return [csvFile, name];
    }
}



//THIS FUNCTION WORKS FOR ALL CSVS EXCEPT SETTLED TRADES, ADDING A SEPARATE ONE FOR THAT  //called convertToCSVandCreateFilesToFoldersSETTLEDONLY
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
  var csvFile = convertToCSVORIGINAL(ss,totalRows, ledgerRange, ledgerOutputRange, ledgerOutputFilter, ledgerOutputColFilter, ledgerInsertValueCol)[0];
  var atsNameo = convertToCSVORIGINAL(ss,totalRows, ledgerRange, ledgerOutputRange, ledgerOutputFilter, ledgerOutputColFilter, ledgerInsertValueCol)[1];
  Logger.log(atsNameo);

  if (atsNameo=="GTS"){
    var atsName = atsNameo;
    var atsNameo = ss.getName().split(" ")[0];
    var outputFileName = ss.getName().replace(atsNameo,'').replace(" ",'').replace(" ",''); 
    var fileName = atsName+"_"+outputFileName+ "_" + ledgerOutputFilter[0] + "_" + dateFormatted + "_" + currentTime + ".csv";
    Logger.log(fileName);
  }else{
    var atsNameo = ss.getName().split(" ")[0];
    var atsName = ss.getName().split(" ")[0].replace("CL","CLEAR").replace("SN","SHARE");
    // create a file in the Docs List with the given name and the csv data
    var outputFileName = ss.getName().replace(atsNameo,'').replace(" ",'').replace(" ",''); 
    try{
      var fileName = atsName+"_"+outputFileName+ "_" + ledgerOutputFilter[0] + "_" + dateFormatted + "_" + currentTime + ".csv";
      Logger.log(fileName);
    } catch (err) {
    var fileName = atsName+"_"+outputFileName+ "_" + dateFormatted + "_" + currentTime + ".csv";
    Logger.log(err);
    }
  }

  
  var file = dest_folder.createFile(fileName, csvFile);
  var file_output2 = clearlist_outgoing_folder.createFile(fileName, csvFile);
  Logger.log("End function convertTodaysTradeIntoCSVWithNEWTradesOnly")
  return fileName;
}





function downloadPendingTradesCSVCL(){
  convertToCSVandCreateFilesToFolders(clearlistTradesLedger, rangeInTradeTab, bauFolderId, todaytradeoutputFolderName, masterOutgoingFolderId,todayTradeRangePending,todayTradeOutputRange,todayTradeOutputPendingFilter, todayTradeOutputColPendingFilter, todayTradeInsertValueColPending)
}

function downloadGTSPendingTradesCSV(){
  convertToCSVandCreateFilesToFolders(clearlistTradesLedger, rangeInTradeTab, bauFolderId, todaytradeoutputFolderName, masterOutgoingFolderId,todayTradeRangePendingGTS,todayTradeOutputRange,todayTradeOutputPendingGTSFilter, todayTradeOutputColPendingGTSFilter, todayTradeInsertValueColPendingGTS)
}

function downloadGTSSettledTradesCSV(){
  convertToCSVandCreateFilesToFolders(clearlistTradesLedger, rangeInTradeTab, bauFolderId, todaytradeoutputFolderName, masterOutgoingFolderId,todayTradeRangePendingGTS,todayTradeOutputRange,todayTradeOutputSettledGTSFilter, todayTradeOutputColSettledGTSFilter, todayTradeInsertValueColSettledGTS)
}

function downloadPendingTradesCSVSN(){
  convertToCSVandCreateFilesToFolders(sharenettTradesLedger, rangeInTradeTab, bauFolderId, todaytradeoutputFolderName, masterOutgoingFolderId,todayTradeRangePending,todayTradeOutputRange,todayTradeOutputPendingFilter, todayTradeOutputColPendingFilter, todayTradeInsertValueColPending)
}

function downloadSettledTradesCSVCL(){
  convertToCSVandCreateFilesToFolders(clearlistTradesLedger, rangeInTradeTab, bauFolderId, todaytradeoutputFolderName, masterOutgoingFolderId,todayTradeRangeSettled,todayTradeOutputRange,todayTradeOutputSettledFilter, todayTradeOutputColSettledFilter, todayTradeInsertValueColSettled)
}

function downloadSettledTradesCSVSN(){
  convertToCSVandCreateFilesToFolders(sharenettTradesLedger, rangeInTradeTab, bauFolderId, todaytradeoutputFolderName, masterOutgoingFolderId,todayTradeRangeSettled,todayTradeOutputRange,todayTradeOutputSettledFilter, todayTradeOutputColSettledFilter, todayTradeInsertValueColSettled)
}

function nofilter(){
  convertToCSVandCreateFilesToFolders(clearlistTradesLedger, rangeInTradeTab, bauFolderId, todaytradeoutputFolderName, masterOutgoingFolderId,todayTradeRangePending,todayTradeOutputRange)
}

  var positionsCL = "CL Positions"
  var positionsSN = "SN Positions"
  var positionsGTS = "GTS Positions"
  var rangeInPositions = "B:B"
  var positionsOutputFolderName = "Positions"
  var rangeUndefined = "A1:P"
  var rangeForCSVInPositioins = "A1:H"

function downloadCLPositions(){
  convertToCSVandCreateFilesToFolders(positionsCL,rangeInPositions,bauFolderId,positionsOutputFolderName,masterOutgoingFolderId,rangeUndefined,rangeForCSVInPositioins) 
}

function downloadSNPositions(){
  convertToCSVandCreateFilesToFolders(positionsSN,rangeInPositions,bauFolderId,positionsOutputFolderName,masterOutgoingFolderId,rangeUndefined,rangeForCSVInPositioins) 
}

function downloadGTSPositions(){
  convertToCSVandCreateFilesToFolders(positionsGTS,rangeInPositions,bauFolderId,positionsOutputFolderName,masterOutgoingFolderId,rangeUndefined,rangeForCSVInPositioins) 
}

//CHLOE pls fill in this function
/***
 * Converts range A:Q in CL Trading History into a CSV with name CLEAR_TradingHistory_YYYYMMDD.csv
 * CSV gets saved to Clearlist-prod-data> Outgoing folder 
 * CSV also gets saved to Operations > BAU> YYYY-MM > YYYY-MM-DD > Trading_History_Export 
 */
function downloadTradingHistoryCSVCL(){
  convertToCSVandCreateFilesToFolders(clearlistTradingHistory, rangeInTradeTab, bauFolderId, todaytradeoutputFolderName, masterOutgoingFolderId,tradingHistoryOutputRange,tradingHistoryOutputRange)
}

function downloadGTSPendingTradingHistoryCSV(){
  convertToCSVandCreateFilesToFolders(clearlistTradingHistory, rangeInTradeTab, bauFolderId, todaytradeoutputFolderName, masterOutgoingFolderId,tradingHistoryOutputRange,tradingHistoryOutputRange, tradeHistoryOutputPendingGTSFilter, tradeHistoryOutputColPendingGTSFilter)
}

function downloadGTSSettledTradingHistoryCSV(){
  convertToCSVandCreateFilesToFolders(clearlistTradingHistory, rangeInTradeTab, bauFolderId, todaytradeoutputFolderName, masterOutgoingFolderId,tradingHistoryOutputRange,tradingHistoryOutputRange, tradeHistoryOutputSettledGTSFilter, tradeHistoryOutputColPendingGTSFilter)
}

//CHLOE pls fill in this function
/***
 * Converts range A:Q in SN Trading History into a CSV with name SHARE_TradingHistory_YYYYMMDD.csv
 * CSV gets saved to Clearlist-prod-data> Outgoing folder 
 * CSV also gets saved to Operations > BAU> YYYY-MM > YYYY-MM-DD > Trading_History_Export 
 */
function downloadTradingHistoryCSVSN(){
  convertToCSVandCreateFilesToFolders(sharenettTradingHistory, rangeInTradeTab, bauFolderId, todaytradeoutputFolderName, masterOutgoingFolderId,tradingHistoryOutputRange,tradingHistoryOutputRange)
}

function eodTradingHistoryCL(){
  downloadTradingHistoryCSVCL()
  moveTradesFromCLTradingHistoryToCLTodaysTradesAtEOD()
}

function eodTradingHistorySN(){
  downloadTradingHistoryCSVSN()
  moveTradesFromSNTradingHistoryToSNTodaysTradesAtEOD()
}








////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//MOVING TRADE DATA FROM TRADE CREATE TO TODAYS TRADES  
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


function moveTradesNew(tabFrom, rowsInTabFrom, startingIndexInTabFrom, startingColInTabFrom, lastColInTabFrom, tabTo, rowsInTabTo,  startingColTabTo,rowIndexTabTo, colIndexTabTo, rangeToClearTabFrom){
  for (var i = startingIndexInTabFrom; i <= rowsInTabFrom; i++) {
    //gets values from tabFrom in a specified range
    var values = tabFrom.getRange(startingColInTabFrom + i + ":"+ lastColInTabFrom + i).getValues();
    //populates tabTo with the values defined in the previous line 
    tabTo.getRange(rowsInTabTo, startingColTabTo, rowIndexTabTo, colIndexTabTo).setValues(values);
    //adds 1 to the variable lastrow in order to make sure that the information from the next trade is written into the next row
    rowsInTabTo +=1
  }
  //clears specified range in tabFrom 
  tabFrom.getRange(rangeToClearTabFrom).clearContent();
}

function moveTradesFromCLTradeCreateToCLTodaysTrades(){
  var ssTabFrom = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tradeCreateCL);
  var rowsInTabFrom = getLastRow(tradeCreateCL, "B:B");
  var rowsInTabFrom = rowsInTabFrom + 1;
  var startingColInTabFrom = "B"
  var lastColInTabFrom = "Q"
  var startingIndexInTabFrom = 3
  
  var ssTabTo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(todaysTradesCL); 
  var rowsInTabTo = getLastRow(todaysTradesCL, "B:B");
  var rowsInTabTo = rowsInTabTo + 1;
  var startingColTabTo = 2
  var rowIndexTabTo = 1
  var colIndexTabTo = 16

  var rangeToClearTabFrom = "B3:Q" + rowsInTabFrom

  moveTradesNew(ssTabFrom, rowsInTabFrom, startingIndexInTabFrom, startingColInTabFrom, lastColInTabFrom, ssTabTo, rowsInTabTo,  startingColTabTo,rowIndexTabTo, colIndexTabTo, rangeToClearTabFrom)
  assignCLUniqueSeqRefID()
}

function moveTradesFromSNTradeCreateToSNTodaysTrades(){
  var ssTabFrom = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tradeCreateSN);
  var rowsInTabFrom = getLastRow(tradeCreateSN, "B:B");
  var rowsInTabFrom = rowsInTabFrom + 1;
  var startingColInTabFrom = "B"
  var lastColInTabFrom = "Q"
  var startingIndexInTabFrom = 3
  
  var ssTabTo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(todaysTradesSN); 
  var rowsInTabTo = getLastRow(todaysTradesSN, "B:B");
  var rowsInTabTo = rowsInTabTo + 1;
  var startingColTabTo = 2
  var rowIndexTabTo = 1
  var colIndexTabTo = 16

  var rangeToClearTabFrom = "B3:Q" + rowsInTabFrom

  moveTradesNew(ssTabFrom, rowsInTabFrom, startingIndexInTabFrom, startingColInTabFrom, lastColInTabFrom, ssTabTo, rowsInTabTo,  startingColTabTo,rowIndexTabTo, colIndexTabTo, rangeToClearTabFrom)
  assignSNUniqueSeqRefID()
}


//ASSIGNING UNIQUE ID TO TRADES IN TODAYS TRADES 
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

//assigns unique ID to newly imported trades using the trade ID and dateTime
//used when trades are moved from Create Trade to Todays Trades
function assignCLUniqueSeqRefID(){
  //get Todays Trades Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(todaysTradesCL);
  var lastRow = getLastRow(todaysTradesCL, "B:B");
  assignUniqueSequenceRefID(ss,lastRow)
}

function assignSNUniqueSeqRefID(){
  //get Todays Trades Spreadsheet SN
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(todaysTradesSN);
  var lastRow = getLastRow(todaysTradesSN, "B:B");
  assignUniqueSequenceRefID(ss,lastRow)
}

//assigns unique ID to newly imported trades using the trade ID and dateTime
//used when trades are moved from Trade Create to Todays Trades
function assignUniqueSequenceRefID(ss, lastRow) {
  var dateFormatted = Utilities.formatDate(new Date(), 'America/New_York', 'MMddyyyyHHmmss');  
  
  var notation = "B2:R"+lastRow
  var data = ss.getRange(notation).getValues()

  //pull columns for trade ID and unique Seq
  var uniqueSequenceColLetter = "R"

  for (var i = 0; i < lastRow-1; i++) {
    //check that unique ID does not already exist 
    Logger.log("index is "+ i)
    Logger.log("data[i][12] is "+ data[i][12])
    if (data[i][12] != "" && data[i][16] == "") {
      Logger.log("IM INSIDE THE IF STATEMENT")
      //then assign a unique id by combining the trade id and datetime
      
      var transactionID = data[i][12]
      var uniqueSequenceID = dateFormatted + transactionID

      Logger.log("We need to populate row " + (i+2))
      //assign unique sequence ID to the trade  
      ss.getRange(uniqueSequenceColLetter + (i+2)).setValue(uniqueSequenceID);

    }
  }
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//MOVING TRADE DATA FROM TODAYS TRADES TO TRADING HISTORY
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


function moveTradesFromCLTodaysTradesToCLTradingHistory(){
  var ssTabFrom = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(todaysTradesCL);
  var rowsInTabFrom = getLastRow(todaysTradesCL, "B:B");
  var rowsInTabFrom = rowsInTabFrom + 1;
  var startingColInTabFrom = "B"
  var lastColInTabFrom = "R"
  var startingIndexInTabFrom = 3
  

  var ssTabTo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tradingHistoryCL); 
  var rowsInTabTo = getLastRow(tradingHistoryCL, "B:B");
  var rowsInTabTo = rowsInTabTo + 1;
  var startingColTabTo = 1
  var rowIndexTabTo = 1
  var colIndexTabTo = 17

  var rangeToClearTabFrom = "B3:R" + rowsInTabFrom

  moveTradesNew(ssTabFrom, rowsInTabFrom, startingIndexInTabFrom, startingColInTabFrom, lastColInTabFrom, ssTabTo, rowsInTabTo,  startingColTabTo,rowIndexTabTo, colIndexTabTo, rangeToClearTabFrom)

  
  var startingColInTabFrom = "AI"
  var lastColInTabFrom = "AK"
  var startingColTabTo = 18
  var colIndexTabTo = 3
  var rangeToClearTabFrom = "AI3:AK" + rowsInTabFrom
  moveTradesNew(ssTabFrom, rowsInTabFrom, startingIndexInTabFrom, startingColInTabFrom, lastColInTabFrom, ssTabTo, rowsInTabTo,  startingColTabTo,rowIndexTabTo, colIndexTabTo, rangeToClearTabFrom)

  var rangeToClearLast = "AV3:AZ"+ rowsInTabFrom
  ssTabFrom.getRange(rangeToClearLast).clearContent();
}

function moveTradesFromSNTodaysTradesToSNTradingHistory(){
  var ssTabFrom = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(todaysTradesSN);
  var rowsInTabFrom = getLastRow(todaysTradesSN, "B:B");
  var rowsInTabFrom = rowsInTabFrom + 1;
  var startingColInTabFrom = "B"
  var lastColInTabFrom = "R"
  var startingIndexInTabFrom = 3
  

  var ssTabTo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tradingHistorySN); 
  var rowsInTabTo = getLastRow(tradingHistorySN, "B:B");
  var rowsInTabTo = rowsInTabTo + 1;
  var startingColTabTo = 1
  var rowIndexTabTo = 1
  var colIndexTabTo = 17

  var rangeToClearTabFrom = "B3:R" + rowsInTabFrom

  moveTradesNew(ssTabFrom, rowsInTabFrom, startingIndexInTabFrom, startingColInTabFrom, lastColInTabFrom, ssTabTo, rowsInTabTo,  startingColTabTo,rowIndexTabTo, colIndexTabTo, rangeToClearTabFrom)

  
  var startingColInTabFrom = "AI"
  var lastColInTabFrom = "AK"
  var startingColTabTo = 18
  var colIndexTabTo = 3
  var rangeToClearTabFrom = "AI3:AK" + rowsInTabFrom
  moveTradesNew(ssTabFrom, rowsInTabFrom, startingIndexInTabFrom, startingColInTabFrom, lastColInTabFrom, ssTabTo, rowsInTabTo,  startingColTabTo,rowIndexTabTo, colIndexTabTo, rangeToClearTabFrom)

  var rangeToClearLast = "AV3:AZ"+ rowsInTabFrom
  ssTabFrom.getRange(rangeToClearLast).clearContent();
}


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//MOVING TRADE DATA FROM TRADING HISTORY TO TODAYS TRADES 
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


function moveTradesFromCLTradingHistoryToCLTodaysTradesAtEOD(){
  var nameTabFrom = tradingHistoryCL
  var nameTabTo = todaysTradesCL
  var colToCheckRowsTabFrom = 'A:A'
  var colToCheckRowsTabTo = 'B:B'
  var range1TabFrom = 'A2:Q'
  var range2TabFrom = 'R2:T'
  var check1 = "PENDING"
  var check2 = "NEW"
  var rangeToClearTabFrom = 'A2:T'

 moveTradesWithChecks(nameTabFrom, nameTabTo, colToCheckRowsTabFrom, colToCheckRowsTabTo,range1TabFrom,range2TabFrom,check1, check2, rangeToClearTabFrom )

}

function moveTradesFromSNTradingHistoryToSNTodaysTradesAtEOD(){
  var nameTabFrom = tradingHistorySN
  var nameTabTo = todaysTradesSN
  var colToCheckRowsTabFrom = 'A:A'
  var colToCheckRowsTabTo = 'B:B'
  var range1TabFrom = 'A2:Q'
  var range2TabFrom = 'R2:T'
  var check1 = "PENDING"
  var check2 = "NEW"
  var rangeToClearTabFrom = 'A2:T'

 moveTradesWithChecks(nameTabFrom, nameTabTo, colToCheckRowsTabFrom, colToCheckRowsTabTo,range1TabFrom,range2TabFrom,check1, check2, rangeToClearTabFrom )

}


function moveTradesWithChecks(nameTabFrom, nameTabTo, colToCheckRowsTabFrom, colToCheckRowsTabTo,range1TabFrom,range2TabFrom,check1, check2, rangeToClearTabFrom ){
   var ssTabFrom = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameTabFrom);
  var ssTabTo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameTabTo);
  var totalRowsTabFrom = getLastRow(nameTabFrom, colToCheckRowsTabFrom);
  var totalRowsTabTo = getLastRow(nameTabTo, colToCheckRowsTabTo);
  var rangeToClear = rangeToClearTabFrom + totalRowsTabFrom

  var notation1 = range1TabFrom + totalRowsTabFrom + 1;
  var dataTabFromRange1 = ssTabFrom.getRange(notation1).getValues(); 
  var notation2 = range2TabFrom + totalRowsTabFrom + 1;
  var dataTabFromRange2 = ssTabFrom.getRange(notation2).getValues();
  
  for (var i = 0; i < dataTabFromRange1.length; i++) {
    if (dataTabFromRange1[i][0] == check1 || dataTabFromRange1[i][0] == check2) {
      ssTabTo.getRange(totalRowsTabTo + 1, 2, 1, 17).setValues([dataTabFromRange1[i]]);
      ssTabTo.getRange(totalRowsTabTo + 1, 35, 1, 3).setValues([dataTabFromRange2[i]]);
      totalRowsTabTo += 1;
    }
  }
  ssTabFrom.getRange(rangeToClear).clearContent();
}




/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//IMPORT DATA (TAC & LIFECYCLE) 
// RE-WRITE THESE FUNCTIONS TO WRITE ANY FORMAT DATA TO ANY SHEET
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// import tac file
function writeDataToSheetTAC(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TAC");
  var last_rows = getLastRow("TAC", "A:A");
  Logger.log(last_rows);
  //ss.getRange(2, 2, data.length, data[1].length).setValues(data);
  ss.getRange(last_rows + 1, 1, data.length, data[0].length).setValues(data);
  //return ss.getName();
}


function importFromCSVForTAC_CreateFolder_MoveFileToArchive() {
  // read all files in the Clearlist -> Ops -> BAU folders -> new clearlist trade files -> folder id: 1ylVnMiFV4pkas5X5sPkcNsI-2VKoU_eW
  // read all files in the https://drive.google.com/drive/folders/1sp2QzTccc7wJR6l-CQr-5D2S-MNAb4NU -> folder id: 1sp2QzTccc7wJR6l-CQr-5D2S-MNAb4NU
  var mainFolder = DriveApp.getFolderById("1sp2QzTccc7wJR6l-CQr-5D2S-MNAb4NU");
  var f = mainFolder.getFiles();
  var blankfile = [];

  var archivefolderid = "1v0TnSyLMEh4Obx6_8-rHSHQJ1hHFGf6C";
  var todaydatefolder = Utilities.formatDate(new Date(), timeZone, "MM-dd-yyyy");
  var todaydatefolderid = createFolder(archivefolderid, todaydatefolder);
  //var uploadclearlistfolderid = createFolder(todaydatefolderid,"Uploaded_Clearlist_Trade_Files");
  var dest_folder = DriveApp.getFolderById(todaydatefolderid);
  var empty_folder = DriveApp.getFolderById("1lz603RwoP-tryMkOqxtsd974dMrPKwm7");
  var todaydatefolderid_bau = createFolder("1s132fsm3mrJX47MLEBzGdVtcCWKLtjbt", todaydatefolder);
  var uploadtacfolderid = createFolder(todaydatefolderid_bau, "TAC_Files");
  var dest_bau_tac_folder = DriveApp.getFolderById(uploadtacfolderid);



  while (f.hasNext()) {
    var file = f.next();
    //var regExp = new RegExp("^CLEAR.20210205.csv$");
    var regExp = new RegExp("^tac_file.csv")

    if (file.getName().search(regExp) != -1) {
      name = file.getName();
      Logger.log(name);
      try {
        var contents = Utilities.parseCsv(file.getBlob().getDataAsString());
        // remove header of the files
        var header = contents.shift();
        writeDataToSheetTAC(contents);
        //file.copyTo(dest_bau_tac_folder);
        file.moveTo(dest_folder);
      } catch (err) {
        Logger.log(err);
        blankfile.push(name);
        file.moveTo(empty_folder);
      }
    }
  };
  return blankfile;
  //Logger.log(blankfile);
  //displayToastAlert("The CSV file was successfully imported into ");
}

// import lifecycle file
function writeDataToSheetLIFECYCLE(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LIFECYCLE");
  var last_rows = getLastRow("LIFECYCLE", "A:A");
  Logger.log(last_rows);
  //ss.getRange(2, 2, data.length, data[1].length).setValues(data);
  ss.getRange(last_rows + 1, 1, data.length, data[0].length).setValues(data);
  //return ss.getName();
}




function importFromCSVForLIFECYCLE_CreateFolder_MoveFileToArchive() {
  // read all files in the Clearlist -> Ops -> BAU folders -> new clearlist trade files -> folder id: 1ylVnMiFV4pkas5X5sPkcNsI-2VKoU_eW
  // read all files in the https://drive.google.com/drive/folders/1sp2QzTccc7wJR6l-CQr-5D2S-MNAb4NU -> folder id: 1sp2QzTccc7wJR6l-CQr-5D2S-MNAb4NU
  var mainFolder = DriveApp.getFolderById("1sp2QzTccc7wJR6l-CQr-5D2S-MNAb4NU");
  var f = mainFolder.getFiles();
  var blankfile = [];

  var archivefolderid = "1Cg6P9npX90iwMsHwsZzPQo1JOGMDRDSu";
  var todaydatefolder = Utilities.formatDate(new Date(), timeZone, "MM-dd-yyyy");
  var todaydatefolderid = createFolder(archivefolderid, todaydatefolder);
  var dest_folder = DriveApp.getFolderById(todaydatefolderid);
  var empty_folder = DriveApp.getFolderById("1Q1Eivg96Vg80FEsYig7gnHKKxKKcJOT-");

  var monthdatefolder = Utilities.formatDate(new Date(), timeZone, "yyyy-MM");
  var monthfolderid_bau = createFolder("1s132fsm3mrJX47MLEBzGdVtcCWKLtjbt", monthdatefolder);
  var todaydatefolderid_bau = createFolder(monthfolderid_bau, todaydatefolder);
  var uploadtacfolderid = createFolder(todaydatefolderid_bau, "LIFECYCLE_Files");
  var dest_bau_tac_folder = DriveApp.getFolderById(uploadtacfolderid);



  while (f.hasNext()) {
    var file = f.next();
    var regExp = new RegExp("LIFECYCLE.csv$")

    if (file.getName().search(regExp) != -1) {
      name = file.getName();
      Logger.log(name);
      try {
        var contents = Utilities.parseCsv(file.getBlob().getDataAsString());
        // remove header of the files
        var header = contents.shift();
        writeDataToSheetLIFECYCLE(contents);
        // file.copyTo(dest_bau_tac_folder);
        file.moveTo(dest_folder);
      } catch (err) {
        Logger.log(err);
        blankfile.push(name);
        file.moveTo(empty_folder);
      }
    }
  };
  return blankfile;
}


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//PROCESSING TRADES: 
//Status changes from NEW to PENDING
//Generates CSV of newly PENDED trades sends to issuer agent via Google Drive > Outgoing folder
//Customers emailed letting them know that the trades are being processed
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function NEWtoPENDINGFunction() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tradesLedger);
  var totalRows = getLastRow(tradesLedger, 'B:B');
  //moveCashAndSecuritiesToHoldingAccounts(ss, totalRows)
  convertTodaysTradeIntoCSVWithNEWTradesOnly(ss, totalRows)
  sendEmailsToSellerBuyerBDsBeforeSettlement(ss, totalRows)

}

function NEWtoPENDINGCL(){
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(todaysTradesCL);
  var totalRows = getLastRow(todaysTradesCL, 'B:B');
  moveCashAndSecuritiesToHoldingAccounts(ss, totalRows)
  sendEmailsToSellerBuyerBDsBeforeSettlement(ss, totalRows)
}

function NEWtoPENDINGSN(){
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(todaysTradesSN);
  var totalRows = getLastRow(todaysTradesSN, 'B:B');
  moveCashAndSecuritiesToHoldingAccounts(ss, totalRows)
  sendEmailsToSellerBuyerBDsBeforeSettlement(ss, totalRows)
}

/***
 * This is one of the main functions, it processes trades
 * Moves cash from Buyer's Personal Account or BD's Omnibus Account to a Holding Account 
 * Moves securiteis from Seller's Account to Seller's Holding Account 
 * 
 * Takes in the Todays Trade tab of the relevant ATS 
 */

//moves cash and securities from customers' main accounts to holding accounts
function moveCashAndSecuritiesToHoldingAccounts(ss, totalRows) {
  var numOfTradesProcessed =0
  //check this variable before updating balances
  var okayToSendCSVtoIAColNum = 32

  var operation = "NEWtoPENDING"
  
  //info for Todays Trades spreadsheet
  var notationTT = "B3:AH" + totalRows
  var data = ss.getRange(notationTT).getValues()

  var totalRowsBalancesHistory = ssBalancesHistory.getLastRow();

  //identifying area of MB to be looked at for the functions that updates MB Cash & Securities
  var ssMB = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(masterBalancesSheet);
  const totalRowsMB = getLastRow(masterBalancesSheet, 'B:B');
  var notationMB = "B1:Q" + totalRowsMB + 1
  
  //use this just to get customer rows in MB
  var ssMasterBalances = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(masterBalancesSheet);
  var customerDataMasterBalances = ssMasterBalances.getRange("C:C").getValues()

  Logger.log("I'm here")
  //loop through the trades 
  for (var i = 0; i < data.length; i++) {
      Logger.log("inside for loop ")

    //if all seller and buyer requirements met, proceed to update balance
    if (data[i][okayToSendCSVtoIAColNum] == "YES" && data[i][tradeStatusColNum] == "NEW" && numOfTradesProcessed < maxTradesPerBatch){
      //the following will be used to update the Balances history
      //get trade ID
      Logger.log("inside if statement ")

      var tradeID = (data[i][tradeIDColNum]);
      var buyerID = data[i][buyerIDCol]
      var sellerID = data[i][sellerIDCol]
      var sellerHoldingID = "Holding_" + sellerID
      var securityCUSIP = data[i][assetCUSIPColNum]

      //check if customer uses omnibus account or personal funds
      var buyerOmniOrPersonal = customerUsesOmnibusOrPersonal(dataCustomerOnboarding,buyerID, customerOnboardingTabCustomerCashAccountColNum)
      var buyerHoldingOmniOrPersonal = customerUsesOmnibusOrPersonal(dataCustomerOnboarding,buyerID, customerOnboardingTabCustomerHoldingCashAccountColNum)

      //getting price and quantity to calc net notional
      var price = data[i][priceColNum]
      var quantityShares = data[i][quantityColNum]
      var netNotional = price * quantityShares;

      //buyer cash requirement
      var ATSBuyerFee = data[i][ATSBuyerFeeColNum]
      var buyerBDFee = data[i][buyerBDFeeColNum]
      var buyerNetNotional = netNotional;
      var buyerCashObligation = ATSBuyerFee + buyerNetNotional + buyerBDFee;
  
      //get fresh information from Master Balances for each trade
      var dataMB = ssMB.getRange(notationMB).getValues()

      //to be used for Balances History updating 
      var buyerRow = getRowNumberOfCustomerInSpreadsheet(customerDataMasterBalances,buyerOmniOrPersonal)
      var buyerHoldingRow = getRowNumberOfCustomerInSpreadsheet(customerDataMasterBalances,buyerHoldingOmniOrPersonal)
      var sellerRow = getRowNumberOfCustomerInSpreadsheet(customerDataMasterBalances, sellerID)
      var sellerHoldingRow = getRowNumberOfCustomerInSpreadsheet(customerDataMasterBalances,sellerHoldingID)
      
      //seller securities requirement
      var sharesQuantity = data[i][sellerSecurityQuantityColNum]
      var securityColNumInMB = data[i][sellerColInMBofSecurityColNum]

      //before updating Master balances, verify that buyer and seller have sufficient cash and securities. If not, skip this trade
      if(checkIfEnoughCashAndAssetsInAccounts(dataMB, buyerRow, buyerCashObligation, sellerRow, securityColNumInMB, sharesQuantity)){
   
        //updateCustomerCashBalance function calls the updateBalanceHistoryNewFormat so the BH get updated automatically
        //debiting cash from buyer's account 
        updateCustomerCashBalance(ssMB, dataMB, buyerRow, -buyerCashObligation, tradeID, buyerOmniOrPersonal, operation,ssBalancesHistory,totalRowsBalancesHistory)
        totalRowsBalancesHistory +=1;

        //crediting cash to buyer's holding account 
        updateCustomerCashBalance(ssMB, dataMB, buyerHoldingRow, buyerCashObligation, tradeID, buyerHoldingOmniOrPersonal, operation,ssBalancesHistory,totalRowsBalancesHistory)
        totalRowsBalancesHistory +=1;

        //debiting securities from seller's account
        updateCustomerSecurityBalance(ssMB, dataMB, sellerRow, -sharesQuantity, securityColNumInMB, tradeID, sellerID, operation, securityCUSIP,ssBalancesHistory,totalRowsBalancesHistory)
        totalRowsBalancesHistory +=1;

        //crediting securities to seller's holding account 
        updateCustomerSecurityBalance(ssMB, dataMB, sellerHoldingRow, sharesQuantity, securityColNumInMB, tradeID, sellerHoldingID, operation, securityCUSIP,ssBalancesHistory,totalRowsBalancesHistory)
        totalRowsBalancesHistory +=1;

        //change status of trade to "PENDING"
        var pointer = i + 3

        ss.getRange(tradeStatusColumnLetter + pointer).setValue("PENDING")
        numOfTradesProcessed +=1
      }
    }
  }

}


function customerUsesOmnibusOrPersonal(dataCustomerOnboarding,buyerID, accountColNum){
  var account = undefined;
  try{
  for(i=0;dataCustomerOnboarding.length; i++){    
    if(dataCustomerOnboarding[i][customerOnboardingTabCustomerTrellisIDColNum] == buyerID){
      var account = dataCustomerOnboarding[i][accountColNum]
      return account
    }
  }
  }
  catch(err){
    Logger.log("user not found")
    displayToastAlert("Customer's account could not be identified")

  }
  
}




function updateCustomerCashBalance(ss, data, customerRow, delta, tradeID, customerID, operation,ssBalancesHistory,totalRowsBalancesHistory) {
  Logger.log("Starting updateCustomerCashBalance")
  //customerRow provides the index in an array, rather than the row in the spreadsheet. in order to update the cell using col Letter and row number, using an offset of +1
  var offset = 1;
  var currentBalance = data[customerRow][3];
  Logger.log("customer "+customerID)
  Logger.log("row "+customerRow)
  Logger.log("currentBalance "+currentBalance)
  var newBalance = (currentBalance + delta);
  Logger.log("new balance "+newBalance)

  ss.getRange('E' + (customerRow+offset)).setValue(newBalance)
  var asset = "USD";
  updateBalancesHistoryNewFormat(tradeID, customerID, operation, asset, currentBalance, delta, newBalance,ssBalancesHistory,totalRowsBalancesHistory)
  Logger.log("Ending updateCustomerCashBalance")

  
}

//updates customer securities balances in Master Balances 
function updateCustomerSecurityBalance(ss, data, customerRow, sharesQuantity, securityColNum, tradeID, customerID, operation, asset, ssBalancesHistory,totalRowsBalancesHistory) {
  Logger.log("UPDATING MB")
  var offsetRow = 1
  var offsetCol = -2
  var currentBalance = data[customerRow][securityColNum + offsetCol];
  var newBalance = currentBalance + sharesQuantity
  //Logger.log("current SECURITIES balance "+currentBalance)
  //Logger.log("newBalance SECURITIES "+newBalance)
  ss.getRange(customerRow+offsetRow, securityColNum).setValue(newBalance)

  updateBalancesHistoryNewFormat(tradeID, customerID, operation, asset, currentBalance, sharesQuantity, newBalance,ssBalancesHistory,totalRowsBalancesHistory)
}



//populates the Balances History with debits and credits that occur in Master Balances
function updateBalancesHistoryNewFormat(tradeID, customerID, operation, asset, previousBalance, delta, newBalance, ss, totalRows) {
  var lastRow = totalRows + 1
  var time = Utilities.formatDate(new Date(), timeZone, "MM-dd-yyyy HH:mm:ss");

  var valuesArray = [[time, tradeID, customerID, operation, asset, previousBalance, delta, newBalance]]

  ss.getRange("A" + lastRow + ":H" + lastRow).setValues(valuesArray)
}


function convertTodaysTradeIntoCSVWithNEWTradesOnly(ssTodaysTrades, totalRows) {
  Logger.log("Start function convertTodaysTradeIntoCSVWithNEWTradesOnly")
  //creating a folder for Master Balances inside MM-DD-YYYY folder
  var timeZone = "EST";
  var baufolderid = "1s132fsm3mrJX47MLEBzGdVtcCWKLtjbt";
  var todaydatefolder = Utilities.formatDate(new Date(), timeZone, "MM-dd-yyyy");
  var todaydatefolderid = createFolder(baufolderid, todaydatefolder);
  var todaysPendingTradesFolderID = createFolder(todaydatefolderid, "Todays_Trades_Export");
  var dest_folder = DriveApp.getFolderById(todaysPendingTradesFolderID);
  var clearlist_outgoing_folder = DriveApp.getFolderById('1Myehii1D3H_sUrvgtuV-yZegp9I1-ru7');

  //converting Todays Trades (referred to as tradesLedger) to CSV
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(tradesLedger);
  var timeZone = "EST";
  var dateFormatted = Utilities.formatDate(new Date(), timeZone, "yyyyMMdd");

  var d = new Date();
  //var currentTime = d.toLocaleTimeString('en-GB'); 
  var currentTime = d.getHours();

  //Kate's note: i think we can either write the name of the sheet directly here as a string or rename the tab to have no spaces and change variable TradesLedger
  var fileName = sheet.getName().replace(' ', '') + "_PENDING_" + dateFormatted + "_" + currentTime + ".csv";
  // convert all available sheet data to csv format
  var csvFile = convertNEWTradesInTodaysTradeToCsvFileAndChangeToPending(ssTodaysTrades, totalRows);

  // create a file in the Docs List with the given name and the csv data
  var file = dest_folder.createFile(fileName, csvFile);
  var file_output2 = clearlist_outgoing_folder.createFile(fileName, csvFile);
  Logger.log("End function convertTodaysTradeIntoCSVWithNEWTradesOnly")
  return fileName;
}


//Generates CSV of trades that were just processed 
function convertNEWTradesInTodaysTradeToCsvFileAndChangeToPending(ss, totalRows) {
  //var totalRows = ss.getLastRow()
  var totalRows = totalRows + 1; // add first row back 

  var notation = "B2:AH" + totalRows
  var notation2 = "B2:R" + totalRows
  var data = ss.getRange(notation).getValues()
  var data2 = ss.getRange(notation2).getValues()
  // get available data range in the spreadsheet


  try {
    //var data = activeRange.getValues();
    var csvFile = undefined;

    // loop through the data in the range and build a string with the csv data
    if (data.length > 1) {
      var csv = "";
      for (var row = 0; row < data.length; row++) {
        //Logger.log("data row "+data[row][0])
        // PROCESSING used to say NEW
        if (data[row][0] == "PENDING" || data[row][0] == "Transaction Type") {

          if (data[row][33] != "SENT" || data[row][32] == "Okay to Process Trades?") {

            var change_row_number = row + 2;

            // join each row's columns
            // add a carriage return to end of each row, except for the last one
            if (row < data2.length - 1) {
              csv += data2[row].join(",") + "\r\n";
            }
            else {
              csv += data2[row];
              Logger.log("Adding row to CSV")
            }


            if (change_row_number != 2) {
              ss.getRange("AI" + change_row_number).setValue("SENT");
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


//sends email to customers & BDs letting them know that a trade was processed 
function sendEmailsToSellerBuyerBDsBeforeSettlement(ss, totalRows) {
  //Logger.log("Start function sendEmailsToSellerBuyerBDsBeforeSettlement")
  //Todays Trades spreadsheet 
  //var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tradesLedger);
  //const totalRows = getLastRow(tradesLedger, 'B:B');
  var notation = 'B3:AP';
  var data = ss.getRange(notation).getValues();

  //used for status of trade & verification if it's ok to send emails 
  var transactionTypeColNum = 0;
  var tradeStatusColNum = 40; 

  //trade id of trade
  var tradeIDColNum = 12; 

  //variables for IA check
  var emailSentToSellerColNum = 34;
  var emailSentToBuyerColNum = 35;

  //Customer Onboarding tab details (used to get Customer email)
  var ssCustOnboarding = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(customerOnboarding);
  var notationCO = 'B3:F';
  var dataCustomerOnboarding = ssCustOnboarding.getRange(notationCO).getValues();
  var customerOnboardingTabCustomerTrellisIDColNum = 3;
  var customerOnboardingTabCustomerEmailColNum = 2;

  //Broker Dealer Onboarding tab details (used to get BD email)
  var ssBrokerDealer = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(brokerDealerOnboarding);
  var notationDB = 'F3:H';
  var dataBrokerDealerOnboarding = ssBrokerDealer.getRange(notationDB).getValues();
  var BDOnboardingTabBDTrellisIDColNum = 2;
  var BDOnboardingTabBDEmailColNum = 1;


  //if trade is NEW, and okayToSendEmails = "YES" --> send emails to IA, Seller&&BD, Buyer&BD
  for (var i = 0; i < totalRows; i++) {
    if (data[i][transactionTypeColNum] == "PENDING" && (data[i][emailSentToSellerColNum] != 'Sent' || data[i][emailSentToBuyerColNum] != 'Sent')) {
    

      //contents of email for all 
      var tradeID = data[i][tradeIDColNum]
      var subjectDate = new Date();
      var tradeTimeColNum = 1;
      var tradeTime = data[i][tradeTimeColNum];
      var priceColNum = 6;
      var price = data[i][priceColNum];
      var buyingNetNotionalColNum = 10;
      var buyingNetNotional = data[i][buyingNetNotionalColNum];
      var sellingNetNotionalColNum = 11;
      var sellingNetNotional = data[i][sellingNetNotionalColNum];
      var quantityColNum = 7;
      var quantity = data[i][quantityColNum];
      var securityColNum = 8;
      var security = data[i][securityColNum];
      var sellerBDFeesColNum = 14;
      var buyerBDFeesColNum = 13;
      var sellerBDFees = data[i][sellerBDFeesColNum];
      var buyerBDFees = data[i][buyerBDFeesColNum];
      var clearlistFeesBuyer = data[i][9];
      var clearlistFeesSeller = data[i][15]
      
      if(data[i][emailSentToSellerColNum] != 'Sent'){
        //Seller email 
        var sellerTrellisIDColNum = 4;
        var sellerTrellisID = data[i][sellerTrellisIDColNum];

        var sellerEmail = returnEmail(dataCustomerOnboarding, sellerTrellisID, customerOnboardingTabCustomerTrellisIDColNum, customerOnboardingTabCustomerEmailColNum)
        Logger.log("SELLER EMAIL IS "+sellerEmail)
        var sellerEmailFormatted = Utilities.formatString('%0s', sellerEmail)
        var sellerBDTrelllisIDColNum = 5;
        var sellerBDTrelllisID = data[i][sellerBDTrelllisIDColNum]
        //var sellerBDEmail = returnBDEmail(data[i][sellerBDTrelllisIDColNum]);
        var sellerBDEmail = returnEmail(dataBrokerDealerOnboarding, sellerBDTrelllisID, BDOnboardingTabBDTrellisIDColNum, BDOnboardingTabBDEmailColNum)
        Logger.log("SELLER BD EMAIL IS " + sellerBDEmail)

        var sellerBDEmailFormatted = Utilities.formatString('%0s', sellerBDEmail);


        var subjectSeller = "Your Sell Trade in " + security + " is being processed";
        var messageSeller = "Hello,\n\nWe’ve received your trade instruction and are currently processing for settlement.\nPlease find the details of the trade below:\n\nTradeID: " + tradeID + "\nTrade Time: " + tradeTime + " (DD:MM:YYYY-HH:MM:SS)\nPrice: " + price +
          "\nQuantity: " + quantity + "\nSecurity: " + security + "\nSelling Net Notional: " +
          sellingNetNotional + "\nFees: " + (sellerBDFees + clearlistFeesSeller) + "\n\nBest,\nPaxos Private Securities Custody Operations"
        //send email to Seller&BD
        sendEmailWithoutAttachmentFromPrivateSecuritiesOps(sellerEmailFormatted, subjectSeller, messageSeller, sellerBDEmailFormatted)
        //mark "Email Sent to Seller" as YES in Todays Trades
        var emailSentToSellerAndBDColNum = 34;
        ss.getRange(i + 3, emailSentToSellerAndBDColNum + 2).setValue("Sent");
      }

      //Buyer email   
      if(data[i][emailSentToBuyerColNum] != 'Sent'){
        var buyerTrellisIDColNum = 2;
        var buyerTrellisID = data[i][buyerTrellisIDColNum];
        var buyerEmail = returnEmail(dataCustomerOnboarding, buyerTrellisID, customerOnboardingTabCustomerTrellisIDColNum, customerOnboardingTabCustomerEmailColNum)
        Logger.log("BUYER EMAIL IS "+buyerEmail)
        var buyerEmailFormatted = Utilities.formatString('%0s', buyerEmail)
        var buyerBDTrelllisIDColNum = 3;
        var buyerBDTrelllisID = data[i][buyerBDTrelllisIDColNum]
        
        var buyerBDEmail = returnEmail(dataBrokerDealerOnboarding, buyerBDTrelllisID, BDOnboardingTabBDTrellisIDColNum, BDOnboardingTabBDEmailColNum)
        Logger.log("BUYER BD EMAIL IS " + buyerBDEmail)
        var buyerBDEmailFormatted = Utilities.formatString('%0s', buyerBDEmail);


        var subjectBuyer = "Your Buy Trade in " + security + " is being processed";
        var messageBuyer = "Hi,\n\nWe’ve received your trade instruction and are currently processing for settlement.\nPlease find the details of the trade below:\n\nTradeID: " + tradeID + "\nTrade Time: " + tradeTime + " (DD:MM:YYYY-HH:MM:SS)\nPrice: " + price +
          "\nQuantity: " + quantity + "\nSecurity: " + security + "\nBuying Net Notional: " +
          buyingNetNotional + "\nFees: " + (buyerBDFees + clearlistFeesBuyer) + "\n\nBest,\nPaxos Private Securities Custody Operations"
        sendEmailWithoutAttachmentFromPrivateSecuritiesOps(buyerEmailFormatted, subjectBuyer, messageBuyer, buyerBDEmailFormatted)
        var emailSentToBuyerAndBDColNum = 35;
        ss.getRange(i + 3, emailSentToBuyerAndBDColNum + 2).setValue("Sent");
      }
    }

  }
}


//returns email address based on matchingValue from any tab. Tab info has to be passed in as data array
function returnEmail(data, matchingValue, colNumForMatchingValue, colNumContainingEmail){
  Logger.log("Start returnEmail")
  var email = undefined;

  for (var i = 0; i < data.length; i++) {
    if (data[i][colNumForMatchingValue] == matchingValue) {
      email = data[i][colNumContainingEmail];
      return email;
    }
  }
  return email
}

function sendEmailWithoutAttachmentFromPrivateSecuritiesOpsORIGINAL(email, subject, message, BD) {
  //getting the emails of user currently logged in
  Logger.log("Start sendEmailWithoutAttachmentFromPrivateSecuritiesOps")

  var me = Session.getActiveUser().getEmail();
  var aliases = GmailApp.getAliases();
  var regExp = new RegExp("^privatesecuritiesops");

  var ccEmails = BD + ",privatesecuritiesops@paxos.com"

  //looping through aliases
  for (i = 0; i < aliases.length; i++) {
    //Logger.log(aliases[i].search(regExp) != -1);
    if (aliases.length > 0 && aliases[i].search(regExp) != -1) {
      GmailApp.sendEmail(email, subject, message, {
        'from': aliases[i],
        cc: ccEmails,
        name: 'Paxos Private Securities Custody Operations'
      })
        ;
    }
  }
  Logger.log("End sendEmailWithoutAttachmentFromPrivateSecuritiesOps")
}

//NEW EMAIL FUNCITON THAT INCLUDES A REPLY-TO FIED
function sendEmailWithoutAttachmentFromPrivateSecuritiesOps(email, subject, message, BD) {
  //getting the emails of user currently logged in
  Logger.log("Start sendEmailWithoutAttachmentFromPrivateSecuritiesOpsTEST")

  var me = Session.getActiveUser().getEmail();
  var aliases = GmailApp.getAliases();
  var regExp = new RegExp("^privatesecuritiesops");

  var ccEmails = BD + ",privatesecuritiesops@paxos.com"

  //looping through aliases
  for (i = 0; i < aliases.length; i++) {
    //Logger.log(aliases[i].search(regExp) != -1);
    if (aliases.length > 0 && aliases[i].search(regExp) != -1) {
      GmailApp.sendEmail(email, subject, message, {
        'from': aliases[i],
        cc: ccEmails,
        name: 'Paxos Private Securities Custody Operations',
        replyTo: "privsec-custody@paxos.com"
      })
        ;
    }
  }
  Logger.log("End sendEmailWithoutAttachmentFromPrivateSecuritiesOpsTEST")
  
}

function testNewEmailLASTMINUTETEEEEST(){
  sendEmailWithoutAttachmentFromPrivateSecuritiesOpsTEST("kchichikashvili@itbit.com", "SubjectTEST","MessageTEST","dgiuntini@paxos.com")
}


//FUNCTIONS THAT RETURN EMAILS FROM SPECIFIC ONBOARDING TABS
//THERE FUNCTIONS SHOULD BE DELETED AFTER CASH & SECURITIES DIGITIZATION / REDEMPTION ARE RE-WRITTEN 
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

//function searches for issuer email in Securities Onboarding tab based off of the security fed into it
//pass the row #s into the return function 
function returnIssuerAgentEmail(securityTicker) {
  Logger.log("Start returnIssuerAgentEmail")
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(securitiesOnboarding);
  const totalRows = getLastRow(securitiesOnboarding, 'B:B');
  var notation = 'B3:M';
  var data = ss.getRange(notation).getValues();

  var securityTickerColNum = 9;
  var issuerEmailColNum = 8;
  var issuerAgentColNum = 4;
  var issuerAgentEmail = undefined;

  for (var i = 0; i < data.length; i++) {
    if (data[i][securityTickerColNum] == securityTicker) {

      issuerAgentEmail = data[i][issuerEmailColNum];
      return issuerAgentEmail;
    }
  }
  Logger.log("End returnIssuerAgentEmail")
  return issuerAgentEmail
}

//function searches for and returns customer email in Customer Onboarding tab based off of the security fed into it
function returnCustomerEmail(customerTrellisID) {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(customerOnboarding);
  const totalRows = getLastRow(customerOnboarding, 'B:B');
  var notation = 'B3:F';
  var data = ss.getRange(notation).getValues();

  var customerOnboardingTabCustomerTrellisIDColNum = 3;
  var customerOnboardingTabCustomerEmailColNum = 2;
  var customerEmail = undefined;

  for (var i = 0; i < data.length; i++) {
    if (data[i][customerOnboardingTabCustomerTrellisIDColNum] == customerTrellisID) {

      customerEmail = data[i][customerOnboardingTabCustomerEmailColNum];
      return customerEmail;
    }
  }
  return customerEmail
}

//function searches for and returns BD email based off of the BDtresllis ID fed into it
function returnBDEmail(BDTrellisID) {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(brokerDealerOnboarding);
  const totalRows = getLastRow(brokerDealerOnboarding, 'B:B');
  var notation = 'F3:H';
  var data = ss.getRange(notation).getValues();

  var BDOnboardingTabBDTrellisIDColNum = 2;
  var BDOnboardingTabBDEmailColNum = 1;
  var BDEmail = undefined;

  for (var i = 0; i < data.length; i++) {
    if (data[i][BDOnboardingTabBDTrellisIDColNum] == BDTrellisID) {

      BDEmail = data[i][BDOnboardingTabBDEmailColNum];
      return BDEmail;
    }
  }
  return BDEmail
}

//KATE TO CHECK IF WE USE THIS FOR ANYTHING, MIGHT REPLACE THE FUNCTION ABOVE WITH THIS. MAY NEED WORK
function returnIAEmail(securityCUSIP) {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(securityOnboardSheet);
  const totalRows = getLastRow(securityOnboardSheet, 'B:B');
  var notation = 'F3:K';
  var data = ss.getRange(notation).getValues();

  var securityCUSIPColNum = 5;
  var issuerAgentEmailColNum = 1;
  var issuerAgentEmail = undefined;

  for (var i = 0; i < data.length; i++) {
    if (data[i][securityCUSIPColNum] == securityCUSIP) {

      issuerAgentEmail = data[i][issuerAgentEmailColNum];
      return issuerAgentEmail;
    }
  }
  return issuerAgentEmail
}


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//SETTLING TRADES: 
//Status changes from PENDING to NEW
//Generates CSV of newly SETTLED trades sends to issuer agent via Google Drive > Outgoing folder
//Customers emailed letting them know that the trades are being settled
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



function getRowNumberOfCustomerInSpreadsheet(dataSS, customerID){
  var row = undefined
  for (i=0;dataSS.length;i++){
    //Logger.log("getting row number for customer "+customerID)
    //Logger.log("dataSS[i] "+dataSS[i])
    if(dataSS[i]==customerID){
      //not doing this for now: adding 1 since will need to look up value in an array and the array starts with 0
      row = i; //IN CASE WE HAVE PROBLEMS WITH PROCESSING OR SETTLEMENT UPDATE THIS NUMBER KATE
      //Logger.log("row of customer "+customerID +" is "+row)
      return row;
    }
  }
}


function PENDINGtoSETTLEDCL(){
  var numOfTradesSettled = 0
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(todaysTradesCL);
  const totalRows = getLastRow(todaysTradesCL, 'B:B');
  //updates master balances & balances history 
  moveCashAndSecuritiesFromHoldingAccountstoCustomersBDsClearlistPaxosSETTLEMENTFUNCTION(ss, totalRows, numOfTradesSettled, clearlistID, paxosID)
  //sends emails
  sendEmailsToSellerBuyerBDsAFTERSettlement(ss, totalRows)

}

function PENDINGtoSETTLEDSN(){
  var numOfTradesSettled = 0
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(todaysTradesSN);
  const totalRows = getLastRow(todaysTradesSN, 'B:B');
  //updates master balances & balances history 
  moveCashAndSecuritiesFromHoldingAccountstoCustomersBDsClearlistPaxosSETTLEMENTFUNCTION(ss, totalRows, numOfTradesSettled, sharenettID, paxosID)
  //sends emails
  sendEmailsToSellerBuyerBDsAFTERSettlement(ss, totalRows)

}

//moves cash and securities from holding accounts to relevant main accounts
//assigns fees to BDs, clearlist & paxos
/***
   * logic of this function:
  Buyer Net Notional = price * quantity 
  Seller Net Notional = Buyer Net Notional 

  If Buyer uses personal account:
    Holding_Buyer $$ = - (Buying Net Notional + clearlist buyer fee + BBD Fee)
  Else If Buyer uses omni:
    Holding_Omnibus_BuyerBD $$ = - (Buying Net Notional + clearlist buyer fee + BBD Fee)

  If Seller uses personal account:
    Seller $$ = + (Selling Net Notional - SBD FEE - Clearlist Seller Fee)*
  Else if Seller uses omni:
    Omnibus_SellerBD $$ = + (Selling Net Notional - SBD FEE - Clearlist Seller Fee)*
  

  BuyerBD = + BBD Fee
  Holding_Seller Sec = - Quantity
  Buyer Sec = +Quantity
  SellerBD = + SBD Fee

  Clearlist = + (Clearlist BD Fee - Paxos fee)
    if Clearlist BBD != 0 --> PAXOS FEE =+(buyer net notional * 0.001)
    if clearliat SBD !=0 --> Paxos fee = + (seller net notional *0.001)" 
  Paxos = + Paxos fee (see above for calc) 
  */
function moveCashAndSecuritiesFromHoldingAccountstoCustomersBDsClearlistPaxosSETTLEMENTFUNCTION(ss, totalRows, numOfTradesSettled, atsID, paxosID) {
  //check this variable before updating balances
  var okayToSettleColNum = 43
  var tradeSettledColNum = 44

  //get Todays Trades spreadsheet
  var notationTT = "B3:AX" + totalRows
  var data = ss.getRange(notationTT).getValues()

  //variable for updating Balances History 
  var operation = "PENDINGtoSETTLED"
  var totalRowsBalancesHistory = ssBalancesHistory.getLastRow();
  
  //identifying area of MB to be looked at for the functions that updates MB Cash
  var ssMB = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(masterBalancesSheet);
  const totalRowsMB = getLastRow(masterBalancesSheet, 'B:B');
  var notationMB = "B1:Q" + totalRowsMB + 1
  
  //use this just to get customer rows in MB
  var ssMasterBalances = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(masterBalancesSheet);
  var customerDataMasterBalances = ssMasterBalances.getRange("C:C").getValues()

  //get row number of ATS account that accrues fees in MB
  var atsAccountRowInMB = getRowNumberOfCustomerInSpreadsheet(customerDataMasterBalances, atsID)

  //get row number of Paxos account that accrues fees in MB
  var paxosAccountRowInMB = getRowNumberOfCustomerInSpreadsheet(customerDataMasterBalances, paxosID)

  //loop through the data in Todays Trades 
  for (var i = 0; i < data.length; i++) {
    //if all seller and buyer requirements met, proceed to update balance
    if (data[i][okayToSettleColNum] == "YES" && data[i][tradeStatusColNum] == "PENDING" && data[i][tradeSettledColNum] == "NotSettled" && numOfTradesSettled < maxTradesPerBatch ) {  

      //after each trade is settled, we pull the newest MB data
      var dataMB = ssMB.getRange(notationMB).getValues()

      //the following will be used to update the Balances history
      //get trade ID
      var tradeID = data[i][tradeIDColNum];
      var buyerID = data[i][buyerIDCol]
      var sellerID = data[i][sellerIDCol]
      var sellerHoldingID = "Holding_" + sellerID
      var securityCUSIP = data[i][assetCUSIPColNum]
      var buyerBDID = data[i][buyerBDIDColNumInTodaysTrades]
      var sellerBDID = data[i][sellerBDIDColNumInTodaysTrades]

      //identify if buyer uses omnibus account or personal funds
      //var buyerOmniOrPersonal = customerUsesOmnibusOrPersonal(dataCustomerOnboarding,buyerID, customerOnboardingTabCustomerCashAccountColNum)
      var buyerHoldingOmniOrPersonal = customerUsesOmnibusOrPersonal(dataCustomerOnboarding,buyerID, customerOnboardingTabCustomerHoldingCashAccountColNum)

      //net notional
      var price = data[i][priceColNum]
      var quantityShares = data[i][quantityColNum]
      var netNotional = price * quantityShares;

      //buyer cash requirement
      var atsBuyerFee = data[i][ATSBuyerFeeColNum]
      var buyerBDFee = data[i][buyerBDFeeColNum]
      var buyerNetNotional = netNotional;
      var buyerCashObligation = atsBuyerFee + buyerNetNotional + buyerBDFee;

      //seller cash requirement
      var atsSellerFee = data[i][ATSSellerFeeColNum]
      var sellerBDFee = data[i][sellerBDFeeColNum]
      var sellerNetNotional = netNotional;

      var sellerCashDue = (netNotional - atsSellerFee - sellerBDFee);


      //to be used for Balances History updating
      var buyerRow = getRowNumberOfCustomerInSpreadsheet(customerDataMasterBalances,buyerID) //using personal as it will receive securities
      Logger.log("buyer row is "+buyerRow)
      Logger.log("buyer id is" + buyerID)
      var buyerHoldingRow = getRowNumberOfCustomerInSpreadsheet(customerDataMasterBalances,buyerHoldingOmniOrPersonal) //using omni as cash comes out of holding_omni
      //var sellerRow = getRowNumberOfCustomerInSpreadsheet(customerDataMasterBalances, sellerID)
      var sellerHoldingRow = getRowNumberOfCustomerInSpreadsheet(customerDataMasterBalances,sellerHoldingID)
      var sharesQuantity = data[i][sellerSecurityQuantityColNum]
      var securityColNumInMB = data[i][sellerColInMBofSecurityColNum]
      var buyerBDRow = getRowNumberOfCustomerInSpreadsheet(customerDataMasterBalances,buyerBDID)
      var sellerBDRow = getRowNumberOfCustomerInSpreadsheet(customerDataMasterBalances,sellerBDID)

      if(checkIfEnoughCashAndAssetsInAccounts(dataMB, buyerHoldingRow, buyerCashObligation, sellerHoldingRow, securityColNumInMB, sharesQuantity)){
        Logger.log("inside if statement")

        //identify if seller uses omnibus account or personal funds. only used for cash. securities of the seller will come out of their personal account
      var sellerOmniOrPersonal = customerUsesOmnibusOrPersonal(dataCustomerOnboarding,sellerID, customerOnboardingTabCustomerCashAccountColNum)
      var sellerRow = getRowNumberOfCustomerInSpreadsheet(customerDataMasterBalances,sellerOmniOrPersonal)
      

        var paxosFee = 0
        if (atsBuyerFee > 0) {
          var paxosFeeBuyerSide = (buyerNetNotional * 0.001);
          paxosFee = paxosFee + paxosFeeBuyerSide
        }
        if (atsSellerFee > 0) {
          var paxosFeeSellerSide = (sellerNetNotional * 0.001);
          paxosFee = paxosFee + paxosFeeSellerSide
        }

        
        var atsFeeBuyerSellerMinusPaxos = (atsBuyerFee + atsSellerFee) - paxosFee

        //updateCustomerCashBalance function calls the updateBalanceHistoryNewFormat so the BH get updated automatically

        //check if the BD is the same for both the buyer and the seller
        //known issue: in case BDs are the same, the code works faster than MB get updated, so when code tries to pull the newly updated the BD MB it still sees the old value. Because of this, we're checking first if BD is the same and then only add the joint fee once 
        if (buyerBDID == sellerBDID) {
          var jointBuyerSellerBDFee = buyerBDFee + sellerBDFee
          updateCustomerCashBalance(ssMB, dataMB, buyerBDRow, jointBuyerSellerBDFee, tradeID, buyerBDID, operation,ssBalancesHistory,totalRowsBalancesHistory)
        totalRowsBalancesHistory +=1;
        }
        else {
          //Logger.log("entering Buyer BD Fee function")
          //BuyerBD = + BBD Fee
          updateCustomerCashBalance(ssMB, dataMB, buyerBDRow, buyerBDFee, tradeID, buyerBDID, operation,ssBalancesHistory,totalRowsBalancesHistory)
        totalRowsBalancesHistory +=1;
          //Logger.log("buyer bd row "+buyerBDRow)
          //Logger.log("buyer fee "+buyerBDFee)
          //Logger.log("buyerBDID "+buyerBDID)

          //Logger.log("entering Seller BD Fee function")
          //SellerBD = + SBD Fee
          updateCustomerCashBalance(ssMB, dataMB, sellerBDRow, +sellerBDFee, tradeID, sellerBDID, operation,ssBalancesHistory,totalRowsBalancesHistory)
        totalRowsBalancesHistory +=1;
          //Logger.log("seller bd row "+sellerBDRow)
          //Logger.log("seller fee "+sellerBDFee)
          //Logger.log("selledbdid "+sellerBDID)
        }



        //Holding_Buyer $$ = - (Buying Net Notional + clearlist BD fee + BBD Fee) = - buyerCashObligation
        updateCustomerCashBalance(ssMB, dataMB, buyerHoldingRow, -buyerCashObligation, tradeID, buyerHoldingOmniOrPersonal, operation,ssBalancesHistory,totalRowsBalancesHistory)
        totalRowsBalancesHistory +=1;

        //Seller $$ = + (Selling Net Notional - SBD FEE - Clearlist Seller Fee) = + sellerCashDue
        updateCustomerCashBalance(ssMB, dataMB, sellerRow, sellerCashDue, tradeID, sellerOmniOrPersonal, operation,ssBalancesHistory,totalRowsBalancesHistory)
        totalRowsBalancesHistory +=1;


        //Holding_Seller Sec = - Quantity
        updateCustomerSecurityBalance(ssMB, dataMB, sellerHoldingRow, -sharesQuantity, securityColNumInMB, tradeID, sellerHoldingID, operation, securityCUSIP,ssBalancesHistory,totalRowsBalancesHistory)
        totalRowsBalancesHistory +=1;

        //Buyer Sec = +Quantity
        updateCustomerSecurityBalance(ssMB, dataMB, buyerRow, sharesQuantity, securityColNumInMB, tradeID, buyerID, operation, securityCUSIP,ssBalancesHistory,totalRowsBalancesHistory)
        totalRowsBalancesHistory +=1;


        //ATS = + (ATS BD Fees - Paxos fee) 
        updateCustomerCashBalance(ssMB, dataMB, atsAccountRowInMB, atsFeeBuyerSellerMinusPaxos, tradeID, atsID, operation,ssBalancesHistory,totalRowsBalancesHistory)
        totalRowsBalancesHistory +=1;
        //Logger.log("clearlist row " + clearlistRowInMB)
        //Logger.log("clearlist fee " + clearlistFeeBuyerSellerMinusPaxos)
        //Logger.log("clearlist ID " + clearlistID)

        //Logger.log("entering Paxos Fee function")

        //update Paxos fee 
        updateCustomerCashBalance(ssMB, dataMB, paxosAccountRowInMB, paxosFee, tradeID, paxosID, operation,ssBalancesHistory,totalRowsBalancesHistory)
        totalRowsBalancesHistory +=1;
        //Logger.log("paxos row " + paxosRowInMB)
        //Logger.log("paxos fee " + paxosFee)
        //Logger.log("paxos ID " + paxosID)




        //change status of trade to "SETTLED"
        var pointer = i + 3

        ss.getRange(tradeStatusColumnLetter + pointer).setValue("SETTLED")
        numOfTradesSettled += 1
        Logger.log("Number of trades after adding 1 "+ numOfTradesSettled)
      }

    }
  }

  Logger.log("END moveCashAndSecuritiesFromHoldingAccountstoCustomersBDsClearlistPaxosSETTLEMENTFUNCTION")

}



//reuse this function for digitization / redemption KATE
function checkIfEnoughCashAndAssetsInAccounts(dataMB, cashRow, requiredCash, assetRow, securityColNum, requiredSecurities){
  Logger.log("Starting checkIfEnoughCashAndAssetsInAccounts")
  var sufficient = false
  var currentCashBalance = dataMB[cashRow][3];
  
  var offsetColSecurities = -2
  //Logger.log("security col num "+ securityColNum)
  var currentSecuritiesBalance = dataMB[assetRow][securityColNum + offsetColSecurities];
  //Logger.log("current CASH balance "+currentCashBalance)
  //Logger.log("require CASH balance "+ requiredCash)
  //Logger.log("current securities " + currentSecuritiesBalance)
  //Logger.log("required securities "+ requiredSecurities)
  if(currentCashBalance>=requiredCash && currentSecuritiesBalance >= requiredSecurities){
    sufficient = true
    Logger.log(sufficient)
    return sufficient
  }
  Logger.log(sufficient)
  return sufficient

}

//reuse this function for digitization / redemption KATE
function checkIfEnoughCash(dataMB, cashRow, requiredCash){
  var sufficient = false
  offsetCash = -2
  var currentCashBalance = dataMB[cashRow][3];
  Logger.log("current balance "+currentCashBalance)
  Logger.log("require balance "+ requiredCash)
  if(currentCashBalance>=requiredCash){
    sufficient = true
    Logger.log(sufficient)
    return sufficient
  }
  Logger.log(sufficient)
  return sufficient

}

function checkIfEnoughAssets(dataMB, assetRow, securityColNum, requiredSecurities){
  var sufficient = false  
  var offsetColSecurities = -2
  var currentSecuritiesBalance = dataMB[assetRow][securityColNum + offsetColSecurities];
  
  Logger.log("current securities " + currentSecuritiesBalance)
  Logger.log("required securities "+ requiredSecurities)
  if(currentSecuritiesBalance >= requiredSecurities){
    sufficient = true
    Logger.log(sufficient)
    return sufficient
  }
  Logger.log(sufficient)
  return sufficient

}




//generates CSV of newly settled trades
function convertTodaysTradeIntoCSVWithNewlySETTLEDTradesOnly(ss, totalRows) {
    Logger.log("Start convertTodaysTradeIntoCSVWithNewlySETTLEDTradesOnly")

  //creating a folder for Master Balances inside MM-DD-YYYY folder
  var timeZone = "EST";
  var baufolderid = "1s132fsm3mrJX47MLEBzGdVtcCWKLtjbt";
  var todaydatefolder = Utilities.formatDate(new Date(), timeZone, "MM-dd-yyyy");
  var todaydatefolderid = createFolder(baufolderid, todaydatefolder);
  var todaysPendingTradesFolderID = createFolder(todaydatefolderid, "Todays_Trades_Export");
  var dest_folder = DriveApp.getFolderById(todaysPendingTradesFolderID);
  var clearlist_outgoing_folder = DriveApp.getFolderById('1Myehii1D3H_sUrvgtuV-yZegp9I1-ru7');

  //converting Todays Trades (referred to as tradesLedger) to CSV
  //var ss = SpreadsheetApp.getActiveSpreadsheet();
  //var sheet = ss.getSheetByName(tradesLedger);
  var timeZone = "EST";
  var dateFormatted = Utilities.formatDate(new Date(), timeZone, "yyyyMMdd");

  var d = new Date();
  //var currentTime = d.toLocaleTimeString('en-GB'); 
  var currentTime = d.getHours();


  var fileName = ss.getName().replace(' ', '') + "_SETTLED_" + dateFormatted + "_" + currentTime + ".csv";
  // convert all available sheet data to csv format
  var csvFile = convertSETTLEDTradesInTodaysTradesToCSV(ss, totalRows);
  // create a file in the Docs List with the given name and the csv data
  var file = dest_folder.createFile(fileName, csvFile);
  var file_output2 = clearlist_outgoing_folder.createFile(fileName, csvFile);
  Logger.log("End convertTodaysTradeIntoCSVWithNewlySETTLEDTradesOnly")

  return fileName;
}


//RE-WRITE: try to make 1 function for all emails 
function convertSETTLEDTradesInTodaysTradesToCSV(ss, totalRows) {
  Logger.log("Start convertSETTLEDTradesInTodaysTradesToCSV")
  var totalRows = totalRows + 1; // add first row back 

  //var notation = "B2:R"+totalRows
  var notation = "B2:AS" + totalRows
  var data = ss.getRange(notation).getValues()

  var notation2 = "B2:R" + totalRows
  var data2 = ss.getRange(notation2).getValues()

  try {
    //var data = activeRange.getValues();
    var csvFile = undefined;

    // loop through the data in the range and build a string with the csv data
    if (data.length > 1) {
      var csv = "";
      csv += data2[0].join(",") + "\r\n";
      for (var row = 0; row < data.length; row++) {
        Logger.log("data row " + "number: " + row + "; detail: " + data[row][0])
        if (data[row][0] == "SETTLED") {
          if (data[row][42] == "YES") {
            if (data[row][43] == "") {
              var change_row_number = row + 2;

              // join each row's columns
              // add a carriage return to end of each row, except for the last one
              if (row < data2.length - 1) {
                csv += data2[row].join(",") + "\r\n";
              }
              else {
                csv += data2[row];
              }
              if (change_row_number != 2) {

                ss.getRange("AS" + change_row_number).setValue("SENT");
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
    Logger.log("Start convertSETTLEDTradesInTodaysTradesToCSV")

}


function sendEmailsToSellerBuyerBDsAFTERSettlement(ss, totalRows) {

  //Todays Trades spreadsheet 
  //var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tradesLedger);
  //const totalRows = getLastRow(tradesLedger, 'B:B');
  var notation = 'B3:AX';
  var data = ss.getRange(notation).getValues();

  //used for status of trade & verification if it's ok to send emails 
  var okayTosendEmailPostSettlement = 42;
  var tradeStatusColNum = 0;

  //trade id of trade
  var tradeIDColNum = 12; 

  //variables for IA check
  var emailSentToSellerColNum = 47;
  var emailSentToBuyerColNum = 48;


  //Customer Onboarding tab details (used to get Customer email)
  var ssCustOnboarding = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(customerOnboarding);
  var notationCO = 'B3:F';
  var dataCustomerOnboarding = ssCustOnboarding.getRange(notationCO).getValues();
  var customerOnboardingTabCustomerTrellisIDColNum = 3;
  var customerOnboardingTabCustomerEmailColNum = 2;

  //Broker Dealer Onboarding tab details (used to get BD email)
  var ssBrokerDealer = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(brokerDealerOnboarding);
  var notationDB = 'F3:H';
  var dataBrokerDealerOnboarding = ssBrokerDealer.getRange(notationDB).getValues();
  var BDOnboardingTabBDTrellisIDColNum = 2;
  var BDOnboardingTabBDEmailColNum = 1;


  //if trade is NEW, and okayToSendEmails = "YES" --> send emails to IA, Seller&&BD, Buyer&BD
  for (var i = 0; i < totalRows; i++) {
    
    if (data[i][tradeStatusColNum] == 'SETTLED' && (data[i][emailSentToSellerColNum] != "Sent" || data[i][emailSentToBuyerColNum] != 'Sent')) {

      //contents of email for all 
      var tradeID = data[i][tradeIDColNum]
      var tradeTimeColNum = 1;
      var tradeTime = data[i][tradeTimeColNum];
      var priceColNum = 6;
      var price = data[i][priceColNum];
      var buyingNetNotionalColNum = 10;
      var buyingNetNotional = data[i][buyingNetNotionalColNum];
      var sellingNetNotionalColNum = 11;
      var sellingNetNotional = data[i][sellingNetNotionalColNum];
      var quantityColNum = 7;
      var quantity = data[i][quantityColNum];
      var securityColNum = 8;
      var security = data[i][securityColNum];
      var emailSentToSellerAndBDColNum = 44;
      var emailSentToBuyerAndBDColNum = 45;
      var sellerBDFeesColNum = 14;
      var buyerBDFeesColNum = 13;
      var sellerBDFees = data[i][sellerBDFeesColNum];
      var buyerBDFees = data[i][buyerBDFeesColNum];
      var clearlistFeesBuyer = data[i][9];
      var clearlistFeesSeller = data[i][15]

  
      //Seller email 
      if(data[i][emailSentToSellerColNum] != 'Sent'){
        var sellerTrellisIDColNum = 4;
        var sellerTrellisID = data[i][sellerTrellisIDColNum];
        var sellerEmail = returnEmail(dataCustomerOnboarding, sellerTrellisID, customerOnboardingTabCustomerTrellisIDColNum, customerOnboardingTabCustomerEmailColNum)
        Logger.log("SELLER EMAIL IS "+sellerEmail)
        var sellerEmailFormatted = Utilities.formatString('%0s', sellerEmail)
        var sellerBDTrelllisIDColNum = 5;
        var sellerBDTrelllisID = data[i][sellerBDTrelllisIDColNum]
        var sellerBDEmail = returnEmail(dataBrokerDealerOnboarding, sellerBDTrelllisID, BDOnboardingTabBDTrellisIDColNum, BDOnboardingTabBDEmailColNum)
        Logger.log("SELLER BD EMAIL IS " + sellerBDEmail)
        var sellerBDEmailFormatted = Utilities.formatString('%0s', sellerBDEmail);


        var subjectSeller = "Your Sell Trade in " + security + " is settled";
        var messageSeller = "Hello,\n\nWe are writing to inform you that your trade has been settled by Paxos.\nPlease find the details of the trade below:\n\nTradeID: " + tradeID + "\nTrade Time: " + tradeTime + " (DD:MM:YYYY-HH:MM:SS)\nPrice: " + price +
          "\nQuantity: " + quantity + "\nSecurity: " + security + "\nSelling Net Notional: " +
          sellingNetNotional + "\nFees: " + (sellerBDFees + clearlistFeesSeller) + "\n\nBest,\nPaxos Private Securities Custody Operations"
        //send email to Seller&BD
        sendEmailWithoutAttachmentFromPrivateSecuritiesOps(sellerEmail, subjectSeller, messageSeller, sellerBDEmailFormatted)
        //mark "Email Sent to Seller" as YES in Todays Trades
        
        ss.getRange(i + 3, emailSentToSellerColNum + 2).setValue("Sent");
      }
    
    

      //Buyer email 
      if(data[i][emailSentToBuyerColNum] != 'Sent'){
        var buyerTrellisIDColNum = 2;
        var buyerTrellisID = data[i][buyerTrellisIDColNum];
        var buyerEmail = returnEmail(dataCustomerOnboarding, buyerTrellisID, customerOnboardingTabCustomerTrellisIDColNum, customerOnboardingTabCustomerEmailColNum)
        Logger.log("BUYER EMAIL IS "+buyerEmail)
        var buyerEmailFormatted = Utilities.formatString('%0s', buyerEmail)
        var buyerBDTrelllisIDColNum = 3;
        var buyerBDTrelllisID = data[i][buyerBDTrelllisIDColNum]
        
        var buyerBDEmail = returnEmail(dataBrokerDealerOnboarding, buyerBDTrelllisID, BDOnboardingTabBDTrellisIDColNum, BDOnboardingTabBDEmailColNum)
        Logger.log("BUYER BD EMAIL IS " + sellerBDEmail)
        var buyerBDEmailFormatted = Utilities.formatString('%0s', buyerBDEmail);


        var subjectBuyer = "Your Buy Trade in " + security + " is settled";
        var messageBuyer = "Hello,\n\nWe are writing to inform you that your trade has been settled by Paxos.\nPlease find the details of the trade below:\n\nTradeID: " + tradeID + "\nTrade Time: " + tradeTime + " (DD:MM:YYYY-HH:MM:SS)\nPrice: " + price +
          "\nQuantity: " + quantity + "\nSecurity: " + security + "\nBuying Net Notional: " +
          buyingNetNotional + "\nFees: " + (buyerBDFees + clearlistFeesBuyer) + "\n\nBest,\nPaxos Private Securities Custody Operations"
        sendEmailWithoutAttachmentFromPrivateSecuritiesOps(buyerEmailFormatted, subjectBuyer, messageBuyer, buyerBDEmailFormatted)

        ss.getRange(i + 3, emailSentToBuyerColNum + 2).setValue("Sent");
      }

    }

  }
    Logger.log("End sendEmailsToSellerBuyerBDsAFTERSettlement")

}


////


/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//CONVERTS ANY TAB INTO CSV
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

//Converts data in sheet to csv format. 
//used to generate clearlist positions & gts positions csvs
function convertRangeToCsvFile_(sheet) {
  // get available data range in the spreadsheet
  var activeRange = sheet.getDataRange();
  try {
    var data = activeRange.getValues();
    var csvFile = undefined;

    // loop through the data in the range and build a string with the csv data
    if (data.length > 1) {
      var csv = "";
      for (var row = 0; row < data.length; row++) {
        for (var col = 0; col < data[row].length; col++) {
          if (data[row][col].toString().indexOf(",") != -1) {
            data[row][col] = "\"" + data[row][col] + "\"";
          }
        }

        // join each row's columns
        // add a carriage return to end of each row, except for the last one
        if (row < data.length - 1) {
          csv += data[row].join(",") + "\r\n";
        }
        else {
          csv += data[row];
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

//////////////////////////////////////////////////////////////////////////////////////////////////////////
// MOVE TRADES FROM TODAYS TRADES TO TRADING HISTORY 
// RE-WRITE, SO IT CAN MOVE TRADES FROM ANY TAB INTO ANY OTHER TAB
//////////////////////////////////////////////////////////////////////////////////////////////////////////

//This function is to bring data from Trade Create tab and append to Todays trades and clear Trade Create after. 
function appendTradingHistory() {

  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tradesLedger);
  var totalRows = getLastRow(tradesLedger, 'B:B');
  var totalRows = totalRows + 1;
  Logger.log(totalRows);
  var totalRows1 = getLastRow(tradingHistory, 'A:A');
  Logger.log(totalRows1);


  var notation = "B3:R" + totalRows
  var notation3 = "AI3:AK" + totalRows

  
  var data = ss.getRange(notation).getValues()
  var data3 = ss.getRange(notation3).getValues()

  
  var ss2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tradingHistory);
  Logger.log(data.length);
  for (var i = 0; i < data.length; i++) {
    ss2.getRange(totalRows1 + 1, 1, 1, 17).setValues([data[i]]);
    
    ss2.getRange(totalRows1 + 1, 20, 1, 3).setValues([data3[i]]);
    totalRows1 += 1;
    //ts.appendRow(data[i]); 
  }


  // clear formula that set "SENT" to specific column
  //ss.getRange("S3:T500").clearContent();
  ss.getRange("AI3:AK500").clearContent();
  ss.getRange("AS3:AU500").clearContent();


  //Clear data from Todays Trades. Starts clearing at B3 so as to not erase the header. Range can be adjusted in global variable section 
  ss.getRange(rangeToClearInTodaysTrades).clearContent();

}


/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// SENDING TRADING HISTORY TO CLEARLIST
// RE-WRITE, SO IT CAN SEND A CSV OF ANY TAB TO RELEVANT GOOGLE DRIVE FOLDER
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//START OF NEW MORE REUSABLE FUNCTION



function convertRangeToCsvFileTradeHistory_() {

  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tradingHistory);
  var totalRows = getLastRow(tradingHistory, 'A:A');
  var totalRows = totalRows + 1; // add first row back 
  Logger.log(totalRows)

  //var notation = "B2:R"+totalRows
  var notation = "A1:AA" + totalRows
  var notation2 = "A1:Q" + totalRows
  var data = ss.getRange(notation).getValues()
  var data2 = ss.getRange(notation2).getValues()
  // get available data range in the spreadsheet

  try {
    //var data = activeRange.getValues();
    var csvFile = undefined;

    // loop through the data in the range and build a string with the csv data
    if (data.length > 1) {
      var csv = "";
      Logger.log(data2[0])
      //csv += data2[0].join(",") + "\r\n";
      for (var row = 0; row < data.length; row++) {
        for (var col = 0; col < data[row].length; col++) {
          if (data[row][col].toString().indexOf(",") != -1) {
            data[row][col] = "\"" + data[row][col] + "\"";
          }
        }

        // join each row's columns
        // add a carriage return to end of each row, except for the last one

        if (row < data2.length - 1) {
          csv += data2[row].join(",") + "\r\n";
        }
        else {
          csv += data2[row];
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


// Downloading Trade History into CSV and saving ot Google Drive 
function convertTradingHistoryToCSV() {
  var timeZone = "EST";
  var baufolderid = "1s132fsm3mrJX47MLEBzGdVtcCWKLtjbt";
  var todaydatefolder = Utilities.formatDate(new Date(), timeZone, "MM-dd-yyyy");
  var todaydatefolderid = createFolder(baufolderid, todaydatefolder);
  var TradeHistoryFolderID = createFolder(todaydatefolderid, "Trading_History_Export");
  var dest_folder = DriveApp.getFolderById(TradeHistoryFolderID);
  //var clearlist_outgoing_folder = DriveApp.getFolderById('1Myehii1D3H_sUrvgtuV-yZegp9I1-ru7');

  //converting trade history sheet into csv
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(tradingHistory);

  var dateFormatted = Utilities.formatDate(new Date(), timeZone, "yyyyMMdd");
  var d = new Date();
  //var currentTime = d.toLocaleTimeString(); 
  var currentTime = d.getHours();

  //var fileName = "Trading_History" + "_"+ dateFormatted + ".csv";
  var fileName = "Trading_History" + "_" + dateFormatted + "_" + currentTime + ".csv";

  // convert all available sheet data to csv format
  //var csvFile = convertRangeToCsvFile_(sheet);
  var csvFile = convertRangeToCsvFileTradeHistory_(sheet);

  // create a file in the Docs List with the given name and the csv data
  var file = dest_folder.createFile(fileName, csvFile);
  //var file_output2 = clearlist_outgoing_folder.createFile(fileName, csvFile);
  return fileName;


}

//moves pending trades back to todays trades
//re-write: so it can move trades from any trading history to any todays trades tab
function movePENDINGTradesFromTradingHistoryToTodaysTrades() {

  // download all trade in trading history tab
  var timeZone = "EST";
  var baufolderid = "1s132fsm3mrJX47MLEBzGdVtcCWKLtjbt";
  var todaydatefolder = Utilities.formatDate(new Date(), timeZone, "MM-dd-yyyy");
  var todaydatefolderid = createFolder(baufolderid, todaydatefolder);
  var TradeHistoryFolderID = createFolder(todaydatefolderid, "Trading_History_Export");
  var dest_folder = DriveApp.getFolderById(TradeHistoryFolderID);
  var clearlist_outgoing_folder = DriveApp.getFolderById('1Myehii1D3H_sUrvgtuV-yZegp9I1-ru7');

  //converting trade history sheet into csv
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(tradingHistory);
  var sheet2 = ss.getSheetByName(tradesLedger);

  var totalRows = getLastRow(tradingHistory, 'A:A');
  Logger.log(totalRows);
  var notation = "A2:Q" + totalRows + 1;
  var data = sheet.getRange(notation).getValues();
  //var notation2 = "R2:S"+totalRows+1;
  //var data2 = sheet.getRange(notation2).getValues();  
  var notation3 = "T2:V" + totalRows + 1;
  var data3 = sheet.getRange(notation3).getValues();
  //var notation4 = "Y2:AA"+totalRows+1;
  //var data4 = sheet.getRange(notation4).getValues();
  Logger.log("123");

  var dateFormatted = Utilities.formatDate(new Date(), timeZone, "yyyyMMdd");
  var d = new Date();
  //var currentTime = d.toLocaleTimeString(); 
  var currentTime = d.getHours();

  var fileName = "TradingHistory" + "_" + dateFormatted + "_" + currentTime + ".csv";

  // convert all available sheet data to csv format
  //var csvFile = convertRangeToCsvFile_(sheet);
  var csvFile = convertRangeToCsvFileTradeHistory_(sheet);
  // create a file in the Docs List with the given name and the csv data
  var file = dest_folder.createFile(fileName, csvFile);
  var file_output2 = clearlist_outgoing_folder.createFile(fileName, csvFile);
  //return fileName;

  // Move only the PENDING trades from Trade History back to Todays Trades
  var totalRows1 = getLastRow(tradesLedger, 'B:B');
  Logger.log(totalRows1);


  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == "PENDING" || data[i][0] == "NEW") {
      sheet2.getRange(totalRows1 + 1, 2, 1, 17).setValues([data[i]]);
      sheet2.getRange(totalRows1 + 1, 35, 1, 3).setValues([data3[i]]);
      totalRows1 += 1;
    }
  }

  // wipe all trade from trade history
  var rangeToClearInTradeHistory = 'A2:AA500';
  Logger.log("123");
  sheet.getRange(rangeToClearInTradeHistory).clearContent();
}


////////////////////////////////////////////////////////////////////////////////////////////
//EMAILS
//THESE EMAILS WORK, HOWEVER NEED TO BE REVIEWED, MIGHT CREATE 1 FUNCTION THAT TAKES IN THE DETAILS OF THE EMAIL
////////////////////////////////////////////////////////////////////////////////////////////

//email sent when a customer is onboarded  
function sendOnboardingEmail() {
  //Customer Onboarding tab details (used to get Customer email)
  var ssCustOnboarding = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(customerOnboarding);
  var totalRows = getLastRow(customerOnboarding,"B:B")
  var notationCO = 'B3:Z'+totalRows+1
  var dataCustomerOnboarding = ssCustOnboarding.getRange(notationCO).getValues();
  var customerOnboardingTabCustomerTrellisIDColNum = 3;
  var customerOnboardingTabCustomerEmailColNum = 2;
  var sentStatusColNum = 24
  var onboardedStatusColNum = 22

  for(var i=0;i<dataCustomerOnboarding.length; i++){
    if(dataCustomerOnboarding[i][onboardedStatusColNum]=="Already Onboarded"&& dataCustomerOnboarding[i][sentStatusColNum]!="Sent"){
      Logger.log(dataCustomerOnboarding[i][sentStatusColNum] + "i is "+ i)
      var trellisID = dataCustomerOnboarding[i][customerOnboardingTabCustomerTrellisIDColNum]
      var recepientEmailAddress = returnEmail(dataCustomerOnboarding, trellisID, customerOnboardingTabCustomerTrellisIDColNum, customerOnboardingTabCustomerEmailColNum)
      var message = "Hello,\n\nYour Paxos custody account has been approved. Please find the details of your account below, including instructions for depositing cash and/or securities and instructions to request a withdrawal.\n\nYour Customer ID is: " + trellisID + "\n\nTo deposit cash, please send a Domestic Wire Transfer from a US bank account with the following instructions:\n\n\t\tBank Name: BMO Harris Bank NA\n\t\tBank Routing Number: 071000288\n\t\tBank Address:\n\t\t\tBMO HARRIS BANK\n\t\t\t111 West Monroe Street\n\t\t\tChicago, IL 60603\n\n\t\tPaxos Business Address:\n\t\t\t450 Lexington Avenue\n\t\t\tSuite 3952\n\t\t\tNew York, NY 10163\n\n\t\tBeneficiary Name: PAXOS TRUST COMPANY LLC - for the benefit of Clearlist customers\n\t\tBeneficiary Account Number: 3738200\n\t\tMessage to Recipient: Customer ID\n\t\t\tNumber is found on your Clearlist Profile and at the top of this email.\n\t\t\tPlease note, it is mandatory to include your Customer ID on your wire transfer so we can credit your account accordingly.\n\nTo deposit securities, please send an email to privsec-custody@paxos.com with a request to initiate the deposit process with your Issuer.\n\nTo request a withdrawal of cash or securities, please send an email to privsec-custody@paxos.com with a request for withdrawal. Please note that cash withdrawals will only be sent to the originating bank from which the initial deposit was sent, and securities withdrawals will only be sent to the Issuer’s custody.\n\nProcessing time for cash transfers is usually one business day, whereas processing time for securities transfers may vary depending on the Issuer. For any questions please contact privsec-custody@paxos.com.\n\nThank you,\nPaxos Private Securities Custody"

      var subject = "Your Paxos Private Securities Custody Account"
      var cc = "privatesecuritiesops@paxos.com"
      sendEmailWithoutAttachmentFromPrivateSecuritiesOps(recepientEmailAddress, subject, message, cc)

      var pointer = i+3  
      ssCustOnboarding.getRange("Z"+pointer).setValue("Sent")
      
    }
  }
  
}

function testSendOnboardingEmail() {
  sendOnboardingEmail("kchichikashvili@paxos.com", "testID", "privatesecuritiesops@paxos.com")
}

//email sent when requesting to digitize securities, cc legal
function sendSecuritiesDigitizationCapTableUpdateEmail(recepientEmailAddress, customerID, numberOfShares, security, cc) {

  var message = "Hello,\n\nWe’ve received a deposit request of " + numberOfShares + " shares of " + security + " from " + customerID + ". Please let us know once your cap table has been updated to indicate the shares reside in Paxos custody.\n\nThank you,\nPaxos Private Securities Custody"

  var subject = "Request for Private Securities Deposit / Cap Table Update:"
  var cc = "ctierno@paxos.com,epalaia@paxos.com"
  sendEmailWithoutAttachmentFromPrivateSecuritiesOps(recepientEmailAddress, subject, message, cc)
}

function testSecuritiesDigitizationCapTableUpdateEmail() {
  sendSecuritiesDigitizationCapTableUpdateEmail("kchichikashvili@paxos.com", "testID", 5, "TESTSECURITY", "privatesecuritiesops@paxos.com")
}

//Used in Sec Digitization, sends Processing emails to Customer & IA 
function securitiesDigitizationCustomerAndIAEmails() {
  //securities create tab
  var ssSecCreate = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(secCreate);
  const totalRowsSC = getLastRow(secCreate, 'C:C');
  var notationSecCreate = "B2:AC" + totalRowsSC + 1
  var dataSecCreate = ssSecCreate.getRange(notationSecCreate).getValues()


  var uniqueIDColNum = 0;
  var amountColNum = 6;
  var trellisIDColNum = 3;
  var securityNameColNum = 7;
  var okayToSendEmails = 13;

  var customerEmailedProcessingColLetter = "P";
  var issuerAgentEmailedColLetter = "Q";

  //Customer Onboarding tab details (used to get Customer email)
  var ssCustOnboarding = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(customerOnboarding);
  var notationCO = 'B3:F';
  var dataCustomerOnboarding = ssCustOnboarding.getRange(notationCO).getValues();
  var customerOnboardingTabCustomerTrellisIDColNum = 3;
  var customerOnboardingTabCustomerEmailColNum = 2;

  //Securities Onboarding tab details (used to get issuer agent email)
  var ssSecuritiesOnboarding = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(securitiesOnboarding);
  var notationSO = 'J3:K';
  var dataSecuritiesOnboarding = ssSecuritiesOnboarding.getRange(notationSO).getValues();
  var securitiesOnboardingTabSecurityColNum = 1;
  var issuerAgentEmailColNum = 0;
  

  //loop through the sec create data 
  for (var i = 0; i < dataSecCreate.length; i++) {
    if (dataSecCreate[i][okayToSendEmails] == "YES") {
      var uniqueID = dataSecCreate[i][uniqueIDColNum];
      var trellisID = dataSecCreate[i][trellisIDColNum];
      var sharesQuantity = dataSecCreate[i][amountColNum];
      var securityCUSIP = dataSecCreate[i][securityNameColNum]


      var pointer = i + 2
      var customerEmailAddress = returnEmail(dataCustomerOnboarding, trellisID, customerOnboardingTabCustomerTrellisIDColNum, customerOnboardingTabCustomerEmailColNum)
      //var customerEmailAddress = returnCustomerEmail(trellisID)
      //var issuerEmailAddress = returnIAEmail(securityCUSIP)
      var issuerEmailAddress = returnEmail(dataSecuritiesOnboarding, securityCUSIP, securitiesOnboardingTabSecurityColNum, issuerAgentEmailColNum)
      Logger.log("customer email is " + customerEmailAddress)
      Logger.log("data 0 "+dataSecuritiesOnboarding[0])
      Logger.log("data 1 "+dataSecuritiesOnboarding[1])
      Logger.log("data 2 "+dataSecuritiesOnboarding[2])


      Logger.log("issuerEmail is "+ issuerEmailAddress)

      sendSecuritiesDigitizationCapTableUpdateEmail(issuerEmailAddress, trellisID, sharesQuantity, securityCUSIP, privateSecuritiesOpsEmail)
      ssSecCreate.getRange(issuerAgentEmailedColLetter + pointer).setValue("Sent")
      //THIS DIGITIZATION FUNCTION NEEDS EDITING FOR THE WORDING 
      sendSecuritiesDigitizationInProgressEmailToCustomer(customerEmailAddress, sharesQuantity, securityCUSIP, privateSecuritiesOpsEmail, uniqueID)
      ssSecCreate.getRange(customerEmailedProcessingColLetter + pointer).setValue("Sent")


    }

  }



}

function sendSecuritiesDigitizationInProgressEmailToCustomer(recepientEmailAddress, numberOfShares, security, cc, actionID) {
  var message = "Hello,\n\nWe’ve received your deposit request of " + numberOfShares + " shares of " + security + ". We have reached out to your Issuer and will confirm once your request is approved.\n\nTransfer ID: " + actionID + "\n\nThank you,\nPaxos Private Securities Custody"
  var subject = "Your Private Securities Deposit: In Progress"
  sendEmailWithoutAttachmentFromPrivateSecuritiesOps(recepientEmailAddress, subject, message, cc)

}

function sendSecuritiesDigitizationCompleteEmail(recepientEmailAddress, numberOfShares, security, cc, actionID) {
  var message = "Hello,\n\nWe can confirm that your recent securities deposit for " + numberOfShares + " shares of " + security + " security has been completed and is now credited to your Paxos custody account. You should see this reflected in your ClearList balance shortly.\n\nTransfer ID: " + actionID + "\n\nThank you,\n\nPaxos Private Securities Custody"
  var subject = "Your Private Securities Deposit: Completed"
  sendEmailWithoutAttachmentFromPrivateSecuritiesOps(recepientEmailAddress, subject, message, cc)
}

function testSendSecuritiesDigitizationCompleteEmail() {
  sendSecuritiesDigitizationCompleteEmail("kchichikashvili@paxos.com", 10, "TESTSECURITY", "privatesecuritiesops@paxos.com")
}

//sent when requesting Issuer Agent to approve customer redemption
function sendSecuritiesRedemptionCapTableUpdateEmail(recepientEmailAddress, customerID, numberOfShares, security, cc, actionID) {
  var message = "Hello,\n\nWe’ve received a withdrawal request of " + numberOfShares + " shares of " + security + " from " + customerID + ". We have validated that the customer has no open trade executions and Paxos has moved these shares into the customer’s holding account. Please let us know once your cap table has been updated to indicate the shares have been returned to your custody.\n\nTransfer ID: " + actionID + "\n\nThank you,\nPaxos Private Securities Custody"
  var subject = "Request for Private Securities Withdrawal / Cap Table Update:"
  sendEmailWithoutAttachmentFromPrivateSecuritiesOps(recepientEmailAddress, subject, message, cc)
}



function sendSecuritiesRedemptionInProgressEmailToCustomer(recepientEmailAddress, numberOfShares, security, cc, actionID) {
  var message = "Hello,\n\nWe’ve received your withdrawal request of " + numberOfShares + " shares of " + security + ". We have reached out to your Issuer and will confirm once your shares have been returned to the Issuer’s custody.\n\nTransfer ID: " + actionID + "\n\nThank you,\nPaxos Private Securities Custody"
  var subject = "Your Private Securities Withdrawal: In Progress"
  sendEmailWithoutAttachmentFromPrivateSecuritiesOps(recepientEmailAddress, subject, message, cc)
}



function sendSecuritiesRedemptionCompleteEmailToCustomer(recepientEmailAddress, numberOfShares, security, cc, actionID) {
  var message = "Hello,\n\nWe can confirm that your request for withdrawal of " + numberOfShares + " shares of " + security + " has been completed. These shares now reside with the Issuer and are no longer in Paxos' custody.\n\nTransfer ID: " + actionID + "\n\nThank you,\nPaxos Private Securities Custody"
  var subject = "Your Private Securities Withdrawal: Completed"
  sendEmailWithoutAttachmentFromPrivateSecuritiesOps(recepientEmailAddress, subject, message, cc)
}



function sendCashDigitizationEmail(recepientEmailAddress, cashQuantity, cc, actionID) {
  var message = "Hello,\n\nWe can confirm that your recent cash deposit for $" + cashQuantity + " has been completed and is now credited to your Paxos custody account. You should see this reflected in your ClearList balance shortly.\n\nTransfer ID: " + actionID + "\n\nThank you,\nPaxos Private Securities Custody"
  var subject = "Your Cash Deposit: Completed"
  sendEmailWithoutAttachmentFromPrivateSecuritiesOps(recepientEmailAddress, subject, message, cc)
}


function sendCashRedemptionEmail(recepientEmailAddress, cashQuantity, cc, actionID) {
  var message = "Hello,\n\nWe can confirm that your recent cash withdrawal for $" + cashQuantity + " has been processed and a wire transfer has been sent to your US bank account. You should receive the transfer shortly.\n\nTransfer ID: " + actionID + "\n\nThank you,\nPaxos Private Securities Custody"
  var subject = "Your Cash Withdrawal: Completed"
  sendEmailWithoutAttachmentFromPrivateSecuritiesOps(recepientEmailAddress, subject, message, cc)
}





///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

//SECURITIES ONBOARDING 
////////////////////////////////////////////////////////////////////////////////////////////////////
// add new security to ledger (master balance sheet, balance history)
function addNewSecurityToLedger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(masterBalancesSheet);
  var ss1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(securityOnboardSheet);
  var ss2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(balancesHistory);
  var last_rows = getLastRow(securityOnboardSheet, "B:B");
  var last_rows1 = getLastRow(masterBalancesSheet, "B:B");

  var secOnboardingTabOkToOnboardColLetter = "O";
  var secOnboardingTabUniqueColLetter = "P";
  var secOnboardingTabNewTickerColLetter = "K";
  var secOnboardingTabDateTimeColLetter = "R";
  var secOnboardingTabOnboardedColLetter = "Q";



  for (var i = 3; i < last_rows + 2; i++) {

    if (ss1.getRange(secOnboardingTabOkToOnboardColLetter + i).getValue() == "YES" && ss1.getRange(secOnboardingTabUniqueColLetter + i).getValue() == "Unique") {

      var new_ticker = ss1.getRange(secOnboardingTabNewTickerColLetter + i).getValue();

      // update master balance sheet by include new ticker as column
      var last_column = ss.getLastColumn();
      ss.insertColumns(last_column + 1);
      ss.getRange(2, last_column + 1).setValue(new_ticker);

      // update balance history by include new ticker as three column (previous, delta, new)
      var last_column1 = ss2.getLastColumn();
      //ss2.insertColumns(last_column1+1); 
      //ss2.insertColumns(last_column1+2);
      //ss2.insertColumns(last_column1+3);

      //ss2.getRange(2,last_column1+1).setValue(new_ticker);
      //ss2.getRange(2,last_column1+1).setValue("previous"+new_ticker);
      //ss2.getRange(2,last_column1+2).setValue("delta"+new_ticker);
      //ss2.getRange(2,last_column1+3).setValue("new"+new_ticker);      

      // update Securities Onboarding column L to "YES" and add date time to col M
      ss1.getRange(secOnboardingTabOkToOnboardColLetter + i).setValue("Already Onboarded");
      ss1.getRange(secOnboardingTabOnboardedColLetter + i).setValue("Onboarded");
      var currentTime = new Date();
      ss1.getRange(secOnboardingTabDateTimeColLetter + i).setValue(currentTime);

    }
  }

  // set 0 for new added ticker in master balance sheet
  for (var i = 3; i < last_rows1 + 1; i++) {
    ss.getRange(i, last_column + 1).setValue(0);
  }

}

function getLastColumn(ss, rowRange){
  var data = ss.getRange(rowRange).getValues();  
  var count = 0
  for(i=0; i<data.length;i++){
    for(j=0; j<data[i].length;j++){
      if (data[i][j]!=""){
        count += 1;
      }
    }
  }
  Logger.log("count " + count)
  return count
}

function testGetLastCOl(){
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(masterBalancesSheet);
  getLastColumn(ss,"A2:S2")

}

function addNewSecurityToLedgerNEW() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(masterBalancesSheet);
  var ss1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(securityOnboardSheet);
  var ss2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(balancesHistory);
  var last_rows = getLastRow(securityOnboardSheet, "B:B");
  var last_rows1 = getLastRow(masterBalancesSheet, "B:B");

  var secOnboardingTabOkToOnboardColLetter = "O";
  var secOnboardingTabUniqueColLetter = "P";
  var secOnboardingTabNewTickerColLetter = "K";
  var secOnboardingTabDateTimeColLetter = "R";
  var secOnboardingTabOnboardedColLetter = "Q";



  for (var i = 3; i < last_rows + 2; i++) {

    if (ss1.getRange(secOnboardingTabOkToOnboardColLetter + i).getValue() == "YES" && ss1.getRange(secOnboardingTabUniqueColLetter + i).getValue() == "Unique") {

      var new_ticker = ss1.getRange(secOnboardingTabNewTickerColLetter + i).getValue();

      // update master balance sheet by include new ticker as column
      //var last_column = ss.getLastColumn();
      //Logger.log("LAST IS" +last_column)
      var last_column = getLastColumn(ss,"A2:S2")
      ss.insertColumns(last_column + 1);
      ss.getRange(2, last_column + 1).setValue(new_ticker);

      // update balance history by include new ticker as three column (previous, delta, new)
      //var last_column1 = ss2.getLastColumn();
      //ss2.insertColumns(last_column1+1); 
      //ss2.insertColumns(last_column1+2);
      //ss2.insertColumns(last_column1+3);

      //ss2.getRange(2,last_column1+1).setValue(new_ticker);
      //ss2.getRange(2,last_column1+1).setValue("previous"+new_ticker);
      //ss2.getRange(2,last_column1+2).setValue("delta"+new_ticker);
      //ss2.getRange(2,last_column1+3).setValue("new"+new_ticker);      

      // update Securities Onboarding column L to "YES" and add date time to col M
      //ss1.getRange(secOnboardingTabOkToOnboardColLetter + i).setValue("Already Onboarded");
      ss1.getRange(secOnboardingTabOnboardedColLetter + i).setValue("Onboarded");
      var currentTime = new Date();
      ss1.getRange(secOnboardingTabDateTimeColLetter + i).setValue(currentTime);

    }
  }

  // set 0 for new added ticker in master balance sheet
  for (var i = 3; i < last_rows1 + 1; i++) {
    ss.getRange(i, last_column + 1).setValue(0);
  }
  

}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// CUSTOMER ONBOARDING (Chloe's structure, re-written by Kate)
/////////////////////////////////////////////////////////////////////////////////
// add Regular line for a customer with their KYC Status, Trellis ID, BD

function addNewCustomerToLedger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(customerOnboarding);
  var ss1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(masterBalancesSheet);
  var last_rows = getLastRow(customerOnboarding, "B:B");
  var last_rows1 = getLastRow(masterBalancesSheet, "B:B");
  var last_column = ss1.getLastColumn();
  var next_row1 = last_rows1 + 1;

  //customer Onboarding tab columns 
  var customerOnboardingTabCustomerEmailColLetter = "D";
  var customerOnboardingTabTrellisIDColLetter = "E";
  var customerOnboardingTabBrokerDealerColLetter = "F";
  var customerOnboardingTabOKtoCreateCustomerColLetter = "R";
  var customerOnboardingTabDateTimeColLetter = "S";
  var customerOnboardingTabEmailSentColLetter = "T";

  //master balances tab columns 
  var masterBalancesKYCStatusColLetter = "B";
  var masterBalancesTrellisIDColLetter = "C";
  var masterBalancesBrokerDealerColLetter = "D";
  //Logger.log("lastrow "+last_rows)
  for (var i = 3; i < last_rows + 2; i++) {
    //Logger.log("i outside is: "+i)
    //check that all data is populated for the customer & customer has not already been onboarded 
    if (ss.getRange(customerOnboardingTabOKtoCreateCustomerColLetter + i).getValue() == "YES") {
      //Logger.log("i inside is: "+i)

      //in Customer Onboarding identifies trellis ID & Broker Dealer
      var Trellis_ID = ss.getRange(customerOnboardingTabTrellisIDColLetter + i).getValue();
      var Broker_Dealer = ss.getRange(customerOnboardingTabBrokerDealerColLetter + i).getValue();
      //Logger.log(Trellis_ID);
      //Logger.log(Broker_Dealer);

      // in Master Balances adds a row for customer by using their Trellis ID & Broker Dealer
      ss1.getRange(masterBalancesKYCStatusColLetter + next_row1).setValue("OK");
      ss1.getRange(masterBalancesTrellisIDColLetter + next_row1).setValue(Trellis_ID);
      ss1.getRange(masterBalancesBrokerDealerColLetter + next_row1).setValue(Broker_Dealer);


      // fill in zero for new added customer
      for (var j = 5; j < last_column + 1; j++) {
        Logger.log(ss1.getRange(next_row1, j));
        ss1.getRange(next_row1, j).setValue(0);
      }

      next_row1 = next_row1 + 1;

      // in Master Balances creates another row in the format "Holding_CustomerTrellisID". This row will be used for redemption of securities 
      ss1.getRange(masterBalancesKYCStatusColLetter + next_row1).setValue("OK");
      ss1.getRange(masterBalancesTrellisIDColLetter + next_row1).setValue("Holding_" + Trellis_ID);
      ss1.getRange(masterBalancesBrokerDealerColLetter + next_row1).setValue(Broker_Dealer);

      // fill in zero for new added holding
      for (var j = 5; j < last_column + 1; j++) {
        Logger.log(ss1.getRange(next_row1, j));
        ss1.getRange(next_row1, j).setValue(0);
      }



      next_row1 = next_row1 + 1;

      // update Customer Onboarding add the date and time of the onboarding  
      var currentTime = new Date();
      ss.getRange(customerOnboardingTabDateTimeColLetter + i).setValue(currentTime);

      var recepientEmailAddress = ss.getRange(customerOnboardingTabCustomerEmailColLetter + i).getValue()
      //Logger.log("email address "+recepientEmailAddress)

      sendOnboardingEmail(recepientEmailAddress, Trellis_ID, privateSecuritiesOpsEmail)
      ss.getRange(customerOnboardingTabEmailSentColLetter + i).setValue("Sent");
    }
  }





}

/// BROKER DEALER ONBOARDING 
//////////////////////////////////////////////////////////////////////////////////
function addNewBDToLedger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(brokerDealerOnboarding);
  var ss1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(masterBalancesSheet);
  var last_rows = getLastRow(brokerDealerOnboarding, "B:B");
  var last_rows1 = getLastRow(masterBalancesSheet, "B:B");
  var next_row1 = last_rows1 + 1;
  var last_column = ss1.getLastColumn();

  //customer Onboarding tab columns 
  var BDOnboardingTabTrellisIDColLetter = "H";
  //var customerOnboardingTabBrokerDealerColLetter = "F";
  var BDOnboardingTabOKtoOnboardColLetter = "T";
  var BDOnboardingTabDateTimeColLetter = "V";

  //master balances tab columns 
  var masterBalancesKYCStatusColLetter = "B";
  var masterBalancesTrellisIDColLetter = "C";
  var masterBalancesBrokerDealerColLetter = "D";
  //Logger.log("lastrow "+last_rows)
  for (var i = 3; i < last_rows + 2; i++) {
    //Logger.log("i outside is: "+i)
    //check that all data is populated for the customer & customer has not already been onboarded 
    if (ss.getRange(BDOnboardingTabOKtoOnboardColLetter + i).getValue() == "YES") {
      //Logger.log("i inside is: "+i)

      //in Customer Onboarding identifies trellis ID & Broker Dealer
      var Trellis_ID = ss.getRange(BDOnboardingTabTrellisIDColLetter + i).getValue();
      var Broker_Dealer = "NA";

      // in Master Balances adds a row for customer by using their Trellis ID & Broker Dealer
      ss1.getRange(masterBalancesKYCStatusColLetter + next_row1).setValue("OK");
      ss1.getRange(masterBalancesTrellisIDColLetter + next_row1).setValue(Trellis_ID);
      ss1.getRange(masterBalancesBrokerDealerColLetter + next_row1).setValue(Broker_Dealer);

      for (var j = 5; j < last_column + 1; j++) {
        Logger.log(ss1.getRange(next_row1, j));
        ss1.getRange(next_row1, j).setValue(0);
      }
      next_row1 = next_row1 + 1;

      // in Master Balances creates another row in the format "Holding_CustomerTrellisID". This row will be used for redemption of securities 
      ss1.getRange(masterBalancesKYCStatusColLetter + next_row1).setValue("OK");
      ss1.getRange(masterBalancesTrellisIDColLetter + next_row1).setValue("Holding_" + Trellis_ID);
      ss1.getRange(masterBalancesBrokerDealerColLetter + next_row1).setValue(Broker_Dealer);

      for (var j = 5; j < last_column + 1; j++) {
        Logger.log(ss1.getRange(next_row1, j));
        ss1.getRange(next_row1, j).setValue(0);
      }
      next_row1 = next_row1 + 1;

      // update Customer Onboarding add the date and time of the onboarding  
      var currentTime = new Date();
      ss.getRange(BDOnboardingTabDateTimeColLetter + i).setValue(currentTime);
      ss.getRange(BDOnboardingTabOKtoOnboardColLetter + i).setValue("Already Onboarded");
    }
  }
}





/**
 * Unpivot a pivot table of any size.
 *
 * @param {A1:D30} data The pivot table.
 * @param {1} fixColumns Number of columns, after which pivoted values begin. Default 1.
 * @param {1} fixRows Number of rows (1 or 2), after which pivoted values begin. Default 1.
 * @param {"city"} titlePivot The title of horizontal pivot values. Default "column".
 * @param {"distance"[,...]} titleValue The title of pivot table values. Default "value".
 * @return The unpivoted table
 * @customfunction
 */
function unpivot(data, fixColumns, fixRows, titlePivot, titleValue) {
  var fixColumns = fixColumns || 1; // how many columns are fixed
  var fixRows = fixRows || 1; // how many rows are fixed
  var titlePivot = titlePivot || 'column';
  var titleValue = titleValue || 'value';
  var ret = [], i, j, row, uniqueCols = 1;

  // we handle only 2 dimension arrays
  if (!Array.isArray(data) || data.length < fixRows || !Array.isArray(data[0]) || data[0].length < fixColumns)
    throw new Error('no data');
  // we handle max 2 fixed rows
  if (fixRows > 2)
    throw new Error('max 2 fixed rows are allowed');

  // fill empty cells in the first row with value set last in previous columns (for 2 fixed rows)
  var tmp = '';
  for (j = 0; j < data[0].length; j++)
    if (data[0][j] != '')
      tmp = data[0][j];
    else
      data[0][j] = tmp;

  // for 2 fixed rows calculate unique column number
  if (fixRows == 2) {
    uniqueCols = 0;
    tmp = {};
    for (j = fixColumns; j < data[1].length; j++)
      if (typeof tmp[data[1][j]] == 'undefined') {
        tmp[data[1][j]] = 1;
        uniqueCols++;
      }
  }

  // return first row: fix column titles + pivoted values column title + values column title(s)
  row = [];
  for (j = 0; j < fixColumns; j++) row.push(fixRows == 2 ? data[0][j] || data[1][j] : data[0][j]); // for 2 fixed rows we try to find the title in row 1 and row 2
  for (j = 3; j < arguments.length; j++) row.push(arguments[j]);
  ret.push(row);

  // processing rows (skipping the fixed columns, then dedicating a new row for each pivoted value)
  for (i = fixRows; i < data.length && data[i].length > 0; i++) {
    // skip totally empty or only whitespace containing rows
    if (data[i].join('').replace(/\s+/g, '').length == 0) continue;

    // unpivot the row
    row = [];
    for (j = 0; j < fixColumns && j < data[i].length; j++)
      row.push(data[i][j]);
    for (j = fixColumns; j < data[i].length; j += uniqueCols)
      ret.push(
        row.concat([data[0][j]]) // the first row title value
          .concat(data[i].slice(j, j + uniqueCols)) // pivoted values
      );
  }

  return ret;
}


/////////////////////////////////////////////////////////////////////////////////////////////////////////
//CASH DIGITIZATION
/////////////////////////////////////////////////////////////////////////////////////////////////////////

//updates the Master Balances + Balances History + marks the digitized cash as "Digitized" to avoid double digitization
function digitizeCash() {
  var ssCashCreate = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cashCreate);
  const totalRowsCashCreate = getLastRow(cashCreate, 'H:H');
  var notationCashCreate = "M3:AD" + totalRowsCashCreate + 1
  var dataCashCreate = ssCashCreate.getRange(notationCashCreate).getValues()

  //column numbers in MB & Cash Create
  //var customerRowIndexinMBColNum = 0;
  var customerIDColNum = 0;
  var cashAmountColNum = 5;
  var uniqueIDColNum = 6;
  var okayToDigitizeColNum = 10;
  var statusColNum = 11;
  var digitizedColLetter = "X"
  var dateTimeDigitizedColLetter = "Z"
  var customerEmailedColLetter = "AC"

  var operation = "CASH DIGITIZATION"

  //use the following to update cash values in Master Balances
  var ssMB = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(masterBalancesSheet);
  const totalRowsMB = getLastRow(masterBalancesSheet, 'B:B');
  var notationMB = "B1:E" + totalRowsMB + 1

  //retrieving information regarding # of rows in Balances History outside the for loop, adding to totalRows inside the if statement so as to minimize run time 
  //var ssBalancesHistory = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(balancesHistory);
  var totalRowsBalancesHistory = ssBalancesHistory.getLastRow();

  //Customer Onboarding tab details (used to get Customer email)
  var ssCustOnboarding = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(customerOnboarding);
  var notationCO = 'B3:F';
  var dataCustomerOnboarding = ssCustOnboarding.getRange(notationCO).getValues();
  var customerOnboardingTabCustomerTrellisIDColNum = 3;
  var customerOnboardingTabCustomerEmailColNum = 2;
  
  //master balances column with trelli IDs / MPIDs
   var ssMasterBalances = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(masterBalancesSheet);
  var customerDataMasterBalances = ssMasterBalances.getRange("C:C").getValues()

  for (var i = 0; i < dataCashCreate.length; i++) {
    //Chekcs that action ID is unique and that the value has not yet been digitized  
    if (dataCashCreate[i][statusColNum] != "Digitized" && dataCashCreate[i][okayToDigitizeColNum] == "YES") {
      //var customerRow = dataCashCreate[i][customerRowIndexinMBColNum]
      var cash = dataCashCreate[i][cashAmountColNum]
      var uniqueID = dataCashCreate[i][uniqueIDColNum]
      var trellisID = dataCashCreate[i][customerIDColNum]

      var customerRow = getRowNumberOfCustomerInSpreadsheet(customerDataMasterBalances, trellisID)
      //data update intentionally inside the for loop so that we have up to date info from MB every time we need to update balances
      var dataMB = ssMB.getRange(notationMB).getValues()

      updateCustomerCashBalance(ssMB, dataMB, customerRow, cash, uniqueID, trellisID, operation, ssBalancesHistory, totalRowsBalancesHistory)
      totalRowsBalancesHistory += 1


      //after balance is updated, change col M to say Digitized so as to avoid double digitization
      var pointer = i + 3
      ssCashCreate.getRange(digitizedColLetter + pointer).setValue("Digitized")

      var currentTime = new Date();
      ssCashCreate.getRange(dateTimeDigitizedColLetter + pointer).setValue(currentTime)
      
      var recepientEmailAddress = returnEmail(dataCustomerOnboarding, trellisID, customerOnboardingTabCustomerTrellisIDColNum, customerOnboardingTabCustomerEmailColNum)
      sendCashDigitizationEmail(recepientEmailAddress, cash, privateSecuritiesOpsEmail, uniqueID)
      ssCashCreate.getRange(customerEmailedColLetter + pointer).setValue("Sent")
    }
  }
}

function digitizeCashRETIRED() {
  var ssCashCreate = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cashCreate);
  const totalRowsCashCreate = getLastRow(cashCreate, 'H:H');
  var notationCashCreate = "H3:X" + totalRowsCashCreate + 1
  var dataCashCreate = ssCashCreate.getRange(notationCashCreate).getValues()

  //column numbers in MB & Cash Create
  //var customerRowIndexinMBColNum = 0;
  var customerIDColNum = 0;
  var cashAmountColNum = 5;
  var uniqueIDColNum = 6;
  var okayToDigitizeColNum = 10;
  var statusColNum = 11;
  var digitizedColLetter = "S"
  var dateTimeDigitizedColLetter = "U"
  var customerEmailedColLetter = "X"

  var operation = "CASH DIGITIZATION"

  //use the following to update cash values in Master Balances
  var ssMB = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(masterBalancesSheet);
  const totalRowsMB = getLastRow(masterBalancesSheet, 'B:B');
  var notationMB = "B1:E" + totalRowsMB + 1

  //retrieving information regarding # of rows in Balances History outside the for loop, adding to totalRows inside the if statement so as to minimize run time 
  //var ssBalancesHistory = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(balancesHistory);
  var totalRowsBalancesHistory = ssBalancesHistory.getLastRow();

  //Customer Onboarding tab details (used to get Customer email)
  var ssCustOnboarding = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(customerOnboarding);
  var notationCO = 'B3:F';
  var dataCustomerOnboarding = ssCustOnboarding.getRange(notationCO).getValues();
  var customerOnboardingTabCustomerTrellisIDColNum = 3;
  var customerOnboardingTabCustomerEmailColNum = 2;
  
  //master balances column with trelli IDs / MPIDs
   var ssMasterBalances = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(masterBalancesSheet);
  var customerDataMasterBalances = ssMasterBalances.getRange("C:C").getValues()

  for (var i = 0; i < dataCashCreate.length; i++) {
    //Chekcs that action ID is unique and that the value has not yet been digitized  
    if (dataCashCreate[i][statusColNum] != "Digitized" && dataCashCreate[i][okayToDigitizeColNum] == "YES") {
      //var customerRow = dataCashCreate[i][customerRowIndexinMBColNum]
      var cash = dataCashCreate[i][cashAmountColNum]
      var uniqueID = dataCashCreate[i][uniqueIDColNum]
      var trellisID = dataCashCreate[i][customerIDColNum]

      var customerRow = getRowNumberOfCustomerInSpreadsheet(customerDataMasterBalances, trellisID)
      //data update intentionally inside the for loop so that we have up to date info from MB every time we need to update balances
      var dataMB = ssMB.getRange(notationMB).getValues()

      updateCustomerCashBalance(ssMB, dataMB, customerRow, cash, uniqueID, trellisID, operation, ssBalancesHistory, totalRowsBalancesHistory)
      totalRowsBalancesHistory += 1


      //after balance is updated, change col M to say Digitized so as to avoid double digitization
      var pointer = i + 3
      ssCashCreate.getRange(digitizedColLetter + pointer).setValue("Digitized")

      var currentTime = new Date();
      ssCashCreate.getRange(dateTimeDigitizedColLetter + pointer).setValue(currentTime)
      
      var recepientEmailAddress = returnEmail(dataCustomerOnboarding, trellisID, customerOnboardingTabCustomerTrellisIDColNum, customerOnboardingTabCustomerEmailColNum)
      sendCashDigitizationEmail(recepientEmailAddress, cash, privateSecuritiesOpsEmail, uniqueID)
      ssCashCreate.getRange(customerEmailedColLetter + pointer).setValue("Sent")
    }
  }
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//CASH REDEMPTION
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function redeemCash(){

  var customerIDColNum = 2;
  var cashAmountColNum = 4;
  //var customerRowIndexinMBColNum = 12;
  var uniqueIDColNum = 0;
  var redeemedColNum = 10;
  var okToRedeemColNum = 9;
  var redeemedColLetter = "L";
  //var okToRedeemColLetter = "J";
  var preservingUniqueIDColLetter = "M";
  var dateTimeRedeemedColLetter = "N";

  var operation = "CASH REDEMPTION"

  //Cash redemption tab 
  var ssCashRedeem = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cashRedeem);
  const totalRowsCashRedeem = getLastRow(cashRedeem, 'C:C');
  var notationCashRedeem = "B3:V" + totalRowsCashRedeem + 1
  var dataCashRedeem = ssCashRedeem.getRange(notationCashRedeem).getValues() 

  //Balances History info
  var totalRowsBalancesHistory = ssBalancesHistory.getLastRow(); 

  //master balances column with trelli IDs / MPIDs
   var ssMasterBalances = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(masterBalancesSheet);
  var customerDataMasterBalances = ssMasterBalances.getRange("C:C").getValues()

  var notationMB = "B1:E" + totalRowsMB + 1


  for (var i = 0; i < dataCashRedeem.length; i++) {

    //Chekcs that the value has not yet been redeemed  
    if (dataCashRedeem[i][redeemedColNum] != "Redeemed" && dataCashRedeem[i][okToRedeemColNum] == "YES") {
      //var customerRow = dataCashRedeem[i][customerRowIndexinMBColNum]
      var cash = dataCashRedeem[i][cashAmountColNum]
      var uniqueID = dataCashRedeem[i][uniqueIDColNum]
      var trellisID = dataCashRedeem[i][customerIDColNum]

      var customerRow =  getRowNumberOfCustomerInSpreadsheet(customerDataMasterBalances, trellisID)
      dataMB = ssMB.getRange(notationMB).getValues()

      if(checkIfEnoughCash(dataMB, customerRow, cash)){
        Logger.log("passed the test")
        updateCustomerCashBalance(ssMB, dataMB, customerRow, -cash, uniqueID, trellisID, operation, ssBalancesHistory, totalRowsBalancesHistory)
        totalRowsBalancesHistory += 1


        //after balance is updated, change Redeemed col to read "Redeemed" to avoid double redemption

        var pointer = i + 3
        ssCashRedeem.getRange(redeemedColLetter + pointer).setValue("Redeemed")
        //ssCashRedeem.getRange(okToRedeemColLetter + pointer).setValue("Already Redeemed")
        ssCashRedeem.getRange(preservingUniqueIDColLetter + pointer).setValue(dataCashRedeem[i][uniqueIDColNum])

        var currentTime = new Date();
        ssCashRedeem.getRange(dateTimeRedeemedColLetter + pointer).setValue(currentTime)
      }

    }
  }
}

//sends email confirming that customer's cash redemption is complete
function emailCustomerConfirmingCashRedemption() {
  
  //var ssCashRedeem = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cashRedeem);

  var customerIDColNum = 2;
  var cashAmountColNum = 4;
  var uniqueIDColNum = 0;
  var okToSendEmailToCustomerColNum = 19;

  var emailSentToCustomerColLetter = "V"

  //Cash redemption tab 
  var ssCashRedeem = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cashRedeem);
  const totalRowsCashRedeem = getLastRow(cashRedeem, 'C:C');
  var notationCashRedeem = "B3:V" + totalRowsCashRedeem + 1
  var dataCashRedeem = ssCashRedeem.getRange(notationCashRedeem).getValues()  

  for (var i = 0; i < dataCashRedeem.length; i++) {

    //Chekcs that the value has not yet been redeemed  
    if (dataCashRedeem[i][okToSendEmailToCustomerColNum] == "YES") {
      var cash = dataCashRedeem[i][cashAmountColNum]
      var uniqueID = dataCashRedeem[i][uniqueIDColNum]
      var trellisID = dataCashRedeem[i][customerIDColNum]

      var recepientEmailAddress = returnEmail(dataCustomerOnboarding, trellisID, customerOnboardingTabCustomerTrellisIDColNum, customerOnboardingTabCustomerEmailColNum)
      sendCashRedemptionEmail(recepientEmailAddress, cash, privateSecuritiesOpsEmail, uniqueID)

      var pointer = i + 3
      ssCashRedeem.getRange(emailSentToCustomerColLetter + pointer).setValue("Sent")
    }
  }
}


///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//DIGITIZE SECURITIES
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//updates the Master Balances + Balances History + marks the digitized shares as "Digitized" to avoid double digitization
 
function digitizeShares() {
  var operation = "ASSET DIGITIZATION"

  var uniqueIDColNum = 0;
  var okToDigitizeColNum = 22;
  var digitizedColNumber = 23;
  var amountColNum = 6;
  var trellisIDColNum = 3;
  var securityNameColNum = 7;
  var securityIndexinMBColNum = 27;

  //used to mark fields post digitization
  var dateTimeColLetter = "AA";
  var digitizedColLetter = "Y"
  var customerEmailedColLetter = "AB";

  //Master balances tab info
  var totalRowsMB = getLastRow(masterBalancesSheet, 'B:B');
  var notationMB = "B1:Z" + totalRowsMB + 1

  //Balances History tab info
  var totalRowsBalancesHistory = ssBalancesHistory.getLastRow(); 

  //Sec Create tab info 
  var ssSecCreate = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(secCreate);
  var totalRowsSecCreate = getLastRow(secCreate, 'C:C');
  var notationSecCreate = "B2:AC" + totalRowsSecCreate + 1
  var dataSecCreate = ssSecCreate.getRange(notationSecCreate).getValues()



  //master balances column with trelli IDs / MPIDs
  var ssMasterBalances = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(masterBalancesSheet);
  var customerDataMasterBalances = ssMasterBalances.getRange("C:C").getValues()


  //loop through the sec create data 
  for (var i = 0; i < dataSecCreate.length; i++) {
    if (dataSecCreate[i][okToDigitizeColNum] == "YES" && dataSecCreate[i][digitizedColNumber] != "Digitized") {
      var uniqueID = dataSecCreate[i][uniqueIDColNum];
      var trellisID = dataSecCreate[i][trellisIDColNum];
      var customerRow =  getRowNumberOfCustomerInSpreadsheet(customerDataMasterBalances, trellisID)
      //var customerRow = dataSecCreate[i][rowIndexOfCustomerInMBColNum];
      var sharesQuantity = dataSecCreate[i][amountColNum];
      var securityColNumInMB = dataSecCreate[i][securityIndexinMBColNum];
      var securityCUSIP = dataSecCreate[i][securityNameColNum]

      var dataMBSecurities = ssMB.getRange(notationMB).getValues();

      //credit shares to customer account
      updateCustomerSecurityBalance(ssMB, dataMBSecurities, customerRow, sharesQuantity, securityColNumInMB, uniqueID, trellisID, operation, securityCUSIP,ssBalancesHistory,totalRowsBalancesHistory)
      totalRowsBalancesHistory +=1 


      //after balance is updated, "Digitized?" reads Digitized so as to avoid double digitization. Add time of digitization and unique ID at time of digitization        
      var pointer = i + 2

      var currentTime = new Date();
      var valuesArray =[["Digitized",uniqueID, currentTime]]
      ssSecCreate.getRange((digitizedColLetter + pointer)+":"+(dateTimeColLetter + pointer)).setValues(valuesArray)

      var recepientEmailAddress = returnEmail(dataCustomerOnboarding, trellisID, customerOnboardingTabCustomerTrellisIDColNum, customerOnboardingTabCustomerEmailColNum)
      
      sendSecuritiesDigitizationCompleteEmail(recepientEmailAddress, sharesQuantity, securityCUSIP, privateSecuritiesOpsEmail, uniqueID)
      ssSecCreate.getRange(customerEmailedColLetter + pointer).setValue("Sent")

    }

  }
}


////////////////////////////////////////////////////////////////////////////////////
//SECURITIES REDEMPTION
/***
 * Consists of 2 functions 
  1) Redeem from Customer's Account to Holding Account
  2) Redeem from Holding Account and Off Platform
 */
/////////////////////////////////////////////////////////////////////////////////// 
function redeemSharesFromCustomertoHolding() {
  var operation = "ASSET REDEMPTION TO HOLDING"
  const totalRowsMB = getLastRow(masterBalancesSheet, 'B:B');

  //this section of the data is used for updating securities balance in MB
  var notationSecurities = "B1:Z" + totalRowsMB + 1

  //securities redeem tab
  var ssSecRedeem = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(secRedeem);
  const totalRowsSR = getLastRow(secRedeem, 'C:C');
  var notationSecRedeem = "B2:AC" + totalRowsSR + 1
  var dataSecRedeem = ssSecRedeem.getRange(notationSecRedeem).getValues()
  
  //variables in Sec Redeem tab
  var uniqueIDColNum = 0;
  var okayToRedeemFromAcctColNum = 13;
  var trellisIDColNum = 2;
  var amountColNum = 4;
  var securityNameColNum = 5;
  var securityIndexinMBColNum = 8;
  //var rowIndexOfCustomerInMBColNum = 10;
  //var rowIndexOfCustomerHoldingInMBColNum = 11;
  var redeemedColLetter = 'P';
  var redeemedColNum = 14;
  var uniqueIDColLetter = 'AB';
  var dateTimeCustToHoldColLetter = "Q"
  var issuerAgentEmailedColLetter = "R";
  var customerEmailedColLetter = "S"

  //Balance History info
  var totalRowsBalancesHistory = ssBalancesHistory.getLastRow();

  //Securities Onboarding tab details (used to get issuer agent email)
  var ssSecuritiesOnboarding = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(securitiesOnboarding);
  var notationSO = 'J3:K';
  var dataSecuritiesOnboarding = ssSecuritiesOnboarding.getRange(notationSO).getValues();
  //KATE MIGHT NEED TO UPDATE THIS IF WE CHANGE THE SECURITIES ONBOARDING TAB STRUCTURE
  var securitiesOnboardingTabSecurityColNum = 1;
  var issuerAgentEmailColNum = 0;

  //master balances column with trelli IDs or MPIDs
  var ssMasterBalances = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(masterBalancesSheet);
  var customerDataMasterBalances = ssMasterBalances.getRange("C:C").getValues()

  //loop through the data 
    for (var i = 0; i < dataSecRedeem.length; i++) {
    if (dataSecRedeem[i][okayToRedeemFromAcctColNum] == "YES" && dataSecRedeem[i][redeemedColNum] != "Redeemed") {
      var uniqueID = dataSecRedeem[i][uniqueIDColNum];
      var trellisID = dataSecRedeem[i][trellisIDColNum];
      var holdingTrellisID = "Holding_" + trellisID;
      //var customerRow = dataSecRedeem[i][rowIndexOfCustomerInMBColNum];
      var customerRow =  getRowNumberOfCustomerInSpreadsheet(customerDataMasterBalances, trellisID)
      //var holdingCustomerRow = dataSecRedeem[i][rowIndexOfCustomerHoldingInMBColNum];
      var holdingCustomerRow =  getRowNumberOfCustomerInSpreadsheet(customerDataMasterBalances, holdingTrellisID)
      var sharesQuantity = dataSecRedeem[i][amountColNum];
      var securityColNumInMB = dataSecRedeem[i][securityIndexinMBColNum];
      var securityCUSIP = dataSecRedeem[i][securityNameColNum]

      var dataMBSecurities = ssMB.getRange(notationSecurities).getValues();
      //debit shares from customer account

      if(checkIfEnoughAssets(dataMBSecurities, customerRow, securityColNumInMB, sharesQuantity)){
        updateCustomerSecurityBalance(ssMB, dataMBSecurities, customerRow, -sharesQuantity, securityColNumInMB, uniqueID, trellisID, operation, securityCUSIP, ssBalancesHistory,totalRowsBalancesHistory)
        totalRowsBalancesHistory +=1
        //credit shares to holding account
        updateCustomerSecurityBalance(ssMB, dataMBSecurities, holdingCustomerRow, +sharesQuantity, securityColNumInMB, uniqueID, holdingTrellisID, operation, securityCUSIP, ssBalancesHistory,totalRowsBalancesHistory)
        totalRowsBalancesHistory +=1

        var pointer = i + 2

        ssSecRedeem.getRange(redeemedColLetter + pointer).setValue("Redeemed")
        ssSecRedeem.getRange(uniqueIDColLetter + pointer).setValue(uniqueID)

        var currentTime = new Date();
        ssSecRedeem.getRange(dateTimeCustToHoldColLetter + pointer).setValue(currentTime)
        
        var customerEmailAddress = returnEmail(dataCustomerOnboarding, trellisID, customerOnboardingTabCustomerTrellisIDColNum, customerOnboardingTabCustomerEmailColNum)
        var issuerEmailAddress = returnEmail(dataSecuritiesOnboarding, securityCUSIP, securitiesOnboardingTabSecurityColNum, issuerAgentEmailColNum)
    

        sendSecuritiesRedemptionCapTableUpdateEmail(issuerEmailAddress, trellisID, sharesQuantity, securityCUSIP, privateSecuritiesOpsEmail, uniqueID)
        ssSecRedeem.getRange(issuerAgentEmailedColLetter + pointer).setValue("Sent")
        sendSecuritiesRedemptionInProgressEmailToCustomer(customerEmailAddress, sharesQuantity, securityCUSIP, privateSecuritiesOpsEmail, uniqueID)

        ssSecRedeem.getRange(customerEmailedColLetter + pointer).setValue("Sent")
      }
    }
  }
}


function redeemSharesFromHoldingtoOFFPlatform() {
  
  var operation = "ASSET REDEMPTION OFF PLATFORM"
 
  //securities redeem tab
  var ssSecRedeem = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(secRedeem);
  const totalRowsSR = getLastRow(secRedeem, 'C:C');
  var notationSecRedeem = "B2:AC" + totalRowsSR + 1
  var dataSecRedeem = ssSecRedeem.getRange(notationSecRedeem).getValues()
  
  

  //rows in Sec Redeem
  var uniqueIDColNum = 0;
  var okayToRedeemFromHoldingColNum = 22;
  var trellisIDColNum = 2;
  var amountColNum = 4;
  var securityNameColNum = 5;
  var redeemedColLetter = 'Y';
  var redeemedColNum = 23;
  var dateTimeOffPlatformColLetter = "Z"
  var customerEmailedConfirmedColLetter = "AA"

  //column of security in master balances 
  var securityIndexinMBColNum = 8;
  

  //Balance History info
  var totalRowsBalancesHistory = ssBalancesHistory.getLastRow();
  
  //Master Balances info
  const totalRowsMB = getLastRow(masterBalancesSheet, 'B:B');
  var notationMB = "B1:Z" + totalRowsMB + 1
  //master balances column with trelli IDs or MPIDs
  var ssMasterBalances = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(masterBalancesSheet);
  var customerDataMasterBalances = ssMasterBalances.getRange("C:C").getValues()


  //loop through the data 
  for (var i = 0; i < dataSecRedeem.length; i++) {
    if (dataSecRedeem[i][okayToRedeemFromHoldingColNum] == "YES" && dataSecRedeem[i][redeemedColNum] != "Redeemed") {
      Logger.log("insideloop")
      var uniqueID = dataSecRedeem[i][uniqueIDColNum];
      var trellisID = dataSecRedeem[i][trellisIDColNum];
      var holdingTrellisID = "Holding_" + trellisID;
      //var customerRow = dataSecRedeem[i][rowIndexOfCustomerInMBColNum];
      //var holdingCustomerRow = dataSecRedeem[i][rowIndexOfCustomerHoldingInMBColNum];
      var holdingCustomerRow =  getRowNumberOfCustomerInSpreadsheet(customerDataMasterBalances, holdingTrellisID)
      var sharesQuantity = dataSecRedeem[i][amountColNum];
      var securityColNumInMB = dataSecRedeem[i][securityIndexinMBColNum];
      var securityCUSIP = dataSecRedeem[i][securityNameColNum]


      var dataMBSecurities = ssMB.getRange(notationMB).getValues();

      if(checkIfEnoughAssets(dataMBSecurities, holdingCustomerRow, securityColNumInMB, sharesQuantity)){

        //debit shares from holding account
        updateCustomerSecurityBalance(ssMB, dataMBSecurities, holdingCustomerRow, -sharesQuantity, securityColNumInMB, uniqueID, holdingTrellisID, operation, securityCUSIP,ssBalancesHistory,totalRowsBalancesHistory)
        totalRowsBalancesHistory +=1


        var pointer = i + 2
        var currentTime = new Date();
        ssSecRedeem.getRange(redeemedColLetter + pointer).setValue("Redeemed")
        ssSecRedeem.getRange(dateTimeOffPlatformColLetter + pointer).setValue(currentTime)

        //email customers letting them know shares have been redeemed 
        var customerEmailAddress = returnEmail(dataCustomerOnboarding, trellisID, customerOnboardingTabCustomerTrellisIDColNum, customerOnboardingTabCustomerEmailColNum)
        sendSecuritiesRedemptionCompleteEmailToCustomer(customerEmailAddress, sharesQuantity, securityCUSIP, privateSecuritiesOpsEmail, uniqueID)
        ssSecRedeem.getRange(customerEmailedConfirmedColLetter + pointer).setValue("Sent")
      }

    }
  }
}
////////////////////////////////////////////////////////////////////////////////////////////////////
//SENDING MASTER BALANCES TO CLEARLIST
////////////////////////////////////////////////////////////////////////////////////////////////////

function convertMasterBalancesToCSV() {
  //creating a folder for Master Balances inside MM-DD-YYYY folder
  var timeZone = "EST";
  var baufolderid = "1s132fsm3mrJX47MLEBzGdVtcCWKLtjbt";
  var todaydatefolder = Utilities.formatDate(new Date(), timeZone, "MM-dd-yyyy");
  var todaydatefolderid = createFolder(baufolderid, todaydatefolder);
  var MasterBalancesFolderID = createFolder(todaydatefolderid, "Master_Balances_Export");
  var dest_folder = DriveApp.getFolderById(MasterBalancesFolderID);
  var clearlist_outgoing_folder = DriveApp.getFolderById('1Myehii1D3H_sUrvgtuV-yZegp9I1-ru7');

  //converting Master Balances to CSV
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(clearlistBalancesTab);
  var timeZone = "EST";
  //var dateFormatted = Utilities.formatDate(new Date(), timeZone, "MM-dd-yyyy");
  var dateFormatted = Utilities.formatDate(new Date(), timeZone, "yyyyMMdd");

  var d = new Date();
  var currentTime = d.getHours();
  Logger.log(currentTime);
  //var currentTime = d.toLocaleTimeString('en-GB').getHours(); 

  //var fileName = "Clearlist_Positions" + "_"+ dateFormatted + "_" + currentTime + ".csv";
  var fileName = "Clearlist_Positions" + "_" + dateFormatted + "_" + currentTime + ".csv";
  // convert all available sheet data to csv format
  var csvFile = convertRangeToCsvFile_(sheet);
  // create a file in the Docs List with the given name and the csv data
  var file = dest_folder.createFile(fileName, csvFile);
  var file_output2 = clearlist_outgoing_folder.createFile(fileName, csvFile);
  return fileName;
}

//This function sends CSV of MB to a specified email. We are no longer sending emails with files to Clearlist so this function is not in use. Files get put on SFTP using convertMasterBalancesToCSV()
function sendMasterBalancesCSVToClearlist() {
  var fileToSendName = convertMasterBalancesToCSV()
  var emailAddress = 'kchichikashvili@itbit.com'; //NEED TO BE CHANGED TO CLEARLIST'S EMAIL 
  var subjectDate = new Date();
  var dateFormatted = Utilities.formatDate(new Date(), timeZone, "MM-dd-yyyy");
  var subject = "Today's Master Balances " + subjectDate;
  var message = "Hi Team, \n\nPlease find a the latest Master Balances attached. \n\nBest, Paxos"
  //var fileName = "MASTER BALANCES_"+dateFormatted+".csv"
  var file = DriveApp.getFilesByName(fileToSendName);
  if (file.hasNext()) {
    MailApp.sendEmail(emailAddress, subject, message, {
      attachments: [file.next().getAs(MimeType.CSV)],
      name: 'Paxos Settlement'
    })
  }
  Browser.msgBox("Balances Mail Sent to " + emailAddress);
}
//????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????

function convertGTSBalancesToCSV() {
  //creating a folder for GTS Balances inside MM-DD-YYYY folder
  var timeZone = "EST";
  var baufolderid = "1s132fsm3mrJX47MLEBzGdVtcCWKLtjbt";
  var todaydatefolder = Utilities.formatDate(new Date(), timeZone, "MM-dd-yyyy");
  var todaydatefolderid = createFolder(baufolderid, todaydatefolder);
  var GTSBalancesFolderID = createFolder(todaydatefolderid, "GTS_Balances_Export");
  var dest_folder = DriveApp.getFolderById(GTSBalancesFolderID);
  var clearlist_outgoing_folder = DriveApp.getFolderById('1Myehii1D3H_sUrvgtuV-yZegp9I1-ru7');

  //converting GTS balances sheet into csv
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(GTSBalances);

  var dateFormatted = Utilities.formatDate(new Date(), timeZone, "yyyyMMdd");
  var d = new Date();
  var currentTime = d.getHours();
  //var currentTime = d.toLocaleTimeString(); 


  //var fileName = "GTS_Positions" + "_"+ dateFormatted + ".csv";
  var fileName = "GTS_Positions" + "_" + dateFormatted + "_" + currentTime + ".csv";

  // convert all available sheet data to csv format
  var csvFile = convertRangeToCsvFile_(sheet);
  // create a file in the Docs List with the given name and the csv data
  var file = dest_folder.createFile(fileName, csvFile);
  var file_output2 = clearlist_outgoing_folder.createFile(fileName, csvFile);
  return fileName;
}


function convertBalanceHistoryToCSV() {
  //creating a folder for GTS Balances inside MM-DD-YYYY folder
  var timeZone = "EST";
  var BalanceHistoryFolderID = "1-htYkUWyWddG8TFm-24dkDWy47bSw2z1";
  var dest_folder = DriveApp.getFolderById(BalanceHistoryFolderID);
  //var clearlist_outgoing_folder = DriveApp.getFolderById('1Myehii1D3H_sUrvgtuV-yZegp9I1-ru7');

  //converting GTS balances sheet into csv
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(balancesHistory);

  var dateFormatted = Utilities.formatDate(new Date(), timeZone, "yyyyMMdd");
  var d = new Date();
  //var currentTime = d.getHours();
  var currentTime = d.toLocaleTimeString('en-GB');


  //var fileName = "GTS_Positions" + "_"+ dateFormatted + ".csv";
  var fileName = "Balance_History" + "_" + dateFormatted + "_" + currentTime + ".csv";

  // convert all available sheet data to csv format
  var csvFile = convertRangeToCsvFile_(sheet);
  // create a file in the Docs List with the given name and the csv data
  var file = dest_folder.createFile(fileName, csvFile);
  //var file_output2 = clearlist_outgoing_folder.createFile(fileName, csvFile);


  // clear content of balance history after download as csv
  sheet.getRange('A2:H1000').clearContent()

}




//SEND EMAILS TO ISSUER AGENTS TO CONFIRM TRADE DETAILS
//////////////////////////////////////////////////////////////////////////////////////////////////////////



//function sends emails to issuer agents 

//check if email already sent before sending new emails 


function sendEmailsToIssuerAgents() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tradesLedger);
  const totalRows = getLastRow(tradesLedger, 'B:B');
  var notation = 'B3:AP';
  var data = ss.getRange(notation).getValues();

  var transactionTypeColNum = 0;
  var okayTosendEmailToIssuerAgentColNum = 32; //was 26

  var emailSentToIAColNum = 33; //was 27 
  var tradeStatusColNum = 40; //was 32
  var securityTokenColNum = 8; //stays the same
  var tradeIDColNum = 12; //stays the same

  var numberOfEmailsSent = 0;
  //data.length
  for (var i = 0; i < totalRows; i++) {
    if (data[i][okayTosendEmailToIssuerAgentColNum] == 'YES' && data[i][emailSentToIAColNum] != 'YES' && data[i][tradeStatusColNum] != 'SETTLED') {
      var issuerAgentEmail = returnIssuerAgentEmail(data[i][securityTokenColNum]);

      var email = Utilities.formatString('%0s', issuerAgentEmail)


      //contents of the email
      var tradeID = data[i][tradeIDColNum]
      var subjectDate = new Date();
      var dateFormatted = Utilities.formatDate(new Date(), timeZone, "MM-dd-yyyy");
      var subject = "[ACTION REQUIRED] Approval Required for Trade " + tradeID + " " + subjectDate;
      var message = "Hi Team, \n\nPlease confirm that trade " + tradeID + " is ok to settle. \n\nBest, Paxos"

      MailApp.sendEmail(email, subject, message, {
        attachments: [],
        name: 'Paxos Settlement'
      })

      ss.getRange(i + 3, emailSentToIAColNum + 2).setValue("YES");
      ss.getRange(i + 3, transactionTypeColNum + 2).setValue("PENDING");
      numberOfEmailsSent = numberOfEmailsSent + 1;
    }

  }
  Logger.log("number of emails sent" + numberOfEmailsSent);
  if (numberOfEmailsSent > 0) {

    Browser.msgBox(numberOfEmailsSent + " email(s) sent to Issuer Agent(s)");
  }
  else {
    Browser.msgBox("There are no emails to send to Issuer Agents");
  }
}



function sendEmailWithoutAttachment(email, subject, message, BD) {
  MailApp.sendEmail(email, subject, message, {
    cc: BD,
    attachments: [],
    name: 'Paxos Private Securities Custody Operations'
  })
}


//TO BE DELETED OR REWRITTEN TO BE MADE MORE MODULAR (VERIFY THAT IMPORT WORKS BEFORE DELETING)
// move clearlist sent file to archive
// next step check if Uploaded_Clearlist_Trade_Files exist in the date folder if not add it
function ImportCSV_CreateFolder_MoveFileToArchive() {

  // import trade csv file to today trade
  var blankfile = importFromCSV();

  //var baufolderid = "1s132fsm3mrJX47MLEBzGdVtcCWKLtjbt";
  var archivefolderid = "18vPPMxZPTIEeXJ-rruva7KAcVy7aXZ37";
  var todaydatefolder = Utilities.formatDate(new Date(), timeZone, "MM-dd-yyyy");
  var todaydatefolderid = createFolder(archivefolderid, todaydatefolder);
  //var uploadclearlistfolderid = createFolder(todaydatefolderid,"Uploaded_Clearlist_Trade_Files");
  var dest_folder = DriveApp.getFolderById(todaydatefolderid);
  Logger.log(dest_folder);
  var empty_folder = DriveApp.getFolderById("1A8C1TJhtjt-hi_4lMAVJTWi2vMYIuCFu");

  // get file from incoming https://drive.google.com/drive/folders/1sp2QzTccc7wJR6l-CQr-5D2S-MNAb4NU (only starts with CLEAR)
  var mainFolder = DriveApp.getFolderById("1sp2QzTccc7wJR6l-CQr-5D2S-MNAb4NU");
  var f = mainFolder.getFiles();



  while (f.hasNext()) {
    var file = f.next();
    //var regExp = new RegExp("^CLEAR");
    var regExp = new RegExp("^CLEAR.2021" + "[0-9]+[0-9]+[0-9]+[0-9]" + ".csv")

    /*
    if (file.getName().search(regExp) != -1 && blankfile.indexOf(file.getName()) != -1) {
      var fileid = file.getId();
      // move blank to archive folder
      file.moveTo(empty_folder);
      //dest_folder.addFile(file);
      //mainFolder.removeFile(file);
    }
    */

    if (file.getName().search(regExp) != -1) {
      name = file.getName();
      Logger.log(name);
      try {
        file.moveTo(dest_folder);
      } catch (err) {
        Logger.log(err);
        file.moveTo(empty_folder);
      }
    }
  };
}


//this function can send email with attachement, no longer sending emails to clearlist so this function is not being used
function sendTradingHistoryCSVToClearlist() {
  var fileToSendName = convertTradingHistoryToCSV()
  var emailAddress = 'kchichikashvili@itbit.com'; //NEED TO BE CHANGED TO CLEARLIST'S EMAIL 
  var subjectDate = new Date();
  var dateFormatted = Utilities.formatDate(new Date(), timeZone, "MM-dd-yyyy");
  var subject = "Today's Trade History " + subjectDate;
  var message = "Hi Team, \n\nPlease find a list of all trades received today. \n\nBest, Paxos"
  var file = DriveApp.getFilesByName(fileToSendName);
  if (file.hasNext()) {
    MailApp.sendEmail(emailAddress, subject, message, {
      attachments: [file.next().getAs(MimeType.CSV)],
      name: 'Paxos Settlement'
    })
  }
  Browser.msgBox("Trades History Mail Sent to " + emailAddress);
}




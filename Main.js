/** To do

 * dealing with old files - not sustainable to keep a trash can?
 * 
 * Implement filter instead of going through everything all the time
 * 
 * Separate logic to determine import file type from importData function
 * 
 * Error handling in general
 * 
 * Externalize sheet search information
 * 
 * This script was written by Oscar Holm of Numbery, reachable at oscar@numbery.se
 * Disclaimer: 
 * - Please be aware that this solution was put into place as an emergency measure over a few days
 * - It was never meant to be a long-term solution and therefore is not constructed with maintainability in mind
 * - The purpose of each function is vaguely documented throughout, for full chart please see link below
 * 
 * Requirements:
 * In order to run importData successfully it needs to match the following documents from its root folder:
 * - 'SAP' : a document containing the invoices to be matched. Should have SAP in the title.
 * - 'Purchase-Orders' : a document containing the purchase orders. Should have Purchase-Orders in title
 * - 'Translation' : a document containing translation of store names. Should have 'Translation' in the title
 * 
 * MAKING A NEW VERSION
 * - Create a new sheet in an empty folder
 * - Name the first tab 'MAIN'
 * - Mark and copy all of the content of this script
 * - Go into Tools -> Script editor, replace function myFunction() {} by pasting
 * - At the top, change 'Untitled project' into desired title
 * - Press Debug, make sure it's set to onOpen, allow the permissions required
 * - Exit out of script editor tab, update the sheet by reloading the page
 * - In the top menu, 'Match invoice' should have appeared. 
 * - Test it by pressing Match invoice -> Import data, it will return that no files were found.
 * - Put a Purchase-order file, name needs to contain purchase order, in the same folder
 * - Put an invoice file, name needs to contain SAP, in the same folder
 * - Put the Translation file, name needs to contain Translation in the same folder
 * - Now run Import data again, to see if successful, press the 3 lines down by the tabs,
 *   and check so that there are now 6 tabs, 1 visible and 5 hidden.
 * - Everything should now be up and running!
 */

function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Match invoice")
    .addItem("Import data", "importTwo")
    .addItem("Find matches", "main")
    .addToUi();
}

// This class is used to recognize sheets with a certain usage
// Example: the sheet 'portal' is identified by 'PONumber' in row 1, column 1
class sheetInfo {
  constructor(name, identifier, lookUpRow, lookUpColumn) {
    this.name=name;
    this.identifier=identifier;
    this.lookUpRow=lookUpRow;
    this.lookUpColumn=lookUpColumn;
  }
}

class invoice {
  constructor(invoiceNo,reference, amount,date, storeName,vendorName,poNumber) {
    this.invoiceNo=invoiceNo;
    this.reference=reference;
    this.amount=amount;
    this.date=date;
    this.storeName = storeName
    this.vendorName=vendorName;
    this.poNumber=poNumber;
  }
}

//Displays an alert as a Toast message
function displayToastAlert(message) {
  SpreadsheetApp.getActive().toast(message, "⚠️ Alert"); 
}
  
function promptUserForInput(promptText) {
  var ui = SpreadsheetApp.getUi();
  var prompt = ui.prompt(promptText);
  var response = prompt.getResponseText();
  return response;
}


function main() {
  var amountDiff = promptUserForInput("Please enter amount criteria, e.g. '0.1' for matches within 10% of invoice amount")
  if (inputValidator(amountDiff) == false) { return; }
  var dateDiff = promptUserForInput("Please enter date criteria, e.g. 5 for within 5 days of the invoice date")
  if (inputValidator(dateDiff) == false) { return; }
  displayMatches(amountDiff,dateDiff)
}

// Checks whether input is valid number, if not flags and returns false
function inputValidator(input) {
  if (isNaN(input) || input.length < 1 || input < 0 ) { 
    displayToastAlert("No valid number input, process cancelled")    
    return false; }    
  else { 
    return true; }
}

function displayMatches(amountDiff, dateDiff) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var main = ss.getSheetByName("MAIN");

  // Identifying which sheet in document which contains purchase orders
  var listOfSheetInfo =
    [new sheetInfo ("portal","PO_Number",1,1)];
  var sheetLookup = identifySheets(listOfSheetInfo);

  var sheets = ss.getSheets();
  var portal = sheets[sheetLookup[listOfSheetInfo[0].name]];
  var portalData = portal.getDataRange().getValues();


  // Fetch and format invocies
  var listOfInvoices = createListOfInvoices();

  // Formatting print sheet
  main.clearContents();
  var listOfHeaders = [];
  for (attribute in listOfInvoices[0]) {
    listOfHeaders.push(attribute);
  }

  var headers = main.getRange(1,1,1,listOfHeaders.length);
  headers.setValues([listOfHeaders]);
  headers.setFontWeight("bold")


  var array = []
  // Evaluate each invoice and push results into array
  for (var item in listOfInvoices) {
    // Returns all the possible matches separated by a comma
    var result = evaluateInvoice(listOfInvoices[item],amountDiff,dateDiff,portalData).join(" , ");
    var row = [listOfInvoices[item].invoiceNo,
              listOfInvoices[item].reference,
              listOfInvoices[item].amount,
              listOfInvoices[item].date,
              listOfInvoices[item].storeName,
              listOfInvoices[item].vendorName,
               result]
    array.push(row)
    }
  
  main.getRange(main.getLastRow()+1,1,item*1+1,row.length).setValues(array);
  main.autoResizeColumns(1,main.getMaxColumns());
}

// Create suggested PO_Numbers for an invoice in order of likelyhood
// invoice -> list of strings
function evaluateInvoice(invoice, amountDiff, dateDiff, matchingData) {
  var list = [];
  Logger.log(invoice.vendorName)
  // Goes through each PO one by one and attempts to match invoice
  for (var row = 1; row < matchingData.length; row++) {
    
    var status = matchingData[row][8];
    if (status == "canceled") { continue }
    
    // !! Really need to use filter
    var storeName = matchingData[row][3];
    if (storeName === invoice.storeName) {
      // Insert exact vendor match here?
      var vendorName = matchingData[row][6]
      if (vendorName === invoice.vendorName) {
        var value = matchingData[row][1];
        if ((Math.abs(value - invoice.amount)/ value) < amountDiff) {
          var date = matchingData[row][7].valueOf()/24/3600/1000; // convert to days
          var invoiceDate = invoice.date.valueOf()/24/3600/1000; // could change datediff to mili   
          if (Math.abs(invoiceDate - date) < dateDiff) {
            list.push(matchingData[row][0]);
          }               
        }
      }
    }
  }
  return list;
}

function levenshteinRatio(s, t, ratioCalc) {
  var rows = s.length + 1;
  var cols = t.length + 1;
  var distance = zero2D(rows,cols)
  

  for (i = 1; i < rows; i++) {
    for (k = 1; k < cols; k++) {
      distance[i][0] = i;
      distance[0][k] = k;
    }
  }
  
  for (var col = 1; col < cols; col++) {
    for (var row = 1; row < rows; row++) {
      if (s[row-1] == t[col-1]) {
        var cost = 0;
      } else {
        if (ratioCalc == true) { 
          var cost = 2 
          } else {
            var cost = 1 }
      }
      distance[row][col] = Math.min(
        distance[row-1][col] + 1,       // Cost of deletions
        distance[row][col-1] + 1,       // Cost of insertions
        distance[row-1][col-1] + cost   // Cost of substitutions
      )
    }
  }
  if (ratioCalc == true) {
    var ratio = ((s.length+t.length) - distance[row-1][col-1]) / (s.length + t.length)
    return ratio;
  } else {
    return distance[row-1][col-1];
  }

}

// Creates a matrix of zeros with given dimensions
function zero2D(rows, cols) {
  var array = [], row = [];
  while (cols--) row.push(0);
  while (rows--) array.push(row.slice());
  return array;
}

// Collects data from 3 tabs to create a listOfInvoices
function createListOfInvoices() {
  // Tell the function what to look for in each tab
  var listOfSheetInfo = 
    [
    new sheetInfo ("headerDataIndex","DP Document Type",1,2),
    new sheetInfo ("lineItemDataIndex","Document Item Id",1,2),
    new sheetInfo ("storeName","Order ID",1,1),
    new sheetInfo ("vendorName", "Supplier_ID", 2,1)
    ];
  var sheetLookup = identifySheets(listOfSheetInfo);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  // Use key-value pairs given by identifySheets to declare sheets
  var headerSheet = sheets[sheetLookup[listOfSheetInfo[0].name]];
  var lineItemSheet = sheets[sheetLookup[listOfSheetInfo[1].name]];
  var stores = sheets[sheetLookup[listOfSheetInfo[2].name]];
  var vendors = sheets[sheetLookup[listOfSheetInfo[3].name]];

  // Generate list of invoices with complete info
  var lineItems = lineItemSheet.getDataRange().getValues();
  var invoicesArray = headerSheet.getDataRange().getValues();
  var storeArray = stores.getDataRange().getValues();
  var vendorArray = vendors.getDataRange().getValues();
  
  var listOfInvoices = []
  for (i = 1; i < invoicesArray.length; i++) {
    var invoiceNo = invoicesArray[i][0];
    var reference = invoicesArray[i][25];
    var amount = invoicesArray[i][28]
    var docDate = invoicesArray[i][95];
    var vendorId = invoicesArray[i][26];
    
    // Finds vendor name in translation tab to provide ID which can be matched in portal system
    for (j = 1; j < vendorArray.length; j++) {
      if (vendorArray[j][2] == vendorId) {
        var vendorName = vendorArray[j][1];
        vendorFound = true;
        break;
      }
    }
    
    // Brings in the storeId from Line Item tab (that's the only use of this tab)
    for (j = 1; j < lineItems.length; j++) {
      if (lineItems[j][0] == invoiceNo) {
        var storeId = lineItems[j][35];
        break;
      }
    }
    for (j = 1; j < storeArray.length; j++) {
      if (storeArray[j][0] == storeId) {
        var storeName = storeArray[j][2]
        break;
      }
    }
    
    // Create invoice
    var add = new invoice(invoiceNo,reference,amount,docDate,storeName,vendorName);
    listOfInvoices.push(add);
  }

  return listOfInvoices;

}

// Returns list of key value pairs identifying the index of different tabs
// listOfSheetInfo -> key value pairs {sheetName : index}
function identifySheets(sheetInfo) {
  ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var listOfSheetIndex = {};
  // loops through all sheets
  
  for (item in sheetInfo) {
    for (i = 0; i < sheets.length; i++) {
      var checkValue = sheets[i].getRange(sheetInfo[item]["lookUpRow"], sheetInfo[item]["lookUpColumn"]).getValues();
      var sheetName = sheetInfo[item]["name"];
      if (checkValue == sheetInfo[item]["identifier"]) {
        listOfSheetIndex[sheetName] = i;
        break;
      }
    }
    if (listOfSheetIndex[sheetName] == null) {
      Logger.log(sheetInfo[item]["identifier"]+" could not be identified using identifySheets. Error1");
      throw EvalError("Could not identify the sheet: "+sheetName+"")
    }
  }
  return listOfSheetIndex;
}

// Calling necessary files
function importTwo() {
  var parentFolderId = findParentFolder();
  // Deleting any unnecessary sheet
  deleteRedundantSheets(); 
  // Takes all xls files and makes them into separated .csv files (if multiple tabs)
  convertXLS(parentFolderId);
  
  // For import to work, the file names have to contain the keywords in sheetsToImport
  var sheetsToImport = ['Purchase-Orders', 'SAP', 'Translation'];
  for (item in sheetsToImport) {
    var sheet = importData(sheetsToImport[item]);
    if (sheet != true) { break }
  }
}

function findParentFolder() {
  // Finding parent folder ID
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = ss.getId();
  var fileInDrive = DriveApp.getFolderById(ssId);
  var folderinDrive = fileInDrive.getParents().next().getId();
  return folderinDrive;
}

//Import new data from CSV file in Google Drive
function importData(fileName) {
  var files = findFilesInDrive(fileName);
    if (files.length === 0) {
      displayToastAlert("No files with name \"" + fileName + "\" were found in Google Drive.");
      return;
    } else if (files.length > 1) {
      displayToastAlert("Multiple files with name " + fileName +" were found. This program does not support picking the right file yet.")//;
      return;
    }

  var file = files[0];
  var name = file.getName();
  var blob = file.getBlob();
  var type = blob.getContentType();
  
  if (type == "text/csv") {
    var contents = Utilities.parseCsv(blob.getDataAsString(),";");
    // This is a terrible fix for date formatting. Docs can't read Euro dd/mm/yy format
    if (name.match("Purchase-Order") != null) { 
      var contents = fixDateFormat(contents);
     }
    
    var sheetName = writeDataToSheet(contents, fileName);
    fixDateFormat(sheetName);
    displayToastAlert("The CSV was successfully imported into " + sheetName + " data."); 
    return true;
  } 
  else if (type == "application/pdf") {
    var fileId = file.getId();
    var ss = SpreadsheetApp.openById(fileId);
    var sheets = ss.getSheets();
    for (i = 0; i < sheets.length; i++) {
      var fileName = sheets[i].getName();
      var contentsRange = sheets[i].getDataRange();
      var contents = contentsRange.getValues();
      sheetName = writeDataToSheet(contents, fileName);
      displayToastAlert(fileName + " was successfully imported into " + sheetName + " data."); 
    }
    return true;
  }
  else { displayToastAlert("The filetype for " + name + " is not recognized. Please only use .csv, .xlsx or Google Sheet format.")
        return false;
  
  }
}

// Fixes date format when importing the purchase orders csv
function fixDateFormat(content) {
  // Looks in first row of content
  for (i in content[0]) {
    // If it contains Date then convert
    if (content[0][i].split("_").indexOf("Date") != -1 ) {
      for (j = 1; j < content.length; j++) {
        var before = content[j][i].split("/");
        var after = before[1]+"/"+before[0]+"/"+before[2]
        content[j][i] = after;
      }
    }
  }
  return content;
}

// Returns list of csv files with name containing search word
function findFilesInDrive(searchWord) {
  var parentFolderId = findParentFolder();
  var folder = DriveApp.getFolderById(parentFolderId);
  var files = folder.searchFiles("title contains '" + searchWord + "'");
  var result = [];
  while (files.hasNext()) {
    result.push(files.next());
  }
  return result;
}


// Creates new sheet and pastes 2d array into A1
function writeDataToSheet(data, sheetName) {
  var ss = SpreadsheetApp.getActive();
  sheet = ss.insertSheet(sheetName);
  sheet.getRange(1,1,data.length, data[0].length).setValues(data);
  sheet.hideSheet();
  return sheet.getName();
}

// Removes all sheets that aren't the MAIN sheet. Breaks if there's no MAIN sheet.
function deleteRedundantSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (i = 0; i < sheets.length; i++) {
    switch(sheets[i].getSheetName()) {
      case "MAIN":
        break;
      default:
        ss.deleteSheet(sheets[i]);
    }
  }
}

// Convert XLS files into Sheet
function convertXLS(id){
  var folderId = id; 
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.searchFiles('title != "nothing"');

  while(files.hasNext()){
    var xFile = files.next();
    var name = xFile.getName();
    if (name.toLowerCase().indexOf('.xlsx')>-1) {
      var ID = xFile.getId();
      var xBlob = xFile.getBlob();
      var newFile = {
        title : name.slice(0,-5) +'_converted',
        parents: [{id: folderId}] //
      };
      file = Drive.Files.insert(newFile, xBlob, {
        convert: true
      });
      moveFileToTrash(ID,folderId);
      // Drive.Files.remove(ID); // Added // If this line is run, the original XLSX file is removed. So please be careful this.
    }
  }
}

// Moves a file to the designated trash folder
function moveFileToTrash(fileId,parentId) {
  var trashFolderId = "1u3cNnY0a3dV8kRWYR4B7dgUW6GrePcNF"
  var file = DriveApp.getFileById(fileId);

  // Remove the file from all parent folders
  var parent = DriveApp.getFolderById(parentId);
  parent.removeFile(file);
  // Add file to trash
  DriveApp.getFolderById(trashFolderId).addFile(file);
}





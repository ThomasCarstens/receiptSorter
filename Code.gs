function DOUBLE(input) {
  return input * 2;
}

function onOpen() {
   var menu = SpreadsheetApp.getUi().createMenu("⚙️ Admin Settings");
   
   menu.addItem("1. Import Receipt List and Link Picture Folder", "main"); //No Picture folder yet.
   menu.addItem("2. Linker: Receipt-Bank Statement", "linkerA");
   menu.addToUi();
  
}

// 1. Import Receipt List //////////////
///////////////////////////////////////

function main() {
  let fileName = promptUser("Enter the name of the Excel file to import:");
  if(fileName === null) {
    toast("Please enter a valid filename.");
    return;
  }
  let sheetName = promptUser(`Enter the name of the sheet in ${fileName} to import:`);
  if(sheetName === null) {
    toast("Please enter a valid sheet.");
    return;
  }
  toast(`Importing ${sheetName} from ${fileName} ...`);
  let spreadsheetId = convertExcelToGoogleSheets(fileName);
  let importedSheetName = importDataFromSpreadsheet(spreadsheetId, sheetName);
  toast(`Successfully imported data from ${sheetName} in ${fileName} to ${importedSheetName}`);

}

function testListLinks() {
  //Tom for example:
  // var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LABELS");
  // cell = sheet.getRange(11,2,1,1).getValue()
  // listFilesInFolder(cell, "Paris_with_ids");
  //find way to combine CSVs afterwards...
  ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Recus Tom.');
  column = getColumnByName(ss, 'Drive Link')
  column.shift()
  console.log(column)
}


function toast(message) {
  SpreadsheetApp.getActive().toast(message);
}

function promptUser(message) {
  let ui = SpreadsheetApp.getUi();
  let response = ui.prompt(message);
  if(response != null && response.getSelectedButton() === ui.Button.OK) {
    return response.getResponseText();
  } else {
    return null;
  }
}

function convertExcelToGoogleSheets(fileName) {
  let files = DriveApp.getFilesByName(fileName);
  let excelFile = null;
  if(files.hasNext())
    excelFile = files.next();
  else
    return null;
  let blob = excelFile.getBlob();
  let config = {
    title: "[Google Sheets] " + excelFile.getName(),
    parents: [{id: excelFile.getParents().next().getId()}],
    mimeType: MimeType.GOOGLE_SHEETS
  };
  let spreadsheet = Drive.Files.insert(config, blob, { convert: true, supportsAllDrives: true } );
  console.log(spreadsheet.id)
  return spreadsheet.id;
}

function importDataFromSpreadsheet(spreadsheetId, sheetName) {
  let spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  let currentSpreadsheet = SpreadsheetApp.getActive();
  let newSheet = currentSpreadsheet.insertSheet();
  let dataToImport = spreadsheet.getSheetByName(sheetName).getDataRange();
  let range = newSheet.getRange(1,1,dataToImport.getNumRows(), dataToImport.getNumColumns());
  range.setValues(dataToImport.getValues());
  return newSheet.getName();
}

// Linker: Receipt-Bank Statement //////////////
///////////////////////////////////////

function listFilesInFolder(receiptFolder_id, newSheetName) {
  //Automatic Link updater. Every time you Add A New Receipts Folder:
  // (1) Add folder id Here --- manual for now.
  // (2) Make sure Excel Sheet is replaced --- manual for now.

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(newSheetName);
  cell = sheet.getRange(1,10,1,1)
  cell.setValues([["Drive Link"]]);

  var folder = DriveApp.getFolderById(receiptFolder_id);
  var contents = folder.getFiles();
  
  // orderInputPicsToReceipts()
  ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(newSheetName);
  var raw_ids = getColumnByName(ss, 'Receipt ID')
  // console.log('RAW IDS', raw_ids)
  var cnt = 0;
  var file;
  var picture_index = new Array()
  while (contents.hasNext()) {
      var file = contents.next();
      // console.log('index is', file.getName().split('_')[0])
      picture_index.push(raw_ids.flat().indexOf(parseInt(file.getName().split('_')[0])))
      // console.log(picture_index)
      data = [
          "https://drive.google.com/file/d/"+file.getId()+"/view?usp=drivesdk"
      ];
      cnt++ 
      row = picture_index.pop();
      if (row != -1) {
        cell = sheet.getRange(1+row,10,1,1)      
        cell.setValues([data]);
      }
      
  };
  return getColumnByName(ss, 'Drive Link').shift()
};

function getColumnByName(ss, colName) {
  var data = ss.getDataRange().getValues();
  var col = data[0].indexOf(colName);
  if (col != -1) {
    var column = 'ABCDEFGHIJKLM'[col]
  }
  var raw_ids = ss.getRange(column+"1:"+column).getValues().filter(function(el) { return el[0]; });

  return raw_ids
}

// function orderInputPicsToReceipts() {
  
//   ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Paris_with_ids');
//   var raw_ids = getColumnByName(ss, 'Receipt ID')
//   console.log(raw_ids)
//   ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Paris_with_ids');
//   // find how we did pictures.
// }


////////////////////////////////////////////

function linkerA() {
  
  // FNB
  // ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Copy of EasyAccountMar2May');
  // var bank_description_raw = ss.getRange("D8:D").getValues()
  // // var bank_description = bank_description_raw.map((x )=> {if (x[0]!=''){return x[0];}});
  // var bank_description =bank_description_raw.filter(function(el) { return el[0]; });
  // // console.log(bank_description)
  // var bank_dates = ss.getRange("A8:A").getValues().filter(function(el) { return el[0]; });
  // var bank_stamp = bank_dates.map((x) => {return new Date(x[0]).getTime();});

  // Rhone-Alpes
  ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Copy of Feb2022');
  var bank_description_raw = ss.getRange("B14:B").getValues()
  console.log('RAW description:', bank_description_raw.flat())
  var bank_dates_raw = ss.getRange("A14:A").getValues()
  // console.log(bank_dates_raw)
  var bank_description = bank_description_raw.filter(function(el) { if(el[0].includes('CARTE')){return el[0];} })
  console.log('Bank description:', bank_description.flat())
  var bank_dates = new Array();
  for (i=0; i<bank_description.length ; i++){
    relevant_index = bank_description_raw.flat().indexOf(bank_description.flat()[i])
    bank_dates.push(bank_dates_raw[relevant_index])
  }
  console.log(bank_dates)
  
  var bank_stamp = bank_dates.map((x) => {return new Date(x[0]).getTime();});

  // ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Copy of EasyAccountMar2May');
  // var linkcolumn = ss.getRange(1,10,bank_dates.length,1)
  // var hyperlinkcolumn = ss.getRange(1,11,bank_dates.length,1)
  // console.log(bank_stamp)
  ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Copy of Feb2022');
  var linkcolumn = ss.getRange(1,10,bank_dates.length,1)
  var hyperlinkcolumn = ss.getRange(1,11,bank_dates.length,1)

  ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Recus Tom.');


  var receipt_cost = getColumnByName(ss, 'Price')
  receipt_cost.shift()
  // console.log(receipt_cost)
  var receipt_dates = getColumnByName(ss, 'Date')
  receipt_dates.shift()
  var receipt_stamp = receipt_dates.map((x) => {return new Date(x[0]).getTime();});
  // console.log(receipt_stamp)
  var linktext = new Array(bank_stamp.length).fill(['0'])
  var linkpic = new Array(bank_description.length).fill(['No receipt detected.'])

  // !!! 
  //Link Receipt Folder. Tom for example:
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LABELS");
  var cell = sheet.getRange(11,2,1,1).getValue()
  console.log('cell is', cell)


  // var receipt_links = listFilesInFolder(cell, 'Recus Tom.'); //issue with Promise.
  var receipt_links = getColumnByName(ss, 'Drive Link')
  receipt_links.shift()
  //find way to combine CSVs afterwards...
  console.log('end of receipt linking on CSV file.')
  

  var index_duplicate = receiptCleaner(receipt_cost, receipt_links);
  // TO ADD: Return above is modified receipt list.

  for (let i = 0; i < receipt_cost.length; i++) {
    for (let j = 0; j < bank_description.length; j++) {    
        console.log(bank_description[j], '|', receipt_cost[i])
        // if (bank_description[j][0].includes(receipt_cost[i])){ //FNB
        if (bank_description[j][0].includes(receipt_cost[i])){

          console.log('found matching costs...')
          console.log(bank_description[j], '|', receipt_cost[i])

          // Display receipt pictures to the user.
          validation = ConfirmStatement(receipt_links[i], bank_description[j][0])
          // Yes/No executes separately.

          linkpic[j] = [receipt_links[i]] // Currently at same time, but can be put in linker C.
          // linktext[j]=['1']
      } 
    }
  }

  return

  console.log('end of search.')
  // console.log(linktext)
  // linkcolumn.setValues(linktext)
  // console.log('end of labelling.')
  linkerB(linktext,bank_description)
  console.log('end of estimating.')
  // linkerC()
  hyperlinkcolumn.setValues(linkpic)
  console.log('end of receipt linking on Bank Statement.')
}

// PicLinking * Incomplete function!
function linkerC(linktext, bank_description) {
  ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Copy of RECEIPTS PARIS 2022');
  var receipt_links = ss.getRange("L8:L").getValues().filter(function(el) { return el[0]; });

  var linkpic = new Array(bank_description.length).fill(['No receipt'])
  for (let i = 0; i < linktext.length; i++) {  
      if (linktext[i] == '1'){
        linkpic[i] = receipt_links[i]
      } 
    }
  ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Copy of RECEIPTS PARIS 2022');
}

// Estimating share of non-tabulated costs.
function linkerB(linktext, bank_description) {
  // assume linktext exists already.
  unknowns_total = 0
  knowns_total = 0
  var text_to_nb = 0

  for (let i = 0; i < linktext.length; i++) {  
    if (linktext[i] == '0'){
      text_to_nb = parseFloat(bank_description[i][0].split(" ")[0])
      if (text_to_nb){
        // console.log(text_to_nb)
        unknowns_total += text_to_nb        }
    } else {
      text_to_nb = parseFloat(bank_description[i][0].split(" ")[0])
      if (text_to_nb){
        // console.log(text_to_nb)
        knowns_total += text_to_nb        }
    }
  }
  toast('Unknowns:', unknowns_total, 'EUR | ', parseInt(100*unknowns_total/(unknowns_total+knowns_total)), '% of Month.')
}


  // (i) CLEANUP: let's check if any values are equal AND within 5 days of one another.
  //      -prompt user. Also save a value (14 receipts) so the next batch can be added (unclear.)
  // (ii) if anything is in the ZAR currency (prep for CFA...)
  //      -tell me what costs were done as cash (ie. not found in both bank accounts.)

  // (iii) one case for 500 euros bank transfer to be linked to business costs.

function multimatch(str, value) {
  var matches = new Array()
  i = str.indexOf(value);
  while(i >= 0) {
      matches.push(i)
      i = str.indexOf(value, i+1);
  }
  return matches
}


function receiptCleaner(receipt_cost, receipt_links) {
  
  var ui = SpreadsheetApp.getUi(); 
  const set = new Set(receipt_cost);
  // Create a set with duplicate values.
  const duplicates = receipt_cost.filter(item => {
      if (set.has(item)) {
          set.delete(item);
      } else {
        return item;        
      }
  });
  
  for (let i = 0; i < duplicates.length; i++) {  
    // Use duplicates to find indexes of a multimatch
    receipt_matches = multimatch(receipt_cost, duplicates[i])
    console.log('MATCHES:', receipt_matches)
    // Display receipt pictures to the user.
    validation = ValidationAlert(receipt_links[receipt_matches[0]], receipt_links[receipt_matches[1]])
    // Yes/No executes separately.
  }
  return duplicates;
}

function userValidate(){
  toast ('Entries combined.')  
  // Do something.
  return 1
}
function userReject(){
  toast ('Left as is.')  
  // Do something.
  return 0
}

function ConfirmStatement(receipt1, description){
  //Get ids from google links
  a = String(receipt1).slice(32,65)
  var ui = SpreadsheetApp.getUi();
  var html = `
    <p class="body" style="font-family: sans-serif; color:gray; text-align:center">
    We found a bank record. Does it correspond to this receipt?</p>
    <center><img src="https://drive.google.com/uc?export=view&id=${a}" width=300 />
    <p> ${description} </p> </center>
    <body>
    <center><input type="button" value="YES" onClick="google.script.run.withSuccessHandler(function(){google.script.host.close();}).userValidate()"/>
    <input type="button" id="no" value="NO" onClick="google.script.run.withSuccessHandler(function(){google.script.host.close();}).userReject()"/></center>
    </body>
    ` 
  var htmlOutput = HtmlService
      .createHtmlOutput(html)
      .setWidth(700)
      .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Possible receipt-bank association.');

}

function ValidationAlert(receipt1, receipt2, receipt_matches) {
  //Get ids from google links
  a = String(receipt1).slice(32,65)
  b = String(receipt2).slice(32,65)
  var ui = SpreadsheetApp.getUi();
  var html = `
    <p class="body" style="font-family: sans-serif; color:gray; text-align:center">
    We found duplicates. Should we combine these entries?</p>
    <center><img src="https://drive.google.com/uc?export=view&id=${a}" width=300 />
    <img src="https://drive.google.com/uc?export=view&id=${b}" width=300 /></center>
    <body>
    <center><input type="button" value="YES" onClick="google.script.run.withSuccessHandler(function(){google.script.host.close();}).userValidate()"/>
    <input type="button" id="no" value="NO" onClick="google.script.run.withSuccessHandler(function(){google.script.host.close();}).userReject()"/></center>
    </body>
    ` 
  var htmlOutput = HtmlService
      .createHtmlOutput(html)
      .setWidth(700)
      .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Possible duplicates.');

}
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
  listFilesInFolder("1edxQoIQsOpoHqrEDpDlV4l1OOdfG8j96", importedSheetName);

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
  // (3) 
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(newSheetName);
  cell = sheet.getRange(1,10,1,1)
  cell.setValues([["Drive Link"]]);

  var folder = DriveApp.getFolderById(receiptFolder_id);
  var contents = folder.getFiles();

  var cnt = 0;
  var file;

  while (contents.hasNext()) {
      var file = contents.next();
      cnt++;
      cell = sheet.getRange(1+cnt,10,1,1)

      data = [
          "https://drive.google.com/file/d/"+file.getId()+"/view?usp=drivesdk"
      ];
      console.log(data)
      
      cell.setValues([data]);
  };
};

function setValueatLocation(val, row, col) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet2');
  var range = sheet.getRange(row,col);
  range.setValue(val);
  
}


function linkerA() {
  
  ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Copy of EasyAccountMar2May');
  var bank_description_raw = ss.getRange("D8:D").getValues()
  // var bank_description = bank_description_raw.map((x )=> {if (x[0]!=''){return x[0];}});
  var bank_description =bank_description_raw.filter(function(el) { return el[0]; });
  // console.log(bank_description)
  var bank_dates = ss.getRange("A8:A").getValues().filter(function(el) { return el[0]; });
  var bank_stamp = bank_dates.map((x) => {return new Date(x[0]).getTime();});


  ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet2');
  var linkcolumn = ss.getRange(8,7,bank_dates.length,1)
  var hyperlinkcolumn = ss.getRange(8,8,bank_dates.length,1)
  // console.log(bank_stamp)


  ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Copy of RECEIPTS PARIS 2022');
  var receipt_description = ss.getRange("C2:D").getValues()
  var receipt_cost = receipt_description.map(x => x[0]);
  // console.log(receipt_cost)
  var receipt_dates = ss.getRange("A2:A").getValues().filter(function(el) { return el[0]; });
  var receipt_stamp = receipt_dates.map((x) => {return new Date(x[0]).getTime();});
  // console.log(receipt_stamp)
  
  var linktext = new Array(bank_stamp.length).fill(['0'])
  var linkpic = new Array(bank_description.length).fill(['No receipt'])
  var receipt_links = ss.getRange("L2:L").getValues().filter(function(el) { return el[0]; });

  var anon_duplicate = receiptCleaner(receipt_cost);

  for (let i = 0; i < receipt_stamp.length; i++) {
    for (let j = 0; j < bank_stamp.length; j++) {    
      // console.log(receipt_stamp[i])
        if (bank_description[j][0].includes(receipt_cost[i])){
          console.log('found matching costs...')
          console.log(bank_description[j], '|', receipt_cost[i])
          linkpic[j] = [receipt_links[i]] // Currently at same time, but can be put in linker C.


          linktext[j]=['1']
      } 
    }
  }

  console.log('end of search.')
  // console.log(linktext)
  linkcolumn.setValues(linktext)
  console.log('end of labelling.')
  linkerB(linktext,bank_description)
  console.log('end of estimating.')
  // linkerC()
  hyperlinkcolumn.setValues(linkpic)
  console.log('end of receipt linking.')
}

// PicLinking.
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


function linkerB(linktext, bank_description) {
  // assume linktext exists already.
  unknowns_total = 0
  knowns_total = 0

  // console.log(bank_description)
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

function receiptCleaner(receipt_cost) {
  const set = new Set(receipt_cost);
  const duplicates = receipt_cost.filter(item => {
      if (set.has(item)) {
          set.delete(item);
      } else {
        return item;        
      }
  });

  console.log(duplicates)
  var ui = SpreadsheetApp.getUi(); // Same variations.


  for (let i = 0; i < duplicates.length; i++) {  
    // runTwo()
    validation = ValidationAlert('We found duplicates.', 'should we combine these entries?')
    if (validation){
      ui.alert('Index is '+ String(receipt_cost.indexOf(duplicates[i])));
    }
  }

  return duplicates;

}

function userValidate(){
  toast ('Entries combined.')  
  // Do something.
  return 1
}
function userReject(){
  // var output = HtmlService.createHtmlOutput('<script>google.script.host.close();</script>');
  // SpreadsheetApp.getUi().showModalDialog(output, 'Loading...');
  toast ('Left as is.')  
  // Do something.
  return 0
}

function runTwo(){
  var ui = SpreadsheetApp.getUi();
  html='<input type="button" value="ok" onClick="google.script.run.withSuccessHandler(function(){google.script.host.close();}).clickOk()" />';
  html+='<input type="button" value="cancel" onClick="google.script.run.withSuccessHandler(function(){google.script.host.close();}).clickCancel()"/>';
  ui.showModalDialog(HtmlService.createHtmlOutput(html), 'sample');
}

function clickCancel() {
  SpreadsheetApp.getUi().alert('Clicked cancel');
}

function clickOk() {
 SpreadsheetApp.getUi().alert('Click Ok')
}

function ValidationAlert(title, question) {
  var ui = SpreadsheetApp.getUi();
  var output = HtmlService.createHtmlOutput('<script>google.script.host.close();</script>');
  var response = 2
  var html = `
    
    <p class="body" style="font-family: sans-serif; color:gray; text-align:center">
    We found duplicates. Should we combine these entries?</p>
    <center><img src="https://drive.google.com/uc?export=view&id=1-Efr-CNn6TysNOdIGptyPbF_jSmLmGjC" width=300 />
    <img src="https://drive.google.com/uc?export=view&id=190LU9e-JjolvFuev3unfOuUBJMw8OKpA" width=300 /></center>
    <body>
    <center><input type="button" value="YES" onClick="google.script.run.withSuccessHandler(function(){google.script.host.close();}).userValidate()"/>
    <input type="button" id="no" value="NO" onClick="var response=0;google.script.host.close();console.log(response)"/></center>
    </body>
    ` 

  var htmlOutput = HtmlService
      .createHtmlOutput(html)
      .setWidth(700)
      .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Duplicates');

}



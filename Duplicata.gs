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

function receiptCleaner() {
  ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Recus Tom.');
  var receipt_links = getColumnByName(ss, 'Drive Link')
  var receipt_cost = getColumnByName(ss, 'Price').flat()
  receipt_cost.shift()
  receipt_links.shift()

  const set = new Set(receipt_cost);
  // Create a set with duplicate values.
  const duplicates = receipt_cost.filter(item => {
      if (set.has(item)) {
          set.delete(item);
      } else {
        return item;        
      }
  });
  match_list = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Process Duplicates');
  cell = match_list.getRange(1,11,1,duplicates.length)      
  cell.setValues([duplicates]);
  // store duplicates for processing.
  theNext()
  // Yes/No executes separately.
  return duplicates;
}

function userValidate(){
  toast ('Entries combined.')  
  // Do something.
  cell = nextCell()
  if (cell.toString().length != 0) {
    theNext()
  } else {
    toast( 'No more duplicates for today.')
  }
  
  return 1
}

function userReject(){
  toast ('Left as is.')  
  // Do something.
  cell = nextCell()
  if (cell.toString().length != 0) {
    theNext()
  } else {
    toast( 'No more duplicates for today.')
  }
  return 0
}

function nextCell(){
  match_list = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Process Duplicates');
  cell = match_list.getRange(1,11,1,1).getValue()
  match_list.deleteRow(1);
  return cell
}
function theNext(){
  ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Recus Tom.');
  var receipt_links = getColumnByName(ss, 'Drive Link')
  receipt_links.shift()
  var receipt_cost = getColumnByName(ss, 'Price').flat()
  receipt_cost.shift()
  duplicate = nextCell()
  receipt_matches = multimatch(receipt_cost, duplicate)
  console.log('MATCHES:', receipt_matches)
  validation = ValidationAlert(receipt_links[receipt_matches[0]], receipt_links[receipt_matches[1]])
}


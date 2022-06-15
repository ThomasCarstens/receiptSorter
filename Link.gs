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
    <center><input type="button" value="YES" onClick="google.script.run.withSuccessHandler(function(){google.script.host.close();}).linkValidate()"/>
    <input type="button" id="no" value="NO" onClick="google.script.run.withSuccessHandler(function(){google.script.host.close();}).linkReject()"/></center>
    </body>
    ` 
  var htmlOutput = HtmlService
      .createHtmlOutput(html)
      .setWidth(700)
      .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Possible receipt-bank association.');

}

function linkValidate(){
  toast ('Entries combined.')  
  // Do something.
  cell = linkContents()

  if (cell.toString().length != 0) {
    nextLink()
  } else {
    toast( 'No more links for today.')
  }
  
  return 1
}

function linkReject(){
  toast ('Left as is.')  
  // Do something.
  cell = linkContents()
  if (cell.toString().length != 0) {
    nextLink()
  } else {
    toast( 'No more links for today.')
  }
  return 0
}

function linkContents(){
  match_list = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Process Duplicates');
  cell = match_list.getRange(2,4,1,1).getValue()
  match_list.deleteColumn(4);
  return cell
}

function nextLink() {
  ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Recus Tom.');
  var receipt_links = getColumnByName(ss, 'Drive Link')
  receipt_links.shift()
  var receipt_cost = getColumnByName(ss, 'Price').flat()
  receipt_cost.shift()


  link = linkContents()
  links=link.split('|')
  // Display receipt pictures to the user.
  validation = ConfirmStatement(links[1], links[0])

}

function linkerFunction() {
  ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Copy of Feb2022');

  var bank_description_raw = ss.getRange("B14:B").getValues()
  var bank_description = bank_description_raw.filter(function(el) { if(el[0].includes('CARTE')){return el[0];} })

  var bank_dates_raw = ss.getRange("A14:A").getValues()
  var bank_dates = new Array();

  for (i=0; i<bank_description.length ; i++){
    relevant_index = bank_description_raw.flat().indexOf(bank_description.flat()[i])
    bank_dates.push(bank_dates_raw[relevant_index])
  }
  var bank_stamp = bank_dates.map((x) => {return new Date(x[0]).getTime();});


  ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Recus Tom.');
  var receipt_links = getColumnByName(ss, 'Drive Link')
  var receipt_cost = getColumnByName(ss, 'Price').flat()
  receipt_cost.shift()
  receipt_links.shift()
  var linkpic = new Array()


  for (let i = 0; i < receipt_cost.length; i++) {
    for (let j = 0; j < bank_description.length; j++) {    
        console.log(bank_description[j], '|', receipt_cost[i][0])
        // if (bank_description[j][0].includes(receipt_cost[i])){ //FNB
        if (bank_description[j][0].includes(receipt_cost[i])){

          console.log('found matching costs...')
          console.log(bank_description[j], '|', receipt_cost[i])

          // if Valid:
          linkpic.push(bank_description[j] + '|' + (receipt_links[i]))
      } 
    }
  }
  match_list = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Process Duplicates');
  cell = match_list.getRange(2,4,1,linkpic.length)      
  cell.setValues([linkpic]);
  // store duplicates for processing.
  nextLink()
  // Yes/No executes separately.
}

function listInFolder(folderName) {

   var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet2');
   sheet.appendRow(["Name", "File-Id"]);

// https://drive.google.com/drive/folders/1edxQoIQsOpoHqrEDpDlV4l1OOdfG8j96?usp=sharing
//change the folder ID below to reflect your folder's ID (look in the URL when you're in your folder)
    var folder = DriveApp.getFolderById("1edxQoIQsOpoHqrEDpDlV4l1OOdfG8j96");
    var contents = folder.getFiles();

    var cnt = 0;
    var file;

    while (contents.hasNext()) {
        var file = contents.next();
        cnt++;

           data = [
                file.getName(),
                file.getId(),
            ];

            sheet.appendRow(data);
    };
};

function listAtLocation() {
  console.log('hi')
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet2');
  cell = sheet.getRange(1,1,1,3)
  // var range = sheet.getRange(row,col);
  cell.setValues([["Name", "File-Id", "Link"]]);

  // https://drive.google.com/drive/folders/1edxQoIQsOpoHqrEDpDlV4l1OOdfG8j96?usp=sharing
  //change the folder ID below to reflect your folder's ID (look in the URL when you're in your folder)
  var folder = DriveApp.getFolderById("1edxQoIQsOpoHqrEDpDlV4l1OOdfG8j96");
  var contents = folder.getFiles();

  var cnt = 0;
  var file;

  while (contents.hasNext()) {
      var file = contents.next();
      cnt++;
      cell = sheet.getRange(1+cnt,1,1,3)

      data = [
          file.getName(),
          file.getId(),
          "https://drive.google.com/file/d/"+file.getId()+"/view?usp=drivesdk"
          // setFormula('=HYPERLINK("https://drive.google.com/file/d/"+file.getId()+"/view?usp=drivesdk", file.getName())')
      ];
      console.log(data)
      
      cell.setValues([data]);
  };
};

function setValueInSpreadsheet(val, row, col) {
  // Insert spreadsheet URL
  var SPREADSHEET_URL = 'INSERT_SPREADSHEET_URL';
  // Name of the specific sheet in the spreadsheet.
  var SHEET_NAME = 'SheetName';

  var ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  var sheet = ss.getSheetByName(SHEET_NAME);

  var range = sheet.getRange(row,col);
  var value = range.getValue();
  
  range.setValue(val);
  
}


// function onOpen() {
//   SpreadsheetApp.getUi().createMenu('⚙ Scripts')
//   .addItem('Get column', 'get_column')
//   .addToUi();
// }

function get_column() {
  var s_src = SpreadsheetApp.openByUrl("your_url");
  var col = s_src.getRange("A1:A").getValues().flat();
  var row = SpreadsheetApp.getActiveSheet().getRange(1,1,1,col.length);
  row.setValues([col]);
}


  // https://docs.google.com/spreadsheets/d/1YAVW3-MFRIS3VtbBMBNnlMGs2SArrOzBuxwoK65p3Ng/edit#gid=1660428332
function CSV_LINKER(input) {
    // SpreadsheetApp.getActive().toast("You selected", col);
  const csvFileId = "1XCHmZ4_o0pWivUMcKJbySxkgy6j8n_M5";
  // https://drive.google.com/file/d/1XCHmZ4_o0pWivUMcKJbySxkgy6j8n_M5/view?usp=sharing
  const csv = DriveApp.getFileById(csvFileId);
  const fileId = "1YAVW3-MFRIS3VtbBMBNnlMGs2SArrOzBuxwoK65p3Ng";
  var ss = SpreadsheetApp.openById(fileId);
  var sh = ss.getSheetByName('Copy of EasyAccountMar2May');
  var rng = sh.getDataRange().getValues();
  console.log(sh)
  var msg = '';

  // const csvBlob = csv.getBlob();
  // console.log("BLOB IS", csvBlob)
  // const csvData = Utilities.parseCsv(csvBlob.getDataAsString());
  // csvData.shift(); // Remove headers
  // console.log("DATA IS", csvData)
  // for (let i = 0; i < csvData.length; i++) {
  //   return csvData[i];
  //   if (csvData[i][0] == "John") {
  //     // All rows with John name will be imported to John's googlesheet
  //   }
  // }

  // DriveApp.createFile('New Text File', 'Hello, world!');

} 
// =IF(SUMPRODUCT(-- ISNUMBER(SEARCH(LABELS!SHOPS,B19))),D19,"00")

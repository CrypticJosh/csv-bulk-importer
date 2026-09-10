function onOpen() {
SpreadsheetApp.getUi()
.createAddonMenu()
.addItem('Bulk Import CSVs', 'showCsvUploader')
.addToUi();
}

/**

Runs when the add-on is installed.
*/
function onInstall(e) {
onOpen(e);
}

/**

Opens the CSV importer.
*/
function showCsvUploader() {

const html = HtmlService
.createHtmlOutputFromFile('CsvUploader')
.setWidth(600)
.setHeight(600);

SpreadsheetApp
.getUi()
.showModalDialog(
html,
'Bulk Import CSVs'
);
}

/**

Gets all tabs from the currently
active Google Spreadsheet.
*/
function getSheetNames() {

const spreadsheet =
SpreadsheetApp.getActiveSpreadsheet();

return spreadsheet
.getSheets()
.map(function(sheet) {
return sheet.getName();
});
}

/**

Imports CSV files into the selected
destination tab.
*/
function processCsvFiles(
files,
destinationSheetName
) {

if (!files || files.length === 0) {
throw new Error(
'No CSV files were selected.'
);
}

if (!destinationSheetName) {
throw new Error(
'Please select a destination tab.'
);
}

const spreadsheet =
SpreadsheetApp.getActiveSpreadsheet();

const sheet =
spreadsheet.getSheetByName(
destinationSheetName
);

if (!sheet) {
throw new Error(
'The destination tab "' +
destinationSheetName +
'" could not be found.'
);
}

let allRows = [];
let headerAdded = false;
let validFiles = 0;

files.forEach(function(file) {

if (
  !file.name ||
  !file.name
    .toLowerCase()
    .endsWith('.csv')
) {
  return;
}


const rows =
  Utilities.parseCsv(file.data);


if (
  !rows ||
  rows.length === 0
) {
  return;
}


validFiles++;


/*
 * Keep the header from the
 * first CSV only.
 */
if (!headerAdded) {

  allRows =
    allRows.concat(rows);

  headerAdded = true;

} else {

  /*
   * Skip the header row
   * from subsequent CSVs.
   */
  allRows =
    allRows.concat(
      rows.slice(1)
    );

}


});

if (allRows.length === 0) {
throw new Error(
'No valid CSV data was found.'
);
}

/*

Find the next empty row.
*/
const lastRow =
sheet.getLastRow();

const startRow =
lastRow === 0
? 1
: lastRow + 1;

/*

Find maximum column count.
*/
const maxColumns =
Math.max(
...allRows.map(
function(row) {
return row.length;
}
)
);

/*

Make every row the same

number of columns.
*/
allRows =
allRows.map(function(row) {

const newRow =
row.slice();

while (
newRow.length <
maxColumns
) {
newRow.push('');
}

return newRow;

});


/*

Make sure the sheet has
enough columns.
*/
if (
sheet.getMaxColumns() <
maxColumns
) {
sheet.insertColumnsAfter(
  sheet.getMaxColumns(),
  maxColumns -
  sheet.getMaxColumns()
);


}

/*

Write all CSV data.
*/
sheet
.getRange(
startRow,
1,
allRows.length,
maxColumns
)
.setValues(allRows);

return {

filesImported:
  validFiles,

rowsImported:
  allRows.length,

sheetName:
  destinationSheetName


};

}
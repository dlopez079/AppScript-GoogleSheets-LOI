function onOpen() {
  SpreadsheetApp.getUi().createMenu("Mailer")
    .addItem("Import Excel file from Drive", "main")
    .addSeparator()
    .addItem('Get Offer Price', 'calculateOfferPrice')
    .addSeparator()
    .addItem('Get EMD', 'calculateEMD')
    .addSeparator()
    .addItem('Mail Merge', 'createNewGoogleDocs')
    .addToUi();
}

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
  let spreadsheet = Drive.Files.insert(config, blob);
  return spreadsheet.id;
}

function importDataFromSpreadsheet(spreadsheetId, sheetName) {
  let spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  let currentSpreadsheet = SpreadsheetApp.getActive();
  let date = Date()
  let newSheet = currentSpreadsheet.insertSheet(date);
  let dataToImport = spreadsheet.getSheetByName(sheetName).getDataRange();
  let range = newSheet.getRange(1,1,dataToImport.getNumRows(), dataToImport.getNumColumns());
  range.setValues(dataToImport.getValues());
  return newSheet.getName();
}



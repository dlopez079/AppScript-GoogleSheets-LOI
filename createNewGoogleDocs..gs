function createNewGoogleDocs() {

  // Using Google Sheets (SpreadsheetApp)
  // Get the active spreadsheet (getActiveSpreadsheet)
  // Get the active spreadsheet tab (getActiveSheet)
  // Save the shee to the container sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Inside of this sheet, obtain the last row and save it within "lastRow"
  var lastRow = sheet.getLastRow();

  // Open the Google Docs Template and obtain the template ID from the URL.
  // Save the URL within a container called 'templateDocID'
  var templateDocId = '1RbEp6qSDlVz_Iz1MN60KL7XdGJdFaKaJgTuYcavLDIM'; // Replace with your template's document ID

  // Document Link Column Title
  var docLinkColTitle = "Document Link";
  var docLinkColCell = sheet.getRange("K1");
  docLinkColCell.setValue(docLinkColTitle)


  // Row 1 is the header so go through each row starting from row 2
  for (var i = 2; i <= lastRow; i++) { 

    // Obtain the values from each rows and define a container called 'row'.
    var row = sheet.getRange(i, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Create a document name for each of the letters.
    // Below will save the document by first name and last name.
    var docName = row[0] + ' ' + row[1]; 

    // Obtain the template using the template doc id and save is to container called 'doc'
    var doc = DriveApp.getFileById(templateDocId).makeCopy(docName);

    // Obtain the body of the template and save it in a container called 'docBody'
    var docBody = DocumentApp.openById(doc.getId()).getBody();

    // Replace placeholders with actual data within docBody
    docBody.replaceText('{{First Name}}', row[0]);
    docBody.replaceText('{{Last Name}}', row[1]);
    docBody.replaceText('{{Address}}', row[2]);
    docBody.replaceText('{{City}}', row[3]);
    docBody.replaceText('{{State}}', row[4]);
    docBody.replaceText('{{Zip}}', row[5]);
    docBody.replaceText('{{APN}}', row[6]);
    docBody.replaceText('{{Asking Price}}', row[7]);
    docBody.replaceText('{{Offer Price}}', row[8]);
    docBody.replaceText('{{EMD}}', row[9]);
    // Add more replacements as needed
    // row[0,1,2,3,4,5,6,7,8,9] are equivalent to the column headers on the page. 
    // So row[0] = Column 1 = First Name; row[1] = Column 2 = Last Name; row[2] = Column 3 = Address; etc.

    // What is last column? Because it is not defined.

    // Save the document URL back to the spreadsheet
    sheet.getRange("K" + i).setValue(doc.getUrl());
  }
}

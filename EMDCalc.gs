// Calculate the offer price.
function calculateEMD() {
  // Assuming your data is in columns G (Asking Price) and H (Offer Price)
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();

  // EMD Column Title
  var emdColTitle = "EMD";
  var emdColCell = sheet.getRange("J1");
  emdColCell.setValue(emdColTitle)

  // Loop through each row
  for (var i = 2; i <= lastRow; i++) { // Assuming row 1 is the header

    // Set EMD Cell
    var emdCell = sheet.getRange("J" + i);
    var offerPriceCell = sheet.getRange("I" + i);
    var offerPrice = offerPriceCell.getValue();

    // Calculate the emd price based on true statement of offer
    var emd
    if (offerPrice) {
      emd = offerPrice * 0.02; // 2%
    } else {
      emd = "Get Offer Price"; // Default "Get Offer Price"
    }
    
    // Set the calculated offer price in the Offer Price column (H)
    emdCell.setValue(emd);
  }

}

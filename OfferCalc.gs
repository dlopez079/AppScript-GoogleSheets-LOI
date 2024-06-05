// Calculate the offer price.
function calculateOfferPrice() {
  // Assuming your data is in columns G (Asking Price) and H (Offer Price)
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();

  // Offer Price Column Title
  var offerPriceColTitle = "Offer Price";
  var offerPriceColCell = sheet.getRange("I1");
  offerPriceColCell.setValue(offerPriceColTitle)

  // Loop through each row
  for (var i = 2; i <= lastRow; i++) { // Assuming row 1 is the header
    var askingPriceCell = sheet.getRange("H" + i);
    var offerPriceCell = sheet.getRange("I" + i);
    var askingPrice = askingPriceCell.getValue();
    
    // Calculate the offer price based on percentage ranges
    var offerPrice;
    if (askingPrice < 150000) {
      offerPrice = askingPrice * 0.6; // 60%
    } else if (askingPrice >= 150000 && askingPrice <= 200000) {
      offerPrice = askingPrice * 0.65; // 65%
    } else if (askingPrice > 200000 && askingPrice <= 300000) {
      offerPrice = askingPrice * 0.7; // 70%
    } else if (askingPrice > 300000 && askingPrice <= 400000) {
      offerPrice = askingPrice * 0.75; // 75%
    } else if (askingPrice > 400000 && askingPrice <= 500000) {
      offerPrice = askingPrice * 0.8; // 80%
    } else {
      offerPrice = askingPrice * 0.8; // Default to 80% for >$500k
    }
    
    // Set the calculated offer price in the Offer Price column (H)
    offerPriceCell.setValue(offerPrice);
  }

}

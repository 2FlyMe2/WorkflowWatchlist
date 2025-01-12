// Function to update ticker data
function updateTicker(sheet, row) {
  const dateColumn = 6; // Column F (date column)
  const priceColumn = 7; // Column G (price column)
  const livePriceColumn = 8; // Column H (live price column)

  const livePrice = sheet.getRange(row, livePriceColumn).getValue();

  // Log the action
  Logger.log(`Updating ticker at row ${row}: Date and Price`);

  // Set the current date and live price in the respective columns
  sheet.getRange(row, dateColumn).setValue(new Date()); // Set current date
  sheet.getRange(row, priceColumn).setValue(livePrice); // Set live price

  Logger.log(`Set Date: ${new Date()}, Price: ${livePrice}`);
}

// Function triggered when the user edits a cell
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const row = range.getRow();
  const column = range.getColumn();

  // Log the action
  Logger.log(`onEdit triggered at row ${row}, column ${column}`);

  // Check if "c" is typed in column C
  if (column === 3 && range.getValue().toLowerCase() === "c") {
    Logger.log(`'c' detected in column C at row ${row}`);
    updateTicker(sheet, row); // Call the function to update the ticker data
    range.setValue(""); // Optionally clear the "c" text after processing
    Logger.log(`Cleared 'c' from column C at row ${row}`);
  }
}

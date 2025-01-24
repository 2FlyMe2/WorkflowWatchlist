function processEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Watch');  // "Watch" tab
  if (!sheet) {
    Logger.log("Sheet 'Watch' not found.");
    return;
  }
  
  const labelName = 'Watchlist'; // Label in Gmail
  const label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    Logger.log(`Label '${labelName}' not found. Create it in Gmail.`);
    return;
  }

  const archiveLabelName = 'Watchlist-Archive'; // Label for processed emails
  let archiveLabel = GmailApp.getUserLabelByName(archiveLabelName);
  if (!archiveLabel) {
    archiveLabel = GmailApp.createLabel(archiveLabelName);
  }

  const threads = label.getThreads(0, 50); // Fetch up to 50 labeled threads
  const today = new Date();

  threads.forEach(thread => {
    const messages = thread.getMessages();
    let tickerFound = false;

    messages.forEach(message => {
      let subject = message.getSubject();
      let body = message.getPlainBody().replace(/\n/g, " "); // Remove line breaks

      // Extract URL from the message body
      let url = extractUrlFromBody(body);

      let ticker = extractTickerFromSubject(subject) || extractTickerFromBody(body);

      if (ticker) {
        tickerFound = true;
        const firstEmptyRow = sheet.getLastRow() + 1;

        // Remove the `$` symbol from the ticker
        ticker = ticker.replace('$', '');

        const tickers = sheet.getRange('B2:B' + sheet.getLastRow()).getValues().flat();
        if (tickers.includes(ticker)) {
          Logger.log(`Ticker ${ticker} already exists. Archiving.`);
          thread.removeLabel(label);
          thread.addLabel(archiveLabel);
          return;
        }

        // Copy formatting from the template row (row 2)
        sheet.getRange('2:2').copyTo(sheet.getRange(firstEmptyRow + ':' + firstEmptyRow), { contentsOnly: false });

        // Add data to the first empty row
        sheet.getRange(firstEmptyRow, 1).setValue(''); // Column A: Marker (blank)
        sheet.getRange(firstEmptyRow, 2).setValue(ticker);  // Column B: Ticker
        sheet.getRange(firstEmptyRow, 4).setValue(url || ""); // Column D: URL
        sheet.getRange(firstEmptyRow, 5).setValue(''); // Column E: Entry (blank)
        sheet.getRange(firstEmptyRow, 6).setValue(today);   // Column F: Date
        sheet.getRange(firstEmptyRow, 10).setValue(''); // Column J: Support (blank)
        sheet.getRange(firstEmptyRow, 11).setValue(''); // Column K: Limit (blank)
        sheet.getRange(firstEmptyRow, 12).setValue(''); // Column L: Prediction (blank)
        sheet.getRange(firstEmptyRow, 15).setValue(body);   // Column T: Source
        sheet.getRange(firstEmptyRow, 14).setValue(subject); // Column S: Note
        sheet.getRange(firstEmptyRow, 16).setFormula(`=HYPERLINK("https://stocktwits.com/symbol/${ticker}", "StockTwits")`); // Column P: StockTwits link

        message.markRead();

        Logger.log(`Added ticker ${ticker} with subject: ${subject}`);

        try {
          updateTicker(sheet, firstEmptyRow);
        } catch (error) {
          Logger.log(`Error updating ticker: ${error.message}`);
        }

        thread.removeLabel(label);
        thread.addLabel(archiveLabel);
      }
    });

    if (!tickerFound) {
      Logger.log(`No ticker found for thread: ${thread.getFirstMessageSubject()}. Marking as processed.`);
      thread.removeLabel(label);
      thread.addLabel(archiveLabel);
    }
  });
}

// Function to extract ticker from the subject
function extractTickerFromSubject(subject) {
  const regex = /\$[A-Za-z0-9]+/;
  const match = subject.match(regex);
  return match ? match[0] : null;
}

// Function to extract ticker from the body
function extractTickerFromBody(body) {
  const regex = /\$[A-Za-z0-9]+/;
  const match = body.match(regex);
  return match ? match[0] : null;
}

// Function to extract the first URL from the body
function extractUrlFromBody(body) {
  const regex = /(https?:\/\/[^\s]+)/; // Match "http" or "https" followed by non-whitespace characters
  const match = body.match(regex);
  return match ? match[0] : null; // Return the first URL found, or null if none exists
}

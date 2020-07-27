/** Accesses UI and creates menu items for the addon.s */
export function createMenu(): void {
  const ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu("Custom Menu")
    .addItem("Download Rows to PDF Purchase Form", "downloadSheet")
    .addSeparator()
    .addItem("Read Row", "readRows")
    .addSeparator()
    .addItem("Send Email to Treasurer", "sendEmail")
    .addToUi();
}

export function downloadSheet(): void {
  // read the selected rows
  // duplicate the Template PO sheet
  // Write the row info to the duplicated sheet.

  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
    .alert("You clicked the download pdf item!");
}

export function sendEmail(): void {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
    .alert("You clicked the send email item!");
}

export function readRows(): void {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  if (range == null) {
    SpreadsheetApp.getUi().alert(
      "Please select a range to read from. You can select multiple rows by pressing shift and clicking."
    );
  } else {
    const data = range.getValues();

    for (let i = 0; i < data.length; i++) {
      Logger.log("Item name: " + data[i][0]);
      Logger.log("Vendor number: " + data[i][1]);
      Logger.log("Vendor URL: " + data[i][2]);
      Logger.log("Cost: " + data[i][3]);

      Logger.log("Quantity: " + data[i][4]);
      Logger.log("Item URL: " + data[i][5]);
      Logger.log("Project: " + data[i][6]);
      Logger.log("Date Needed: " + data[i][7]);
    }
    SpreadsheetApp.getUi().alert(Logger.getLog());
  }
}

// TODO: Add data validation

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
    // check if range is valid
    let rangeValid = false;
    // Make sure at least 8 Columns

    if (range.getNumRows() <= 12) {
      if (range.getNumColumns() >= 8) {
        rangeValid = true;
        // Loop through each row and check if the Cell Notation starts with A and 8th Column is H
        for (let j = 0; j < range.getNumRows(); j++) {
          if (
            !(
              range
                .getCell(j + 1, 1)
                .getA1Notation()
                .startsWith("A") &&
              range
                .getCell(j + 1, 8)
                .getA1Notation()
                .startsWith("H")
            )
          ) {
            rangeValid = false;
            SpreadsheetApp.getUi().alert(
              "One of the rows in the selected range does include the correct column selection. 1st Column should be A and 8th Column should be H"
            );
          }
        }
      } else {
        SpreadsheetApp.getUi().alert("Please select the entire row.");
      }
    } else {
      SpreadsheetApp.getUi().alert(
        "If you have more than 12 items, please only select 12 at a time for one vendor and repeat process for more items."
      );
    }

    if (rangeValid) {
      for (let i = 0; i < 8; i++) {
        Logger.log("Item name: " + range.getCell(i + 1, 1).getValue());
        Logger.log("Vendor number: " + range.getCell(i + 1, 2).getValue());
        Logger.log("Vendor URL: " + range.getCell(i + 1, 3).getValue());
        Logger.log("Cost: " + range.getCell(i + 1, 4).getValue());

        Logger.log("Quantity: " + range.getCell(i + 1, 5).getValue());
        Logger.log("Item URL: " + range.getCell(i + 1, 6).getValue());
        Logger.log("Project: " + range.getCell(i + 1, 7).getValue());
        Logger.log("Date Needed: " + range.getCell(i + 1, 8).getValue());
      }
      SpreadsheetApp.getUi().alert(Logger.getLog());
    } else {
      //update alert
      SpreadsheetApp.getUi().alert(
        "You have selected invalid columns to read from. Please click a row number to select entire row."
      );
    }
  }
}

// TODO: Add data validation

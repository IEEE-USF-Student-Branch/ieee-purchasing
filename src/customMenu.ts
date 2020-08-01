/** Accesses UI and creates menu items for the addon.s */
export function createMenu(): void {
  const ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu("Custom Menu")
    .addItem("Duplicate Template P.O.", "duplicateTemplate")
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
  const rangeList = sheet.getActiveRangeList();
  let rangeArray;
  let rangeValid = true;
  if (rangeList != null) {
    rangeArray = rangeList.getRanges();

    let numberOfRows = 0;

    for (let r = 0; r < rangeArray?.length; r++) {
      // check if each range is valid
      const range = rangeArray[r];
      numberOfRows += range.getNumRows();
      if (range == null) {
        SpreadsheetApp.getUi().alert(
          "Please select a range to read from. You can select multiple rows by pressing shift and clicking."
        );
        rangeValid = false;
      } else {
        // Make sure at least 8 Columns
        if (range.getNumColumns() >= 8) {
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
              break;
            }
          }
        } else {
          SpreadsheetApp.getUi().alert("Please select the entire row.");

          rangeValid = false;
          break;
        }
      }
    }
    if (numberOfRows > 12) {
      SpreadsheetApp.getUi().alert(
        "If you have more than 12 items, please only select 12 at a time for one vendor and repeat process for more items."
      );
      rangeValid = false;
    }
    if (rangeValid) {
      for (let r = 0; r < rangeArray?.length; r++) {
        const validRange = rangeArray[r];
        for (let i = 0; i < validRange.getNumRows(); i++) {
          Logger.log("Item name: " + validRange.getCell(i + 1, 1).getValue());
          Logger.log(
            "Vendor number: " + validRange.getCell(i + 1, 2).getValue()
          );
          Logger.log("Vendor URL: " + validRange.getCell(i + 1, 3).getValue());
          Logger.log("Cost: " + validRange.getCell(i + 1, 4).getValue());

          Logger.log("Quantity: " + validRange.getCell(i + 1, 5).getValue());
          Logger.log("Item URL: " + validRange.getCell(i + 1, 6).getValue());
          Logger.log("Project: " + validRange.getCell(i + 1, 7).getValue());
          Logger.log("Date Needed: " + validRange.getCell(i + 1, 8).getValue());
        }
      }
    }
  }
  SpreadsheetApp.getUi().alert(Logger.getLog());
}

// TODO: Add data validation

// Method to duplicate the template PO and rename with Current Date
export function duplicateTemplate(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const template = spreadsheet.getSheetByName("Template PO");
  if (template !== null) {
    spreadsheet.setActiveSheet(template);
  } else {
    SpreadsheetApp.getUi().alert("Sheet 'Template PO' Not Found");
    return;
  }
  const date = new Date();
  const dateString = date.toLocaleString("en-us", {
    year: "numeric",
    month: "numeric",
    day: "numeric",
  });
  spreadsheet.duplicateActiveSheet();
  try {
    spreadsheet.renameActiveSheet(dateString + " IEEE USF PO");
  } catch (error) {
    spreadsheet.toast("SheetName already taken. Please rename manually.");
  }
}

/** Accesses UI and creates menu items for the addon.s */
export function createMenu(): void {
  const ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu("Custom Menu")
    .addItem("Duplicate Template P.O.", "duplicateTemplate")
    .addItem("Download Rows to PDF Purchase Form", "downloadSheet")
    .addSeparator()
    .addItem("Create a new PO", "createNewPO")
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
  MailApp.sendEmail(
    "ieee.usf.tampa@gmail.com",
    "PO Email",
    "Here is a PO from IEEE Purchasing Portal. Ignore this email."
  );
}

function readRows(): Array<Array<string | number>> | null {
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
    // Check if more than 12 rows because that's the max for 1 PO
    if (numberOfRows > 12) {
      SpreadsheetApp.getUi().alert(
        "If you have more than 12 items, please only select 12 at a time for one vendor and repeat process for more items."
      );
      rangeValid = false;
    }
    if (rangeValid) {
      // Create a 1d array with number of rows as columns
      const items = new Array(numberOfRows);
      for (let i = 0; i < items.length; i++) {
        items[i] = new Array(8); // Create a 2d array from 1d array for each of the values.
      }
      let currentRow = 0;
      for (let r = 0; r < rangeArray?.length; r++) {
        const validRange = rangeArray[r];
        for (let i = 0; i < validRange.getNumRows(); i++) {
          items[currentRow][0] = validRange.getCell(i + 1, 1).getValue(); // Item Name
          items[currentRow][1] = validRange.getCell(i + 1, 2).getValue(); // Vendor Name and Number
          items[currentRow][2] = validRange.getCell(i + 1, 3).getValue(); // Vendor URL
          items[currentRow][3] = validRange.getCell(i + 1, 4).getValue(); // Cost
          items[currentRow][4] = validRange.getCell(i + 1, 5).getValue(); // Quantity
          items[currentRow][5] = validRange.getCell(i + 1, 6).getValue(); // Item URL
          items[currentRow][6] = validRange.getCell(i + 1, 7).getValue(); // Project
          items[currentRow][7] = validRange.getCell(i + 1, 8).getValue(); // Date Needed
          currentRow++;
        }
      }
      return items;
    } else {
      return null;
    }
  } else {
    return null;
  }
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
  // Get the current Date and Format to Rename PO
  const date = new Date();
  const dateString = date.toLocaleString("en-us", {
    year: "numeric",
    month: "numeric",
    day: "numeric",
  });
  spreadsheet.duplicateActiveSheet(); // Create copy of Template
  try {
    spreadsheet.renameActiveSheet(dateString + " IEEE USF PO");
  } catch (error) {
    spreadsheet.toast("SheetName already taken. Please rename manually.");
  }
}

// Populate PO with data from read rows
function populatePO(data: Array<Array<string | number>>): void {
  const po = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Loop through each item and add proper fields
  for (let i = 0; i < data.length; i++) {
    po.getRange("B" + (50 + i)).setValue(data[i][0]); // Set Item Name
    po.getRange("H" + (50 + i)).setValue(data[i][5]); // Set Link or Item No.
    po.getRange("M" + (50 + i)).setValue(data[i][4]); // Set Qty.
    po.getRange("O" + (50 + i)).setValue(data[i][3]); // Set Cost
  }
  po.getRange("M38").setValue(data[0][7]); // Set Need by Date. Should be the same for all items. (Validated before)
  po.getRange("J42").setValue(data[0][1]); // Set Vendor Name. Should be the same for all items. (Validated before)
  po.getRange("J44").setValue(data[0][2]); // Set Vendor URL. Should be the same for all items. (Validated before)
  po.getRange("F14").setValue(data[0][6]); // Set Vendor URL. Should be the same for all items. (Validated before)
}

export function createNewPO(): void {
  // Read Rows and save to an array of values.
  // Duplicate Template PO
  // Populate New PO.
  const data = readRows();
  if (data !== null) {
    duplicateTemplate(); // Duplicates PO and sets active sheet to new PO

    // Uncomment below to see what data is.
    // Logger.log(data);
    // SpreadsheetApp.getUi().alert(Logger.getLog());

    // Populate PO with Items.
    populatePO(data);
  } else {
    SpreadsheetApp.getUi().alert("Data returned is NULL");
  }
}

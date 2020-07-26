// You can access any of the global GAS objects in this file. You can also
// import local files or external dependencies:
import { helloWorld } from "./example";

console.log(helloWorld);

// Simple Triggers: These five export functions are reserved export function names that are
// called by Google Apps when the corresponding event occurs. You can safely
// delete them if you won't be using them, but don't use the same export function names
// for anything else.
// See: https://developers.google.com/apps-script/guides/triggers

function menuItem1(): void {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
    .alert("You clicked the download pdf item!");
}

function menuItem2(): void {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
    .alert("You clicked the send email item!");
}

function readRows(): void {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
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

export function onOpen(
  e:
    | GoogleAppsScript.Events.DocsOnOpen
    | GoogleAppsScript.Events.SlidesOnOpen
    | GoogleAppsScript.Events.SheetsOnOpen
    | GoogleAppsScript.Events.FormsOnOpen
): void {
  console.log(e);
  const ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu("Custom Menu")
    .addItem("Download Rows to PDF Purchase Form", "menuItem1")
    .addItem("Read Row", "readRows")
    .addSeparator()
    .addItem("Send Email to Treasurer", "menuItem2")
    .addToUi();
}

export function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit): void {
  console.log(e);
}

export function onInstall(e: GoogleAppsScript.Events.AddonOnInstall): void {
  console.log(e);
}

export function doGet(e: GoogleAppsScript.Events.DoGet): void {
  console.log(e);
}

export function doPost(e: GoogleAppsScript.Events.DoPost): void {
  console.log(e);
}

global.onOpen = onOpen;
global.onEdit = onEdit;
global.onInstall = onInstall;
global.doPost = doPost;
global.doGet = doGet;
// Be sure to define all custom functions as global variables
global.menuItem1 = menuItem1;
global.menuItem2 = menuItem2;
global.readRows = readRows;

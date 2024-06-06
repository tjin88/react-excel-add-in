/* eslint-disable office-addins/load-object-before-read */
// eslint-disable-next-line no-redeclare
/* global Excel, document */

import { handleCellChange } from "./eventHandlers";

// ########################### MAIN FUNCTIONS ###########################

// Make the background of the selected range white --> Currently unused.
export async function backgroundWhite() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("address");
      range.format.fill.color = "white";
      await context.sync();
    });
  } catch (error) {
    console.error(error);
    document.getElementById("message").innerText += `Error: ${error.message}\n`;
  }
}

// Validate the type for each cell in the selected range
export async function validate() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["address", "values"]);
      await context.sync();

      // Check each cell in the 2D matrix (of selected values)
      for (let i = 0; i < range.values.length; i++) {
        for (let j = 0; j < range.values[i].length; j++) {
          const value = range.values[i][j];

          // All types: string, number, boolean, date
          const all_types = ["string", "number", "boolean", "date"];
          const type = all_types[0];
          const isValid = checkIsValid(value, type);
          const message = isValid ? `Valid ${type}` : `Invalid ${type}`;

          // Determine the cell to the right of the current cell
          const currentCell = range.getCell(i, j);
          const messageCell = currentCell.getOffsetRange(0, 1);
          messageCell.values = [[message]];

          currentCell.format.font.color = isValid ? "black" : "red";
          messageCell.format.font.color = isValid ? "black" : "red";
          document.getElementById("message").innerText += `Cell[${i}][${j}] = ${isValid ? "B" : "R"}.\n`;
        }
      }

      await context.sync();
    });
  } catch (error) {
    console.error(error);
    document.getElementById("message").innerText += `Error: ${error.message}\n`;
  }
}

// Hide rows based on the answer in column E  --> Currently unused.
export async function hideRows() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      // TODO: Update this range to match the range of the data. Ideally it's dynamic.
      const range = sheet.getRange("E3:E1000");
      range.load("values");
      await context.sync();

      if (!range.values) {
        document.getElementById("message").innerText += "Range values are null or empty.\n";
        return;
      }

      for (let i = 0; i < range.values.length; i++) {
        const answer = range.values[i][0];

        // Check if the answer is a valid string for logging
        if (typeof answer !== "string") {
          document.getElementById("message").innerText += `Cell[E${i + 3}] is not a valid string.\n`;
        } else {
          const shouldHide = checkIsValid(answer, "boolean") && answer.toLowerCase() === "no";
          if (shouldHide) {
            const rowToHide = i + 3 + 1;
            const range = sheet.getRange(`${rowToHide}:${rowToHide}`);
            range.rowHidden = true;
            document.getElementById("message").innerText += `Row ${rowToHide}:${rowToHide} hidden.\n`;
          }
        }
      }

      await context.sync();
    });
  } catch (error) {
    console.error(error);
    document.getElementById("message").innerText += `Error: ${error.message}\n`;
  }
}

// Unhide rows 1:100
export async function unhideRows() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.worksheets.getActiveWorksheet().getRange("1:100");
      range.rowHidden = false;
      document.getElementById("message").innerText += `Rows 1:100 are unhidden.\n`;
    });
  } catch (error) {
    console.error(error);
    document.getElementById("message").innerText += `Error: ${error.message}\n`;
  }
}

// Clear any messages
export async function clearMessage() {
  document.getElementById("message").innerText = "";
}

// Delete the sheet {sheetName} if it exists
export async function deleteSheet(sheetName) {
  try {
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      const sheet = sheets.getItemOrNullObject(sheetName);
      sheet.load("name");
      await context.sync();
      if (!sheet.isNullObject) {
        sheet.delete();
        await context.sync();
        document.getElementById("message").innerText += `Sheet "${sheetName}" has been deleted.\n`;
      } else {
        document.getElementById("message").innerText += `Sheet "${sheetName}" does not exist.\n`;
      }
    });
  } catch (error) {
    console.error(error);
    document.getElementById("message").innerText += `Error: ${error.message}\n`;
  }
}

// Create a "Questionaire" sheet with questions for borrower
// This function currently encounters errors if there is an existing sheet called "Questionaire"
export async function questionaire(questions) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.add("Questionaire");

      // Make the entire sheet white
      const entireRange = sheet.getRange("A1:ZZ1000");
      entireRange.format.fill.color = "white";

      // Table Header
      const headerRange = sheet.getRange("B2:E2");
      headerRange.values = [["Num", "Question", "Method", "Answer"]];
      headerRange.format.font.bold = true;

      // Add the questions to table, excluding the last element (used for Hidden/Visible row)
      const slicedQuestions = questions.map((q) => q.slice(0, -2));
      const questionsRange = sheet.getRange(`B3:E${3 + slicedQuestions.length - 1}`);
      questionsRange.values = slicedQuestions;

      // Add border around the table --> B2:E2 = header, B3:E${3 + slicedQuestions.length - 1} = questions
      const tableRange = sheet.getRange(`B2:E${3 + slicedQuestions.length - 1}`);

      // Set the borders for the table range
      const borderItems = [
        Excel.BorderIndex.edgeTop,
        Excel.BorderIndex.edgeBottom,
        Excel.BorderIndex.edgeLeft,
        Excel.BorderIndex.edgeRight,
      ];

      borderItems.forEach((borderItem) => {
        const border = tableRange.format.borders.getItem(borderItem);
        border.style = Excel.BorderLineStyle.thin;
        border.color = "black";
      });

      // Adjust column widths
      sheet.getRange("B:D").getEntireColumn().format.autofitColumns();
      sheet.getRange("E:E").getEntireColumn().format.columnWidth = 500;

      const range = sheet.getRange("D:D");
      range.columnHidden = true;

      // Left justify all values in column E (numbers get sent to the right by default)
      const columnERange = sheet.getRange("E:E");
      columnERange.format.horizontalAlignment = Excel.HorizontalAlignment.left;

      // Make the table headers (B2:E2), and questions (B3:D{3 + question length}) read-only
      sheet.getUsedRange().format.protection.locked = false;
      const lockedHeaders = sheet.getRange(`B2:E2`);
      lockedHeaders.format.protection.locked = true;
      const lockedQuestions = sheet.getRange(`B3:D${3 + slicedQuestions.length - 1}`);
      lockedQuestions.format.protection.locked = true;
      sheet.protection.protect();

      await context.sync();

      // Hide rows with "Hidden" in the last element of each question array
      for (let i = 0; i < questions.length; i++) {
        const lastValue = questions[i][questions[i].length - 1];
        if (lastValue === "Hidden") {
          const rowIndex = i + 3; // Adjust for the header row and 0-based index
          const rowRange = sheet.getRange(`${rowIndex}:${rowIndex}`);
          rowRange.rowHidden = true;
        }
      }
      await context.sync();

      // Add event listener for changes in the "Questionaire" sheet
      sheet.onChanged.add(handleCellChange);
      await context.sync();

      document.getElementById("message").innerText += `Questionaire sheet created and event listener added.\n`;
    });
  } catch (error) {
    console.error(error);
    document.getElementById("message").innerText += `Error: ${error.message}\n`;
  }
}

// Create a "Questionaire_v2" sheet with questions for borrower
// This function currently encounters errors if there is an existing sheet called "Questionaire_v2"
export async function questionaire_v2(questions) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.add("Questionaire_v2");

      // Make the entire sheet white
      const entireRange = sheet.getRange("A1:ZZ1000");
      entireRange.format.fill.color = "white";

      // Table Header
      const headerRange = sheet.getRange("B2:D2");
      headerRange.values = [["Num", "Question", "Answer"]];
      headerRange.format.font.bold = true;

      // Add the questions to table, excluding the last element (used for Hidden/Visible row)
      const slicedQuestions = questions.map((q) => q.slice(0, -2));
      const questionsRange = sheet.getRange(`B3:D${3 + slicedQuestions.length - 1}`);
      questionsRange.values = slicedQuestions;

      // Add border around the table --> B2:E2 = header, B3:E${3 + slicedQuestions.length - 1} = questions
      const tableRange = sheet.getRange(`B2:D${3 + slicedQuestions.length - 1}`);

      // Set the borders for the table range
      const borderItems = [
        Excel.BorderIndex.edgeTop,
        Excel.BorderIndex.edgeBottom,
        Excel.BorderIndex.edgeLeft,
        Excel.BorderIndex.edgeRight,
      ];

      borderItems.forEach((borderItem) => {
        const border = tableRange.format.borders.getItem(borderItem);
        border.style = Excel.BorderLineStyle.thin;
        border.color = "black";
      });

      // Adjust column widths
      sheet.getRange("B:C").getEntireColumn().format.autofitColumns();
      sheet.getRange("D:D").getEntireColumn().format.columnWidth = 500;

      // Left justify all values in column D (numbers get sent to the right by default)
      const columnERange = sheet.getRange("D:D");
      columnERange.format.horizontalAlignment = Excel.HorizontalAlignment.left;

      // Make the table headers (B2:E2), and questions (B3:D{3 + question length}) read-only
      sheet.getUsedRange().format.protection.locked = false;
      const lockedHeaders = sheet.getRange(`B2:D2`);
      lockedHeaders.format.protection.locked = true;
      const lockedQuestions = sheet.getRange(`B3:C${3 + slicedQuestions.length - 1}`);
      lockedQuestions.format.protection.locked = true;
      sheet.protection.protect();

      await context.sync();

      // TODO: Add Flowpoint DSL here to hide any questions/rows
      await context.sync();

      // Add event listener for changes in the "Questionaire" sheet
      sheet.onChanged.add(handleCellChange);
      await context.sync();

      document.getElementById("message").innerText += `Questionaire sheet created and event listener added.\n`;
    });
  } catch (error) {
    console.error(error);
    document.getElementById("message").innerText += `Error: ${error.message}\n`;
  }
}

// ########################### HELPER FUNCTIONS ###########################
// Helper function to check if a value is of a certain type
// Boolean isn't T/F, but Yes/No
// numberCapped is number capped at 100
export function checkIsValid(value, type) {
  switch (type) {
    case "numberCapped":
      return typeof value === "number" && !isNaN(value) && value <= 100;
    case "number":
      return typeof value === "number" && !isNaN(value);
    case "date":
      return value instanceof Date && !isNaN(value.getTime());
    case "boolean":
      return typeof value === "string" && ["yes", "no"].includes(value.toLowerCase());
    default:
      return typeof value === type;
  }
}

// eslint-disable-next-line no-redeclare, @typescript-eslint/no-unused-vars
/* global document, Excel, Office */

import { checkIsValid } from "./functions";

export const initializeOffice = () => {
  document.getElementById("app-body").style.display = "flex";
};

// Event handler for cell changes
// TODO: Update handler to allow changedRange to be a range of cells rather than just one cell
export async function handleCellChange(event) {
  try {
    await Excel.run(async (context) => {
      document.getElementById("message").innerText += `handleCellChange called!.\n`;
      const sheet = context.workbook.worksheets.getItem("Questionaire");
      document.getElementById("message").innerText += `handleCellChange called!.\n`;

      // Unprotect the sheet to make changes
      sheet.protection.unprotect();

      const changedRange = sheet.getRange(event.address);
      changedRange.load(["values", "address"]);
      await context.sync();

      // Check if the changed cell is in column E
      if (changedRange.address.split("!")[1].startsWith("E")) {
        const answer = changedRange.values[0][0];
        document.getElementById("message").innerText += `answer: ${answer}.\n`;

        const rowIndex = parseInt(event.address.match(/\d+/)[0]);
        document.getElementById("message").innerText += `rowIndex: ${rowIndex}.\n`;

        // Load the value from column D of the same row
        const methodCell = sheet.getRange(`D${rowIndex}`);
        methodCell.load("values");
        await context.sync();
        const method = methodCell.values[0][0];
        document.getElementById("message").innerText += `method: ${method}.\n`;

        let shouldHide = false;
        let isValid = false;

        switch (method) {
          case "Num & capped":
            isValid = checkIsValid(answer, "numberCapped");
            document.getElementById("message").innerText += `Cell ${event.address} is a valid number: ${isValid}.\n`;
            break;
          case "Num":
            isValid = checkIsValid(answer, "number");
            document.getElementById("message").innerText += `Cell ${event.address} is a valid number: ${isValid}.\n`;
            break;
          case "String":
            isValid = checkIsValid(answer, "string");
            document.getElementById("message").innerText += `Cell ${event.address} is a valid string: ${isValid}.\n`;
            break;
          case "Bool & Hide No":
            isValid = checkIsValid(answer, "boolean");
            document.getElementById("message").innerText += `Cell ${event.address} is a valid boolean: ${isValid}.\n`;

            shouldHide = isValid && answer.toLowerCase() === "no";
            break;
          case "Bool & Hide Yes":
            isValid = checkIsValid(answer, "boolean");
            document.getElementById("message").innerText += `Cell ${event.address} is a valid boolean: ${isValid}.\n`;

            shouldHide = isValid && answer.toLowerCase() === "yes";
            break;
          default:
            document.getElementById("message").innerText += `Cell ${event.address} is not a valid method.\n`;
            return;
        }

        changedRange.format.font.color = isValid ? "black" : "red";

        const rowRange = sheet.getRange(`${rowIndex + 1}:${rowIndex + 1}`);
        rowRange.rowHidden = shouldHide;
        document.getElementById("message").innerText += `Row ${rowIndex + 1} ${
          shouldHide ? "hidden" : "shown"
        } due to cell change.\n`;

        // Left justify all values (numbers get sent to the right by default)
        changedRange.format.horizontalAlignment = Excel.HorizontalAlignment.left;
      }

      await context.sync();
      // Protect the sheet again
      sheet.protection.protect();
    });
  } catch (error) {
    console.error(error);
    document.getElementById("message").innerText += `Error: ${error.message}\n`;
  }
}
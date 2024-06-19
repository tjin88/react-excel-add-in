// eslint-disable-next-line no-redeclare, @typescript-eslint/no-unused-vars
/* global document, Excel, Office */
import { checkIsValid } from "./functions";

export const initializeOffice = () => {
  document.getElementById("app-body").style.display = "flex";
};

// Load any script from a CDN
// const loadScript = (src) => {
//   return new Promise((resolve, reject) => {
//     const script = document.createElement("script");
//     script.src = src;
//     script.onload = resolve;
//     script.onerror = reject;
//     document.head.appendChild(script);
//   });
// };

// Event handler for cell changes
// TODO: Update handler to allow changedRange to be a range of cells rather than just one cell
// TODO: Once Mark gives the go-ahead, remove the jQuery test and replace with his DSL
export async function handleCellChange(event) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

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

      // Testing jQuery (to see if any CDN script can be run in the onChange function)
      // loadScript("https://ajax.googleapis.com/ajax/libs/jquery/3.7.1/jquery.min.js")
      //   .then(() => {
      //     console.log("script loaded!");
      //     // simple jQuery test
      //     const all_colours = [
      //       "lightblue",
      //       "lightgreen",
      //       "lightyellow",
      //       "lightcoral",
      //       "lightcyan",
      //       "lightgoldenrodyellow",
      //       "lightgray",
      //       "lightpink",
      //       "lightsalmon",
      //       "lightseagreen",
      //       "lightskyblue",
      //       "lightslategray",
      //       "lightsteelblue",
      //       "lightyellow",
      //     ];
      //     const colour = all_colours[Math.floor(Math.random() * (all_colours.length - 1))];
      //     window.$(".ms-welcome").css("background-color", colour);
      //     window.$("body").css("background-color", colour);
      //   })
      //   .catch((error) => {
      //     console.error("Failed to load the jQuery script:", error);
      //   });

      await context.sync();
      // Protect the sheet again
      sheet.protection.protect();
    });
  } catch (error) {
    console.error(error);
    document.getElementById("message").innerText += `Error: ${error.message}\n`;
  }
}

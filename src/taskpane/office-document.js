/* global Excel console */

const insertText = async (text, cell) => {
  // Write text to the top left cell.
  try {
    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(cell);
      range.values = [[text]];
      range.format.font.color = "black";
      range.format.autofitColumns();
      return context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
};

export default insertText;

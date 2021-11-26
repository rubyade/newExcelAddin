/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // run code when page is ready
    document.getElementById("btn").addEventListener("click", writeData);
  }
});

// the actual function that runs above when ready
export async function writeData() {
  // This is a global function that calls the workbook
  Excel.run((context) => {
    const ws = context.workbook.worksheets.getActiveWorksheet();

    const range = ws.getRange("A1:M2");

    range.values = [
      ["=A1*12", 78, 50, "My single Value", 78, 50, "My single Value", 78, 50, "My single Value", 78, 50, 24],
      ["=A2*12", 78, 50, "My single Value", 78, 50, "My single Value", 78, 50, "My single Value", 78, 50, 24],
    ];

    return context.sync();
    // range.format.fill.color = "#4472C4";
    // range.format.font.color = "white";
    // range.format.autofitColumns();
  });
}



Excel.run(function (context) {
  var sheet = context.workbook.worksheets.getActiveWorksheet();

  var range = sheet.getRange("E3");
  if (range.formulas == [["=D3 / 12"]]) {
  sheet.getRange("F3:M3").copyFrom("E3", Excel.RangeCopyType.all, false, false);
  }
 

  

  return context.sync();
}).catch(errorHandlerFunction);

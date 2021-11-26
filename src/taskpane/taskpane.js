/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // run code when page is ready
    document.getElementById("btn").addEventListener("click", calculateData);
  }
});

// the actual function that runs above when ready
export async function calculateData() {
  // This is a global function that calls the workbook
  Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    const range = sheet.getRange("E3");
    range.formulas = [["=D3 / 12"]];
    range.format.autofitColumns();

    sheet.getRange("F3:M3").copyFrom("E3", Excel.RangeCopyType.values, false, false);

    return context.sync();
  })
}



  
  Excel.run(function (context) {

    const clearRange = sheet.getRange("F3:M3");

    const cell = sheet.getCell(2, 3);

    if (cell === 0) {
      clearRange.clear();
    }

    return context.sync();
  })
 // This is a global function that calls the workbook
// Excel.run(function (context) {
//   const sheet = context.workbook.worksheets.getActiveWorksheet();

//   const range = sheet.getRange("E3");
//   range.formulas = [["=D3 / 12"]];
//   range.format.autofitColumns();

//   sheet.getRange("F3:M3").copyFrom("E3", Excel.RangeCopyType.values, false, false);
  

//   return context.sync();
// }).catch(errorHandlerFunction);

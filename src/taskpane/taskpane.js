/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

let button = document.getElementById("create-table");
button.addEventListener("click", async () => {
  await run();
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getItem("Sheet1");
      let range = sheet.getRange("C3");
      range.values = [[5]];
      await context.sync();
    });
  }
  catch (error) {
    console.error(error);
  }
}

// export async function run() {
//   try {
//     await Excel.run(async (context) => {
//       alert("teste");
//       // const range = context.workbook.getSelectedRange();
//       // range.load("C3");
//       // range.format.fill.color = "yellow";

//       // await context.sync();
//       // console.log(`The range address was ${range.address}.`);
//       let sheet = context.workbook.worksheets.getItem("Sheet1");
//       let range = sheet.getRange("C3");
//       range.values = [[5]];
//       await context.sync();
//     });
//   }
//   catch (error) {
//     console.error(error);
//   }
// }
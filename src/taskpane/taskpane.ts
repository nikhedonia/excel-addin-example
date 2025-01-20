/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */




Office.initialize = async function init()  {
  console.log('init');
  Office.addin.setStartupBehavior(Office.StartupBehavior.load);
  run();
}



Office.onReady(async function doit(info) {

  Office.addin.setStartupBehavior(Office.StartupBehavior.load);

  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;

    run();
  }
});

export async function run() {
  await Excel.run(async (context) => {
    const res = await fetch("http://localhost:3001/")
      .then(x=>x.json())
      .catch(()=>({action:"write", cell: "A7", value: Math.random()}));

    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const targetRange = sheet.getRange(res.cell);

    // Select the range (this will scroll to it)
    targetRange.select();
    targetRange.values = [[res.value]];
    
    targetRange.format.fill.color = "yellow";

    await context.sync();

  }).catch(e=>console.error(e));

  await new Promise(done=>setTimeout(done, 1000));
  await run();
}

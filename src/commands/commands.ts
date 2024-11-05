/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

async function insertParagraph(event) {
  // Implement your custom code here. The following code is a simple Excel example.
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.insertParagraph("Hello World, this is the line added automatically when document is opened.", Word.InsertLocation.end);
      await context.sync();
    });
  } catch (error) {
    // Note: In a production add-in, notify the user through your add-in's UI.
    //console.error(error);
  }

  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
}

async function insertPPT(event) {
  try{
    await PowerPoint.run(async (context) => {
      const shapes: PowerPoint.ShapeCollection = context.presentation.slides.getItemAt(0).shapes;
      const shapeOptions: PowerPoint.ShapeAddOptions = {
        left: 100,
        top: 300,
        height: 300,
        width: 450
      };
      const textbox: PowerPoint.Shape = shapes.addTextBox("Hello, World! This is the first Shape automatically added when opening.", shapeOptions);
  
      await context.sync();
    });
  } catch (error) {
    // Note: In a production add-in, notify the user through your add-in's UI.
    //console.error(error);
  }

  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
}

async function insertExcelTable() {
  await Excel.run(async (context) => {
    context.workbook.worksheets.getItemOrNullObject("Sample").delete();
    const sheet = context.workbook.worksheets.add("Sample");

    const data = [
      ["Table data", "Added", "Through", "Autorun"],
      ["Product", "Qty", "Unit Price", "Total Price"],
      ["Almonds", 2, 7.5, "=C3 * D3"],
      ["Coffee", 1, 34.5, "=C4 * D4"],
      ["Chocolate", 5, 9.56, "=C5 * D5"]
    ];

    const range = sheet.getRange("B1:E5");
    range.values = data;
    range.format.autofitColumns();

    const header = range.getRow(1);
    header.format.fill.color = "#4472C4";
    header.format.font.color = "white";

    sheet.activate();

    await context.sync();
  });
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal() as any;

Office.actions.associate("insertParagraph", insertParagraph);
Office.actions.associate("insertPPT", insertPPT);
Office.actions.associate("insertExcelTable", insertExcelTable);

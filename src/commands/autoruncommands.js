/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

async function insertParagraph(event) {
  Word.run(async (context) => {
    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph(
      "XiangminExecuteFunction works. Button ID=" /*+ event.source.id*/,
      Word.InsertLocation.end
    );
    //context.document.onParagraphAdded.add(paragraphAdded);
    //await context.sync();
    // change the paragraph color to blue.
    //paragraph.font.color = "blue";

    await context.sync();
  });

  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
}
/*
async function paragraphAdded(event: Word.ParagraphAddedEventArgs) {
  await Word.run(async (context) => {
    const paragraph = context.document.body.insertParagraph(
      "onParagraphAdded" /*+ event.source.id*/,
      Word.InsertLocation.end
    );
  });
}*/
  
function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope

Office.actions.associate("insertParagraph", insertParagraph);

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global document, Office, Word */

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("create-content-control").onclick = createContentControl;
    document.getElementById("replace-content-in-control").onclick = replaceContentInControl; 
    document.getElementById("retitle-content-in-control").onclick = renameTitleOfControl;
    document.getElementById("digest-content").onclick = digestContent;
    document.getElementById("run").onclick = run;
  }
});

function createContentControl() {
  Word.run(function (context) {

      // TODO1: Queue commands to create a content control.
      var serviceNameRange = context.document.getSelection();
      var serviceNameContentControl = serviceNameRange.insertContentControl();
      serviceNameContentControl.title = "Signature";
      serviceNameContentControl.tag = "serviceName";
      serviceNameContentControl.appearance = "Tags";
      serviceNameContentControl.color = "blue";
      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function replaceContentInControl() {
  Word.run(function (context) {

      // TODO1: Queue commands to replace the text in the Service Name
      //        content control.
      var serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
      serviceNameContentControl.insertText("Akintunde Pounds", "Replace");

      return context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

async function renameTitleOfControl(){
  return Word.run( async context => {
    const newTitle = document.getElementById("new-title").value;
    const oldTitle = document.getElementById("old-title").value;
    console.log(newTitle, oldTitle)

    let contentControls = context.document.contentControls.getByTitle(oldTitle);
    contentControls.load(`title`);
  
    await context.sync();
    
    contentControls.items.forEach(sig => {
      sig.title = newTitle
      //sig.insertText("Enter Signature Here", "Replace")
    })

    return context.sync();
  })
}

async function digestContent() {
  await Word.run(async (context) => {
    const controlTitle = document.getElementById("control-type").value;
    let contentControls = context.document.contentControls.getByTitle(controlTitle);
    contentControls.load(`text, title, id`);
  
    await context.sync();
    
    contentControls.items.forEach(sig => {
      console.log(`text within ${sig.id} -- ${sig.title}: ${sig.text}`)
    })
    
    await context.sync();
  });
}

export async function run() {
  return Word.run(async context => {
    /**
     * Insert your Word code here
     */
    
    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph(`${Math.random() * 10}`, Word.InsertLocation.end);

    // change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();
  });
}

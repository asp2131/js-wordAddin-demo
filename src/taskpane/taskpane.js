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
    let executableClass = document.getElementsByClassName("clickable-executable")
    console.log(executableClass[0].id)
    for (var i=0; i < executableClass.length; i++) {
      console.log(executableClass[i].id)
      let execContext = executableClass[i].id;
      switch(execContext){
        case 'signature': 
          executableClass[i].onclick = digestSignature
          break;
        case "date-created":
          executableClass[i].onclick = digestCreated
          break;
        case "time-frame":
          executableClass[i].onclick = digestTimeFrame
          break;
        default:
          executableClass[i].onclick = digestOther
      }
      
    }
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

async function digestSignature() {
  return Word.run(async (context) => {
    var contentControlsPrev
    var serviceNameRange = context.document.getSelection();
    
    serviceNameRange.load('text')
    await context.sync();
    if(serviceNameRange.text != ""){
        
      var serviceNameContentControl = serviceNameRange.insertContentControl();
      serviceNameContentControl.title = "Signature"
        serviceNameContentControl.tag = "serviceName";
        serviceNameContentControl.appearance = "Tags";
        serviceNameContentControl.color = "blue";
        
        let inputVal = serviceNameRange.text
        let inputText = document.getElementById('signature')
        inputText.value = ''
        inputText.value += inputVal
      console.log(serviceNameRange.text)
    }
  });
}

async function digestCreated() {
  return Word.run(async (context) => {
    var serviceNameRange = context.document.getSelection();
    serviceNameRange.load('text')
    await context.sync();
    if(serviceNameRange.text != ""){
      var serviceNameContentControl = serviceNameRange.insertContentControl();
      serviceNameContentControl.title = "Created At"
        serviceNameContentControl.tag = "serviceName";
        serviceNameContentControl.appearance = "Tags";
        serviceNameContentControl.color = "green";
        
        let inputVal = serviceNameRange.text
        let inputText = document.getElementById('date-created')
        inputText.value = ''
        inputText.value += inputVal
      console.log(serviceNameRange.text)
    }
  });
}

async function digestTimeFrame() {
  return Word.run(async (context) => {
    var serviceNameRange = context.document.getSelection();
    serviceNameRange.load('text')
    await context.sync();
    if(serviceNameRange.text != ""){
      var serviceNameContentControl = serviceNameRange.insertContentControl();
      serviceNameContentControl.title = "Time Frame"
        serviceNameContentControl.tag = "serviceName";
        serviceNameContentControl.appearance = "Tags";
        serviceNameContentControl.color = "red";
        
        let inputVal = serviceNameRange.text
        let inputText = document.getElementById('time-frame')
        inputText.value = ''
        inputText.value += inputVal
      console.log(serviceNameRange.text)
    }
  });
}


async function digestOther() {
  return Word.run(async (context) => {
    var serviceNameRange = context.document.getSelection();
    serviceNameRange.load('text')
    await context.sync();
    if(serviceNameRange.text != ""){
      var serviceNameContentControl = serviceNameRange.insertContentControl();
      serviceNameContentControl.title = "Other Tags"
        serviceNameContentControl.tag = "serviceName";
        serviceNameContentControl.appearance = "Tags";
        serviceNameContentControl.color = "purple";
        
        let inputVal = serviceNameRange.text
        let inputText = document.getElementById('time-frame')
        inputText.value = ''
        inputText.value += inputVal
      console.log(serviceNameRange.text)
    }
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

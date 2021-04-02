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
    // document.getElementById("digest-content").onclick = digestContent;
    document.getElementById('submit').onclick = sendFile;
    updateStatus("Ready to send file.");
    document.getElementById("run").onclick = run;
  }
});

Office.initialize = function (reason) {

  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {

      // Execute sendFile when submit is clicked
      $('#submit').click(function () {
          sendFile();
      });

      // Update status
      updateStatus("Ready to send file.");
  });
}

// Create a function for writing to the status div.

function updateStatus(message) {
  console.log(message)
  var statusInfo = $('#status');
  statusInfo.innerHTML += message + "<br/>";
}

function sendFile() {
  Office.context.document.getFileAsync("compressed",
      { sliceSize: 100000 },
      function (result) {

          if (result.status == Office.AsyncResultStatus.Succeeded) {

              // Get the File object from the result.
              var myFile = result.value;
              var state = {
                  file: myFile,
                  counter: 0,
                  sliceCount: myFile.sliceCount
              };

              updateStatus("Getting file of " + myFile.size + " bytes");
              getSlice(state);
          }
          else {
              updateStatus(result.status);
          }
      });
}

function getSlice(state) {
  state.file.getSliceAsync(state.counter, function (result) {
      if (result.status == Office.AsyncResultStatus.Succeeded) {
          updateStatus("Sending piece " + (state.counter + 1) + " of " + state.sliceCount);
          sendSlice(result.value, state);
      }
      else {
          updateStatus(result.status);
      }
  });
}

function sendSlice(slice, state) {
  var data = slice.data;

  // If the slice contains data, create an HTTP request.
  if (data) {

      // Encode the slice data, a byte array, as a Base64 string.
      // NOTE: The implementation of myEncodeBase64(input) function isn't
      // included with this example. For information about Base64 encoding with
      // JavaScript, see https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding.
      var fileData = myEncodeBase64(data);

      // Create a new HTTP request. You need to send the request
      // to a webpage that can receive a post.
      var request = new XMLHttpRequest();

      // Create a handler function to update the status
      // when the request has been sent.
      request.onreadystatechange = function () {
          if (request.readyState == 4) {

              updateStatus("Sent " + slice.size + " bytes.");
              state.counter++;

              if (state.counter < state.sliceCount) {
                  getSlice(state);
              }
              else {
                  closeFile(state);
              }
          }
      }

      request.open("POST", "[Your receiving page or service]");
      request.setRequestHeader("Slice-Number", slice.index);

      // Send the file as the body of an HTTP POST
      // request to the web server.
      request.send(fileData);
  }
}


function closeFile(state) {
  // Close the file when you're done with it.
  state.file.closeAsync(function (result) {

      // If the result returns as a success, the
      // file has been successfully closed.
      if (result.status == "succeeded") {
          updateStatus("File closed.");
      }
      else {
          updateStatus("File couldn't be closed.");
      }
  });
}

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

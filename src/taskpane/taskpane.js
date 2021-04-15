/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";


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
  console.log(statusInfo)
  statusInfo.innerHTML += message + "<br/>";
}

function sendFile() {
  console.log(Office.context.document)
  Office.context.document.getFileAsync("compressed",
      { sliceSize: 100000 },
      function (result) {
          console.log(result)
          if (result.status == Office.AsyncResultStatus.Succeeded) {

              // Get the File object from the result.
              var myFile = result.value;
              console.log(result.value)
              var state = {
                  file: myFile,
                  counter: 0,
                  sliceCount: myFile.sliceCount
              };
              console.log('line 65')
              updateStatus("Getting file of " + myFile.size + " bytes");
              getSlice(state);
          }
          else {
            console.log('line 70')
              updateStatus(result.status);
          }
      });
}

function getSlice(state) {
  console.log('getting Slice...')
  state.file.getSliceAsync(state.counter, function (result) {
    console.log(result)
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
  console.log('sending slice...')
  var data = slice.data;

  // If the slice contains data, create an HTTP request.
  if (data) {

      // Encode the slice data, a byte array, as a Base64 string.
      // NOTE: The implementation of myEncodeBase64(input) function isn't
      // included with this example. For information about Base64 encoding with
      // JavaScript, see https://developer.mozilla.org/docs/Web/JavaScript/Base64_encoding_and_decoding.
      console.log(data)
      var fileData = btoa(data);
      console.log(fileData)

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

      request.open("POST", "http://localhost:4000", true);
      request.setRequestHeader("Slice-Number", slice.index);
      //request.setRequestHeader('Content-Type', 'application/xml')

      // Send the file as the body of an HTTP POST
      // request to the web server.
      request.send(fileData);
  }
}


function closeFile(state) {
  console.log('closing file...')
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


/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
Office.initialize = () => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = run;
  document.getElementById("writeTextButton").onclick = writeText;
  document.getElementById("scrollDownButton").onclick = scrollDown;
};

async function run() {
  try {
    await Excel.run(async context => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

// Reads data from current document selection and displays a notification
function writeText() {
  Excel.run(function(context) {
    var range = context.workbook.getSelectedRange();
    var saturation = (<HTMLInputElement>document.getElementById("saturation")).value;
    if (isNaN(+saturation) || saturation == "") {
      //TODO: Display error on dialog or something
      range.getCell(0, 0).values = [["Saturation is NaN"]];
    } else if (+saturation < 0 || +saturation > 255) {
      //TODO: Display error on dialog or something
      range.getCell(0, 0).values = [["Saturation is out of range"]];
    } else {
      var colorSelected = (<HTMLInputElement>document.getElementById("color")).value;

      var saturationInHex = parseInt(saturation).toString(16);
      if (saturationInHex.length < 2) {
        saturationInHex = "0" + saturationInHex;
      }

      if (+saturation > 128) {
        range.format.font.color = "black";
      } else {
        range.format.font.color = "white";
      }
      //range.format.fill.color = "#873F11";
      switch (colorSelected) {
        case "Red":
          range.format.fill.color = "#" + saturationInHex + "0000";
          //range.values = [["#" + saturationInHex + "0000"]];
          break;
        case "Green":
          range.format.fill.color = "#00" + saturationInHex + "00";
          // range.values = [["#00" + saturationInHex + "00"]];
          break;
        case "Blue":
          range.format.fill.color = "#0000" + saturationInHex;
          //range.values = [["#0000" + saturationInHex]];
          break;
        case "Gray":
          range.format.fill.color = "#" + saturationInHex + saturationInHex + saturationInHex;
          //range.values = [["#" + saturationInHex + saturationInHex + saturationInHex]];
          break;
      }
    }

    return context.sync();
  }).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
  // Office.context.document.setSelectedDataAsync("Data hereee",
  //     function (asyncResult) {
  //         var error = asyncResult.error;
  //         if (asyncResult.status === "failed") {
  //             //show error. Upcoming displayDialog API will help here.
  //         }
  //         else {
  //             //show success.Upcoming displayDialog API will help here.
  //         }
  //     });
}
//Scroll down for easier viewing input list
function scrollDown() {
  Excel.run(async function(context) {
    var scrollDownInput = (<HTMLInputElement>document.getElementById("scrollDownInput")).value;
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = context.workbook.getSelectedRange();
    if (scrollDownInput == "") {
      range.getCell(0, 0).values = [["Please input something"]];
      await context.sync();
    } else {
      range.load(["rowCount", "columnCount"]);
      // range.getCell(0,0).values = [[scrollDownInput]];
      await context.sync();
      for (var i = 0; i < range.rowCount; i++) {
        for (var j = 0; j < range.columnCount; j++) {
          range.getCell(i, j).values = [[scrollDownInput]];
        }
      }
      await context.sync();
      range = range.getRowsBelow(15);
      range.select();
      await context.sync();
      range = range.getRowsBelow();
      range.select();
      await context.sync();
      range = range.getRowsAbove(14);
      range.select();
      await context.sync();
      range = range.getRowsAbove();
      range.select();
      await context.sync();
    }
    focusInputScrolldown();
  }).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function focusInputScrolldown() {
  var input = document.getElementById("scrollDownInput");
  (<HTMLInputElement>input).select();
}

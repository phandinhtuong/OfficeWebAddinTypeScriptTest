/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
import { colorize } from "../taskpane/taskpane";
function colorizeCommandFunction(event: Office.AddinCommands.Event) {
  // Your code goes here
  console.log("Colorizing");

  colorize();
  console.log("Colorized");
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}
function ttsCommandFunction(event: Office.AddinCommands.Event) {
  // Your code goes here
  console.log("TTS Working");
  document.getElementById("textToSpeech").click();
  // textToSpeech();
  console.log("TTS Worked");
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}
function docsPropsCommandFunction(event: Office.AddinCommands.Event) {
  // Your code goes here
  console.log("docsProps Working");
  document.getElementById("docsPropsButton").click();
  // textToSpeech();
  console.log("docsProps Worked");
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}
function scrollDownCommandFunction(event: Office.AddinCommands.Event) {
  // Your code goes here
  console.log("scrollDown Working");
  document.getElementById("scrollDownButton").click();
  // textToSpeech();
  console.log("scrollDown Worked");
  // Be sure to indicate when the add-in command function is complete
  event.completed();
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

// the add-in command functions need to be available in global scope
g.colorizeCommandFunction = colorizeCommandFunction;
g.ttsCommandFunction = ttsCommandFunction;
g.docsPropsCommandFunction = docsPropsCommandFunction;
g.scrollDownCommandFunction = scrollDownCommandFunction;

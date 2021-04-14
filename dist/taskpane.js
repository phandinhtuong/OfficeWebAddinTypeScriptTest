/******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, { enumerable: true, get: getter });
/******/ 		}
/******/ 	};
/******/
/******/ 	// define __esModule on exports
/******/ 	__webpack_require__.r = function(exports) {
/******/ 		if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 			Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 		}
/******/ 		Object.defineProperty(exports, '__esModule', { value: true });
/******/ 	};
/******/
/******/ 	// create a fake namespace object
/******/ 	// mode & 1: value is a module id, require it
/******/ 	// mode & 2: merge all properties of value into the ns
/******/ 	// mode & 4: return value when already ns object
/******/ 	// mode & 8|1: behave like require
/******/ 	__webpack_require__.t = function(value, mode) {
/******/ 		if(mode & 1) value = __webpack_require__(value);
/******/ 		if(mode & 8) return value;
/******/ 		if((mode & 4) && typeof value === 'object' && value && value.__esModule) return value;
/******/ 		var ns = Object.create(null);
/******/ 		__webpack_require__.r(ns);
/******/ 		Object.defineProperty(ns, 'default', { enumerable: true, value: value });
/******/ 		if(mode & 2 && typeof value != 'string') for(var key in value) __webpack_require__.d(ns, key, function(key) { return value[key]; }.bind(null, key));
/******/ 		return ns;
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = "./src/taskpane/taskpane.ts");
/******/ })
/************************************************************************/
/******/ ({

/***/ "./assets/icon-16.png":
/*!****************************!*\
  !*** ./assets/icon-16.png ***!
  \****************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

module.exports = __webpack_require__.p + "assets/icon-16.png";

/***/ }),

/***/ "./assets/icon-32.png":
/*!****************************!*\
  !*** ./assets/icon-32.png ***!
  \****************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

module.exports = __webpack_require__.p + "assets/icon-32.png";

/***/ }),

/***/ "./assets/icon-80.png":
/*!****************************!*\
  !*** ./assets/icon-80.png ***!
  \****************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

module.exports = __webpack_require__.p + "assets/icon-80.png";

/***/ }),

/***/ "./src/taskpane/taskpane.ts":
/*!**********************************!*\
  !*** ./src/taskpane/taskpane.ts ***!
  \**********************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";


var __awaiter = this && this.__awaiter || function (thisArg, _arguments, P, generator) {
  function adopt(value) {
    return value instanceof P ? value : new P(function (resolve) {
      resolve(value);
    });
  }

  return new (P || (P = Promise))(function (resolve, reject) {
    function fulfilled(value) {
      try {
        step(generator.next(value));
      } catch (e) {
        reject(e);
      }
    }

    function rejected(value) {
      try {
        step(generator["throw"](value));
      } catch (e) {
        reject(e);
      }
    }

    function step(result) {
      result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected);
    }

    step((generator = generator.apply(thisArg, _arguments || [])).next());
  });
};

var __generator = this && this.__generator || function (thisArg, body) {
  var _ = {
    label: 0,
    sent: function () {
      if (t[0] & 1) throw t[1];
      return t[1];
    },
    trys: [],
    ops: []
  },
      f,
      y,
      t,
      g;
  return g = {
    next: verb(0),
    "throw": verb(1),
    "return": verb(2)
  }, typeof Symbol === "function" && (g[Symbol.iterator] = function () {
    return this;
  }), g;

  function verb(n) {
    return function (v) {
      return step([n, v]);
    };
  }

  function step(op) {
    if (f) throw new TypeError("Generator is already executing.");

    while (_) try {
      if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
      if (y = 0, t) op = [op[0] & 2, t.value];

      switch (op[0]) {
        case 0:
        case 1:
          t = op;
          break;

        case 4:
          _.label++;
          return {
            value: op[1],
            done: false
          };

        case 5:
          _.label++;
          y = op[1];
          op = [0];
          continue;

        case 7:
          op = _.ops.pop();

          _.trys.pop();

          continue;

        default:
          if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) {
            _ = 0;
            continue;
          }

          if (op[0] === 3 && (!t || op[1] > t[0] && op[1] < t[3])) {
            _.label = op[1];
            break;
          }

          if (op[0] === 6 && _.label < t[1]) {
            _.label = t[1];
            t = op;
            break;
          }

          if (t && _.label < t[2]) {
            _.label = t[2];

            _.ops.push(op);

            break;
          }

          if (t[2]) _.ops.pop();

          _.trys.pop();

          continue;
      }

      op = body.call(thisArg, _);
    } catch (e) {
      op = [6, e];
      y = 0;
    } finally {
      f = t = 0;
    }

    if (op[0] & 5) throw op[1];
    return {
      value: op[0] ? op[1] : void 0,
      done: true
    };
  }
};

Object.defineProperty(exports, "__esModule", {
  value: true
});
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
// images references in the manifest

__webpack_require__(/*! ../../assets/icon-16.png */ "./assets/icon-16.png");

__webpack_require__(/*! ../../assets/icon-32.png */ "./assets/icon-32.png");

__webpack_require__(/*! ../../assets/icon-80.png */ "./assets/icon-80.png");
/* global console, document, Excel, Office */
// The initialize function must be run each time a new page is loaded


Office.initialize = function () {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = run;
  document.getElementById("colorizeButton").onclick = colorize;
  document.getElementById("scrollDownButton").onclick = scrollDown;
};

function run() {
  return __awaiter(this, void 0, void 0, function () {
    var error_1;

    var _this = this;

    return __generator(this, function (_a) {
      switch (_a.label) {
        case 0:
          _a.trys.push([0, 2,, 3]);

          return [4
          /*yield*/
          , Excel.run(function (context) {
            return __awaiter(_this, void 0, void 0, function () {
              var range;
              return __generator(this, function (_a) {
                switch (_a.label) {
                  case 0:
                    range = context.workbook.getSelectedRange(); // Read the range address

                    range.load("address"); // Update the fill color

                    range.format.fill.color = "yellow";
                    return [4
                    /*yield*/
                    , context.sync()];

                  case 1:
                    _a.sent();

                    console.log("The range address was " + range.address + ".");
                    return [2
                    /*return*/
                    ];
                }
              });
            });
          })];

        case 1:
          _a.sent();

          return [3
          /*break*/
          , 3];

        case 2:
          error_1 = _a.sent();
          console.error(error_1);
          return [3
          /*break*/
          , 3];

        case 3:
          return [2
          /*return*/
          ];
      }
    });
  });
} // colorize cells with cells' values


function colorize() {
  // Excel.run(function(context) {
  //   var range = context.workbook.getSelectedRange();
  //   var saturation = (<HTMLInputElement>document.getElementById("saturation")).value;
  //   if (isNaN(+saturation) || saturation == "") {
  //     //TODO: Display error on dialog or something
  //     range.getCell(0, 0).values = [["Saturation is NaN"]];
  //   } else if (+saturation < 0 || +saturation > 255) {
  //     //TODO: Display error on dialog or something
  //     range.getCell(0, 0).values = [["Saturation is out of range"]];
  //   } else {
  //     var colorSelected = (<HTMLInputElement>document.getElementById("color")).value;
  //     var saturationInHex = parseInt(saturation).toString(16);
  //     if (saturationInHex.length < 2) {
  //       saturationInHex = "0" + saturationInHex;
  //     }
  //     if (+saturation > 128) {
  //       range.format.font.color = "black";
  //     } else {
  //       range.format.font.color = "white";
  //     }
  //     //range.format.fill.color = "#873F11";
  //     switch (colorSelected) {
  //       case "Red":
  //         range.format.fill.color = "#" + saturationInHex + "0000";
  //         //range.values = [["#" + saturationInHex + "0000"]];
  //         break;
  //       case "Green":
  //         range.format.fill.color = "#00" + saturationInHex + "00";
  //         // range.values = [["#00" + saturationInHex + "00"]];
  //         break;
  //       case "Blue":
  //         range.format.fill.color = "#0000" + saturationInHex;
  //         //range.values = [["#0000" + saturationInHex]];
  //         break;
  //       case "Gray":
  //         range.format.fill.color = "#" + saturationInHex + saturationInHex + saturationInHex;
  //         //range.values = [["#" + saturationInHex + saturationInHex + saturationInHex]];
  //         break;
  //     }
  //   }
  //   return context.sync();
  // }).catch(function(error) {
  //   console.log("Error: " + error);
  //   if (error instanceof OfficeExtension.Error) {
  //     console.log("Debug info: " + JSON.stringify(error.debugInfo));
  //   }
  // });
  Excel.run(function (context) {
    return __awaiter(this, void 0, void 0, function () {
      var range, colorSelected, i, j, tmpRangeValueColorInHex;
      return __generator(this, function (_a) {
        switch (_a.label) {
          case 0:
            range = context.workbook.getSelectedRange();
            range.load(["columnCount", "rowCount", "values"]);
            return [4
            /*yield*/
            , context.sync()];

          case 1:
            _a.sent();

            colorSelected = document.getElementById("color").value;

            switch (colorSelected //REMEMBER TO CHANGE ALL CASES
            ) {
              case "Red":
                for (i = 0; i < range.rowCount; i++) {
                  for (j = 0; j < range.columnCount; j++) {
                    if (typeof range.values[i][j] != "number" || range.values[i][j] < 0 || range.values[i][j] > 255) {
                      //not a number or out of range
                      range.getCell(i, j).format.fill.color = "white";
                      range.getCell(i, j).format.font.color = "black";
                    } else {
                      tmpRangeValueColorInHex = "#" + parseInt(range.values[i][j]).toString(16).padStart(2, "0") + "0000";
                      range.getCell(i, j).format.fill.color = tmpRangeValueColorInHex;

                      if (parseInt(range.values[i][j]) > 128) {
                        range.getCell(i, j).format.font.color = "black";
                      } else {
                        range.getCell(i, j).format.font.color = "white";
                      }
                    }
                  }
                }

                break;

              case "Green":
                for (i = 0; i < range.rowCount; i++) {
                  for (j = 0; j < range.columnCount; j++) {
                    if (typeof range.values[i][j] != "number" || range.values[i][j] < 0 || range.values[i][j] > 255) {
                      //not a number or out of range
                      range.getCell(i, j).format.fill.color = "white";
                      range.getCell(i, j).format.font.color = "black";
                    } else {
                      tmpRangeValueColorInHex = "#00" + parseInt(range.values[i][j]).toString(16).padStart(2, "0") + "00";
                      range.getCell(i, j).format.fill.color = tmpRangeValueColorInHex;

                      if (parseInt(range.values[i][j]) > 128) {
                        range.getCell(i, j).format.font.color = "black";
                      } else {
                        range.getCell(i, j).format.font.color = "white";
                      }
                    }
                  }
                } // range.format.fill.color = "#00" + saturationInHex + "00";
                //range.values = "#00" + saturationInHex + "00";


                break;

              case "Blue":
                for (i = 0; i < range.rowCount; i++) {
                  for (j = 0; j < range.columnCount; j++) {
                    if (typeof range.values[i][j] != "number" || range.values[i][j] < 0 || range.values[i][j] > 255) {
                      //not a number or out of range
                      range.getCell(i, j).format.fill.color = "white";
                      range.getCell(i, j).format.font.color = "black";
                    } else {
                      tmpRangeValueColorInHex = "#0000" + parseInt(range.values[i][j]).toString(16).padStart(2, "0");
                      range.getCell(i, j).format.fill.color = tmpRangeValueColorInHex;

                      if (parseInt(range.values[i][j]) > 128) {
                        range.getCell(i, j).format.font.color = "black";
                      } else {
                        range.getCell(i, j).format.font.color = "white";
                      }
                    }
                  }
                } // range.format.fill.color = "#0000" + saturationInHex;
                //range.values = "#0000" + saturationInHex;


                break;

              case "Gray":
                for (i = 0; i < range.rowCount; i++) {
                  for (j = 0; j < range.columnCount; j++) {
                    if (typeof range.values[i][j] != "number" || range.values[i][j] < 0 || range.values[i][j] > 255) {
                      //not a number or out of range
                      range.getCell(i, j).format.fill.color = "white";
                      range.getCell(i, j).format.font.color = "black";
                    } else {
                      tmpRangeValueColorInHex = parseInt(range.values[i][j]).toString(16).padStart(2, "0");
                      tmpRangeValueColorInHex = "#" + tmpRangeValueColorInHex + tmpRangeValueColorInHex + tmpRangeValueColorInHex;
                      range.getCell(i, j).format.fill.color = tmpRangeValueColorInHex;

                      if (parseInt(range.values[i][j]) > 128) {
                        range.getCell(i, j).format.font.color = "black";
                      } else {
                        range.getCell(i, j).format.font.color = "white";
                      }
                    }
                  }
                } // range.format.fill.color =
                //   "#" + saturationInHex + saturationInHex + saturationInHex;
                //range.values =
                //"#" + saturationInHex + saturationInHex + saturationInHex;


                break;
            }

            return [4
            /*yield*/
            , context.sync()];

          case 2:
            _a.sent();

            return [2
            /*return*/
            ];
        }
      });
    });
  }).catch(function (error) {
    console.log("Error: " + error);

    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
} //Scroll down for easier viewing input list


function scrollDown() {
  Excel.run(function (context) {
    return __awaiter(this, void 0, void 0, function () {
      var scrollDownInput, sheet, range, i, j;
      return __generator(this, function (_a) {
        switch (_a.label) {
          case 0:
            scrollDownInput = document.getElementById("scrollDownInput").value;
            sheet = context.workbook.worksheets.getActiveWorksheet();
            range = context.workbook.getSelectedRange();
            if (!(scrollDownInput == "")) return [3
            /*break*/
            , 2];
            range.getCell(0, 0).values = [["Please input something"]];
            return [4
            /*yield*/
            , context.sync()];

          case 1:
            _a.sent();

            return [3
            /*break*/
            , 9];

          case 2:
            range.load(["rowCount", "columnCount"]); // range.getCell(0,0).values = [[scrollDownInput]];

            return [4
            /*yield*/
            , context.sync()];

          case 3:
            // range.getCell(0,0).values = [[scrollDownInput]];
            _a.sent();

            for (i = 0; i < range.rowCount; i++) {
              for (j = 0; j < range.columnCount; j++) {
                range.getCell(i, j).values = [[scrollDownInput]];
              }
            }

            return [4
            /*yield*/
            , context.sync()];

          case 4:
            _a.sent();

            range = range.getRowsBelow(15);
            range.select();
            return [4
            /*yield*/
            , context.sync()];

          case 5:
            _a.sent();

            range = range.getRowsBelow();
            range.select();
            return [4
            /*yield*/
            , context.sync()];

          case 6:
            _a.sent();

            range = range.getRowsAbove(14);
            range.select();
            return [4
            /*yield*/
            , context.sync()];

          case 7:
            _a.sent();

            range = range.getRowsAbove();
            range.select();
            return [4
            /*yield*/
            , context.sync()];

          case 8:
            _a.sent();

            _a.label = 9;

          case 9:
            focusInputScrolldown();
            return [2
            /*return*/
            ];
        }
      });
    });
  }).catch(function (error) {
    console.log("Error: " + error);

    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function focusInputScrolldown() {
  var input = document.getElementById("scrollDownInput");
  input.select();
}

/***/ })

/******/ });
//# sourceMappingURL=taskpane.js.map
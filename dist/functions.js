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
/******/ 	return __webpack_require__(__webpack_require__.s = "./src/functions/functions.ts");
/******/ })
/************************************************************************/
/******/ ({

/***/ "./src/functions/functions.ts":
/*!************************************!*\
  !*** ./src/functions/functions.ts ***!
  \************************************/
/*! no static exports found */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */

/* global clearInterval, console, setInterval */

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
exports.TyGia = exports.TacGia = exports.add15 = exports.show = exports.logMessage = exports.increment = exports.currentTime = exports.clock = exports.add = void 0;

function add(first, second) {
  return first + second;
}

exports.add = add;
/**
 * Displays the current time once a second.
 * @customfunction
 * @param invocation Custom function handler
 */

function clock(invocation) {
  var timer = setInterval(function () {
    var time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = function () {
    clearInterval(timer);
  };
}

exports.clock = clock;
/**
 * Returns the current time.
 * @returns String with the current time formatted for the current locale.
 */

function currentTime() {
  return new Date().toLocaleTimeString();
}

exports.currentTime = currentTime;
/**
 * Increments a value once a second.
 * @customfunction
 * @param incrementBy Amount to increment
 * @param invocation Custom function handler
 */

function increment(incrementBy, invocation) {
  var result = 0;
  var timer = setInterval(function () {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = function () {
    clearInterval(timer);
  };
}

exports.increment = increment;
/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param message String to write.
 * @returns String to write.
 */

function logMessage(message) {
  console.log(message);
  return message;
}

exports.logMessage = logMessage;
/**
 * Returns the input string.
 * @customfunction
 * @param message String to be returned.
 * @returns Input string.
 */

function show(message) {
  return message;
}

exports.show = show;
/**
 * Returns the input number plus 15.
 * @customfunction
 * @param number The input number.
 * @returns The input number plus 15.
 */

function add15(x) {
  return x + 15;
}

exports.add15 = add15;
/**
 * Returns author's name and ID in 4 cells.
 * @customfunction
 * @param nameAndID - The author's name and ID, ex: "Phan Dinh Tuong 20164582".
 * @returns The name and ID of author separatedly.
 */

function TacGia(nameAndID) {
  if (Number(nameAndID)) //the input is only number
    return [["ERROR: input only number"]];
  var splitted = nameAndID.split(" "); //split input by spaces

  if (splitted.length <= 1) // input length is not enough
    return [["ERROR: input length <=1"]];
  var authorID = splitted.pop(); // pop the last element / author's ID

  if (isNaN(Number(authorID))) //author's ID is not number
    return [["ERROR: input wrong ID (not number)"]];
  var authorName = splitted; //get fullname array

  if (authorName.length == 1) //input only last name
    return [["ERROR: only one in name, please input fullname"]];

  for (var _i = 0, authorName_1 = authorName; _i < authorName_1.length; _i++) {
    var entry = authorName_1[_i]; //check if name array contains number

    if (!isNaN(Number(entry))) return [["ERROR: number inside name"]];
  }

  var lastNameArray = authorName.pop(); //get last name array

  var lastName = lastNameArray.toString(); //last name in string type

  var firstName;

  if (authorName.length == 1) {
    //the remain name is one left / only first name, no middle name
    firstName = authorName[0].toString();
    return [[firstName, ""], [lastName, authorID]]; //this author has no middle name
  } //get first name, middle name


  var firstNameArray = authorName[0]; //get first name in array type

  firstName = firstNameArray.toString(); //first name in string type

  authorName.shift(); //shift to the left / remove first name in array

  var middleName = authorName.join(" "); // get middle name in string

  return [[firstName, middleName], [lastName, authorID]];
}

exports.TacGia = TacGia;
/**
 * Lấy tỷ giá hối đoái ngoại tệ và vnđ theo niêm yết tại pỏtal.vietcombank.com.vn
 * @customfunction
 * @param currency Mã ngoại tệ.
 * @param type Loại tỷ giá.
 * @param date Ngày lấy tỷ giá.
 * @param invocation Invocation for updating cell's value
 */

function TyGia(currency, type, date, invocation) {
  return __awaiter(this, void 0, void 0, function () {
    var url, xhttp;
    return __generator(this, function (_a) {
      // Loại ngoại tệ
      currency = currency.toUpperCase();

      if (currency == "VND") {
        invocation.setResult("1"); // Vietnam Dong

        return [2
        /*return*/
        ];
      }

      if (!(currency == "AUD" || currency == "CAD" || currency == "CHF" || currency == "CNY" || currency == "DKK" || currency == "EUR" || currency == "GBP" || currency == "HKD" || currency == "INR" || currency == "JPY" || currency == "KRW" || currency == "KWD" || currency == "MYR" || currency == "NOK" || currency == "RUB" || currency == "SAR" || currency == "SEK" || currency == "SGD" || currency == "THB" || currency == "USD")) {
        invocation.setResult("Unknown currency");
        return [2
        /*return*/
        ];
      } // Loại ngày lấy tỷ giá


      if (date != "") {//check valid date
        //var excha = Date.parse("01/01/2020");
        //var newDate = excha.toLocaleString();
        //return newDate; //1,577,811,600,000
        //return excha.toString(); // 157781600000
        //return "01/01/2020".toLocaleUpperCase();
      } // Loại tỷ giá mua/bán


      type = type.toUpperCase();
      if (type == "MUA" || type == "BUY") type = "Buy";else if (type == "BÁN" || type == "SELL") type = "Sell";else if (type == "CHUYỂN KHOẢN" || type == "TRANSFER") type = "Transfer";else {
        invocation.setResult("Unknown type");
        return [2
        /*return*/
        ];
      }
      url = "http://127.0.0.1:10010/crawl?currency=" + currency + "&type=" + type + "&date=" + date;
      xhttp = new XMLHttpRequest();
      return [2
      /*return*/
      , new Promise(function (resolve, reject) {
        xhttp.onreadystatechange = function () {
          // invocation.setResult(xhttp.responseText);
          // invocation.setResult(xhttp.responseText);
          if (xhttp.readyState !== 4) return;

          if (xhttp.status == 200) {
            invocation.setResult(xhttp.responseText); // resolve(xhttp.responseText);
          } else {
            reject({
              status: xhttp.status,
              statusText: xhttp.statusText
            }); // invocation.setResult("rejectedffff");
          }
        };

        xhttp.open("GET", url, true);
        xhttp.send();
      })];
    });
  });
}

exports.TyGia = TyGia;
/**
 * Gets the star count for a given Github organization or user and repository.
 * @customfunction
 * @param invocation invo
 */

function getStarCount(invocation) {
  return __awaiter(this, void 0, void 0, function () {
    var url, xhttp;
    return __generator(this, function (_a) {
      url = "http://127.0.0.1:10010/crawl";
      xhttp = new XMLHttpRequest();
      return [2
      /*return*/
      , new Promise(function (resolve, reject) {
        xhttp.onreadystatechange = function () {
          // invocation.setResult(xhttp.responseText);
          // invocation.setResult(xhttp.responseText);
          if (xhttp.readyState !== 4) return;

          if (xhttp.status == 200) {
            invocation.setResult(xhttp.responseText); // resolve(xhttp.responseText);
          } else {
            reject({
              status: xhttp.status,
              statusText: xhttp.statusText
            }); // invocation.setResult("rejectedffff");
          }
        };

        xhttp.open("GET", url, true);
        xhttp.send();
      })];
    });
  });
}
CustomFunctions.associate("ADD", add);
CustomFunctions.associate("CLOCK", clock);
CustomFunctions.associate("INCREMENT", increment);
CustomFunctions.associate("LOG", logMessage);
CustomFunctions.associate("SHOW", show);
CustomFunctions.associate("ADD15", add15);
CustomFunctions.associate("TACGIA", TacGia);
CustomFunctions.associate("TYGIA", TyGia);
CustomFunctions.associate("GETSTARCOUNT", getStarCount);

/***/ })

/******/ });
//# sourceMappingURL=functions.js.map
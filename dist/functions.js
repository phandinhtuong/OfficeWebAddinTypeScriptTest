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

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.TacGia = exports.add15 = exports.show = exports.logMessage = exports.increment = exports.currentTime = exports.clock = exports.add = void 0;

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
 * @param nameAndID The author's name and ID as a string.
 * @returns The name and ID of author separated.
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
    //check if name array contains number
    var entry = authorName_1[_i];
    if (!isNaN(Number(entry))) return [["ERROR: number inside name"]];
  }

  var lastNameArray = authorName.pop(); //get last name array

  var lastName = lastNameArray.toString(); //last name in string type

  if (authorName.length == 1) {
    //the remain name is one left / only first name, no middle name
    var firstName = authorName[0].toString();
    return [[firstName, ''], [lastName, authorID]];
  } //get first name, middle name


  var firstNameArray = authorName[0]; //get first name in array type

  var firstName = firstNameArray.toString(); //first name in string type

  authorName.shift(); //shift to the left / remove first name in array

  var middleName = authorName.join(" "); // get middle name in string

  return [[firstName, middleName], [lastName, authorID]]; //return [['first','second'], ['thirth','fourth']];
}

exports.TacGia = TacGia;
CustomFunctions.associate("ADD", add);
CustomFunctions.associate("CLOCK", clock);
CustomFunctions.associate("INCREMENT", increment);
CustomFunctions.associate("LOG", logMessage);
CustomFunctions.associate("SHOW", show);
CustomFunctions.associate("ADD15", add15);
CustomFunctions.associate("TACGIA", TacGia);

/***/ })

/******/ });
//# sourceMappingURL=functions.js.map
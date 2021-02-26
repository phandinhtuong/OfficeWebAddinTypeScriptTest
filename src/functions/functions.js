"use strict";
/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
/* global clearInterval, console, setInterval */
exports.__esModule = true;
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
    for (var _i = 0, authorName_1 = authorName; _i < authorName_1.length; _i++) { //check if name array contains number
        var entry = authorName_1[_i];
        if (!isNaN(Number(entry)))
            return [["ERROR: number inside name"]];
    }
    var lastNameArray = authorName.pop(); //get last name array
    var lastName = lastNameArray.toString(); //last name in string type
    if (authorName.length == 1) { //the remain name is one left / only first name, no middle name
        var firstName = authorName[0].toString();
        return [[firstName, ''], [lastName, authorID]]; //this author has no middle name
    }
    //get first name, middle name
    var firstNameArray = authorName[0]; //get first name in array type
    var firstName = firstNameArray.toString(); //first name in string type
    authorName.shift(); //shift to the left / remove first name in array
    var middleName = authorName.join(" "); // get middle name in string
    return [[firstName, middleName], [lastName, authorID]];
}
exports.TacGia = TacGia;

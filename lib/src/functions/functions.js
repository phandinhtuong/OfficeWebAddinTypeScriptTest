"use strict";
/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
/* global clearInterval, console, setInterval */
Object.defineProperty(exports, "__esModule", { value: true });
exports.show = exports.logMessage = exports.increment = exports.currentTime = exports.clock = exports.add = void 0;
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
//# sourceMappingURL=functions.js.map
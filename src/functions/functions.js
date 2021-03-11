"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
exports.__esModule = true;
exports.TyGia = exports.TacGia = void 0;
/**
 * Returns author's name and ID in 4 cells.
 * @customfunction
 * @param nameAndID - The author's name and ID, ex: "Phan Dinh Tuong 20164582".
 * @returns The name and ID of author separatedly.
 */
function TacGia(nameAndID) {
    if (Number(nameAndID))
        //the input is only number
        return [["ERROR: input only number"]];
    var splitted = nameAndID.split(" "); //split input by spaces
    if (splitted.length <= 1)
        // input length is not enough
        return [["ERROR: input length <=1"]];
    var authorID = splitted.pop(); // pop the last element / author's ID
    if (isNaN(Number(authorID)))
        //author's ID is not number
        return [["ERROR: input wrong ID (not number)"]];
    var authorName = splitted; //get fullname array
    if (authorName.length == 1)
        //input only last name
        return [["ERROR: only one in name, please input fullname"]];
    for (var _i = 0, authorName_1 = authorName; _i < authorName_1.length; _i++) {
        var entry = authorName_1[_i];
        //check if name array contains number
        if (!isNaN(Number(entry)))
            return [["ERROR: number inside name"]];
    }
    var lastNameArray = authorName.pop(); //get last name array
    var lastName = lastNameArray.toString(); //last name in string type
    var firstName;
    if (authorName.length == 1) {
        //the remain name is one left / only first name, no middle name
        firstName = authorName[0].toString();
        return [
            [firstName, ""],
            [lastName, authorID]
        ]; //this author has no middle name
    }
    //get first name, middle name
    var firstNameArray = authorName[0]; //get first name in array type
    firstName = firstNameArray.toString(); //first name in string type
    authorName.shift(); //shift to the left / remove first name in array
    var middleName = authorName.join(" "); // get middle name in string
    return [
        [firstName, middleName],
        [lastName, authorID]
    ];
}
exports.TacGia = TacGia;
/**
 * Lấy tỷ giá hối đoái ngoại tệ và vnđ theo niêm yết tại portal.vietcombank.com.vn
 * @customfunction
 * @param currency Mã ngoại tệ.
 * @param type Loại tỷ giá.
 * @param date Ngày lấy tỷ giá.
 * @param invocation Invocation for updating cell's value
 */
function TyGia(currency, type, date, invocation) {
    return __awaiter(this, void 0, void 0, function () {
        var moment, exchangeDate, today, url, xhttp;
        return __generator(this, function (_a) {
            // Loại ngoại tệ
            currency = currency.toUpperCase();
            if (currency == "VND") {
                invocation.setResult("1"); // Vietnam Dong
                return [2 /*return*/];
            }
            if (!(currency == "AUD" ||
                currency == "CAD" ||
                currency == "CHF" ||
                currency == "CNY" ||
                currency == "DKK" ||
                currency == "EUR" ||
                currency == "GBP" ||
                currency == "HKD" ||
                currency == "INR" ||
                currency == "JPY" ||
                currency == "KRW" ||
                currency == "KWD" ||
                currency == "MYR" ||
                currency == "NOK" ||
                currency == "RUB" ||
                currency == "SAR" ||
                currency == "SEK" ||
                currency == "SGD" ||
                currency == "THB" ||
                currency == "USD")) {
                invocation.setResult("Unknown currency");
                return [2 /*return*/];
            }
            // Loại ngày lấy tỷ giá
            if (date != "") {
                moment = require('moment');
                exchangeDate = moment(date, "DD/MM/YYYY", true);
                today = moment();
                if (isNaN(exchangeDate)) {
                    invocation.setResult("Invalid date");
                    return [2 /*return*/];
                }
                else if (exchangeDate > today) {
                    invocation.setResult('Out of date');
                    return [2 /*return*/];
                }
                else {
                    date = exchangeDate.format('DD/MM/YYYY');
                }
            }
            // Loại tỷ giá mua/bán
            type = type.toUpperCase();
            if (type == "MUA" || type == "BUY")
                type = "Buy";
            else if (type == "BÁN" || type == "SELL")
                type = "Sell";
            else if (type == "CHUYỂN KHOẢN" || type == "TRANSFER")
                type = "Transfer";
            else {
                invocation.setResult("Unknown type");
                return [2 /*return*/];
            }
            url = "http://127.0.0.1:10010/crawlTyGia?currency=" + currency + "&type=" + type + "&date=" + date;
            xhttp = new XMLHttpRequest();
            return [2 /*return*/, new Promise(function (resolve, reject) {
                    xhttp.onreadystatechange = function () {
                        if (xhttp.readyState !== 4)
                            return;
                        if (xhttp.status == 200) {
                            // resolve(JSON.parse(xhttp.responseText).result);
                            invocation.setResult(JSON.parse(xhttp.responseText).result);
                        }
                        else {
                            reject({
                                status: xhttp.status,
                                statusText: xhttp.statusText
                            });
                            invocation.setResult("Request was rejected");
                        }
                    };
                    xhttp.open("GET", url, true);
                    xhttp.send();
                })];
        });
    });
}
exports.TyGia = TyGia;

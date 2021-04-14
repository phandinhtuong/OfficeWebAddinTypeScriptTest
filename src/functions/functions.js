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
exports.decimalToText = exports.decimalToSpeech = exports.numberToText = exports.numberToSpeech = exports.QRCode = exports.StringFontFormatter = exports.StringCellFormatter = exports.KipThi = exports.TyGia = exports.TacGia = void 0;
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
                moment = require("moment");
                exchangeDate = moment(date, "DD/MM/YYYY", true);
                today = moment();
                if (isNaN(exchangeDate)) {
                    invocation.setResult("Invalid date");
                    return [2 /*return*/];
                }
                else if (exchangeDate > today) {
                    invocation.setResult("Out of date");
                    return [2 /*return*/];
                }
                else {
                    date = exchangeDate.format("DDMMYYYY");
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
            url = "https://mycelwebapi.vimoitruong.xyz/crawlTyGia?currency=" + currency + "&type=" + type + "&date=" + date;
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
/**
 * Returns starting time of the exam
 * @customfunction
 * @param kip Kip thi index
 * @returns Starting time of the exam
 */
function KipThi(kip) {
    var startingTime;
    switch (kip) {
        case 1:
            startingTime = "07:00";
            break;
        case 2:
            startingTime = "09:30";
            break;
        case 3:
            startingTime = "12:30";
            break;
        case 4:
            startingTime = "15:00";
            break;
        default:
            startingTime = "Invalid";
            break;
    }
    return startingTime;
}
exports.KipThi = KipThi;
/**
 * Return input text and change background color.
 * @customfunction
 * @param text Input text.
 * @param cellBackgroundColor Background color of cell to be applied.
 * @param invocation Invocation object to get current cell.
 * @requiresAddress
 * @returns Input text and change background color.
 */
function StringCellFormatter(text, cellBackgroundColor, invocation) {
    var address = invocation.address; //get address of the invocation / current cell, eg: Sheet1!A4
    var addressWithoutSheet = address.split("!")[1]; //split to get address without sheet
    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var range = sheet.getRange(addressWithoutSheet);
        range.select();
        range.format.fill.color = cellBackgroundColor;
        return context.sync();
    })["catch"](function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
    return text;
}
exports.StringCellFormatter = StringCellFormatter;
/**
 * Return the input text and change the font type, font size, font color.
 * @customfunction
 * @param text Input text.
 * @param fontName Font name.
 * @param fontSize Font size.
 * @param fontColor Font color.
 * @param invocation Invocation object to get current cell.
 * @requiresAddress
 * @returns Input text.
 */
function StringFontFormatter(text, fontName, fontSize, fontColor, invocation) {
    var address = invocation.address; //get address of the invocation / current cell, eg: Sheet1!A4
    var addressWithoutSheet = address.split("!")[1]; //split to get address without sheet
    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        var range = sheet.getRange(addressWithoutSheet);
        range.select();
        range.format.font.name = fontName; //eg 'Times New Roman'
        range.format.font.color = fontColor;
        range.format.font.size = fontSize;
        return context.sync();
    })["catch"](function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
    return text;
}
exports.StringFontFormatter = StringFontFormatter;
/**
 * Generate QR code from text.
 * @customfunction
 * @param text Context of QR code.
 * @param ShapeName Name of the shape will be filled with QR code.
 * @param invocation Invocation object to get current cell.
 * @requiresAddress
 */
function QRCode(text, ShapeName, invocation) {
    Excel.run(function (context) {
        // request to WebAPI / must start WebAPI first
        // const url = "http://127.0.0.1:10010/qrcode?text=" + text;
        var url = "https://mycelwebapi.vimoitruong.xyz/qrcode?text=" + text;
        var xhttp = new XMLHttpRequest();
        return new Promise(function (resolve, reject) {
            xhttp.onreadystatechange = function () {
                return __awaiter(this, void 0, void 0, function () {
                    var result, address, addressWithoutSheet, sheet, range, shapes, myShape;
                    return __generator(this, function (_a) {
                        switch (_a.label) {
                            case 0:
                                if (xhttp.readyState !== 4)
                                    return [2 /*return*/];
                                if (!(xhttp.status == 200)) return [3 /*break*/, 3];
                                result = JSON.parse(xhttp.responseText).result;
                                result = result.substring(22); //get substring from index 22 to avoid 'data:image/png;base64,'
                                address = invocation.address;
                                addressWithoutSheet = address.split("!")[1];
                                sheet = context.workbook.worksheets.getActiveWorksheet();
                                range = sheet.getRange(addressWithoutSheet);
                                range.select();
                                range.load(["left", "top", "height", "width"]); // load left, top, height, width of the cell to place QR code image
                                return [4 /*yield*/, context.sync()];
                            case 1:
                                _a.sent(); //always call sync after loading
                                shapes = context.workbook.worksheets.getActiveWorksheet().shapes;
                                myShape = sheet.shapes.addImage(result);
                                myShape.name = ShapeName;
                                //placement: two cell: shape is moved with the cell.
                                //left, top, height, width: properties of the cell that the shape follows
                                myShape.set({
                                    placement: "TwoCell",
                                    left: range.left,
                                    top: range.top,
                                    height: range.height,
                                    width: range.width
                                });
                                // }else{
                                // }
                                // var j = shapes.getCount();
                                // // var stemp[j]
                                // context.sync();
                                // console.log("Sss");
                                // console.log(j.value);
                                // var i, stemp;
                                // for (i = 0; i < j.value; i++) {
                                //   stemp = shapes.getItemAt(i);
                                //   stemp.load("name");
                                //   context.sync();
                                //   console.log(i);
                                //   if (stemp.name == "newImage") {
                                //     context.sync();
                                //     console.log("yes");
                                //   } else {
                                //     console.log("no");
                                //   }
                                // }
                                //var image = sheet.shapes.addImage(result);
                                //TODO: find shape by name and fill
                                //console.log(result);
                                //image.name = "image";
                                //streamingInvocation.setResult(text);
                                //return JSON.parse(xhttp.responseText).result;
                                return [4 /*yield*/, context.sync()];
                            case 2:
                                // }else{
                                // }
                                // var j = shapes.getCount();
                                // // var stemp[j]
                                // context.sync();
                                // console.log("Sss");
                                // console.log(j.value);
                                // var i, stemp;
                                // for (i = 0; i < j.value; i++) {
                                //   stemp = shapes.getItemAt(i);
                                //   stemp.load("name");
                                //   context.sync();
                                //   console.log(i);
                                //   if (stemp.name == "newImage") {
                                //     context.sync();
                                //     console.log("yes");
                                //   } else {
                                //     console.log("no");
                                //   }
                                // }
                                //var image = sheet.shapes.addImage(result);
                                //TODO: find shape by name and fill
                                //console.log(result);
                                //image.name = "image";
                                //streamingInvocation.setResult(text);
                                //return JSON.parse(xhttp.responseText).result;
                                _a.sent();
                                return [3 /*break*/, 4];
                            case 3:
                                reject({
                                    status: xhttp.status,
                                    statusText: xhttp.statusText
                                });
                                console.log("Request was rejected");
                                _a.label = 4;
                            case 4: return [2 /*return*/];
                        }
                    });
                });
            };
            xhttp.open("GET", url, true);
            xhttp.send();
        });
    })["catch"](function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
    return text;
}
exports.QRCode = QRCode;
/**
 * Speak the input number.
 * @customfunction
 * @param number Number to be spoken.
 * @returns Number in text.
 */
function numberToSpeech(number) {
    console.log(typeof number);
    console.log(number);
    var chuSo = [" không", " một", " hai", " ba", " bốn", " năm", " sáu", " bẩy", " tám", " chín"];
    var audioChuSo = [
        new Audio("../../sound/0.wav"),
        new Audio("../../sound/1.wav"),
        new Audio("../../sound/2.wav"),
        new Audio("../../sound/3.wav"),
        new Audio("../../sound/4.wav"),
        new Audio("../../sound/5.wav"),
        new Audio("../../sound/6.wav"),
        new Audio("../../sound/7.wav"),
        new Audio("../../sound/8.wav"),
        new Audio("../../sound/9.wav")
    ];
    var tien = [" GH", " nghìn", " triệu", " tỷ", " nghìn tỷ", " triệu tỷ"];
    var audioTien = [
        new Audio("../../sound/GolfHit.wav"),
        new Audio("../../sound/nghin.wav"),
        new Audio("../../sound/trieu.wav"),
        new Audio("../../sound/ty.wav"),
        new Audio("../../sound/nghinTy.wav"),
        new Audio("../../sound/trieuTy.wav")
    ];
    var biggestNumber = 8999999999999999;
    var lan, i, so, ketqua = "", tmp = "";
    var vitri = new Array(6);
    var sounds = new Array();
    if (typeof number != "number") {
        console.log("check number: not number");
    }
    else {
        console.log("check number: number");
        if (number == 0) {
            console.log("number == 0");
            sounds.push(audioChuSo[0]);
            sounds.push(audioTien[0]);
            ketqua += chuSo[0];
        }
        if (number > 0) {
            console.log("number > 0");
            so = number;
        }
        else {
            console.log("number <0");
            so = -number;
        }
        if (number > biggestNumber) {
            return "";
        }
        vitri[5] = Math.floor(so / 1000000000000000);
        console.log("vitri[5] = " + vitri[5]);
        so -= vitri[5] * 1000000000000000;
        vitri[4] = Math.floor(so / 1000000000000);
        console.log("vitri[4] = " + vitri[4]);
        so -= vitri[4] * 1000000000000;
        vitri[3] = Math.floor(so / 1000000000);
        console.log("vitri[3] = " + vitri[3]);
        so -= vitri[3] * 1000000000;
        vitri[2] = Math.floor(so / 1000000);
        console.log("vitri[2] = " + vitri[2]);
        vitri[1] = Math.floor((so % 1000000) / 1000);
        console.log("vitri[1] = " + vitri[1]);
        vitri[0] = so % 1000;
        console.log("vitri[0] = " + vitri[0]);
        if (vitri[5] > 0) {
            lan = 5;
        }
        else if (vitri[4] > 0) {
            lan = 4;
        }
        else if (vitri[3] > 0) {
            lan = 3;
        }
        else if (vitri[2] > 0) {
            lan = 2;
        }
        else if (vitri[1] > 0) {
            lan = 1;
        }
        else {
            lan = 0;
        }
        console.log("lan = " + lan);
        for (i = lan; i >= 0; i--) {
            console.log("i in for = " + i);
            tmp = docSo3ChuSo(vitri[i]);
            console.log("tmp = " + tmp);
            ketqua += tmp;
            console.log("ketqua inside for loop lan= " + ketqua);
            if (vitri[i] != 0 || i == 0) {
                if (i == 0 && ketqua.substring(ketqua.length - 1) == ",")
                    ketqua = ketqua.substring(0, ketqua.length - 1);
                ketqua += tien[i];
                sounds.push(audioTien[i]);
            }
            console.log("ketqua before adding , = " + ketqua);
            if (i > 0 && tmp != "")
                ketqua += ",";
            console.log("ketqua = after adding , = " + ketqua);
        }
        //console.log("ketqua.substring(ketqua.length - 1) = " + ketqua.substring(ketqua.length - 1));
        if (ketqua.substring(ketqua.length - 1) == ",")
            ketqua = ketqua.substring(0, ketqua.length - 1);
        ketqua = ketqua.trim();
        if (number < 0) {
            ketqua = "âm " + ketqua;
            sounds.unshift(new Audio("../../sound/am.wav"));
        }
        console.log("ketqua after add - = " + ketqua);
        // var audio = new Audio("../../sound/0.wav");
        // var audio2 = new Audio("../../sound/1.wav");
        // sounds.push(audio);
        // sounds.push(audio2);
        // sounds.push(audio);
        var soundIndex = -1;
        //console.log("sounds.length outside func = " + sounds.length);
        playSnd();
        return ketqua.substring(0, 1).toUpperCase() + ketqua.substring(1);
    }
    function playSnd() {
        //console.log("sounds.length inside func = " + sounds.length);
        soundIndex++;
        if (soundIndex == sounds.length) {
            return;
        }
        sounds[soundIndex].addEventListener("ended", playSnd);
        sounds[soundIndex].play();
    }
    function docSo3ChuSo(baso) {
        console.log("baso = " + baso);
        var tram, chuc, donvi;
        var ketQua = "";
        tram = Math.floor(baso / 100);
        console.log("tram = " + tram);
        chuc = Math.floor((baso % 100) / 10);
        console.log("chuc = " + chuc);
        donvi = baso % 10;
        console.log("donvi = " + donvi);
        if (tram == 0 && chuc == 0 && donvi == 0) {
            return "";
        }
        if (ketqua.substring(ketqua.length - 1) == "," || tram != 0) {
            ketQua += chuSo[tram] + " trăm";
            sounds.push(audioChuSo[tram]);
            sounds.push(new Audio("../../sound/tram.wav"));
            if (chuc == 0 && donvi != 0) {
                ketQua += " linh";
                sounds.push(new Audio("../../sound/linh.wav"));
            }
        }
        if (chuc != 0 && chuc != 1) {
            ketQua += chuSo[chuc] + " mươi";
            sounds.push(audioChuSo[chuc]);
            sounds.push(new Audio("../../sound/muoi.wav"));
            if (chuc == 0 && donvi != 0) {
                ketQua += " linh";
                sounds.push(new Audio("../../sound/linh.wav"));
            }
        }
        if (chuc == 1) {
            ketQua += " mười";
            sounds.push(new Audio("../../sound/10.wav"));
        }
        switch (donvi) {
            case 1:
                if (chuc != 0 && chuc != 1) {
                    ketQua += " mốt";
                    sounds.push(new Audio("../../sound/mot.wav"));
                }
                else {
                    ketQua += chuSo[donvi];
                    sounds.push(audioChuSo[donvi]);
                }
                break;
            case 5:
                if (chuc == 0) {
                    ketQua += chuSo[donvi];
                    sounds.push(audioChuSo[donvi]);
                }
                else {
                    ketQua += " lăm";
                    sounds.push(new Audio("../../sound/lam.wav"));
                }
                break;
            default:
                if (donvi != 0) {
                    ketQua += chuSo[donvi];
                    sounds.push(audioChuSo[donvi]);
                }
                break;
        }
        return ketQua;
    }
}
exports.numberToSpeech = numberToSpeech;
/**
 * Number to text in Vietnamese.
 * @customfunction
 * @param number Number to text.
 * @returns Number in text.
 */
function numberToText(number) {
    console.log(typeof number);
    console.log(number);
    var chuSo = [" không", " một", " hai", " ba", " bốn", " năm", " sáu", " bẩy", " tám", " chín"];
    var tien = [" GH", " nghìn", " triệu", " tỷ", " nghìn tỷ", " triệu tỷ"];
    var biggestNumber = 8999999999999999;
    var lan, i, so, ketqua = "", tmp = "";
    var vitri = new Array(6);
    if (typeof number != "number") {
        console.log("check number: not number");
    }
    else {
        console.log("check number: number");
        if (number == 0) {
            console.log("number == 0");
            ketqua += chuSo[0];
        }
        if (number > 0) {
            console.log("number > 0");
            so = number;
        }
        else {
            console.log("number <0");
            so = -number;
        }
        if (number > biggestNumber) {
            return "";
        }
        vitri[5] = Math.floor(so / 1000000000000000);
        console.log("vitri[5] = " + vitri[5]);
        so -= vitri[5] * 1000000000000000;
        vitri[4] = Math.floor(so / 1000000000000);
        console.log("vitri[4] = " + vitri[4]);
        so -= vitri[4] * 1000000000000;
        vitri[3] = Math.floor(so / 1000000000);
        console.log("vitri[3] = " + vitri[3]);
        so -= vitri[3] * 1000000000;
        vitri[2] = Math.floor(so / 1000000);
        console.log("vitri[2] = " + vitri[2]);
        vitri[1] = Math.floor((so % 1000000) / 1000);
        console.log("vitri[1] = " + vitri[1]);
        vitri[0] = so % 1000;
        console.log("vitri[0] = " + vitri[0]);
        if (vitri[5] > 0) {
            lan = 5;
        }
        else if (vitri[4] > 0) {
            lan = 4;
        }
        else if (vitri[3] > 0) {
            lan = 3;
        }
        else if (vitri[2] > 0) {
            lan = 2;
        }
        else if (vitri[1] > 0) {
            lan = 1;
        }
        else {
            lan = 0;
        }
        console.log("lan = " + lan);
        for (i = lan; i >= 0; i--) {
            console.log("i in for = " + i);
            tmp = docSo3ChuSo(vitri[i]);
            console.log("tmp = " + tmp);
            ketqua += tmp;
            console.log("ketqua inside for loop lan= " + ketqua);
            if (vitri[i] != 0 || i == 0) {
                if (i == 0 && ketqua.substring(ketqua.length - 1) == ",")
                    ketqua = ketqua.substring(0, ketqua.length - 1);
                ketqua += tien[i];
            }
            console.log("ketqua before adding , = " + ketqua);
            if (i > 0 && tmp != "")
                ketqua += ",";
            console.log("ketqua = after adding , = " + ketqua);
        }
        //console.log("ketqua.substring(ketqua.length - 1) = " + ketqua.substring(ketqua.length - 1));
        if (ketqua.substring(ketqua.length - 1) == ",")
            ketqua = ketqua.substring(0, ketqua.length - 1);
        ketqua = ketqua.trim();
        if (number < 0) {
            ketqua = "âm " + ketqua;
        }
        console.log("ketqua after add - = " + ketqua);
        return ketqua.substring(0, 1).toUpperCase() + ketqua.substring(1);
    }
    function docSo3ChuSo(baso) {
        console.log("baso = " + baso);
        var tram, chuc, donvi;
        var ketQua = "";
        tram = Math.floor(baso / 100);
        console.log("tram = " + tram);
        chuc = Math.floor((baso % 100) / 10);
        console.log("chuc = " + chuc);
        donvi = baso % 10;
        console.log("donvi = " + donvi);
        if (tram == 0 && chuc == 0 && donvi == 0) {
            return "";
        }
        if (ketqua.substring(ketqua.length - 1) == "," || tram != 0) {
            ketQua += chuSo[tram] + " trăm";
            if (chuc == 0 && donvi != 0) {
                ketQua += " linh";
            }
        }
        if (chuc != 0 && chuc != 1) {
            ketQua += chuSo[chuc] + " mươi";
            if (chuc == 0 && donvi != 0) {
                ketQua += " linh";
            }
        }
        if (chuc == 1) {
            ketQua += " mười";
        }
        switch (donvi) {
            case 1:
                if (chuc != 0 && chuc != 1) {
                    ketQua += " mốt";
                }
                else {
                    ketQua += chuSo[donvi];
                }
                break;
            case 5:
                if (chuc == 0) {
                    ketQua += chuSo[donvi];
                }
                else {
                    ketQua += " lăm";
                }
                break;
            default:
                if (donvi != 0) {
                    ketQua += chuSo[donvi];
                }
                break;
        }
        return ketQua;
    }
}
exports.numberToText = numberToText;
/**
 * Decimal number to speech.
 * @customfunction
 * @param number Number to be spoken.
 * @returns Decimal number in text.
 */
function decimalToSpeech(number) {
    console.log(typeof number);
    console.log(number);
    var chuSo = [" không", " một", " hai", " ba", " bốn", " năm", " sáu", " bẩy", " tám", " chín"];
    var audioChuSo = [
        new Audio("../../sound/0.wav"),
        new Audio("../../sound/1.wav"),
        new Audio("../../sound/2.wav"),
        new Audio("../../sound/3.wav"),
        new Audio("../../sound/4.wav"),
        new Audio("../../sound/5.wav"),
        new Audio("../../sound/6.wav"),
        new Audio("../../sound/7.wav"),
        new Audio("../../sound/8.wav"),
        new Audio("../../sound/9.wav")
    ];
    var tien = [" GH", " nghìn", " triệu", " tỷ", " nghìn tỷ", " triệu tỷ"];
    var audioTien = [
        new Audio("../../sound/GolfHit.wav"),
        new Audio("../../sound/nghin.wav"),
        new Audio("../../sound/trieu.wav"),
        new Audio("../../sound/ty.wav"),
        new Audio("../../sound/nghinTy.wav"),
        new Audio("../../sound/trieuTy.wav")
    ];
    var biggestNumber = 8999999999999999;
    var lan, i, so, ketqua = "", tmp = "";
    var vitri = new Array(6);
    var sounds = new Array();
    if (typeof number != "number") {
        console.log("check number: not number");
    }
    else {
        console.log("check number: number");
        if (number == 0) {
            console.log("number == 0");
            sounds.push(audioChuSo[0]);
            sounds.push(audioTien[0]);
            ketqua += chuSo[0];
        }
        if (number > 0) {
            console.log("number > 0");
            so = number;
        }
        else {
            console.log("number <0");
            so = -number;
        }
        if (number > biggestNumber) {
            return "";
        }
        var decimalPart = so;
        decimalPart = decimalPart - Math.floor(decimalPart);
        console.log("decimalPart be4 = " + decimalPart);
        decimalPart = Math.round((decimalPart + Number.EPSILON) * 10000) / 10000;
        console.log("decimalPart aft = " + decimalPart);
        var decimalPartInString = decimalPart.toString();
        decimalPartInString = decimalPartInString.substring(2);
        console.log("decimalPartInString = " + decimalPartInString);
        // decimalPart = parseInt(decimalPartInString);
        // console.log("decimalPart after transform = "+decimalPart);
        so = Math.floor(so); // get integer part
        vitri[5] = Math.floor(so / 1000000000000000);
        console.log("vitri[5] = " + vitri[5]);
        so -= vitri[5] * 1000000000000000;
        vitri[4] = Math.floor(so / 1000000000000);
        console.log("vitri[4] = " + vitri[4]);
        so -= vitri[4] * 1000000000000;
        vitri[3] = Math.floor(so / 1000000000);
        console.log("vitri[3] = " + vitri[3]);
        so -= vitri[3] * 1000000000;
        vitri[2] = Math.floor(so / 1000000);
        console.log("vitri[2] = " + vitri[2]);
        vitri[1] = Math.floor((so % 1000000) / 1000);
        console.log("vitri[1] = " + vitri[1]);
        vitri[0] = so % 1000;
        console.log("vitri[0] = " + vitri[0]);
        if (vitri[5] > 0) {
            lan = 5;
        }
        else if (vitri[4] > 0) {
            lan = 4;
        }
        else if (vitri[3] > 0) {
            lan = 3;
        }
        else if (vitri[2] > 0) {
            lan = 2;
        }
        else if (vitri[1] > 0) {
            lan = 1;
        }
        else {
            lan = 0;
        }
        console.log("lan = " + lan);
        for (i = lan; i >= 0; i--) {
            console.log("i in for = " + i);
            tmp = docSo3ChuSo(vitri[i]);
            console.log("tmp = " + tmp);
            ketqua += tmp;
            console.log("ketqua inside for loop lan= " + ketqua);
            if (vitri[i] != 0 || i == 0) {
                if (i == 0 && ketqua.substring(ketqua.length - 1) == ",")
                    ketqua = ketqua.substring(0, ketqua.length - 1);
                ketqua += tien[i];
                sounds.push(audioTien[i]);
            }
            console.log("ketqua before adding , = " + ketqua);
            if (i > 0 && tmp != "")
                ketqua += ",";
            console.log("ketqua = after adding , = " + ketqua);
        }
        //console.log("ketqua.substring(ketqua.length - 1) = " + ketqua.substring(ketqua.length - 1));
        if (ketqua.substring(ketqua.length - 1) == ",")
            ketqua = ketqua.substring(0, ketqua.length - 1);
        ketqua = ketqua.trim();
        if (number < 0) {
            ketqua = "âm " + ketqua;
            sounds.unshift(new Audio("../../sound/am.wav"));
        }
        console.log("ketqua after add - = " + ketqua);
        // var audio = new Audio("../../sound/0.wav");
        // var audio2 = new Audio("../../sound/1.wav");
        //add decimal part
        ketqua += " phẩy";
        sounds.push(new Audio("../../sound/phay.wav"));
        ketqua += docSo(decimalPartInString);
        // sounds.push(audio);
        // sounds.push(audio2);
        // sounds.push(audio);
        var soundIndex = -1;
        //console.log("sounds.length outside func = " + sounds.length);
        playSnd();
        return ketqua.substring(0, 1).toUpperCase() + ketqua.substring(1);
    }
    function docSo(num) {
        var ketQua = "";
        var i;
        console.log("num.length = " + num.length);
        for (i = 0; i < num.length; i++) {
            console.log(num.substr(i, 1));
            switch (num.substr(i, 1)) {
                case "0":
                    ketQua += " không";
                    sounds.push(audioChuSo[0]);
                    break;
                case "1":
                    ketQua += " một";
                    sounds.push(audioChuSo[1]);
                    break;
                case "2":
                    ketQua += " hai";
                    sounds.push(audioChuSo[2]);
                    break;
                case "3":
                    ketQua += " ba";
                    sounds.push(audioChuSo[3]);
                    break;
                case "4":
                    ketQua += " bốn";
                    sounds.push(audioChuSo[4]);
                    break;
                case "5":
                    ketQua += " năm";
                    sounds.push(audioChuSo[5]);
                    break;
                case "6":
                    ketQua += " sáu";
                    sounds.push(audioChuSo[6]);
                    break;
                case "7":
                    ketQua += " bẩy";
                    sounds.push(audioChuSo[7]);
                    break;
                case "8":
                    ketQua += " tám";
                    sounds.push(audioChuSo[8]);
                    break;
                case "9":
                    ketQua += " chín";
                    sounds.push(audioChuSo[9]);
                    break;
                default: break;
            }
        }
        return ketQua;
    }
    function playSnd() {
        //console.log("sounds.length inside func = " + sounds.length);
        soundIndex++;
        if (soundIndex == sounds.length) {
            return;
        }
        sounds[soundIndex].addEventListener("ended", playSnd);
        sounds[soundIndex].play();
    }
    function docSo3ChuSo(baso) {
        console.log("baso = " + baso);
        var tram, chuc, donvi;
        var ketQua = "";
        tram = Math.floor(baso / 100);
        console.log("tram = " + tram);
        chuc = Math.floor((baso % 100) / 10);
        console.log("chuc = " + chuc);
        donvi = baso % 10;
        console.log("donvi = " + donvi);
        if (tram == 0 && chuc == 0 && donvi == 0) {
            return "";
        }
        if (ketqua.substring(ketqua.length - 1) == "," || tram != 0) {
            ketQua += chuSo[tram] + " trăm";
            sounds.push(audioChuSo[tram]);
            sounds.push(new Audio("../../sound/tram.wav"));
            if (chuc == 0 && donvi != 0) {
                ketQua += " linh";
                sounds.push(new Audio("../../sound/linh.wav"));
            }
        }
        if (chuc != 0 && chuc != 1) {
            ketQua += chuSo[chuc] + " mươi";
            sounds.push(audioChuSo[chuc]);
            sounds.push(new Audio("../../sound/muoi.wav"));
            if (chuc == 0 && donvi != 0) {
                ketQua += " linh";
                sounds.push(new Audio("../../sound/linh.wav"));
            }
        }
        if (chuc == 1) {
            ketQua += " mười";
            sounds.push(new Audio("../../sound/10.wav"));
        }
        switch (donvi) {
            case 1:
                if (chuc != 0 && chuc != 1) {
                    ketQua += " mốt";
                    sounds.push(new Audio("../../sound/mot.wav"));
                }
                else {
                    ketQua += chuSo[donvi];
                    sounds.push(audioChuSo[donvi]);
                }
                break;
            case 5:
                if (chuc == 0) {
                    ketQua += chuSo[donvi];
                    sounds.push(audioChuSo[donvi]);
                }
                else {
                    ketQua += " lăm";
                    sounds.push(new Audio("../../sound/lam.wav"));
                }
                break;
            default:
                if (donvi != 0) {
                    ketQua += chuSo[donvi];
                    sounds.push(audioChuSo[donvi]);
                }
                break;
        }
        return ketQua;
    }
}
exports.decimalToSpeech = decimalToSpeech;
/**
 * Decimal number to text in Vietnamese.
 * @customfunction
 * @param number Number to text.
 * @returns Decimal number in text.
 */
function decimalToText(number) {
    console.log(typeof number);
    console.log(number);
    var chuSo = [" không", " một", " hai", " ba", " bốn", " năm", " sáu", " bẩy", " tám", " chín"];
    var tien = [" GH", " nghìn", " triệu", " tỷ", " nghìn tỷ", " triệu tỷ"];
    var biggestNumber = 8999999999999999;
    var lan, i, so, ketqua = "", tmp = "";
    var vitri = new Array(6);
    if (typeof number != "number") {
        console.log("check number: not number");
    }
    else {
        console.log("check number: number");
        if (number == 0) {
            console.log("number == 0");
            ketqua += chuSo[0];
        }
        if (number > 0) {
            console.log("number > 0");
            so = number;
        }
        else {
            console.log("number <0");
            so = -number;
        }
        if (number > biggestNumber) {
            return "";
        }
        var decimalPart = so;
        decimalPart = decimalPart - Math.floor(decimalPart);
        console.log("decimalPart be4 = " + decimalPart);
        decimalPart = Math.round((decimalPart + Number.EPSILON) * 10000) / 10000;
        console.log("decimalPart aft = " + decimalPart);
        var decimalPartInString = decimalPart.toString();
        decimalPartInString = decimalPartInString.substring(2);
        console.log("decimalPartInString = " + decimalPartInString);
        // decimalPart = parseInt(decimalPartInString);
        // console.log("decimalPart after transform = "+decimalPart);
        so = Math.floor(so); // get integer part
        console.log("so = " + so);
        vitri[5] = Math.floor(so / 1000000000000000);
        console.log("vitri[5] = " + vitri[5]);
        so -= vitri[5] * 1000000000000000;
        vitri[4] = Math.floor(so / 1000000000000);
        console.log("vitri[4] = " + vitri[4]);
        so -= vitri[4] * 1000000000000;
        vitri[3] = Math.floor(so / 1000000000);
        console.log("vitri[3] = " + vitri[3]);
        so -= vitri[3] * 1000000000;
        vitri[2] = Math.floor(so / 1000000);
        console.log("vitri[2] = " + vitri[2]);
        vitri[1] = Math.floor((so % 1000000) / 1000);
        console.log("vitri[1] = " + vitri[1]);
        vitri[0] = so % 1000;
        console.log("vitri[0] = " + vitri[0]);
        if (vitri[5] > 0) {
            lan = 5;
        }
        else if (vitri[4] > 0) {
            lan = 4;
        }
        else if (vitri[3] > 0) {
            lan = 3;
        }
        else if (vitri[2] > 0) {
            lan = 2;
        }
        else if (vitri[1] > 0) {
            lan = 1;
        }
        else {
            lan = 0;
        }
        console.log("lan = " + lan);
        for (i = lan; i >= 0; i--) {
            console.log("i in for = " + i);
            tmp = docSo3ChuSo(vitri[i]);
            console.log("tmp = " + tmp);
            ketqua += tmp;
            console.log("ketqua inside for loop lan= " + ketqua);
            if (vitri[i] != 0 && i != 0) {
                // if (i == 0 && ketqua.substring(ketqua.length - 1) == ",") ketqua = ketqua.substring(0, ketqua.length - 1);
                ketqua += tien[i];
            }
            console.log("ketqua before adding , = " + ketqua);
            if (i > 0 && tmp != "")
                ketqua += ",";
            console.log("ketqua = after adding , = " + ketqua);
        }
        //console.log("ketqua.substring(ketqua.length - 1) = " + ketqua.substring(ketqua.length - 1));
        if (ketqua.substring(ketqua.length - 1) == ",")
            ketqua = ketqua.substring(0, ketqua.length - 1);
        ketqua = ketqua.trim();
        if (number < 0) {
            ketqua = "âm " + ketqua;
        }
        console.log("ketqua after add - = " + ketqua);
        //add decimal part
        ketqua += " phẩy";
        ketqua += docSo(decimalPartInString);
        return ketqua.substring(0, 1).toUpperCase() + ketqua.substring(1);
    }
    function docSo(num) {
        var ketQua = "";
        var i;
        console.log("num.length = " + num.length);
        for (i = 0; i < num.length; i++) {
            console.log(num.substr(i, 1));
            switch (num.substr(i, 1)) {
                case "0":
                    ketQua += " không";
                    break;
                case "1":
                    ketQua += " một";
                    break;
                case "2":
                    ketQua += " hai";
                    break;
                case "3":
                    ketQua += " ba";
                    break;
                case "4":
                    ketQua += " bốn";
                    break;
                case "5":
                    ketQua += " năm";
                    break;
                case "6":
                    ketQua += " sáu";
                    break;
                case "7":
                    ketQua += " bẩy";
                    break;
                case "8":
                    ketQua += " tám";
                    break;
                case "9":
                    ketQua += " chín";
                    break;
                default: break;
            }
        }
        return ketQua;
    }
    function docSo3ChuSo(baso) {
        console.log("baso = " + baso);
        var tram, chuc, donvi;
        var ketQua = "";
        tram = Math.floor(baso / 100);
        console.log("tram = " + tram);
        chuc = Math.floor((baso % 100) / 10);
        console.log("chuc = " + chuc);
        donvi = baso % 10;
        console.log("donvi = " + donvi);
        if (tram == 0 && chuc == 0 && donvi == 0) {
            return "";
        }
        if (ketqua.substring(ketqua.length - 1) == "," || tram != 0) {
            ketQua += chuSo[tram] + " trăm";
            if (chuc == 0 && donvi != 0) {
                ketQua += " linh";
            }
        }
        if (chuc != 0 && chuc != 1) {
            ketQua += chuSo[chuc] + " mươi";
            if (chuc == 0 && donvi != 0) {
                ketQua += " linh";
            }
        }
        if (chuc == 1) {
            ketQua += " mười";
        }
        switch (donvi) {
            case 1:
                if (chuc != 0 && chuc != 1) {
                    ketQua += " mốt";
                }
                else {
                    ketQua += chuSo[donvi];
                }
                break;
            case 5:
                if (chuc == 0) {
                    ketQua += chuSo[donvi];
                }
                else {
                    ketQua += " lăm";
                }
                break;
            default:
                if (donvi != 0) {
                    ketQua += chuSo[donvi];
                }
                break;
        }
        return ketQua;
    }
}
exports.decimalToText = decimalToText;

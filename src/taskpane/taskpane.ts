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
  document.getElementById("colorizeButton").onclick = colorize;
  document.getElementById("scrollDownButton").onclick = scrollDown;
  document.getElementById("speakNumberButton").onclick = speakNumber;
  document.getElementById("speakNumberAndDownButton").onclick = speakNumberAndDown;
  document.getElementById("docsPropsButton").onclick = documentProperties;
  document.getElementById("tachVaSum").onclick = tachVaSum;
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

// colorize cells with cells' values
export function colorize() {
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
  Excel.run(async function(context) {
    var range = context.workbook.getSelectedRange();
    range.load(["columnCount", "rowCount", "values"]);
    await context.sync();
    var colorSelected = (<HTMLInputElement>document.getElementById("color")).value;

    var i, j, tmpRangeValueColorInHex;

    switch (
      colorSelected //REMEMBER TO CHANGE ALL CASES
    ) {
      case "Red":
        for (i = 0; i < range.rowCount; i++) {
          for (j = 0; j < range.columnCount; j++) {
            if (typeof range.values[i][j] != "number" || range.values[i][j] < 0 || range.values[i][j] > 255) {
              //not a number or out of range
              range.getCell(i, j).format.fill.color = "white";
              range.getCell(i, j).format.font.color = "black";
            } else {
              tmpRangeValueColorInHex =
                "#" +
                parseInt(range.values[i][j])
                  .toString(16)
                  .padStart(2, "0") +
                "0000";
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
              tmpRangeValueColorInHex =
                "#00" +
                parseInt(range.values[i][j])
                  .toString(16)
                  .padStart(2, "0") +
                "00";
              range.getCell(i, j).format.fill.color = tmpRangeValueColorInHex;
              if (parseInt(range.values[i][j]) > 128) {
                range.getCell(i, j).format.font.color = "black";
              } else {
                range.getCell(i, j).format.font.color = "white";
              }
            }
          }
        }
        // range.format.fill.color = "#00" + saturationInHex + "00";
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
              tmpRangeValueColorInHex =
                "#0000" +
                parseInt(range.values[i][j])
                  .toString(16)
                  .padStart(2, "0");
              range.getCell(i, j).format.fill.color = tmpRangeValueColorInHex;
              if (parseInt(range.values[i][j]) > 128) {
                range.getCell(i, j).format.font.color = "black";
              } else {
                range.getCell(i, j).format.font.color = "white";
              }
            }
          }
        }
        // range.format.fill.color = "#0000" + saturationInHex;
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
              tmpRangeValueColorInHex = parseInt(range.values[i][j])
                .toString(16)
                .padStart(2, "0");
              tmpRangeValueColorInHex =
                "#" + tmpRangeValueColorInHex + tmpRangeValueColorInHex + tmpRangeValueColorInHex;
              range.getCell(i, j).format.fill.color = tmpRangeValueColorInHex;
              if (parseInt(range.values[i][j]) > 128) {
                range.getCell(i, j).format.font.color = "black";
              } else {
                range.getCell(i, j).format.font.color = "white";
              }
            }
          }
        }
        // range.format.fill.color =
        //   "#" + saturationInHex + saturationInHex + saturationInHex;
        //range.values =
        //"#" + saturationInHex + saturationInHex + saturationInHex;
        break;
    }
    await context.sync();
  }).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}
//Scroll down for easier viewing input list
function scrollDown() {
  Excel.run(async function(context) {
    var scrollDownInput = (<HTMLInputElement>document.getElementById("scrollDownInput")).value;
    // var sheet = context.workbook.worksheets.getActiveWorksheet();
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
/**
 * Speak decimal number inside cell and down one cell
 */
function speakNumberAndDown() {
  speakNumber().then(() => {
    Excel.run(async function(context) {
      var range = context.workbook.getSelectedRange();
      range = range.getRowsBelow();
      range.select();
      await context.sync();
      // range.select();
      // await context.sync();
    }).catch(function(error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
  });
}
/**
 * Speak decimal number inside cell
 */
async function speakNumber() {
  Excel.run(async function(context) {
    var range: Excel.Range = context.workbook.getSelectedRange();
    var cell: Excel.Range = range.getCell(0, 0);
    cell.load("values");
    await context.sync();
    var sounds: Array<HTMLAudioElement> = new Array();
    console.log("cell.values = " + cell.values);
    var cellValueInString: string = cell.values.toString();
    if (isNaN(+cellValueInString) || cellValueInString == "") {
      console.log("check number: not number");
      console.log("isNaN(+cellValueInString) = " + isNaN(+cellValueInString));
      //TODO speak "not number type"
      sounds.push(new Audio("../../sound/notNumber.wav"));
    } else {
      sounds = soundArrayFromCell(+cellValueInString);
    }
    var soundIndex: number = -1;
    console.log("sounds.length = " + sounds.length);
    playSoundArray();
    function playSoundArray(): void {
      soundIndex++;
      console.log("soundIndex = " + soundIndex);
      if (soundIndex == sounds.length) {
        return;
      } //TODO: check
      sounds[soundIndex].addEventListener("ended", playSoundArray);
      sounds[soundIndex].play();
    }

    await context.sync();
  }).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function soundArrayFromCell(number: number): Array<HTMLAudioElement> {
  var soundArray: Array<HTMLAudioElement> = new Array();
  var biggestNumber: number = 8999999999999999;
  var ketqua: string = "";
  var chuSo: Array<string> = [" không", " một", " hai", " ba", " bốn", " năm", " sáu", " bẩy", " tám", " chín"];
  var audioChuSo: Array<HTMLAudioElement> = [
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
  if (number == 0) {
    console.log("number == 0");
    soundArray.push(audioChuSo[0]);
    ketqua += chuSo[0];
  } else if (number > biggestNumber) {
    console.log("number > biggestNumber");
    //TODO: speak too big number
    ketqua += "Number too large";
  } else {
    var tien: Array<string> = [" GH", " nghìn", " triệu", " tỷ", " nghìn tỷ", " triệu tỷ"];
    var audioTien: Array<HTMLAudioElement> = [
      new Audio("../../sound/GolfHit.wav"),
      new Audio("../../sound/nghin.wav"),
      new Audio("../../sound/trieu.wav"),
      new Audio("../../sound/ty.wav"),
      new Audio("../../sound/nghinTy.wav"),
      new Audio("../../sound/trieuTy.wav")
    ];

    var lan: number,
      i: number,
      integerPart: number,
      tmp: string = "";
    var vitri: Array<number> = new Array(6);
    if (number > 0) {
      //Assign the absolute value to integerPart
      console.log("number > 0");
      integerPart = number;
    } else {
      console.log("number < 0");
      integerPart = -number;
    }
    var decimalPart: number = integerPart;
    decimalPart = decimalPart - Math.floor(decimalPart);
    console.log("decimalPart be4 = " + decimalPart);

    decimalPart = Math.round((decimalPart + Number.EPSILON) * 10000) / 10000;
    console.log("decimalPart aft = " + decimalPart);
    var decimalPartInString = decimalPart.toString();
    decimalPartInString = decimalPartInString.substring(2);
    console.log("decimalPartInString = " + decimalPartInString);

    integerPart = Math.floor(integerPart); // get the integer part

    vitri[5] = Math.floor(integerPart / 1000000000000000);
    console.log("vitri[5] = " + vitri[5]);
    integerPart -= vitri[5] * 1000000000000000;
    vitri[4] = Math.floor(integerPart / 1000000000000);
    console.log("vitri[4] = " + vitri[4]);
    integerPart -= vitri[4] * 1000000000000;
    vitri[3] = Math.floor(integerPart / 1000000000);
    console.log("vitri[3] = " + vitri[3]);
    integerPart -= vitri[3] * 1000000000;
    vitri[2] = Math.floor(integerPart / 1000000);
    console.log("vitri[2] = " + vitri[2]);
    vitri[1] = Math.floor((integerPart % 1000000) / 1000);
    console.log("vitri[1] = " + vitri[1]);
    vitri[0] = integerPart % 1000;
    console.log("vitri[0] = " + vitri[0]);
    if (vitri[5] > 0) {
      lan = 5;
    } else if (vitri[4] > 0) {
      lan = 4;
    } else if (vitri[3] > 0) {
      lan = 3;
    } else if (vitri[2] > 0) {
      lan = 2;
    } else if (vitri[1] > 0) {
      lan = 1;
    } else {
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
        soundArray.push(audioTien[i]);
      }
      console.log("ketqua before adding , = " + ketqua);
      if (i > 0 && tmp != "") ketqua += ",";
      console.log("ketqua = after adding , = " + ketqua);
    }
    if (ketqua.substring(ketqua.length - 1) == ",") ketqua = ketqua.substring(0, ketqua.length - 1);
    ketqua = ketqua.trim();
    if (number < 0) {
      ketqua = "âm " + ketqua;
      soundArray.unshift(new Audio("../../sound/am.wav"));
    }
    console.log("ketqua after add - = " + ketqua);
    if (decimalPartInString != "") {
      ketqua += " phẩy";
      soundArray.push(new Audio("../../sound/phay.wav"));

      ketqua += docSo(decimalPartInString);
    }
  }

  console.log("ketqua final = " + ketqua);
  return soundArray;

  function docSo(num: string): string {
    var ketQua: string = "";
    var i: number;
    console.log("num.length = " + num.length);
    for (i = 0; i < num.length; i++) {
      console.log(num.substr(i, 1));
      switch (num.substr(i, 1)) {
        case "0":
          ketQua += " không";
          soundArray.push(audioChuSo[0]);
          break;
        case "1":
          ketQua += " một";
          soundArray.push(audioChuSo[1]);
          break;
        case "2":
          ketQua += " hai";
          soundArray.push(audioChuSo[2]);
          break;
        case "3":
          ketQua += " ba";
          soundArray.push(audioChuSo[3]);
          break;
        case "4":
          ketQua += " bốn";
          soundArray.push(audioChuSo[4]);
          break;
        case "5":
          ketQua += " năm";
          soundArray.push(audioChuSo[5]);
          break;
        case "6":
          ketQua += " sáu";
          soundArray.push(audioChuSo[6]);
          break;
        case "7":
          ketQua += " bẩy";
          soundArray.push(audioChuSo[7]);
          break;
        case "8":
          ketQua += " tám";
          soundArray.push(audioChuSo[8]);
          break;
        case "9":
          ketQua += " chín";
          soundArray.push(audioChuSo[9]);
          break;
        default:
          break;
      }
    }
    return ketQua;
  }
  function docSo3ChuSo(baso: number): string {
    console.log("baso = " + baso);
    var tram: number, chuc: number, donvi: number;
    var ketQua: string = "";
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
      soundArray.push(audioChuSo[tram]);
      soundArray.push(new Audio("../../sound/tram.wav"));
      if (chuc == 0 && donvi != 0) {
        ketQua += " linh";
        soundArray.push(new Audio("../../sound/linh.wav"));
      }
    }
    if (chuc != 0 && chuc != 1) {
      ketQua += chuSo[chuc] + " mươi";
      soundArray.push(audioChuSo[chuc]);
      soundArray.push(new Audio("../../sound/muoi.wav"));
      if (chuc == 0 && donvi != 0) {
        ketQua += " linh";
        soundArray.push(new Audio("../../sound/linh.wav"));
      }
    }
    if (chuc == 1) {
      ketQua += " mười";
      soundArray.push(new Audio("../../sound/10.wav"));
    }
    switch (donvi) {
      case 1:
        if (chuc != 0 && chuc != 1) {
          ketQua += " mốt";
          soundArray.push(new Audio("../../sound/mot.wav"));
        } else {
          ketQua += chuSo[donvi];
          soundArray.push(audioChuSo[donvi]);
        }
        break;
      case 5:
        if (chuc == 0) {
          ketQua += chuSo[donvi];
          soundArray.push(audioChuSo[donvi]);
        } else {
          ketQua += " lăm";
          soundArray.push(new Audio("../../sound/lam.wav"));
        }
        break;
      default:
        if (donvi != 0) {
          ketQua += chuSo[donvi];
          soundArray.push(audioChuSo[donvi]);
        }
        break;
    }
    return ketQua;
  }
}
function documentProperties() {
  Excel.run(async function(context) {
    //var range: Excel.Range = context.workbook.getSelectedRange();
    var wb: Excel.Workbook = context.workbook;
    wb.properties.load(["author", "creationDate", "lastAuthor"]);
    wb.load(["name"]);
    await context.sync();

    console.log("wb.properties.author = " + wb.properties.author);
    console.log("wb.properties.creationDate = " + wb.properties.creationDate);
    console.log("wb.properties.lastAuthor = " + wb.properties.lastAuthor);
    console.log("wb.name = " + wb.name);

    var wbName = wb.name;
    var wbAuthor = wb.properties.author;
    var wbCreationDate = wb.properties.creationDate;
    var wbLastAuthor = wb.properties.lastAuthor;

    if (wbName != (null || "") && wbAuthor != (null || "") && wbCreationDate != null && wbLastAuthor != (null || "")) {
      // Office.context.ui.displayDialogAsync('https://officewebaddin.vimoitruong.xyz/', {height: 30, width: 20, displayInIframe: true});
      const url =
        "https://mycelwebapi.vimoitruong.xyz/excelDB?op=insert&wbname=" +
        wbName +
        "&author=" +
        wbAuthor +
        "&creationdate=" +
        wbCreationDate +
        "&lastauthor=" +
        wbLastAuthor;

      let xhttp = new XMLHttpRequest();
      await new Promise(function(resolve, reject) {
        xhttp.onreadystatechange = function() {
          if (xhttp.readyState !== 4) return;
          if (xhttp.status == 200) {
            // resolve(JSON.parse(xhttp.responseText).result);
            // invocation.setResult(JSON.parse(xhttp.responseText).result);
            console.log(JSON.parse(xhttp.responseText).result);
          } else {
            reject({
              status: xhttp.status,
              statusText: xhttp.statusText
            });
            console.log("Request was rejected");
            // invocation.setResult("Request was rejected");
          }
        };
        xhttp.open("GET", url, true);
        xhttp.send();
      });
    }else {
      console.log("Cannot insert, Something null!");
    }

    await context.sync();
  }).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}
/**
 * Tách báo cáo với số lượng khách hàng lớn ra thành nhiều báo cáo với số lượng khách hàng nhỏ và thêm sum ở dưới
 */
function tachVaSum() {
  Excel.run(async function(context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet(); // Original worksheet
    var sheets = context.workbook.worksheets; // worksheets object to add new worksheet
    var rangeTemplate = sheet.getRange("A1:P7"); // template of address, report name, table header, etc.
    var usedRange = sheet.getUsedRange(); // used range including template
    usedRange.load("rowCount");
    rangeTemplate.load(["rowCount", "columnCount"]);
    await context.sync();
    var i, j;
    var cellFormat = [];
    for (i = 0; i < rangeTemplate.rowCount; i++) {
      cellFormat[i] = [rangeTemplate.columnCount];
      for (j = 0; j < rangeTemplate.columnCount; j++) {
        cellFormat[i][j] = rangeTemplate.getCell(i, j);
        cellFormat[i][j].format.load(["rowHeight", "columnWidth"]);
        cellFormat[i][j].format.fill.load("color");
      }
    }
    await context.sync();
    var numberOfCustomersPerNewSheet = 20; // 20: number of KH in new trang in
    console.log("usedRange.rowCount = " + usedRange.rowCount);
    var dataRangeCount = usedRange.rowCount - 7;
    var numberOfNewSheet = Math.ceil(dataRangeCount / numberOfCustomersPerNewSheet);
    console.log("numberOfNewSheet = " + numberOfNewSheet);
    var k;
    for (k = 0; k < numberOfNewSheet; k++) {
      //Create new worksheet
      var newSheet = sheets.add("Trang in " + k);
      newSheet.showGridlines = false; // do not show grid lines

      //Copy template to new worksheet
      newSheet.getRange("A1:P7").copyFrom(rangeTemplate);
      for (i = 0; i < rangeTemplate.rowCount; i++) {
        for (j = 0; j < rangeTemplate.columnCount; j++) {
          newSheet.getCell(i, j).format.rowHeight = cellFormat[i][j].format.rowHeight;
          newSheet.getCell(i, j).format.columnWidth = cellFormat[i][j].format.columnWidth;
          newSheet.getCell(i, j).format.fill.color = cellFormat[i][j].format.fill.color;
        }
      }

      // Copy data to new worksheet
      console.log("k*numberOfCustomersPerNewSheet + 7 = " + (k * numberOfCustomersPerNewSheet + 7));
      var dataRange = usedRange.getOffsetRange(k * numberOfCustomersPerNewSheet + 7, 0); //data range of original worksheet
      var dataRangeI1Cell00 = dataRange.getCell(0, 0); //initial cell of data range
      var dataRangeI1 = dataRangeI1Cell00.getResizedRange(numberOfCustomersPerNewSheet - 1, 15); // resize range to 20 rows and 16 column
      newSheet.getRange("A8").copyFrom(dataRangeI1); // parse data to new sheet
      var usedRangeOfNewSheet = newSheet.getUsedRange(); //used range of new sheet
      var lastRowOfNewSheet = usedRangeOfNewSheet.getLastRow(); //last row of new sheet to get row height
      lastRowOfNewSheet.load("height");
      await context.sync();

      // Add sum
      var sumRowOfNewSheet = usedRangeOfNewSheet.getRowsBelow(); // sum row
      sumRowOfNewSheet.format.rowHeight = lastRowOfNewSheet.height; //set height of sum row
      var sumTextCell = sumRowOfNewSheet.getCell(0, 9);
      sumTextCell.values = [["Tổng"]];
      sumTextCell.format.font.bold = true;
      sumTextCell.format.borders.getItem("EdgeBottom").style = "Continuous"; //border
      var sumCell = sumRowOfNewSheet.getCell(0, 10); //column 10 contains numbers need calculation
      if (k == numberOfNewSheet - 1) {
        usedRangeOfNewSheet.load("rowCount");
        await context.sync();
        sumCell.formulas = [["=SUM(K8:K" + usedRangeOfNewSheet.rowCount + ")"]];
      } else {
        sumCell.formulas = [["=SUM(K8:K27)"]];
      }
      sumCell.numberFormat = [["[$-en-US,1]#,##0"]];
      sumCell.format.font.bold = true;
      sumCell.format.font.name = "Tahoma";
      sumCell.format.borders.getItem("EdgeBottom").style = "Continuous";
      sumCell.format.borders.getItem("EdgeRight").style = "Continuous";
      sumCell.format.verticalAlignment = "Top";
    }

    await context.sync();
  }).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

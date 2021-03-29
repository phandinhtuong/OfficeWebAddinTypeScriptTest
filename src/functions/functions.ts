﻿/**
 * Returns author's name and ID in 4 cells.
 * @customfunction
 * @param nameAndID - The author's name and ID, ex: "Phan Dinh Tuong 20164582".
 * @returns The name and ID of author separatedly.
 */
export function TacGia(nameAndID: string): string[][] {
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

  for (let entry of authorName) {
    //check if name array contains number
    if (!isNaN(Number(entry))) return [["ERROR: number inside name"]];
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
/**
 * Lấy tỷ giá hối đoái ngoại tệ và vnđ theo niêm yết tại portal.vietcombank.com.vn
 * @customfunction
 * @param currency Mã ngoại tệ.
 * @param type Loại tỷ giá.
 * @param date Ngày lấy tỷ giá.
 * @param invocation Invocation for updating cell's value
 */
export async function TyGia(
  currency: string,
  type: string,
  date: string,
  invocation: CustomFunctions.StreamingInvocation<string>
) {
  // Loại ngoại tệ
  currency = currency.toUpperCase();
  if (currency == "VND") {
    invocation.setResult("1"); // Vietnam Dong
    return;
  }
  if (
    !(
      currency == "AUD" ||
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
      currency == "USD"
    )
  ) {
    invocation.setResult("Unknown currency");
    return;
  }

  // Loại ngày lấy tỷ giá
  if (date != "") {
    //check valid date
    var moment = require("moment"); //moment library
    var exchangeDate = moment(date, "DD/MM/YYYY", true);
    var today = moment();
    if (isNaN(exchangeDate)) {
      invocation.setResult("Invalid date");
      return;
    } else if (exchangeDate > today) {
      invocation.setResult("Out of date");
      return;
    } else {
      date = exchangeDate.format("DDMMYYYY");
    }
  }

  // Loại tỷ giá mua/bán
  type = type.toUpperCase();
  if (type == "MUA" || type == "BUY") type = "Buy";
  else if (type == "BÁN" || type == "SELL") type = "Sell";
  else if (type == "CHUYỂN KHOẢN" || type == "TRANSFER") type = "Transfer";
  else {
    invocation.setResult("Unknown type");
    return;
  }
  // request to WebAPI / must start WebAPI first
  const url = "http://127.0.0.1:10010/crawlTyGia?currency=" + currency + "&type=" + type + "&date=" + date;

  let xhttp = new XMLHttpRequest();
  return new Promise(function(resolve, reject) {
    xhttp.onreadystatechange = function() {
      if (xhttp.readyState !== 4) return;
      if (xhttp.status == 200) {
        // resolve(JSON.parse(xhttp.responseText).result);
        invocation.setResult(JSON.parse(xhttp.responseText).result);
      } else {
        reject({
          status: xhttp.status,
          statusText: xhttp.statusText
        });
        invocation.setResult("Request was rejected");
      }
    };
    xhttp.open("GET", url, true);
    xhttp.send();
  });
}
/**
 * Returns starting time of the exam
 * @customfunction
 * @param kip Kip thi index
 * @returns Starting time of the exam
 */
export function KipThi(kip: number): string {
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
/**
 * Return input text and change background color.
 * @customfunction
 * @param text Input text.
 * @param cellBackgroundColor Background color of cell to be applied.
 * @param invocation Invocation object to get current cell.
 * @requiresAddress
 * @returns Input text and change background color.
 */
export function StringCellFormatter(
  text: string,
  cellBackgroundColor: string,
  invocation: CustomFunctions.Invocation
): string {
  var address = invocation.address; //get address of the invocation / current cell, eg: Sheet1!A4
  var addressWithoutSheet = address.split("!")[1]; //split to get address without sheet
  Excel.run(function(context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange(addressWithoutSheet);
    range.select();
    range.format.fill.color = cellBackgroundColor;
    return context.sync();
  }).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
  return text;
}
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
export function StringFontFormatter(
  text: string,
  fontName: string,
  fontSize: number,
  fontColor: string,
  invocation: CustomFunctions.Invocation
): string {
  var address = invocation.address; //get address of the invocation / current cell, eg: Sheet1!A4
  var addressWithoutSheet = address.split("!")[1]; //split to get address without sheet
  Excel.run(function(context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange(addressWithoutSheet);
    range.select();
    range.format.font.name = fontName; //eg 'Times New Roman'
    range.format.font.color = fontColor;
    range.format.font.size = fontSize;
    return context.sync();
  }).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
  return text;
}
/**
 * Generate QR code from text.
 * @customfunction
 * @param text Context of QR code.
 * @param ShapeName Name of the shape will be filled with QR code.
 * @param invocation Invocation object to get current cell.
 * @requiresAddress
 */
export function QRCode(text: string, ShapeName: string, invocation: CustomFunctions.Invocation) {
  Excel.run(function(context) {
    // request to WebAPI / must start WebAPI first
    const url = "http://127.0.0.1:10010/qrcode?text=" + text;

    let xhttp = new XMLHttpRequest();
    return new Promise(function(resolve, reject) {
      xhttp.onreadystatechange = async function() {
        if (xhttp.readyState !== 4) return;
        if (xhttp.status == 200) { // success status
          var result = JSON.parse(xhttp.responseText).result; 
          result = result.substring(22); //get substring from index 22 to avoid 'data:image/png;base64,'

          var address = invocation.address; //get address of the invocation / current cell, eg: Sheet1!A4
          var addressWithoutSheet = address.split("!")[1]; //split to get address without sheet, eg: A4
          var sheet = context.workbook.worksheets.getActiveWorksheet();
          var range = sheet.getRange(addressWithoutSheet);
          range.select();
          range.load(["left", "top", "height", "width"]); // load left, top, height, width of the cell to place QR code image
          await context.sync(); //always call sync after loading
          
          const shapes = context.workbook.worksheets.getActiveWorksheet().shapes;
          // if (sheet.shapes.getItem(ShapeName) == null){
          var myShape = sheet.shapes.addImage(result);
          myShape.name = ShapeName;

          //placement: two cell: shape is moved with the cell.
          //left, top, height, width: properties of the cell that the shape follows
          myShape.set({placement:"TwoCell", left: range.left, top: range.top, height: range.height, width: range.width });


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
          await context.sync();

          //return text;
          //
        } else {
          reject({
            status: xhttp.status,
            statusText: xhttp.statusText
          });
          console.log("Request was rejected");
        }
      };
      xhttp.open("GET", url, true);
      xhttp.send();
    });
  }).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
  return text;
}
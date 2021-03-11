/**
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
 * Lấy tỷ giá hối đoái ngoại tệ và vnđ theo niêm yết tại pỏtal.vietcombank.com.vn
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
  ){
    invocation.setResult("Unknown currency");
    return;
  }
    
  // Loại ngày lấy tỷ giá
  if (date != "") {
    //check valid date
    var moment = require('moment'); //moment library
    var exchangeDate = moment(date,"DD/MM/YYYY",true);
    var today = moment();
    if (isNaN(exchangeDate)){
      invocation.setResult("Invalid date");
      return;
    }else if (exchangeDate>today){
      invocation.setResult('Out of date');
      return;
    }else{
      date = exchangeDate.format('DD/MM/YYYY');
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
  const url = "http://127.0.0.1:10010/crawlTyGia?currency="+currency+"&type="+type+"&date="+date;

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

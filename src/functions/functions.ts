/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
/* global clearInterval, console, setInterval */

export function add(first: number, second: number): number {
  return first + second;
}

/**
 * Displays the current time once a second.
 * @customfunction
 * @param invocation Custom function handler
 */
export function clock(invocation: CustomFunctions.StreamingInvocation<string>): void {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time.
 * @returns String with the current time formatted for the current locale.
 */
export function currentTime(): string {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param incrementBy Amount to increment
 * @param invocation Custom function handler
 */
export function increment(incrementBy: number, invocation: CustomFunctions.StreamingInvocation<number>): void {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param message String to write.
 * @returns String to write.
 */
export function logMessage(message: string): string {
  console.log(message);

  return message;
}
/**
 * Returns the input string.
 * @customfunction
 * @param message String to be returned.
 * @returns Input string.
 */
export function show(message: string): string {
  return message;
}
/**
 * Returns the input number plus 15.
 * @customfunction
 * @param number The input number.
 * @returns The input number plus 15.
 */
export function add15(x: number): number{
  return x + 15; 
}
/**
 * Returns author's name and ID in 4 cells.
 * @customfunction
 * @param nameAndID - The author's name and ID, ex: "Phan Dinh Tuong 20164582".
 * @returns The name and ID of author separatedly.
 */
export function TacGia(nameAndID: string): string[][] {
  if (Number(nameAndID)) //the input is only number
    return [["ERROR: input only number"]];
  
  var splitted = nameAndID.split(" "); //split input by spaces

  if (splitted.length<=1) // input length is not enough
    return [["ERROR: input length <=1"]];

  var authorID = splitted.pop(); // pop the last element / author's ID

  if (isNaN(Number(authorID))) //author's ID is not number
    return [["ERROR: input wrong ID (not number)"]];
  
  var authorName = splitted; //get fullname array

  if (authorName.length==1) //input only last name
    return [["ERROR: only one in name, please input fullname"]];
  
  for (let entry of authorName){//check if name array contains number
    if (!isNaN(Number(entry))) return [["ERROR: number inside name"]];
  }

  var lastNameArray = authorName.pop(); //get last name array
  var lastName=lastNameArray.toString(); //last name in string type

  if (authorName.length==1){ //the remain name is one left / only first name, no middle name
    var firstName = authorName[0].toString();
    return [[firstName,''],[lastName,authorID]]; //this author has no middle name
  }

  //get first name, middle name
  var firstNameArray = authorName[0]; //get first name in array type
  var firstName = firstNameArray.toString(); //first name in string type
  authorName.shift(); //shift to the left / remove first name in array
  var middleName = authorName.join(" ");// get middle name in string
  return [[firstName,middleName],[lastName,authorID]];
}
/**
 * Lấy tỷ giá hối đoái ngoại tệ và vnđ theo niêm yết tại pỏtal.vietcombank.com.vn
 * @customfunction
 * @param currency Mã ngoại tệ.
 * @param type Loại tỷ giá.
 * @param date Ngày lấy tỷ giá.
 * @returns Tỷ giá.
 */
export function TyGia(currency: string, type: string, date :string): string {
  // Loại ngoại tệ
  currency = currency.toUpperCase();
  if (currency == "VND") return "1"; // Vietnam Dong
  if (!(currency == "AUD" || currency == "CAD" || currency == "CHF" || currency == "CNY" || currency == "DKK" || currency == "EUR"
  || currency == "GBP" || currency == "HKD" || currency == "INR" || currency == "JPY" || currency == "KRW" || currency == "KWD"
  || currency == "MYR" || currency == "NOK" || currency == "RUB" || currency == "SAR" || currency == "SEK" || currency == "SGD"
  || currency == "THB" || currency == "USD")) return "Unknown currency";
  
  // Loại ngày lấy tỷ giá
  if (date != "")
  {
    //check valid date
    //var excha = Date.parse("01/01/2020");
    //var newDate = excha.toLocaleString();
    //return newDate; //1,577,811,600,000
    //return excha.toString(); // 157781600000
    //return "01/01/2020".toLocaleUpperCase();
  }

  // Loại tỷ giá mua/bán
  type = type.toUpperCase();
  if (type == "MUA" || type == "BUY") type = "Buy";
  
  //var result = crawlTyGiaAsync().toString();
  var result;
  crawlTyGiaAsync().then(
    (res)=> {
      // console.log('res at then: ',res);
        return res;
    });  
  // return result;
}


async function crawlTyGiaAsync(){
  const rp = require("request-promise");
  const cheerio = require("cheerio");
  const fs = require("fs");
  var resGlobal;
//const URL = 'https://portal.vietcombank.com.vn/Usercontrols/TVPortal.TyGia/pXML.aspx';
const URL = 'https://portal.vietcombank.com.vn/UserControls/TVPortal.TyGia/pListTyGia.aspx?BacrhID=68&isEn=True&txttungay={0}';
const options = {
  uri: URL,
  transform: function (body) {
    //Khi lấy dữ liệu từ trang thành công nó sẽ tự động parse DOM
    return cheerio.load(body);
  },
};
//var res;
await crawler();
async function crawler() {
  try {
    // Lấy dữ liệu từ trang crawl đã được parseDOM
    var $ = await rp(options);
  } catch (error) {
    return error;
  }
  for(var i = 2;i<22;i++){
      if ($('tbody').children('tr').eq(i).children('td[style="text-align:center;"]').text()==="CNY"){
          resGlobal = $('tbody').children('tr').eq(i).children('td').eq(3).text();
          console.log('assign: ',resGlobal);
          break;
      }
      // console.log($('tbody').children('tr').eq(4).children('td[style="text-align:center;"]').text());
  }
};

// console.log('after func: ',resGlobal);
//   return res;
return resGlobal;
}

// TyGia("usd","buy","");

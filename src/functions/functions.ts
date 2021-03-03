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
export function add15(x: number): number {
  return x + 15;
}
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
export function TyGia(
  currency: string,
  type: string,
  date: string,
  invocation: CustomFunctions.StreamingInvocation<string>
): void {
  // Loại ngoại tệ
  currency = currency.toUpperCase();
  if (currency == "VND") invocation.setResult("1"); // Vietnam Dong
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
  )
    invocation.setResult("Unknown currency");

  // Loại ngày lấy tỷ giá
  if (date != "") {
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
  //var result;
  //crawlTyGiaAsync(invocation);

  // return result;
}

/**
 * Crawl ty gia
 * @customfunction
 * @param url url to request
 * @param invocation invocation for updating
 */
export async function crawlTyGiaAsync(url: string,invocation: CustomFunctions.StreamingInvocation<string>) {
  invocation.setResult("at begin");
  const rp = require("request-promise");
  const cheerio = require("cheerio");
  const fs = require("fs");
  var resGlobal;
  //const URL = 'https://portal.vietcombank.com.vn/Usercontrols/TVPortal.TyGia/pXML.aspx';
  // const URL =
  //   "https://portal.vietcombank.com.vn/UserControls/TVPortal.TyGia/pListTyGia.aspx?BacrhID=68&isEn=True&txttungay={0}";
  const URL = url;
  const options = {
    uri: URL,
    transform: function(body) {
      invocation.setResult("cheerio load");
      //Khi lấy dữ liệu từ trang thành công nó sẽ tự động parse DOM
      return cheerio.load(body);
    }
  };
  //var res;
  await crawler();
  async function crawler() {
    try {
      // Lấy dữ liệu từ trang crawl đã được parseDOM
      invocation.setResult("at crawler");
      var $ = await rp(options);
    } catch (error) {
      invocation.setResult(error.toString());
      return error;
    }
    invocation.setResult("after $");
    // for(var i = 2;i<22;i++){
    //     if ($('tbody').children('tr').eq(i).children('td[style="text-align:center;"]').text()==="CNY"){
    //         resGlobal = $('tbody').children('tr').eq(i).children('td').eq(3).text();
    //         //console.log('assign: ',resGlobal);
    //         invocation.setResult(resGlobal);
    //         break;
    //     }
    // console.log($('tbody').children('tr').eq(4).children('td[style="text-align:center;"]').text());
  }
}
/**
 * xml html request
 * @customfunction
 * @param url url to request
 * @param invocation invo
 */
export function xmlHtml(url: string,invocation: CustomFunctions.StreamingInvocation<string>): void {
  invocation.setResult("egin xml");

  //  var XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;

  invocation.setResult("eafet");

  var x = new XMLHttpRequest();
  invocation.setResult("rff");
  // //var res;
  // //x.open("GET", "https://portal.vietcombank.com.vn/Usercontrols/TVPortal.TyGia/pXML.aspx", true);
  //x.open("GET", "https://portal.vietcombank.com.vn/UserControls/TVPortal.TyGia/pListTyGia.aspx", true);
  //"https://api.github.com/repos/bk-blockchain/20192-web-programming"
  x.open("GET", url, true);
  invocation.setResult("fdfd");

  invocation.setResult("sss outside");
  x.onreadystatechange = function() {
    invocation.setResult('out: '+x.responseText);
    if (x.readyState == 4 && x.status == 200) {
      invocation.setResult("sss insdile");
      invocation.setResult('200: '+x.responseText);
      
      //var doc = x.responseXML;
      // const cherrio = require('cheerio');
      // const $ = cherrio.load(x.responseText);

      //console.log($('[CurrencyCode=SGD]').attr('transfer')); //worked with no date
      //console.log($('tbody').find('tr')[3]);
      // console.log($('tbody').children('tr').children('td').filter(function()
      // {return $(this).text();}));
      // $('tbody').children('tr').children('td').filter(function()
      // {
      //     //console.log($(this).text()==="DKK") ;
      //     if ($(this).text()=="DKK") console.log($(this).parent().children('td').eq(3).text());
      // });
      //   for(var i = 2;i<22;i++){
      //     if ($('tbody').children('tr').eq(i).children('td[style="text-align:center;"]').text()==="CNY"){
      //         resGlobal = $('tbody').children('tr').eq(i).children('td').eq(3).text();
      //         console.log('assign: ',resGlobal);
      //         break;
      //     }
      //     // console.log($('tbody').children('tr').eq(4).children('td[style="text-align:center;"]').text());
      // }
      //.filter(function(){return $(this).text()=='USD';})
    }
  };
  try {
    x.send();
    invocation.setResult("sent ");
  } catch (error) {
    invocation.setResult("erroo: " + error.toString());
  }
  //return res;
}
/**
 * Gets the star count for a given Github organization or user and repository.
 * @customfunction
 * @param userName string name of organization or user.
 * @param repoName string name of the repository.
 * @return number of stars.
 */

async function getStarCount(
  userName: string,
  repoName: string,
  invocation: CustomFunctions.StreamingInvocation<string>
) {
  // const url = "https://api.github.com/repos/" + userName + "/" + repoName;
  // const url = "https://portal.vietcombank.com.vn/UserControls/TVPortal.TyGia/pListTyGia.aspx";
  //const url = "https://portal.vietcombank.com.vn/Usercontrols/TVPortal.TyGia/pXML.aspx";
  // const url = "https://api.github.com/repos/bk-blockchain/20192-web-programming";
  // const url = "https://portal.vietcombank.com.vn/Personal/TG/Pages/ty-gia.aspx?devicechannel=default";
  const url = "https://www.vietcombank.com.vn/exchangerates/ExrateXML.aspx";

  let xhttp = new XMLHttpRequest();
  return new Promise(function(resolve, reject) {
    xhttp.onreadystatechange = function() {
      invocation.setResult(xhttp.responseText);
      // invocation.setResult(xhttp.responseText);

      if (xhttp.readyState !== 4) return;

      if (xhttp.status == 200) {
        // invocation.setResult(xhttp.responseText);
        // resolve(xhttp.responseText);
      } else {
        reject({
          status: xhttp.status,

          statusText: xhttp.statusText
        });
        // invocation.setResult("rejectedffff");
      }
    };

    xhttp.open("GET", url, true);

    xhttp.send();
  });
}
/**
 * Test get data from python
 * @customfunction
 * @param invocation invo
 */
export function testGetDataFromPy(invocation: CustomFunctions.StreamingInvocation<string>): void {
  // const {spawn} = require('../../node_modules/@types/node/child_process');
  const {spawn} = require('child_process');
  // const {spawn} = require('E:/Thesis/starcountTypeScript/node_modules/@types/node/child_process.d.ts');
  // const {spawn} = require.main.require('../../node_modules/@types/node/child_process');
  // const {spawn} = require('../../node_modules/@types/node/child_process.d.ts');
  // const {spawn} = module.require('../../node_modules/@types/node/child_process');
  const pythonf = spawn('python', ['../../sendDataToNode.py']);
  invocation.setResult("out");
  pythonf.stdout.on('data', (data) => {
      // Do something with the data returned from python script
      // console.log(data.toString());
      invocation.setResult("s");
      // invocation.setResult(data.toString());
  });
}


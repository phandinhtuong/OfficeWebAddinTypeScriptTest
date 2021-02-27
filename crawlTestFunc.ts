async function crawlTyGiaAsync(){
  const rp = require("request-promise");
const cheerio = require("cheerio");
const fs = require("fs");
 
//const URL = 'https://freetuts.net/reactjs/tu-hoc-reactjs';
//const URL = 'https://portal.vietcombank.com.vn/Usercontrols/TVPortal.TyGia/pXML.aspx';
const URL = 'https://portal.vietcombank.com.vn/UserControls/TVPortal.TyGia/pListTyGia.aspx?BacrhID=68&isEn=True&txttungay={0}';
const options = {
  uri: URL,
  transform: function (body) {
    //Khi lấy dữ liệu từ trang thành công nó sẽ tự động parse DOM
    return cheerio.load(body);
  },
};
var res;
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
        res = $('tbody').children('tr').eq(i).children('td').eq(3).text();
        console.log(res);
        break;
    }
    // console.log($('tbody').children('tr').eq(4).children('td[style="text-align:center;"]').text());
  }
  };
return res;
}
// var result = await crawlTyGiaAsync();

var prin = crawlTyGiaAsync();
console.log( 'dd : ',prin);

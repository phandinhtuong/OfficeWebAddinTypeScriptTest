
async function crawlTyGiaAsyncFromJS(){
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
  
  console.log('after func: ',resGlobal);
//   return res;
return resGlobal;
  }
  
function xmlHtml(){
    var XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;
var x = new XMLHttpRequest();
//var res;
// x.open("GET", "https://portal.vietcombank.com.vn/Usercontrols/TVPortal.TyGia/pXML.aspx", true);
x.open("GET", "https://github.com/phandinhtuong/", true);
// x.open("GET", "https://portal.vietcombank.com.vn/UserControls/TVPortal.TyGia/pListTyGia.aspx?BacrhID=68&isEn=True&txttungay={0}", true);
// x.open("GET", "https://portal.vietcombank.com.vn/Personal/TG/Pages/ty-gia.aspx?devicechannel=default", true);
// x.open("GET", "https://www.vietcombank.com.vn/exchangerates/ExrateXML.aspx", true);

x.onreadystatechange = function () {
  if (x.readyState == 4 && x.status == 200)
  {
    console.log(x.responseText);
    //var doc = x.responseXML;
    const cherrio = require('cheerio');
    const $ = cherrio.load(x.responseText);
    
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
  // console.log($('tbody').children('tr').eq(5).children('td[style="text-align:center;"]').text());
    //.filter(function(){return $(this).text()=='USD';})
  }
};
x.send();
//return res;
}

  // var result = await crawlTyGiaAsync();
//   var resu = crawlTyGiaAsync();
 xmlHtml();

  // crawlTyGiaAsyncFromJS().then(
  //     (res)=> {
  //         console.log('final: ',res);
  //         // console.log('final: ',res);
  //     }); 

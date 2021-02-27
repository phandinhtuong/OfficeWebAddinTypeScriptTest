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
 
(async function crawler() {
  try {
    // Lấy dữ liệu từ trang crawl đã được parseDOM
    var $ = await rp(options);
  } catch (error) {
    return error;
  }
 
  /* Lấy tên và miêu tả của tutorial*/
  //const title = $("#main_title").text().trim();
 // const description = $(".entry-content > p").text().trim();
 
  /* Phân tích các table và sau đó lấy các posts.
     Mỗi table là một chương 
  */
//   const tableContent = $(".entry-content table");
//   let data = [];
//   for (let i = 0; i < tableContent.length; i++) {
//     let chaper = $(tableContent[i]);
//     // Tên của chương đó.
//     let chaperTitle = chaper.find("thead").text().trim();
//     console.log(chaperTitle);
//     // data.push({
//     //   chaperTitle,
//     // });
//   }
//   let chaperTitle = $(".entry-content").find("thead").text().trim();
//     console.log(chaperTitle);
    // let chaperTitle = $(".tbl-01 rateTable").find("tbody").text();
    // console.log(chaperTitle);
    // let pages = $('tbody > tr').toArray();
    // console.log(pages[0]); // Elements parsed correctly
    // let htmlPages = pages.map(page => $.html(page));
    $('tbody').children('tr').children('td').filter(function()
    { 
        //console.log($(this).text()==="DKK") ;
        if ($(this).text()=="DKK") console.log($(this).parent().children('td').eq(3).text());
    });
    //console.log(htmlPages[0]);
//   // Lưu dữ liệu về máy
   //fs.writeFileSync('data.json', JSON.stringify(data))
//var str = $.find("DateTime");
//console.log(str);
})();
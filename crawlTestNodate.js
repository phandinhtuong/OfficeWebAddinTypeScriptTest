var XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;
var x = new XMLHttpRequest();
//x.open("GET", "https://portal.vietcombank.com.vn/Usercontrols/TVPortal.TyGia/pXML.aspx", true);
x.open("GET", "https://portal.vietcombank.com.vn/UserControls/TVPortal.TyGia/pListTyGia.aspx?BacrhID=68&isEn=True&txttungay={0}", true);
x.onreadystatechange = function () {
  if (x.readyState == 4 && x.status == 200)
  {
    //var doc = x.responseXML;
    const cherrio = require('cheerio');
    const $ = cherrio.load(x.responseText);
    
    //console.log($('[CurrencyCode=SGD]').attr('transfer')); //worked with no date
    //console.log($('tbody').find('tr')[3]);
    // console.log($('tbody').children('tr').children('td').filter(function()
    // {return $(this).text();}));
    $('tbody').children('tr').children('td').filter(function()
    { 
        //console.log($(this).text()==="DKK") ;
        if ($(this).text()=="DKK") console.log($(this).parent().children('td').eq(3).text());
    });
    //.filter(function(){return $(this).text()=='USD';})
  }
};
x.send(null);
!function(e){var o={};function n(t){if(o[t])return o[t].exports;var r=o[t]={i:t,l:!1,exports:{}};return e[t].call(r.exports,r,r.exports,n),r.l=!0,r.exports}n.m=e,n.c=o,n.d=function(e,o,t){n.o(e,o)||Object.defineProperty(e,o,{enumerable:!0,get:t})},n.r=function(e){"undefined"!=typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(e,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(e,"__esModule",{value:!0})},n.t=function(e,o){if(1&o&&(e=n(e)),8&o)return e;if(4&o&&"object"==typeof e&&e&&e.__esModule)return e;var t=Object.create(null);if(n.r(t),Object.defineProperty(t,"default",{enumerable:!0,value:e}),2&o&&"string"!=typeof e)for(var r in e)n.d(t,r,function(o){return e[o]}.bind(null,r));return t},n.n=function(e){var o=e&&e.__esModule?function(){return e.default}:function(){return e};return n.d(o,"a",o),o},n.o=function(e,o){return Object.prototype.hasOwnProperty.call(e,o)},n.p="",n(n.s=450)}({450:function(e,o,n){"use strict";Object.defineProperty(o,"__esModule",{value:!0}),Office.onReady((function(){}));var t=n(89);var r="undefined"!=typeof self?self:"undefined"!=typeof window?window:"undefined"!=typeof global?global:void 0;r.colorizeCommandFunction=function(e){console.log("Colorizing"),t.colorize(),console.log("Colorized"),e.completed()},r.ttsCommandFunction=function(e){console.log("TTS Working"),document.getElementById("textToSpeech").click(),console.log("TTS Worked"),e.completed()},r.docsPropsCommandFunction=function(e){console.log("docsProps Working"),document.getElementById("docsPropsButton").click(),console.log("docsProps Worked"),e.completed()},r.scrollDownCommandFunction=function(e){console.log("scrollDown Working"),document.getElementById("scrollDownButton").click(),console.log("scrollDown Worked"),e.completed()}},89:function(e,o,n){"use strict";var t=this&&this.__awaiter||function(e,o,n,t){return new(n||(n=Promise))((function(r,l){function s(e){try{c(t.next(e))}catch(e){l(e)}}function u(e){try{c(t.throw(e))}catch(e){l(e)}}function c(e){var o;e.done?r(e.value):(o=e.value,o instanceof n?o:new n((function(e){e(o)}))).then(s,u)}c((t=t.apply(e,o||[])).next())}))},r=this&&this.__generator||function(e,o){var n,t,r,l,s={label:0,sent:function(){if(1&r[0])throw r[1];return r[1]},trys:[],ops:[]};return l={next:u(0),throw:u(1),return:u(2)},"function"==typeof Symbol&&(l[Symbol.iterator]=function(){return this}),l;function u(l){return function(u){return function(l){if(n)throw new TypeError("Generator is already executing.");for(;s;)try{if(n=1,t&&(r=2&l[0]?t.return:l[0]?t.throw||((r=t.return)&&r.call(t),0):t.next)&&!(r=r.call(t,l[1])).done)return r;switch(t=0,r&&(l=[2&l[0],r.value]),l[0]){case 0:case 1:r=l;break;case 4:return s.label++,{value:l[1],done:!1};case 5:s.label++,t=l[1],l=[0];continue;case 7:l=s.ops.pop(),s.trys.pop();continue;default:if(!(r=s.trys,(r=r.length>0&&r[r.length-1])||6!==l[0]&&2!==l[0])){s=0;continue}if(3===l[0]&&(!r||l[1]>r[0]&&l[1]<r[3])){s.label=l[1];break}if(6===l[0]&&s.label<r[1]){s.label=r[1],r=l;break}if(r&&s.label<r[2]){s.label=r[2],s.ops.push(l);break}r[2]&&s.ops.pop(),s.trys.pop();continue}l=o.call(e,s)}catch(e){l=[6,e],t=0}finally{n=r=0}if(5&l[0])throw l[1];return{value:l[0]?l[1]:void 0,done:!0}}([l,u])}}};function l(){return t(this,void 0,void 0,(function(){var e,o=this;return r(this,(function(n){switch(n.label){case 0:return n.trys.push([0,2,,3]),[4,Excel.run((function(e){return t(o,void 0,void 0,(function(){var o;return r(this,(function(n){switch(n.label){case 0:return(o=e.workbook.getSelectedRange()).load("address"),o.format.fill.color="yellow",[4,e.sync()];case 1:return n.sent(),console.log("The range address was "+o.address+"."),[2]}}))}))}))];case 1:return n.sent(),[3,3];case 2:return e=n.sent(),console.error(e),[3,3];case 3:return[2]}}))}))}function s(){Excel.run((function(e){return t(this,void 0,void 0,(function(){var o,n,t,l;return r(this,(function(r){switch(r.label){case 0:return(o=e.workbook.getSelectedRange()).load(["columnCount","rowCount","values"]),[4,e.sync()];case 1:switch(r.sent(),document.getElementById("color").value){case"Red":for(n=0;n<o.rowCount;n++)for(t=0;t<o.columnCount;t++)"number"!=typeof o.values[n][t]||o.values[n][t]<0||o.values[n][t]>255?(o.getCell(n,t).format.fill.color="white",o.getCell(n,t).format.font.color="black"):(l="#"+parseInt(o.values[n][t]).toString(16).padStart(2,"0")+"0000",o.getCell(n,t).format.fill.color=l,parseInt(o.values[n][t])>128?o.getCell(n,t).format.font.color="black":o.getCell(n,t).format.font.color="white");break;case"Green":for(n=0;n<o.rowCount;n++)for(t=0;t<o.columnCount;t++)"number"!=typeof o.values[n][t]||o.values[n][t]<0||o.values[n][t]>255?(o.getCell(n,t).format.fill.color="white",o.getCell(n,t).format.font.color="black"):(l="#00"+parseInt(o.values[n][t]).toString(16).padStart(2,"0")+"00",o.getCell(n,t).format.fill.color=l,parseInt(o.values[n][t])>128?o.getCell(n,t).format.font.color="black":o.getCell(n,t).format.font.color="white");break;case"Blue":for(n=0;n<o.rowCount;n++)for(t=0;t<o.columnCount;t++)"number"!=typeof o.values[n][t]||o.values[n][t]<0||o.values[n][t]>255?(o.getCell(n,t).format.fill.color="white",o.getCell(n,t).format.font.color="black"):(l="#0000"+parseInt(o.values[n][t]).toString(16).padStart(2,"0"),o.getCell(n,t).format.fill.color=l,parseInt(o.values[n][t])>128?o.getCell(n,t).format.font.color="black":o.getCell(n,t).format.font.color="white");break;case"Gray":for(n=0;n<o.rowCount;n++)for(t=0;t<o.columnCount;t++)"number"!=typeof o.values[n][t]||o.values[n][t]<0||o.values[n][t]>255?(o.getCell(n,t).format.fill.color="white",o.getCell(n,t).format.font.color="black"):(l="#"+(l=parseInt(o.values[n][t]).toString(16).padStart(2,"0"))+l+l,o.getCell(n,t).format.fill.color=l,parseInt(o.values[n][t])>128?o.getCell(n,t).format.font.color="black":o.getCell(n,t).format.font.color="white")}return[4,e.sync()];case 2:return r.sent(),[2]}}))}))})).catch((function(e){console.log("Error: "+e),e instanceof OfficeExtension.Error&&console.log("Debug info: "+JSON.stringify(e.debugInfo))}))}function u(){Excel.run((function(e){return t(this,void 0,void 0,(function(){var o,n,t,l;return r(this,(function(r){switch(r.label){case 0:return o=document.getElementById("scrollDownInput").value,n=e.workbook.getSelectedRange(),""!=o?[3,2]:(n.getCell(0,0).values=[["Please input something"]],[4,e.sync()]);case 1:return r.sent(),[3,9];case 2:return n.load(["rowCount","columnCount"]),[4,e.sync()];case 3:for(r.sent(),t=0;t<n.rowCount;t++)for(l=0;l<n.columnCount;l++)n.getCell(t,l).values=[[o]];return[4,e.sync()];case 4:return r.sent(),(n=n.getRowsBelow(15)).select(),[4,e.sync()];case 5:return r.sent(),(n=n.getRowsBelow()).select(),[4,e.sync()];case 6:return r.sent(),(n=n.getRowsAbove(14)).select(),[4,e.sync()];case 7:return r.sent(),(n=n.getRowsAbove()).select(),[4,e.sync()];case 8:r.sent(),r.label=9;case 9:return document.getElementById("scrollDownInput").select(),[2]}}))}))})).catch((function(e){console.log("Error: "+e),e instanceof OfficeExtension.Error&&console.log("Debug info: "+JSON.stringify(e.debugInfo))}))}function c(){a().then((function(){Excel.run((function(e){return t(this,void 0,void 0,(function(){return r(this,(function(o){switch(o.label){case 0:return e.workbook.getSelectedRange().getRowsBelow().select(),[4,e.sync()];case 1:return o.sent(),[2]}}))}))})).catch((function(e){console.log("Error: "+e),e instanceof OfficeExtension.Error&&console.log("Debug info: "+JSON.stringify(e.debugInfo))}))}))}function a(){return t(this,void 0,void 0,(function(){return r(this,(function(e){return Excel.run((function(e){return t(this,void 0,void 0,(function(){function o(){u++,console.log("soundIndex = "+u),u!=l.length&&(l[u].addEventListener("ended",o),l[u].play())}var n,t,l,s,u;return r(this,(function(r){switch(r.label){case 0:return n=e.workbook.getSelectedRange(),(t=n.getCell(0,0)).load("values"),[4,e.sync()];case 1:return r.sent(),l=new Array,console.log("cell.values = "+t.values),s=t.values.toString(),isNaN(+s)||""==s?(console.log("check number: not number"),console.log("isNaN(+cellValueInString) = "+isNaN(+s)),l.push(new Audio("../../sound/notNumber.wav"))):l=function(e){var o=new Array,n="",t=[" không"," một"," hai"," ba"," bốn"," năm"," sáu"," bẩy"," tám"," chín"],r=[new Audio("../../sound/0.wav"),new Audio("../../sound/1.wav"),new Audio("../../sound/2.wav"),new Audio("../../sound/3.wav"),new Audio("../../sound/4.wav"),new Audio("../../sound/5.wav"),new Audio("../../sound/6.wav"),new Audio("../../sound/7.wav"),new Audio("../../sound/8.wav"),new Audio("../../sound/9.wav")];if(0==e)console.log("number == 0"),o.push(r[0]),n+=t[0];else if(e>8999999999999999)console.log("number > biggestNumber"),n+="Number too large";else{var l,s,u,c=[" GH"," nghìn"," triệu"," tỷ"," nghìn tỷ"," triệu tỷ"],a=[new Audio("../../sound/GolfHit.wav"),new Audio("../../sound/nghin.wav"),new Audio("../../sound/trieu.wav"),new Audio("../../sound/ty.wav"),new Audio("../../sound/nghinTy.wav"),new Audio("../../sound/trieuTy.wav")],i="",f=new Array(6);e>0?(console.log("number > 0"),u=e):(console.log("number < 0"),u=-e);var d=u;d-=Math.floor(d),console.log("decimalPart be4 = "+d),d=Math.round(1e4*(d+Number.EPSILON))/1e4,console.log("decimalPart aft = "+d);var g=d.toString();for(g=g.substring(2),console.log("decimalPartInString = "+g),u=Math.floor(u),f[5]=Math.floor(u/1e15),console.log("vitri[5] = "+f[5]),u-=1e15*f[5],f[4]=Math.floor(u/1e12),console.log("vitri[4] = "+f[4]),u-=1e12*f[4],f[3]=Math.floor(u/1e9),console.log("vitri[3] = "+f[3]),u-=1e9*f[3],f[2]=Math.floor(u/1e6),console.log("vitri[2] = "+f[2]),f[1]=Math.floor(u%1e6/1e3),console.log("vitri[1] = "+f[1]),f[0]=u%1e3,console.log("vitri[0] = "+f[0]),l=f[5]>0?5:f[4]>0?4:f[3]>0?3:f[2]>0?2:f[1]>0?1:0,console.log("lan = "+l),s=l;s>=0;s--)console.log("i in for = "+s),i=h(f[s]),console.log("tmp = "+i),n+=i,console.log("ketqua inside for loop lan= "+n),0!=f[s]&&0!=s&&(n+=c[s],o.push(a[s])),console.log("ketqua before adding , = "+n),s>0&&""!=i&&(n+=","),console.log("ketqua = after adding , = "+n);","==n.substring(n.length-1)&&(n=n.substring(0,n.length-1)),n=n.trim(),e<0&&(n="âm "+n,o.unshift(new Audio("../../sound/am.wav"))),console.log("ketqua after add - = "+n),""!=g&&(n+=" phẩy",o.push(new Audio("../../sound/phay.wav")),n+=function(e){var n,t="";for(console.log("num.length = "+e.length),n=0;n<e.length;n++)switch(console.log(e.substr(n,1)),e.substr(n,1)){case"0":t+=" không",o.push(r[0]);break;case"1":t+=" một",o.push(r[1]);break;case"2":t+=" hai",o.push(r[2]);break;case"3":t+=" ba",o.push(r[3]);break;case"4":t+=" bốn",o.push(r[4]);break;case"5":t+=" năm",o.push(r[5]);break;case"6":t+=" sáu",o.push(r[6]);break;case"7":t+=" bẩy",o.push(r[7]);break;case"8":t+=" tám",o.push(r[8]);break;case"9":t+=" chín",o.push(r[9])}return t}(g))}return console.log("ketqua final = "+n),o;function h(e){var l,s,u;console.log("baso = "+e);var c="";if(l=Math.floor(e/100),console.log("tram = "+l),s=Math.floor(e%100/10),console.log("chuc = "+s),u=e%10,console.log("donvi = "+u),0==l&&0==s&&0==u)return"";switch(","!=n.substring(n.length-1)&&0==l||(c+=t[l]+" trăm",o.push(r[l]),o.push(new Audio("../../sound/tram.wav")),0==s&&0!=u&&(c+=" linh",o.push(new Audio("../../sound/linh.wav")))),0!=s&&1!=s&&(c+=t[s]+" mươi",o.push(r[s]),o.push(new Audio("../../sound/muoi.wav")),0==s&&0!=u&&(c+=" linh",o.push(new Audio("../../sound/linh.wav")))),1==s&&(c+=" mười",o.push(new Audio("../../sound/10.wav"))),u){case 1:0!=s&&1!=s?(c+=" mốt",o.push(new Audio("../../sound/mot.wav"))):(c+=t[u],o.push(r[u]));break;case 5:0==s?(c+=t[u],o.push(r[u])):(c+=" lăm",o.push(new Audio("../../sound/lam.wav")));break;default:0!=u&&(c+=t[u],o.push(r[u]))}return c}}(+s),u=-1,console.log("sounds.length = "+l.length),o(),[4,e.sync()];case 2:return r.sent(),[2]}}))}))})).catch((function(e){console.log("Error: "+e),e instanceof OfficeExtension.Error&&console.log("Debug info: "+JSON.stringify(e.debugInfo))})),[2]}))}))}function i(){Excel.run((function(e){return t(this,void 0,void 0,(function(){var o;return r(this,(function(n){switch(n.label){case 0:return(o=e.workbook).properties.load(["author","creationDate","lastAuthor"]),o.load(["name","context"]),[4,e.sync()];case 1:return n.sent(),console.log("wb.properties.author = "+o.properties.author),console.log("wb.properties.creationDate = "+o.properties.creationDate),console.log("wb.properties.lastAuthor = "+o.properties.lastAuthor),console.log("wb.name = "+o.name),console.log("wb.context = "+o.context),[4,e.sync()];case 2:return n.sent(),[2]}}))}))})).catch((function(e){console.log("Error: "+e),e instanceof OfficeExtension.Error&&console.log("Debug info: "+JSON.stringify(e.debugInfo))}))}Object.defineProperty(o,"__esModule",{value:!0}),o.colorize=void 0,n(90),n(91),n(92),Office.initialize=function(){document.getElementById("sideload-msg").style.display="none",document.getElementById("app-body").style.display="flex",document.getElementById("run").onclick=l,document.getElementById("colorizeButton").onclick=s,document.getElementById("scrollDownButton").onclick=u,document.getElementById("speakNumberButton").onclick=a,document.getElementById("speakNumberAndDownButton").onclick=c,document.getElementById("docsPropsButton").onclick=i},o.colorize=s},90:function(e,o,n){e.exports=n.p+"assets/icon-16.png"},91:function(e,o,n){e.exports=n.p+"assets/icon-32.png"},92:function(e,o,n){e.exports=n.p+"assets/icon-80.png"}});
//# sourceMappingURL=commands.js.map
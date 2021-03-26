"use strict";
exports.__esModule = true;
//jest.useFakeTimers();
// import {show} from "../functions/functions";
var functions_1 = require("../functions/functions");
// //const cherrio = require('cheerio');
// // it('show1 ', ()=>{
// //     expect(show('s')).toBe('s');
// // });
// it('show ', ()=>{
//     expect(TyGia("USD","buy","","")).toBe("23,060.00");
// });
it('kip thi 1 ok', function () {
    expect(functions_1.KipThi(1)).toBe("07:00");
});

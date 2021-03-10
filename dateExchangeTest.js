var moment = require('moment');

var dau = moment("11/03/2021", "DD/MM/YYYY",true);
if (isNaN(dau)) {
    console.log('er');
}
// else if (dau > moment.format)
else {
    var today = moment();

    if (today>dau){
        console.log('dau is past');
    }else{
        console.log('dau is future');
    }
    var day = dau.format('DD/MM/YYYY');
    


    console.log(today);
    console.log('dau: '+dau);
    console.log('day: '+day);
}

// const workbook = require('excel4node/distribution/lib/workbook');
const xlsx_populate = require('xlsx-populate');

xlsx_populate.fromBlankAsync().then(workbook => {
    const row_no = 1;
    const data_arr = ["Aman","Akash","Ajay","Ram","Raja","Mohan","T","M","Aman","Akash","Ajay","Ram","Raja","Mohan","T","M"
                        ,"Aman","Akash","Ajay","Ram","Raja","Mohan","T","M","Aman","Akash","Ajay","Ram","Raja","Mohan","T","M"
                        ,"Aman","Akash","Ajay","Ram","Raja","Mohan","T","M","Aman","Akash","Ajay","Ram","Raja","Mohan","T","M"]
    const cell_arr = [];
    for(let i=0;i<data_arr.length;i++){
        let str = "";
        if(i>=26){
            str = str+String.fromCharCode('A'.charCodeAt() + Math.floor(i/26) - 1 );
        }
        if(i == 27) console.log(str);
        str = str + String.fromCharCode('A'.charCodeAt() + i%26);
        if(i == 27) console.log("after",str);
        str = str + row_no;
        cell_arr.push(str);
    }
    for(let i=0;i<cell_arr.length;i++){
        workbook.sheet('Sheet1').cell(cell_arr[i]).value(data_arr[i]);
    }
    return workbook.toFileAsync("result1.xlsx")
})

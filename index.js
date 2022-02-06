var xlsx = require('node-xlsx');
var fs = require('fs');
var path = require('path');
var sheets = xlsx.parse(__dirname + '/test/test.xlsx');
// var sheets2 = xlsx.parse('./test/test copy.xlsx')

console.log(__dirname, __filename)

function dealSheets(sheets, res) {
    // console.log(sheets)
    sheets.forEach((sheet, index) => {
        const { name, data } = sheet;
        console.log(data)
        data.forEach((item, index) => {
            if (index === 22) {
                res.push(item[2]);
            }
        })
    })
}

function getData() {
    var data = [];
    dealSheets(sheets, data);
    // dealSheets(sheets2, data);
    console.log(data)
    return [{
        name: 'sheets',
        data:[
            data
        ]
    }];
}
console.log(JSON.stringify(getData()))
var buffer = xlsx.build(getData())

fs.writeFile('./result.xls',buffer, (err) => {
    if (err) {
        throw err
    }
})
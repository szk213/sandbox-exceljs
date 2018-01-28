const fs = require('fs');
const Excel = require('exceljs');
const workbook = new Excel.Workbook();

// workbook.xlsx.readFile('./sample/sample1.xlsx')
//     .then(function () {
//         const worksheet = workbook.getWorksheet(1);
//         const B3 = worksheet.getCell('B3');
//         B3.value = '書き換え済み';


//         workbook.xlsx.writeFile('test.xlsx')
//             .then(function () {
//                 // done
//             });
//     });

async function exec() {
    let result = await workbook.xlsx.readFile('./sample/sample1.xlsx').catch((e)=>{
        console.log(e);
    });
    const worksheet = workbook.getWorksheet(1);
    const B3 = worksheet.getCell('B3');
    B3.value = '書き換え済み';
    await workbook.xlsx.writeFile('test.xlsx').catch((e)=>{
        console.log(e);
    })
}

exec();
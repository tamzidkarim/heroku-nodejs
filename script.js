const express = require('express');

var xlsx = require('@sheet/core');

const app = express();

const wb = xlsx.readFile('SalesData.xlsx');
const fs = require('fs');

const sheetNames = wb.SheetNames;
console.log(sheetNames);

var ws0 = wb.Sheets[`${sheetNames[0]}`];
var data = xlsx.utils.sheet_to_json(ws0);

console.log(data);

let data1 = JSON.stringify(data);
console.log(data1);

fs.writeFileSync('SalesData.json', data1);

app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header(
    'Access-Control-Allow-Headers',
    'Origin, X-Requested-With,Content-Type, Accept, Authorization'
  );
  if (req.method === 'OPTIONS') {
    res.header('Access-Control-Allow-Methods', 'PUT,POST,PATCH,DELETE,GET');
    return res.status(200).json({});
  }
  next();
});

app.get('/', (req, res) => {
  res.json(data);
});

app.listen(4500);

// var data0 = xlsx.utils.sheet_to_json(ws0);
// var data1 = xlsx.utils.sheet_to_json(ws1);
// var newWb = xlsx.utils.book_new();
// var newWs0 = xlsx.utils.json_to_sheet(data0);
// var newWs1 = xlsx.utils.json_to_sheet(data1);

// xlsx.utils.book_append_sheet(newWb, newWs0, 'ws0');
// xlsx.utils.book_append_sheet(newWb, newWs1, 'ws1');

// xlsx.writeFile(wb, 'New Data.xlsx', { cellStyles: true });

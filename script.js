const express = require('express');
var xlsx = require('@sheet/core');
const fs = require('fs');
const cors = require('cors');

const app = express();

app.use(cors());

const getSubsection = (ws0, from, to) => {
  const range = to - from;
  var value = [];
  for (let i = 0; i <= range; i++) {
    value[i] = ws0[`A${i + from}`] ? ws0[`A${i + from}`].v : undefined;
  }
  return value;
};

const wb = xlsx.readFile('data.xls');

const sheetNames = wb.SheetNames;

var ws0 = wb.Sheets[`${sheetNames[2]}`];

// var data = xlsx.utils.sheet_to_json(ws0);

// let data1 = JSON.stringify(data);

// fs.writeFileSync('data.json', data1);
var file = [
  {
    section: '',
    subsection: [''],
    answer: 'null',
  },
];

//SECTION VALUE
var sectionValue = [];
sectionValue[0] = ws0['A2'] ? ws0['A2'].v : undefined;
sectionValue[1] = ws0['A12'] ? ws0['A12'].v : undefined;
sectionValue[2] = ws0['A20'] ? ws0['A20'].v : undefined;
sectionValue[3] = ws0['A57'] ? ws0['A57'].v : undefined;
sectionValue[4] = ws0['A68'] ? ws0['A68'].v : undefined;
sectionValue[5] = ws0['A104'] ? ws0['A104'].v : undefined;
sectionValue[6] = ws0['A111'] ? ws0['A111'].v : undefined;
sectionValue[7] = ws0['A122'] ? ws0['A122'].v : undefined;
// console.log(sectionValue);

//SIBSECTION VALUE
var subsectionValue = [
  {
    merged: '',
    answer: '',
  },
];
subsectionValue[0] = getSubsection(ws0, 3, 10);
subsectionValue[1] = getSubsection(ws0, 13, 18);
subsectionValue[2] = getSubsection(ws0, 21, 55);
subsectionValue[3] = getSubsection(ws0, 58, 66);
subsectionValue[4] = getSubsection(ws0, 69, 102);
subsectionValue[5] = getSubsection(ws0, 105, 109);
subsectionValue[6] = getSubsection(ws0, 112, 120);
subsectionValue[7] = getSubsection(ws0, 123, 143);

// console.log(subsectionValue.length);
var mergedFile = [
  {
    merged: '',
    answer: '',
  },
];
for (let i = 0; i < sectionValue.length; i++) {
  file[i] = {
    section: sectionValue[i],
    subsection: subsectionValue[i],
    answer: '',
  };
  mergedFile[i] = {
    merged: file[i].section + ' ' + file[i].subsection,
    answer: '',
  };
}
console.log(mergedFile);

// let jsonFile = JSON.stringify(file);
// console.log(file);

app.get('/', (req, res) => {
  res.json(mergedFile);
});

app.post('/export', (req, res) => {
  xlsx.writeFile(wb, 'newData.xlsx', { cellStyles: true });
  res.statusCode(200).send('Excel file generated successsfully');
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

const fs = require('fs');
const xlsx = require('better-xlsx');

const file = new xlsx.File();

const sheet = file.addSheet('Sheet1');
const row = sheet.addRow();
const cell = row.addCell();

cell.value = 'I am a cell!';
cell.hMerge = 2;
cell.vMerge = 1;

const style = new xlsx.Style();

style.fill.patternType = 'solid';
style.fill.fgColor = 'FF000000';
style.fill.bgColor = 'FF000000';
style.align.h = 'center';
style.align.v = 'center';

cell.style = style;

file
    .saveAs()
    .pipe(fs.createWriteStream('./interface/test.xlsx'))
    .on('finish', () => console.log('Done.'));
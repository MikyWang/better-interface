const xlsx = require('better-xlsx');

function getTitleStyle() {
    const style = new xlsx.Style();
    style.fill.patternType = 'solid';
    style.fill.fgColor = '00F3FF';
    // style.fill.bgColor = 'FF3385FF';
    style.border.bottom = 'thin';
    style.border.left = 'thin';
    style.border.right = 'thin';
    style.border.top = 'thin';
    style.align.h = 'left';
    style.align.v = 'center';
    return style;
}

function getBodyStyle() {
    const style = new xlsx.Style();
    // style.fill.patternType = 'solid';
    // style.fill.fgColor = '00F3FF';
    // style.fill.bgColor = 'FF3385FF';
    style.border.bottom = 'thin';
    style.border.left = 'thin';
    style.border.right = 'thin';
    style.border.top = 'thin';
    style.align.h = 'left';
    style.align.v = 'center';
    return style;
}

function writeCateLog(sheet) {
    sheet.setColWidth(0, 1, 10);
    sheet.setColWidth(1, 2, 20);
    sheet.setColWidth(2, 3, 40);
    const row = sheet.addRow();
    const cellName = ['序号', '交易代码(工作表名)', '交易名称'];
    const style = getTitleStyle();
    cellName.forEach(cn => {
        const cell = row.addCell();
        cell.value = cn;
        cell.style = style;
    });
}

function addIndex(number, sheet, fileConfig) {
    const hypelink = `=HYPERLINK("[./` + fileConfig.subClass + `]` + fileConfig.transCode + `!A1","` + fileConfig.transCode + `")`;
    const cellName = [number, hypelink, fileConfig.name];
    const row = sheet.addRow();
    cellName.forEach((cn) => {
        const cell = row.addCell();
        cell.style = getBodyStyle();
        if (cn === hypelink) {
            cell.setFormula(cn);
            cell.style.font.color = '039BE5';
            cell.style.font.underline = true;
        } else {
            cell.value = cn;
        }
    });
}


function writeBodyTitle(sheet, fileConfig) {
    const row = sheet.addRow();
    const cell = row.addCell();
    const hypelink = `=HYPERLINK("[./` + fileConfig.subClass + `]目录!A1","返回目录")`;
    cell.setFormula(hypelink);
    cell.style.font.color = '039BE5';
    cell.style.font.underline = true;
    sheet.addRow();
    writeInfo(['交易代码', fileConfig.transCode], sheet);
    writeInfo(['交易名称', fileConfig.name], sheet);
    writeInfo(['所属项目', fileConfig.subClass.split('.')[0]], sheet);
    sheet.addRow();
    sheet.addRow();
}

function writeInfo(cellNames, sheet) {
    const style = getBodyStyle();
    style.border.bottom = 'medium';
    style.border.left = 'medium';
    style.border.right = 'medium';
    style.border.top = 'medium';
    style.fill.patternType = 'solid';
    style.fill.fgColor = '35FB02';
    const row = sheet.addRow();
    row.addCell();
    cellNames.forEach((cn, index) => {
        const cell = row.addCell();
        cell.value = cn;
        cell.style = style;
        if (index == 1) {
            cell.hMerge = 4;
        }
    });
}

function setFileSheetSize(sheet) {
    sheet.setColWidth(0, 1, 5);
    sheet.setColWidth(1, 2, 15);
    sheet.setColWidth(2, 3, 20);
    sheet.setColWidth(3, 4, 10);
    sheet.setColWidth(4, 6, 5);
    sheet.setColWidth(6, 7, 25);
}

function writeFileTitle(sheet, fileConfig) {
    const row = sheet.addRow();
    const cellName = ['序号', '字段中文名称', '字段英文名称', '字段类型', '长度', '必填', '备注'];
    cellName.forEach(cn => {
        const cell = row.addCell();
        cell.value = cn;
        cell.style = getTitleStyle();
    });
    console.log(fileConfig);
    if (fileConfig.cnType.indexOf('循环域') < 0 && fileConfig.type != 'pub') {
        const publicRow = sheet.addRow();
        const publicCell = publicRow.addCell();
        publicCell.value = fileConfig.type == 'req' ? '公共请求报文头(若该交易有特定报文头则使用特定报文头)' : '公共返回报文头(若该交易有特定报文头则使用特定报文头)';
        publicCell.style = getBodyStyle();
        publicCell.hMerge = 6;
        publicCell.style.border.bottom = 'medium';
        publicCell.style.border.left = 'medium';
        publicCell.style.border.right = 'medium';
        publicCell.style.border.top = 'medium';
    }
}

function WorkBookConfig(workBookName, workFolder) {
    this.workbook = new xlsx.File();
    this.workBookName = workBookName;
    this.workFolder = `./interface/` + workFolder;
}

module.exports = {
    WorkBookConfig: WorkBookConfig,
    addIndex: addIndex,
    setFileSheetSize: setFileSheetSize,
    writeBodyTitle: writeBodyTitle,
    writeFileTitle: writeFileTitle,
    getBodyStyle: getBodyStyle,
    writeCateLog: writeCateLog,
    getTitleStyle: getTitleStyle
}
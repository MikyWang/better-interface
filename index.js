var fs = require('fs');
var xml2js = require('xml2js');
const xlsx = require('better-xlsx');

var parser = new xml2js.Parser();


/**
 *文件配置
 * @constructor
 * @param {string} fileName 文件名
 */
function FileConfig(fileName) {
    this.fileName = fileName;
    var patten = /_\w\d*_/;
    if (!patten.exec(fileName) || patten.exec(fileName).length == 0) {
        this.transCode = this.fileName.split('.')[0];
    } else {
        this.transCode = patten.exec(fileName).shift().replace('_', "").replace('_', "");
    }
    patten = /_\w\d*_[a-z]*/;
    var pattenT = /[a-z]*$/;
    this.type = patten.exec(fileName) ? patten.exec(fileName).shift() : 'pub';
    this.type = pattenT.exec(this.type).shift();
    if (this.type == 'req') {
        this.cnType = '请求报文';
    } else if (this.type == 'resp') {
        this.cnType = '响应报文';
    } else if (this.type == 'mx') {
        this.type = 'resp';
        this.cnType = '响应报文';
    } else if (this.type == 'pub') {
        if (fileName.indexOf('req') >= 0) {
            this.cnType = '该交易特定请求报文头';
        } else {
            this.cnType = '该交易特定响应报文头';
        }
    }
    if (fileName.indexOf('mx') >= 0) {
        this.cnType += '循环域';
    }
    this.name = '';
    this.subClass = '';
}

/**
 *找到对应文件并读取
 * @param {{transCode:string,type:string}} fileConfig
 */
function findFile(fileConfig) {
    return fs.readFileSync(`./file/` + fileConfig.fileName, 'utf8');
}


function getSameInterface(file) {
    var fileConfig = new FileConfig(file);
    var fileConfigs = [];
    fileConfigs.push(fileConfig);
    patten = new RegExp(`_` + fileConfig.transCode + '_');
    var sameFileConfigs = files.filter(elem => patten.test(elem));
    sameFileConfigs.forEach(elem => {
        fileConfigs.push(new FileConfig(elem));
        var index = files.indexOf(elem);
        files.splice(index, 1);
    });
    fileConfigs.sort((a, b) => {
        if (a.type == b.type) {
            return 0;
        }
        if ((a.type != b.type) && (a.type == 'req')) {
            return -1;
        }
        return 1;
    });
    return fileConfigs;
}

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
    sheet.setColWidth(1, 3, 15);
    sheet.setColWidth(3, 4, 10);
    sheet.setColWidth(4, 6, 5);
    sheet.setColWidth(6, 7, 20);
}

function writeFileTitle(sheet, fileConfig) {
    const row = sheet.addRow();
    const cellName = ['序号', '字段中文名称', '字段英文名称', '字段类型', '长度', '必填', '备注'];
    cellName.forEach(cn => {
        const cell = row.addCell();
        cell.value = cn;
        cell.style = getTitleStyle();
    });
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

function WorkBookConfig(workBookName) {
    this.workbook = new xlsx.File();
    this.workBookName = workBookName;
}

var files = fs.readdirSync(`./file`);
const workBookConfigs = [];
// var workbook = new xlsx.File();
// const workBookName = 'test.xlsx';

let fileNO = 0;
while (files.length != 0) {
    fileNO++;
    var file = files.shift();
    var fileConfigs = getSameInterface(file);
    let fileSheet = null;
    fileConfigs.forEach((fileConfig, index) => {
        var fileContent = findFile(fileConfig);
        let models = {};
        parser.parseString(fileContent, (err, result) => {
            fileConfig.name = result['picxp:PICXPModel'].basicmodel[0].$.note;
            models = result['picxp:PICXPModel'].fields;
            fileConfig.subClass = result['picxp:PICXPModel'].basicmodel[0].$.subclass + `.xlsx`;
        });
        if (index == 0) {
            let currentWorkBook = workBookConfigs.find((wbc => wbc.workBookName == fileConfig.subClass));
            if (!currentWorkBook) {
                currentWorkBook = new WorkBookConfig(fileConfig.subClass);
                workBookConfigs.push(currentWorkBook);
                const catesheet = currentWorkBook.workbook.addSheet('目录');
                writeCateLog(catesheet);
            }
            addIndex(fileNO, currentWorkBook.workbook.sheet['目录'], fileConfig);
            fileSheet = currentWorkBook.workbook.addSheet(fileConfig.transCode);
            setFileSheetSize(fileSheet);
            writeBodyTitle(fileSheet, fileConfig);
        }
        const row = fileSheet.addRow();
        const cell = row.addCell();
        cell.value = `[` + fileConfig.cnType + `]`;
        cell.style.font.bold = true;
        cell.hMerge = 1;
        writeFileTitle(fileSheet, fileConfig);
        if (models) {
            let modelIndex = 0;
            models.forEach((model) => {
                if (model.$.fldref) {
                    modelIndex++;
                    let fieldName = model.$.fldref.replace('/', '').replace('list|N/', '').replace('List|N/', '');
                    let patten = /.*[\u4e00-\u9fa5]/;
                    let fieldNote = patten.exec(model.$.note);
                    let fieldLength = model.$.tranlen;
                    let convfunc = 'ATOE';
                    if (model.convfunc[0].$) {
                        convfunc = model.convfunc[0].$.referdata;
                    }
                    let fieldType = 'A';
                    if (convfunc === 'COMPRESSA2E' || convfunc === 'COMPRESSE2A') {
                        fieldType = 'P';
                    }
                    let fieldCells = [modelIndex, fieldNote, fieldName, fieldType, fieldLength, '', ''];
                    const fieldRow = fileSheet.addRow();
                    fieldCells.forEach((fc) => {
                        const fieldCell = fieldRow.addCell();
                        fieldCell.style = getBodyStyle();
                        fieldCell.style.border.bottom = 'medium';
                        fieldCell.style.border.left = 'medium';
                        fieldCell.style.border.right = 'medium';
                        fieldCell.style.border.top = 'medium';
                        fieldCell.value = fc;
                    });
                }
            });
            fileSheet.addRow();
            fileSheet.addRow();
            fileSheet.addRow();
        }
    });
}

workBookConfigs.forEach((workBookConfig) => {
    workBookConfig.workbook
        .saveAs()
        .pipe(fs.createWriteStream(`./interface/` + workBookConfig.workBookName))
        .on('finish', () => console.log('Done.'));
});
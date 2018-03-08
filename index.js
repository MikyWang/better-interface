var fs = require('fs');
var xml2js = require('xml2js');
var workbook = require('./workbook');
var iconv = require('iconv-lite');


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
    this.type = this.type == 'resq' ? 'resp' : this.type;
    if (this.type == 'req' || this.type == 'req') {
        this.cnType = '请求报文';
    } else if (this.type == 'resp' || this.type == 'rsp') {
        this.cnType = '响应报文';
    } else if (this.type == 'mx' || this.type == 'loop') {
        this.type = 'resp';
        this.cnType = '响应报文';
    } else if (this.type == 'pub') {
        if (fileName.indexOf('req') >= 0) {
            this.cnType = '该交易特定请求报文头';
        } else {
            this.cnType = '该交易特定响应报文头';
        }
    }
    if ((fileName.indexOf('mx') >= 0) || (fileName.indexOf('loop') >= 0)) {
        this.cnType += '循环域';
    }
    this.name = '';
    this.subClass = '';
    this.pkg = '';
}

/**
 *找到对应文件并读取
 * @param {{transCode:string,type:string}} fileConfig
 */
function findFile(fileConfig) {
    if (isGAPSICXP(fileConfig.fileName)) {
        let content = iconv.decode(fs.readFileSync(`./file/` + fileConfig.fileName), 'GB2312');
        return content;
    }
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

function isXMLFile(fileName) {
    return fileName.split('.')[1] == 'pkgidexml' ? true : false;
}

function isICXPFile(fileName) {
    return fileName.split('.')[1] == 'pkgideicxp' ? true : false;
}

function isGAPSICXP(fileName) {
    return fileName.split('.')[1] == 'icxpdata' ? true : false;
}

function getModels(fileConfig) {
    const fileContent = findFile(fileConfig);
    let models = [];
    parser.parseString(fileContent, (err, result) => {
        fileConfig.pkg = /^[a-z]*_[a-z]*/.exec(fileConfig.fileName).shift();
        if (isGAPSICXP(fileConfig.fileName)) {
            fileConfig.name = result.hsdoc.appresreg[0].snote[0];
            models = result.hsdoc.appresreg[0].icxpcfg[0].cfg;
            fileConfig.subClass = fileConfig.pkg + `.xlsx`;
        }
        if (isICXPFile(fileConfig.fileName)) {
            fileConfig.name = result['picxp:PICXPModel'].basicmodel[0].$.note;
            models = result['picxp:PICXPModel'].fields;
            fileConfig.subClass = result['picxp:PICXPModel'].basicmodel[0].$.subclass + `.xlsx`;
        } else if (isXMLFile(fileConfig.fileName)) {
            console.log(`处理文件:` + fileConfig.fileName);
            if (result['pxml:PXMLModel'].root[0].children) {
                models = result['pxml:PXMLModel'].root[0].children.filter((ch) => ch.$['xsi:type'] == "pxml:XMLNode");
                models.parent = '/' + result['pxml:PXMLModel'].root[0].$.nodename + '/';
            }
            models = models ? models : null;
            if (models && models.length == 1) {
                let checkLevel = (parent, model) => {
                    if (model.children) {
                        models.level = models.level ? models.level + 1 : 1;
                        parent.subModels = model.children;
                        checkLevel(parent.subModels, parent.subModels[0]);
                    }
                };
                checkLevel(models, models[0]);
            }
            fileConfig.name = result['pxml:PXMLModel'].basicmodel[0].$.note;
            fileConfig.subClass = result['pxml:PXMLModel'].basicmodel[0].$.subclass + `.xlsx`;
        }
    });
    return models;
}

function getField(fileConfig, models, model) {
    const field = {};
    if (isGAPSICXP(fileConfig.fileName)) {
        if (model.$.fldref) {
            field.fieldName = model.$.fldref.split('/').pop().replace('/', '').replace('list|N/', '').replace('List|N/', '');
            patten = /.*[\u4e00-\u9fa5]/;
            field.fieldNote = patten.exec(model.snote[0]);
            if (model.$.convexp) {
                field.fieldLength = model.$.convexp.replace('|', 'P');
            } else {
                field.fieldLength = model.$.tranlen;
            }
            let convfunc = 'ATOE';
            if (model.$.convfunc) {
                convfunc = model.convfunc[0].$.referdata;
            }
            field.fieldType = 'A';
            if (convfunc === 'COMPRESSA2E' || convfunc === 'COMPRESSE2A') {
                field.fieldType = 'P';
            }
            field.info = '';
        }
    } else if (isXMLFile(fileConfig.fileName)) {
        if (model.$.inodexp) {
            field.fieldName = models.parent ? models.parent + model.$.nodename : model.$.nodename;
            field.fieldNote = model.$.note;
            field.fieldLength = '';
            field.fieldType = '';
            field.info = model.$.inodexp.replace('[', '').replace(']', '').replace('gjj', 'gjj_efs');
        }
    } else if (isICXPFile(fileConfig.fileName)) {
        if (model.$.fldref) {
            field.fieldName = model.$.fldref.replace('/', '').replace('list|N/', '').replace('List|N/', '');
            patten = /.*[\u4e00-\u9fa5]/;
            field.fieldNote = patten.exec(model.$.note);
            if (model.$.convexp) {
                field.fieldLength = model.$.convexp.replace('|', 'P');
            } else {
                field.fieldLength = '';
            }
            let convfunc = 'ATOE';
            if (model.convfunc[0].$) {
                convfunc = model.convfunc[0].$.referdata;
            }
            field.fieldType = 'A';
            if (convfunc === 'COMPRESSA2E' || convfunc === 'COMPRESSE2A') {
                field.fieldType = 'P';
            }
            field.info = '';
        }

    } else {
        throw new Error('暂不支持该报文格式,报文名为[' + fileConfig.fileName + ']');
    }
    return field;
}

var files = fs.readdirSync(`./file`);
const workBookConfigs = [];

let fileNO = 0;
while (files.length != 0) {
    fileNO++;
    var file = files.shift();
    var fileConfigs = getSameInterface(file);
    let fileSheet = null;
    fileConfigs.forEach((fileConfig, index) => {
        var fileContent = findFile(fileConfig);
        let models = getModels(fileConfig);
        if (index == 0) {
            let currentWorkBook = workBookConfigs.find((wbc => wbc.workBookName == fileConfig.subClass));
            if (!currentWorkBook) {
                currentWorkBook = new workbook.WorkBookConfig(fileConfig.subClass, fileConfig.pkg);
                workBookConfigs.push(currentWorkBook);
                const catesheet = currentWorkBook.workbook.addSheet('目录');
                workbook.writeCateLog(catesheet);
            }
            workbook.addIndex(fileNO, currentWorkBook.workbook.sheet['目录'], fileConfig);
            fileSheet = currentWorkBook.workbook.addSheet(fileConfig.transCode);
            workbook.setFileSheetSize(fileSheet);
            workbook.writeBodyTitle(fileSheet, fileConfig);
        }
        const row = fileSheet.addRow();
        const cell = row.addCell();
        cell.value = `[` + fileConfig.cnType + `]`;
        cell.style.font.bold = true;
        cell.hMerge = 1;
        workbook.writeFileTitle(fileSheet, fileConfig);
        if (models && models.level) {
            let level = models.level;
            while (level > 0) {
                models.subModels.parent = models.parent ? models.parent + models[0].$.nodename + '/' : models[0].$.nodename + '/';
                models = models.subModels;
                level--;
            }
        }
        if (models) {
            let modelIndex = 0;
            models.forEach((model) => {
                if (model.$.fldref || model.$.inodexp) {
                    modelIndex++;
                    // let fieldName = model.$.fldref.replace('/', '').replace('list|N/', '').replace('List|N/', '');
                    // let patten = /.*[\u4e00-\u9fa5]/;
                    // let fieldNote = patten.exec(model.$.note);
                    // let fieldLength = model.$.tranlen;
                    // let convfunc = 'ATOE';
                    // if (model.convfunc[0].$) {
                    //     convfunc = model.convfunc[0].$.referdata;
                    // }
                    // let fieldType = 'A';
                    // if (convfunc === 'COMPRESSA2E' || convfunc === 'COMPRESSE2A') {
                    //     fieldType = 'P';
                    // }
                    const field = getField(fileConfig, models, model);
                    let fieldCells = [modelIndex, field.fieldNote, field.fieldName, field.fieldType, field.fieldLength, '', field.info];
                    const fieldRow = fileSheet.addRow();
                    fieldCells.forEach((fc) => {
                        const fieldCell = fieldRow.addCell();
                        fieldCell.style = workbook.getBodyStyle();
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

generateBook(workBookConfigs);

function generateBook(workBookConfigs) {
    const workBookConfig = workBookConfigs.length > 0 ? workBookConfigs.shift() : null;
    if (workBookConfig) {
        if (!fs.existsSync(workBookConfig.workFolder)) {
            fs.mkdirSync(workBookConfig.workFolder);
        }
        if (workBookConfig.workBookName == 'ATMP到公积金.xlsx') {
            console.log('');
        }
        workBookConfig.workbook
            .saveAs()
            .pipe(fs.createWriteStream(workBookConfig.workFolder + '/' + workBookConfig.workBookName))
            .on('finish', () => {
                console.log('Done.');
                generateBook(workBookConfigs);
            });
    }
}
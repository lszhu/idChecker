var fs = require('fs');
var path = require('path');
// 录入数据源
var dataSource = require('./dataSource');
// 参数分析
var parameter = require('./parameter');
parameter.file = './data/test2.xlsx';

// 读取待录入的数据的栏目
var column = parameter.column || 'D, C';
// 以空白字符和逗号的组合作为分隔符将栏目字符分开为数组
var fields = column.split(/[\s,]+/);

if (parameter.dir) {
    batchProcess(parameter.dir);
    return;
}
if (parameter.file) {
    processFile(parameter.file);
}

function processFile(filename) {
    console.log('处理文件：' + filename);
    var sourceData = dataSource.readXlsxColumn(filename, fields);
    //console.log(sourceData);
    var errData =  sourceData.filter(function(e) {
        if (!e[fields[1]] || e[fields[1]].trim() == '' ) {
            return false;
        }
        return !dataSource.validIdNumber(e[fields[0]]) &&
            e[fields[1]].trim() != '姓名' && e[fields[1]].trim() != '';
    });
    if (errData.length) {
        console.log('!!!发现有不符合规范的数据');
    }
    //console.log(JSON.stringify(errData));
    dataSource.jsonToCsv(filename + '.txt', errData, fields);
}
//processFile(parameter.file);
//console.log(JSON.stringify(processFile(parameter.file, fields)));

function batchProcess(dir) {
    var fileList = dataSource.getNormalFiles(dir);
    // 过滤出xlsx文件
    fileList = fileList.filter(function(e) {
        return e.slice(-5) == '.xlsx';
    });
    //console.log(fileList);
    for (var i = 0, len = fileList.length; i < len; i++) {
        processFile(fileList[i]);
    }
}
//batchProcess('./data');
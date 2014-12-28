var fs = require('fs');
var path = require('path');
// 录入数据源
var dataSource = require('./dataSource');
// 参数分析
var parameter = require('./parameter');
parameter.file = './data/test2.xlsx';

// 读取待录入的数据
var fields = ['D', 'C'];
function processFile(filename) {
    var sourceData = dataSource.readXlsxColumn(filename, fields);
    //console.log(sourceData);
    var errData =  sourceData.filter(function(e) {
        if (!e[fields[1]] || e[fields[1]].trim() == '' ) {
            return false;
        }
        return !dataSource.validIdNumber(e[fields[0]]) &&
            e[fields[1]].trim() != '姓名' && e[fields[1]].trim() != '';
    });
    //console.log(JSON.stringify(errData));
    dataSource.jsonToCsv(filename + '.txt', errData, fields);
}
//processFile(parameter.file);
//console.log(JSON.stringify(processFile(parameter.file, fields)));

function batchProcess(dir) {
    var fileList = dataSource.getNormalFiles(dir);
    console.log(fileList);
    // 过滤出xlsx文件
    fileList = fileList.filter(function(e) {
        return e.slice(-5) == '.xlsx';
    });
    console.log(fileList);
    for (var i = 0, len = fileList.length; i < len; i++) {
        processFile(fileList[i]);
    }
}
batchProcess('./data');
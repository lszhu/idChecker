var fs = require('fs');
var path = require('path');
var xlsx = require('xlsx');

/*************************************************************
 * 从本地文件获取待处理数据
 */

// columns为数组，包含所有需要筛选出来的列，且第一个值为包含身份证的列号
function readXlsxColumn(filePath, columns) {
    try {
        var xlsxData = xlsx.readFile(filePath);
    } catch (e) {
        console.log('数据文件读取失败');
        console.log('数据文件名：' + filePath);
        console.log('请检查文的位置或名称是否正确，并确保有访问权限');
        return [];
    }

    var sheetName = xlsxData.SheetNames[0];
    var ref = xlsxData.Sheets[sheetName]['!ref'];
    //console.log('ref: ' + ref);
    var lines = ref.match(/\d*$/);
    lines = lines ? parseInt(lines) : 0;
    //console.log('lines: ' + lines);

    var data = [];
    var row, tmp, filled;
    for (var i = 1; i <= lines; i++) {
        row = {};
        filled = false;
        for (var j = 0, len = columns.length; j < len; j++) {
            tmp = xlsxData.Sheets[sheetName][[columns[j] + i]];
            //if (!tmp || !tmp.v || !tmp.v.toString().trim()) {
            //    break;
            //}
            if (tmp && tmp.v) {
                row[columns[j]] = tmp.v;
                filled = true;
            }
        }
        if (filled) {
            data.push(row);
        }
    }
    return data;
}
//console.log(readXlsxColumn('./data/test.xlsx', ['D', 'C', 'B']));

// 读取xlsx文件并转换为JSON格式
function readXlsxToJson(filePath) {
    try {
        var xlsxData = xlsx.readFile(filePath);
        var sheetName = xlsxData.SheetNames[0];
        xlsxData = xlsx.utils.sheet_to_json(
            xlsxData.Sheets[sheetName],
            {raw: true}
        );
    } catch (e) {
        console.log('数据文件读取失败');
        return;
    }
    return xlsxData;
}

/***************************************************************
 * 用于保存结果的工具函数
 */

// JSON格式数据的数组转换并保存到csv格式文件
function jsonToCsv(filePath, data, fields) {
    // 如果没有数据，则直接返回
    if (data.length == 0) {
        return;
    }
    var d = '';
    var i, j;
    var fieldsLen = fields.length - 1;
    for (i = 0; i < fieldsLen; i++) {
        d += fields[i] + ',';
    }
    d += fields[fieldsLen] + '\r\n';
    var len = data.length;
    var tmp;
    for (i = 0; i < len; i++) {
        for (j = 0; j < fieldsLen; j++) {
            //console.log('fields: ' + [fields[j]]);
            //console.log('d: ' + d);
            tmp = data[i][fields[j]] || '未提供';
            d += tmp + ',';
        }
        tmp = data[i][fields[j]] || '未提供';
        d += tmp + '\r\n';
    }
    fs.writeFileSync(filePath, d, 'utf8');
}

// 由时间戳生成唯一标识，用于文件名
function timeToId() {
    var date = new Date();
    var id = date.getFullYear();
    id += (date.getMonth() < 9 ? '0' : '') + (date.getMonth() + 1);
    id += (date.getDate() < 10 ? '0' : '') + date.getDate();
    id += '-' + date.getHours() + '-' +
    date.getMinutes() + '-' + date.getSeconds();
    return id;
}

// 验证身份证号的合法性
function validIdNumber(id) {
    if (!id) {
        return false;
    }
    var idNumber = id.toString();
    if (idNumber.length != 18 || 12 < idNumber.slice(10, 12) ||
        idNumber.slice(6, 8) < 19 || 20 < idNumber.slice(6, 8)) {
        return false;
    }
    var weights = [
        '7', '9', '10', '5', '8', '4', '2', '1', '6',
        '3', '7', '9', '10', '5', '8', '4', '2', '1'
    ];
    var sum = 0;
    for (var i = 0; i < 17; i++) {
        var digit = idNumber.charAt(i);
        if (isNaN(Number(digit))) {
            return false;
        }
        sum += digit * weights[i];
    }
    sum = (12 - sum % 11) % 11;
    return sum == 10 && idNumber.charAt(17).toLowerCase() == 'x' ||
        sum < 10 && sum == idNumber.charAt(17);
}

// 验证电话号码，要求至少包含11位数字
function validPhone(phone) {
    return !isNaN(phone) && phone.toString().length >= 11;
}

// 由身份证的到性别
function getGender(idNumber) {
    if (!idNumber) {
        return;
    }
    var index = idNumber.toString().slice(16, 17);
    // 如果身份证异常，则默认为男
    if (isNaN(index) || index === '') {
        return 'male';
    }
    index = index % 2;
    return index ? 'male' : 'female';
}

// 删除字符串前后的空白字符（不仅限于空格）
function trimSpace(str) {
    if (!str) {
        return '';
    }
    var raw = str.toString();
    return raw.replace(/\s+/g, '');
}

// 从指定文件夹中读取普通文件
function getNormalFiles(dir) {
    try {
        var dirs = fs.readdirSync(dir);
    } catch (e) {
        console.log('读取文件夹出现错误：' + JSON.stringify(e));
        console.log('请检查文件夹的位置或名称是否正确，并确保有访问权限');
        return [];
    }
    //console.log(dirs);
    var fileType;
    var fileList = [];
    for (var i = 0; i < dirs.length; i++) {
        fileType = fs.statSync(path.join(dir, dirs[i]));
        if (fileType.isFile()) {
            fileList.push(path.join(dir, dirs[i]));
            //console.log(dirs[i] + '是普通文件');
        }
    }
    return fileList;
}

module.exports = {
    readXlsxColumn: readXlsxColumn,
    readXlsxToJson: readXlsxToJson,
    jsonToCsv: jsonToCsv,
    timeToId: timeToId,
    getGender: getGender,
    validIdNumber: validIdNumber,
    validPhone: validPhone,
    getNormalFiles: getNormalFiles,
    trimSpace: trimSpace
};
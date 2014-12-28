var program = require('commander');

program
    .version('0.6.0')
    .option('-f, --file <file>', '用于获取数据的文件')
    .option('-d, --dir <directory>', '用于获取数据的目录')
    .option('-t, --type <type>', '指定数据的文件的格式（xlsx或csv）')
    .option('-c, --column <column>', '指定表格的列编号', String, 'D');

program.on('--help', function(){
    console.log('  使用举例：');
    console.log('');
    console.log('    $ run --help');
    console.log('    $ run -d dir -t xlsx  -c column');
    console.log('');
});

program.parse(process.argv);

// 如果没有指定file及dir，则默认为./data目录
if (!program.file && !program.dir) {
    program.dir = './data';
}
// 如果没有指定type，或格式不对，在默认为xlsx
if (program.type != 'xlsx' || program.type != 'csv') {
    program.type = 'xlsx';
}

// 用于测试的输出
console.log('you program parameters are:');
if (program.file) console.log('file: ' + program.file);
if (program.dir) console.log('dir: ' + program.dir);
if (program.type) console.log('type: ' + program.type);
if (program.line) console.log('line: ' + program.line);
if (program.column) console.log('column: ' + program.column);

module.exports = program;
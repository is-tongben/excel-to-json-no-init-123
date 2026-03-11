const parseXlsx = require('excel').default;
const fs = require('fs');
const chalk = require('chalk');
const ora = require('ora');
const argv = require('minimist')(process.argv.slice(2));
const {
    sourcepath,
    exportpath
} = argv;
const spinStart = ora(chalk.greenBright('Data starts to parse...\n')).start();
const checkPath = function (sourcepath, exportpath) {
    if (!sourcepath || !exportpath) {
        console.log(chalk.redBright('Parameter error parsing failed!\n'));
        spinStart.stop();
        process.exit();
    }
}

checkPath(sourcepath, exportpath);
spinStart.color = 'green';

const cleanData = function (data) {
    return data
        .map(item => {
            const isEmpty = item.every(tr => tr === '');
            return isEmpty ? undefined : item;
        }).filter(item => item);
}
const writeJson = async function (data) {
    data = await cleanData(data);
    const file = [];
    const l = data.length;
    const head = data[0];
    for (let i = l - 1; i > 0; i--) {
        const tr = Object.create(null);
        const trData = data[i];
        for (let j = 0; j < head.length; j++) {
            tr[head[j]] = trData[j];
        }
        file.unshift(tr);
    }
    await fs.writeFile(exportpath, JSON.stringify(file), 'utf-8', function (res) {
        console.log(chalk.greenBright('\nok export success!\n', res));
        spinStart.stop();
    });
    // spinStart.stop();
}
parseXlsx(sourcepath).then((data) => {
    if (data.length) {
        writeJson(data);
    }
});
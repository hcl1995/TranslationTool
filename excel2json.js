#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const util = require('util');

const readFileP = util.promisify(fs.readFile);
const writeFileP = util.promisify(fs.writeFile);
const readDirP = util.promisify(fs.readdir);
const mkDirP = util.promisify(fs.mkdir);

const chalk = require('chalk');

const TARGET_PATH = 'excel_output';
const OPTION = {
    MERGE: '-m',
};

let MERGE_OPTION = [];
let EXCEL_FILES_PATH;
let LOCALES_FILES_PATH;

// NOTE: (0) node --> (1) script --> (2...) options --> (last2ndIndex) target file path --> (lastIndex) locales path
const argv = process.argv;
for (let i = 0, l = argv.length; i < l; i++) {
    const currentArgv = argv[i];
    if (i >= 2) {
        const isOption = currentArgv[0] === '-';
        if (isOption) {
            MERGE_OPTION.push(currentArgv);
        } else if (currentArgv.indexOf('locales') !== -1) {
            LOCALES_FILES_PATH = currentArgv;
        } else {
            EXCEL_FILES_PATH = currentArgv;
        }
    }
}

async function main() {
    const excelFiles = await readDirP(EXCEL_FILES_PATH);
    if (excelFiles.length === 0) {
        console.error('Nothing inside the folder path');
        process.exit(1);
    }

    if (!fs.existsSync(TARGET_PATH)) mkDirP(TARGET_PATH);

    const isMergeOption = MERGE_OPTION.indexOf(OPTION.MERGE) !== -1;
    if (isMergeOption) console.log('Generating file... (w/ merge option)');
    else console.log('Generating file...');

    for (const excelFile of excelFiles) {
        // NOTE: merge & individual have the same output --> hence split to one process at a time.
        if (isMergeOption && excelFile === '_merge.xlsx') _generateJson(excelFile, _updateMergeJsonData);
        else if (!isMergeOption && excelFile !== '_merge.xlsx') _generateJson(excelFile, _updateNonmergeJsonData);
    }
}
main();

async function _generateJson(excelFile, updateJsonDataCallback) {
    const excelFilePath = path.join(EXCEL_FILES_PATH, excelFile);
    const workBook = xlsx.readFile(excelFilePath);

    await Promise.all(
        workBook.SheetNames.map((name) => {
            const langFolderPath = path.join(TARGET_PATH, name);
            if (!fs.existsSync(langFolderPath)) mkDirP(langFolderPath);

            const sheet = workBook.Sheets[name];
            const sheetJson = xlsx.utils.sheet_to_json(sheet);

            updateJsonDataCallback(sheetJson, name, excelFile);
        })
    ).catch((reason) => {
        throw new Error(reason);
    });
}

// NOTE: convert _merge.xlsx
async function _updateMergeJsonData(sheetJson, name, excelFile) {
    let json = {};
    let previousKey;

    let isLocaleFileExist = true;
    for (const sheetObj of sheetJson) {
        const { Key, English, Translation } = sheetObj;
        if (!Key) _yellowLog('Missing key in excel sheet.');

        const splittedKey = Key.split(':');
        const jsonFileName = splittedKey[0];
        const objectKeys = splittedKey.pop().split('.');

        const _readLocaleFile = async () => {
            if (LOCALES_FILES_PATH) {
                const basePath = path.join(LOCALES_FILES_PATH, name);
                const localePath = path.join(basePath, `${path.basename(jsonFileName, '.xlsx')}.json`);

                // means new language translation --> no need check locales --> hence true
                if (!fs.existsSync(basePath)) {
                    isLocaleFileExist = true;
                } else if (!fs.existsSync(localePath)) {
                    _redLog(`File not exist. --> ${localePath}`);

                    isLocaleFileExist = false;
                } else {
                    json = await readFileP(localePath);
                    json = JSON.parse(json);

                    isLocaleFileExist = true;
                }
            }
        };

        const _writeFile = () => {
            const outputFileName = path.join(TARGET_PATH, name, `${previousKey}.json`);
            const data = JSON.stringify(json, undefined, 4);

            _greenLog(`Translated json file generated. --> ${outputFileName}`);
            writeFileP(outputFileName, data);
        };

        if (!previousKey) {
            await _readLocaleFile();
            previousKey = jsonFileName;
        }
        if (previousKey !== jsonFileName) {
            // write previous key data
            if (isLocaleFileExist) _writeFile();

            // read current key data
            json = {};
            await _readLocaleFile();

            previousKey = jsonFileName;
        }

        if (!isLocaleFileExist) continue;

        const _translation = name === 'en' ? English : Translation;
        if (_translation) _createNestedObject(json, objectKeys, _translation);
        if (sheetJson.lastIndexOf(sheetObj) === sheetJson.length - 1) _writeFile();
    }
}

// NOTE: convert individual xlsx
async function _updateNonmergeJsonData(sheetJson, name, excelFile) {
    let json = {};

    if (LOCALES_FILES_PATH) {
        const basePath = path.join(LOCALES_FILES_PATH, name);
        const localePath = path.join(basePath, `${path.basename(excelFile, '.xlsx')}.json`);

        // means new language translation --> no need check locales
        if (!fs.existsSync(basePath)) {
            // do nothing
        } else if (!fs.existsSync(localePath)) {
            _redLog(`File not exist. --> ${localePath}`);
            return;
        } else {
            json = await readFileP(localePath);
            json = JSON.parse(json);
        }
    }

    for (const sheetObj of sheetJson) {
        const { Key, English, Translation } = sheetObj;
        if (!Key) _yellowLog('Missing key in excel sheet.');

        const objectKeys = Key.split(':').pop().split('.');
        const _translation = name === 'en' ? English : Translation;
        if (_translation) _createNestedObject(json, objectKeys, _translation);
    }

    const outputFileName = path.join(TARGET_PATH, name, `${path.basename(excelFile, '.xlsx')}.json`);
    const data = JSON.stringify(json, undefined, 4);

    _greenLog(`Translated json file generated. --> ${outputFileName}`);
    writeFileP(outputFileName, data);
}

function _createNestedObject(base, keys, value) {
    for (let i = 0, l = keys.length; i < l; i++) {
        if (i === l - 1) base = base[keys[i]] = value ? value : '';
        else base = base[keys[i]] = base[keys[i]] || {};
    }
}

function _redLog(string) {
    console.log(chalk.red(string));
}

function _yellowLog(string) {
    console.log(chalk.yellow(string));
}

function _greenLog(string) {
    console.log(chalk.green(string));
}

#!/usr/bin/env node

const fs = require('fs');
const util = require('util');
const readFileP = util.promisify(fs.readFile);
const readDirP = util.promisify(fs.readdir);
const mkDirP = util.promisify(fs.mkdir);
const path = require('path');
const xlsx = require('xlsx');

const TARGET_PATH = process.argv[2]; // node --> script --> params
const OUTPUT_PATH = 'json_output';

let BASE_EN_OBJ;

const MERGE_DATA = { en: [] };
const MERGE_OUTPUT_FILENAME = '_merge.json';

async function main() {
    // NOTE: checking
    const targetFiles = await readDirP(TARGET_PATH);
    if (targetFiles.length === 0) {
        console.error('Nothing inside the folder path');
        process.exit(1);
    }

    if (!fs.existsSync(OUTPUT_PATH)) {
        mkDirP(OUTPUT_PATH);
    }

    // NOTE: start processing
    const isDir = targetFiles.every((e) => e.split('.').length === 1);
    if (!isDir) {
        console.log('Generating base file...');
        _generateExcel(targetFiles);
        _generateExcelMerge(targetFiles);
    } else {
        /**
         * NOTE:
         * PRO: no need to retranslate
         * CON: object key might not sync with 'en', resulting missing translation. (Create tool to cross check, DON'T DO HERE.)
         */
        console.log('Generating translated file...');
        const multiSheetConfig = await _getMultiSheetConfig(targetFiles);
        _generateExcelMultiSheet(multiSheetConfig);
        _generateExcelMultiSheetMerge(multiSheetConfig);
    }
}
main();

// ------------------------------ ORIGINAL ------------------------------
// NOTE: one book one sheet
async function _generateExcel(jsonFiles) {
    for (const jsonFile of jsonFiles) {
        if (jsonFile === '.json') continue;

        const excel = xlsx.utils.book_new();

        // NOTE: default en, BA gonna duplicate sheet from 'en' & rename to example 'zh' for translation
        const jsonFilePath = path.join(TARGET_PATH, jsonFile);
        const sheetData = await _getSheetData(jsonFilePath, jsonFile);
        _generateSheet(excel, sheetData);

        /**
         * NOTE: output must target to another dir due to read dir for json purpose
         * files in folder must stay original to return same output & prevent error
         */
        _xlsxWriteFile(excel, jsonFile);
    }
}

// ------------------------------ ORIGINAL MERGE ------------------------------
// NOTE: merge all json file into one book one sheet
async function _generateExcelMerge(jsonFiles) {
    const excel = xlsx.utils.book_new();

    const baseLang = 'en';
    for (const jsonFile of jsonFiles) {
        if (jsonFile === '.json') continue;

        const jsonFilePath = path.join(TARGET_PATH, jsonFile);
        const sheetData = await _getSheetData(jsonFilePath, jsonFile);
        MERGE_DATA[baseLang].push(...sheetData);
    }
    _generateSheet(excel, MERGE_DATA[baseLang]);

    _xlsxWriteFile(excel, MERGE_OUTPUT_FILENAME);
}

// ------------------------------ MULTI LANG SHEET ------------------------------
async function _getMultiSheetConfig(targetFiles) {
    let config = {};
    for (let i = 0, l = targetFiles.length; i < l; i++) {
        const file = targetFiles[i];
        const langPath = path.join(TARGET_PATH, file);
        const jsonFiles = await readDirP(langPath);

        for (const jsonFile of jsonFiles) {
            if (jsonFile === '.json') continue;

            if (!config[jsonFile]) {
                config[jsonFile] = [];
            }

            // NOTE: en must be the 1st one, to store base translation.
            if (file === 'en') {
                config[jsonFile].unshift(file);
            } else {
                config[jsonFile].push(file);
            }
        }
    }
    return config;
}

// NOTE: one book multi sheet
// NOTE: for const file :: langs { agent: [en, zh, th...] }
async function _generateExcelMultiSheet(multiSheetConfig) {
    const jsonFiles = Object.keys(multiSheetConfig);
    for (const jsonFile of jsonFiles) {
        const excel = xlsx.utils.book_new();

        const jsonLangList = multiSheetConfig[jsonFile];
        for (const jsonLang of jsonLangList) {
            const sheetName = jsonLang;
            const jsonFilePath = path.join(TARGET_PATH, jsonLang, jsonFile);
            const sheetData = await _getSheetData(jsonFilePath, jsonFile, sheetName);
            _generateSheet(excel, sheetData, sheetName);
        }

        _xlsxWriteFile(excel, jsonFile);
    }
}

// ------------------------------ MULTI LANG SHEET MERGE ------------------------------
// NOTE: merge all json file into one book multi sheet
async function _generateExcelMultiSheetMerge(multiSheetConfig) {
    const excel = xlsx.utils.book_new();

    const jsonFiles = Object.keys(multiSheetConfig);
    for (const jsonFile of jsonFiles) {
        const jsonLangList = multiSheetConfig[jsonFile];
        for (const jsonLang of jsonLangList) {
            const sheetName = jsonLang;
            const jsonFilePath = path.join(TARGET_PATH, jsonLang, jsonFile);

            const sheetData = await _getSheetData(jsonFilePath, jsonFile, sheetName);
            if (!MERGE_DATA[jsonLang]) MERGE_DATA[jsonLang] = [];
            MERGE_DATA[jsonLang].push(...sheetData);
        }
    }

    const mergeDataKeyList = Object.keys(MERGE_DATA);
    for (const key of mergeDataKeyList) {
        _generateSheet(excel, MERGE_DATA[key], key);
    }

    _xlsxWriteFile(excel, MERGE_OUTPUT_FILENAME);
}

// ------------------------------ UTIL ------------------------------
// NOTE: modify sheet data here
async function _getSheetData(file, jsonFile, sheetName = 'en') {
    const content = await readFileP(file, 'utf-8');
    const json = JSON.parse(content);
    const data = [];

    const isEN = sheetName === 'en';
    if (isEN) BASE_EN_OBJ = json;

    const uniqueFilename = path.basename(jsonFile, '.json');
    const iterateObject = (baseObject, object, previousKey) => {
        for (const key in object) {
            const obj = object[key];
            const baseObj = baseObject && baseObject[key];
            if (obj) {
                if (typeof obj === 'object') {
                    const storedKey = previousKey ? previousKey + '.' + key : key;
                    iterateObject(baseObj, obj, storedKey);
                } else {
                    // NOTE: en = base, else export to translated.
                    let dataObj;
                    const previousKeyBridge = previousKey ? previousKey + '.' : ''; // NOTE: some json straight is string without nested object.
                    const keyData = uniqueFilename + ':' + previousKeyBridge + key;
                    if (isEN) {
                        dataObj = {
                            Key: keyData,
                            English: obj,
                        };
                    } else {
                        dataObj = {
                            Key: keyData,
                            English: baseObj,
                            Translation: obj,
                        };
                    }
                    data.push(dataObj);
                }
            }
        }
    };
    iterateObject(BASE_EN_OBJ, json);

    return data;
}

async function _generateSheet(excel, sheetData, sheetName = 'en') {
    const sheet = xlsx.utils.json_to_sheet(sheetData, {
        header: ['Key', 'English', 'Translation'],
    });
    xlsx.utils.book_append_sheet(excel, sheet, sheetName);
}

function _xlsxWriteFile(excel, jsonFile) {
    const outputFileName = path.basename(jsonFile, '.json') + '.xlsx';
    const output = path.join(OUTPUT_PATH, outputFileName);
    xlsx.writeFile(excel, output);
    console.log(`Excel file generated. --> ${output}`);
}

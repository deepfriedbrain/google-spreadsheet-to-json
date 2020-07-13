
var fs = require('fs')
const { GoogleSpreadsheet } = require('google-spreadsheet');

var Promise = require('bluebird')

// constants used for converting column names into number/index
var ALPHABET = 'abcdefghijklmnopqrstuvwxyz'
var ALPHABET_BASE = ALPHABET.length

function capitalize(str) {
    return str.charAt(0).toUpperCase() + str.slice(1)
}

function getWords(phrase) {
    return phrase.replace(/[- ]/ig, ' ').split(' ')
}

// Service Account credentials are first parsed as JSON and, in case this fails,
// they are considered a file path
function parseServiceAccountCredentials(credentials) {

    if (typeof credentials === 'string') {
        try {
            return JSON.parse(credentials)
        } catch(ex) {
            return JSON.parse(fs.readFileSync(credentials, 'utf8'))
        }
    }

    return credentials
}

function handlePropertyName(cellValue, handleMode) {

    var handleModeType = typeof handleMode

    if (handleModeType === 'function')
        return handleMode(cellValue)

    var propertyName = (cellValue || '').trim()

    if (handleMode === 'camel' || handleModeType === 'undefined')
        return getWords(propertyName.toLowerCase()).map(function(word, index) {
            return !index ? word : capitalize(word)
        }).join('')

    if (handleMode === 'pascal')
        return getWords(propertyName.toLowerCase()).map(function(word) {
            return capitalize(word)
        }).join('')

    if (handleMode === 'nospace')
        return getWords(propertyName).join('')

    return propertyName
}

function handleIntValue(val) {
    return parseInt(val, 10) || 0
}

// returns a number if the string can be parsed as an integer
function handlePossibleIntValue(val) {
    if (typeof val === 'string' && /^\d+$/.test(val))
        return handleIntValue(val)
    return val
}

function normalizePossibleIntList(option, defaultValue) {
    return normalizeList(option, defaultValue).map(handlePossibleIntValue)
}

// should always return an array
function normalizeList(option, defaultValue) {

    if (typeof option === 'undefined')
        return defaultValue || []

    return Array.isArray(option) ? option : [option]
}

function setPropertyTree(object, tree, value) {

    if (!Array.isArray(tree))
        tree = [tree]

    var prop = tree[0]
    if (!prop)
        return

    object[prop] = tree.length === 1 ? value : (typeof object[prop] === 'object' ? object[prop] : {})

    setPropertyTree(object[prop], tree.slice(1), value)

}

function parseColIdentifier(col) {

    var colType = typeof col

    if (colType === 'string') {
        return col.trim().replace(/[ \.]/i, '').toLowerCase().split('').reverse().reduce(function(totalValue, letter, index) {

            var alphaIndex = ALPHABET.indexOf(letter)

            if (alphaIndex === -1)
                throw new Error('Column identifier format is invalid')

            var value = alphaIndex + 1

            return totalValue + value * Math.pow(ALPHABET_BASE, index)
        }, 0)
    }

    if (colType !== 'number')
        throw new Error('Column identifier value type is invalid')

    return col
}

function cellIsValid(cell) {
    return !!cell && typeof cell.value === 'string' && cell.value !== ''
}

// google spreadsheet cells into json
exports.cellsToJson = async function(allCells, options) {

    // setting up some options, such as defining if the data is horizontal or vertical
    options = options || {}

    // var rowProp = options.vertical ? 'col' : 'row'
    // var colProp = options.vertical ? 'row' : 'col'
    // var isHashed = options.hash && !options.listOnly
    // var includeHeaderAsValue = options.listOnly && options.includeHeader
    // var headerStartNumber = options.headerStart ? parseColIdentifier(options.headerStart) : 0
    // var headerSize = Math.min(handleIntValue(options.headerSize)) || 1
    var ignoredRows = normalizePossibleIntList(options.ignoreRow)
    var ignoredCols = normalizePossibleIntList(options.ignoreCol).map(parseColIdentifier)
    var ignoredDataNumbers = options.vertical ? ignoredRows : ignoredCols
    ignoredDataNumbers.sort().reverse()

    var rows = await allCells.getRows();
    console.log("rowCount : " + allCells.rowCount);

    let headers = allCells.headerValues;
    console.log("headers : " + headers);

    //Note: row doesn't have any formatting or formula info; cell contains all that
    //console.log(allCells.getCellByA1("A2").valueType);

    var nonEmptyRows = rows.filter((row, index) => {
        //console.log("row.id : " + row[headers[0]]);
        return row[headers[0]];
    });

    console.log("nonEmptyRows length : " + nonEmptyRows.length);

    return nonEmptyRows.map((row, index) => {
        return Object.assign(...headers.map(k => row[k] && {[k]: row[k]}));
        //return Object.assign(...headers.map(k => row[k] && {[k]: handlePossibleIntValue(row[k])}));
    })
}

exports.spreadsheetToJson = async function(options) {

    var allWorksheets = !!options.allWorksheets
    var expectMultipleWorksheets = allWorksheets || Array.isArray(options.worksheet)

    var spreadsheet = new GoogleSpreadsheet(options.spreadsheetId);

    if (options.credentials) {
        await spreadsheet.useServiceAccountAuth(parseServiceAccountCredentials(options.credentials));
    }

    await spreadsheet.loadInfo();
    console.log("spreadsheet.title : " + spreadsheet.title);

    var worksheets = spreadsheet.sheetsByIndex;
    var selectedWorksheets;

    if (!allWorksheets) {
        var identifiers = normalizePossibleIntList(options.worksheet, [0])

        selectedWorksheets = worksheets.filter(function(worksheet, index) {
            return identifiers.indexOf(index) !== -1 || identifiers.indexOf(worksheet.title) !== -1
        })

        if (!expectMultipleWorksheets) {
            selectedWorksheets = selectedWorksheets.slice(0, 1)
        }

        if (selectedWorksheets.length === 0) {
            throw new Error('No worksheet found!')
        }
    }

    var finalList = selectedWorksheets.map(async function(worksheet) {
        console.log ("worksheet.title : " + worksheet.title)
        await worksheet.loadCells();
        console.log(worksheet.cellStats);
        return exports.cellsToJson(worksheet, options);
    })

    return expectMultipleWorksheets ? finalList : finalList[0];
}

const ExcelJS = require('exceljs')

/**
 * @typedef {Object} HeaderData
 * @property {string} name
 * @property {number} width
 */

/**
 * Populates excel workbook with data from headers and data parameters.
 * Parameters can be arrays of objects, in this case each object will be used to create separate worksheet.
 * @param {Object<string, HeaderData>|[Object<string, HeaderData>]} headers mapping field name => header value and column width
 * @param {[Object]|[[Object]]} data array of objects with fields as properties of headers parameter
 * @param {string|[string]} worksheetNames
 * @throws {Error} if sized of array parameters are not equal
 * @returns {Promise<Buffer>}
 */
async function createBuffer(headers, data, worksheetNames) {
    headers = Array.isArray(headers) ? headers : [headers]
    data = Array.isArray(data) ? data : [data]
    worksheetNames = Array.isArray(worksheetNames) ? worksheetNames : [worksheetNames]

    if ((headers.length != data.length) != worksheetNames.length) {
        throw new Error('Size of headers, data and worksheetNames must be equal')
    }

    const workbook = createWorkbook()

    for (let i = 0; i < headers.length; i++) {
        const sheetHeaders = headers[i]
        const sheetData = data[i]

        const worksheet = addWorksheet(workbook)

        makeRowsBold(worksheet, [1])

        const headerValues = Object.values(sheetHeaders).map(item => item.name)
        setHeaders(worksheet, headerValues)
        const columnWidths = Object.values(sheetHeaders).map(item => item.width)
        setColumnWidths(worksheet, columnWidths)

        const fieldNames = Object.keys(sheetHeaders)

        const startDataRow = 2
        sheetData.forEach((dataItem, index) => {
            const values = []
            fieldNames.forEach(field => {
                values.push(dataItem[field])
            })
            addDataRow(worksheet, startDataRow + index, values)
        })
    }

    return await getBuffer(workbook)
}

function createWorkbook() {
    return new ExcelJS.Workbook()
}

/**
 * @param {ExcelJS.Workbook} workboork
 */
function addWorksheet(workboork, worksheetName) {
    return workboork.addWorksheet(worksheetName, {
        pageSetup: { paperSize: 9, orientation: 'landscape' },
    })
}

/**
 * Make data in specified rows bold
 * @param {import('exceljs').Worksheet} worksheet
 * @param {number|[number]} rows numbers of rows to make bold
 */
function makeRowsBold(worksheet, rows) {
    rows = Array.isArray(rows) ? rows : [rows]
    for (const rowNum of rows) {
        worksheet.getRow(rowNum).font = { bold: true }
    }
}

/**
 * Set column widths for all columns in worksheet
 * @param {import('exceljs').Worksheet} worksheet
 * @param {[number]} widths
 */
function setColumnWidths(worksheet, widths) {
    widths.forEach((width, index) => {
        const column = worksheet.getColumn(index + 1)
        column.width = width
        column.alignment = {
            vertical: 'middle',
            horizontal: 'center',
            wrapText: true,
        }
    })
}

/**
 * Set widths for specified columns
 * @param {import('exceljs').Worksheet} worksheet
 * @param {[{columns: [string], width: number}]} widths columns to set width for
 */
function setPartialColumnWidths(worksheet, widths) {
    widths.forEach(item => {
        for (const colName of item.columns) {
            column = worksheet.getColumn(colName)
            column.width = item.width
        }
    })
}

/**
 * Adds headers to the first row of the worksheet
 * @param {import('exceljs').Worksheet} worksheet
 * @param {[string]} headers
 */
function setHeaders(worksheet, headers) {
    headers.forEach((header, index) => {
        worksheet.getCell(String.fromCharCode(65 + index) + '1').value = header
    })
}

/**
 * Merge cells in worksheet
 * @param {import('exceljs').Worksheet} worksheet
 * @param {string} cellFrom name of the first cell in the merged range
 * @param {string} cellTo name of the last cell in the merged range
 * @param {string} value value to put in the merged cell
 */
function mergeCells(worksheet, cellFrom, cellTo, value) {
    worksheet.mergeCells(`${cellFrom}:${cellTo}`)
    const mergedCell = worksheet.getCell(cellFrom)
    mergedCell.value = value
}

/**
 * Create merged header cells in worksheet
 * @param {import('exceljs').Worksheet} worksheet
 * @param {[{from: string?, to: string?, cell: string?, value: string}]} headers
 */
function setComplexHeaders(worksheet, headers) {
    for (const header of headers) {
        if (header.cell) {
            worksheet.getCell(header.cell).value = header.value
        } else if (header.from && header.to) {
            mergeCells(worksheet, header.from, header.to, header.value)
        } else {
            throw new Error('Either cell or from & to fields must be present in headers array objects:', header)
        }
    }
}

/**
 * Add data row to worksheet starting from rowIndex from data array consequently to columns starting from A
 * @param {import('exceljs').Worksheet} worksheet
 * @param {number} rowIndex
 * @param {[string]} data
 */
function addDataRow(worksheet, rowIndex, data) {
    const row = worksheet.getRow(rowIndex)
    data.forEach((item, index) => {
        row.getCell(String.fromCharCode(65 + index)).value = item
    })
}

/**
 * Set alignment for all columns in worksheet
 * @param {import('exceljs').Worksheet} worksheet
 * @param {{vertical: string?, horizontal: string?}?} specialCells rules for special alignment for some columns. Object with column name (letter) as key and alignment object as value
 */
function alignColumns(worksheet, specialCells = {}) {
    const totalColumns = worksheet.columnCount
    for (let columnIndex = 1; columnIndex <= totalColumns; columnIndex++) {
        const columnLetter = getExcelColumnLetter(columnIndex)
        const specialAlignment = specialCells[columnLetter]
        const column = worksheet.getColumn(columnLetter)
        column.alignment = specialAlignment || {
            vertical: 'middle',
            horizontal: 'center',
        }
    }
}

/**
 * Set alignment for specified rows in worksheet
 * @param {import('exceljs').Worksheet} worksheet
 * @param {Object} alignment
 * @param {number?} rowTo number of last row to align. If not passed, all rows will be aligned
 */
function alignRows(worksheet, alignment, rowTo) {
    rowTo = rowTo || worksheet.rowCount
    for (let rowNumber = 1; rowNumber <= rowTo; rowNumber++) {
        const row = worksheet.getRow(rowNumber)
        row.alignment = alignment
    }
}

/**
 * Convert column index to Excel column letter
 * @param {number} colIndex - column index (1-based)
 * @returns {string} column letter
 */
function getExcelColumnLetter(colIndex) {
    let letter = ''
    while (colIndex > 0) {
        const mod = (colIndex - 1) % 26
        letter = String.fromCharCode(65 + mod) + letter
        colIndex = Math.floor((colIndex - 1) / 26)
    }
    return letter
}

/**
 * Create buffer from workbook
 * @param {ExcelJS.Workbook} workbook
 */
async function getBuffer(workbook) {
    return await workbook.xlsx.writeBuffer()
}

module.exports = {
    createBuffer,
    createWorkbook,
    addWorksheet,
    makeRowsBold,
    setColumnWidths,
    setPartialColumnWidths,
    setHeaders,
    mergeCells,
    setComplexHeaders,
    addDataRow,
    alignColumns,
    alignRows,
    getExcelColumnLetter,
    getBuffer,
}

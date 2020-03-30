const readXlsxFile = require('read-excel-file/node')
const gettextParser = require('gettext-parser')
const fs = require('fs')
const path = require('path')

const args = process.argv.slice(2)
if (args.length < 1) {
    console.log('Please indicate an Excel file.')
    console.log('Example: npm run import my-excel-file.xlsx')
    process.exit(1)
}

// Get a header for a PO file.
function getHeaders() {
    const headers = {
        'MIME-Version': '1.0',
        'Content-Type': 'text/plain; charset=UTF-8',
        'Content-Transfer-Encoding': '8bit'
    }
    return headers
}

// Generate translatable units.
function getUnit(source, context) {
    if (typeof string !== 'string') {
        source = String(source)
    }
    return {
        msgctxt: context,
        msgid: source.replace(/\r\n/g, "\n"),
        msgstr: '',
    }
}

// Helper function to a PO file.
function writePo(indicatorId, units) {
    const data = {
        headers: getHeaders(),
        translations: {
            "": units
        }
    }
    const fileData = gettextParser.po.compile(data)
    indicatorId = dotsToDashes(indicatorId)
    const filePath = path.join('translations', 'templates', indicatorId + '.pot')
    fs.writeFileSync(filePath, fileData)
}

function dotsToDashes(indicatorId) {
    return indicatorId.replace(/\./g, '-')
}

function printError(sheetName, message) {
    console.log(`ERROR in sheet "${sheetName}"`)
    console.log(message)
    console.log('')
}

function getConceptId(row) {
    return row[0]
}

function getConceptValue(row) {
    return row[3]
}

async function main() {

    const excelFilePath = path.join(process.env.INIT_CWD, args[0])
    const sheets = await readXlsxFile(excelFilePath, { getSheets: true })
    const instructions = sheets.shift()
    const units = {}

    for (const sheet of sheets) {
        const sheetName = sheet.name

        const rows = await readXlsxFile(excelFilePath, { sheet: sheetName })

        const mainConcept = rows[rows.findIndex(row => row[0] === 'Main Concept') + 2]
        const detailedStart = rows.findIndex(row => row[0] === 'Detailed Concepts')
        const detailedConcepts = (detailedStart === -1) ? [] : rows.slice(detailedStart + 2)

        const mainConceptValue = getConceptValue(mainConcept)
        const mainConceptEntered = mainConceptValue !== null
        const detailedConceptsEntered = detailedConcepts.some(row => getConceptValue(row))

        if (mainConceptEntered && detailedConceptsEntered) {
            const message = 'Please complete the Main Concept OR the Detailed Concepts, but not both.'
            printError(sheetName, message)
            continue
        }

        if (!mainConceptEntered && !detailedConceptsEntered) {
            const message = 'Please complete the Main Concept OR the Detailed Concepts.'
            printError(sheetName, message)
            continue
        }

        if (mainConceptEntered) {
            const conceptId = getConceptId(mainConcept)
            units[conceptId] = getUnit(mainConceptValue, conceptId)
        }
        else {
            for (const concept in detailedConcepts) {
                const conceptId = getConceptId(concept)
                units[conceptId] = getUnit(getConceptValue(concept), conceptId)
            }
        }
    }

    // For now we assume that the Excel file name is the indicator id. This
    // logic will likely be changed as the template is refined.
    const baseName = path.basename(excelFilePath).split('.').slice(0, -1).join('.')
    const indicatorId = dotsToDashes(baseName)
    if (indicatorId.split('-').length < 3) {
        console.log(`ERROR in file "${baseName}"`)
        console.log('Unable to convert filename to indicator ID')
        return
    }

    writePo(indicatorId, units)
}

main().catch(e => {
    console.error(e)
})

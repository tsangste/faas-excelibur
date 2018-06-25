const fs = require('fs')
const createWorkbook = require('./create-workbook')

function WriteWorkbook(options) {
    let workbook
    try {
        workbook = createWorkbook(options)
    } catch (err) {        
        return Promise.reject(err.message)
    }
    
    const workbookFile = workbook[options.fileExtension]

    if (!workbookFile) {
        return Promise.reject('Cannot create an excel file for this file extension')
    }

    const writeStream = fs.createWriteStream('spreadsheet')

    return new Promise((resolve, reject) => workbookFile.write(writeStream, data => resolve(data)))
}

module.exports = WriteWorkbook
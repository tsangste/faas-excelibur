const createWorkbook = require('./create-workbook')
const unstream = require('unstream')

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

    return new Promise((resolve, reject) => workbookFile.write(unstream({}, data => resolve(data))))
}

module.exports = WriteWorkbook
module.exports = (payload) => {
    let errors = []

    if (!payload.fileExtension) {
        errors.push('A file extension must be specified')
    } else if (payload.fileExtension !== 'xlsx' && payload.fileExtension !== 'csv') {
        errors.push('File extension must be either csv or xlsx')
    }

    if (!payload.workbook) {
        errors.push('No workbook detected')
        return errors
    }

    let workbook = payload.workbook
    if (!workbook.worksheets || workbook.worksheets.length === 0) {
        errors.push('At least 1 worksheet must be attached to a workbook')
    }

    return errors
}
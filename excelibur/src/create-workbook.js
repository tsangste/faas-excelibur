const Excel = require('exceljs')
const dimensionCalculator = require('./image-cell-dimension-calc')

function CreateWorkbook(options) {
    let workbook = new Excel.Workbook()
    let worksheets = options.workbook.worksheets

    worksheets.forEach(ws => {
        let worksheet = workbook.addWorksheet(ws.title)

        // Assigning doesn't seem to work and objects don't copy across right
        // so for now, manually assign so we have something working
        if (ws.cols) {
            worksheet.columns = Array(ws.cols.length).fill({})
            ws.cols.forEach((colDefinition, colNum) => {
                let column = worksheet.getColumn(colNum + 1)

                Object.keys(colDefinition).forEach(key => {
                    column[key] = colDefinition[key]
                })
            })
        }

        if (ws.rows) {
            ws.rows.forEach(rowDefinition => {
                let row = worksheet.addRow()
                Object.assign(row, rowDefinition)

                if (rowDefinition.cells) {
                    rowDefinition.cells.forEach((cellDefinition, colNum) => {
                        let cell = row.getCell(colNum + 1)
                        Object.assign(cell, cellDefinition)
                    })
                }

                row.commit()
            })
        }

        // Add the images after the columns so we know the widths
        if (ws.images) {
            ws.images.forEach(img => {
                addImage(workbook, worksheet, img)
            })
        }

        if (ws.mergeCells) {
            ws.mergeCells.forEach(merge => {
                worksheet.mergeCells(merge.top, merge.left, merge.bottom, merge.right)
            })
        }
    })

    return workbook;
}

function addImage(workbook, worksheet, image) {
    let imageBuffer = Buffer.from(image.data)
    let imageCellDimensions = dimensionCalculator(imageBuffer, worksheet, image.anchor, image.cellHeight)

    let imageId = workbook.addImage({
        buffer: imageBuffer,
        extension: imageCellDimensions.type
    })

    worksheet.addImage(imageId, imageCellDimensions)
}

module.exports = CreateWorkbook
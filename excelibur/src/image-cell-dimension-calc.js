const sizeOf = require('image-size')

const excelUnitToPixelHeightRatio = 0.75
const excelUnitToPixelWidthRatio = 0.12

function ImageCellDimensionCalculator(buffer, worksheet, anchor = {col: 0, row: 0},  imageHeight = 10) {
    const excelCellHeightInUnits = 15
    let dimensions = sizeOf(buffer)

    let imageHeightInUnits = imageHeight * excelCellHeightInUnits

    let imageHeightInPixels = imageHeightInUnits / excelUnitToPixelHeightRatio
    let imageWidthInPixels = imageHeightInPixels * dimensions.width / dimensions.height

    let imageWidthInUnits = imageWidthInPixels * excelUnitToPixelWidthRatio

    // Calculate how many columns span the required width
    let totalColSpan, totalColWidth = 0
    for (let i = 0; i < 100; i++) {
        let colWidth = 8.38
        if (worksheet.actualColumnCount > i) {
            colWidth = worksheet.columns[i].width
        }

        // If the next cell contains more than enough space, calculate what 
        // percentage of that cell is necessary to contain the rest of the image
        if (totalColWidth + colWidth >= imageWidthInUnits) {
            let widthToCover = imageWidthInUnits - totalColWidth
            let colWidthPercentage = 1 - (colWidth - widthToCover) / ((colWidth + widthToCover) / 2)
            
            totalColSpan = i + colWidthPercentage
            break;
        }

        totalColWidth += colWidth
    }

    return {
        tl: { col: anchor.col, row: anchor.row },
        // Knock off 0.05 to allow the header border to be visible
        br: { col: anchor.col + totalColSpan, row: anchor.row + imageHeight - 0.05 },
        type: dimensions.type
    }
}

module.exports = ImageCellDimensionCalculator
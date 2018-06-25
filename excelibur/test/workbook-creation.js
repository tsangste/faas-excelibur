const createWorkbook = require('../src/create-workbook')

describe('Workbook Creation', function () {
    it('should create a worksheet with the correct title', function () {
        let options = {
            "fileExtension": "xlsx",
            "workbook": {
                "worksheets": [
                    {
                        "title": "My First Worksheet"
                    }
                ]
            }
        }

        let workbook = createWorkbook(options)
        workbook.worksheets.should.have.lengthOf(1)
        workbook.getWorksheet(1).name.should.equal('My First Worksheet')
    })

    it('should create multiple worksheets with the correct title', function () {
        let options = {
            "fileExtension": "xlsx",
            "workbook": {
                "worksheets": [
                    {
                        "title": "My First Worksheet"
                    },
                    {
                        "title": "My Second Worksheet"
                    }
                ]
            }
        }

        let workbook = createWorkbook(options)
        workbook.worksheets.should.have.lengthOf(2)
        workbook.getWorksheet(1).name.should.equal('My First Worksheet')
        workbook.getWorksheet(2).name.should.equal('My Second Worksheet')
    })
})
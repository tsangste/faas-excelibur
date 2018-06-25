const createWorkbook = require('../src/create-workbook')

describe('Workbook Cell Values', function () {
    it('should create the correct number of empty cells', function () {
        let options = {
            "fileExtension": "xlsx",
            "workbook": {
                "worksheets": [
                    {
                        "title": "My First Worksheet",
                        "rows": [
                            {
                                "cells": [
                                    // Values are required otherwise exceljs ignores entirely
                                    { 'value': 'Test'},
                                    { 'value': 'Test'},
                                    { 'value': 'Test'},
                                    { 'value': 'Test'},
                                    { 'value': 'Test'},
                                    { 'value': 'Test'},
                                ]
                            }
                        ]
                    }
                ]
            }
        }

        let workbook = createWorkbook(options)
        let worksheet = workbook.getWorksheet(1)
        let row = worksheet.getRow(1)
        
        row.actualCellCount.should.equal(6)
        row.getCell(1).value.should.equal('Test')
        row.getCell(2).value.should.equal('Test')
        row.getCell(3).value.should.equal('Test')
        row.getCell(4).value.should.equal('Test')
        row.getCell(5).value.should.equal('Test')
        row.getCell(6).value.should.equal('Test')

        worksheet.getCell('A1').value.should.equal('Test')
    })
})
const createWorkbook = require('../src/create-workbook')

describe('Workbook Cell Creation', function () {
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
    })

    it('should set the correct properties on each row', function () {
        let options = {
            "fileExtension": "xlsx",
            "workbook": {
                "worksheets": [
                    {
                        "title": "My First Worksheet",
                        "rows": [
                            {
                                "cells": [
                                    {
                                        "numFmt": '"£"#,##0.00;[Red]\-"£"#,##0.00',
                                        'value' : 'Test'
                                        
                                    },
                                    {
                                        "numFmt": '"£"#,##0.00;[Blue]\-"£"#,##0.00',
                                        'value' : 'Test'
                                    }
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

        row.actualCellCount.should.equal(2)
        row.getCell(1).numFmt.should.equal('"£"#,##0.00;[Red]\-"£"#,##0.00')
        row.getCell(2).numFmt.should.equal('"£"#,##0.00;[Blue]\-"£"#,##0.00')
    })

    it('should set the font', function () {
        let options = {
            "fileExtension": "xlsx",
            "workbook": {
                "worksheets": [
                    {
                        "title": "My First Worksheet",
                        "rows": [
                            {
                                "cells": [
                                    {
                                        "font": {
                                            'name': 'Calibri',
                                            'size': 18,
                                            "color": {
                                                "argb": "FFAAAAAA"
                                            }
                                        }
                                    }
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
        let cell = row.getCell(1)

        cell.style.should.have.key('font')
        cell.style.font.name.should.equal('Calibri')
        cell.style.font.size.should.equal(18)
        cell.style.font.color.should.have.key('argb')
        cell.style.font.color.argb.should.equal('FFAAAAAA')
    })

    it('should set the alignment', function () {
        let options = {
            "fileExtension": "xlsx",
            "workbook": {
                "worksheets": [
                    {
                        "title": "My First Worksheet",
                        "rows": [
                            {
                                "cells": [
                                    {
                                        "alignment": {
                                            'vertical': 'middle',
                                            'horizontal': 'center',
                                            "wrapText": true
                                        }
                                    }
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
        let cell = row.getCell(1)

        cell.style.should.have.key('alignment')
        cell.style.alignment.vertical.should.equal('middle')
        cell.style.alignment.horizontal.should.equal('center')
        cell.style.alignment.wrapText.should.be.true
    })

    it('should set the border', function () {
        let options = {
            "fileExtension": "xlsx",
            "workbook": {
                "worksheets": [
                    {
                        "title": "My First Worksheet",
                        "rows": [
                            {
                                "cells": [
                                    {
                                        "border": {
                                            'top': { 'style': 'medium', 'color': { 'argb': 'AAA3A3A3' } },
                                            'left': { 'style': 'medium', 'color': { 'argb': 'AAA3A3A3' } },
                                            'bottom': { 'style': 'medium', 'color': { 'argb': 'AAA3A3A3' } },
                                            'right': { 'style': 'medium', 'color': { 'argb': 'AAA3A3A3' } },
                                        },
                                    }
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
        let cell = row.getCell(1)

        cell.style.should.have.key('border')
        cell.style.border.should.have.keys(['top', 'left', 'bottom', 'right'])

        cell.style.border.top.should.have.keys(['style', 'color'])
        cell.style.border.top.style.should.equal('medium')
        cell.style.border.top.color.should.have.key('argb')
        cell.style.border.top.color.argb.should.equal('AAA3A3A3')

        cell.style.border.left.should.have.keys(['style', 'color'])
        cell.style.border.left.style.should.equal('medium')
        cell.style.border.left.color.should.have.key('argb')
        cell.style.border.left.color.argb.should.equal('AAA3A3A3')

        cell.style.border.bottom.should.have.keys(['style', 'color'])
        cell.style.border.bottom.style.should.equal('medium')
        cell.style.border.bottom.color.should.have.key('argb')
        cell.style.border.bottom.color.argb.should.equal('AAA3A3A3')

        cell.style.border.right.should.have.keys(['style', 'color'])
        cell.style.border.right.style.should.equal('medium')
        cell.style.border.right.color.should.have.key('argb')
        cell.style.border.right.color.argb.should.equal('AAA3A3A3')
    })

    it('should set the fill', function () {
        let options = {
            "fileExtension": "xlsx",
            "workbook": {
                "worksheets": [
                    {
                        "title": "My First Worksheet",
                        "rows": [
                            {
                                "cells": [
                                    {
                                        "fill": {
                                            'type': 'pattern',
                                            'pattern': 'solid',
                                            "fgColor": {
                                                "argb": 'FFFF0000'
                                            }
                                        }
                                    }
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
        let cell = row.getCell(1)

        cell.style.should.have.key('fill')
        cell.style.fill.type.should.equal('pattern')
        cell.style.fill.pattern.should.equal('solid')
        cell.style.fill.fgColor.should.contain.key('argb')
        cell.style.fill.fgColor.argb.should.equal('FFFF0000')
    })
})
const createWorkbook = require('../src/create-workbook')

describe('Workbook Row Creation', function() {
    it('should create the correct number of empty rows', function() {
        let options = {
            "fileExtension": "xlsx",
            "workbook": {
                "worksheets": [
                    {
                        "title": "My First Worksheet",
                        "rows": [ {}, {}, {}, {} ]
                    }
                ]
            }
        }

        let workbook = createWorkbook(options)
        let worksheet = workbook.getWorksheet(1)

        worksheet.rowCount.should.equal(4)
    })

    it('should set the correct properties on each row', function () {
        let options = {
            "fileExtension": "xlsx",
            "workbook": {
                "worksheets": [
                    {
                        "title": "My First Worksheet",
                        "rows": [
                            { "height": 20 }, { "height": 40 }, { "height": 60 }
                        ]
                    }
                ]
            }
        }

        let workbook = createWorkbook(options)
        let worksheet = workbook.getWorksheet(1)

        worksheet.rowCount.should.equal(3)
        worksheet.getRow(1).height.should.equal(20)
        worksheet.getRow(2).height.should.equal(40)
        worksheet.getRow(3).height.should.equal(60)
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
        }

        let workbook = createWorkbook(options)
        let worksheet = workbook.getWorksheet(1)
        let row = worksheet.getRow(1)
        
        row.style.should.have.key('font')
        row.style.font.name.should.equal('Calibri')
        row.style.font.size.should.equal(18)    
        row.style.font.color.should.have.key('argb')
        row.style.font.color.argb.should.equal('FFAAAAAA')
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
        }

        let workbook = createWorkbook(options)
        let worksheet = workbook.getWorksheet(1)
        let row = worksheet.getRow(1)
        
        row.style.should.have.key('alignment')
        row.style.alignment.vertical.should.equal('middle')
        row.style.alignment.horizontal.should.equal('center')
        row.style.alignment.wrapText.should.be.true
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
        }

        let workbook = createWorkbook(options)
        let worksheet = workbook.getWorksheet(1)
        let row = worksheet.getRow(1)
        
        row.style.should.have.key('border')
        row.style.border.should.have.keys(['top', 'left', 'bottom', 'right'])

        row.style.border.top.should.have.keys(['style', 'color'])
        row.style.border.top.style.should.equal('medium')
        row.style.border.top.color.should.have.key('argb')
        row.style.border.top.color.argb.should.equal('AAA3A3A3')
       
        row.style.border.left.should.have.keys(['style', 'color'])
        row.style.border.left.style.should.equal('medium')
        row.style.border.left.color.should.have.key('argb')
        row.style.border.left.color.argb.should.equal('AAA3A3A3')

        row.style.border.bottom.should.have.keys(['style', 'color'])
        row.style.border.bottom.style.should.equal('medium')
        row.style.border.bottom.color.should.have.key('argb')
        row.style.border.bottom.color.argb.should.equal('AAA3A3A3')

        row.style.border.right.should.have.keys(['style', 'color'])
        row.style.border.right.style.should.equal('medium')
        row.style.border.right.color.should.have.key('argb')
        row.style.border.right.color.argb.should.equal('AAA3A3A3')
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
        }

        let workbook = createWorkbook(options)
        let worksheet = workbook.getWorksheet(1)
        let row = worksheet.getRow(1)
        
        row.style.should.have.key('fill')
        row.style.fill.type.should.equal('pattern')
        row.style.fill.pattern.should.equal('solid')
        row.style.fill.fgColor.should.contain.key('argb')
        row.style.fill.fgColor.argb.should.equal('FFFF0000')
    })
})
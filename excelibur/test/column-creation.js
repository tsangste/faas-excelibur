const createWorkbook = require('../src/create-workbook')

describe('Workbook Column Creation', function () {
    it('should create the correct number of empty columns', function () {
        let options = {
            "fileExtension": "xlsx",
            "workbook": {
                "worksheets": [
                    {
                        "title": "My First Worksheet",
                        "cols": [
                            {}, {}, {}, {}, {}
                        ]
                    }
                ]
            }
        }

        let workbook = createWorkbook(options)
        let worksheet = workbook.getWorksheet(1)

        // Can't use columnCount here as they don't have values
        // we just want to ensure something is there
        worksheet.columns.should.have.lengthOf(5)
    })

    it('should set the correct properties on each column', function () {
        let options = {
            "fileExtension": "xlsx",
            "workbook": {
                "worksheets": [
                    {
                        "title": "My First Worksheet",
                        "cols": [
                            { "width": 20 }, { "width": 40 }, { "width": 60 }
                        ]
                    }
                ]
            }
        }

        let workbook = createWorkbook(options)
        let worksheet = workbook.getWorksheet(1)

        worksheet.columns.should.have.lengthOf(3)
        worksheet.getColumn(1).width.should.equal(20)
        worksheet.getColumn(2).width.should.equal(40)
        worksheet.getColumn(3).width.should.equal(60)
    })

    it('should set the font', function () {
        let options = {
            "fileExtension": "xlsx",
            "workbook": {
                "worksheets": [
                    {
                        "title": "My First Worksheet",
                        "cols": [
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
        let column = worksheet.getColumn(1)
        
        column.style.should.have.key('font')
        column.style.font.name.should.equal('Calibri')
        column.style.font.size.should.equal(18)    
        column.style.font.color.should.have.key('argb')
        column.style.font.color.argb.should.equal('FFAAAAAA')
    })

    it('should set the alignment', function () {
        let options = {
            "fileExtension": "xlsx",
            "workbook": {
                "worksheets": [
                    {
                        "title": "My First Worksheet",
                        "cols": [
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
        let column = worksheet.getColumn(1)
        
        column.style.should.have.key('alignment')
        column.style.alignment.vertical.should.equal('middle')
        column.style.alignment.horizontal.should.equal('center')
        column.style.alignment.wrapText.should.be.true
    })

    it('should set the border', function () {
        let options = {
            "fileExtension": "xlsx",
            "workbook": {
                "worksheets": [
                    {
                        "title": "My First Worksheet",
                        "cols": [
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
        let column = worksheet.getColumn(1)
        
        column.style.should.have.key('border')
        column.style.border.should.have.keys(['top', 'left', 'bottom', 'right'])

        column.style.border.top.should.have.keys(['style', 'color'])
        column.style.border.top.style.should.equal('medium')
        column.style.border.top.color.should.have.key('argb')
        column.style.border.top.color.argb.should.equal('AAA3A3A3')
       
        column.style.border.left.should.have.keys(['style', 'color'])
        column.style.border.left.style.should.equal('medium')
        column.style.border.left.color.should.have.key('argb')
        column.style.border.left.color.argb.should.equal('AAA3A3A3')

        column.style.border.bottom.should.have.keys(['style', 'color'])
        column.style.border.bottom.style.should.equal('medium')
        column.style.border.bottom.color.should.have.key('argb')
        column.style.border.bottom.color.argb.should.equal('AAA3A3A3')

        column.style.border.right.should.have.keys(['style', 'color'])
        column.style.border.right.style.should.equal('medium')
        column.style.border.right.color.should.have.key('argb')
        column.style.border.right.color.argb.should.equal('AAA3A3A3')
    })

    it('should set the fill', function () {
        let options = {
            "fileExtension": "xlsx",
            "workbook": {
                "worksheets": [
                    {
                        "title": "My First Worksheet",
                        "cols": [
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
        let column = worksheet.getColumn(1)
        
        column.style.should.have.key('fill')
        column.style.fill.type.should.equal('pattern')
        column.style.fill.pattern.should.equal('solid')
        column.style.fill.fgColor.should.contain.key('argb')
        column.style.fill.fgColor.argb.should.equal('FFFF0000')
    })
})
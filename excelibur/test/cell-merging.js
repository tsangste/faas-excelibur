const chai = require('chai')
const should = chai.should()
const Excel = require('exceljs')

const createWorkbook = require('../src/create-workbook')

describe('Workbook Cell Merging', function () {
    it('Should merge columns and values when 3 blank cells are merged with a value cell', function () {
        let options = {
            "fileExtension": "xlsx",
            "workbook": {
                "worksheets": [
                    {
                        "title": "My First Worksheet",
                        "rows": [
                            {
                                "cells": [
                                    { 'value': 'Test' },
                                ]
                            }
                        ],
                        "mergeCells": [
                            {
                                "top": 1, "left": 1, "bottom": 1, "right": 3
                            }
                        ]
                    }
                ]
            }
        }

        let workbook = createWorkbook(options);
        let a1Cell = workbook.getWorksheet(1).getCell('A1');
        let b1Cell = workbook.getWorksheet(1).getCell('B1');
        let c1Cell = workbook.getWorksheet(1).getCell('C1');

        a1Cell.value.should.equal('Test');
        a1Cell.type.should.equal(Excel.ValueType.String)
        b1Cell.value.should.equal('Test');
        b1Cell.type.should.equal(Excel.ValueType.Merge)
        c1Cell.value.should.equal('Test');
        c1Cell.type.should.equal(Excel.ValueType.Merge)
    })

    it('Should error when a merge value is negative', function () {
        let options = {
            "fileExtension": "xlsx",
            "workbook": {
                "worksheets": [
                    {
                        "title": "My First Worksheet",
                        "rows": [
                            {
                                "cells": [
                                    { 'value': 'Test' },
                                ]
                            }
                        ],
                        "mergeCells": [
                            {
                                "top": -1, "left": -1, "bottom": 1, "right": 3
                            }
                        ]
                    }
                ]
            }
        }

        should.throw(() => createWorkbook(options), Error, '-1 is out of bounds. Excel supports columns from 1 to 16384');
    })

    it('Should not perform a merge if bottom & right is missing', function () {
        let options = {
            "fileExtension": "xlsx",
            "workbook": {
                "worksheets": [
                    {
                        "title": "My First Worksheet",
                        "rows": [
                            {
                                "cells": [
                                    { 'value': 'Test' },
                                ]
                            }
                        ],
                        "mergeCells": [
                            {
                                "top": 1, "left": 1,
                            }
                        ]
                    }
                ]
            }
        }

        let workbook = createWorkbook(options);
        let a1Cell = workbook.getWorksheet(1).getCell('A1');
        let b1Cell = workbook.getWorksheet(1).getCell('B1');

        a1Cell.value.should.equal('Test');
        a1Cell.type.should.equal(Excel.ValueType.String)
        should.not.exist(b1Cell.value)
    })

    it('Should default merge to 1 if a value is missing', function () {
        let options = {
            "fileExtension": "xlsx",
            "workbook": {
                "worksheets": [
                    {
                        "title": "My First Worksheet",
                        "rows": [
                            {
                                "cells": [
                                    { 'value': 'Test' },
                                ]
                            }
                        ],
                        "mergeCells": [
                            {
                                "top": 1, "left": 1, "right": 2
                            }
                        ]
                    }
                ]
            }
        }

        let workbook = createWorkbook(options);
        let a1Cell = workbook.getWorksheet(1).getCell('A1');
        let b1Cell = workbook.getWorksheet(1).getCell('B1');

        a1Cell.value.should.equal('Test');
        a1Cell.type.should.equal(Excel.ValueType.String)
        b1Cell.value.should.equal('Test');
        b1Cell.type.should.equal(Excel.ValueType.Merge)
    })

    it('Should merge cell values to use the master cell value', function () {
        let options = {
            "fileExtension": "xlsx",
            "workbook": {
                "worksheets": [
                    {
                        "title": "My First Worksheet",
                        "rows": [
                            {
                                "cells": [
                                    { 'value': 'Test' },
                                    { 'value': 'This value should disappear' },
                                ]
                            }
                        ],
                        "mergeCells": [
                            {
                                "top": 1, "left": 1, "right": 2
                            }
                        ]
                    }
                ]
            }
        }

        let workbook = createWorkbook(options);
        let a1Cell = workbook.getWorksheet(1).getCell('A1');
        let b1Cell = workbook.getWorksheet(1).getCell('B1');

        a1Cell.value.should.equal('Test');
        a1Cell.type.should.equal(Excel.ValueType.String)
        b1Cell.value.should.equal('Test');
        b1Cell.type.should.equal(Excel.ValueType.Merge)
    })
})
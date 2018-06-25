"use strict"

const validate = require('./src/validator')
const writeWorkbook = require('./src/write-workbook')

module.exports = (context, callback) => {
  let options = JSON.parse(context)
  let errors = validate(options)

  if (errors.length > 0) {
    return callback(errors, null)
  }

  writeWorkbook(options)
    .then(result => result.pipe(process.stdout))
    .catch((err) => callback(err, null))
}

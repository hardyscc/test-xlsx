var fs = require('fs')
var concat = require('concat-stream')
var request = require('request')
var XLSX = require('xlsx')

function gotXls(xlsBuffer) {
  var workbook = XLSX.read(xlsBuffer, { type: 'buffer' })
  var csv = toCsv(workbook)
  console.log(csv)
}

function toCsv(workbook) {
  var result = [];
  workbook.SheetNames.forEach((sheetName) => {
    var csv = XLSX.utils.sheet_to_csv(workbook.Sheets[sheetName]);
    if(csv.length > 0){
      result.push('SHEET: ' + sheetName)
      result.push('')
      result.push(csv)
    }
  })
  return result.join('\n')
}

var concatStream = concat(gotXls)
var url = 'http://warrants.com.hk/en/data/warrant_search_data?action=excel&ucode=700&order=1&sname=all&wtype=all'
request(url).pipe(concatStream)

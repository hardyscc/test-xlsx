var fs = require('fs')
var concat = require('concat-stream')

var XLSX = require('xlsx')

var readStream = fs.createReadStream('warrant_search_result_xls.xls')
var concatStream = concat(gotXls)

readStream.on('error', handleError)
readStream.pipe(concatStream)

function gotXls(xlsBuffer) {
  var workbook = XLSX.read(xlsBuffer, {type:"buffer"})
  var csv = toCsv(workbook)
  console.log(csv)
}

function handleError(err) {
  // handle your error appropriately here, e.g.:
  console.error(err) // print the error to STDERR
  process.exit(1) // exit program with non-zero exit code
}

function toCsv(workbook) {
  var result = [];
  workbook.SheetNames.forEach(function(sheetName) {
    var csv = XLSX.utils.sheet_to_csv(workbook.Sheets[sheetName]);
    if(csv.length > 0){
      result.push("SHEET: " + sheetName);
      result.push("");
      result.push(csv);
    }
  });
  return result.join("\n");
}

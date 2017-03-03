require('blob.js')
var saveAs = require('file-saver').saveAs


var Workbook = require('./workbook')

const ws_name = 'SheetJS'
const data = [[1, 2, 3],
  [true, false, null, 'sheetjs'],
  ['foo', 'bar', new Date('2014-02-19T14:30Z'), '0.3'], ['baz', null, 'qux']]

const wb = new Workbook()
const ws = wb.createSheet(data)
wb.SheetNames.push(ws_name)
wb.Sheets [ws_name] = ws

// FIXME: FileSaver doesn't work in Safari, need another option
// Do we need npm install --save blob.js?
saveAs(new Blob(wb.getAsArrayBuffer(),{type:"application/octet-stream"}), "test.xlsx")
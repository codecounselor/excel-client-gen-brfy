// Had to fork the https://github.com/protobi/js-xlsx version of this module
// to get style functionality when saving a workbook

// And then manually apply the fix in 'xlsx-browserify-shim'
// because of some dependency issues
// https://github.com/SheetJS/js-xlsx/issues/143#issuecomment-154567695

var XLSX = require('../../js-xlsx-spider')

/*

### Needed Features

- Styles
  - https://github.com/SheetJS/js-xlsx/pull/66/files#diff-a470cfa184451a7854f1edc917351da3
    - Style Fork: https://github.com/SheetJS/js-xlsx/issues/128#issuecomment-76243917
  - Writing Styles: https://github.com/SheetJS/js-xlsx/issues/128

#### Requirements
    - Cell Background Color
    - Font Family
    - Font Weight
    - Font Color

[https://github.com/protobi/js-xlsx#cell-styles](Reference Guide)

- Cell Merge
  - https://github.com/SheetJS/js-xlsx/issues/41 (merged, and works; see below)

 Key	Description
 v	  raw value (see Data Types section for more info)
 w	  formatted text (if applicable)
 t	  cell type: b Boolean, n Number, e error, s String, d Date
 f	  cell formula encoded as an A1-style string (if applicable)
 F	  range of enclosing array if formula is array formula (if applicable)
 r	  rich text encoding (if applicable)
 h	  HTML rendering of the rich text (if applicable)
 c	  comments associated with the cell
 z	  number format string associated with the cell (if requested)
 l	  cell hyperlink object (.Target holds link, .tooltip is tooltip)
 s	  the style/theme of the cell (if applicable)

### Dummies Guide to ARGB

 Transparency is controlled by the alpha channel (AA in #AARRGGBB).
 Maximal value (255 dec, FF hex) means fully opaque.
 Minimum value (0 dec, 00 hex) means fully transparent.
 Values in between are semi-transparent, i.e. the color is mixed with the background color.

 Here is the table of % to hex values E.g. For 50% white you would use #80FFFFFF.

 100% — FF
 95% — F2
 90% — E6
 85% — D9
 80% — CC
 75% — BF
 70% — B3
 65% — A6
 60% — 99
 55% — 8C
 50% — 80
 45% — 73
 40% — 66
 35% — 59
 30% — 4D
 25% — 40
 20% — 33
 15% — 26
 10% — 1A
 5% — 0D
 0% — 00


 Colors from designers:

 <color rgb="FF4C4C4C" /> red
 <color rgb="FFEFF1F4" /> light pink
 <color rgb="FFD5DAE1" /> dark pink
 <color rgb="FFF9FAFB" /> white?
 <color rgb="FF720E0E" /> orange
 <color rgb="FFF9A898" /> yellow
 <color rgb="FF87690F" /> orange?
 <color rgb="FFFBE39C" /> yellow
 <color rgb="FF57720E" /> more pink?
 <color rgb="FFCAE67F" /> purple pink

 */


/**
 * @param s
 * @returns {ArrayBuffer}
 */
function s2ab (s) {
  var buf = new ArrayBuffer(s.length)
  var view = new Uint8Array(buf)
  for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
  return buf
}

function datenum (v, date1904) {
  if (date1904) v += 1462
  var epoch = Date.parse(v)
  return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000)
}

var Workbook = function () {

  this.SheetNames = []
  this.Sheets = {}
}

/**
 *
 * @param data a two dimensional array of data (rows and columns)
 * @param opts
 */
Workbook.prototype.createSheet = function (data, opts) {
  var ws = {}
  var range = {s: {c: 10000000, r: 10000000}, e: {c: 0, r: 0}}
  for (var R = 0; R != data.length; ++R) {
    for (var C = 0; C != data[R].length; ++C) {
      if (range.s.r > R) range.s.r = R
      if (range.s.c > C) range.s.c = C
      if (range.e.r < R) range.e.r = R
      if (range.e.c < C) range.e.c = C
      var cell = {v: data[R][C]}
      if (cell.v == null) continue
      var cell_ref = XLSX.utils.encode_cell({c: C, r: R})

      if (typeof cell.v === 'number') cell.t = 'n'
      else if (typeof cell.v === 'boolean') cell.t = 'b'
      else if (cell.v instanceof Date) {
        cell.t = 'n'
        cell.z = XLSX.SSF._table[14]
        cell.v = datenum(cell.v)
      }
      else cell.t = 's'

      // styles
      cell.s = {
        fill: {fgColor: { rgb: 'FF87690F'} }
        //font: {bold: true, color: { rgb: 'FFD5DAE1'}}
      }

      ws[cell_ref] = cell
    }
  }
  if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range)
  var merges = ws['!merges'] = []
  merges.push( { s: 'A1', e: 'F1' } )
  return ws
}

Workbook.prototype.getAsArrayBuffer = function () {
  return [s2ab(XLSX.write(this, {bookType: 'xlsx', bookSST: true, type: 'binary'}))]
}

module.exports = Workbook


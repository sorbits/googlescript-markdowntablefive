/**
 * This is free and unencumbered software released into the public domain.
 *
 * Anyone is free to copy, modify, publish, use, compile, sell, or
 * distribute this software, either in source code form or as a compiled
 * binary, for any purpose, commercial or non-commercial, and by any
 * means.
 *
 * In jurisdictions that recognize copyright laws, the author or authors
 * of this software dedicate any and all copyright interest in the
 * software to the public domain. We make this dedication for the benefit
 * of the public at large and to the detriment of our heirs and
 * successors. We intend this dedication to be an overt act of
 * relinquishment in perpetuity of all present and future rights to this
 * software under copyright law.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
 * EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
 * MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
 * IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR
 * OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
 * ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
 * OTHER DEALINGS IN THE SOFTWARE.
 *
 * For more information, please refer to <http://unlicense.org/>
 */

/**
 * name     : MarkdownTableMakerFive.gs
 * version  : 16
 * updated  : 2017-06-20
 * license  : http://unlicense.org/ The Unlicense
 * git      : https://github.com/pffy/googlescript-markdowntablefive
 *
 */
var MarkdownTableMaker = function () {

  // monospace Google fonts: https://www.google.com/fonts
  const MONOSPACE_FONTS = [
    'anonymous pro',
    'courier new',
    'cousine',
    'cutive mono',
    'droid sans mono',
    'fira mono',
    'inconsolata',
    'monospace',
    'nova mono',
    'oxygen mono',
    'pt mono',
    'roboto mono',
    'share tech mono',
    'source code pro',
    'ubuntu mono',
    'vt323'
  ];

  // Code solution based on info found here and here:
  // https://help.github.com/articles/github-flavored-markdown/#tables
  // https://github.com/adam-p/markdown-here/wiki/Markdown-Cheatsheet#tables
  const TABLE_EMPTY_RANGE = '|     |\r\n| --- |';

  var _derp = 'derp',

    // flag to crop input Range to last rows and/or columns with content
    _cropInputRange = false,

    // flag to force Markdown hyperlinks in strange places
    _forceHyperlinks = false;

  // input-output
  var _markdown = TABLE_EMPTY_RANGE,
    _range = {};


  // sets the range for this object
  function _setRange(rng) {
    _range = (typeof rng === 'object') ? rng : {};

    if(_range && _cropInputRange) {
      _range = _reduceRangeToBounds(_range);
    }

    _convert();
  }

  // reduces the range to a bounding box
  function _reduceRangeToBounds(rng) {

    // fixed top-left corner
    var firstRow = rng.getRow(); // 1-indexed
    var firstColumn = rng.getColumn(); // 1-indexed

    // proposed bottom-right corner
    // the largest possible recommendation is the n x m range of the spreadsheet
    var lastRow = rng.getLastRow(); // 1-indexed
    var lastColumn = rng.getLastColumn(); // 1-indexed

    var numRows = rng.getNumRows();
    var numColumns = rng.getNumColumns();
    var cellValues = rng.getValues();

    var currentValue = '';

    // smallest possible difference is 0 x 0
    // it is possible for firstRow to equal lastRow, etc
    var upperRowBoundDifference = 1;
    var upperColumnBoundDifference = 1;

    for(var i = 1; i <= numRows; i++) {
      for(var j = 1; j <= numColumns; j++) {
        if(cellValues && cellValues[i-1][j-1]) {
          upperRowBoundDifference = Math.max(upperRowBoundDifference, i);
          upperColumnBoundDifference = Math.max(upperColumnBoundDifference, j);
        }
      }
    }

    // the new bottom-right corner
    var newRow = firstRow + upperRowBoundDifference - 1;
    var newColumn = firstColumn + upperColumnBoundDifference - 1;

    var r1c1 = 'R' + firstRow + 'C' + firstColumn + ':'
      + 'R' + newRow + 'C' + newColumn;

    SpreadsheetApp.getActive().setActiveSelection(r1c1);
    return SpreadsheetApp.getActive().getActiveRange();
  }

  function getIndexesOfFilteredRows(ssId, sheetId) {
    var hiddenRows = [];

    // limit what’s returned from the API
    var fields = "sheets(data(rowMetadata(hiddenByFilter)),properties/sheetId)";
    var sheets = Sheets.Spreadsheets.get(ssId, {fields: fields}).sheets;  

    for(var i = 0; i < sheets.length; i++) {
      if(sheets[i].properties.sheetId == sheetId) {
        var data = sheets[i].data;
        var rows = data[0].rowMetadata;
        for(var j = 0; j < rows.length; j++) {
          if(rows[j].hiddenByFilter)
            hiddenRows.push(j);
        }
      }
    }
    return hiddenRows;
  }

  // converts Range object into Markdown string
  function _convert() {

    if(_range.isBlank()) {
      _markdown = TABLE_EMPTY_RANGE;
      _ready = true;
      return false;
    }

    var temp = [];

    var sheet = _range.getSheet();
    var ss = sheet.getParent();
    var hiddenRowIndexes = getIndexesOfFilteredRows(ss.getId(), sheet.getSheetId());

    for(col = 0; col < _range.getNumColumns(); col++) {
      var hidden = hiddenRowIndexes.slice();
      var hiddenRowIndex = hidden.count == 0 ? -1 : hidden.shift();

      var cells = [], align = [];
      for(row = 0; row < _range.getNumRows(); row++) {
        if(hiddenRowIndex == row + _range.getRow()-1) {
          hiddenRowIndex = hidden.count == 0 ? -1 : hidden.shift();
          continue;
        }

        var cell = _range.offset(row, col);
        cells.push(_cellToMarkdown(cell, row == 0));
        align.push(cell.getHorizontalAlignment().replace(/^general-/, ''));
      }

      // pad all cells to be the same width
      var width = cells.reduce(function (w, str) { return Math.max(str.length, w); }, 0);
      cells = cells.map(function (str, idx) {
        var padSize = width - str.length;
        var padRight = align[idx] == 'left' ? padSize : (align[idx] == 'center' ? Math.round(padSize / 2) : 0);
        return Array(padSize-padRight+1).join(' ') + str + Array(padRight+1).join(' ');
      });

      // insert table header/body divider if we have more than one row
      if(cells.length > 1) {
        // find the alignment of the column by majority, ignoring header row
        var leftCount   = align.filter(function (str, idx) { return idx != 0 && str == 'left';   }).length;
        var rightCount  = align.filter(function (str, idx) { return idx != 0 && str == 'right';  }).length;
        var centerCount = align.filter(function (str, idx) { return idx != 0 && str == 'center'; }).length;

        var left  = rightCount < Math.max(leftCount, centerCount)  ? ':' : '-';
        var right = leftCount <= Math.max(rightCount, centerCount) ? ':' : '-';
        cells.splice(1, 0, left + Array(width-2+1).join('-') + right);
      }

      temp.push(cells);
    }

    // Transpose temp array: col × row → row × col and turn into markdown
    temp = temp[0].map(function (_, c) { return temp.map(function (r) { return r[c]; }); });
    _markdown = temp.map(function (row) { return '| ' + row.join(' | ') + ' |'; }).join('\r\n');
  }

  // creates markdown for a cell
  function _cellToMarkdown(cell, isHeader) {
    var displayValue = cell.getDisplayValue();
    if(displayValue.length == 0) {
      return '';
    }

    var textFormat = '';

    if(cell.getFontLine() == 'line-through') {
      textFormat += '~~';
    }

    // table headers are styled by default, so skip explicit bold and italic
    if(isHeader == false) {
      if(cell.getFontStyle() == 'italic') {
        textFormat += '*';
      }

      if(cell.getFontWeight() == 'bold') {
        textFormat += '**';
      }
    }

    var textFormatOut = _reverse(textFormat);
    var orgDisplayValue = displayValue;

    // EXPERIMENTAL
    if(displayValue.search(/^https?:\/\//) == 0 && _forceHyperlinks) {
      displayValue = '<' + displayValue + '>';
    }
    else {
      displayValue = displayValue.replace(/[\r\n]/g, '<br/>');
      displayValue = displayValue.replace(/\|/g, '&#124;');

      var url = cell.getFormula().match(/^=HYPERLINK\("([^"]+)", .*\)/i);
      if(url && url.length == 2) {
        displayValue = '[' + displayValue + '](' + url[1] +')';
      }
    }

    if(MONOSPACE_FONTS.indexOf(cell.getFontFamily().toLowerCase()) != -1) {
      // we can only treat cell as raw if we did not insert special markup
      if(displayValue == orgDisplayValue) {
        displayValue = '`' + displayValue + '`';
      }
      else {
        displayValue = '<code>' + displayValue + '</code>';
      }
    }

    return textFormat + displayValue + textFormatOut;
  }

  // reverses a string
  function _reverse(str) {
    return str.split('').reverse().join('');
  }

  // crops sheet data
  function _cropSheetAsRange(){
    var sheet = SpreadsheetApp.getActiveSheet();
    var lastRow = sheet.getLastRow(); // method looks for content
    var lastColumn = sheet.getLastColumn(); // method looks for content
    return sheet.getRange(1, 1, lastRow, lastColumn);
  }

  return {

    /**
     * Returns the string representation of this object.
     * @return string text
     */
    toString: function() {
      return this.getMarkdown();
    },

    /**
     * Returns Markdown table text string.
     * @return string text
     */
    getMarkdown: function() {
      return _markdown;
    },

    /**
     * Returns spreadsheet Range of values.
     * @return Range
     */
    getRange: function() {
      return _range;
    },

    /**
     * Sets spreadsheet Range of values.
     * @return this object
     */
    setRange: function(range) {
      _setRange(range);
      return this;
    },

    /**
     * Sets entire spreadsheet as the Range
     * @return this object
     */
    setSheetAsRange: function(range) {
      _setRange(_cropSheetAsRange());
      return this;
    },

    /**
     * EXPERIMENTAL *
     * Sets flag to crop input Range.
     * @return this object
     */
    setCropInputRangeEnabled: function(enabled) {
      _cropInputRange = !!enabled;
      return this;
    },

    /**
     * EXPERIMENTAL *
     * Returns true if the Range is cropped upon input; false, otherwise.
     * @return boolean value
     */
    isCropInputRangeEnabled: function() {
      return _cropInputRange;
    },

    /**
     * EXPERIMENTAL *
     * Sets flag to force Markdown hyperlinks in strange places.
     * @return this object
     */
    setForceHyperlinksEnabled: function(enabled) {
      _forceHyperlinks = !!enabled;
      return this;
    },

    /**
     * EXPERIMENTAL *
     * Returns true if Markdown hyperlinks are forced; false, otherwise.
     * @return boolean value
     */
    isForceHyperlinksEnabled: function() {
      return _forceHyperlinks;
    }
  };
};
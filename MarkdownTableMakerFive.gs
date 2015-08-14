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
 * version  : 1
 * updated  : 2015-08-14
 * license  : http://unlicense.org/ The Unlicense
 * git      : https://github.com/pffy/googlescript-markdowntablefive
 *
 */
var MarkdownTableMaker = function () {

  // parts
  const _borderPipe = '|',

    // Code solution based on info found here and here:
    // https://help.github.com/articles/github-flavored-markdown/#tables
    // https://github.com/adam-p/markdown-here/wiki/Markdown-Cheatsheet#tables
    _tableColumnGeneral = ' ------ |',
    _tableColumnLeft = ' :------ |',
    _tableColumnCenter = ' :------: |',
    _tableColumnRight = ' ------: |',

    // space-space-pipe
    _tableCellEmpty = '  |',

    // CRLF-pipe-space
    _tableNewRow = '\r\n| ';


  var _derp = 'derp',

    // rows and columns
    _numRows = 0,
    _numColumns = 0,

    // flag to crop input Range to last rows and/or columns with content
    _cropInputRange = false,

    // flag to force Markdown hyperlinks in strange places
    _forceHyperlinks = false,

    // font strikethrough, italic, and bold
    _fontStyles = [],
    _fontWeights = [],
    _fontLines = [],

    // font typefaces
    _fontFamilies = [],

    // cell attributes
    _cellValues = [],
    _cellAlignments = [],
    _cellFormulas = [];

  // input-output
  var _markdown = '',
    _range = {};

  // let's go!
  _setRange(_cropSheetAsRange());



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

  // converts Range object into Markdown string
  function _convert() {

    if(_range.isBlank()) {
      _range = _cropSheetAsRange();
      if(_range) {
        _ready = true;
      }
    }

    _getMetaData();

    var output = '', // an accumulator

        // iterative-storage
        textFormat = '',
        textFormatClose = '',
        currentValue = '',
        faceValue = '',
        currentFormula = '';

    for (var i = 1; i <= _numRows; i++) {

      output += _tableNewRow;

      for (var j = 1; j <= _numColumns; j++) {

        // strikethrough
        if(_fontLines && (_fontLines[i-1][j-1] == 'line-through')) {
          textFormat += '~~';
        }

        // italic
        if(_fontStyles && (_fontStyles[i-1][j-1] == 'italic')) {
          textFormat += '*';
        }

        // bold
        if(_fontWeights && (_fontWeights[i-1][j-1] == 'bold')) {
          textFormat += '**';
        }

        // inline code backticks
        if(_fontFamilies && (_fontFamilies[i-1][j-1] == 'courier new,monospace')) {
          textFormat = '`'; // this OVERRIDES other formats
        }

        // formatting finished. add reversed string for the other bookend.
        textFormatClose = _reverse(textFormat);

        // cell values
        if(_cellValues) {
          currentValue = _cellValues[i-1][j-1];
        }

        // add a cell value OR add bupkis
        if(currentValue) {
          faceValue = textFormat + currentValue + textFormatClose;

          // cell formulas (optional)
          if(_cellFormulas) {
            currentFormula = _cellFormulas[i-1][j-1];
            if(_isValidHyperlink(currentFormula)) {
              var url = _getHyperlinkUrl(currentFormula);
              if( (url != currentValue)) {
                 var title = textFormat + currentValue + textFormatClose;
                 faceValue = '[' + title + '](' + url +')';
              }
            }
          }

          // EXPERIMENTAL
          if(_hasUrlScheme(currentValue) && _forceHyperlinks) {
            faceValue = '[' + faceValue + '](' + currentValue +')';
          }

          output += ' ' +  faceValue + ' ' + _borderPipe;
        } else {
          output += _tableCellEmpty;
        }

        // reset formatting each time
        textFormat = '';
        textFormatClose = '';
        currentValue = '';
        faceValue = '';
        currentFormula = '';
      }

      // table column alignment
      if(i < 2) {
        output += _tableNewRow;
        for (var k = 1; k <= _numColumns; k++) {
          switch(_cellAlignments[i-1][k-1]) {
            case 'center':
              output += _tableColumnCenter;
              break;
            case 'right':
              output += _tableColumnRight;
              break;
            case 'left':
              output += _tableColumnLeft;
              break;
            default:
              output += _tableColumnGeneral;
              break;
          }
        }
      }
    }

    _markdown = output;
  }

  // reverse a  string
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

  // accumulate layered meta data
  function _getMetaData() {

    _numRows = _range.getNumRows();
    _numColumns = _range.getNumColumns();

    _fontStyles = _range.getFontStyles();
    _fontWeights = _range.getFontWeights();
    _fontLines = _range.getFontLines();

    _fontFamilies = _range.getFontFamilies();

    _cellFormulas = _range.getFormulas();
    _cellAlignments = _range.getHorizontalAlignments();
    _cellValues = _range.getValues();
  }

  // detects HYPERLINK formula
  // does not process hyperlink with cell references
  // processes hyperlinks with url and title as strings
  function _isValidHyperlink(str) {
    // todo: add validation for URL with cell references (or the negative of that)
    return ((str.indexOf('=HYPERLINK') == 0 ) && (str.indexOf('","') > 0 )) ? true : false;
  }

  // looks like URL scheme
  function _hasUrlScheme(str) {

    str = '' + str;
    const URL_SCHEMES = [
      'http://',
      'https://'
    ];

    for(var s in URL_SCHEMES) {
      if(str.indexOf(URL_SCHEMES[s]) == 0) {
        return true;
      }
     return false;
    }
  }

  // extracts HYPERLINK url
  function _getHyperlinkUrl(str) {
    var arr = JSON.parse('[' + str.slice(11, -1) + ']');
    var url = arr[0];
    return url;
  }

  // builds Markdown hyperlink
  function _createHyperlinkMarkdown(title, url) {
    return '[' + title + '](' + url +')';
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
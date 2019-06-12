/**
 * Utility functions for Google sheets stock tracker
 *
 * Copyright 2019 Chris Hundley
 *
 * MIT LICENSE
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation
 * files (the 'Software'), to deal in the Software without restriction, including without limitation the rights to use, copy,
 * modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the
 * Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

 * THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
 * WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
 * COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
 * ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 *
 */


var sheetCache = {};
var debugRow = 2;

/**
 * Gets a sheet by name
 */
function getSheetByName(sheetName) {
  if (!sheetCache[sheetName]) {
    sheetCache[sheetName] = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  }
  return sheetCache[sheetName];
}

/**
 * Given a sheet name and cell, return the value in that cell
 */
function getCell(sheetName, cell) {
  return getSheetByName(sheetName).getRange(cell).getValue();
}

/**
 * Given a sheet name, cell, and value, set the value in that cell
 */
function setCell(sheetName, cell, value) {
  getSheetByName(sheetName).getRange(cell).setValue(value);
}

/**
 * Given a sheet name, start cell, and end cell, set the value in that cell to blank
 */
function clearRange(sheetName, start, end) {
  getSheetByName(sheetName).getRange(sheetName + '!' + start + ':' + end).setValue('');
}

/**
 * Gets the next column in sequence, works up to column ZZ (676 columns)
 *
 * Examples:  Pass in  'C' --> 'D'
 *            Pass in 'AZ' --> 'BA'
 */
function getNextColumn(currentColumn) {
  var nextColumn = '';
  var nextIndex = currentColumn.charCodeAt(currentColumn.length - 1) + 1;

  // went past Z
  if (nextIndex == 91) {
    if (currentColumn.length == 1) {
      nextColumn = 'AA';
    }
    else {
      nextColumn = getNextColumn(currentColumn[0]) + 'A';
    }
  }
  else {
    nextColumn = currentColumn.substring(0, currentColumn.length - 1) + String.fromCharCode(nextIndex);
  }

  return nextColumn;
}

/**
 * Given a sheet, starting column, and row, determine the last valid value at the end of the row
 */
function getLastValueInRow(sheetName, startColumn, row) {
  var done = false;
  var retVal = null;
  while (!done) {
    var val = getCell(sheetName, startColumn + row);
    if (val) {
      retVal = val;
      startColumn = getNextColumn(startColumn);
    }
    else {
      done = true;
    }
  }
  return retVal;
}

/**
 * Given a range, find the minimum value in the range. Used for setting minimum range on chart vertical axis
 */
function getMinValueInRange(range) {
  var minVal = 100000000;
  var curVal = 0;
  var rows = range.getNumRows();
  var cols = range.getNumColumns();
  for (var r=1; r<= rows; r++) {
    for (var c=1; c<= cols; c++) {
      curVal = range.getCell(r, c).getValue();
      if (curVal && curVal < minVal) {
        minVal = curVal;
      }
    }
  }
  return minVal;
}

/**
 * Attempts to select a minimum y-axis scale cutoff appropriate for the data so the chart is
 * readable. Without doing this prices would not show much variance, especially the daily data
 */
function getScaleCutoff(value) {
  if (value < .1) {
    return 0;
  }
  else if (value < 1) {
    return value;
  }
  else if (value < 10) {
    return parseInt(value / 1, 10) * 1;
  }
  else if (value < 50) {
    return parseInt(value / 2, 10) * 2;
  }
  else if (value < 100) {
    return parseInt(value / 5, 10) * 5;
  }
  else if (value < 1000) {
    return parseInt(value / 20, 10) * 20;
  }
  else if (value < 10000) {
    return parseInt(value / 1000, 10) * 1000;
  }
  else if (value < 100000) {
    return parseInt(value / 10000, 10) * 10000;
  }
}

/**
 * Given a range of cells, format the numbers in that range based on passed in format
 */
function formatRange(range, format) {
  range.setNumberFormat(format);
}

/**
 * Toast alert
 */
function alert(message, title, timeout) {
  if (!title) {
    title = 'Ticker status';
  }
  if (!timeout) {
    timeout = 5;
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title, timeout);
}

/**
 * Add a debug message to the debug tab
 */
function debugMessage(call, message) {
  setCell(sheetNames.debug, 'A' + debugRow, call);
  setCell(sheetNames.debug, 'B' + debugRow, message);
  debugRow++;
}

/**
 * TICKER v1.1
 *
 * Utility functions for managing/manipulating data in a Google sheet
 *
 * Based on the following docs:
 *   - https://developers.google.com/apps-script/reference/spreadsheet/sheet
 *   - https://developers.google.com/apps-script/reference/spreadsheet/range
 *
 *
 * Copyright 2019-2020 Chris Hundley
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
 * @param {string} sheetName - the name of the sheet as labeled in the tabs at the bottom of the view
 * @return {Sheet} - see https://developers.google.com/apps-script/reference/spreadsheet/sheet
 */
function getSheetByName(sheetName) {
  if (!sheetCache[sheetName]) {
    sheetCache[sheetName] = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  }
  return sheetCache[sheetName];
}

/**
 * Given a sheet name and cell, return the value in that cell
 * @param {string} sheetName - the sheet containing the cell
 * @param {string} cell - the cell in A1 Notation, eg 'B7'
 * @returns {string|number|date} - the string, number, or date in the cell
 */
function getCell(sheetName, cell) {
  return getSheetByName(sheetName).getRange(cell).getValue();
}

/**
 * Given a sheet name, cell, and value, set the value in that cell
 * @param {string} sheetName - the sheet containing the cell
 * @param {string} cell - the cell in A1 Notation, eg 'B7'
 * @param {string|number|date} value - the value to put in the cell
 */
function setCell(sheetName, cell, value) {
  getSheetByName(sheetName).getRange(cell).setValue(value);
}

/**
 * Given a sheet name and range, remove the values in the range
 * @param {string} sheetName - the sheet containing the range to clear
 * @param {string} range - the range in A1 Notation, eg 'A8:F16'
 */
function clearRangeValues(sheetName, range) {
  getSheetByName(sheetName).getRange(range).setValue('');
}

/**
 * Given a sheet name and range, remove all formatting
 * @param {string} sheetName - the sheet containing the range to clear
 * @param {string} range - the range in A1 Notation, eg 'A8:F16'
 */
function clearRangeFormat(sheetName, range) {
  getSheetByName(sheetName).getRange(range).clearFormat();
}

/**
 * Gets the next column in sequence, works up to column ZZ (676 columns)
 * @param {string} currentColumn - the current column
 * @returns {string} - the next column (to the right of the column passed in)
 *
 * Examples:  getNextColumn('C')  --> 'D'
 *            getNextColumn('AZ') --> 'BA'
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
 * Given a sheet, starting column, and row, determine the last non-blank value at the end of the row
 * @param {string} sheetName - the sheet containing the row to check
 * @param {string} startColumn - the column to start searching from
 * @param {string} row - the row to check for the first non-blank value
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
 * Given a sheet, starting row, and column, determine the first row that doesn't have a value
 * @param {string} sheetName - the sheet containing the column to check
 * @param {string} startRow - the row to start searching from
 * @param {string} column - the column to check for the first non-blank value
 */
function getFirstEmptyRow(sheetName, startRow, column) {
  var done = false;
  var retRow = startRow;
  while (!done) {
    var val = getCell(sheetName, column + retRow);
    if (val) {
      retRow++;
    }
    else {
      done = true;
    }
  }
  return retRow;
}

/**
 * Given a range, find the minimum value in the range. Used for setting minimum range on chart vertical axis
 * @param {Range} range - a range of cells to search and determine the minimum value
 * @returns {number} - the minimum value found in the range of cells
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
 * readable. Without doing this prices would not show much variance, especially the daily data. Works together
 * with the getMinValueInRange function to help format the y-axis of charts
 * @param {number} value - the numerical value representing the minimum data point in a series
 * @returns {number} - the scale cutoff value deteremined to be the best choice
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
 * @param {string} range - the range in A1 Notation, eg 'A8:F16'
 * @param {string} format - the format desired, eg '#,###.##'
 */
function formatRange(range, format) {
  range.setNumberFormat(format);
}

/**
 * Given a sheet, range, and color, create full borders including internal in the color specified
 * @param {string} sheetName - the sheet containing the range to format
 * @param {string} range - the range in A1 Notation, eg 'A8:F16'
 * @param {string} color - the color for the border, eg 'red' or '#ff0000'
 */
function setRangeBorder(sheetName, range, color) {
  getSheetByName(sheetName).getRange(range).setBorder(true, true, true, true, true, true, color, SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * Clear borders in the passed in range
 * @param {string} sheetName - the sheet containing the range to format
 * @param {string} range - the range in A1 Notation, eg 'A8:F16'
 */
function clearRangeBorder(sheetName, range) {
  getSheetByName(sheetName).getRange(range).setBorder(false, false, false, false, false, false, null, null);
}

/**
 * Set the font size of the specified range
 * @param {string} sheetName - the sheet containing the range to format
 * @param {string} range - the range in A1 Notation, eg 'A8:F16'
 * @param {string} size - the font size in points, eg 14
 */
function setRangeFontSize(sheetName, range, size) {
  getSheetByName(sheetName).getRange(range).setFontSize(size);
}

/**
 * Set the font weight of the specified range. Possible weights are 'bold', 'normal', or null (to reset)
 * @param {string} sheetName - the sheet containing the range to format
 * @param {string} range - the range in A1 Notation, eg 'A8:F16'
 * @param {string} weight - the weight to set, eg 'bold'
 */
function setRangeWeight(sheetName, range, weight) {
  getSheetByName(sheetName).getRange(range).setFontWeight(weight);
}

/**
 * Set the horizontal text alignment of the specified range. Possible values are 'left', 'center', 'right', or null (to reset)
 * @param {string} sheetName - the sheet containing the range to format
 * @param {string} range - the range in A1 Notation, eg 'A8:F16'
 * @param {string} alignment - the alignment to set, eg 'center'
 */
function setRangeHorizontalAlignment(sheetName, range, alignment) {
  getSheetByName(sheetName).getRange(range).setHorizontalAlignment(alignment);
}

/**
 * Set the font color of the specified range
 * @param {string} sheetName - the sheet containing the range to format
 * @param {string} range - the range in A1 Notation, eg 'A8:F16'
 * @param {string} color - the color to set, eg 'red' or '#ff0000'
 */
function setRangeColor(sheetName, range, color) {
  getSheetByName(sheetName).getRange(range).setFontColor(color);
}

/**
 * Set the cell background color of the specified range
 * @param {string} sheetName - the sheet containing the range to format
 * @param {string} range - the range in A1 Notation, eg 'A8:F16'
 * @param {string} color - the color to set, eg 'red' or '#ff0000'
 */
function setRangeBackground(sheetName, range, color) {
  getSheetByName(sheetName).getRange(range).setBackgroundColor(color);
}

/**
 * Merge the specified range of cells horizontally
 * @param {string} sheetName - the sheet containing the range to merge
 * @param {string} range - the range in A1 Notation, eg 'B4:B10'
 */
function mergeHorizontal(sheetName, range) {
  getSheetByName(sheetName).getRange(range).mergeAcross();
}

/**
 * Formats the cells in a range with up (green) or down (red) formatting to show net gain or loss in assets
 * @param {string} sheetName - the sheet containing the range to format
 * @param {string} range - the range in A1 Notation, eg 'B4:B10'
 */
function formatUpDown(sheetName, range) {
  var rangeLocal = getSheetByName(sheetName).getRange(range);
  var rows = rangeLocal.getNumRows();
  var cols = rangeLocal.getNumColumns();

  for (var r=1; r<= rows; r++) {
    for (var c=1; c<= cols; c++) {
      val = rangeLocal.getCell(r, c).getValue();
      if (val && val < 0) {
        setRangeColor(sheetName, rangeLocal.getCell(r, c).getA1Notation(), colors.down);
      }
      else {
        setRangeColor(sheetName, rangeLocal.getCell(r, c).getA1Notation(), colors.up);
      }
    }
  }
}

/**
 * GSheet toast alert (the little box that pops up in the lower-right corner). Whenever this is called the new
 * message immediately displays and replaces the message already there even if the timeout isn't reached yet
 * @param {string} message - the message to display
 * @param {string} title - the title of the popup
 * @param {number} timeout - how many seconds to show the popup. Optional, defaults to 8 seconds
 */
function alert(message, title, timeout) {
  if (!title) {
    title = 'Ticker status';
  }
  if (!timeout) {
    timeout = 8;
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title, timeout);
}

/**
 * Add a debug message to the debug tab - effectively raw output from API requests to help troubleshoot when
 * something is failing
 * @param {string} call - the API call that was made
 * @param {string} message - the message to log in the debug sheet
 */
function debugMessage(call, message) {
  setCell(sheetNames.debug, 'A' + debugRow, call);
  setCell(sheetNames.debug, 'B' + debugRow, message);
  debugRow++;
}

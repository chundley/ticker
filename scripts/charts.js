/**
 * TICKER v1.0
 *
 * This script is a library for rendering charts for the Ticker worksheet
 *
 * Based on the following docs:
 *   - https://developers.google.com/apps-script/reference/spreadsheet/embedded-area-chart-builder
 *   - https://developers.google.com/chart/interactive/docs/gallery/areachart#configuration-options
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


/**
 * Remove any existing charts from a worksheet (tab)
 * @param {string} sheetName - the name of the sheet as labeled in the tabs at the bottom of the view
 */
function removeChartsFromSheet(sheetName) {
  var sheet = getSheetByName(sheetName);
  var charts = sheet.getCharts();
  charts.forEach(function(chart) {
    sheet.removeChart(chart);
  });
}

/**
 * Generate an area chart with optional parameters. NOTE: in this case we are transposing rows and columns due to the shape
 * of the data on our worksheet
 * @param {string} sheetName - the sheet to render the chart on
 * @param {string} range - the range containing the data in A1 Notation, eg 'A8:F16'
 * @param {Object} options - a set of options to manipulate the look of the chart
 * @param {string} options.title - a title to give the chart, default none
 * @param {number} options.row - the row number (1-based) for the top-left corner of the chart, default 1
 * @param {number} options.column - the column number (1-based) for the top-left corner of the chart, eg column B is 2, default 1
 * @param {string} options.offsetLeft - offset (in pixels) to move the chart right from the starting point, default 0
 * @param {string} options.offsetTop - offset (in pixels) to move the chart down from the starting point, default 0
 * @param {string} options.width - width of the chart in pixels, default 300
 * @param {string} options.height - height of the chart in pixels, default 200
 * @param {Array[string]} options.colors - an array of colors to use for the chart series, eg ['#ff0000', '#00ff00'], default GSheet theme colors
 * @param {string} options.backgroundColor - background color of the chart, eg 'white' or '#ff0000', default 'white'
 * @param {string} options.titleColor - title color of the chart, default 'black'
 * @param {string} options.yAxisLabelColor - color of the y-axis labels, default '#666666'
 * @param {string} options.gridLineColor - color of the grid lines, default '#cccccc'
 */
function areaChart(sheetName, range, options) {
  var title = options.title || '';
  var row = options.row || 1;
  var column = options.column || 1;
  var offsetLeft = options.offsetLeft || 0;
  var offsetTop = options.offsetTop || 0;
  var width = options.width || 300;
  var height = options.height || 200;
  var colors = options.colors || [];
  var backgroundColor = options.backgroundColor || 'white';
  var titleColor = options.titleColor || 'black';
  var yAxisLabelColor = options.yAxisLabelColor || '#666666';
  var gridLineColor = options.gridLineColor || '#cccccc';

  var sheet = getSheetByName(sheetName);
  var cutoff = getScaleCutoff(getMinValueInRange(range));

  // if the data set is all over $1000, remove the pennies for better chart formatting
  if (cutoff >= 1000) {
    formatRange(range, '$#,###');
  }

  var chart = sheet.newChart()
         .setChartType(Charts.ChartType.AREA)
         .addRange(range)
         .setTransposeRowsAndColumns(true)
         .setPosition(row, column, offsetLeft, offsetTop)
         .setOption('backgroundColor', backgroundColor)
         .setOption('colors', colors)
         .setOption('height', height)
         .setOption('width', width)
         .setOption('title', title)
         .setOption('titleTextStyle', { fontSize: 11, color: titleColor, bold: true })
         .setOption('hAxis.direction', -1)
         .setOption('vAxes', { 0: {
                 viewWindow: { min: cutoff },
                 textStyle: { color: yAxisLabelColor },
                 gridlines: { color: gridLineColor }
          } } )
         .build();

  sheet.insertChart(chart);
}

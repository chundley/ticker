/**
 * This script is a library for rendering charts for the Ticker worksheet
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
 */
function removeChartsFromSheet(sheetName) {
  var sheet = getSheetByName(sheetName);
  var charts = sheet.getCharts();
  charts.forEach(function(chart) {
    sheet.removeChart(chart);
  });
}

/**
 * Generate an area chart with optional parameters
 */
function areaChart(range, sheetName, options) {
  var title = options.title || '';
  var row = options.row || 1;
  var column = options.column || 1;
  var offsetLeft = options.offsetLeft || 0;
  var offsetTop = options.offsetTop || 0;
  var width = options.width || 300;
  var height = options.height || 200;
  var colors = options.colors || [];
  var backgroundColor = options.backgroundColor || 'white';
  var titleColor = options.titleColor || '#ffffff';
  var yAxisLabelColor = options.yAxisLabelColor || '#bbbbbb';
  var gridLineColor = options.gridLineColor || '#666666';

  var sheet = getSheetByName(sheetName);
  var cutoff = getScaleCutoff(getMinValueInRange(range));
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

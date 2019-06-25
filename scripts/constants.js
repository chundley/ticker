/**
 * TICKER v1.0
 *
 * Constants used throughout the Ticker scripts
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
 * Name of sheets in the workbook
 */
var sheetNames = {
  config: 'Config',
  dashboard: 'Dashboard',
  detail: 'Detail',
  stockPurchased: 'StockPurchased',
  stockDividend: 'StockDividend',
  cryptoPurchased: 'CryptoPurchased',
  stockCurrent: 'StockCurrent',
  stockHistory: 'StockHistory',
  cryptoCurrent: 'CryptoCurrent',
  cryptoHistory: 'CryptoHistory',
  debug: 'Debug'
};

/**
 * Colors used for rendering data and charts
 */
var colors = {
  dashboardText: '#ffffff',
  dashboardHeaderBackground: '#292929',
  dashboardSubheaderBackground: '#333333',
  dashboardBackground: '#434343',
  dashboardOutlines: '#888888',
  up: '#00dd00',
  down: '#e09595',
  stockDayChart: '#0075c9',
  stockHourChart: '#f06eaa',
  cryptoDayChart: '#00a6b6',
  cryptoHourChart: '#9157d8',
  chartYAxisLabel: '#bbbbbb',
  chartgridLineColor: '#666666',
  chartTitleColor: '#ffffff'
}

// TODO: Add all the column limits/constants so they aren't hard-coded all over the place

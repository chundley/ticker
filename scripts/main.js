/**
 * TICKER v1.0
 *
 * This script uses 3rd-party API's for stock and crypto pricing data to keep assets
 * up to date in a Google sheet.
 *
 * Documentation: https://github.com/chundley/ticker
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
 * On every run there are values we need to track to support dynamic layout
 */
var context = {
  lastDashboardRowUsed: 0,
  stockTotal: {},
  cryptoTotal: {},
  stockDetail: {},
  cryptoDetail: {}
};

/**
 * Refresh all data and update the dashboard
 */
function refreshAll() {
  removeChartsFromSheet(sheetNames.dashboard);
  clearRangeValues(sheetNames.dashboard, 'A2:Z1000');
  clearRangeFormat(sheetNames.dashboard, 'A2:Z1000');
  setRangeBackground(sheetNames.dashboard, 'A2:Z1000', colors.dashboardBackground);
  setRangeColor(sheetNames.dashboard, 'A2:Z1000', colors.dashboardText);

  alert('Updating current stock prices');
  refreshStockCurrent();

  alert('Updating historical stock prices');
  refreshStockHistory();

  alert('Updating dashboard stock returns');
  updateStockDashboard();

  alert('Updating current crypto prices');
  refreshCryptoCurrent();

  alert('Updating historical crypto prices');
  refreshCryptoHistory();

  alert('Updating dashboard crypto returns');
  updateCryptoDashboard();

  alert('Updating dashboard charts');
  refreshDashboardCharts();

  alert('Updating detail tab');
  refreshDetail();

  alert('Refresh complete');
}

/**
 * Refresh all real-time prices and update the dashboard
 */
function refreshCurrent() {
  removeChartsFromSheet(sheetNames.dashboard);
  clearRangeValues(sheetNames.dashboard, 'A2:Z1000');
  clearRangeFormat(sheetNames.dashboard, 'A2:Z1000');
  setRangeBackground(sheetNames.dashboard, 'A2:Z1000', colors.dashboardBackground);
  setRangeColor(sheetNames.dashboard, 'A2:Z1000', colors.dashboardText);

  alert('Updating current stock prices');
  refreshStockCurrent();

  alert('Updating dashboard stock returns');
  updateStockDashboard();

  alert('Updating current crypto prices');
  refreshCryptoCurrent();

  alert('Updating dashboard crypto returns');
  updateCryptoDashboard();

  alert('Updating dashboard charts');
  refreshDashboardCharts();

  alert('Updating detail tab');
  refreshDetail();

  alert('Refresh complete');
}

/**
 * Refreshes the stock pricing - current trading day in five minute intervals. The general process is:
 * 1) Get unqiue symbols from the config tab
 * 2) Request stock price data from Alpaca
 * 3) Iterate the list of stock symbols and prices, filling in the StockCurrent worksheet
 * 4) The FIRST iteration will also fill in the date row
 * 5) After pulling all data, update the StockPurchased tab with updated returns
 */
function refreshStockCurrent() {
  var symbols = getUniqueSymbols(sheetNames.config, 'D', 4);
  var data = getCurrentStockPrices(symbols, 80);
  var row = 3;
  var datesDone = false;

  clearRangeValues(sheetNames.stockCurrent, 'A3:CC103');
  symbols.forEach(function(symbol) {
    // account for bad symbols
    if (data[symbol]) {
      var column = 'A';
      setCell(sheetNames.stockCurrent, column + row, symbol);
      data[symbol] = data[symbol].reverse();
      data[symbol].forEach(function(day) {
        column = getNextColumn(column);
        if (!datesDone) {
          setCell(sheetNames.stockCurrent, column + '2', new Date(day.t*1000));
        }
        setCell(sheetNames.stockCurrent, column + row, day.c);
      });
      datesDone = true;
      row++;
    }
  });

  alert('Updating stock purchased');
  updateStockPurchased();
}

/**
 * Using the prices from StockCurrent, update the StockPurchased tab to get gains and losses
 */
function updateStockPurchased() {
  var done = false;
  var row = 3;
  var symbols = {};

  // grab most current price for all tracked stock
  while (!done) {
    var symbol = getCell(sheetNames.stockCurrent, 'A' + row);
    if (symbol) {
      symbols[symbol] = getCell(sheetNames.stockCurrent, 'B' + row);
    }
    else {
      done = true;
    }
    row++;
  }

  // Update the StockPurchased tab with latest prices
  done = false;
  row = 3;
  while (!done) {
    var symbol = getCell(sheetNames.stockPurchased, 'A' + row);
    if (symbol) {
      if (symbols[symbol]) {
        setCell(sheetNames.stockPurchased, 'E' + row, symbols[symbol]);
      }
      else {
        setCell(sheetNames.stockPurchased, 'E' + row, 'n/a');
      }
    }
    else {
      done = true;
    }
    row++;
  }
}

/**
 * Update the dashboard with stock prices and returns
 */
function updateStockDashboard() {
  clearRangeValues(sheetNames.dashboard, 'A5:H105');
  done = false;
  row = 3;
  var sData = {};
  while (!done) {
    var symbol = getCell(sheetNames.stockPurchased, 'A' + row);
    if (symbol) {
      if (!sData[symbol]) {
        sData[symbol] = {
          shares: 0,
          price: 0,
          cost: 0,
          value: 0,
          div: 0
        };
      }
      sData[symbol].shares += getCell(sheetNames.stockPurchased, 'B' + row);
      sData[symbol].price = getCell(sheetNames.stockPurchased, 'E' + row);
      sData[symbol].cost += getCell(sheetNames.stockPurchased, 'D' + row);
      sData[symbol].value += getCell(sheetNames.stockPurchased, 'F' + row);
      row++;
    }
    else {
      done = true;
    }
  }

  // now calculate dividends
  row = 3;
  done = false;
  while (!done) {
    var symbol = getCell(sheetNames.stockDividend, 'B' + row);
    if (symbol && symbol.length > 0) {
      if (!sData[symbol]) {
        sData[symbol] = {
          cost: 0,
          value: 0,
          div: 0
        };
      }
      sData[symbol].div += getCell(sheetNames.stockDividend, 'C' + row);
      row++;
    }
    else {
      done = true;
    }
  }

  var totals = {
    cost: 0,
    value: 0,
    div: 0
  };

  // stock summary header
  row = 2;
  setCell(sheetNames.dashboard, 'A' + row, 'Stock Summary');
  mergeHorizontal(sheetNames.dashboard, 'A' + row + ':H' + row);
  setRangeFontSize(sheetNames.dashboard, 'A' + row, 14);
  setRangeBackground(sheetNames.dashboard, 'A' + row, colors.dashboardHeaderBackground);
  setRangeBorder(sheetNames.dashboard, 'A' + row, colors.dashboardOutlines);
  setRangeHorizontalAlignment(sheetNames.dashboard, 'A' + row, 'center');
  row++;

  // row headers
  setCell(sheetNames.dashboard, 'A' + row, 'Symbol');
  setCell(sheetNames.dashboard, 'B' + row, 'Shares');
  setCell(sheetNames.dashboard, 'C' + row, 'Price');
  setCell(sheetNames.dashboard, 'D' + row, 'Cost Basis');
  setCell(sheetNames.dashboard, 'E' + row, 'Value');
  setCell(sheetNames.dashboard, 'F' + row, 'Dividends');
  setCell(sheetNames.dashboard, 'G' + row, '$ Net');
  setCell(sheetNames.dashboard, 'H' + row, '% Net');
  setRangeHorizontalAlignment(sheetNames.dashboard, 'B' + row + ':H' + row, 'right');
  setRangeBorder(sheetNames.dashboard, 'A' + row + ':H' + row, colors.dashboardOutlines);
  setRangeBackground(sheetNames.dashboard, 'A' + row + ':H' + row, colors.dashboardSubheaderBackground);
  row++;

  var startRow = row;
  Object.keys(sData).forEach(function(item) {
    totals.cost += sData[item].cost;
    totals.value += sData[item].value;
    totals.div += sData[item].div;

    setCell(sheetNames.dashboard, 'A' + row, item);
    setCell(sheetNames.dashboard, 'B' + row, sData[item].shares);
    setCell(sheetNames.dashboard, 'C' + row, sData[item].price);
    setCell(sheetNames.dashboard, 'D' + row, sData[item].cost);
    setCell(sheetNames.dashboard, 'E' + row, sData[item].value);
    setCell(sheetNames.dashboard, 'F' + row, sData[item].div);
    setCell(sheetNames.dashboard, 'G' + row, (sData[item].value + sData[item].div) - sData[item].cost);
    setCell(sheetNames.dashboard, 'H' + row, (sData[item].value + sData[item].div - sData[item].cost) / sData[item].cost);
    setRangeBorder(sheetNames.dashboard, 'A' + row + ':' + 'H' + row, colors.dashboardOutlines);
    row++;
  });

  // add totals
  setCell(sheetNames.dashboard, 'D' + row, totals.cost);
  setCell(sheetNames.dashboard, 'E' + row, totals.value);
  setCell(sheetNames.dashboard, 'F' + row, totals.div);
  setCell(sheetNames.dashboard, 'G' + row, (totals.value + totals.div) - totals.cost);
  setCell(sheetNames.dashboard, 'H' + row, (totals.value + totals.div - totals.cost) / totals.cost);
  setRangeBorder(sheetNames.dashboard, 'D' + row + ':' + 'H' + row, colors.dashboardOutlines);
  setRangeBackground(sheetNames.dashboard, 'D' + row + ':' + 'H' + row, colors.dashboardSubheaderBackground);
  setRangeWeight(sheetNames.dashboard, 'D' + row + ':' + 'H' + row, 'bold');

  // set up/down colors
  formatUpDown(sheetNames.dashboard, 'G' + startRow + ':H' + row);

  // save context for use later
  context.lastDashboardRowUsed = row;
  context.stockTotal = totals;
  context.stockDetail = sData;
}

/**
 * Updates the StockHistory tab, getting daily closing prices for the last year for each stock symbol
 * in the config tab
 */
function refreshStockHistory() {
  var symbols = getUniqueSymbols(sheetNames.config, 'D', 4);
  var data = getHistoricalStockPrices(symbols, 250);
  var row = 3;
  var datesDone = false;

  clearRangeValues(sheetNames.stockHistory, 'A3:IQ103');
  symbols.forEach(function(symbol) {
    // account for bad symbols
    if (data[symbol]) {
      var column = 'A';
      setCell(sheetNames.stockHistory, column + row, symbol);
      data[symbol] = data[symbol].reverse();
      data[symbol].forEach(function(day) {
        column = getNextColumn(column);
        if (!datesDone) {
          setCell(sheetNames.stockHistory, column + '2', new Date(day.t*1000));
        }
        setCell(sheetNames.stockHistory, column + row, day.c);
      });
      datesDone = true;
      row++;
    }
  });
}

/**
 * Refreshes crypto pricing - current trading day in one minute intervals. The general process is:
 * 1) Get unqiue symbols from the config tab
 * 2) Request coin price data from CryptoCompare
 * 3) Iterate the list of coin symbols and prices, filling in the CryptoCurrent worksheet
 * 4) The FIRST iteration will also fill in the date row
 * 5) After pulling all data, update the CryptoPurchased tab with updated returns
 */
function refreshCryptoCurrent() {
  var symbols = getUniqueSymbols(sheetNames.config, 'E', 4);
  var data = getCryptoCurrent(symbols, 480);
  var row = 3;
  var datesDone = false;

  clearRangeValues(sheetNames.cryptoCurrent, 'A3:RN103');
  symbols.forEach(function(symbol) {
    // account for bad symbols
    if (data[symbol]) {
      var column = 'A';
      setCell(sheetNames.cryptoCurrent, column + row, symbol);
      data[symbol] = data[symbol].reverse();
      data[symbol].forEach(function(day) {
        column = getNextColumn(column);
        if (!datesDone) {
          setCell(sheetNames.cryptoCurrent, column + '2', new Date(day.time*1000));
        }
        setCell(sheetNames.cryptoCurrent, column + row, day.close);
      });
      datesDone = true;
      row++;
    }
  });

  alert('Updating crypto purchased');
  updateCryptoPurchased();
}

/**
 * Update the CryptoPurchased tab and refresh crypto returns on the dashboard
 */
function updateCryptoPurchased() {
  var done = false;
  var row = 3;
  var symbols = {};

  // grab most current price for all tracked crypto
  while (!done) {
    var symbol = getCell(sheetNames.cryptoCurrent, 'A' + row);
    if (symbol) {
      symbols[symbol] = getCell(sheetNames.cryptoCurrent, 'B' + row);
    }
    else {
      done = true;
    }
    row++;
  }

  // Update the CryptoPurchased tab with latest prices
  done = false;
  row = 3;
  while (!done) {
    var symbol = getCell(sheetNames.cryptoPurchased, 'A' + row);
    if (symbol) {
      if (symbols[symbol]) {
        setCell(sheetNames.cryptoPurchased, 'E' + row, symbols[symbol]);
      }
      else {
        setCell(sheetNames.cryptoPurchased, 'E' + row, 'n/a');
      }
    }
    else {
      done = true;
    }
    row++;
  }
}

/**
 * Update the dashboard with Crypto prices and returns
 */
function updateCryptoDashboard() {
  clearRangeValues(sheetNames.dashboard, 'J5:N14');
  done = false;
  row = 3;
  var sData = {};
  while (!done) {
    var symbol = getCell(sheetNames.cryptoPurchased, 'A' + row);
    if (symbol) {
      if (!sData[symbol]) {
        sData[symbol] = {
          shares: 0,
          price: 0,
          cost: 0,
          value: 0
        };
      }
      sData[symbol].shares += getCell(sheetNames.cryptoPurchased, 'B' + row);
      sData[symbol].price = getCell(sheetNames.cryptoPurchased, 'E' + row);
      sData[symbol].cost += getCell(sheetNames.cryptoPurchased, 'D' + row);
      sData[symbol].value += getCell(sheetNames.cryptoPurchased, 'F' + row);
      row++;
    }
    else {
      done = true;
    }
  }

  var totals = {
    cost: 0,
    value: 0,
    div: 0
  };

  row = getFirstEmptyRow(sheetNames.dashboard, 5, 'D') + 1;

  // crypto summary header
  setCell(sheetNames.dashboard, 'A' + row, 'Crypto Summary');
  mergeHorizontal(sheetNames.dashboard, 'A' + row + ':H' + row);
  setRangeFontSize(sheetNames.dashboard, 'A' + row, 14);
  setRangeBackground(sheetNames.dashboard, 'A' + row, colors.dashboardHeaderBackground);
  setRangeBorder(sheetNames.dashboard, 'A' + row, colors.dashboardOutlines);
  setRangeHorizontalAlignment(sheetNames.dashboard, 'A' + row, 'center');
  row++;

  // row headers
  setCell(sheetNames.dashboard, 'A' + row, 'Symbol');
  setCell(sheetNames.dashboard, 'B' + row, 'Shares');
  setCell(sheetNames.dashboard, 'C' + row, 'Price');
  setCell(sheetNames.dashboard, 'D' + row, 'Cost Basis');
  setCell(sheetNames.dashboard, 'E' + row, 'Value');
  setCell(sheetNames.dashboard, 'F' + row, 'Dividends');
  setCell(sheetNames.dashboard, 'G' + row, '$ Net');
  setCell(sheetNames.dashboard, 'H' + row, '% Net');
  setRangeHorizontalAlignment(sheetNames.dashboard, 'B' + row + ':H' + row, 'right');
  setRangeBorder(sheetNames.dashboard, 'A' + row + ':H' + row, colors.dashboardOutlines);
  setRangeBackground(sheetNames.dashboard, 'A' + row + ':H' + row, colors.dashboardSubheaderBackground);
  row++;

  var startRow = row;
  Object.keys(sData).forEach(function(item) {
    totals.cost += sData[item].cost;
    totals.value += sData[item].value;

    setCell(sheetNames.dashboard, 'A' + row, item);
    setCell(sheetNames.dashboard, 'B' + row, sData[item].shares);
    setCell(sheetNames.dashboard, 'C' + row, sData[item].price);
    setCell(sheetNames.dashboard, 'D' + row, sData[item].cost);
    setCell(sheetNames.dashboard, 'E' + row, sData[item].value);
    setCell(sheetNames.dashboard, 'F' + row, 0);
    setCell(sheetNames.dashboard, 'G' + row, sData[item].value - sData[item].cost);
    setCell(sheetNames.dashboard, 'H' + row, (sData[item].value - sData[item].cost) / sData[item].cost);
    setRangeBorder(sheetNames.dashboard, 'A' + row + ':' + 'H' + row, colors.dashboardOutlines);
    row++;
  });

  // add totals
  setCell(sheetNames.dashboard, 'D' + row, totals.cost);
  setCell(sheetNames.dashboard, 'E' + row, totals.value);
  setCell(sheetNames.dashboard, 'F' + row, 0);
  setCell(sheetNames.dashboard, 'G' + row, totals.value - totals.cost);
  setCell(sheetNames.dashboard, 'H' + row, (totals.value - totals.cost) / totals.cost);

  // format totals row
  setRangeBorder(sheetNames.dashboard, 'D' + row + ':' + 'H' + row, colors.dashboardOutlines);
  setRangeBackground(sheetNames.dashboard, 'D' + row + ':' + 'H' + row, colors.dashboardSubheaderBackground);
  setRangeWeight(sheetNames.dashboard, 'D' + row + ':' + 'H' + row, 'bold');

  // set up/down colors
  formatUpDown(sheetNames.dashboard, 'G' + startRow + ':H' + row);

  // save context for use later
  context.lastDashboardRowUsed = row;
  context.cryptoTotal = totals;
  context.cryptoDetail = sData;
}

/**
 * Updates the CryptoHistory tab
 */
function refreshCryptoHistory() {
  var symbols = getUniqueSymbols(sheetNames.config, 'E', 4);
  var data = getCryptoHistory(symbols, 365);
  var row = 3;
  var datesDone = false;

  clearRangeValues(sheetNames.cryptoHistory, 'A3:NC103');
  symbols.forEach(function(symbol) {
    // account for bad symbols
    if (data[symbol]) {
      var column = 'A';
      setCell(sheetNames.cryptoHistory, column + row, symbol);
      data[symbol] = data[symbol].reverse();
      data[symbol].forEach(function(day) {
        column = getNextColumn(column);
        if (!datesDone) {
          setCell(sheetNames.cryptoHistory, column + '2', new Date(day.time*1000));
        }
        setCell(sheetNames.cryptoHistory, column + row, day.close);
      });
      datesDone = true;
      row++;
    }
  });
}

/**
 * Refresh stock data on the detail tab. Once complete, also calls the refreshCrypto function. It's done here
 * because we need to pass in the starting row after having refreshed the stock detail
 */
function refreshDetail() {
  var symbols = getUniqueSymbols(sheetNames.config, 'D', 4);
  var row = 4;

  // hold latest prices for other calculations
  var latestPrices = {};

  clearRangeValues(sheetNames.detail, 'A4:O104');
  symbols.forEach(function(symbol) {
    var cRow = 3;
    var done = false;
    while (!done) {
      var cSymbol = getCell(sheetNames.stockCurrent, 'A' + cRow);
      if (!cSymbol) {
        done = true;
      }
      if (cSymbol == symbol) {
        done = true;
        setCell(sheetNames.detail, 'A' + row, cSymbol);
        setCell(sheetNames.detail, 'I' + row, cSymbol);
        var today = getCell(sheetNames.stockCurrent, 'B' + cRow);
        latestPrices[symbol] = today;
        var compare = getLastValueInRow(sheetNames.stockCurrent, 'B', cRow);
        setCell(sheetNames.detail, 'B' + row, (today - compare)/compare);
        setCell(sheetNames.detail, 'J' + row, today - compare);
      }
      cRow++;
    }

    var hRow = 3;
    done = false;
    while (!done) {
      var hSymbol = getCell(sheetNames.stockHistory, 'A' + hRow);
      if (!hSymbol) {
        done = true;
      }
      if (hSymbol == symbol) {
        done = true;

        // week change
        var compare = getCell(sheetNames.stockHistory, 'F'+ hRow);
        if (compare) {
          setCell(sheetNames.detail, 'C' + row, (latestPrices[symbol] - compare)/compare);
          setCell(sheetNames.detail, 'K' + row, latestPrices[symbol] - compare);
        }
        else {
          setCell(sheetNames.detail, 'C' + row, 'n/a');
          setCell(sheetNames.detail, 'K' + row, 'n/a');
        }

        // month change
        compare = getCell(sheetNames.stockHistory, 'V'+ hRow);
        if (compare) {
          setCell(sheetNames.detail, 'D' + row, (latestPrices[symbol] - compare)/compare);
          setCell(sheetNames.detail, 'L' + row, latestPrices[symbol] - compare);
        }
        else {
          setCell(sheetNames.detail, 'D' + row, 'n/a');
          setCell(sheetNames.detail, 'L' + row, 'n/a');
        }

        // three month change
        compare = getCell(sheetNames.stockHistory, 'BM'+ hRow);
        if (compare) {
          setCell(sheetNames.detail, 'E' + row, (latestPrices[symbol] - compare)/compare);
          setCell(sheetNames.detail, 'M' + row, latestPrices[symbol] - compare);
        }
        else {
          setCell(sheetNames.detail, 'E' + row, 'n/a');
          setCell(sheetNames.detail, 'M' + row, 'n/a');
        }

        // six month change
        compare = getCell(sheetNames.stockHistory, 'DV'+ hRow);
        if (compare) {
          setCell(sheetNames.detail, 'F' + row, (latestPrices[symbol] - compare)/compare);
          setCell(sheetNames.detail, 'N' + row, latestPrices[symbol] - compare);
        }
        else {
          setCell(sheetNames.detail, 'F' + row, 'n/a');
          setCell(sheetNames.detail, 'N' + row, 'n/a');
        }

        // twelve month change
        compare = getCell(sheetNames.stockHistory, 'IQ'+ hRow);
        if (compare) {
          setCell(sheetNames.detail, 'G' + row, (latestPrices[symbol] - compare)/compare);
          setCell(sheetNames.detail, 'O' + row, latestPrices[symbol] - compare);
        }
        else {
          setCell(sheetNames.detail, 'G' + row, 'n/a');
          setCell(sheetNames.detail, 'O' + row, 'n/a');
        }

        row++;
      }
      hRow++;
    }
  });

  refreshCryptoDetail(row);
}

/**
 * Refresh crypto data on the detail tab
 */
function refreshCryptoDetail(row) {
  var symbols = getUniqueSymbols(sheetNames.config, 'E', 4);

  // hold latest prices for other calculations
  var latestPrices = {};

  symbols.forEach(function(symbol) {
    var cRow = 3;
    var done = false;
    while (!done) {
      var cSymbol = getCell(sheetNames.cryptoCurrent, 'A' + cRow);
      if (!cSymbol) {
        done = true;
      }
      if (cSymbol == symbol) {
        done = true;
        setCell(sheetNames.detail, 'A' + row, cSymbol);
        setCell(sheetNames.detail, 'I' + row, cSymbol);
        var today = getCell(sheetNames.cryptoCurrent, 'B' + cRow);
        latestPrices[symbol] = today;
        var compare = getLastValueInRow(sheetNames.cryptoCurrent, 'B', cRow);
        setCell(sheetNames.detail, 'B' + row, (today - compare)/compare);
        setCell(sheetNames.detail, 'J' + row, today - compare);
      }
      cRow++;
    }

    var hRow = 3;
    done = false;
    while (!done) {
      var hSymbol = getCell(sheetNames.cryptoHistory, 'A' + hRow);
      if (!hSymbol) {
        done = true;
      }
      if (hSymbol == symbol) {
        done = true;

        // week change
        var compare = getCell(sheetNames.cryptoHistory, 'H'+ hRow);
        if (compare) {
          setCell(sheetNames.detail, 'C' + row, (latestPrices[symbol] - compare)/compare);
          setCell(sheetNames.detail, 'K' + row, latestPrices[symbol] - compare);

        }
        else {
          setCell(sheetNames.detail, 'C' + row, 'n/a');
          setCell(sheetNames.detail, 'K' + row, 'n/a');
        }

        // month change
        compare = getCell(sheetNames.cryptoHistory, 'AF'+ hRow);
        if (compare) {
          setCell(sheetNames.detail, 'D' + row, (latestPrices[symbol] - compare)/compare);
          setCell(sheetNames.detail, 'L' + row, latestPrices[symbol] - compare);
        }
        else {
          setCell(sheetNames.detail, 'D' + row, 'n/a');
          setCell(sheetNames.detail, 'L' + row, 'n/a');
        }

        // three month change
        compare = getCell(sheetNames.cryptoHistory, 'CO'+ hRow);
        if (compare) {
          setCell(sheetNames.detail, 'E' + row, (latestPrices[symbol] - compare)/compare);
          setCell(sheetNames.detail, 'M' + row, latestPrices[symbol] - compare);
        }
        else {
          setCell(sheetNames.detail, 'E' + row, 'n/a');
          setCell(sheetNames.detail, 'M' + row, 'n/a');
        }

        // six month change
        compare = getCell(sheetNames.cryptoHistory, 'EV'+ hRow);
        if (compare) {
          setCell(sheetNames.detail, 'F' + row, (latestPrices[symbol] - compare)/compare);
          setCell(sheetNames.detail, 'N' + row, latestPrices[symbol] - compare);
        }
        else {
          setCell(sheetNames.detail, 'F' + row, 'n/a');
          setCell(sheetNames.detail, 'N' + row, 'n/a');
        }

        // twelve month change
        compare = getCell(sheetNames.cryptoHistory, 'NB'+ hRow);
        if (compare) {
          setCell(sheetNames.detail, 'G' + row, (latestPrices[symbol] - compare)/compare);
          setCell(sheetNames.detail, 'O' + row, latestPrices[symbol] - compare);
        }
        else {
          setCell(sheetNames.detail, 'G' + row, 'n/a');
          setCell(sheetNames.detail, 'O' + row, 'n/a');
        }

        row++;
      }
      hRow++;
    }
  });
}

/**
 * Removes existing charts from the dashboard and re-creates them for the four categories of charts
 */
function refreshDashboardCharts() {
  removeChartsFromSheet(sheetNames.dashboard);

  alert('Updating historical stock charts');
  refreshYearStockCharts();

  alert('Updating day stock charts');
  refreshDayStockCharts();

  alert('Updating historical crypto charts');
  refreshYearCryptoCharts();

  alert('Updating day crypto charts');
  refreshDayCryptoCharts();
}

/**
 * Refreshes the year-long stock charts
 */
function refreshYearStockCharts() {
  var historySheet = getSheetByName(sheetNames.stockHistory);
  var done = false;
  var row = 3;

  setCell(sheetNames.dashboard, 'C' + (context.lastDashboardRowUsed + 2), 'One Year');
  setRangeColor(sheetNames.dashboard, 'C' + (context.lastDashboardRowUsed + 2), colors.stockDayChart);
  setRangeFontSize(sheetNames.dashboard, 'C' + (context.lastDashboardRowUsed + 2), 14);
  setRangeHorizontalAlignment(sheetNames.dashboard, 'C' + (context.lastDashboardRowUsed + 2), 'center');

  var options = {
    title: '',
    row: context.lastDashboardRowUsed + 3,
    column: 1,
    offsetLeft: 5,
    offsetTop: 5,
    width: 280,
    height: 180,
    backgroundColor: colors.dashboardBackground,
    colors: [colors.stockDayChart],
    titleColor: colors.chartTitleColor,
    yAxisLabelColor: colors.chartYAxisLabel,
    gridLineColor: colors.chartgridLineColor
  };

  while (!done) {
    var symbol = getCell(sheetNames.stockHistory, 'A' + row);
    if (symbol) {
      var range = historySheet.getRange('B' + row + ':IQ' + row);
      options.title = symbol;
      areaChart(sheetNames.dashboard, range, options);
      row++;
      options.row += 9;
    }
    else {
      done = true;
    }
  }
}

/**
 * Refreshes the daily stock charts
 */
function refreshDayStockCharts() {
  var currentSheet = getSheetByName(sheetNames.stockCurrent);
  var done = false;
  var row = 3;

  setCell(sheetNames.dashboard, 'F' + (context.lastDashboardRowUsed + 2), 'Today');
  setRangeColor(sheetNames.dashboard, 'F' + (context.lastDashboardRowUsed + 2), colors.stockHourChart);
  setRangeFontSize(sheetNames.dashboard, 'F' + (context.lastDashboardRowUsed + 2), 14);
  setRangeHorizontalAlignment(sheetNames.dashboard, 'F' + (context.lastDashboardRowUsed + 2), 'center');

  var options = {
    title: '',
    row: context.lastDashboardRowUsed + 3,
    column: 5,
    offsetLeft: 0,
    offsetTop: 5,
    width: 280,
    height: 180,
    backgroundColor: colors.dashboardBackground,
    colors: [colors.stockHourChart],
    titleColor: colors.chartTitleColor,
    yAxisLabelColor: colors.chartYAxisLabel,
    gridLineColor: colors.chartgridLineColor
  };
  while (!done) {
    var symbol = getCell(sheetNames.stockCurrent, 'A' + row);
    if (symbol) {
      var range = currentSheet.getRange('B' + row + ':CC' + row);
      options.title = symbol;
      areaChart(sheetNames.dashboard, range, options);
      row++;
      options.row += 9;
    }
    else {
      done = true;
    }
  }
}

/**
 * Refreshes the year-long crypto charts
 */
function refreshYearCryptoCharts() {
  var historySheet = getSheetByName(sheetNames.cryptoHistory);
  var done = false;
  var row = 3;

  setCell(sheetNames.dashboard, 'L' + (context.lastDashboardRowUsed + 2), 'One Year');
  setRangeColor(sheetNames.dashboard, 'L' + (context.lastDashboardRowUsed + 2), colors.cryptoDayChart);
  setRangeFontSize(sheetNames.dashboard, 'L' + (context.lastDashboardRowUsed + 2), 14);
  setRangeHorizontalAlignment(sheetNames.dashboard, 'L' + (context.lastDashboardRowUsed + 2), 'center');

  var options = {
    title: '',
    row: context.lastDashboardRowUsed + 3,
    column: 10,
    offsetLeft: 5,
    offsetTop: 5,
    width: 280,
    height: 180,
    backgroundColor: colors.dashboardBackground,
    colors: [colors.cryptoDayChart],
    titleColor: colors.chartTitleColor,
    yAxisLabelColor: colors.chartYAxisLabel,
    gridLineColor: colors.chartgridLineColor
  };
  while (!done) {
    var symbol = getCell(sheetNames.cryptoHistory, 'A' + row);
    if (symbol) {
      var range = historySheet.getRange('B' + row + ':NC' + row);
      options.title = symbol;
      areaChart(sheetNames.dashboard, range, options);
      row++;
      options.row += 9;
    }
    else {
      done = true;
    }
  }
}

/**
 * Refreshes the daily crypto charts
 */
function refreshDayCryptoCharts() {
  var historySheet = getSheetByName(sheetNames.cryptoCurrent);
  var done = false;
  var row = 3;

  setCell(sheetNames.dashboard, 'O' + (context.lastDashboardRowUsed + 2), 'Today');
  setRangeColor(sheetNames.dashboard, 'O' + (context.lastDashboardRowUsed + 2), colors.cryptoHourChart);
  setRangeFontSize(sheetNames.dashboard, 'O' + (context.lastDashboardRowUsed + 2), 14);
  setRangeHorizontalAlignment(sheetNames.dashboard, 'O' + (context.lastDashboardRowUsed + 2), 'center');

  var options = {
    title: '',
    row: context.lastDashboardRowUsed + 3,
    column: 14,
    offsetLeft: 5,
    offsetTop: 5,
    width: 280,
    height: 180,
    backgroundColor: colors.dashboardBackground,
    colors: [colors.cryptoHourChart],
    titleColor: colors.chartTitleColor,
    yAxisLabelColor: colors.chartYAxisLabel,
    gridLineColor: colors.chartgridLineColor
  };
  while (!done) {
    var symbol = getCell(sheetNames.cryptoCurrent, 'A' + row);
    if (symbol) {
      var range = historySheet.getRange('B' + row + ':RN' + row);
      options.title = symbol;
      areaChart(sheetNames.dashboard, range, options);
      row++;
      options.row += 9;
    }
    else {
      done = true;
    }
  }
}

/**
 * Get the unique symbols given a sheet name, column, and starting row and returns
 * in an array: ['MSFT', 'GOOG', 'AAPL']
 */
function getUniqueSymbols(sheetName, column, startRow) {
  var done = false;
  var row = startRow;
  var symbols = [];
  while (!done) {
    var symbol = getCell(sheetName,column + row);
    if (symbol && symbol.length > 0) {
      symbols.push(symbol);
      row++;
    }
    else {
      done = true;
    }
  }

  return symbols.filter(function(item, pos, self) {
    return self.indexOf(item) == pos;
  });
}

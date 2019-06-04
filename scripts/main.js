/**
 * This script uses 3rd-party API's for stock and crypto pricing data to keep assets
 * up to date in the Ticker spreadsheet
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
 * Refresh all data and update the dashboard
 */
function refreshAll() {
  clearRange(sheetNames.debug, 'A2', 'B50');
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

  alert('Updates complete');
}

/**
 * Refresh all real-time prices and update the dashboard
 */
function refreshCurrent() {
  clearRange(sheetNames.debug, 'A2', 'B50');
  alert('Updating current stock prices');
  refreshStockCurrent();

  alert('Updating dashboard stock returns');
  updateStockDashboard();

  alert('Updating current crypto prices');
  refreshCryptoCurrent();

  alert('Updating dashboard crypto returns');
  updateCryptoDashboard();

  alert('Updates complete');
}

/**
 * Refreshes the stock pricing - current trading day and up to the minute price for everything in the StockCurrent
 * tab. Then updates latest price for anywhere else necessary
 *
 */
function refreshStockCurrent() {
  var symbols = getUniqueSymbols(sheetNames.stockCurrent, 'A', 3);
  var data = getCurrentStockPrices(symbols, 80);

  var done = false;
  var row = 3;
  var datesDone = false;

  while (!done) {
    var symbol = getCell(sheetNames.stockCurrent, 'A' + row);
    if (symbol && symbol.length > 0) {
      var column = 'A';
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
    else {
      done = true;
    }
  }
  alert('Updating stock purchased');
  updateStockPurchased();
}

/**
 * Using the prices from StockCurrent, update the StockPurchased tab
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
  clearRange(sheetNames.dashboard, 'A5', 'F14');
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
  row = 2;
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

  row = 5;
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
    row++;
  });

  // throw totals in
  setCell(sheetNames.dashboard, 'D' + row, totals.cost);
  setCell(sheetNames.dashboard, 'E' + row, totals.value);
  setCell(sheetNames.dashboard, 'F' + row, totals.div);
}

/**
 * Updates the StockHistory tab
 */
function refreshStockHistory() {
  var symbols = getUniqueSymbols(sheetNames.stockHistory, 'A', 3);
  var data = getHistoricalStockPrices(symbols, 250);

  var done = false;
  var row = 3;
  var datesDone = false;

  while (!done) {
    var symbol = getCell(sheetNames.stockHistory, 'A' + row);
    if (symbol) {
      var column = 'A';
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
    else {
      done = true;
    }
  }
}

/**
 * Updates the CryptoCurrent tab
 */
function refreshCryptoCurrent() {
  var symbols = getUniqueSymbols(sheetNames.cryptoCurrent, 'A', 3);
  var data = getCryptoCurrent(symbols, 480);

  var done = false;
  var row = 3;
  var datesDone = false;

  while (!done) {
    var symbol = getCell(sheetNames.cryptoCurrent, 'A' + row);
    if (symbol) {
      var column = 'A';
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
    else {
      done = true;
    }
  }
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
  clearRange(sheetNames.dashboard, 'J5', 'N14');
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

  row = 5;
  Object.keys(sData).forEach(function(item) {
    totals.cost += sData[item].cost;
    totals.value += sData[item].value;

    setCell(sheetNames.dashboard, 'J' + row, item);
    setCell(sheetNames.dashboard, 'K' + row, sData[item].shares);
    setCell(sheetNames.dashboard, 'L' + row, sData[item].price);
    setCell(sheetNames.dashboard, 'M' + row, sData[item].cost);
    setCell(sheetNames.dashboard, 'N' + row, sData[item].value);
    row++;
  });

  // throw totals in
  setCell(sheetNames.dashboard, 'M' + row, totals.cost);
  setCell(sheetNames.dashboard, 'N' + row, totals.value);
}

/**
 * Updates the CryptoHistory tab
 */
function refreshCryptoHistory() {
  var symbols = getUniqueSymbols(sheetNames.cryptoHistory, 'A', 3);
  var data = getCryptoHistory(symbols, 365);

  var done = false;
  var row = 3;
  var datesDone = false;

  while (!done) {
    var symbol = getCell(sheetNames.cryptoHistory, 'A' + row);
    if (symbol) {
      var column = 'A';
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

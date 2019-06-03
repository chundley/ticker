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
 * Refreshes the stock pricing - current trading day and up to the minute price for everything in the StockCurrent
 * tab. Then updates latest price for anywhere else necessary
 *
 */
function refreshStockCurrent() {
  alert('Updating current stock prices');
  clearRange(sheetNames.debug, 'A2', 'B50');
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
  alert('Update Complete');
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

  // Update the dashboard with cost basis and returns for all stock
  alert('Updating dashboard');
  clearRange(sheetNames.dashboard, 'A5', 'G50');
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

  row = 5;
  Object.keys(sData).forEach(function(item) {
    setCell(sheetNames.dashboard, 'A' + row, item);
    setCell(sheetNames.dashboard, 'B' + row, sData[item].shares);
    setCell(sheetNames.dashboard, 'C' + row, sData[item].price);
    setCell(sheetNames.dashboard, 'D' + row, sData[item].cost);
    setCell(sheetNames.dashboard, 'E' + row, sData[item].value);
    setCell(sheetNames.dashboard, 'F' + row, sData[item].div);
    row++;
  });
}

/**
 * Updates the StockHistory tab
 */
function refreshStockHistory() {
  alert('Updating stock history');
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
  alert('Stock history finished');
}

/**
 * Update all
 */
function updateAll() {

  // clear debug and other ephemeral output

  //setCell('Dashboard', 'I3', '');
  //setCell('Dashboard', 'I4', '');
  //setCell('Stock', 'K2', '');
  //setCell('Stock', 'K3', '');
  //setCell('Crypto', 'K2', '');
  //setCell('Crypto', 'K3', '');
  //setCell('Debug', 'B2', '');
  //setCell('Debug', 'B3', '');

  // start updates
  //setCell('Dashboard', 'I3', 'Updating Stock Prices...');
  //updateStock();

  //setCell('Dashboard', 'I3', 'Updating Stock Returns...');
  //updateStockReturnsBySymbol();

  //setCell('Dashboard', 'I3', 'Updating Stock Watch...');
  //updateStockHistory();

  //setCell('Dashboard', 'I3', 'Updating Crypto Prices...');
  //updateCrypto();

  //setCell('Dashboard', 'I3', 'Updating Crypto Returns...');
  //updateCryptoReturnsBySymbol();

  //setCell('Dashboard', 'I3', 'Updating Crypto Watch...');
  //updateCryptoHistory();

  //setCell('Dashboard', 'I3', 'Finished Updating');
}





function refreshCryptoPricesOLD() {
  alert('Updating crypto history');
  var symbols = getUniqueSymbols(sheetNames.cryptoHistory, 'A', 3);
  var prices = getCryptoPrices(symbols, 1);
  setCell('Debug', 'B3', JSON.stringify(prices));

  // Rate limits - unclear why we hit this
  if (prices.Response && prices.Response == 'Error') {
    setCell('Crypto', 'K2', 'Update Finished');
    setCell('Dashboard', 'I4', 'Error when requesting Crypto Updates');
    setCell('Crypto', 'K3', 'Error when requesting Crypto Updates');
    return;
  }

  var done = false;
  var row = 3;
  while (!done) {
    var symbol = getCell('Crypto', 'A' + row);
    if (symbol && symbol.length > 0) {
      setCell('Crypto', 'E' + row, prices[symbol].USD);
      row++;
    }
    else {
      done = true;
    }
  }
  setCell('Crypto', 'K2', 'Update Finished');
}

/**
 * Updates Dashboard crypto returns by symbol
 */
function refreshCryptoHistoryOLD() {
  var sData = {};
  var done = false;
  var row = 3;
  while (!done) {
    var symbol = getCell('Crypto', 'A' + row);
    if (symbol && symbol.length > 0) {
      if (!sData[symbol]) {
        sData[symbol] = {
          shares: 0,
          price: 0,
          cost: 0,
          value: 0
        };
      }
      sData[symbol].shares += getCell('Crypto', 'B' + row);
      sData[symbol].price = getCell('Crypto', 'E' + row);
      sData[symbol].cost += getCell('Crypto', 'D' + row);
      sData[symbol].value += getCell('Crypto', 'F' + row);
      row++;
    }
    else {
      done = true;
    }
  }

  var outputRow = 11;
  Object.keys(sData).forEach(function(item) {
    setCell('Dashboard', 'J' + outputRow, item);
    setCell('Dashboard', 'K' + outputRow, sData[item].shares);
    setCell('Dashboard', 'L' + outputRow, sData[item].price);
    setCell('Dashboard', 'M' + outputRow, sData[item].cost);
    setCell('Dashboard', 'N' + outputRow, sData[item].value);
    outputRow++;
  });
}

function refreshCryptoCurrent() {
  alert('Updating crypto current prices');
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
  alert('Crypto current prices updated');
}

/**
 * Updates the CryptoHistory tab
 */
function refreshCryptoHistory() {
  alert('Updating crypto historical prices');
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
  alert('Crypto history complete');
}

function cleanDashboardPositions() {
  clearRange(sheetNames.dashboard, 'A5', 'F14');
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

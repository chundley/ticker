/**
 * TICKER v1.1
 *
 * Interface to the CryptoCompare API for Google Sheets implementation
 *
 * https://min-api.cryptocompare.com/documentation
 *
 * Rate limit: 50/sec 2500/minute, 25K/hour, 100K/month
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


/**
 * @private
 * Make a request to the Crypto Compare API
 * @param {string} path - the url path of the request
 * @param {Object} params - key/value pairs to be mapped as query string parameters on the requst
 * @return {Object} - the resulting data set from the request
 */
function cryptoRequest(path, params) {
  var endpoint = 'https://min-api.cryptocompare.com';
  var headers = {
    'Authorization': 'Apikey ' + getCell('Config', 'B6')
  }

  var options = {
    headers: headers,
    muteHttpExceptions: true
  };

  var url = endpoint + path;
  if (params) {
    var kv = [];
    for (var k in params) {
      kv.push(k + '=' + encodeURIComponent(params[k]));
    }
    url += '?' + kv.join('&');
  }

  var response = UrlFetchApp.fetch(url, options);
  var json = response.getContentText();

  try {
    return JSON.parse(json);
  }
  catch (e) {
    return json;
  }
}

/**
 * Get crypto prices in one-minute interval for x number of data points. We can only request one symbol
 * at a time, so this function will make multiple API requests and combine the data. When requesting
 * https://min-api.cryptocompare.com/documentation?key=Historical&cat=dataHistominute
 * @param {string} symbols - an array of symbols to request pricing on, eg ['BTC','ETH','LTC']
 * @param {number} limit - the number of intervals to request, eg 480 for 8 hours of market data (max=600)
 * @return {Object} - the resulting data set
 * @example
 * var data = getHistoricalStockPrices(['BTC','ETH'], 480);
 * // returns
 * {
 *   "BTC": [
 *    {
 *      time: 1561144080
 *      close: 7788.24
 *      high: 7792.6
 *      low: 7788.24
 *      open: 7792.6
 *      volumefrom: 0.212
 *      volumeto: 1652.3
 *    }, // 480 total values
 *  ],
 *  "ETH": [...]
 * }
 */
function getCryptoCurrent(symbols, limit) {
  if (limit > 600) {
    limit = 600;
  }

  var results = {};
  symbols.forEach(function(symbol) {
    var params = {
      fsym: symbol,
      tsym: 'USD',
      limit: limit
    };

    var resp = cryptoRequest('/data/histominute', params);
    debugMessage('CryptoCompare /data/histominute', JSON.stringify(resp).substring(0,255));
    results[symbol] = resp.Data;
  });
  return results;
}

/**
 * Get daily closing price for crypto symbols for x days
 * https://min-api.cryptocompare.com/documentation?key=Historical&cat=dataHistoday
 * @param {string} symbols - an array of symbols to request pricing on, eg ['BTC','ETH','LTC']
 * @param {number} limit - the number of intervals to request, eg 365 for 1 year of market data (max=600)
 * @return {Object} - the resulting data set
 * @example
 * var data = getHistoricalStockPrices(['BTC','ETH'], 365);
 * // returns
 * {
 *   "BTC": [
 *    {
 *      time: 1561144080
 *      close: 7788.24
 *      high: 7792.6
 *      low: 7788.24
 *      open: 7792.6
 *      volumefrom: 0.212
 *      volumeto: 1652.3
 *    }, // 365 total values
 *  ],
 *  "ETH": [...]
 * }
 */
function getCryptoHistory(symbols, limit) {
  if (limit > 600) {
    limit = 600;
  }

  var results = {};
  symbols.forEach(function(symbol) {
    var params = {
      fsym: symbol,
      tsym: 'USD',
      limit: limit
    };

    var resp = cryptoRequest('/data/histoday', params);
    debugMessage('CryptoCompare /data/histoday', JSON.stringify(resp).substring(0,255));
    results[symbol] = resp.Data;
  });
  return results;
}

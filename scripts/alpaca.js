/**
 * TICKER v1.0
 *
 * Interface to the Alpaca API
 *
 * Docs - https://docs.alpaca.markets/api-documentation/
 *
 * Rate limit: 200 requests per minute
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
 * @private
 * Request to the standard Alpaca API endpoints
 * https://docs.alpaca.markets/api-documentation/
 * @param {string} path - the url path of the request
 * @param {Object} params - key/value pairs to be mapped as query string parameters on the requst
 * @return {Object} - the resulting data set from the request
 * @example
 * var data = alpacaApiRequest('/v1/assets/' + symbol);
 */
function alpacaApiRequest(path, params) {
  return alpacaRequest('https://api.alpaca.markets', path, params);
}

/**
 * @private
 * Request to the Alpaca data API endpoints
 * https://docs.alpaca.markets/api-documentation/web-api/market-data/bars/
 * @param {string} path - the url path of the request
 * @param {Object} params - key/value pairs to be mapped as query string parameters on the requst
 * @return {Object} - the resulting data set from the request
 * @example
 * var data = alpacaDataRequest('/v1/bars/5Min', { symbols: 'MSFT,GOOG', limit: 100});
 */
function alpacaDataRequest(path, params) {
  return alpacaRequest('https://data.alpaca.markets', path, params);
}

/**
 * @private
 * Execute request to one of the Alpaca APIs. We've abstracted away the particular endpoint and methods of calling to make the public
 * functions easier to understand and use
 * @param {string} path - the url path of the request
 * @param {Object} params - key/value pairs to be mapped as query string parameters on the requst
 * @return {Object} - the resulting data set from the request
 */
function alpacaRequest(endpoint, path, params) {
  var headers = {
    'APCA-API-KEY-ID': getCell('Config', 'B4'),
    'APCA-API-SECRET-KEY': getCell('Config', 'B5')
  };

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
 * Gets basic data about a stock
 * https://docs.alpaca.markets/api-documentation/web-api/assets/#get-an-asset
 * @param {string} symbol - the symbol of the stock/asset to retrieve
 * @return {Object} - the result
 * @example
 * var data = getAsset('AAPL');
 * // returns
 * {
 *   id: "904837e3-3b76-47ec-b432-046db621571b",
 *   asset_class: "us_equity",
 *   exchange: "NASDAQ",
 *   symbol: "AAPL",
 *   status: "active",
 *   tradable: true
 * }
 */
function getAsset(symbol) {
  var resp = alpacaApiRequest('/v1/assets/' + symbol);
  debugMessage('Alpaca /v1/assets/', JSON.stringify(resp).substring(0,255));
  return resp;
}

/**
 * Gets daily closing price for a set of symbols for x number of days
 * https://docs.alpaca.markets/api-documentation/web-api/market-data/bars/
 * @param {string} symbols - an array of symbols to request pricing on, eg ['AAPL','MSFT','GOOG']
 * @param {number} days - the number of days to request, eg 90 for 90 days of market data
 * @return {Object} - the resulting data set
 * @example
 * var data = getHistoricalStockPrices(['AAPL','MSFT'], 90);
 * // returns
 * {
 *   "AAPL": [
 *    {
 *      "t": 1544129220,
 *      "o": 172.26,
 *      "h": 172.3,
 *      "l": 172.16,
 *      "c": 172.18,
 *      "v": 3892,
 *    }, // 90 total values
 *  ],
 *  "MSFT": [...]
 * }
 */
function getHistoricalStockPrices(symbols, days) {
  var params = {
    symbols: symbols.join(','),
    limit: days
  };

  var resp = alpacaDataRequest('/v1/bars/day', params);
  debugMessage('Alpaca /v1/bars/day', JSON.stringify(resp).substring(0,255));
  return resp;
}

/**
 * Get stock price at five minute intervals for a set of symbols for x number of data points
 * https://docs.alpaca.markets/api-documentation/web-api/market-data/bars/
 * @param {string} symbols - an array of symbols to request pricing on, eg ['AAPL','MSFT','GOOG']
 * @param {number} count - the number of data points to request, eg 288 for 1 day of market data
 * @return {Object} - the resulting data set
 * @example
 * var data = getHistoricalStockPrices(['AAPL','MSFT'], 288);
 * // returns
 * {
 *   "AAPL": [
 *    {
 *      "t": 1544129220,
 *      "o": 172.26,
 *      "h": 172.3,
 *      "l": 172.16,
 *      "c": 172.18,
 *      "v": 3892,
 *    }, // 288 total values
 *  ],
 *  "MSFT": [...]
 * }
 */
function getCurrentStockPrices(symbols, count) {
  var params = {
    symbols: symbols.join(','),
    limit: count
  };

  var resp = alpacaDataRequest('/v1/bars/5Min', params);
  debugMessage('Alpaca /v1/bars/5Min', JSON.stringify(resp).substring(0,255));
  return resp;
}

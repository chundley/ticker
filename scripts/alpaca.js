/**
 * Interface to the Alpaca API for Google Sheets implementation
 *
 * https://docs.alpaca.markets/api-documentation/
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
 * Standard request to the Alpaca API
 */
function alpacaApiRequest(path, params) {
  return alpacaRequest('https://api.alpaca.markets', path, params);
}

/**
 * Request to the Alpaca data API
 */
function alpacaDataRequest(path, params) {
  return alpacaRequest('https://data.alpaca.markets', path, params);
}

/**
 * Execute request to one of the Alpaca APIs
 */
function alpacaRequest(endpoint, path, params) {
  var headers = {
    'APCA-API-KEY-ID': getCell('Config', 'B1'),
    'APCA-API-SECRET-KEY': getCell('Config', 'B2')
  };

  var options = {
    headers: headers,
    muteHttpExceptions: true
  };

  var url = endpoint + path;
  if (params) {
    if (params.qs) {
      var kv = [];
      for (var k in params.qs) {
        kv.push(k + '=' + encodeURIComponent(params.qs[k]));
      }
      url += '?' + kv.join('&');
      delete params.qs
    }
    for (var k in params) {
      options[k] = params[k];
    }
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
 */
function getAsset(symbol) {
  var resp = alpacaApiRequest('/v1/assets/' + symbol);
  return JSON.stringify(resp);
}

/**
 * Gets price data for a set of symbols for x number of days
 */
function getHistoricalStockPrices(symbols, days) {
  var params = {
    qs: {
      symbols: symbols.join(','),
      limit: days
    }
  };

  var resp = alpacaDataRequest('/v1/bars/day', params);
  debugMessage('Alpaca /v1/bars/day', JSON.stringify(resp).substring(0,255));
  return resp;
}

function getCurrentStockPrices(symbols, count) {
  var params = {
    qs: {
      symbols: symbols.join(','),
      limit: count
    }
  };

  var resp = alpacaDataRequest('/v1/bars/5Min', params);
  debugMessage('Alpaca /v1/bars/5Min', JSON.stringify(resp).substring(0,255));
  return resp;
}

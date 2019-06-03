/**
 * Interface to the CryptoCompare API for Google Sheets implementation
 *
 * https://min-api.cryptocompare.com/documentation
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
 * Make a request to the Crypto Compare API
 */
function cryptoRequest(path, params) {
  var endpoint = 'https://min-api.cryptocompare.com';
  var headers = {
    'Authorization': 'Apikey ' + getCell('Config', 'B3')
  }

  var options = {
    headers: headers,
    muteHttpExceptions: true
  };

  var url = endpoint + path;
  if (params) {
    if (params.qs) {
      var kv = [];
      for (var k in params.qs) {
        kv.push(k + '=' + params.qs[k]);
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
 * Get current price for an array of Crypto symbols
 */
function getCryptoPrices(symbols) {
  var params = {
    qs: {
      fsyms: symbols.join(','),
      tsyms: 'USD'
    }
  };

  var resp = cryptoRequest('/data/pricemulti', params);
  debugMessage('CryptoCompare /v1/bars/5Min', JSON.stringify(resp).substring(0,255));
  return resp;
}

function getCryptoCurrent(symbols, limit) {
  var results = {};
  symbols.forEach(function(symbol) {
    var params = {
      qs: {
        fsym: symbol,
        tsym: 'USD',
        limit: limit
      }
    };

    var resp = cryptoRequest('/data/histominute', params);
    debugMessage('CryptoCompare /data/histominute', JSON.stringify(resp).substring(0,255));
    results[symbol] = resp.Data;
  });
  return results;
}

function getCryptoHistory(symbols, limit) {
  var results = {};
  symbols.forEach(function(symbol) {
    var params = {
      qs: {
        fsym: symbol,
        tsym: 'USD',
        limit: limit
      }
    };

    var resp = cryptoRequest('/data/histoday', params);
    debugMessage('CryptoCompare /data/histoday', JSON.stringify(resp).substring(0,255));
    results[symbol] = resp.Data;
  });
  return results;
}

## Ticker
This project started as a small utility to track investments in the Stock Market and Cryptocurrency side-by-side with the ability to refresh prices in real-time.

It turned into an experiment to learn some programmatic interaction with [Google Sheets](https://www.google.com/sheets/about/) and dynamically loading to support long-term investing and refreshing.

### Requirements
Out of the box the worksheet supports integrations with two different platforms: Alpaca for stock market data and CryptoCompare for crypto prices. Data is pulled on-demand and refreshes the underlying data sheets which in turn triggers a dashboard update.

If you set up accounts at both of these default vendors you'll be able to use the worksheet as-is. If you have a different API you'll need to do some work on the requests and data model.

#### Alpaca
Alpaca is a robust API that supports not only market pricing data, but both fake and real API-based trading.

* Web site: [https://alpaca.markets](https://alpaca.markets/)
* API docs: [https://docs.alpaca.markets/api-documentation/web-api/](https://docs.alpaca.markets/api-documentation/web-api/)

It's free and the only rate limit is 200 requests per minute. This worksheet only makes two API requests every time it's refreshed.

#### CryptoCompare
CryptoCompare has a free version that provides more than enough data and rate limits for a small project like this.

Web site: [https://www.cryptocompare.com/](https://www.cryptocompare.com/)
API docs: [https://min-api.cryptocompare.com/](https://min-api.cryptocompare.com/)

Free edition rate limits: 50/sec | 2500/minute | 25K/hour | 100K/month

### Project structure
The project is a Google sheet with javascript responsible for API requests and refreshing various tabs to support the dashboard. It's meant to be an extensible template you can modify to suit your needs.

To manage changes and pull requests I'll create versions of the worksheet and tag release versions of the Javascript source to keep things in sync. The tagged releases will contain a link to the updated worksheet. The worksheet will be read-only since it won't contain API keys and can't be refreshed.

For personal use and experimenting:

1. Make a copy of the worksheet via the menu: File -> Make a copy
2. Add your API keys to the **Debug** tab
3. Add the symbols you want to track to the **Debug** tab
4. Add your stock purchases, stock dividends, and crypto purchased data to their corresponding tabs
5. Click the **Refresh All History** button on the Dashboard tab

### Tabs and purpose
The tabs in this worksheet act like database tables for the most part.

#### Config
The config tab is where you'll add your API keys once you've set up accounts with Alpaca and CryptoCompare. This is also where you'll create a list of stock and crypto symbols you'd like to pull data for.

> Note: you can put whatever symbols you'd like to track, even those you haven't purchased. The rows and columns from the config tab are hard-coded in the supporting script, so be mindful to keep things in the correct cells and/or columns.

#### Dashboard
The dashboard is a summary of all positions, dividends earned (if any), and two trend charts for each symbol. A one-year trend alongside a daily trend. Of course these can be tuned to your liking with some minor tweaks to the script and API calls.

#### Detail
This tab shows additional detail to help determine best days/times to buy or sell. The current price for each symbol is compared to the price from one year ago, six months ago, three months ago, one month ago, one week ago, and one day ago. Again, this can be customized to your preference via minor tweaking.

#### StockPurchased
List all of your stock purchases in this tab. The **Last Price** column is the only one updated by the script. You should manually fill in columns A, B, C, and I. The rest of the columns should be copied down since they contain formulas. Whenever the script is run, the last price is populated with updated data which refreshes value and returns on the dashboard and detail tabs.

#### StockDividend
If you are earning dividends on any purchased stock, manually fill in the dividends whenever they are paid. This part is tricky to automate because the worksheet would need access to your brokerage account. I chose to keep this manual since it only needs an update once a quarter. You don't need to track dividends if you don't want to, everything will still work without data in this tab.

#### CryptoPurchased
Similar to the StockPurchased tab, this tab tracks all of your crypto purchases in order to track gains and losses over time.

#### StockCurrent
This tab tracks prices for all symbols in the config tab. It tracks back up to 24 hours worth of prices in five minute intervals. These prices are used for keeping the current price for net gain/loss calculations as well as single day trend charts.

#### StockHistory
This tab tracks the daily closing price for all symbols in the config tab. It tracks back up to one year of prices and is used for several data points in the dashboard and detail tabs.

#### CryptoCurrent
This tab tracks prices for all crypto symbols in the config tab. It tracks back up to 8 hours worth of prices in one minute intervals. These prices are used for keeping the current price for net gain/loss calculations as well as single day trend charts.

#### CryptoHistory
This tab tracks the daily closing price for all crypto symbols in the config tab. It tracks back up to one year of prices and is used for several data points in the dashboard and detail tabs.

#### Debug
To help troubleshoot problems, the debug tab logs all API calls and the results of the call. Each time the worksheet is refreshed the debug tab is cleared and each call is logged sequentially.

### License
Copyright 2019 Chris Hundley

MIT LICENSE
 Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the 'Software'), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED 'AS IS', WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

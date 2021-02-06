# yahoofinancials-excel-autoscraper
short program that utilizes pandas, yahoofinancials, sys, and date time to read in an excel file with a list of stock tickers and desired metrics, and then exports the filled out excel spreadsheet with the most recent market data (complete with a timestamp at time of export) to the same folder. 

At the current moment, the program scrapes current price, 50-day moving average, 52-week high,  52-week low,Beta,10-day average volume, 90-day average volume, Market Cap, Price-to-Earnings Ratio (P/E), Payout Rate, Earnings per Share (EPS), Earnings Before Income Taxes (EBIT), Gross Profit, and Book Value. 

Ideally, I would like to expand this stock screener in three major ways:

1) speed up the scraping of the data (currently the program takes ~5 seconds per ticker), 
2) Add additional information, perhaps even creating different excel tabs for each stock, complete with charts, historical data and projections
3) Add a section that would gather general "sentiment" surrounding the company, which would be achieved by scraping company news off a reputable market news site (likely either Wall Street Journal, Bloomberg, or MarketWatch.) The company news would then be fed through a sentiment analysis library like texblob or VADER, or would be run through an ML algorithm that has been trained on thousands of stock market news reports to indicate the general sentiments of news outlets. 

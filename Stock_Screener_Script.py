#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Feb  5 09:40:37 2021

@author: samueljocas

Description: This program reads in an excel spreadsheet (.xlsx) 
with a column of stock tickers, scrapes key pieces of stock information
from yahoofinance and marketwatch, and enters that information into a pandas data frame
and exports it to an excel file in the same folder, making a one-stop shop for financial data 
"""
#import relevant libraries 
import sys
from datetime import datetime
import pandas as pd 
from yahoofinancials import YahooFinancials
import datetime as dt

#list path to read in and output files 
folder_path = "/Users/samueljocas/Desktop/Data Science Portfolio/Stock Screener Files"
file_path = "/Stock_Ticker_Input_File.xlsx"


def importFile(path):#function that imports starting .xslx file 
    try:
        df = pd.read_excel(path, header = 1, index_col = 0) #reads in .xlsx file to a pd df 
        print("Successfully read in file...") 
    except FileNotFoundError:
        sys.exit(1) #exit program if file not found
    return df

def fetchTickers(df): #I am isolating the list of tickers in case I want to do more with this program in the future 
    all_tickers = []
    all_tickers = df.index
    all_tickers = list(all_tickers) #typecast df as a list
    return all_tickers 

def fetchData(tickers,df):
    for ticker in tickers: 
        print("Fetching data for " + str(ticker) + "....")
        ##Fetch all of the relvant stats for each stock 
        yf = YahooFinancials(ticker)
        #enter data into each respective column
        df.loc[ticker]["Price"] = yf.get_current_price()
        df.loc[ticker]["50d Avg."] = yf.get_50day_moving_avg()
        df.loc[ticker]["52W H"] = yf.get_yearly_high()
        df.loc[ticker]["52W L"] = yf.get_yearly_low()
        df.loc[ticker]["Beta"] = yf.get_beta()
        df.loc[ticker]["10Vol"] = yf.get_ten_day_avg_daily_volume()
        df.loc[ticker]["90Vol"] = yf.get_three_month_avg_daily_volume()
        #all MM columns divide value by 1 Million for easy analysis
        df.loc[ticker]["Mkt. Cap (MM)"] = yf.get_market_cap() / 1000000
        df.loc[ticker]["P/E"] = yf.get_pe_ratio()
        df.loc[ticker]["Payout Rat."] = yf.get_payout_ratio()
        df.loc[ticker]["EPS"] = yf.get_earnings_per_share()
        df.loc[ticker]["EBIT (MM)"] = yf.get_ebit() /1000000
        df.loc[ticker]["Gross Profit (MM)"] =  yf.get_gross_profit() / 1000000
        df.loc[ticker]["Book Val. (MM)"] = yf.get_book_value() / 1000000     

#########OUTPUT#######################
print("Starting Program....")
df = importFile(folder_path + file_path)
tickers = fetchTickers(df)
fetchData(tickers,df)
#export to folder with timestamp 
df.to_excel(folder_path + "/Stock_Screener_Output_File-" + str(dt.datetime.now())+ ".xlsx", sheet_name = "Sheet 1", float_format="%.2f")
print("Program finished, output file saved to folder:\n" + folder_path ) #saves in same folder as input file 
import datetime as dt
import pandas as pd
from pandas_datareader import data as pdr
import yfinance as yf
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import os
from pandas import ExcelWriter


def main():

    yf.pdr_override()
    start = dt.datetime(2019, 1, 1)
    now = dt.datetime.now()

    # file path for spreadsheet of stocks
    filePath = r"C:\Users\LAM\Desktop\RichardStocks.xlsx"

    stocklist = pd.read_excel(filePath)
    stocklist = stocklist.head(20)  # get first 50 stocks in RichardStocks.xlsx
    # print(stocklist)

    exportList = pd.DataFrame(columns=['Stock', "RS_Rating", "50 Day MA",
                                       "150 Day Ma", "200 Day MA", "52 Week Low", "52 week High"])

    for i in stocklist.index:
        stock = str(stocklist["Symbol"][i])
        RS_Rating = stocklist["RS Rating"][i]

        try:
            df = pdr.get_data_yahoo(stock, start, now)
            smaUsed = [50, 150, 200]
            for sma in smaUsed:
                df["SMA_"+str(sma)] = round(df.iloc[:,
                                                    4].rolling(window=sma).mean(), 2)

            currentClose = df["Adj Close"][-1]
            moving_average_50 = df["SMA_50"][-1]
            moving_average_150 = df["SMA_150"][-1]
            moving_average_200 = df["SMA_200"][-1]
            low_of_52week = min(df["Adj Close"][-260:])
            high_of_52week = max(df["Adj Close"][-260:])
            try:
                moving_average_200_20 = df["SMA_200"][-20]
            except Exception:
                moving_average_200_20 = 0

            if(checkConditions(df, RS_Rating, stock)):
                exportList = exportList.append({'Stock': stock, "RS_Rating": RS_Rating, "50 Day MA": moving_average_50, "150 Day Ma": moving_average_150,
                                                "200 Day MA": moving_average_200, "52 Week Low": low_of_52week, "52 week High": high_of_52week}, ignore_index=True)

        except Exception:
            print("No data on "+stock)

    print(exportList)

    newFile = os.path.dirname(filePath)+"/ScreenOutput.xlsx"

    writer = ExcelWriter(newFile)
    exportList.to_excel(writer, "Sheet1")
    writer.save()


def checkConditions(df, RS_Rating, stock):

    currentClose = df["Adj Close"][-1]
    moving_average_50 = df["SMA_50"][-1]
    moving_average_150 = df["SMA_150"][-1]
    moving_average_200 = df["SMA_200"][-1]
    low_of_52week = min(df["Adj Close"][-260:])
    high_of_52week = max(df["Adj Close"][-260:])

    try:
        moving_average_200_20 = df["SMA_200"][-20]

    except Exception:
        moving_average_200_20 = 0

    print(stock, currentClose, moving_average_150, moving_average_50,
          moving_average_200, moving_average_200_20)
    # Condition 1: Current Price > 150 SMA and > 200 SMA
    if(currentClose <= moving_average_150 and currentClose <= moving_average_200):
        print("cond_1 is false for stock: " + str(stock))
        return False

    # Condition 2: 150 SMA and > 200 SMA
    if(moving_average_150 <= moving_average_200):
        print("cond_2 is false for stock: " + str(stock))
        return False

    # Condition 3: 200 SMA trending up for at least 1 month (ideally 4-5 months)
    if(moving_average_200 <= moving_average_200_20):
        print("cond_3 is false for stock: " + str(stock))
        return False

    # Condition 4: 50 SMA> 150 SMA and 50 SMA> 200 SMA
    if(moving_average_50 <= moving_average_150 and moving_average_50 < moving_average_200):
        print("cond_4 is false for stock: " + str(stock))
        return False

    # Condition 5: Current Price > 50 SMA
    if(currentClose <= moving_average_50):
        print("cond_5 is false for stock: " + str(stock))
        return False

    # Condition 6: Current Price is at least 30% above 52 week low (Many of the best are up 100-300% before coming out of consolidation)
    if(currentClose < (1.3*low_of_52week)):
        print("cond_6 is false for stock: " + str(stock))
        return False

    # Condition 7: Current Price is within 25% of 52 week high
    if(currentClose < (.75*high_of_52week)):
        print("cond_7 is false for stock: " + str(stock))
        return False

    # Condition 8: IBD RS rating >70 and the higher the better
    if(RS_Rating <= 70):
        print("cond_8 is false for stock: " + str(stock))
        return False

    # all 8 conditions above must be true
    return True


main()

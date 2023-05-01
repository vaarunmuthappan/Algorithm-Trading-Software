from yahoo_fin.stock_info import get_data
import tkinter as tk
import pandas as pd
import os, sys
from datetime import date
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg,
NavigationToolbar2Tk)

## FOR FILE PATHS
try: # running using executable
    path = sys._MEIPASS

except: # running using.py sript
    path = os.path.abspath('.')

# Home - button to see ticker
root = tk.Tk()
root.geometry("1000x750")
root.columnconfigure(3, {'minsize': 50})

deskTopPath = os.getcwd()
if not os.path.exists(deskTopPath):
    os.makedirs(deskTopPath)

# variables for strategy
buyCallPrice = 0
sellCallPrice = 0
tickerVal = ""

def drawGraph(tickerVal):
    graphText.set("Graph based on historical data, buying in blocks of 25 shares:")

    BuyLowFactor = 0.85
    lastBuyPrice = 0.0
    noStocks = 0
    money = 100.0
    targetAll = 250
    aveCSB = 0.0

    filePath = os.path.join(deskTopPath, tickerVal + ".csv")
    df = pd.read_csv(filePath)
    money = 100
    stocks = 0
    lastBPrice = 0
    avePrice = 0
    buyCall = 0
    sellCall = 0
    df["buyCalls"] = 0
    df["sellCalls"] = 0
    df["Money"] = money
    df["lastBPrice"] = 10000
    df["avePrice"] = 10000

    for index, row in df.iloc[1:, ].iterrows():
        buyCall = 0.85 * df.iloc[index - 1]["252ave"]  # 0.85* ave252 days ending yesterday
        df.at[index, "lastBPrice"] = df.at[index - 1, "lastBPrice"]
        df.at[index, "avePrice"] = df.at[index - 1, "avePrice"]
        if buyCall <= row["high"] and buyCall >= row["low"]:
            money -= buyCall * 25  # buy blocks of 25 at call price
            stocks += 25
            # update variables
            lastBPrice = buyCall
            avePrice = ((avePrice * (stocks - 25)) + lastBPrice * 25) / stocks
            df.at[index, "lastBPrice"] = buyCall
            df.at[index, "avePrice"] = avePrice

        avePriceTarget = 1.15 * df.iloc[index - 1]["252ave"]
        lastBuyTarget = 1.1 * df.iloc[index - 1]["lastBPrice"]
        aveSBTarget = 1.5 * df.iloc[index - 1]["avePrice"]
        sellCall = max(aveSBTarget, lastBuyTarget, avePriceTarget)
        if sellCall <= row["high"] and sellCall >= row["low"] and stocks >= 100:
            money += 25 * sellCall
            stocks -= 25

        df.at[index, "buyCalls"] = buyCall
        df.at[index, "sellCalls"] = sellCall
        df.at[index, "Money"] = money

    df = df.tail(1008)
    fig = Figure(figsize=(5, 5), dpi=100)

    # adding the subplot
    plot1 = fig.add_subplot(111)

    # plotting the graph
    plot1.plot(df[["high","low","buyCalls","sellCalls"]])

    # creating the Tkinter canvas
    # containing the Matplotlib figure
    canvas = FigureCanvasTkAgg(fig,
                               master=root)
    canvas.draw()

    # placing the canvas on the Tkinter window
    canvas.get_tk_widget().grid(row=10, column=2)

    # creating the Matplotlib toolbar
    toolbar = NavigationToolbar2Tk(canvas,
                                   root)
    toolbar.update()

    # placing the toolbar on the Tkinter window
    canvas.get_tk_widget().grid(row=10, column=2)

def submitButton():
    tickerVal = tickerEntry.get()
    dateVal = dateEntry.get()
    filePath = os.path.join(deskTopPath, tickerVal + ".csv")

    # if file exists, delete file, download data for new date
    if os.path.exists(filePath):
        os.remove(filePath)

    data = get_data(tickerVal, start_date=dateVal)
    # print("getting data 3")
    data = data.reset_index()
    data["Average"] = (data["high"] + data["low"]) / 2.0
    data["252ave"] = data.Average.rolling(252).mean().fillna(0)
    data["252ave"].fillna(0, inplace=True)
    data.to_csv(filePath, sep=",",index=False)

    #open file and use values
    data1 = pd.read_csv(filePath)

    # Check that it is up to date
    sFormat = str(data1["index"].iloc[-1])
    lastBussDay = pd.datetime.today() - pd.tseries.offsets.BDay(0)

    if sFormat in str(lastBussDay):
        data = get_data(tickerVal)
        # print("getting data 1", type(data1["index"].iloc[-1]), sFormat)
        data = data.reset_index()
        data["Average"] = (data["high"] + data["low"]) / 2.0
        data["252ave"] = data.Average.rolling(252).mean()
        data1 = data
        data.to_csv(filePath, sep=",", index=False)

    # test
    #print(filePath)
    #print(data1.columns)
    data2 = pd.read_csv(filePath)

    # calculate buy price
    if data1["252ave"].iloc[-1] == 0: # ie is newly inserted
        buyCallPrice =  0.85 * data1["252ave"].iloc[-2]
    else:
        buyCallPrice = 0.85 * data1["252ave"].iloc[-1]
    cText.set(tickerVal + ' Buy Call Price: $' + str(round(buyCallPrice,2)))
    buyExpText.set('( 85% of $' + str(round(buyCallPrice/0.85,2)) + ' (1-year average) = $' + str(round(buyCallPrice,2))+ ')')

    # Calculate Sell price
    # from yearly average price
    avePriceTarget = 1.15 * data1["252ave"].iloc[-1]
    avePText.set('( 1.15 * $' + str(round(data1["252ave"].iloc[-1],2)) + ' (1-year average) = $' + str(round(avePriceTarget,2)) + ')')

    # from last buy price
    lastBuyTarget=0
    histFilePath = os.path.join(deskTopPath, "history.xlsx")
    try:
        hist = pd.read_excel(histFilePath, sheet_name=tickerVal)
        bp = hist['buy_price'].iloc[hist['buy_quantity'].to_numpy().nonzero()[0]]
        bq = hist['buy_quantity'].iloc[hist['buy_quantity'].to_numpy().nonzero()[0]]
        # print(bp, bq)
        lastBuyTarget = 1.1 * (bp.iloc[-1]/bq.iloc[-1])
    except OSError:
        lastBuyTarget = 0
    except Exception as err:
        print(err)

    aveSBTarget=0
    try:
        data1 = pd.read_excel(histFilePath, sheet_name=tickerVal)
        totalSold = data1['sell_quantity'].sum()
        totalP = 0
        totalQ = 0
        for index,row in data1.iterrows():
            # print(row,row['buy_quantity'])
            minima = min(totalSold, row['buy_quantity'])
            totalSold -= minima
            netAdded = (row['buy_quantity'] - minima)
            totalQ += netAdded
            if row['buy_quantity'] > 0:
                totalP += (row['buy_price'] / row['buy_quantity']) * netAdded

        if totalQ!=0:
            aveSBTarget = totalP / totalQ
    except OSError:
        aveSBTarget=0
    except Exception as err:
        print(err)

    aveSBTarget = 1.5* aveSBTarget

    # print("averPriceYear", avePriceTarget, type(avePriceTarget))
    # print("last buy",lastBuyTarget, type(lastBuyTarget))
    # print("averSTOCKbuy",aveSBTarget, type(aveSBTarget))

    sellCallPrice = max(aveSBTarget, lastBuyTarget, avePriceTarget)

    lastBText.set('( 1.10 * $' + str(round(lastBuyTarget/1.1,2))+ ' (last buy price) = $'+str(round(lastBuyTarget,2)) + ')')
    aveCostText.set('( 1.50 * $' + str(round(aveSBTarget/1.5,2))+ ' (average stock cost) = $'+str(round(aveSBTarget,2)) + ')')
    sText.set(tickerVal + ' Sell Call Price: $' + str(round(sellCallPrice,2)))

    drawGraph(tickerVal)

# ticker entry
tk.Label(root, text='Enter stock ticker for calls:').grid(row=0, column=1)
tickerEntry = tk.Entry(root)
tickerEntry.grid(row=0, column=3)

# date entry
tk.Label(root, text='Enter start date in format yyyy-mm-dd:').grid(row=1, column=1)
dateEntry = tk.Entry(root)
dateEntry.grid(row=1, column=3)

submitB = tk.Button(root, text='Enter', command=submitButton)
submitB.grid(row=2, column=2)

# print buy price text + logic
cText = tk.StringVar()
cText.set("")
cLabel = tk.Label(root, textvariable=cText,font=("Helvetica",15,"bold")).grid(row=3, column=2)
buyExpText = tk.StringVar()
buyExpText.set("")
buyExpLabel = tk.Label(root, textvariable=buyExpText).grid(row=4, column=2)

# print sell price text + logic
sText = tk.StringVar()
sText.set("")
sLabel = tk.Label(root, textvariable=sText,font=("Helvetica",15,"bold")).grid(row=5, column=2)
avePText = tk.StringVar()
avePText.set("")
avePLabel = tk.Label(root, textvariable=avePText).grid(row=6, column=2)
lastBText = tk.StringVar()
lastBText.set("")
lastBLabel = tk.Label(root, textvariable=lastBText).grid(row=7, column=2)
aveCostText = tk.StringVar()
aveCostText.set("")
aveCostTextLabel = tk.Label(root, textvariable=aveCostText).grid(row=8, column=2)

# POP UP
def open_popup():
    top= tk.Toplevel(root)
    top.geometry("750x250")
    top.title("Enter Data")

    stTickTr = tk.Entry(top)
    buyAmount = tk.Entry(top)
    sellAmount = tk.Entry(top)
    buyQuantity = tk.Entry(top)
    sellQuantity = tk.Entry(top)
    tk.Label(top, text='Enter stock ticker:').grid(row=0, column=0)
    tk.Label(top, text='Enter total buy price:').grid(row=1, column=0)
    tk.Label(top, text='Enter total buy quantity:').grid(row=2, column=0)
    tk.Label(top, text='Enter total sell price:').grid(row=3, column=0)
    tk.Label(top, text='Enter total sell amount:').grid(row=4, column=0)


    def enter_button():
        d = date.today()
        sTick = stTickTr.get()
        bPrice = buyAmount.get()
        bQuant = buyQuantity.get()
        sPrice = sellAmount.get()
        sQuant = sellQuantity.get()

        filePath = os.path.join(deskTopPath, "history.xlsx")

        if not os.path.exists(filePath):
            d = pd.DataFrame()
            d.to_excel(filePath)

        df1 = pd.read_excel(filePath, None)

        if sTick in df1.keys():
            existingHist = pd.read_excel(filePath, sheet_name=sTick)
            d = {'Date': [pd.datetime.today()], 'buy_price': [bPrice], 'buy_quantity': [bQuant], 'sell_price': [sPrice],
             'sell_quantity': [sQuant]}
            df = pd.DataFrame(data=d)
            existingHist = pd.concat([existingHist, df])
        else:
            d = {'Date': [pd.datetime.today()], 'buy_price': [bPrice], 'buy_quantity': [bQuant], 'sell_price': [sPrice],
             'sell_quantity': [sQuant]}
            existingHist = pd.DataFrame(data=d)

        with pd.ExcelWriter(filePath, mode='a', if_sheet_exists='overlay') as writer:
            existingHist.to_excel(writer, sheet_name=sTick,index=False)

    myButton = tk.Button(top, text='Submit', command=enter_button)
    stTickTr.grid(row=0, column=2)
    buyAmount.grid(row=1, column=2)
    buyQuantity.grid(row=2, column=2)
    sellAmount.grid(row=3, column=2)
    sellQuantity.grid(row=4, column=2)
    myButton.grid(row=5, column=3)


newDataButton = tk.Button(root, text= "Enter buy call / sell call completion", command= open_popup)
newDataButton.grid(row=9, column=2)

# Graph label
graphText = tk.StringVar()
graphText.set("")
graphLabel = tk.Label(root, textvariable=graphText).grid(row=10, column=2)

root.title("Stock Call Estimator")
root.mainloop()



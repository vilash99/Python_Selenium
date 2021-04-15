from selenium import webdriver
from selenium.webdriver.chrome.options import Options

import time
import csv
from datetime import datetime

import win32com.client
speaker = win32com.client.Dispatch("SAPI.SpVoice")

amountTrade = 10000
IP_csv = r"F:\Python Scraping Scripts\08_Intraday_Bot_Candles\stocks.csv"


checkedMinute = 0
shortlistedMinute = 0
timeFrame = 15


### --- Loop infinite until user click START button --- ###
def WaitForUserInput(driver_user):
    myCommand = ""
    
    while True:
        try:
            myCommand = driver_user.find_element_by_id("user-input").get_attribute('innerHTML')
        except:
            myCommand = ""    
        
        myCommand = myCommand.strip()
        
        if myCommand == "START" or myCommand == "STOP":
            return myCommand
            break        
        time.sleep(2) 
########End of Function: WaitForUserInput


### --- Show shortlisted data on custome designed web page --- ###
def ShowLiveDataOnPage(driver_user, ShortListData):
    
    ####Get symbols which dont' need to display on final list
    excludeSymbols = ""
    try:
        excludeSymbols = driver_user.find_element_by_id("stockExclude").get_attribute('value')
    except:
        excludeSymbols = ""
        
    script = "arguments[0].insertAdjacentHTML('beforeend', arguments[1])"        
    
    
    ###Show Shortlisted items on webpage
    driver_user.execute_script("document.getElementById('ShortListed').innerHTML = '';")  
    ShortList = driver_user.find_element_by_id("ShortListed")
    
    tableHTML = '<thead class="text-danger"> <th>Status</th> <th>Symbol</th> <th>Name</th> <th>Close Price</th> <th>Shares</th> </thead> <tbody>'
    
    ###ShortlistStocks["Long/Short, Symbol, StockName, ClosePrice, TotalShare]
    for symbol in ShortListData:        
        if excludeSymbols.find(ShortListData[symbol][1]) == -1: #Check if symbol is exculuded
            newData = '<tr>'
            if ShortListData[symbol][0] == "Long":
                newData = newData + '<td><label class="label label-success">Long</label></td>'		
            else:
                newData = newData + '<td><label class="label label-danger">Short</label></td>'        
            
            newData = newData + '<td>' + str(ShortListData[symbol][1]) + '</td>'
            newData = newData + '<td class="script-name">' + str(ShortListData[symbol][2]) + '</td>'
            newData = newData + '<td>' + str(ShortListData[symbol][3]) + '</td>'
            newData = newData + '<td>' + str(ShortListData[symbol][4]) + '</td>'            
            newData = newData + '</tr>'
            
            tableHTML = tableHTML + newData + '</tbody>'    
    
    driver_user.execute_script(script, ShortList, tableHTML)  
        
########End of Function: ShopLiveDataOnPage

        
if __name__ == "__main__":        
    chrome_option = Options()
    chrome_option.add_argument(r"user-data-dir=C:\ChromeDriver\tradingbot")
    chrome_option.add_argument('--ignore-certificate-errors')
    chrome_option.add_argument("--test-type")
    
    driver = webdriver.Chrome(executable_path=r"C:\ChromeDriver\chromedriver.exe", options=chrome_option)  
    
    driver_user = webdriver.Chrome(executable_path=r"C:\ChromeDriver\chromedriver.exe")  
    driver_user.get("file:///F:/Python%20Scraping%20Scripts/08_Intraday_Bot_Candles/TradingPage.html")
    
    
    ##Investing Symbol: NSESymbol, Name, [open, high, low, close], [open, high, low, close]
    CSVStocks = {}
    
    ###For skipping first row of CSV file and htmlData
    flag = False    
    
    with open(IP_csv) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        
        for row in csv_reader:
            if flag == False:
                flag = True
            else:                
                CSVStocks[row[1]] = [row[2], row[0], [0.0, 0.0, 0.0, 0.0], [0.0, 0.0, 0.0, 0.0]]
                
    
    driver.get("https://in.investing.com/portfolio/?portfolioID=MzRiM2I1ZDFkO21hZD81Mg%3D%3D")
    
    ###Wait Till exact given timeframe minute
    while True:
        now = datetime.now()
        current_minute = now.strftime("%M")
        if int(current_minute) % timeFrame == 0:
            break    
    
    print("Trading Bot is Started!")
    
    checkedMinute = -1
    shortlistedMinute = -1
    
    while True:        
        UserCommand = ""        
        UserCommand = WaitForUserInput(driver_user)        
        
        if UserCommand == "STOP":
            print("User pressed Stop button! Exiting script...")
            break
        
        ###Get live data from Investing.com website of protfolio page
        tmpHTML =  driver.find_element_by_id("tbody_overview_22199085").get_attribute("innerHTML").strip()
        StocksList = tmpHTML.split('<td class="flag">')        
        
        
        ###Get current Minute
        now = datetime.now()
        current_minute = int(now.strftime("%M"))
        
        ####Check if it is 1 minute earlier of current 15 minute timeframe
        ####Symbol: Short/Long, NSESymbol, StockName, ClosePrice, totalSharetoBuy
        ShortlistStocks = {}
        
        if ((current_minute+1) % timeFrame == 0) and (shortlistedMinute != current_minute):  
            
            shortlistedMinute = current_minute
        
            ###Show all candles who have bullish/bearish engulfing candlsticks
            for symbols, info in CSVStocks.items():
                
                prevOpen = info[2][0]
                prevHigh = info[2][1]
                prevLow = info[2][2]
                prevClose = info[2][3]
                
                currOpen = info[3][0]    
                currHigh = info[3][1]
                currLow = info[3][2]
                currClose = info[3][3]
                
                ####Check if previous candle is empty
                if prevOpen == 0:
                    break
                else:
                    if (prevOpen > prevClose) and (prevClose < currClose) and (prevClose > currOpen) and (prevLow >= currLow) and (currOpen < currClose):
                        #Add in shortlisted Stocks table
                        ShortlistStocks[symbols] = ["Long", info[0], info[1], currClose, int(amountTrade/currClose)]
                        
                    elif (prevOpen < prevClose) and (prevClose < currOpen) and (prevClose > currClose) and (prevHigh <= currHigh) and (currOpen > currClose):
                        ShortlistStocks[symbols] = ["Short", info[0], info[1], currClose, int(amountTrade/currClose)]            
                
            
            ###Show all Shortlisted stocks in HTML
            ShowLiveDataOnPage(driver_user, ShortlistStocks)
            speaker.Speak("Stocks are shortlisted! Please check.")        
            
        ###End If: Show live shortlisted data 1 minute before end of 15 min candle
        
        
        ##For Skipping first record
        flag = False
        
        ###Extract all stocks and their current prices
        for currentRow in StocksList:
                
            ##Skip first record
            if flag == False:
                flag = True
                continue
            
            stockName = stockSymbol = stockPrice = ""
            
            ###Extract Stock Name
            stockName = currentRow.split('data-column-name="name"')[1]
            stockName = stockName.split(' title="')[1]
            stockName = stockName.split('"')[0]
            stockName = stockName.replace("&amp;", "&")
            
            ###Extract Stock Symbol
            stockSymbol = currentRow.split('data-column-name="symbol"')[1]
            stockSymbol = stockSymbol.split(' title="')[1]
            stockSymbol = stockSymbol.split('"')[0]        
            
            ###Extract Stock current Price   
            stockPrice = currentRow.split('data-column-name="last"')[1]
            stockPrice = stockPrice.split('">')[1]
            stockPrice = stockPrice.split('</td>')[0]   
            stockPrice = stockPrice.replace(",", "")
            stockPrice = float(stockPrice)
            
            
            ####Check if current time is new start time. Then
            ####Switch current candle to previous candle             
            
            if current_minute % timeFrame == 0 and checkedMinute != current_minute:
                ###Update current candle close price
                CSVStocks[stockSymbol][3][3] = stockPrice
                
                #Current candle is previous candle
                CSVStocks[stockSymbol][2] = CSVStocks[stockSymbol][3]
                
                ####Set all prices to current price of current candles
                CSVStocks[stockSymbol][3] = [0.0, stockPrice, stockPrice, stockPrice]                
            else:                    
                ###Add Open price
                if CSVStocks[stockSymbol][3][0] == 0:
                    CSVStocks[stockSymbol][3][0] = stockPrice
                
                ###Check if current price is higher than previous                
                if CSVStocks[stockSymbol][3][1] < stockPrice:
                    CSVStocks[stockSymbol][3][1] = stockPrice
                elif CSVStocks[stockSymbol][3][2] > stockPrice:
                    CSVStocks[stockSymbol][3][2] = stockPrice
                
                ##Change close price
                CSVStocks[stockSymbol][3][3] = stockPrice                      
        
        
        ###End For: Scraping data from webpage
        checkedMinute = current_minute
        
        time.sleep(5)
    ####<<<<<End While: Read live data from webpage infinte times
                    
                
driver.close()
driver_user.close()



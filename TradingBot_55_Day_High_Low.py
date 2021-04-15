from selenium import webdriver
from selenium.webdriver.chrome.options import Options


import time
import csv

import win32com.client
speaker = win32com.client.Dispatch("SAPI.SpVoice")


IP_csv = r"F:\Python Scraping Scripts\07_Trading_Bot\HighLow.csv"
amountTrade = 10000

### --- Loop infinite until user click START button --- ###
def WaitForUserToBegin(driver_user):
    myCommand = ""
    
    while True:
        try:
            myCommand = driver_user.find_element_by_id("user-input").get_attribute('innerHTML')
        except:
            myCommand = ""    
        
        if myCommand == "START" or myCommand == "STOP":
            return myCommand
            break        
        time.sleep(2) 
        
        
def ShowLiveDataOnPage(driver_user, ShortListData):
    
    excludeSymbols = ""
    try:
        excludeSymbols = driver_user.find_element_by_id("stockExclude").get_attribute('value')
    except:
        excludeSymbols = ""
        
    script = "arguments[0].insertAdjacentHTML('beforeend', arguments[1])"        
    
    ###Show Shortlisted items
    driver_user.execute_script("document.getElementById('ShortListed').innerHTML = '';")  
    ShortList = driver_user.find_element_by_id("ShortListed")
    
    tableHTML = '<thead class="text-danger"> <th>Status</th> <th>Symbol</th> <th>Name</th> <th>Price</th> <th>High</th> <th>Low</th> <th>Shares</th> </thead> <tbody>'
    
    ###ShortlistStocks["Long/Short, Symbol, StockName, Price, High, Low, TotalShare, Target, Stoploss]
    for symbol in ShortListData:        
        if excludeSymbols.find(ShortListData[symbol][1]) == -1:
            newData = '<tr>'
            if ShortListData[symbol][0] == "Long":
                newData = newData + '<td><label class="label label-success">Long</label></td>'		
            else:
                newData = newData + '<td><label class="label label-danger">Short</label></td>'        
            
            newData = newData + '<td>' + str(ShortListData[symbol][1]) + '</td>'
            newData = newData + '<td class="script-name">' + str(ShortListData[symbol][2]) + '</td>'
            newData = newData + '<td>' + str(ShortListData[symbol][3]) + '</td>'
            newData = newData + '<td>' + str(ShortListData[symbol][4]) + '</td>'
            newData = newData + '<td>' + str(ShortListData[symbol][5]) + '</td>'
            newData = newData + '<td>' + str(ShortListData[symbol][6]) + '</td>'
            newData = newData + '</tr>'
            
            tableHTML = tableHTML + newData + '</tbody>'
    
    
    driver_user.execute_script(script, ShortList, tableHTML)  
        


if __name__ == "__main__":
    chrome_option = Options()
    chrome_option.add_argument(r"user-data-dir=C:\ChromeDriver\tradingbot")
    chrome_option.add_argument('--ignore-certificate-errors')
    chrome_option.add_argument("--test-type")
    
    driver = webdriver.Chrome(executable_path=r"C:\ChromeDriver\chromedriver.exe", options=chrome_option)  
    
    driver_user = webdriver.Chrome(executable_path=r"C:\ChromeDriver\chromedriver.exe")  
    driver_user.get("file:///F:/Python%20Scraping%20Scripts/07_Trading_Bot/TradingPage.html")                     
    
    
    ###Read Previous Day high/low prices    
    
    ##Investing Symbol: NSESymbol, Name, High, Low, currPrice, 
    CSVStocks = {}
    ShortlistStocks = {}
    
    flag = False    
    with open(IP_csv) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        
        for row in csv_reader:
            if flag == False:
                flag = True
            else:                
                CSVStocks[row[1]] = [row[2], row[0], row[3].replace(",", ""), row[4].replace(",", ""), "0.00"]
    
    
    ###Show CSV data in my custome webpage
    ShowLiveDataOnPage(driver_user, ShortlistStocks)    
        
    driver.get("https://in.investing.com/portfolio/?portfolioID=MzRiM2I1ZDFkO21hZD81Mg%3D%3D")                            
    
    while True:        
        UserCommand = ""        
        UserCommand = WaitForUserToBegin(driver_user)        
        
        if UserCommand == "STOP":
            print("User pressed Stop button! Exiting script...")
            break        
        
    
        tmpHTML =  driver.find_element_by_id("tbody_overview_22199085").get_attribute("innerHTML").strip()   
        StocksList = tmpHTML.split('<td class="flag">')    
        
        flag = False
        
        ###Use for refresh page, if there is need to refresh
        refreshFlag = False
        
        for currentRow in StocksList:
            ##Scip first record
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
            

            
            
            ###ShortlistStocks["Long/Short, Symbol, StockName, Price, High, Low, TotalShare]
            ###CSVStock: [NSESymbol, Name, High, Low, currPrice]
            
            ##Check if Symbolname is in CSV Dictonary File
            if stockSymbol in CSVStocks:   
                #Update current Price
                CSVStocks[stockSymbol][4] = stockPrice
                
                #Check if current price is high than previous day high
                if stockPrice > float(CSVStocks[stockSymbol][2]):                    
                    speaker.Speak(CSVStocks[stockSymbol][1] + " is ready for long!")
                    refreshFlag = True
                    
                    #Add in shortlisted Stocks table
                    ShortlistStocks[stockSymbol] = ["Long", CSVStocks[stockSymbol][0], stockName, stockPrice, CSVStocks[stockSymbol][2], "-", int(amountTrade/stockPrice)]
                                        
                    #Remove from Dictionary
                    try:
                        del CSVStocks[stockSymbol]
                    except:
                        pass
                    
                elif stockPrice < float(CSVStocks[stockSymbol][3]):                                
                    speaker.Speak(CSVStocks[stockSymbol][1] + " is ready for short!")
                    refreshFlag = True
                    
                    #Add in shortlisted Stocks table
                    ShortlistStocks[stockSymbol] = ["Short", CSVStocks[stockSymbol][0], stockName, stockPrice, "-", CSVStocks[stockSymbol][3], int(amountTrade/stockPrice)]
                    
                    #Remove from Dictionary
                    try:
                        del CSVStocks[stockSymbol]
                    except:
                        pass
                    
            ####<<<End IF: Add/check item in dictionary
        ####<<<<End For: Extract live Prices from website
        
        ####Show live data in custome web page  
        if refreshFlag == True:
            ShowLiveDataOnPage(driver_user, ShortlistStocks)

        
        ###Wait for 5 seconds for next loop
        time.sleep(5)
    ####<<<<<End While: Read live data from webpage infinte times
            
    
    #driver.close()
    #driver_user.close()
    
    
    



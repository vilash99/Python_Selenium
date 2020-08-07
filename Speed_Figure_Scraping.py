from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.select import Select
import csv

from time import sleep
from datetime import date

OP_csv = r'C:\ChromeDriver\speed figures ' + str(date.today()) + ".csv"
monthsName = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

def getMonthIndex(mName):    
    for index in range(0, 12):        
        if monthsName[index] == mName:
            return index
    return -1



def SaveDataInCSV(mRow):    
    with open(OP_csv, 'a', newline='') as appendFile:
        writer = csv.writer(appendFile)
        writer.writerow(mRow)
    
    appendFile.close() 


def scrapHeader(headerData):
    #Scrap Header Row
    tableRow = []
    for tmpRow in headerData:
        if tmpRow.find("</th>") > 0:
            tableRow.append(tmpRow.split('</th>')[0].strip())
    
    #Save Scraped Table Data in CSV
    SaveDataInCSV(tableRow)

    
def scrapDataRow(RowData):
    #Save Actual Data
    tableRow = []
    for tmpRow in RowData:
        if tmpRow.find('<td class="text-center">') >= 0:
            tmpCell = tmpRow.split('<td class="text-center">')
            
            #Extract Cells of current Row
            i = 1
            for cell in tmpCell:
                if cell.find("</td>") > 0:
                    if i == 1 or i == 3:
                        tableRow.append("'" + cell.split('</td>')[0].strip())
                    else:
                        tableRow.append(cell.split('</td>')[0].strip())                    
                    i = i + 1                
            
            SaveDataInCSV(tableRow)
            tableRow = []
        
            
if __name__ == "__main__":
    
    fromDay = input("Enter start date (dd mmm yyyy): ")
    endDay = input("Enter end date (dd mmm yyyy): ")
    headerSaved = False
    
    #Split the day, month and year      
    fromDate, fromMonth, fromYear = fromDay.split(" ")
    endDate, endMonth, endYear = endDay.split(" ")
    
    startMonthIndex = getMonthIndex(fromMonth)
    endMonthIndex = getMonthIndex(endMonth)
    
    if startMonthIndex == -1 or endMonthIndex == -1:
        print("Plese enter valid month shown in calender! Unable to begin script.")
        exit
    
    
    #Create Chrome Instance
    chrome_option = Options()
    chrome_option.add_argument("--user-data-dir=C:\\ChromeDriver\\userdata")    
    #chrome_option.add_argument("--headless")  
    driver = webdriver.Chrome(options=chrome_option, executable_path=r"C:\ChromeDriver\chromedriver.exe")
    driver.minimize_window()
    
    #fromDay = "20 Jun 2020"
    #toDay = "1 Aug 2020"
    
    #currentMonth = "Jun 2020"
    #currentDay = "20"    
    
    #Loop through all month from beggining date
    while startMonthIndex <= endMonthIndex:
        
        #Get format shown in calender e.g. Jan 2020
        currentMonth = monthsName[startMonthIndex] + " " + fromYear
        
        #Loop throuh all days in month
        while int(fromDate) <= 31:
            
            #scrap till given date and month
            if int(fromDate) > int(endDate) and startMonthIndex == endMonthIndex:
                break            
            
            print("Scraping for date:", fromDate + " " + currentMonth)
            
            #Open given URL
            driver.get('https://andyholdingspeedfigures.co.uk/race-search/') 
            driver.implicitly_wait(20)  
            sleep(5) #Wait Extra 5 seconds
            
            
            #########Selection Start from HERE###################
            #Select DatePicker
            raceDatePicker = driver.find_element_by_xpath("//input[@id='datepickerid']").click()
            sleep(5)    
            
            #Select Month
            while True:
                raceMonth = driver.find_element_by_class_name("day__month_btn")
                raceMonth.get_attribute("innerHTML")
                
                if raceMonth.text == currentMonth:
                    break
                else:
                    #Click on previou month arrow
                    driver.find_element_by_xpath("/html/body/div[1]/div/div/div/div/form/div/div/div[2]/div[1]/div/div/div[2]/header/span[1]").click()
                    sleep(3)
            #-------------End of Month selector while loop-------------#
            
            #Choose Date
            try:
                driver.find_element_by_xpath("/html/body/div[1]/div/div/div/div/form/div/div/div[2]/div[1]/div/div/div[2]/div/span[contains(text(), "+ fromDate + ")]").click()
                sleep(5)
            except:
                break            
            
            
            #Select Course if any availabe
            courseList = Select(driver.find_element_by_css_selector("select.form-control"))
            
            #Loop through all options and scrap the data            
            for course in courseList.options:
                courseList.select_by_visible_text(course.text)
                sleep(2)
                
                driver.find_element_by_class_name("btn-primary").click() #Click on Search button
                sleep(5)
                
                #Get HTML of Data
                dataTable = driver.find_element_by_css_selector("table.table")
                dataTableHTML = dataTable.get_attribute("innerHTML")
                
                #Get Header Row
                if headerSaved == False:
                    scrapHeader(dataTableHTML.split('<th scope="col" class="text-center">'))
                    headerSaved = True
                
                #Get 1st page table
                scrapDataRow(dataTableHTML.split('<tr class="small">'))        
                
                
                ##Loop Through all pages to get all data
                while True:
                    try:
                        driver.find_element_by_css_selector("ul.pagination li.next a").click()
                        sleep(5)
                        
                        dataTable = driver.find_element_by_css_selector("table.table")
                        dataTableHTML = dataTable.get_attribute("innerHTML")
                        
                        #Scrap Actuall Data
                        scrapDataRow(dataTableHTML.split('<tr class="small">'))
                    
                    except:
                        break
                #--------------End of while: scanning all pages------------#
            #--------------End of For: choose all options in course------------#
            
            
            #Incrment to get next date of month
            fromDate = str(int(fromDate) + 1)
        #--------------End while: selecting Date----------------------#
        
        
        #Increment month after each completion of month
        startMonthIndex = startMonthIndex + 1
        fromDate = "1"
    
    driver.quit()
    
    print("All data is scraped successfully!")
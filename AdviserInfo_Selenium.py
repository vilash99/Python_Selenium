from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import time
import csv

OP_csv = r'D:\dcfile6413.csv'

def SaveDataInCSV(mRow):    
    with open(OP_csv, 'a', newline='') as appendFile:
        writer = csv.writer(appendFile)
        writer.writerow(mRow)
    
    appendFile.close()
    

def searchZipCode(crdNo, zipCode):
    
    #Search CRD in respected textbox
    cNo = driver.find_element_by_id('firm-input')
    cNo.send_keys(crdNo)
    
    #Search ZIP Code
    zipC = driver.find_element_by_id('acIndividualLocationId')
    zipC.send_keys(zipCode)
    
    #Press Enter to getch results
    time.sleep(3)
    zipC.send_keys(Keys.ENTER)
    time.sleep(3)
    
    #Extract all CRD Number    
    crdNumberList = driver.find_elements_by_css_selector(".names > div.font-dark-gray > div.smaller > span.ng-binding")    
    myCRD = []
    for cNo in crdNumberList:
        myCRD.append(cNo.text)
    #Required form of CRD url
    #https://adviserinfo.sec.gov/individual/summary/1575389
    
    driver.execute_script("window.open('','_blank');")   
    time.sleep(3)
    driver.switch_to.window(driver.window_handles[1])
    time.sleep(2)
    
    for cNo in myCRD:        
        cURL = "https://adviserinfo.sec.gov/individual/summary/" + cNo
        print(cURL)
        
        driver.get(cURL)  
        driver.implicitly_wait(5)
        time.sleep(5)
                
        try:
            advisorName = driver.find_element_by_class_name("namesummary").text
        except:
            advisorName = ""
        
        try:
            advisorMoreName = driver.find_element_by_css_selector("#bio-geo-summary > div.names.layout-xs-column.layout-column.flex-auto > div:nth-child(2) > div").text
        except:
            advisorMoreName = ""
            
        try:
            advisorCompany = driver.find_element_by_css_selector("#bio-geo-summary > div.layout-gt-xs-row.layout-column > div.employment.md-body-1.offset-margintop-1x.layout-column.flex-gt-xs-30.flex > div.ng-scope > div > div.bold.crdtextcolor.ng-binding").text        
        except:
            advisorCompany = ""
        
        try:
            companyCRD = driver.find_element_by_css_selector("#bio-geo-summary > div.layout-gt-xs-row.layout-column > div.employment.md-body-1.offset-margintop-1x.layout-column.flex-gt-xs-30.flex > div.ng-scope > div > div.smaller > span.crdtextcolor.ng-binding").text
        except:
            companyCRD = ""        
        
        try:
            companyAddress = driver.find_element_by_css_selector("#bio-geo-summary > div.layout-gt-xs-row.layout-column > div.employment.md-body-1.offset-margintop-1x.layout-column.flex-gt-xs-30.flex > div.ng-scope > div > div.smaller > div").text
            companyAddress = companyAddress.replace("\n", ", ")
        except:
            companyAddress = ""
        
        #Save Data in CSV
        SaveDataInCSV([crdNo, zipCode, cNo, advisorName, advisorMoreName, advisorCompany, companyCRD, companyAddress])
    
    driver.close()
    time.sleep(1)
    driver.switch_to.window(driver.window_handles[0])


if __name__ == "__main__":
    chrome_option = Options()
    chrome_option.add_argument("--user-data-dir=C:\\ChromeDriver\\userdata")    
    #chrome_option.add_argument("--headless")    
    driver = webdriver.Chrome(options=chrome_option, executable_path=r"C:\ChromeDriver\chromedriver.exe")    
    
    driver.get("https://adviserinfo.sec.gov/")    
    searchZipCode("6413", "10001")
        
    #driver.implicitly_wait(1)
    
    driver.quit()
    
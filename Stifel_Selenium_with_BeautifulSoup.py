from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import csv
from random import randint
from time import sleep


OP_csv = r'D:\dcfile793_Selenium.csv'

def SaveDataInCSV(mRow):    
    with open(OP_csv, 'a', newline='') as appendFile:
        writer = csv.writer(appendFile)
        writer.writerow(mRow)
    
    appendFile.close() 
    

def GetBranchData(tmpBranchURL):
    driver.get(tmpBranchURL)
    
    bsObj = BeautifulSoup(driver.page_source, "lxml")
    
    #Get Branch Name
    try:
        branchName = bsObj.find("h1", class_="bold-headline").get_text()
    except:
        branchName = ""
    
    #Get branch manager
    try:
        branchManager = bsObj.select(".branch-landing-info-border div span a")[0].get_text()
    except:
        branchManager = ""        
    
    #Get contact numbers
    try:
        branchContact = [contactNo.text for contactNo in bsObj.select(".branch-landing-phone-desktop a")]
    except:
        branchContact=[]       
        
    #Get Address
    try:
        branchAddress = bsObj.find("div", class_="branch-landing-address").get_text()
        branchAddress = ",".join(branchAddress.split("\n"))
        branchAddress = branchAddress.replace(',Get Directions,', '')
        branchAddress = branchAddress[1:]        
    except:
        branchAddress = ""
        
    try:
        advisorURL = bsObj.select("div.branch-landing-financial-advisors-columns div.branch-landing-financial-advisors-branchFA a")
    except:
        advisorURL = []
        
    #Open each advisorURL and extract its details
    for aURL in advisorURL:
        print("\tOpening advisor page...")
        
        tmpAdvisorURL = "https://www.stifel.com"+aURL['href']
        
        driver.get(tmpAdvisorURL)
        bsObj = BeautifulSoup(driver.page_source, "lxml")
        
        #Get Advisor name
        try:
            advisorName = bsObj.find("span", class_="fa-landing-name").get_text()
        except:
            advisorName = ""
                        
        #Advisor Title
        try:
            advisorTitle = bsObj.select(".fa-landing-name-wrapper div p")[0]
            
            #Remove advisor name tag
            toRemove = advisorTitle.find("span")
            _ = toRemove.extract()
            advisorTitle = advisorTitle.get_text(strip=True)
        except:
            advisorTitle = ""
        
        SaveDataInCSV([tmpBranchURL, branchName, branchManager, branchContact, branchAddress, tmpAdvisorURL, advisorName, advisorTitle])        
    
    if len(advisorURL) == 0:
        SaveDataInCSV([tmpBranchURL, branchName, branchManager, branchContact, branchAddress, "", "", ""])
        

if __name__ == "__main__":
    chrome_option = Options()
    chrome_option.add_argument("--user-data-dir=C:\\ChromeDriver\\userdata")    
    driver = webdriver.Chrome(options=chrome_option, executable_path=r"C:\ChromeDriver\chromedriver.exe")
    
    driver.get('https://www.stifel.com/branch/')
    
    bsObj = BeautifulSoup(driver.page_source, "lxml")
    branch_urls = bsObj.select(".branch-landing-locations-list li a")
    
    recordCount = 0
    for bURL in branch_urls:
        if recordCount >= 2:
            break
        
        recordCount = recordCount + 1
        
        print("Running {0} URL...".format(recordCount))
        
        tmpBranchURL = "https://www.stifel.com"+bURL['href']
        GetBranchData(tmpBranchURL)
        
        waitTime = randint(1, 5)
        print("Waiting for:", waitTime)
        sleep(waitTime)
    
    driver.quit()
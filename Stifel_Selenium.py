from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import csv
from random import randint
from time import sleep


OP_csv = r'D:\dcfile793_Selenium_Pure.csv'

def SaveDataInCSV(mRow):    
    with open(OP_csv, 'a', newline='') as appendFile:
        writer = csv.writer(appendFile)
        writer.writerow(mRow)
    
    appendFile.close() 
    

def GetBranchData(tmpBranchURL):
    driver.get(tmpBranchURL)      
    
    #Get Branch Name
    try:
        branchName = driver.find_element_by_class_name("bold-headline").text
    except:
        branchName = ""
    
    #Get branch manager
    try:
        branchManager = driver.find_element_by_css_selector(".branch-landing-info-border > div > span > a").text
        
    except:
        branchManager = ""        
    
    #Get contact numbers
    try:
        tmpContact = driver.find_elements_by_css_selector(".branch-landing-phone-desktop > a")        
        branchContact = [tmpCnt.get_attribute('innerHTML') for tmpCnt in tmpContact]
    except:
        branchContact = []       
        
    #Get Address
    try:
        branchAddress = driver.find_element_by_class_name("branch-landing-address").text
        branchAddress = ",".join(branchAddress.split("\n"))
        branchAddress = branchAddress.replace(',Get Directions,', '')
        branchAddress = branchAddress[1:]  
        ##We can seprate address, city, state zip with , with no space after it
    except:
        branchAddress = ""
      
    #Get all Advisor URL, if any present
    try:
        tmpURL = driver.find_elements_by_css_selector(".branch-landing-financial-advisors-columns > div.branch-landing-financial-advisors-branchFA [href]")
        advisorURL = [tmpU.get_attribute('href') for tmpU in tmpURL]
    except:
        advisorURL = []
        
        
    #Open each advisorURL and extract its details
    for tmpAdvisorURL in advisorURL:
        print("\tOpening advisor page...")
        
        driver.get(tmpAdvisorURL)
        
        #Get Advisor name
        try:
            advisorName = driver.find_element_by_class_name("fa-landing-name").text            
        except:
            advisorName = ""
                        
        #Advisor Title
        try:
            advisorTitle = driver.find_element_by_css_selector(".fa-landing-name-wrapper div p").text            
                        
            #Remove advisor name tag            
            advisorTitle = advisorTitle.split("\n")[1]
        except:
            advisorTitle = ""
        
        SaveDataInCSV([tmpBranchURL, branchName, branchManager, branchContact, branchAddress, tmpAdvisorURL, advisorName, advisorTitle])        
        
        #Wait for some time
        waitTime = randint(1, 3)
        print("Waiting for:", waitTime, "seconds...")
        sleep(waitTime)        
    
    if len(advisorURL) == 0:
        SaveDataInCSV([tmpBranchURL, branchName, branchManager, branchContact, branchAddress, "", "", ""])

if __name__ == "__main__":
    chrome_option = Options()
    chrome_option.add_argument("--user-data-dir=C:\\ChromeDriver\\userdata")    
    driver = webdriver.Chrome(options=chrome_option, executable_path=r"C:\ChromeDriver\chromedriver.exe")
    
    driver.get('https://www.stifel.com/branch/')    
    
    branch_urls = driver.find_elements_by_css_selector(".branch-landing-locations-list [href]")
    links = [tmpUrl.get_attribute('href') for tmpUrl in branch_urls]
    
    recordCount = 0
    for tmpBranchURL in links:
        if recordCount >= 2:
            break
        
        recordCount = recordCount + 1
        print("Running {0} URL...".format(recordCount))
        
        GetBranchData(tmpBranchURL)
        
        #Wait time after each branch details
        waitTime = randint(2, 5)
        print("Waiting for:", waitTime, "seconds...")
        sleep(waitTime)        
    
    driver.quit()
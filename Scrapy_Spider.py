from scrapy.http import Request
from scrapy.spiders import CSVFeedSpider
import re

class ScraperSpider(CSVFeedSpider):
    name = 'scraper'
    start_urls = ['file:///E:/Scrapy%20Practice/Oreleans_Scraping/oreleans/orelansurl.csv']
    delimiter = ','
    headers = ['URL']    

    
    def parse_row(self, response, row):         
        yield Request(url = row['URL'], callback = self.scrapData)           
     

    def scrapData(self, response):
        print("Working URL - " + response.url)     
        
        def deleteHTMLTags(raw_html):
            cleanr = re.compile('<.*?>')
            cleantext = re.sub(cleanr, '', raw_html)
            return cleantext
            
            
        def extract_with_css(query, index):            
            extractData = ""
            try:
                extractData = response.css(query).extract()[index]
                extractData = deleteHTMLTags(extractData)
            except:
                extractData = ""
            
            return extractData
        
        #Extrct building Image        
        try:
            data = response.css("td img").extract()
            photo = data[len(data)-1] #Get Last Image   
        except:
            photo = ""        
        
        #Extract Tax Bill URL
        TaxBillURL = ""
        try:
            TaxBillURL = response.css('a[href*="http://services.nola.gov/service.aspx?load=treasury"]::attr(href)').get()
        except:
            TaxBillURL = ""
        
        yield {
            'URL': response.url,            
            'Ownername': extract_with_css("td.owner_value", 0),   
            'Todaydate': extract_with_css("td.owner_value", 1),
            'Mailingaddress': extract_with_css("td.owner_value", 2),
            'Municipaldistrict': extract_with_css("td.owner_value", 3),
            'Locationaddress': extract_with_css("td.owner_value", 4),
            'TaxBillURL': TaxBillURL,
            'Propertyclass': extract_with_css("td.owner_value", 6),
            'Subdivisionname': extract_with_css("td.owner_value", 8),
            'Landarea': extract_with_css("td.owner_value", 9),
            'Buildingarea': extract_with_css("td.owner_value", 11),
            'Parcelmap': extract_with_css("td.owner_value", 17),
            'Assessmentarea': extract_with_css("td.owner_value", 19),
            'Viyear': extract_with_css("td.tax_value", 0),
            'Vilandvalue': extract_with_css("td.tax_value", 1),
            'Vibuildingvalue': extract_with_css("td.tax_value", 2),
            'Vitotalvalue': extract_with_css("td.tax_value", 3),
            'Viassessedlandvalue': extract_with_css("td.tax_value", 4),
            'Viassessedbuildingvalue': extract_with_css("td.tax_value", 5),
            'Vitotalassessedvalue': extract_with_css("td.tax_value", 6),
            'Vihomesteadexemptionvalue': extract_with_css("td.tax_value", 7),
            'Vitaxableassessment': extract_with_css("td.tax_value", 8),
            'Stdate1': extract_with_css("tr.even td.sales_value", 0),
            'Stprice1': extract_with_css("tr.even td.sales_value", 1),
            'Stgrantor1': extract_with_css("tr.even td.sales_value", 2),
            'Stgrantee1': extract_with_css("tr.even td.sales_value", 3),
            'Stdate2': extract_with_css("tr.odd td.sales_value", 0),
            'Stprice2': extract_with_css("tr.odd td.sales_value", 1),
            'Stgrantor2': extract_with_css("tr.odd td.sales_value", 2),
            'Stgrantee2': extract_with_css("tr.odd td.sales_value", 3),
            'Stdate3': extract_with_css("tr.even td.sales_value", 6),
            'Stprice3': extract_with_css("tr.even td.sales_value", 7),
            'Stgrantor3': extract_with_css("tr.even td.sales_value", 8),
            'Stgrantee3': extract_with_css("tr.even td.sales_value", 9),            
            'Buildingphoto': photo,
        }
        
        

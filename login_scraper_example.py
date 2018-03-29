#CTRL+ALT+N to Run the python script
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import requests
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from openpyxl import Workbook
from openpyxl import load_workbook
import re

#x = urllib.request.urlopen('https://pythonprogramming.net')

try:

    #Stop remember me from showing up
    chrome_options = Options()
    chrome_options.add_experimental_option('prefs', {
        'credentials_enable_service': False,
        'profile': {
            'password_manager_enabled': False
        }
    })
    #Open the chrome driver, and navigate to the page
    driver = webdriver.Chrome("newchromedriver/chromedriver.exe",chrome_options=chrome_options)
    

    #Load original Excel file - and select the needed sheet - get maximum row number
    Result_File = load_workbook('source_file/IPOs.xlsx')
    Result_sheet = Result_File.worksheets[0]
    Max_Row_Count = Result_sheet.max_row

    #Count not found records  on site
    Not_Found_Tickers = 0
    
    #Filing type
    Filing_Type = "424B4"

    #Program Name
    Program_Name = "directed share program"

    for i in range(1650, Max_Row_Count):

        driver.get("https://www.sec.gov/edgar/searchedgar/companysearch.html")

        #Get the ticker and EPS date from data sheet
        print(i)  

        #Save a final results sheet
        if(i%10 == 0):
            Result_File.save('results_file/FINAL.xlsx')  

        Ticker_Symbol = Result_sheet.cell(row=i, column=3).value
        


        #Locate the search box on top of the page
        Ticker_Symbol_Search = driver.find_element_by_name("CIK")

        #send ticker symbol value to the search box
        Ticker_Symbol_Search.send_keys(str(Ticker_Symbol))
        
        #Press enter on the search box
        Ticker_Symbol_Search.send_keys(Keys.ENTER)

        try:
            time.sleep(2)  
            #Find the Filing type type search box
            Filing_Type_Search = driver.find_element_by_name('type')

            #send filing type to search box
            Filing_Type_Search.send_keys(Filing_Type)

            #press enter on the search box
            Filing_Type_Search.send_keys(Keys.ENTER)

            time.sleep(2)           

            #Grab the documents link and click it
            Documents_Link = driver.find_element_by_xpath('//*[@id="documentsbutton"]').click()

            time.sleep(2)
            #424B4 document link
            Final_Document_Link = driver.find_element_by_xpath('//*[@id="formDiv"]/div/table/tbody/tr[2]/td[3]/a').click() 


            time.sleep(2)
            #get text from the page body
            Body = driver.find_element_by_xpath("/html/body")
            Body_To_Text = Body.text
            Body_To_Text = Body_To_Text.lower()
            

            Program_Text = re.findall(r'([^.]*'+Program_Name+'[^.]*)', Body_To_Text)

            if not Program_Text:
                
                Result_sheet.cell(row=i, column=13).value = 0
                Result_sheet.cell(row=i, column=14).value = 'No Shares Program'
                Result_sheet.cell(row=i, column=15).value = 'No Shares Program'

            else:
                Result_sheet.cell(row=i, column=13).value = 1
                Result_sheet.cell(row=i, column=15).value = '.'.join(Program_Text)

        except Exception as e:
            Result_sheet.cell(row=i, column=13).value = 'No SEC Record'
            Result_sheet.cell(row=i, column=14).value = 'No SEC Record'
            Result_sheet.cell(row=i, column=15).value = 'No SEC Record'
            Not_Found_Tickers += 1
            print('not found '+ str(Not_Found_Tickers))
            
    
    #Save a final results sheet
    Result_File.save('results_file/FINAL.xlsx')   

    #close the chrome page
    driver.close()

except Exception as e:
    print(str(e))
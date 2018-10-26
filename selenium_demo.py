#import selenium libs and file handler 
#'xxx' in code represents personal information and is thus hidden
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
import pandas as pd
from pandas import ExcelWriter
import random

driver = webdriver.Chrome(executable_path="xxx/chromedriver.exe")


#login
driver.get("xxx")
driver.find_element_by_xpath('//*[@id="userId"]').send_keys('xxx')  
driver.find_element_by_xpath('//*[@id="formLogin"]/table/tbody/tr[3]/td/input[1]').send_keys('xxx') 
driver.find_element_by_xpath('//*[@id="formLogin"]/table/tbody/tr[4]/td/input').click()

#page process
driver.find_element_by_xpath('//*[@id="headKeyword"]').send_keys('xxx')  
driver.find_element_by_xpath('/html/body/div[4]/div/div[1]/input[4]').click()
driver.find_element_by_xpath('/html/body/div[4]/div/div[1]/input[4]').click()   



#open file
df_excel = pd.read_excel("xxx.xlsx")
writer = ExcelWriter("xxx.xlsx") 
index = 1

#get_text function to get all text within one web_element
def get_text(tag):
    children = tag.find_elements_by_xpath('*')
    original_text = tag.text
    for child in children:
        original_text = original_text + child.text
    return original_text

#loop with page	
for page in range(xxx):    
	#loop with table rows in a page
    for ele in driver.find_elements_by_xpath('//*[@id="dlAjaxRecordList"]/dl/dd'):
	
        #do exception handling with every element
        try:
            half_website = ele.find_element_by_xpath('.//a').get_attribute('href')
            website  = 'xxx'+half_website.replace('xxx','') #save redirecting website       
            df_excel.loc[index,'website'] = website
            print('WW'+website[0:4])
        except:
            half_website = '获取失败' #locating element failure
            df_excel.loc[index,'website'] = half_website
            print('WW'+half_website[0:4])
                        
        #anotehr element
        try:
            title = ele.find_element_by_xpath('.//a').text
            df_excel.loc[index,'title'] = title
            print('AA'+title[0:4]) 
        except:
            title = '获取失败'
            df_excel.loc[index,'title'] = title
            print('AA'+title[0:4]) 
                    
        #another element
        try:
            number = ele.find_element_by_xpath('.//div[2]/p/span[2]').text
            df_excel.loc[index,'number'] = number
            print('BB'+number[0:4])
        except:
            number = '获取失败'
            df_excel.loc[index,'number'] = number
            print('BB'+number[0:4])
        
        try: 
			
            #switch to the redirecting page
            ele.find_element_by_xpath('.//a').click()
			#select new window and switch
            WebDriverWait(driver, 5).until(EC.number_of_windows_to_be(2))
            newWindow = driver.window_handles
            newNewWindow = newWindow[1]
            driver.switch_to.window(newNewWindow)
    
            #sleep to wait for the page processed 
            time.sleep(round(random.uniform(2,3), 1))
            
			#locate another element
            try:
                tag = driver.find_element_by_xpath('//*[@id="divFullText"]')
                fulltext = get_text(tag)
                print('CC'+fulltext[0:4])  
                print('======='+str(index+1)+'=======')
                df_excel.loc[index,'fulltext'] = fulltext.replace('\u3000','')
            except:
                fulltext = '获取失败'
                df_excel.loc[index,'fulltext'] = fulltext.replace('\u3000','')
                print('CC'+fulltext[0:4])  
                print('======='+str(index+1)+'=======')
            
			#close current window and switch back
            driver.close()
            driver.switch_to.window(newWindow[0])
        except:
            fulltext = '获取失败'
            df_excel.loc[index,'fulltext'] = fulltext.replace('\u3000','')
            print('CC'+fulltext[0:4]) 
            print('======='+str(index+1)+'=======')
        
		#increment excel index
        index = index + 1
                
    df_excel.to_excel(writer,'Sheet1',index=False)
    writer.save()
    
	#limit times of clicking next page with page number
    if page<=xx-2:    
        driver.find_element_by_xpath('//*[@id="btnPageNext"]').click()
    
    time.sleep(12)
          
df_excel.to_excel(writer,'Sheet1',index=False)
writer.save()


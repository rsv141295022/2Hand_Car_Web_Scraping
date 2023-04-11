import numpy as np
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import time
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException

EXCEL_PATH = r'C:\Users\patcharapol.y\Desktop\car_ttb\car_price.xlsx'
col_name = ['ลำดับ','ประเภทรถ','ยี่ห้อ','กลุ่ม','รุ่น','เกียร์','ปี','ราคา']
pd.DataFrame([], columns=col_name).to_excel(EXCEL_PATH, index=False)

def stop():
    time.sleep((np.random.rand()+0.25))

driver = webdriver.Chrome()
driver.get("https://bluebook.ttbbank.com/search.asp")
stop()

driver.find_element(By.XPATH, '//a[@href="#"]').click()
stop()

car_info = []
select_type = Select(driver.find_element(By.XPATH, '//*[@id="selectType"]'))
for i in range(1, len(select_type.options)):
    
    select_type = Select(driver.find_element(By.XPATH, '//*[@id="selectType"]'))
    select_type.select_by_index(i)
    stop()
    
    select_brand = Select(driver.find_element(By.XPATH, '//*[@id="selectBrand"]'))
    for j in range(1, len(select_brand.options)):
        select_brand = Select(driver.find_element(By.XPATH, '//*[@id="selectBrand"]'))
        select_brand.select_by_index(j)
        stop()

        img = driver.find_element(By.XPATH, '//img[@src="images/btt.gif"]')
        ActionChains(driver).move_to_element(img).click().perform()
        stop()
        
        for page in range(1,300):
            window_before = driver.window_handles[0]
            x_path_tr = '//html/body/div/table[2]/tbody/tr[1]/td/table/tbody/tr[7]/td/table/tbody/tr'
            trs = driver.find_elements(By.XPATH, x_path_tr)
            
            # Get 
            for row in range(2, len(trs) - 2):
                data_list = []
                for col in range(1, 8):
                    td = driver.find_element(By.XPATH, x_path_tr + f'[{row}]/td[{col}]')
                    data_list.append(td.text)
                
                # Open Popup to Extract Price
                try:
                    ActionChains(driver).move_to_element(driver.find_element(By.XPATH, x_path_tr + f'[{row}]')).click().perform()
                    window_after = driver.window_handles[1]
                    driver.switch_to.window(window_after)
                    stop()
                    td = driver.find_element(By.XPATH, '/html/body/form/div/table/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr[7]/td[2]')
                    data_list.append(td.text[:-4])
                    print(data_list)
                    car_info.append(data_list)
                    driver.close()
                    driver.switch_to.window(window_before)
                except:
                    pass
                    print("pass")
                    
            # Click next page
            try:
                driver.find_element(By.XPATH, f'/html/body/div/table[2]/tbody/tr[1]/td/table/tbody/tr[7]/td/table/tbody/tr[18]/td/a[contains(text(), "{page + 1}")]').click()
            except NoSuchElementException:
                break
        
        df1 = pd.read_excel(EXCEL_PATH, header=0)
        df2 = pd.DataFrame(car_info, columns=col_name)
        df3 = pd.concat([df1, df2]).to_excel(EXCEL_PATH, index=False)


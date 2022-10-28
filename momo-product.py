# import moudle
import selenium
from selenium import webdriver
import time
import re
import os
import bs4
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

# setting web driver path
options = webdriver.ChromeOptions()
driver = webdriver.Chrome('C:/Users/1900422/.ipython/chromedriver.exe',options = options)

# link to momo
url = "https://www.momoshop.com.tw/main/Main.jsp"
driver.get(url)

# search
search = driver.find_element(By.ID, "keyword")
search.clear()
search.send_keys("專科")
search.send_keys(Keys.RETURN)

# clasify
cateTooth = WebDriverWait(driver, 10).until(
    EC.presence_of_all_elements_located((By.XPATH, "/html/body/div[1]/div[4]/div[3]/table[1]/tbody/tr[1]/td/div/ul/li[3]/a"))
)
cateTooth = driver.find_element(By.XPATH, ("/html/body/div[1]/div[4]/div[3]/table[1]/tbody/tr[1]/td/div/ul/li[3]/a")).click()
root = bs4.BeautifulSoup(driver.page_source, "html.parser")

PrdCode_list = []
PrdName_list = []

#crawler product imformation
    # Page1
prdCode = WebDriverWait(driver, 10).until(
    EC.presence_of_all_elements_located((By.XPATH, '//*[@id="BodyBase"]/div[4]/div[5]/div[4]/ul/li[1]/a/div[2]/h3'))
)
goods_code = driver.find_elements(By.XPATH, '//*[@id="BodyBase"]/div[4]/div[5]/div[4]/ul/li')
for code in goods_code:
    PrdCode = code.get_attribute("gcode")
    PrdCode_list.append(PrdCode)
goods_name = driver.find_elements(By.CLASS_NAME, "prdName")
for name in goods_name:
    PrdName = name.text
    PrdName_list.append(PrdName)

    # Page2
driver.find_element(By.PARTIAL_LINK_TEXT, "下一頁").click()
prdCode = WebDriverWait(driver, 10).until(
    EC.presence_of_all_elements_located((By.XPATH, '//*[@id="BodyBase"]/div[4]/div[5]/div[4]/ul/li[1]/a/div[2]/h3'))
)
goods_code = driver.find_elements(By.XPATH, '//*[@id="BodyBase"]/div[4]/div[5]/div[4]/ul/li')
for code in goods_code:
    PrdCode = code.get_attribute("gcode")
    PrdCode_list.append(PrdCode)
goods_name = driver.find_elements(By.CLASS_NAME, "prdName")
for name in goods_name:
    PrdName = name.text
    PrdName_list.append(PrdName)

# Pandas
df = pd.DataFrame(data ={"PrdCode" : PrdCode_list,"PrdName" : PrdName_list})
df.to_excel('MOMOproduct.xlsx')
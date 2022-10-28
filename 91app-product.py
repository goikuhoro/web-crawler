# import moudle
from asyncio.windows_events import NULL
import selenium
from selenium import webdriver
import time
import re
import os
import bs4
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

# setting web driver path
options = webdriver.ChromeOptions()
driver = webdriver.Chrome('C:/Users/1900422/.ipython/chromedriver.exe',options = options)

# link to fbshop
url = "https://www.fbshop.com.tw/"
driver.get(url)
time.sleep(3)
driver.refresh()

# search
clasify = WebDriverWait(driver, 10).until(
    EC.presence_of_all_elements_located((By.PARTIAL_LINK_TEXT, "商品分類"))
)
clasify = driver.find_element(By.PARTIAL_LINK_TEXT, "商品分類")
ActionChains(driver).move_to_element(clasify).perform()
Drformula = driver.find_element(By.PARTIAL_LINK_TEXT, "選品牌_Dr's Formula生活用品").click()
shampoo = driver.find_element(By.PARTIAL_LINK_TEXT, "此分類全部商品").click()
for i in range(5):
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(3)

#crawler product imformation
PrdName_list = []
PrdName = driver.find_elements(By.CLASS_NAME, "sc-fzXfNQ.ckeLKw")
for Name in PrdName:
    PrdName = Name.text
    PrdName_list.append(PrdName)
print(PrdName_list)

SuggestProductPrice_list = []
NormalProductPrice_list = []
ProductPrice_list = []
ProductPrice = driver.find_elements(By.CLASS_NAME, "sc-fzXfNR.cknioF")
for Price in ProductPrice :
    ProductPrice = Price.text.split('\n')
    ProductPrice_list.append(ProductPrice)
    if len(ProductPrice) == 1 :
        ProductPrice = ProductPrice * 2
    SuggestProductPrice_list.append(ProductPrice[0])
    NormalProductPrice_list.append(ProductPrice[1])

ProductCode_list = []
ProductCode = driver.find_elements(By.CSS_SELECTOR, 'div.product-card__vertical a')
for code in ProductCode :
    ProductCode = code.get_attribute("href")
    ProductCode = ProductCode.replace('https://www.fbshop.com.tw/SalePage/Index/', "")
    ProductCode_list.append(ProductCode)
print(ProductCode_list)

# store to pandas
df = pd.DataFrame(data ={"ProductCode" : ProductCode_list, "PrdName" : PrdName_list, "SuggestPrice" : SuggestProductPrice_list, "NormalPrice" : NormalProductPrice_list})
df.to_excel('91appproduct.xlsx')
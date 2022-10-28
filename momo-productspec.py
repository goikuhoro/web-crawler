# import moudle
import selenium
from selenium import webdriver
import time
import re
import os
import wget
import bs4
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import execjs
import urllib.request as req
import requests
import pandas as pd
import numpy as np

# read excel
df_PrdCode = pd.read_excel('MOMOproduct.xlsx', usecols = "B")
df_PrdCode = df_PrdCode['PrdCode'].values.tolist()
for code in df_PrdCode:
    url = "https://www.momoshop.com.tw/goods/GoodsDetail.jsp?i_code={}".format(code)

    # setting web driver path
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome('C:/Users/1900422/.ipython/chromedriver.exe',options = options)
    
    # link to product page
    driver.get(url)

    #crawler product imformation
    goodscode_list = []
    prdCode = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.CLASS_NAME, 'btnArea'))
    )
    goodscode = driver.find_element(By.XPATH, '//*[@id="categoryActivityInfo"]/li[1]')
    goods_code = goodscode.text
    goodscode_list.append(goods_code)

    goodsname_list = []
    goodsname = driver.find_element(By.CLASS_NAME, "fprdTitle")
    goods_name = goodsname.text
    goodsname_list.append(goods_name)
    
    goodsprice_list = []
    goodsprice = driver.find_element(By.CLASS_NAME, "prdPrice")
    goods_price = goodsprice.text
    goodsprice_list.append(goods_price)

    # Spec
    spec_list = []
    driver.execute_script("document.getElementsByClassName('vendordetailview specification')[0].style.display='block';") # display
    time.sleep(5)
    table = driver.find_element(By.XPATH, '//*[@id="attributesTable"]/tbody')
    spec = table.text
    spec_list.append(spec)
    
    # Pandas
    df = pd.DataFrame(data = {"PrdCode" : goodscode_list, "PrdName" : goodsname_list, "Price" : goodsprice_list, "spec" : spec_list})
    df.to_excel('MOMOproductspec.xlsx')

    # Picture
    driver.execute_script("document.getElementsByClassName('vendordetailview')[0].style.display='block';") # display
    time.sleep(5)
    picurl = driver.find_element(By.XPATH, '//*[@id="ifrmGoods"]')
    pics = picurl.get_attribute("src")
    request = req.Request(pics, headers={
        "User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36"
    })
    with req.urlopen(request) as response :
        data = response.read()
    import bs4
    root = bs4.BeautifulSoup(data, "html.parser")
    directory = "EDM"
    parent_dir = "C:/Users/1900422/.ipython"
    path = os.path.join(parent_dir, directory)
    images = root.find_all('img')
    link_list = []
    count = 0
    for img in images:
        link = 'https:' + img['src']
        save_as = os.path.join(path, '{}'.format(code) + '_' + str(count) + '.jpg')
        wget.download(link, save_as)
        count += 1
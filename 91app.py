# import moudle
import urllib.request as req
import json
from selenium import webdriver
import time
import re
import os
import bs4
import openpyxl
import xlrd
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

# url1
url = "https://fts-api.91app.com/pythia-cdn/graphql?shopId=40027&lang=zh-TW&query=query%20cms_shopCategory(%24shopId%3A%20Int!%2C%20%24categoryId%3A%20Int!%2C%20%24startIndex%3A%20Int!%2C%20%24fetchCount%3A%20Int!%2C%20%24orderBy%3A%20String%2C%20%24isShowCurator%3A%20Boolean%2C%20%24locationId%3A%20Int%2C%20%24tagFilters%3A%20%5BItemTagFilter%5D%2C%20%24tagShowMore%3A%20Boolean%2C%20%24serviceType%3A%20String%2C%20%24minPrice%3A%20Float%2C%20%24maxPrice%3A%20Float%2C%20%24payType%3A%20%5BString%5D%2C%20%24shippingType%3A%20%5BString%5D)%20%7B%0A%20%20shopCategory(shopId%3A%20%24shopId%2C%20categoryId%3A%20%24categoryId)%20%7B%0A%20%20%20%20salePageList(startIndex%3A%20%24startIndex%2C%20maxCount%3A%20%24fetchCount%2C%20orderBy%3A%20%24orderBy%2C%20isCuratorable%3A%20%24isShowCurator%2C%20locationId%3A%20%24locationId%2C%20tagFilters%3A%20%24tagFilters%2C%20tagShowMore%3A%20%24tagShowMore%2C%20minPrice%3A%20%24minPrice%2C%20maxPrice%3A%20%24maxPrice%2C%20payType%3A%20%24payType%2C%20shippingType%3A%20%24shippingType%2C%20serviceType%3A%20%24serviceType)%20%7B%0A%20%20%20%20%20%20salePageList%20%7B%0A%20%20%20%20%20%20%20%20salePageId%0A%20%20%20%20%20%20%20%20title%0A%20%20%20%20%20%20%20%20picUrl%0A%20%20%20%20%20%20%20%20salePageCode%0A%20%20%20%20%20%20%20%20price%0A%20%20%20%20%20%20%20%20suggestPrice%0A%20%20%20%20%20%20%20%20isFav%0A%20%20%20%20%20%20%20%20isComingSoon%0A%20%20%20%20%20%20%20%20isSoldOut%0A%20%20%20%20%20%20%20%20soldOutActionType%0A%20%20%20%20%20%20%20%20sellingQty%0A%20%20%20%20%20%20%20%20__typename%0A%20%20%20%20%20%20%7D%0A%20%20%20%20%20%20totalSize%0A%20%20%20%20%20%20shopCategoryId%0A%20%20%20%20%20%20shopCategoryName%0A%20%20%20%20%20%20statusDef%0A%20%20%20%20%20%20listModeDef%0A%20%20%20%20%20%20orderByDef%0A%20%20%20%20%20%20dataSource%0A%20%20%20%20%20%20tags%20%7B%0A%20%20%20%20%20%20%20%20isGroupShowMore%0A%20%20%20%20%20%20%20%20groups%20%7B%0A%20%20%20%20%20%20%20%20%20%20groupId%0A%20%20%20%20%20%20%20%20%20%20groupDisplayName%0A%20%20%20%20%20%20%20%20%20%20isKeyShowMore%0A%20%20%20%20%20%20%20%20%20%20keys%20%7B%0A%20%20%20%20%20%20%20%20%20%20%20%20keyId%0A%20%20%20%20%20%20%20%20%20%20%20%20keyDisplayName%0A%20%20%20%20%20%20%20%20%20%20%20%20__typename%0A%20%20%20%20%20%20%20%20%20%20%7D%0A%20%20%20%20%20%20%20%20%20%20__typename%0A%20%20%20%20%20%20%20%20%7D%0A%20%20%20%20%20%20%20%20__typename%0A%20%20%20%20%20%20%7D%0A%20%20%20%20%20%20priceRange%20%7B%0A%20%20%20%20%20%20%20%20min%0A%20%20%20%20%20%20%20%20max%0A%20%20%20%20%20%20%20%20__typename%0A%20%20%20%20%20%20%7D%0A%20%20%20%20%20%20__typename%0A%20%20%20%20%7D%0A%20%20%20%20__typename%0A%20%20%7D%0A%7D%0A&operationName=cms_shopCategory&variables=%7B%22shopId%22%3A40027%2C%22categoryId%22%3A297914%2C%22startIndex%22%3A0%2C%22fetchCount%22%3A100%2C%22orderBy%22%3A%22Sales%22%2C%22isShowCurator%22%3Afalse%2C%22locationId%22%3Anull%2C%22tagFilters%22%3A%5B%5D%2C%22tagShowMore%22%3Afalse%2C%22minPrice%22%3Anull%2C%22maxPrice%22%3Anull%2C%22payType%22%3A%5B%5D%2C%22shippingType%22%3A%5B%5D%7D"

request = req.Request(url, headers={
    "content-type":"application/json",
    "user-agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36 Edg/106.0.1370.37"
})
with req.urlopen(request) as response :
    result = response.read().decode("utf-8")
result = json.loads(result)
items = result["data"]["shopCategory"]["salePageList"]["salePageList"]

PrdName_list = []
for item in items :
    PrdName = item["title"]
    PrdName_list.append(PrdName)

ProductCode_list = []
for item in items :
    ProductCode = item["salePageId"]
    ProductCode_list.append(ProductCode)

SuggestProductPrice_list = []
for item in items :
    SuggestPrice = item["suggestPrice"]
    SuggestPrice = "NT$" + str(SuggestPrice)
    SuggestProductPrice_list.append(SuggestPrice)

NormalProductPrice_list = []
for item in items :
    NormalPrice = item["price"]
    NormalPrice = "NT$" + str(NormalPrice)
    NormalProductPrice_list.append(NormalPrice)

df1 = pd.DataFrame(data = {"ProductCode" : ProductCode_list, "PrdName" : PrdName_list, "SuggestPrice" : SuggestProductPrice_list, "NormalPrice" : NormalProductPrice_list})
# url2
url = "https://fts-api.91app.com/pythia-cdn/graphql?shopId=40027&lang=zh-TW&query=query%20cms_shopCategory(%24shopId%3A%20Int!%2C%20%24categoryId%3A%20Int!%2C%20%24startIndex%3A%20Int!%2C%20%24fetchCount%3A%20Int!%2C%20%24orderBy%3A%20String%2C%20%24isShowCurator%3A%20Boolean%2C%20%24locationId%3A%20Int%2C%20%24tagFilters%3A%20%5BItemTagFilter%5D%2C%20%24tagShowMore%3A%20Boolean%2C%20%24serviceType%3A%20String%2C%20%24minPrice%3A%20Float%2C%20%24maxPrice%3A%20Float%2C%20%24payType%3A%20%5BString%5D%2C%20%24shippingType%3A%20%5BString%5D)%20%7B%0A%20%20shopCategory(shopId%3A%20%24shopId%2C%20categoryId%3A%20%24categoryId)%20%7B%0A%20%20%20%20salePageList(startIndex%3A%20%24startIndex%2C%20maxCount%3A%20%24fetchCount%2C%20orderBy%3A%20%24orderBy%2C%20isCuratorable%3A%20%24isShowCurator%2C%20locationId%3A%20%24locationId%2C%20tagFilters%3A%20%24tagFilters%2C%20tagShowMore%3A%20%24tagShowMore%2C%20minPrice%3A%20%24minPrice%2C%20maxPrice%3A%20%24maxPrice%2C%20payType%3A%20%24payType%2C%20shippingType%3A%20%24shippingType%2C%20serviceType%3A%20%24serviceType)%20%7B%0A%20%20%20%20%20%20salePageList%20%7B%0A%20%20%20%20%20%20%20%20salePageId%0A%20%20%20%20%20%20%20%20title%0A%20%20%20%20%20%20%20%20picUrl%0A%20%20%20%20%20%20%20%20salePageCode%0A%20%20%20%20%20%20%20%20price%0A%20%20%20%20%20%20%20%20suggestPrice%0A%20%20%20%20%20%20%20%20isFav%0A%20%20%20%20%20%20%20%20isComingSoon%0A%20%20%20%20%20%20%20%20isSoldOut%0A%20%20%20%20%20%20%20%20soldOutActionType%0A%20%20%20%20%20%20%20%20sellingQty%0A%20%20%20%20%20%20%20%20__typename%0A%20%20%20%20%20%20%7D%0A%20%20%20%20%20%20totalSize%0A%20%20%20%20%20%20shopCategoryId%0A%20%20%20%20%20%20shopCategoryName%0A%20%20%20%20%20%20statusDef%0A%20%20%20%20%20%20listModeDef%0A%20%20%20%20%20%20orderByDef%0A%20%20%20%20%20%20dataSource%0A%20%20%20%20%20%20tags%20%7B%0A%20%20%20%20%20%20%20%20isGroupShowMore%0A%20%20%20%20%20%20%20%20groups%20%7B%0A%20%20%20%20%20%20%20%20%20%20groupId%0A%20%20%20%20%20%20%20%20%20%20groupDisplayName%0A%20%20%20%20%20%20%20%20%20%20isKeyShowMore%0A%20%20%20%20%20%20%20%20%20%20keys%20%7B%0A%20%20%20%20%20%20%20%20%20%20%20%20keyId%0A%20%20%20%20%20%20%20%20%20%20%20%20keyDisplayName%0A%20%20%20%20%20%20%20%20%20%20%20%20__typename%0A%20%20%20%20%20%20%20%20%20%20%7D%0A%20%20%20%20%20%20%20%20%20%20__typename%0A%20%20%20%20%20%20%20%20%7D%0A%20%20%20%20%20%20%20%20__typename%0A%20%20%20%20%20%20%7D%0A%20%20%20%20%20%20priceRange%20%7B%0A%20%20%20%20%20%20%20%20min%0A%20%20%20%20%20%20%20%20max%0A%20%20%20%20%20%20%20%20__typename%0A%20%20%20%20%20%20%7D%0A%20%20%20%20%20%20__typename%0A%20%20%20%20%7D%0A%20%20%20%20__typename%0A%20%20%7D%0A%7D%0A&operationName=cms_shopCategory&variables=%7B%22shopId%22%3A40027%2C%22categoryId%22%3A297914%2C%22startIndex%22%3A100%2C%22fetchCount%22%3A100%2C%22orderBy%22%3A%22Sales%22%2C%22isShowCurator%22%3Afalse%2C%22locationId%22%3Anull%2C%22tagFilters%22%3A%5B%5D%2C%22tagShowMore%22%3Afalse%2C%22minPrice%22%3Anull%2C%22maxPrice%22%3Anull%2C%22payType%22%3A%5B%5D%2C%22shippingType%22%3A%5B%5D%7D"

request = req.Request(url, headers={
    "content-type":"application/json",
    "user-agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36 Edg/106.0.1370.37"
})
with req.urlopen(request) as response :
    result = response.read().decode("utf-8")
result = json.loads(result)
items = result["data"]["shopCategory"]["salePageList"]["salePageList"]

PrdName_list = []
for item in items :
    PrdName = item["title"]
    PrdName_list.append(PrdName)

ProductCode_list = []
for item in items :
    ProductCode = item["salePageId"]
    ProductCode_list.append(ProductCode)

SuggestProductPrice_list = []
for item in items :
    SuggestPrice = item["suggestPrice"]
    SuggestProductPrice_list.append(SuggestPrice)

NormalProductPrice_list = []
for item in items :
    NormalPrice = item["price"]
    NormalProductPrice_list.append(NormalPrice)

df2 = pd.DataFrame(data = {"ProductCode" : ProductCode_list, "PrdName" : PrdName_list, "SuggestPrice" : SuggestProductPrice_list, "NormalPrice" : NormalProductPrice_list})

path = os.path.join(os.getcwd(), '91app.xlsx')
writer = pd.ExcelWriter(path, engine='openpyxl')
df1.to_excel(writer, sheet_name='url1')
df2.to_excel(writer, sheet_name='url2')
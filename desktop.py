from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from PIL import Image, ImageChops, ImageStat
import pandas as pd
import time
import os

# getting before screenshot
# i gives the row numbers
def getBeforeScreenshot(url, i):
    chrome_options = Options ()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options= chrome_options)
    driver.get(url)
    time.sleep(2)
    width = 1320 #driver.execute_script('return document.body.parentNode.scrollWidth')
    height = driver.execute_script('return document.body.parentNode.scrollHeight')
    driver.set_window_size(width, height)
    time.sleep(2)
    driver.get_screenshot_as_file('./Before/Chrome_bf_'+str(i)+'.png')
    driver.quit()

# getting after screenshot
def getAfterScreenshot(url, i):
    chrome_options = Options ()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options= chrome_options)
    driver.get(url)
    url_list.append(url)
    time.sleep(2)
    width = 1320 #driver.execute_script('return document.body.parentNode.scrollWidth')
    height = driver.execute_script('return document.body.parentNode.scrollHeight')
    driver.set_window_size(width, height)
    driver.get_screenshot_as_file('./After/Chrome_af_'+str(i)+'.png')
    driver.quit()

    # image comparison section
    beforeImage = ('./Before/Chrome_bf_'+str(i)+'.png')
    afterImage = ('./After/Chrome_af_'+str(i)+'.png')
    diffImage = ('./Diff/Chrome_df_'+str(i)+'.png')
    Img1 = Image.open(beforeImage).convert('RGB')
    Img2 = Image.open(afterImage).convert('RGB')
    imageDifference = ImageChops.difference(Img1,Img2)
    stat = ImageStat.Stat(imageDifference)
    diffValue = sum(stat.mean)
    if diffValue > 1:
        result = 'Fail'
    else:
        result = 'Pass'
    similarity_list.append (diffValue)
    result_list.append (result)
    imageDifference.save(diffImage)

# importing a data from excel named Book1.xlsx and sheet name is Sheet1
excelData = pd.read_excel("Book1.xlsx", sheet_name="Sheet1")

# creating a list for value storage
url_list = []
similarity_list =[]
result_list =[]

# making directory for saving screenshots
if not (os.path.isdir('./Before') and os.path.isdir('./After') and os.path.isdir('./Diff')):
    os.makedirs('Before')
    os.makedirs('After')
    os.makedirs('Diff')

# for loop for calling the function
for x in excelData.index:
    # before code drop
    url = excelData['Links'][x]
    getBeforeScreenshot(url, x+2)
    # after code drop
    getAfterScreenshot(url, x+2)

# saving the result in excel file called output.xlsx
result_df = pd.DataFrame({'links':url_list, 'Difference': similarity_list, 'result':result_list})
writer = pd.ExcelWriter('output.xlsx')
result_df.to_excel(writer)
writer.save()

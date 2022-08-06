from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from PIL import Image, ImageChops, ImageStat
import pandas as pd
import time

# getting the height of the webpage in ipad emulation for that particular url
# mobile emulation details can get from chrome devtool
def gettingTheHeightOfIpad(url):
    # setting the device metrics
    mobile_emulation = {"deviceMetrics": {"width": 768, "height": 1024, "pixelRatio": 2.0},
        "userAgent": "Mozilla/5.0 (iPad; CPU OS 13_2 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) CriOS/104.0.0.0 Mobile/15E148 Safari/604.1"}
    chrome_options = Options()
    chrome_options.add_experimental_option("mobileEmulation", mobile_emulation)
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=chrome_options)
    driver.get(url)
    driver.find_element(By.TAG_NAME, 'body')
    return (driver.execute_script('return document.body.parentNode.scrollHeight'))

# getting before screenshot
def getBeforeScreenshot(url, i):
    mobile_emulation = {"deviceMetrics": {"width": 768, "height": 1024, "pixelRatio": 2.0},
                        "userAgent": "Mozilla/5.0 (iPad; CPU OS 13_2 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) CriOS/104.0.0.0 Mobile/15E148 Safari/604.1"}
    chrome_options = Options()
    chrome_options.add_experimental_option("mobileEmulation", mobile_emulation)
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-gpu")
    driver = webdriver.Chrome(options= chrome_options)
    driver.get(url)
    url_list.append(url)
    driver.find_element(By.TAG_NAME, 'body').screenshot('./Before/Chrome_bf_'+str(i)+'.png')
    driver.quit()

# getting after screenshot
def getAfterScreenshot(url, i):
    mobile_emulation = {"deviceMetrics": {"width": 768, "height": 1024, "pixelRatio": 2.0},
                        "userAgent": "Mozilla/5.0 (iPad; CPU OS 13_2 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) CriOS/104.0.0.0 Mobile/15E148 Safari/604.1"}
    chrome_options = Options()
    chrome_options.add_experimental_option("mobileEmulation", mobile_emulation)
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-gpu")
    driver = webdriver.Chrome(options=chrome_options)
    driver.get(url)
    driver.find_element(By.TAG_NAME, 'body').screenshot('./After/Chrome_af_' + str(i) + '.png')
    driver.quit()

    # image comparison section
    beforeImage = ('./Before/Chrome_bf_'+str(i)+'.png')
    afterImage = ('./After/Chrome_af_'+str(i)+'.png')
    diffImage = './Diff/Chrome_df_'+str(i)+'.png')
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
height = 0
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
    #before code drop
    url = excelData['links'][x]
    height = gettingTheHeightOfIpad(url)
    getBeforeScreenshot(url, x+2)
    # after code drop
    height = gettingTheHeightOfIpad(url)
    getAfterScreenshot(url, x+2)

# saving the result in excel file called output.xlsx
result_df = pd.DataFrame({'links':url_list, 'Difference': similarity_list, 'result':result_list})
writer = pd.ExcelWriter('output.xlsx')
result_df.to_excel(writer)
writer.save()

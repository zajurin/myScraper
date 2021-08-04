from selenium import webdriver 
from selenium.webdriver.common.keys import Keys
import time

import pyautogui

from urllib.request import urlopen
from urllib.error import HTTPError
from bs4 import BeautifulSoup

import requests
from csv import writer

# headers: {"User-Agent"="Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36"}
headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.77 Safari/537.36'}


# Opening Google
companyName = input('Tell us the name of your company ')
myAddress = input('type the address of your company ')

driver = webdriver.Chrome("C:\\Program Files (x86)\\chromedriver.exe")
driver.get("https://www.google.com")

driver.find_element_by_name('q').click()
driver.find_element_by_name('q').send_keys("phone number of " + companyName + " " + myAddress)
driver.find_element_by_name('q').send_keys(Keys.ENTER)

myCurrentURL = driver.current_url
# myPhoneNumber = driver.find_element_by_class_name("mw31Ze") # GETTING LATITUDE ****************
# print(myPhoneNumber.get_attribute('text'))

time.sleep(2)

response = requests.get(myCurrentURL, headers=headers)
soup = BeautifulSoup(response.text, 'html.parser')

tagSearched = soup.find_all('span', attrs={'class': 'mw31Ze'})
for tag in tagSearched:
    print(tag.text.strip())


# def getTitle(url):
#     try:
#         html = urlopen(url)
#     except HTTPError as e:
#         return e
#     try:
#         bsObj = BeautifulSoup(html.read())
#         title = bsObj.body
#     except AttributeError as e:
#         return None
#     return title

# print(myCurrentURL)    

# driver.find_element_by_xpath("/html/body/div/div[2]/form/div/div/div/div/div[1]/input[0]").send_keys(myAddress)
# time.sleep(4)


# try:
#     element = WebDriverWait(driver, 5).until( 
#         EC.presence_of_element_located((By.ID, "searchbox")) #Verifying the existence of search box 
#     )
#     element.click()
#     time.sleep(4)

#     driver.find_element_by_xpath("/html/body/div/div[2]/form/div/div/div/div/div[1]/div/input").click()
#     # driver.find_element_by_id('searchbox').click() 
#     # driver.find_element_by_id('searchbox').send_keys(myAddress)
#     driver.find_element_by_xpath("/html/body/div/div[2]/form/div/div/div/div/div[1]/div/input").send_keys(myAddress)
#     time.sleep(4)


#     # element = WebDriverWait(driver, 10).until( 
#     #     EC.presence_of_element_located((By.ID, "password1")) #Verifying the existence of password box 
#     # )
#     # element.click()

#     # driver.find_element_by_id('password1').send_keys(mypass)
#     driver.find_element_by_xpath("/html/body/jsl/div/div/div/div/div/div/div/button").click()


# except:
#     print('there was a mistake')
#     time.sleep(3)


# html = urlopen("http://www.pythonscraping.com/pages/page1.html")
# bsObj = BeautifulSoup(html.read())
# myUrl = "http://www.elfinanciero.com"




# title = getTitle("https://www.google.com")
# if title == None:
#     print("Title could not be found")
# else:
#     print(title)

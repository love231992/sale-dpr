from selenium import webdriver
from selenium.webdriver.chrome.options import Options

import time


def load():
    options = Options()
    options.headless = True
    driver = webdriver.Chrome(chrome_options=options)
    driver.get("http://www.punjabsldc.org/realtimepbGen.aspx")
    time.sleep(3)
    mw =[]
    power_plant = ["ippRajpura1","ippRajpura2","ippTS1","ippTS2","ippTS3","ippGVK1","ippGVK2","GGSSTP3","GGSSTP4","GGSSTP5","GGSSTP6","GHTP1","GHTP2","GHTP3","GHTP4"]
    for i in power_plant:
        element = driver.find_element_by_xpath(f'//*[@id=\"{i}\"]')
        mw.append(element.text)
    driver.close()
    return mw


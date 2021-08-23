from Adresse import Adresse
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from bs4 import BeautifulSoup
import requests
import xlsxwriter
import tkinter
from tkinter import *
import time
import xlrd


plz = "30823"
city = "Garbsen"

streetOptions = []
numberOptions = []

driver = webdriver.Chrome()


def main():
    driver.get("http://127.0.0.1:5500/hv/city_option_list.html")
    select = driver.find_element_by_name("adrList")
    for option in select.find_elements_by_tag_name('option'):
        adr = option.text
        print(adr)
        arr = adr.split(',')
        print(len(arr))
        # streetOptions.append(Adresse(arr[0], arr[1], arr[2], arr[3], ''))
        # for adr in streetOptions:
        #     print(adr.plz)


main()
driver.close()

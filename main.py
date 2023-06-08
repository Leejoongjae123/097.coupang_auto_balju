import json

import openpyxl
import pandas as pd
from pyautogui import size
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import subprocess
import shutil
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from bs4 import BeautifulSoup
import time
import datetime
import pyautogui
import pyperclip
import csv
import sys
import os
import math
import requests
import re
import random
import chromedriver_autoinstaller
from PyQt5.QtWidgets import QWidget, QApplication, QTreeView, QFileSystemModel, QVBoxLayout, QPushButton, QInputDialog, \
    QLineEdit, QMainWindow, QMessageBox, QFileDialog
from PyQt5.QtCore import QCoreApplication
from selenium.webdriver import ActionChains
from datetime import datetime, date, timedelta
import numpy
import datetime
# from window import Ui_MainWindow
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
import pprint


def load_excel(fname,startNumber,endNumber):
    # fname = 'result.xlsx'
    wb = openpyxl.load_workbook(fname)
    ws = wb.active
    no_row = ws.max_row
    print("행갯수:", no_row)
    data_list = []
    for i in range(startNumber, endNumber+1):
        sellerNo = ws.cell(row=i, column=3).value
        if sellerNo == None:
            print('데이타 더 이상 없음')
            break
        receiverName = ws.cell(row=i, column=4).value
        receiverPhone = ws.cell(row=i, column=5).value
        basicAddress = ws.cell(row=i, column=6).value
        detailAddress = ws.cell(row=i, column=7).value
        productName=ws.cell(row=i, column=8).value
        productQuantity = int(ws.cell(row=i, column=9).value)
        url = ws.cell(row=i, column=11).value
        data = {'sellerNo':sellerNo,'receiverName':receiverName,'receiverPhone':receiverPhone,'basicAddress':basicAddress,'detailAddress':detailAddress,'productName':productName,'productQuantity':productQuantity,'url':url,'rowNum':i}
        data_list.append(data)
    print(data_list)
    return data_list
def chrome_browser(url):
    chrome_ver = chromedriver_autoinstaller.get_chrome_version().split('.')[0]
    driver_path = f'./{chrome_ver}/chromedriver.exe'

    if os.path.exists(driver_path):
        print(f"chromedriver is installed: {driver_path}")
    else:
        print(f"install the chrome driver(ver: {chrome_ver})")
        chromedriver_autoinstaller.install(True)

    try:
        shutil.rmtree(r"c:\chrometemp")  # 쿠키 / 캐쉬파일 삭제
    except FileNotFoundError:
        pass

    subprocess.Popen(
        r'C:\Program Files\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\chrometemp"')  # 디버거 크롬 구동

    print("옵션설정")
    option = Options()
    user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36'
    option.add_argument('user-agent='+user_agent)
    option.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    option.add_argument("--disable - blink - features = AutomationControlled")
    chrome_ver = chromedriver_autoinstaller.get_chrome_version().split('.')[0]
    try:
        browser = webdriver.Chrome(driver_path, options=option)
        browser.maximize_window()
        url_start = url
        browser.get(url_start)
    except:
        chromedriver_autoinstaller.install(True)
        browser = webdriver.Chrome(options=option)
    return browser


firstFlag=True
fName='list.xlsx'
startNumber=6
endNumber=99999
productList=load_excel(fName,startNumber,endNumber)
wb=openpyxl.load_workbook(fName)
ws=wb.active


url_login='https://login.coupang.com/login/login.pang?rtnUrl=https%3A%2F%2Fwww.coupang.com%2Fnp%2Fpost%2Flogin%3Fr%3Dhttps%253A%252F%252Fwww.coupang.com%252F'
browser=chrome_browser(url_login)


# id='ljj3347@naver.com'
# pw='dlwndwo2'
id='lek740815@naver.com'
pw='1q2w3e4r5t@'


browser.implicitly_wait(3)
input_id=browser.find_element(By.ID,'login-email-input')
input_id.send_keys(id)
time.sleep(0.5)

input_pw=browser.find_element(By.ID,'login-password-input')
input_pw.send_keys(pw)
time.sleep(0.5)

btn_login=browser.find_element(By.CSS_SELECTOR,'body > div.member-wrapper.member-wrapper--flex > div.member-main > div.member-login._loginRoot.sms-login-target.style-v2 > form > div.login__content.login__content--trigger > button')
browser.execute_script("arguments[0].click();", btn_login)  #

while True:
    print("로그인시도중..")
    soup=BeautifulSoup(browser.page_source,'lxml')
    isInputKeyword=len(soup.find_all('input',attrs={'id':'headerSearchKeyword'}))
    if isInputKeyword>=1:
        print("로그인완료")
        break
    time.sleep(1)


for productElem in productList:
    url_purchase=productElem['url']
    browser.get(url_purchase)
    browser.implicitly_wait(3)

    try:
        btnUp=browser.find_element(By.CLASS_NAME,'prod-quantity__plus')
        actionBtn=productElem['productQuantity']-1
        print("증가횟수:",actionBtn)
        for i in range(0,actionBtn):
            print(f"{i}번째 클릭")
            btnUp = browser.find_element(By.CLASS_NAME, 'prod-quantity__plus')
            btnUp.click()
            time.sleep(3)
    except:
        print("갯수 증가 안됨")

    # 구매가능갯수 파악하기
    soup=BeautifulSoup(browser.page_source,'lxml')
    scripts=soup.find_all('script')
    result=""
    for script in scripts:
        if str(script).find("wishList")>=0:
            result=script
            break

    rawscript=str(result)
    # print(rawscript)
    splitPositionFr=rawscript.find("=")
    splitPositionRr=rawscript.find(";")
    # print(splitPositionFr,splitPositionRr)
    rawscriptChanged=rawscript[splitPositionFr+1:splitPositionRr].strip()
    # print(rawscriptChanged)
    jsonRawScript=json.loads(rawscriptChanged)
    # pprint.pprint(jsonRawScript)
    buyableQuantity=jsonRawScript['buyableQuantity']
    print("구매가능수량:",buyableQuantity)
    print("구매할수량:",productElem['productQuantity'])


    if buyableQuantity==None:
        print('구매불가능')
        ws.cell(row=productElem['rowNum'], column=10).value = "주문불가"
        wb.save(fName)
        continue
    elif buyableQuantity==0 or buyableQuantity<productElem['productQuantity']:
        print('구매불가능')
        ws.cell(row=productElem['rowNum'], column=10).value = "주문불가"
        wb.save(fName)
        continue


    btnPurchase=browser.find_element(By.CLASS_NAME,'prod-buy-btn__txt')
    btnPurchase.click()
    print("구매버튼클릭하기")
    browser.implicitly_wait(5)

    while True:
        try:
            btnAddress=browser.find_element(By.CSS_SELECTOR,'#body > div.middle > div:nth-child(4) > h2 > button')
            btnAddress.click()
            time.sleep(1)
            print("구매버튼누르기성공")
            break
        except:
            print("구매버튼누르기실패")
        time.sleep(1)



    while True:
        print("창 뜨는것 대기, 창갯수:{}".format(len(browser.window_handles)))
        if len(browser.window_handles)>=2:

            break
        time.sleep(0.5)

    browser.switch_to.window(browser.window_handles[-1])
    print("마지막창으로 이동")


    btnModify=browser.find_element(By.CSS_SELECTOR,'body > div > div > div.content-wrapper > div.content-body.content-body--fixed > div > div.address-card.address-card--picked > div.address-card__foot > form.address-card__form.address-card__form--edit._addressBookAddressCardEditForm > button')
    btnModify.click()
    browser.implicitly_wait(3)

    btnDelete=browser.find_element(By.CSS_SELECTOR,'body > div > div > div.content-wrapper > div.content-body.content-body--fixed > div.content-body__corset > div > form > button')
    btnDelete.click()
    browser.implicitly_wait(3)


    interval=1
    prev_height=browser.execute_script('return document.body.scrollHeight')
    while True:
        browser.execute_script('window.scrollTo(0,document.body.scrollHeight)')
        time.sleep(interval)
        curr_height = browser.execute_script('return document.body.scrollHeight')
        if curr_height == prev_height:
            break
            prev_height=curr_height



    btnAddAddress=browser.find_element(By.CSS_SELECTOR,'body > div > div > div.content-wrapper > div.content-body.content-body--fixed > div > form > div > button')
    btnAddAddress.click()
    # browser.execute_script("arguments[0].click();", btnAddAddress)  #
    time.sleep(1)



    inputName=browser.find_element(By.ID,'addressbookRecipient')
    inputName.send_keys(productElem['receiverName'])
    time.sleep(0.5)
    print("이름입력")

    inputPhone=browser.find_element(By.ID,'addressbookCellphone')
    inputPhone.send_keys(productElem['receiverPhone'])
    time.sleep(0.5)
    print("전화번호입력")

    searchAddress=browser.find_element(By.CSS_SELECTOR,'body > div > div > div.content-wrapper > div.content-body.content-body--fixed > div.content-body__corset > form > div.icon-text-field__frame-box._addressBookAddressErrorStatus > div > div.icon-text-field__button-container > a')
    searchAddress.click()
    browser.implicitly_wait(3)
    print("서칭하기2")


    soup=BeautifulSoup(browser.page_source,'lxml')
    # print(soup.prettify())

    element = browser.find_element(By.CLASS_NAME,"identity__iframe")
    browser.switch_to.frame(element) #프레임 이동

    inputAddressBasic=browser.find_element(By.NAME,'searchKey')
    inputAddressBasic.send_keys(productElem['basicAddress'])
    time.sleep(0.5)
    print("기본주소정보입력")

    btnSearch=browser.find_element(By.CSS_SELECTOR,'body > section > div.zipcode__wrapper > div > div > header > div > form > div.zipcode__search-trigger > button')
    browser.execute_script("arguments[0].click();", btnSearch)  #
    browser.implicitly_wait(3)
    print("서치버튼누르기3")
    btnRowAddress=browser.find_element(By.CSS_SELECTOR,'body > section > div.zipcode__wrapper > div > div > div > div.zipcode__slide-view._zipcodeResultSlideRoot > div.zipcode__slide-track._zipcodeResultSlide > div.zipcode__slide-item.zipcode__slide-item--address._zipcodeResultSlideItem > div._zipcodeResultListAddress > div:nth-child(1) > span.zipcode__result__item.zipcode__result__item--road._zipcodeResultSendTrigger')
    browser.execute_script("arguments[0].click();", btnRowAddress)  #
    browser.implicitly_wait(3)
    print("첫번째행선택")


    browser.switch_to.default_content()
    soup=BeautifulSoup(browser.page_source,'lxml')
    # print(soup.prettify())
    inputDetail=browser.find_element(By.CSS_SELECTOR,'#addressbookAddressDetail')
    inputDetail.send_keys(productElem['detailAddress'])
    time.sleep(0.5)

    checkBase=browser.find_element(By.ID,'_addressBookSaveAsDefault')
    browser.execute_script("arguments[0].click();", checkBase)  #
    time.sleep(0.5)

    btnSave=browser.find_element(By.CSS_SELECTOR,'body > div > div > div.content-wrapper > div.content-body.content-body--fixed > div.content-body__corset > form > div.addressbook__button-fixer > button')
    browser.execute_script("arguments[0].click();", btnSave)  #
    time.sleep(0.5)

    browser.switch_to.window(browser.window_handles[0])


    if firstFlag==True:
        paymentType=browser.find_element(By.CSS_SELECTOR,'#payType9')
        browser.execute_script("arguments[0].click();", paymentType)  #
        time.sleep(0.5)

        btnChangeInfo=browser.find_element(By.CSS_SELECTOR,'#body > div.middle > div:nth-child(8) > div > div > button')
        btnChangeInfo.click()
        time.sleep(0.5)

        selectOption=browser.find_element(By.CSS_SELECTOR,'#body > div.middle > div:nth-child(8) > div > div > div:nth-child(2) > div.cash-receipt__resiter-type__wrap > span:nth-child(1) > select')
        selectOption.click()
        time.sleep(0.5)

        selectSaupja=browser.find_element(By.CSS_SELECTOR,'#body > div.middle > div:nth-child(8) > div > div > div:nth-child(2) > div.cash-receipt__resiter-type__wrap > span:nth-child(1) > select > option:nth-child(2)')
        selectSaupja.click()
        time.sleep(0.5)

        inputSaupja=browser.find_element(By.CSS_SELECTOR,'#body > div.middle > div:nth-child(8) > div > div > div:nth-child(2) > div.cash-receipt__resiter-type__wrap > span:nth-child(2) > input')

        inputSaupja.click()
        time.sleep(0.5)
        ActionChains(browser).key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).perform()
        time.sleep(0.5)
        ActionChains(browser).send_keys('delete')

        inputSaupja.send_keys(productElem['sellerNo'])
    firstFlag=False

    btnPay=browser.find_element(By.CSS_SELECTOR,'#paymentBtn')
    browser.execute_script("arguments[0].click();", btnPay)  #
    successFlag=False
    successCount=0
    while True:
        soup=BeautifulSoup(browser.page_source,'lxml')
        if successCount>=30:
            print("조회 한도 초과로 실패")
            break
        # print(soup.prettify())
        try:
            payStatus=soup.find('span',attrs={'class':'i18n-wrapper'})
            payStatusText=payStatus.get_text()
            print('payStatusText:',payStatusText)
            if payStatusText.find("완료")>=0:
                print("주문완료됨")
                successFlag=True
                break
        except:
            print(f"아직안뜸_{successCount}")
        successCount+=1

        time.sleep(1)

    if successFlag==True:
        print("주문완료로 엑셀에 저장")
        ws.cell(row=productElem['rowNum'],column=10).value="주문성공"
    else:
        print("주문실패로 엑셀에 저장")
        ws.cell(row=productElem['rowNum'], column=10).value = "주문불가"
    wb.save(fName)
    determinant=input("또 결제할까요?")
    if determinant=="Y":
        print("반복")






















# -*- coding: utf-8 -*-
"""
Created on Wed Dec 25 21:42:05 2019

@author: user
"""

# =============================================================================
# Import Library
# =============================================================================
import datetime
import pandas as pd
from hanziconv import HanziConv
from selenium import webdriver
from selenium.webdriver.chrome.options import Options #for headless 
from selenium.webdriver.common.by import By
import os #for directory
import xlwings as xw
from GoogleHomemadeAPI import Create_Service
import win32com.client as win32

def setHeadless(mainPAth,headless = False):
    chromeOption = Options()
    
    if(headless):
        print("Headless")
        chromeOption.add_argument('--headless')
    else:
        print("Not Headless")
    
    driver=webdriver.Chrome(executable_path=mainPAth + "\chromedriver.exe",options=chromeOption)
    
    return driver

def xpathTableHeader(i,j):
#    return '//*[@id="alternatecolor"]/tbody/tr['+str(i)+']/td['+str(i)+']'
    return '/html/body/center/table/tbody/tr['+str(i)+']/th['+str(j)+']'

def xpathTable(i,j):
    return '/html/body/center/table/tbody/tr['+str(i)+']/td['+str(j)+']'

def downloadIPO(mydir,today):
    if isinstance(today,str): 
        if today != "":
            today = datetime.datetime.strptime(today,"%Y-%m-%d")
    
    ipo = "http://www.ipohk.com.cn/"
    
    log = open("log.txt","a") 
    
    try:
        driver = setHeadless(mydir,False)
    except Exception:
        log.write("[{}] #Error: cannot open web driver.\n".format(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
        return
    
    driver.get(ipo)
    run = True
    col = 1
    header = ['Update Time']
    dateIdx = -1
    stockCodeIdx = -1
    
    while run:
        try:
            tmp = driver.find_element_by_xpath(xpathTableHeader(1,col)).text
            if tmp == "上市日期": dateIdx = col
            if tmp == "代码": stockCodeIdx = col
            tmp = HanziConv.toTraditional(tmp)
            header.append(tmp)
            
            col += 1
        except Exception:
            run = False
    
    df = pd.DataFrame(columns = header)
    run = True
    i = 2
    while run:
        try:
            data = [datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')]
            for j in [x+1 for x in range(col-1)]:
                tmp = driver.find_element_by_xpath(xpathTable(i,j)).text
                tmp = HanziConv.toTraditional(tmp)
                data.append(tmp)
            
            if isinstance(today,datetime.date):
                if datetime.datetime.strptime(data[dateIdx],"%Y-%m-%d") < today: 
                    run = False
            
            if run:
                dfTmp = pd.DataFrame([data],columns = header)
                df = df.append(dfTmp , ignore_index=True)
                log.write("[{}] Finished download data for stock {}.\n".format(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),data[stockCodeIdx]))
            i+=2 #it skips 1 row
        except Exception:
            log.write("[{}] error in row {}.\n".format(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),i))
            run = False
    
    driver.quit()
    df.to_csv(mydir+"\pythonOutput.csv" ,encoding='utf_8_sig')
    if today == "": today = datetime.datetime.now()
    log.write("[{}] Finished program with date {}.\n".format(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),today.strftime('%Y-%m-%d')))
    log.close()

def getLatest(file = "log.txt"):
    try:
        log = open("log.txt","r")
        dates = []
        for line in log: 
            if "Finished program with date" in line:
                dates.append(datetime.datetime.strptime(line.split()[-1][:-1],'%Y-%m-%d'))
        
        return max(dates).strftime('%Y-%m-%d')
    except Exception:
        return ""

def updateDB(mydir):
    book = xw.Book(mydir + "/database.xlsm")
    myMarco = book.macro("update")
    myMarco(False)
    book.save()
    app = xw.apps.active 
    app.quit()

def updateGoogleSheet(mydir):
    xlApp = win32.Dispatch('Excel.Application')
    wb = xlApp.Workbooks.Open(mydir+"/database.xlsm")
    ws = wb.Worksheets('DB')
    rngData = ws.Range('A1').CurrentRegion()
    
    # Google Sheet Id
    gsheet_id = '1XH-OJBcOr7GlvGu9yFnLU4blGSNnpWTRnYdIjhKvjf8'
    CLIENT_SECRET_FILE = 'client_secret.json'
    API_SERVICE_NAME = 'sheets'
    API_VERSION = 'v4'
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    service = Create_Service(CLIENT_SECRET_FILE, API_SERVICE_NAME, API_VERSION, SCOPES)
    
    response = service.spreadsheets().values().append(
        spreadsheetId=gsheet_id,
        valueInputOption='RAW',
        range='DB!A1',
        body=dict(
            majorDimension='ROWS',
            values=rngData
        )
    ).execute()


if __name__=="__main__":
    mydir = os.getcwd()
    downloadIPO(mydir,getLatest())
    updateDB(mydir)
    updateGoogleSheet(mydir)

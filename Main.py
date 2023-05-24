from bs4 import BeautifulSoup
import pandas as pd
import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
import tkinter as tk
from tkinter import ttk
import threading
root = tk.Tk()

tk.Label(root, text='EnterName:').grid(row = 0, column = 0)
eUsername = tk.Entry(root, width=15, borderwidth=2)
eUsername.grid(row = 0, column = 1)
tk.Label(root, text='Password:').grid(row = 1, column=0)
ePassword = tk.Entry(root, width=15, borderwidth=2, show='*')
ePassword.grid(row = 1, column = 1)
pb = ttk.Progressbar(orient='horizontal', length=15, mode='indeterminate')
pb.grid(row = 3, column = 1)

def run(Account, Password):
    print('begin')
    Chromdriver = 'chromedriver.exe'
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("start-maximized")
    chrome_options.add_argument("--disable-blink-features")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")


    #login info
    LOGIN_PAGE = "https://www.flightmemory.com/"
    ACCOUNT = Account
    PASSWORD = Password

    driver = webdriver.Chrome(executable_path=Chromdriver, chrome_options=chrome_options)


    #login
    driver.get(LOGIN_PAGE)
    wait = WebDriverWait(driver, 30)
    wait.until(EC.element_to_be_clickable((By.NAME, "username"))).send_keys(ACCOUNT)
    print('enteredName')
    wait.until(EC.element_to_be_clickable((By.NAME, "passwort"))).send_keys(PASSWORD)
    print('enteredPassword')
    wait.until(EC.element_to_be_clickable((By.XPATH, ".//input[@value='SignIn' and @type='submit']"))).click()
    print("Logged in")
    #go to flights

    wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'FLIGHTDATA')]"))).click()
    print('found flightdata')
    pages = []
    pages.append(driver.execute_script("return document.documentElement.outerHTML"))


    while True:
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//img[contains(@src,'/images/next.gif')]"))).click()
            pages.append(driver.execute_script("return document.documentElement.outerHTML"))
        except TimeoutException:
            print("Pages Found:")
            print(int(len(pages)))
            break


    #Create workbook object
    wb = openpyxl.Workbook()
    sheet = wb.get_sheet_names()[0]
    sheet = wb.get_sheet_by_name(sheet)
    sheet.title='Flights'

    #TableInfo
    sheet.cell(row = 1, column = 1).value='No.'
    sheet.cell(row = 1, column = 2).value='Date - Dep./Arr.'
    sheet.cell(row = 1, column = 3).value='Dep.'
    sheet.cell(row = 1, column = 4).value='Arr.'
    sheet.cell(row = 1, column = 5).value='Dis. [Km]'
    sheet.cell(row = 1, column = 6).value='Time'
    sheet.cell(row = 1, column = 7).value='Airline/FlightNumber'
    sheet.cell(row = 1, column = 8).value='Airplane'
    sheet.cell(row = 1, column = 9  ).value='Seat'



    global lastidx 
    lastidx = 0

    for page in pages:
        #find table with BS
        data = BeautifulSoup(page, 'html.parser')
        container = data.select_one('.container')
        tablebody = container.find_all('tbody')[2]
        trs = tablebody.find_all('tr', recursive=False)

        #loop through each line
        for idx, tr in enumerate(trs):
            idx += lastidx
            idx += 1
            
            if idx != lastidx + 1:
                #get info
                flight_number = tr.find_all('td')[0]
                date_dep_arr = tr.find_all('td')[1]
                Departure = tr.find_all('td')[2]
                Arrival = tr.find_all('td')[4]
                Distance = tr.find_all('td')[6]
                FlightTime =tr.find_all('td')[8]
                Airline_Flightinfo = tr.find_all('td')[10]
                Airplane = tr.find_all('td')[11]
                Seat = tr.find_all('td')[12]
                #addInfoToSheet
                sheet.cell(row = idx, column = 1).value=flight_number.get_text()
                sheet.cell(row = idx, column = 2).value=date_dep_arr.get_text()
                sheet.cell(row = idx, column = 3).value=Departure.get_text()
                sheet.cell(row = idx, column = 4).value=Arrival.get_text()
                sheet.cell(row = idx, column = 5).value=Distance.get_text()
                sheet.cell(row = idx, column = 6).value=FlightTime.get_text()
                sheet.cell(row = idx, column = 7).value=Airline_Flightinfo.get_text()
                sheet.cell(row = idx, column = 8).value=Airplane.get_text()
                sheet.cell(row = idx, column = 9).value=Seat.get_text()
            print(idx)    
        lastidx = idx

    wb.save('Flights.xlsx')
    pb.end()
    print('finished')
        
def OK():
    account = eUsername.get()
    password = ePassword.get()
    x = threading.Thread(target=run, args=(account, password))
    x.start()
    pb.start()

#GUI


tk.Button(root, text="OK",  command=lambda : OK()).grid(row=3, column=0)


root.mainloop()
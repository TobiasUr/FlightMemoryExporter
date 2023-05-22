from bs4 import BeautifulSoup
import pandas as pd
import openpyxl

#Create workbook object
wb = openpyxl.Workbook()
sheet = wb.create_sheet()
sheet.title='Flights'


#open txt file
filepath = "FlightMemory - FlightData.htm"
file = open(filepath, 'r')
data = file.read()
#print(file.read())

#remove first part'
#index = data.find('Options')
#data[index:]

data = BeautifulSoup(data, 'html.parser')
container = data.select_one('.container')
tablebody = container.find_all('tbody')[2]
trs = tablebody .find_all('tr', recursive=False)
global i
i = 1
for idx, tr in enumerate(trs):
    idx += 1
    sheet.cell(row = idx, column = 1).value=tr.get_text()
    print(idx)    
    #print(tr.get_text())
wb.save('Flights.xlsx')
    


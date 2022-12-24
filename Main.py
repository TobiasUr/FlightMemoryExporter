import pyautogui
import cv2
import pytesseract
import pandas as pd
from PIL import Image
import numpy as np
import xlwt

pytesseract.pytesseract.tesseract_cmd = r'C:\Users\Tobia\AppData\Local\Tesseract-OCR\tesseract.exe'

img= cv2.imread("TEST.png")
imgGray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)


wb = xlwt.Workbook()
ws = wb.add_sheet('Main')




def ROI(img, vertices):
    mask=np.zeros_like(img)
    cv2.fillPoly(mask, vertices, 255)
    masked = cv2.bitwise_and(img, mask)
    print('hello')
    return masked


# Get the size
w = img.shape[1]
h = img.shape[0]
resolution = np.array([w, h])

def readText(cropStart: float, cropEnd: float, outputName:str):
    vertices = np.array([[resolution[0]/cropStart, 0],[resolution[0]/cropEnd, 0], [resolution[0]/cropEnd, resolution[1]], [resolution[0]/cropStart, resolution[1]]], np.int32)
    imgCropped =ROI(imgGray, [vertices])

    flt = cv2.adaptiveThreshold(imgCropped,
                            100, cv2.ADAPTIVE_THRESH_MEAN_C,
                            cv2.THRESH_BINARY, 15, 16)


    config = ('-l eng — oem 1 — psm 3')

    text = pytesseract.image_to_string(flt, config=config)

    with open(outputName, 'w') as f:
        f.write(text)

    #cv2.imshow('window', flt)
    #cv2.waitKey()








readText(w, 35, '1.txt')
readText(35, float(9.46), '2.txt')
readText(9.46, 6.899, '3.txt')
readText(6.899, 3.63, '4.txt')
readText(3.63, 3.17, '5.txt')
readText(3.17, 2.04, '6.txt')
readText(2.04, 1.79, '7.txt')
readText(1.79, 1.45, '8.txt')
readText(1.45, 1.189, '9.txt')
readText(1.189, 1.08, '10.txt')


def excel(file, column):
    f = open(file, 'r+')

    data = f.readlines() # read all lines at once
    for i in range(len(data)):
        row = data[i].split()  # This will return a line of string data, you may need to convert to other formats depending on your use case
    for j in range(len(row)):
        ws.write(i, j, row[j])  # Write to cell i, j
    f.close()

excel('1.txt', 0)

wb.save('Excelfile' + '.xls')

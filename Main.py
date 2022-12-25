import pyautogui
import cv2
import pytesseract
import pandas as pd
from PIL import Image
import numpy as np
import xlwt
import csv
import openpyxl

pytesseract.pytesseract.tesseract_cmd = r'C:\Users\Tobia\AppData\Local\Tesseract-OCR\tesseract.exe'

img= cv2.imread("TEST.png")
imgGray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

wb = openpyxl.Workbook()
ws = wb.worksheets[0]




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
    #imgCropped =ROI(imgGray, [vertices])

    y1 = 0
    y2 = int(resolution[1])
    x1= int(resolution[0]/cropStart)
    x2= int(resolution[0]/cropEnd)
    imgCropped = imgGray[y1:y2, x1:x2]
    flt = cv2.adaptiveThreshold(imgCropped,
                            100, cv2.ADAPTIVE_THRESH_MEAN_C,
                           cv2.THRESH_BINARY, 15, 16,
                           )

    scale_percent = 150 # percent of original size
    width = int(imgCropped.shape[1] * scale_percent / 100)
    height = int(imgCropped.shape[0] * scale_percent / 100)
    dim = (width, height)
      
    flt = cv2.resize(flt, dim, interpolation = cv2.INTER_AREA)

    
    
    

    flt = imgCropped

    config = ('-l eng — oem 1 — -psm 6')

    text = pytesseract.image_to_string(flt, config=config)

    with open(outputName, 'w') as f:
        f.write(text)

    inputFile = outputName
    outputFile = outputName.replace(".txt", ".csv")

    with open(inputFile, 'r') as data:
        reader = csv.reader(data, delimiter='\t')
        for row in reader:
            ws.append(row)
            
    wb.save(outputFile)

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




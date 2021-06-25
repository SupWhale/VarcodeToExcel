from PIL import Image
from PIL import ImageTk
import keyboard
import tkinter as tk
import threading
import datetime
import os
import time
import cv2
import pyzbar.pyzbar as pyzbar
from openpyxl import Workbook
from datetime import datetime

def macadd(mac):
    
    result = ''
    
    if len(mac) == 12:
        
        list = [0,0,0,0,0,0,0,0,0,0,0]
        
        for i in range(0,11):
            if i%2 == 0:
                list[i] = mac[i:i+2]
            elif i%2 != 0:
                list[i] = ':'
        
        for j in list:
            result += j
    else:
        result = mac
    return result

def CheckType_Serial(code):
    try:   
        result = ""
        if code[0] == 84:
            result = "TMS"
        elif code[0] == 77:
            result = "PMC"
        elif code[0] == 71:
            result = "GateWay"
        elif code[0] == 82:
            result = "IRC"
        else:
            result = ""
    except IndexError as e:
        print(e)
    
    return result
        
def Doctor_processing(write_ws, decodedObjects,width,high): #의사협회 관리 모드
        
    for obj in decodedObjects:
        code = obj.data
        write_ws.cell(high,width,macadd(code.decode()))     
       
def Stock_Manage(write_ws, decodedObjects,width,high): #재고 관리 모드

        write_ws.cell(high,width,decodedObjects)
        write_ws.cell(high,width+1,'회사 7층')

        stri = datetime.today().strftime("%Y-%m-%d")
        
        write_ws.cell(high,width+2,stri)
        write_ws.cell(1,10,high)

def install_Int(write_ws, decodedObjects,width,high): #GEM 설치정보 양식에 맞춘 모드
    
    for obj in decodedObjects:
        code = obj.data
        str_high = str(high)
        location = 'A'+str_high
        write_ws[location] = obj.data
        merge_location = 'A'+str_high+':C'+str_high
        write_ws.merge_cells(merge_location)
    

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

def Decode_Serial(im): #이미지내에서 바코드를 찾아내고 해당 타입과 데이터를 출력하는 함수
    decodedObjects = pyzbar.decode(im)
    result = ""

    for obj in decodedObjects:
        print('Type : ', obj.type)
        print('Data : ', obj.data, '\n')
        result = obj.data

    return result

def Decode_Image(im): #이미지내에서 바코드를 찾아내고 해당 타입과 데이터를 출력하는 함수
    decodedObjects = pyzbar.decode(im)

    for obj in decodedObjects:
        print('Type : ', obj.type)
        print('Data : ', obj.data, '\n')

    return decodedObjects


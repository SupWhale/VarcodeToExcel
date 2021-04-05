from PIL import Image
from PIL import ImageTk
import keyboard
import tkinter as tk
import threading
import datetime
import os
import cv2
import pyzbar.pyzbar as pyzbar
import numpy as np
from openpyxl import Workbook


def decode(im): #이미지내에서 바코드를 찾아내고 해당 타입과 데이터를 출력하는 함수
    decodedObjects = pyzbar.decode(im)
    for obj in decodedObjects:
        print('Type : ', obj.type)
        print('Data : ', obj.data, '\n')
    return decodedObjects

def macadd(mac):
    list = [0,0,0,0,0,0,0,0,0,0,0]
    for i in range(0,11):
        if i%2 == 0:
            list[i] = mac[i:i+2]
        elif i%2 != 0:
            list[i] = ':'

    result = ''
    for j in list:
        result += j
    return result

def camThread():
    cap = cv2.VideoCapture(1) #외부 웹캠을 불러옴
    write_wb = Workbook() #워크시트를 생성할 객체 생성
    panel = None
    write_ws = write_wb.create_sheet('생성시트') #워크 시트 생성
    
    print('width :%d, height : %d' % (cap.get(3), cap.get(4))) #현재 카메라의 해상도 출력

    a = 10 #지정된 열 번호
    b = 4 #지정된 행 번호
    check_code = False
    while(True):
        ret, frame = cap.read()    # Read 결과와 frame
        if(ret) :
            
           #cv2.imshow('frame_color', frame) # 컬러 화면 출력
            image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            image = Image.fromarray(image)
            image = ImageTk.PhotoImage(image)

            if panel is None:
                panel = tk.Label(image=image)
                panel.image = image
                panel.pack(side="left")
            else:
                panel.configure(image=image)
                panel.image = image
            #if cv2.waitKey(1) == ord('q'): #q버튼 누를시 이미지 촬영
            if keyboard.is_pressed('q'):
                cv2.imwrite(os.path.join(os.path.abspath('jpg') , 'bacode.jpg'), frame, params=[cv2.IMWRITE_PNG_COMPRESSION,0]) #이미지파일 출력
                im = cv2.imread(os.path.join(os.path.abspath('jpg') , 'bacode.jpg'))
                    
                decodedObjects = decode(im) #바코드 해석
                write_ws = write_wb.active #엑셀 파일에 쓰기 활성화

                
                for obj in decodedObjects:
                    if obj: #바코드가 탐지가 되지 않았다면 false
                        check_code = True
                    else:
                        chekc_code = False

                    if a == 10 and check_code == True: #바코드가 탐지되었다면 데이터 작성 후  다음 열로 이동
                        code = obj.data
                        write_ws.cell(b,a,macadd(code.decode()))
                        a = a+1
                        print('물리 주소인식 완료')
                    elif a < 13 and check_code == True: #바코드가 탐지되었다면 데이터 작성 후  다음 열로 이동
                        write_ws.cell(b,a,obj.data)
                        a = a+1
                        print('인식 완료')
                    elif a == 13 and check_code == True: #바코드가 탐지되었다면 데이터 작성 후 줄 바꿈
                        write_ws.cell(b,a,obj.data)
                        a = 10
                        b = b+1
                        print('인식 완료')
                    elif check_code == False: #바코드가 탐지 되지 않았다면 헤당 셀에서 대기
                        a = a
                        b = b
                        print('인식 불가')
                    
                write_wb.save(os.path.join(os.path.abspath('xlsx') , 'test.xlsx')) #엑셀 파일 저장

            elif cv2.waitKey(1) == ord('w'): #w를 누룬다면 프로그램 종료
                break

                    
    
    cap.release() #프로그램 최종종료
    cv2.destroyAllWindows()






if __name__ == '__main__': #Main
    thread_img = threading.Thread(target=camThread, args=())
    thread_img.daemon = True
    thread_img.start()

    root = tk.Tk()
    root.mainloop()

    
    


    


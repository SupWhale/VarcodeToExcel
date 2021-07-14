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
from openpyxl import *
from datetime import datetime
from toExcel import *
from decode import *
from GEM import *
import sys

def Setmode_Doc(ch):
    global choice
    State.config(text = '의사협회모드')
    choice = ch + 1
    
def Setmode_Stock(ch):
    global choice
    State.config(text = '재고관리모드')
    choice = ch + 2
    
def Setmode_Install(ch):
    global choice
    State.config(text = '설치정보모드')
    choice = ch + 3

def save_filename(types):
    
    stri = ""
    
    if types == 2:
        stri = "재고 관리 파일"
    else:
        stri = datetime.today().strftime("%Y년%m월%d일%H시")
        
    return stri

def camThread():

    try:
            cap = cv2.VideoCapture(1) #외부 웹캠을 불러옴
            write_wb = Workbook() #워크시트를 생성할 객체 생성
            write_ws = write_wb.create_sheet('기본시트') #워크 시트 생성
            panel = None #GUI를 생성할 판넬 생성
            width2 = 2
            width = 2
            high = 3
            pmc_h=1
            irc_h=1
            tms_h=1
            gateway_h=1

                    
            print('width :%d, height : %d' % (cap.get(3), cap.get(4))) #현재 카메라의 해상도 출력

            
            
            check_code = False
            
            while(True):

                
                        
                    

                end_point = False
                ret, frame = cap.read()    # Read 결과와 frame
                if(ret) :
                    
                    image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                    image = Image.fromarray(image)
                    image = ImageTk.PhotoImage(image)

                    ch = 0

                    button1 = tk.Button(root, overrelief="solid", text = "의사협회 모드" ,width=15, command = lambda: Setmode_Doc(ch))
                    button1.place(x=1000, y=500)

                    button2 = tk.Button(root, overrelief="solid", text = "재고관리 모드" ,width=15, command = lambda: Setmode_Stock(ch))
                    button2.place(x=1000, y=400)

                    button3 = tk.Button(root, overrelief="solid", text = "설치정보 모드" ,width=15, command = lambda: Setmode_Install(ch))
                    button3.place(x=1000, y=300)

                    if panel is None:
                        panel = tk.Label(image=image)
                        panel.image = image
                        panel.pack(side="left")
    
                    else:
                        panel.configure(image=image)
                        panel.image = image
                    
                    if choice == 2:
                        write_wb = load_workbook("재고 관리 파일.xlsx")
                        pmc_ws = write_wb['PMC']
                        gateway_ws = write_wb['Gateway']
                        tms_ws = write_wb['TMS']
                        irc_ws = write_wb['IRC']

                        width = 3
                        high = 3
                        pmc_h=pmc_ws.cell(2,20).value
                        irc_h=irc_ws.cell(2,20).value
                        tms_h= tms_ws.cell(2,20).value
                        gateway_h=gateway_ws.cell(2,20).value
                    else:
                        write_wb2 = load_workbook("GEM 데이터베이스.xlsx") 
                        write_ws2 = write_wb2['Sheet']
                        db_h = write_ws2.cell(2,20).value
                        width2 = 2




                    if keyboard.is_pressed('q'):
                        time.sleep(0.1)
                        cv2.imwrite('bacode.jpg', frame) #이미지파일 출력
                        im = cv2.imread('bacode.jpg')
                            
                        decodedObjects = decode(im)#바코드 해석 
                        result.config(text = decode_serial(im))
                        
                        if decodedObjects == []:
                            result.config(text = "감지된 바코드가 없습니다.")
                            continue

                        else:
                        
                            if choice == 1:
                                
                                DB_sell(write_ws2, decodedObjects, width2, db_h)

                                if width2 < 3:
                                    width2 = width2+1
                                elif width2 == 3:
                                    width2 = 2
                                    db_h = db_h+1

                                Doctor_processing(write_ws, decodedObjects,width,high)
                                
                                if width < 7:
                                    width = width+1
                                elif width == 7:
                                    width = 3
                                    high = high+1
                                                                    
                            elif choice == 2:
    
                                decodedObjects_s = decode_serial(im)
                                print(decodedObjects_s)

                                types = CheckType_Serial(decodedObjects_s)
                                
                                if types == "TMS":
                                    
                                    Stock_Manage(tms_ws, decodedObjects_s,width,tms_h)
                                    tms_h = tms_h+1
                                    
                                elif types == "PMC":
                                    
                                    Stock_Manage(pmc_ws, decodedObjects_s,width,pmc_h)
                                    pmc_h = pmc_h+ 1
                                    
                                elif types == "IRC":

                                    Stock_Manage(irc_ws, decodedObjects_s,width,irc_h)
                                    irc_h = irc_h+ 1
                                    
                                elif types == "GateWay":
                                    
                                    Stock_Manage(gateway_ws, decodedObjects_s,width,gateway_h)
                                    gateway_h = gateway_h+1
                                    
                                high = high+1
                                
                            elif choice == 3:
                                DB_sell(write_ws2, decodedObjects, width2, db_h)

                                if width2 < 3:
                                    width2 = width2+1
                                elif width2 == 3:
                                    width2 = 2
                                    db_h = db_h+1


                                install_Int(write_ws, decodedObjects,high)
                                high = high+1
                            
                            
                            stri = save_filename(choice)
                            
                            write_wb.save(stri + '.xlsx') #엑셀 파일 저장
                            write_wb2.save("GEM 데이터베이스.xlsx")
                        
            cap.release() #프로그램 최종종료
            cv2.destroyAllWindows()

    except PermissionError as e:
        print(e)

        result.config(text = "현재 활성화 된 엑셀 파일을 닫아주세요! 잠시후 프로그램이 재시작됩니다.")
        ErrorMessage.config(text = e)

        time.sleep(15)
        os.execl(sys.executable, sys.executable, *sys.argv)
        

if __name__ == '__main__': #Main
    try:
            thread_img = threading.Thread(target=camThread, args=())
            thread_img.daemon = True
            thread_img.start()
            
            root = tk.Tk()
            root.title("GEM 자동화 엑셀 작성 툴")
            root.geometry("1200x600+100+100")

            State = tk.Label(root, text = '모드를 입력하십시오', font = 'TkFixedFont')
            State.pack()

            result = tk.Label(root, text = 'Result', font = 'TkFixedFont')
            result.pack()

            ErrorMessage = tk.Label(root, text = 'Result', font = 'TkFixedFont')
            ErrorMessage.pack()
            
           

            root.mainloop()
    except NameError as e:
        print(e)

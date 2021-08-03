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
    ch = 0
    button1 = tk.Button(root, overrelief="solid", text = "의사협회 모드" ,width=15, command = lambda: Setmode_Doc(ch))
    button1.place(x=1000, y=500)

    button2 = tk.Button(root, overrelief="solid", text = "재고관리 모드" ,width=15, command = lambda: Setmode_Stock(ch))
    button2.place(x=1000, y=400)

    button3 = tk.Button(root, overrelief="solid", text = "설치정보 모드" ,width=15, command = lambda: Setmode_Install(ch))
    button3.place(x=1000, y=300)
    try:
            cap = cv2.VideoCapture(1) #외부 웹캠을 불러옴
            write_wb = Workbook() #워크시트를 생성할 객체 생성
            write_ws = write_wb.create_sheet('기본시트') #워크 시트 생성
            panel = None #GUI를 생성할 판넬 생성
            lable_WH = [2,2,3,1,1,1,1] #width2(0), width(1), high(2), pmc_h(3), irc_h(4), tms_h(5), gateway_h(6)
                    
            print('width :%d, height : %d' % (cap.get(3), cap.get(4))) #현재 카메라의 해상도 출력

            check_code = False
            
            while(True):
                end_point = False
                ret, frame = cap.read()    # Read 결과와 frame
                if(ret) :
                    
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
                    
                    if choice == 2:
                            write_wb = load_workbook("재고 관리 파일.xlsx")
                            pmc_ws = write_wb['PMC']
                            gateway_ws = write_wb['Gateway']
                            tms_ws = write_wb['TMS']
                            irc_ws = write_wb['IRC']

                            lable_WH[1] = 3
                            lable_WH[2] = 3
                            lable_WH[3] = pmc_ws.cell(3,20).value
                            lable_WH[4] = irc_ws.cell(3,20).value
                            lable_WH[5] = tms_ws.cell(3,20).value
                            lable_WH[6] = gateway_ws.cell(3,20).value
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
                                
                                DB_sell(write_ws2, decodedObjects, lable_WH[0], db_h)

                                if lable_WH[0] < 3:
                                    lable_WH[0] = lable_WH[0]+1
                                elif lable_WH[0] == 3:
                                    lable_WH[0] = 2
                                    db_h = db_h+1

                                Doctor_processing(write_ws, decodedObjects,lable_WH[1],lable_WH[2])
                                
                                if lable_WH[1] < 7:
                                    lable_WH[1] = lable_WH[1]+1
                                elif lable_WH[1] == 7:
                                    lable_WH[1] = 3
                                    lable_WH[2] = lable_WH[2]+1
                                                                    
                            elif choice == 2:
    
                                decodedObjects_s = decode_serial(im)
                                print(decodedObjects_s)

                                types = CheckType_Serial(decodedObjects_s)
                                
                                if types == "TMS":
                                    
                                    Stock_Manage(tms_ws, decodedObjects,lable_WH[1],lable_WH[5])
                                    lable_WH[5] = lable_WH[5]+1
                                    
                                elif types == "PMC":
                                    
                                    Stock_Manage(pmc_ws, decodedObjects,lable_WH[1],lable_WH[3])
                                    lable_WH[3] = lable_WH[3]+ 1
                                    
                                elif types == "IRC":

                                    Stock_Manage(irc_ws, decodedObjects,lable_WH[1],lable_WH[4])
                                    lable_WH[4] = lable_WH[4]+ 1
                                    
                                elif types == "GateWay":
                                    
                                    Stock_Manage(gateway_ws, decodedObjects,lable_WH[1],lable_WH[6])
                                    lable_WH[6] = lable_WH[6]+1
                                    
                                lable_WH[2] = lable_WH[2]+1
                                
                            elif choice == 3:
                                DB_sell(write_ws2, decodedObjects, lable_WH[0], db_h)

                                if lable_WH[0] < 3:
                                    lable_WH[0] = lable_WH[0]+1
                                elif lable_WH[0] == 3:
                                    lable_WH[0] = 2
                                    db_h = db_h+1


                                install_Int(write_ws, decodedObjects,lable_WH[2])
                                lable_WH[2] = lable_WH[2]+1
                            
                            
                            stri = save_filename(choice)
                            
                            write_wb.save(stri + '.xlsx') #엑셀 파일 저장
                            ##write_wb2.save("GEM 데이터베이스.xlsx")
                        
            cap.release() #프로그램 최종종료
            cv2.destroyAllWindows()
    except NameError as e:
        print(e)

        result.config(text = "잠시후 프로그램이 재시작됩니다.")
        ErrorMessage.config(text = e)

        time.sleep(15)
        os.execl(sys.executable, sys.executable, *sys.argv)

    except PermissionError as e:
        print(e)

        result.config(text = "현재 활성화 된 엑셀 파일을 닫아주세요! 잠시후 프로그램이 재시작됩니다.")
        ErrorMessage.config(text = e)

        time.sleep(15)
        os.execl(sys.executable, sys.executable, *sys.argv)
        

if __name__ == '__main__': #Main
    try:
            
            root = tk.Tk()
            root.title("GEM 자동화 엑셀 작성 툴")
            root.geometry("1200x600+100+100")

            State = tk.Label(root, text = '모드를 입력하십시오', font = 'TkFixedFont')
            State.pack()

            result = tk.Label(root, text = 'Result', font = 'TkFixedFont')
            result.pack()

            ErrorMessage = tk.Label(root, text = 'Result', font = 'TkFixedFont')
            ErrorMessage.pack()
            
            thread_img = threading.Thread(target=camThread, args=())
            thread_img.daemon = True
            thread_img.start()
            

            root.mainloop()
    except NameError as e:
        print(e)

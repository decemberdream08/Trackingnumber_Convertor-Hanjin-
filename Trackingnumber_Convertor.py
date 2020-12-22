#import datetime, time, os, re, win32com.client, shutil, telepot
import datetime, win32com.client
#import pyautogui
import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic

###################################################################
#   Define List, Variable
###################################################################
#form_class = uic.loadUiType("D:\\03_Study\\01_Python\\01_Code\\02_Auto\\Trackingnumber_Convertor.ui")[0]
form_class = uic.loadUiType("Trackingnumber_Convertor.ui")[0]
app = 0
item_number_each_person = []

###################################################################
#   각 열의 위치를 지정
#   문자에 대한 ASCII 코드 변환 후 64를 빼서 A부터 1로 카운트
###################################################################
SOURCE_DELIVERY_RECEIVER = ord('K') - 64 # 수취인명(K열)
SOURCE_DELIVERY_ITEM = ord('Q') - 64 # 상품명(Q열)
SOURCE_DELIVERY_OPTION = ord('S') - 64 # 옵션정보(S열)
SOURCE_DELIVERY_NUMBER = ord('U') - 64 # 수량(U열)
SOURCE_DELIVERY_TEL_NO1 = ord('O') - 64 + 26 # 수취인 연락처1(AO)
SOURCE_DELIVERY_TEL_NO2 = ord('P') - 64 + 26 # 수취친 연락처2(AP)
SOURCE_DELIVERY_RECEIVER_ADDRESS = ord('Q') - 64 + 26 # 배송지(AQ)
SOURCE_DELIVERY_POST_NO = ord('S') - 64 + 26 # 우편번호(AS)
SOURCE_DELIVERY_MESSAGE = ord('T') - 64 + 26 # 배송메세지(AT)

TARGET_DELIVERY_RECEIVER = ord('A') - 64 # 수취인명
TARGET_DELIVERY_ITEM = ord('R') - 64 # 상품명1
TARGET_DELIVERY_NUMBER = ord('S') - 64 # 수량1
TARGET_DELIVERY_TEL_NO1 = ord('B') - 64 # 수취인 연락처1
TARGET_DELIVERY_TEL_NO2 = ord('D') - 64 # 수취친 연락처2
TARGET_DELIVERY_RECEIVER_ADDRESS = ord('F') - 64 # 배송지
TARGET_DELIVERY_POST_NO = ord('E') - 64 # 우편번호
TARGET_DELIVERY_MESSAGE = ord('Q') - 64 # 배송메세지

TARGET_DELIVERY_BOX_NUMBER = ord('G') - 64 # 배송 Box 수

###################################################################
#   Insection File Name Only
###################################################################
def Insection_Filename(full_path):
    x = full_path.split('/')
    x.reverse()
    ret_filename = x[0]
    return ret_filename

###################################################################
#   Working Directory
###################################################################
def Working_code(source, target):
    ### Excel File 정보 ###
    Source_Excel_PATH = source    
    Target_Excel_PATH = target
    
    SaveAS_PATH = target.replace('한진택배_Template.xlsx', '')
    SaveAS_PATH = SaveAS_PATH.replace('/','\\')    
    SaveAS_File_Name = '한진택배'
    File_extension = '.xlsx'
    date = str(datetime.date.today())
    SaveAS_Excel_PATH = SaveAS_PATH + SaveAS_File_Name + '_' + date + File_extension
    
    excel = win32com.client.Dispatch('Excel.Application')
    wb = excel.Workbooks.Open(Source_Excel_PATH)
    ws = wb.Worksheets('발주발송관리')

    wb1 = excel.Workbooks.Open(Target_Excel_PATH)
    ws1 = wb1.Worksheets('Sheet1')
    
    ### Looking for number of item from Excel file ###
    row = 3

    while True:
        cell_value = ws.Cells(row, 1).Value

        if cell_value == None:
            item_number = row - 3
            last_row = row - 1
            break

        row += 1
    
    # 한 수취인이 몇개의 상품인지 확인
    row = 3
    prev_cell_value = 0
    count = 0

    while True:
        cell_value = ws.Cells(row, SOURCE_DELIVERY_RECEIVER).Value

        if cell_value != prev_cell_value:
            prev_cell_value = cell_value

            if count != 0:
                item_number_each_person.append(count) ### 수취인명이 다른때 마다 총 갯수를 저장

            count = 1
        else:
            count += 1

        if row == last_row + 1: ## 마지막 행+1(마지막행 데이터를 저장하기 위함)까지 검색이 끝나면 종료)
            break

        row += 1

    ### Data copy from Source to Target

    target_data_last_row = len(item_number_each_person) # Target 엑셀 파일의 데이터 행 숫자는 리스트 크기와 동일

    total_count = 0

    for i in range(target_data_last_row):
        
        for j in range(item_number_each_person[i]):
            
            if j == 0: # 수취인이 동일하기 때문에 하기 정보들은 1회만 저장
                ws1.Cells(i+4, TARGET_DELIVERY_RECEIVER).Value = ws.Cells(total_count+3, SOURCE_DELIVERY_RECEIVER).Value # 수취인명 : Target 데이터는 4행 부터 시작 / Source 데이터는 3행 부터 시작
                ws1.Cells(i+4, TARGET_DELIVERY_TEL_NO1).Value = ws.Cells(total_count+3, SOURCE_DELIVERY_TEL_NO1).Value # 수취인 연락처1 : Target 데이터는 4행 부터 시작 / Source 데이터는 3행 부터 시작
                ws1.Cells(i+4, TARGET_DELIVERY_TEL_NO2).Value = ws.Cells(total_count+3, SOURCE_DELIVERY_TEL_NO2).Value # 수취친 연락처2 : Target 데이터는 4행 부터 시작 / Source 데이터는 3행 부터 시작
                ws1.Cells(i+4, TARGET_DELIVERY_RECEIVER_ADDRESS).Value = ws.Cells(total_count+3, SOURCE_DELIVERY_RECEIVER_ADDRESS).Value # 배송지 : Target 데이터는 4행 부터 시작 / Source 데이터는 3행 부터 시작
                ws1.Cells(i+4, TARGET_DELIVERY_POST_NO).Value = ws.Cells(total_count+3, SOURCE_DELIVERY_POST_NO).Value # 우편번호 : Target 데이터는 4행 부터 시작 / Source 데이터는 3행 부터 시작
                ws1.Cells(i+4, TARGET_DELIVERY_MESSAGE).Value = ws.Cells(total_count+3, SOURCE_DELIVERY_MESSAGE).Value # 배송메세지 : Target 데이터는 4행 부터 시작 / Source 데이터는 3행 부터 시작
                
                ws1.Cells(i+4, TARGET_DELIVERY_BOX_NUMBER).Value = 1 # 배송 박스 개수는 1개로 지정
                
                if ws.Cells(total_count+3, SOURCE_DELIVERY_OPTION).Value == None:            
                    ws1.Cells(i+4, TARGET_DELIVERY_ITEM).Value = ws.Cells(total_count+3, SOURCE_DELIVERY_ITEM).Value # 첫번째 상품명 : Target 데이터는 4행 부터 시작 / Source 데이터는 3행 부터 시작
                else:
                    ws1.Cells(i+4, TARGET_DELIVERY_ITEM).Value = ws.Cells(total_count+3, SOURCE_DELIVERY_ITEM).Value + '/' + ws.Cells(total_count+3, SOURCE_DELIVERY_OPTION).Value # 첫번째 상품명

                ws1.Cells(i+4, TARGET_DELIVERY_NUMBER).Value = ws.Cells(total_count+3, SOURCE_DELIVERY_NUMBER).Value
            else:
                # 상품명만 추가로 기입
                if ws.Cells(total_count+3, SOURCE_DELIVERY_OPTION).Value == None: 
                    ws1.Cells(i+4, TARGET_DELIVERY_ITEM + 2*j).Value = ws.Cells(total_count+3, SOURCE_DELIVERY_ITEM).Value # n번째 상품명 : Target 데이터는 4행 부터 시작 / Source 데이터는 3행 부터 시작
                else:
                    ws1.Cells(i+4, TARGET_DELIVERY_ITEM + 2*j).Value = ws.Cells(total_count+3, SOURCE_DELIVERY_ITEM).Value + '/' + ws.Cells(total_count+3, SOURCE_DELIVERY_OPTION).Value # n번째 상품명
                    
                ws1.Cells(i+4, TARGET_DELIVERY_NUMBER + 2*j).Value = ws.Cells(total_count+3, SOURCE_DELIVERY_NUMBER).Value # n번째 상품명의 수량
            
            #print("Total_count = ", total_count)  
            total_count += 1


    ### 엑셀 파일을 저장 후 종료
    wb.Close()
    wb1.SaveAs(SaveAS_Excel_PATH) # 편집한 내용은 다른이름으로 저장 '한진택배_오늘날짜.xlsx"
    wb1.Close(Target_Excel_PATH)
    #excel.Quit()


class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        
        self.Source_File_Path = 0
        self.Target_File_Path = 0
        self.Source_input = False
        self.Target_input = False
        
        self.pushButton.clicked.connect(self.pushButtonClicked)
        self.pushButton_2.clicked.connect(self.pushButtonClicked2)
        self.pushButton_3.clicked.connect(self.pushButtonClicked3)
        self.pushButton_4.clicked.connect(app.quit)

    def pushButtonClicked(self):
        fname = QFileDialog.getOpenFileName(self)
        self.Source_File_Path = fname[0] # 파일 이름을 포함한 전체 경로
        Source_Filename = Insection_Filename(self.Source_File_Path) # 파일 이름
        
        if '스마트스토어' in Source_Filename:
            self.textEdit.setText(Source_Filename)
            self.label.setText('스마트스토어 파일 선택 완료')
            self.Source_input = True
        else:
            #self.textEdit.clearText()
            self.label.setText('스마트스토어 파일을 선택해주세요 !!!')
            self.Source_input = False

    def pushButtonClicked2(self):
        fname = QFileDialog.getOpenFileName(self)
        self.Target_File_Path = fname[0] # 파일 이름을 포함한 전체 경로
        Target_Filename = Insection_Filename(self.Target_File_Path) # 파일 이름
        
        if '한진택배_Template' in Target_Filename:
            self.textEdit_2.setText(Target_Filename)
            self.label.setText('한진택배_Template 파일 선택 완료')
            self.Target_input = True
        else:
            #self.textEdit_2.clearText()
            self.label.setText('한진택배_Template 파일을 선택해주세요 !!!')
            self.Target_input = False

    def pushButtonClicked3(self):
        if (self.Source_input == True) and (self.Target_input == True):            
            self.label.setText('파일 변환 진행 중 ...')
            Working_code(self.Source_File_Path, self.Target_File_Path)
            self.label.setText('완료 ! 파일은 한진택배_Template 파일 폴더에 저장됨 !! 종료 버튼을 누르세요 !!')
            
            self.Source_input = False
            self.Target_input = False
        else:
            if self.Source_input == True:
                self.label.setText('스마트스토어 파일을 선택해주세요 !!')
            elif self.Target_input == True:                
                self.label.setText('한진택배_Template 파일을 선택해주세요 !!')
            else:
                self.label.setText('스마트스토어 파일과 한진택배_Template 파일을 선택해주세요 !!')
        

if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
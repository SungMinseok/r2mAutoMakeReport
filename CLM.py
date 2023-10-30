


import os
import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5 import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import*
import pandas as pd
from PyQt5 import QtCore, QtWidgets
import shutil
from datetime import datetime
import time
from creation3 import *
import traceback
form_class = uic.loadUiType(f'./CLM_UI.ui')[0]
FROM_CLASS_Loading = uic.loadUiType("load.ui")[0]

class WindowClass(QMainWindow, form_class) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)

        self.setWindowTitle("CheckListMaker 0.1")
        self.statusLabel = QLabel(self.statusbar)

        self.setGeometry(1470,28,400,400)
        self.setFixedSize(450,350)
        

        '''기본값입력■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■'''
        self.make_ref_info_dict()
        #소스가 있는 폴더(r2m 쉐어 포인트)        
        self.btn_sourcePath_1.clicked.connect(lambda : self.파일열기(self.input_sourcePath.text()))
        self.btn_sheetName.clicked.connect(self.load_sheetnames)
        self.combo_sheetName.currentTextChanged.connect(self.load_enablecolnames)
        self.btn_execute.clicked.connect(self.execute)

        self.load_sheetnames()

        # self.input_departure_dir_path.setText(fr"c:\Users\mssung\OneDrive - Webzen Inc\라이브 서비스(국내)\KR R2M\2023 3분기\230831 업데이트")
        # #self.input_departure_dir_path.setText(fr"C:\Users\mssung\OneDrive - Webzen Inc\라이브 서비스(대만)\TW R2M 2023년\2023년 3분기\230822 패치")
        # #최종 목적지 폴더(팀 쉐어 포인트)
        # self.input_destination_dir_path.setText(fr"C:\Users\mssung\OneDrive - Webzen Inc\R2M\QA\2023년")
        # self.input_resultdir.setText(fr'D:\파이썬결과물저장소\ARM\{datetime.now().strftime("%y%m%d")}')
        # '''메뉴탭■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■'''
        # self.menu_howtouse.triggered.connect(lambda : self.파일열기("사용법_ARM.txt"))
        # self.menu_patchnote.triggered.connect(lambda : self.파일열기("패치노트_ARM.txt"))

        # '''폴더세팅부'''
        # #self.btn_makedir.clicked.connect(self.set_folder)
        
        
        # self.btn_open_0.clicked.connect(self.open_qa_file)
        # self.btn_make_2.clicked.connect(self.create_qa_result)
        # '''빌드파일명'''
        # self.btn_getbuild_0.clicked.connect(lambda : self.get_filename('Korea'))
        # self.btn_getbuild_1.clicked.connect(lambda : self.get_filename('Taiwan'))
    
    def load_sheetnames(self):
        #print('353')
        xlsx_filename = self.input_sourcePath.text()

        #엑셀파일 내 시트명을 불러와서 lines로 저장하는 코드 추가
        global xls
        # xls = pd.ExcelFile(xlsx_filename)
        # sheet_names = xls.sheet_names


        # self.combo_sheetName.clear()
        # for sheet_name in sheet_names:
        #     #print(sheet_name)
        #     self.combo_sheetName.addItem(sheet_name)


        xls = pd.read_excel(xlsx_filename, sheet_name=None)
        sheet_names = xls.keys() if isinstance(xls, dict) else xls.sheet_names

        self.combo_sheetName.clear()
        for sheet_name in sheet_names:
            if sheet_name in ref_info_dict:
                self.combo_sheetName.addItem(sheet_name)


        self.load_enablecolnames()
        self.apply_colname(self.combo_sheetName.currentText())
    def load_enablecolnames(self):
        
        xlsx_filename = self.input_sourcePath.text()
        sheet_name = self.combo_sheetName.currentText()

        # xls의 sheet_name의 열 이름을 불러와서 입력하는 코드 추가
        try:    
            df = xls[sheet_name] if isinstance(xls, dict) else pd.read_excel(xlsx_filename, sheet_name=sheet_name)
        except: 
            return
            
        # Filter out "Unnamed" columns
        valid_col_names = [col for col in df.columns if not col.startswith('Unnamed')]
    
        col_names = ', '.join(valid_col_names)
        
        self.input_enableColName.clear()
        self.input_enableColName.insertPlainText(col_names)

        self.apply_colname(sheet_name)

    def make_ref_info_dict(self):
        global ref_info_dict
        ref_info_dict = {}
        with open("ref_info.txt", "r", encoding='utf-8') as file:
            for line in file:
                parts = line.strip().split(',')
                sheet_name = parts[0]
                col_names = [col.strip() for col in parts[1:]]
                ref_info_dict[sheet_name] = col_names

    def apply_colname(self,cur_sheet_name):

        #cur_sheet_name이 ref_info_dic의 키와 일치하는게 있으면, 아래의 코드를 실행함
        if cur_sheet_name in ref_info_dict:
            print(cur_sheet_name)
            col_names = ref_info_dict[cur_sheet_name]

            self.input_mainColName.setText(col_names[0])
            self.input_targetColName.setText(','.join(col_names[1:]))
        # else:
        #     # Handle the case where cur_sheet_name is not found in ref_info_dic
        #     pass
        #일치하는게 없으면 아무것도안함.

    def execute(self):
        try:
            input_file = self.input_sourcePath.text()
            sheet_name = self.combo_sheetName.currentText()
            output_file = os.path.join(self.input_resultPath.text(),f"{sheet_name}_{time.strftime('%y%m%d_%H%M%S')}.xlsx")
            criterion = self.input_mainColName.text()
            #text = "퀘스트 목표,퀘스트 내용,경험치,보상1,보상2"
            required_parts = self.input_targetColName.text().split(',')
            
            self.loading = loading(self)
            self.worker_thread = WorkerThread(myWindow,sheet_name,input_file, output_file, criterion, required_parts)
            self.worker_thread.finished.connect(self.cleanup)
            self.worker_thread.start()
            
            #create_checklist(sheet_name,input_file, output_file, criterion, required_parts)

        except Exception as e:
            
            self.popUp(desText= traceback.format_exc())
            print(f'생성실패 : {e}')

    def cleanup(self):
        self.worker_thread = None

    def start_loading(self,qma):
        loading(self)

    def make_process(self,a,b,c,d,e):
        create_checklist(a,b,c,d,e)

        self.worker_thread.finished.emit()
        self.worker_thread.quit()
        self.loading
        self.loading.deleteLater()

        if self.check_0.isChecked():
            os.startfile(c)

    def 파일열기(self,filePath):
        try:
            os.startfile(filePath)
        except : 
            print("파일 없음 : "+filePath)    

    def get_latest_file_in_directory(self, source_path, target_file):

        def find_latest_file(folder):
            latest_file = None
            latest_time = datetime.min

            for root, dirs, files in os.walk(folder):
                if target_file in files:
                    file_path = os.path.join(root, target_file)
                    modified_time = datetime.fromtimestamp(os.path.getmtime(file_path))
                    if modified_time > latest_time:
                        latest_file = file_path
                        latest_time = modified_time

            return latest_file

        latest_file_path = find_latest_file(source_path)
        return latest_file_path
        # if latest_file_path:
        #     # 파일 실행 코드 작성
        #     print(f"가장 최신의 파일 실행: {latest_file_path}")
        #     os.startfile(os.path.normpath(latest_file_path))

        # else:
        #     print(f"'{target_file}' 파일을 찾을 수 없습니다.")


    def find_folders_by_name(self, source_path, folder_name):
        #matching_folders = []
        
        for root, dirs, files in os.walk(source_path):
            if folder_name in dirs:
                folder_path = os.path.join(root, folder_name)
                #matching_folders.append(folder_path)
                return folder_path
            
        print(f"'{folder_name}' 이름을 가진 폴더를 찾을 수 없습니다.")
        

    source_path = fr'C:\Users\mssung\OneDrive - Webzen Inc\R2M_Build\KR'
    #folder_name = 'YourFolderName'  # 검색할 폴더 이름

    # matching_folders = find_folders_by_name(source_path, folder_name)

    # if matching_folders:
    #     for folder in matching_folders:
    #         print(f"폴더 경로: {folder}")
    # else:
    #     print(f"'{folder_name}' 이름을 가진 폴더를 찾을 수 없습니다.")


    def print_log(self, log): # / - \ / - \ / ㅡ ㄷ
        self.progressLabel.setText(log)
        QApplication.processEvents()
        
    def popUp(self,desText,titleText="error"):
        #if type == "about" :
        msg = QtWidgets.QMessageBox()  
        msg.setGeometry(1520,28,400,2000)
        msg.setText(desText)

        #msg.setTextInteractionFlags(QtCore.Qt.TextInteractionFlag.TextSelectableByMouse)
        msg.setTextInteractionFlags(QtCore.Qt.TextInteractionFlag.TextSelectableByMouse)

        # if type == "report" :
        #     msg.setFixedSize(500,500)

            
        # if type == "searchItem" :
        #     msg.setStandardButtons(QtWidgets.QMessageBox.Ok | QtWidgets.QMessageBox.Cancel)
        #     #msg.buttonClicked.connect(self.messageBoxButton,desText)
        # #msg.setIcon(QtWidgets.QMessageBox.Information)
        # msg.setWindowTitle(titleText)

        # if type == "searchItem" :
        #     msg.setText("ItemID : " + str(desText) + "\n('OK'를 눌러 바로 생성)")      
        # else:
        #     msg.setText(desText)

        x = msg.exec_()
        

class loading(QWidget,FROM_CLASS_Loading):
    
    def __init__(self,parent):
        super(loading, self).__init__(parent)    
        self.setupUi(self) 
        #self.resize(parent.size())
        self.setFixedSize(parent.size())
        self.center()
        # Get the size of the parent widget and set the loading widget to the same size
        
        self.show()
        
        self.movie = QMovie('lcu_ui_ready_check.gif', QByteArray(), self)
        self.movie.setCacheMode(QMovie.CacheAll)
        self.label.setMovie(self.movie)
        self.label.setScaledContents(True)
        #self.movie.set(500,500)
        self.movie.start()
        self.setWindowFlags(Qt.FramelessWindowHint)
    # 위젯 정중앙 위치
    def center(self):
        
        size=self.size()
        ph = self.parent().geometry().height()
        pw = self.parent().geometry().width()
        self.move(int(pw/2 - size.width()/2), int(ph/2 - size.height()/2))
        self.move(int(pw/2 - size.width()/2), int(ph/2 - size.height()/2))
        
class WorkerThread(QThread):
    finished = pyqtSignal()

    def __init__(self,window, sheet_name,input_file, output_file, criterion, required_parts):
        super().__init__()
        
        self.window = window
        self.a = sheet_name
        self.b = input_file
        self.c = output_file
        self.d = criterion
        self.e = required_parts

    def run(self):
        #create_checklist(sheet_name,input_file, output_file, criterion, required_parts)
        #create_checklist(self.a,self.b, self.c, self.d, self.e)
        self.window.make_process(self.a, self.b, self.c, self.d, self.e)
        self.finished.emit()

if __name__ == "__main__" :
    #QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv) 

    #WindowClass의 인스턴스 생성
    myWindow = WindowClass() 

    #프로그램 화면을 보여주는 코드
    myWindow.show()

    #프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec_()
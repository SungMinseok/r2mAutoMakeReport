


import os
import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtCore import Qt, QDate
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtGui import QKeySequence,QPixmap, QColor
from PyQt5.QtWidgets import QLabel, QApplication, QWidget, QVBoxLayout
import pandas as pd
import shutil
from CL파일읽고정보저장 import *
from 정보대로워드문서작성 import *

form_class = uic.loadUiType(f'./AutoReport_UI.ui')[0]


class WindowClass(QMainWindow, form_class) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)

        self.setWindowTitle("AutoReportMaker 0.1")
        self.statusLabel = QLabel(self.statusbar)

        self.setGeometry(1470,28,400,400)
        self.setFixedSize(400,400)
        
        self.dateedit_project.setDate(QDate.currentDate())

        '''기본값입력■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■'''
        #소스가 있는 폴더(r2m 쉐어 포인트)
        self.input_departure_dir_path.setText(fr"c:\Users\mssung\OneDrive - Webzen Inc\라이브 서비스(국내)\KR R2M\2023 3분기\230817 패치")
        #최종 목적지 폴더(팀 쉐어 포인트)
        self.input_destination_dir_path.setText(fr"C:\Users\mssung\OneDrive - Webzen Inc\R2M\QA\2023년")
        
        '''메뉴탭■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■'''
        #self.menu_patchnote.triggered.connect(lambda : self.파일열기("패치노트.txt"))
        #self.combox_country.currentTextChanged.connect(self.set_data_path)


        '''폴더세팅부'''
        self.btn_makedir.clicked.connect(self.set_folder)


    def set_folder(self):
        self.makedir()
        self.복붙문서적용()

        country = self.combox_country.currentText()
        project_name = self.combox_projectname.currentText()
        date = pd.to_datetime(self.dateedit_project.text())
        patch_type = self.combox_patchtype.currentText()    
        doc_type = self.combox_doctype.currentText()        


        source_directory = self.input_departure_dir_path.text()
        #destination_path = self.input_destination_dir_path.text()
        destination_path = f'./{doc_type}/{datetime.now().strftime("%y%m%d_%H%M%S")}/{date.strftime("%y%m%d")} {patch_type}'
        if not os.path.isdir(destination_path):                                                           
            os.mkdir(destination_path)
        #source_directory = os.path.join(departure_path,f'{date.strftime("%y%m%d")} {patch_type}')
        #share_point_path = os.path.join(destination_path,f'{date.strftime("%y%m%d")} {patch_type}')
        self.폴더내용물전체복사(source_directory,destination_path)

        checklist_file_name = fr'CL_{country}_{project_name}_{date.strftime("%y%m%d")} {patch_type} QA.xlsx'
        checklist_file_path = os.path.join(destination_path, checklist_file_name)
        save_qa_info_to_csv(checklist_file_path)

        result_file_name = fr'QA_결과 보고_{country}_{project_name}_{date.strftime("%y%m%d")} {patch_type} QA 결과.docx'
        result_file_path = os.path.join(destination_path, result_file_name)
        update_qa_report(result_file_path)


    def makedir(self):
        '''
        폴더생성
        '''
        country = self.combox_country.currentText()
        project_name = self.combox_projectname.currentText()
        date = pd.to_datetime(self.dateedit_project.text())
        patch_type = self.combox_patchtype.currentText()        
        doc_type = self.combox_doctype.currentText()        

        #C:\Users\mssung\OneDrive - Webzen Inc\R2M\QA\2023년        

        dir_path = self.input_destination_dir_path.text()
        total_path = f'{dir_path}\{country} {project_name}\{date.strftime("%y%m%d")} {patch_type}'
        
        
        if not os.path.isdir(total_path):                                                           
            os.mkdir(total_path)


        destination_folder = f'./{doc_type}/{datetime.now().strftime("%y%m%d_%H%M%S")}'#/{date.strftime("%y%m%d")} {patch_type}'

        if not os.path.isdir(destination_folder):                                                           
            os.mkdir(destination_folder)

    def 복붙문서적용(self):
        '''
        이메일복붙 등의 문서일 경우 디폴트폴더에서 복사/붙여넣기
        킥오프문서/qa요청문서 등...
        '''
        country = self.combox_country.currentText()
        project_name = self.combox_projectname.currentText()
        date = pd.to_datetime(self.dateedit_project.text())
        patch_type = self.combox_patchtype.currentText()        
        doc_type = self.combox_doctype.currentText()        
        
        #dir_path = self.input_destination_dir_path.text()
        #total_path = f'{dir_path}\{country} {project_name}\{date.strftime("%y%m%d")} {patch_type}'
        
        source_folder = f'./디폴트양식_{doc_type}' 
        destination_folder = f'./{doc_type}/{datetime.now().strftime("%y%m%d_%H%M%S")}/{date.strftime("%y%m%d")} {patch_type}'
        if not os.path.isdir(destination_folder):                                                           
            os.mkdir(destination_folder)
        try:
            # Make sure the destination folder exists
            # if not os.path.exists(destination_folder):
            #     os.makedirs(destination_folder)
            
            # Define the file mappings
            file_mappings = [
                ('QA_Kick-off 문서.docx', f'QA_Kick-off_{country}_{project_name}_{date.strftime("%y%m%d")} {patch_type} QA.docx'),
                ('QA_요청 문서.docx', f'QA_요청_{country}_{project_name}_{date.strftime("%y%m%d")} {patch_type} QA.docx')
                # Add more file mappings here
            ]
            
            for source_file, destination_file in file_mappings:
                source_path = os.path.join(source_folder, source_file)
                destination_path = os.path.join(destination_folder, destination_file)
                
                # Copy the file with renaming
                shutil.copy2(source_path, destination_path)
                
                print(f"File '{source_file}' copied and renamed to '{destination_file}'.")
        except Exception as e:
            print(f"An error occurred: {e}")


        # departure_path = self.input_departure_dir_path.text()
        # os.startfile(departure_path)
        # checklist_file_name = fr'CL_{country}_{project_name}_{date.strftime("%y%m%d")} {patch_type} QA.xlsx'
        # checklist_file_path = os.path.join(departure_path,f'{date.strftime("%y%m%d")} {patch_type}' ,checklist_file_name)

        # destination_file_path = os.path.join(dir_path, f'{country} {project_name}' ,checklist_file_name)
        # shutil.copy(checklist_file_path, destination_file_path)

    def 폴더내용물전체복사(self,source_folder, destination_folder):
        try:
            # Make sure the destination folder exists
            if not os.path.exists(destination_folder):
                os.makedirs(destination_folder)
            
            # Get a list of files in the source folder
            file_list = os.listdir(source_folder)
            
            for file_name in file_list:
                source_path = os.path.join(source_folder, file_name)
                destination_path = os.path.join(destination_folder, file_name)
                
                # Copy the file to the destination folder
                shutil.copy(source_path, destination_path)
                
                print(f"File '{file_name}' copied to '{destination_folder}'.")
        except Exception as e:
            print(f"An error occurred: {e}")


if __name__ == "__main__" :
    #QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv) 

    #WindowClass의 인스턴스 생성
    myWindow = WindowClass() 

    #프로그램 화면을 보여주는 코드
    myWindow.show()

    #프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec_()
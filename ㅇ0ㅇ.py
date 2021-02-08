import sys
import shutil
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QFileDialog, QGridLayout, QInputDialog
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt
import random

import win32com.client as win32  # 한/글 열기 위한 모듈
import pandas as pd  # 그 유명한 판다스. 엑셀파일을 다루기 위함


class xcelToHwpGenerator:
    def __init__(self, xlsPath, hwpPath):
        self.xlsPath = xlsPath
        self.hwpPath = hwpPath
        self.hwpResultPath = self.hwpPath[:-4] + "_result.hwp"

    def generate(self):
        excel = pd.read_excel(self.xlsPath, engine='openpyxl')  # 엑셀로 데이터프레임 생성
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")

        shutil.copyfile(self.hwpPath,
                        self.hwpResultPath)

        hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")

        hwp.Open(self.hwpResultPath)
        field_list = [i for i in hwp.GetFieldList().split("\x02")
                      ]  # 한글 필드리스트 변수에 저장
        field_list_set = set(field_list)  # 중복 제거됨
        field_list_dic = {}
        for field in field_list_set:
            field_list_dic[field] = field_list.count(field)  # 필드이름 : 원본에서의 필드수
        print(field_list_dic)
        print(hwp.GetFieldList())
        hwp.Run('SelectAll')  # Ctrl-A (전체선택)
        hwp.Run('Copy')  # Ctrl-C (복사)
        hwp.MovePos(3)  # 문서 끝으로 이동
        print(len(excel))
        # 엑셀파일 행갯수-1 만큼 한/글 페이지를 복사(기존에 한쪽이 있으니까)
        for i in range(len(excel) - 1):
            hwp.Run('Paste')  # Ctrl-V (붙여넣기)
            hwp.MovePos(3)  # 문서 끝으로 이동

        # 한/글 모든 페이지를 전부 순회하면서,
        for field in field_list_set:
            for order in range(field_list_dic[field] * len(excel)):
                hwp.PutFieldText(f'{field}{{{{{order}}}}}',  # f"{{{{{ㅇㅇㅇ}}}}}" "{{1}}" {를 출력하려면 {{를 입력.
                                 excel[field].iloc[order // field_list_dic[field]])


class MainWindow(QWidget):
    xlsPath = "select xls(x) file"
    hwpPath = "select hwp file"

    def __init__(self):
        super().__init__()
        self.setupUI()

    def setupUI(self):
        self.setGeometry(650, 350, 600, 300)
        self.setWindowTitle("ㅇ0ㅇ")

        self.title = QLabel("엑셀 데이터 한글문서 자동삽입 프로그램 ver 0.0.2")

        self.xcelFileBtn = QPushButton("xls(x) File Open")
        self.xcelFileBtn.setStyleSheet('background:#A9D18E')
        self.xcelFileBtn.clicked.connect(self.xcelFileBtnClicked)
        self.xlsPathTxt = QLabel(self.xlsPath)

        self.hangleFileBtn = QPushButton("hwp File Open")
        self.hangleFileBtn.setStyleSheet('background:#478FD1')
        self.hangleFileBtn.clicked.connect(self.hangleFileBtnClicked)
        self.hwpPathTxt = QLabel(self.hwpPath)

        self.generateBtn = QPushButton("I'm okay, let's create")
        self.generateBtn.clicked.connect(self.generateBtnCliked)

        grid = QGridLayout()
        grid.setSpacing(15)
        grid.addWidget(self.title, 0, 0, 1, 2, alignment=Qt.AlignHCenter)
        grid.addWidget(self.xcelFileBtn, 1, 0, alignment=Qt.AlignHCenter)
        grid.addWidget(self.xlsPathTxt, 1, 1, alignment=Qt.AlignHCenter)
        grid.addWidget(self.hangleFileBtn, 2, 0, alignment=Qt.AlignHCenter)
        grid.addWidget(self.hwpPathTxt, 2, 1, alignment=Qt.AlignHCenter)
        grid.addWidget(self.generateBtn, 2, 0, 2, 2, alignment=Qt.AlignHCenter)

        self.setLayout(grid)
        self.show()

    def xcelFileBtnClicked(self):
        fname = QFileDialog.getOpenFileName(self)
        print(fname)
        self.xlsPath = fname[0]
        self.xlsPathTxt.setText(self.xlsPath)

    def hangleFileBtnClicked(self):
        fname = QFileDialog.getOpenFileName(self)
        print(fname)
        self.hwpPath = fname[0]
        self.hwpPathTxt.setText(self.hwpPath)

    def generateBtnCliked(self):
        # self.generateBtn.setEnabled(False)
        generator = xcelToHwpGenerator(
            self.xlsPath, self.hwpPath)
        generator.generate()


프로그램무한반복 = QApplication(sys.argv)
실행인스턴스 = MainWindow()
프로그램무한반복.exec_()

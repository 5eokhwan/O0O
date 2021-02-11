import sys
import shutil
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QFileDialog, QGridLayout, QInputDialog
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import random

import win32com.client as win32
import pandas as pd


class xcelToHwpGenerator(QThread):
    progressUpdateSignal = pyqtSignal(object)

    xlsPath = None  # "select xls(x) file"
    xlsFieldList = None
    hwpPath = None  # "select hwp file"
    hwpFieldList = None

    # self.hwpResultPath = self.hwpPath[:-4] + "_result.hwp"
    hwpResultPath = None
    isCanGenerate = False

    def setHwpField(self):
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
        hwp.Open(self.hwpPath)
        field_list = [i for i in hwp.GetFieldList().split("\x02")
                      ]  # 한글 필드리스트 변수에 저장
        field_list = list(set(field_list))
        field_list.sort()  # 중복 제거한 후 사전정렬
        self.hwpFieldList = field_list  # 맴버변수에 저장
        hwp.Quit()

    def setXlsField(self):
        excel = pd.read_excel(self.xlsPath, engine='openpyxl')
        self.xlsFieldList = excel.columns.tolist()
        self.xlsFieldList.sort()

    def activeButton(self, UIbtn):
        # 한글리스트와 엑셀리스트의 요소를 비교
        # 한글에 있는 필드명이 엑셀에 없으면 활성화하지 않음
        self.isCanGenerate = True
        message = "생성하기"
        for field in self.hwpFieldList:
            if field in self.xlsFieldList:
                continue
            else:
                self.isCanGenerate = False
                message = f"'{field}' 필드이름이 엑셀에 존재하지 않습니다."
                break
        UIbtn.setEnabled(self.isCanGenerate)
        UIbtn.setText(message)

    def run(self):
        excel = pd.read_excel(self.xlsPath, engine='openpyxl')  # 엑셀로 데이터프레임 생성
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")

        self.hwpResultPath = self.hwpPath[:-4] + "_result.hwp"

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
        excelRowNum = len(excel)
        progressPercent = 0
        for page in range(excelRowNum - 1):
            progressPercent = page / (excelRowNum - 1) * 100
            # UIbtn.setText(f"row수만큼 페이지 복사중...{progressPercent}%")
            self.progressUpdateSignal.emit(
                ["엑셀 row수 만큼 페이지 복사중... ", round(progressPercent, 1)])
            hwp.Run('Paste')  # Ctrl-V (붙여넣기)
            hwp.MovePos(3)  # 문서 끝으로 이동
        # 한/글 모든 페이지를 전부 순회하면서,
        lst = list(field_list_dic.values())
        emptyFieldCount = sum(lst) * excelRowNum
        repeatCount = 0
        progressPercent = 0
        for field in field_list_set:
            # UIbtn.setText(f"{field}필드값 채우는 중... 총 {progressPercent}% 완료")
            for order in range(field_list_dic[field] * excelRowNum):
                self.progressUpdateSignal.emit(
                    [f"모든 필드값 채우는 중(현재 '{field}' 진행)... 총", round(progressPercent, 1)])
                hwp.PutFieldText(f'{field}{{{{{order}}}}}',  # f"{{{{{ㅇㅇㅇ}}}}}" "{{1}}" {를 출력하려면 {{를 입력.
                                 excel[field].iloc[order // field_list_dic[field]])
                repeatCount = repeatCount+1
                progressPercent = repeatCount / emptyFieldCount * 100
        self.progressUpdateSignal.emit(
            ["모든 작업이 완료되었습니다! 결과물을 확인하세요!", round(100.0, 1)])


class MainWindow(QWidget):
    xlsInfo = "select xls(x) file"
    hwpInfo = "select hwp file"

    def __init__(self):
        super().__init__()
        self.setupUI()
        self.generator = xcelToHwpGenerator()

    def setupUI(self):
        self.setGeometry(650, 350, 600, 300)
        self.setWindowTitle("ㅇ0ㅇ")

        self.title = QLabel("엑셀 데이터 한글문서 자동삽입 프로그램 ㅇ0ㅇ ver 0.0.3")

        self.xcelFileBtn = QPushButton("xls(x) File Open")
        self.xcelFileBtn.setStyleSheet('background:#A9D18E')
        self.xcelFileBtn.clicked.connect(self.xcelFileBtnClicked)
        self.xlsPathTxt = QLabel(self.xlsInfo)
        self.xlsValueList = QLabel("파일을 넣어주세요")

        self.hangleFileBtn = QPushButton("hwp File Open")
        self.hangleFileBtn.setStyleSheet('background:#478FD1')
        self.hangleFileBtn.clicked.connect(self.hangleFileBtnClicked)
        self.hwpPathTxt = QLabel(self.hwpInfo)
        self.hwpValueList = QLabel("파일을 넣어주세요")

        self.generateBtn = QPushButton("파일 선택을 먼저 해주세요")
        self.generateBtn.clicked.connect(self.generateBtnCliked)
        self.generateBtn.setEnabled(False)

        grid = QGridLayout()
        grid.setSpacing(15)
        grid.addWidget(self.title, 0, 0, 1, 2, alignment=Qt.AlignHCenter)
        grid.addWidget(self.xcelFileBtn, 1, 0, alignment=Qt.AlignHCenter)
        grid.addWidget(self.xlsPathTxt, 1, 1, alignment=Qt.AlignHCenter)
        grid.addWidget(self.xlsValueList, 2, 0, 1,
                       2, alignment=Qt.AlignHCenter)
        grid.addWidget(self.hangleFileBtn, 3, 0, alignment=Qt.AlignHCenter)
        grid.addWidget(self.hwpPathTxt, 3, 1, alignment=Qt.AlignHCenter)
        grid.addWidget(self.hwpValueList, 4, 0, 1,
                       2, alignment=Qt.AlignHCenter)
        grid.addWidget(self.generateBtn, 5, 0, 2, 2, alignment=Qt.AlignHCenter)

        self.setLayout(grid)
        self.show()

    def setGenerator(self):
        pass

    def xcelFileBtnClicked(self):
        fname = QFileDialog.getOpenFileName(self)
        print(fname)
        self.xlsInfo = fname[0]
        print(self.xlsInfo)
        if self.xlsInfo != '':
            if self.xlsInfo[-4:] == '.xls' or self.xlsInfo[-5:] == '.xlsx':
                self.xlsPathTxt.setText(self.xlsInfo)
                self.generator.xlsPath = self.xlsInfo
                self.generator.setXlsField()
                self.xlsValueList.setText(
                    "|".join(self.generator.xlsFieldList))
                # 버튼 활성
                if self.generator.hwpPath != None and self.generator.hwpPath != "":
                    self.generator.activeButton(self.generateBtn)
                    # self.generateBtn.setEnabled(self.generator.isCanGenerate)
                    # self.generateBtn.setText(message)
            else:
                self.xlsValueList.setText("xls(x) 파일이 아닙니다.")

    def hangleFileBtnClicked(self):
        fname = QFileDialog.getOpenFileName(self)
        print(fname)
        self.hwpInfo = fname[0]
        if self.hwpInfo != '':
            if self.hwpInfo[-4:] == '.hwp':
                self.hwpPathTxt.setText(self.hwpInfo)
                self.hwpPathTxt.setText(self.hwpInfo)
                self.generator.hwpPath = self.hwpInfo
                self.generator.setHwpField()
                print(self.generator.hwpFieldList)
                self.hwpValueList.setText(
                    "|".join(self.generator.hwpFieldList))
                # 버튼 활성
                if self.generator.xlsPath != None and self.generator.xlsPath != "":
                    self.generator.activeButton(self.generateBtn)
                    # self.generateBtn.setEnabled(self.generator.isCanGenerate)
                    # self.generateBtn.setText(message)
            else:
                self.hwpValueList.setText(".hwp 파일이 아닙니다.")

    def generateBtnCliked(self):
        self.generateBtn.setEnabled(False)
        self.generateBtn.setText("진행중")
        self.generator.progressUpdateSignal.connect(self.updateProgress)
        self.generator.start()

    def updateProgress(self, msg):
        self.generateBtn.setText(f"{msg[0]} {msg[1]}% 완료")


프로그램무한반복 = QApplication(sys.argv)
실행인스턴스 = MainWindow()
프로그램무한반복.exec_()

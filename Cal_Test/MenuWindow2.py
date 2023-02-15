import os
import sys
#import numpy as np
#import pandas as pd

import xlsxwriter
import xlrd

from ReportGenerator_EPS import CdfxExporter
from ReportGenerator_EPS import PrmXmlImpoter
from ReportGenerator_EPS import ReportGenrCDFX
from ReportGenerator_EPS import ReportGenr
from ReportGenerator_EPS import GetVariant

from Plausibility_EPS import PlausibilityCheck
from cdfxexporter_EPS import CdfxExporter
from cdfxexporter_EPS import PrmXmlImpoter
import xlsxwriter.utility

from tkinter import messagebox
from tkinter.filedialog import *
from tkinter.dialog import *

import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic

import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QCoreApplication

import openpyxl as op
import shutil

from openpyxl.styles import colors
from openpyxl.styles import Color
from openpyxl.styles import Font
from openpyxl import Workbook
from openpyxl.styles import PatternFill

from PyQt5.QtWidgets import *
from PyQt5 import uic

from collections import defaultdict

#사내 문서는 보안처리 되어있어서 openpyxl이나 xlsxwriter로 excel파일 못연다.
import xlwings as xw
from xlwings.utils import rgb_to_int
import resources_rc
param_exporter = CdfxExporter()
cdfx_filename = ""
param_importer = PrmXmlImpoter()

Newfile1 = ""
Newfile2 = ""

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
form_class, QtBaseClass = uic.loadUiType(BASE_DIR + r'\ui\menu_window_2.ui')

class SubWindow2(QMainWindow, form_class) :
    def __init__(self) :
        super().__init__()

        self.ImportCDFXPath = ''
        self.ImportAnalysisPath = ''

        self.setupUi(self)

        self.AnalysisFileImport.clicked.connect(self.ImportAnlysis)
        self.CDFXFileImport.clicked.connect(self.ImportCDFX)
        self.CheckBtn.clicked.connect(self.CheckFile)

        #분석할 SWC List
        self.SWCname = []
        self.SWC_list = []
        self.SWC_Applied = []
        #적용되지 않는다고 표시된 SWC List(나중에 CDFX에는 있는데 휴먼에러로 빠뜨린거 체크할 용도로 만들어둠)
        self.notAnlySWC = []

        self.df = defaultdict(list)

    def ImportAnlysis(self):
        SWAnlyFile = QFileDialog.getOpenFileName(self, "SW Analysis File 선택",'', "Exel files(*.xlsx);; 모든 파일(*)")
        NewFile1 = SWAnlyFile[0]
        self.ImportAnalysisPath.setText('')
        self.SWC_list = []
        self.SWC_Applied = []
        self.notAnlySWC = []
        self.SWCname = []

        if not SWAnlyFile[0]:
            QMessageBox.about(self, 'Notice', '파일을 업데이트 하세요')
            pass

        else:
            self.ImportAnalysisPath.setText(NewFile1)
            # 새로 불러올 때마다 값 초기화
            self.df = defaultdict(list)

            app = xw.App(visible=False)
            workbook = xw.Book(NewFile1)
            ws = workbook.sheets(1)

            self.df['SWC'] = ws.range('D4').expand('down').value
            self.df['Applied'] = ws.range('E4').expand('down').value

            for i in self.df["SWC"]:
                self.SWC_list.append(i)
            for i in self.df["Applied"]:
                self.SWC_Applied.append(i)

            for i in range(len(self.SWC_list)):
                #해당 SWC가 O 표시 되어 있을 때
                if self.SWC_Applied[i] == 'O' or self.SWC_Applied[i] == 'o' or self.SWC_Applied[i] == '0' or self.SWC_Applied[i] == 'ㅇ':
                    self.SWCname.append(self.SWC_list[i])

            app.kill()

    def ImportCDFX(self):
        fileImportCDFX = QFileDialog.getOpenFileName(self, "CDFX 파일 선택", '', "CDF20 files(*.cdfx);; 모든 파일(*)")
        NewFile2 = fileImportCDFX[0]
        self.ImportCDFXPath.setText('')

        if not fileImportCDFX[0]:
            QMessageBox.about(self, 'Notice', '파일을 업데이트 하세요')
            pass

        else:
            self.ImportCDFXPath.setText(NewFile2)
            param_exporter.__load_from_cdfx__(NewFile2)
            param_exporter.__export_to_xml__(NewFile2 + ".xml")
            param_importer.__load_xml__(NewFile2 + ".xml")

        return

    def CheckFile(self):
        # if self.ImportAnalysisPath == '' or self.ImportCDFXPath == '':
        #     QMessageBox.about(self, 'Notice', '파일을 업데이트 하세요!!!')
        #     pass
        # else:

        filesave = QFileDialog.getSaveFileName(self, "저장", 'SW Change Analysis_ing_Check.xlsx', "excel files(*.xlsx);; 모든 파일(*)")

        if not filesave[0]:
            QMessageBox.about(self, "Notice", "저장 파일 업데이트 하세요")
        else:
            chk_diff = 0
            shutil.copy(self.ImportAnalysisPath.toPlainText(), filesave[0])

            # CDFX에 적용되어 있는 SWC명 목록 사용 준비
            GetCDFX = GetVariant(param_importer)

            for i in range(len(self.SWC_list)):
                AnlySWC = self.SWC_list[i]

                #공백 제거
                if ' ' in AnlySWC:
                    AnlySWC = AnlySWC.replace(' ','')

                # 현재 검사할 swc가 실제 CDFX파일에 들어있는지 검사
                GetCDFX.__setvariant__("{}".format(AnlySWC))
                result = GetCDFX.__getname__()

                #해당 SWC가 적용된다고 기재했는데 실제로는 없음
                if AnlySWC in self.SWCname and result == 0:
                    chk_diff += 1
                    self.df['Applied'][i] = 'X'
                #해당 SWC가 적용되지 않는다고 기재했는데 실제로는 있음
                elif AnlySWC not in self.SWCname and result != 0 :
                    chk_diff += 1
                    self.df['Applied'][i] = 'O'

            app = xw.App(visible=False)
            workbook = xw.Book(filesave[0])
            ws = workbook.sheets(1)

            if chk_diff > 0:
                for i in range(len(self.SWC_Applied)):
                    if (self.SWC_Applied[i] == 'O' or self.SWC_Applied[i] == 'o' or self.SWC_Applied[i] == '0' or self.SWC_Applied[i] == 'ㅇ') and self.df['Applied'][i] == 'X':
                        num = 'E'+str(i+4)
                        a1 = ws.range(num)
                        a1.value = 'X'
                        a1.api.Font.Color = rgb_to_int((255,0,0)) #red

                    elif (self.SWC_Applied[i] == 'X' or self.SWC_Applied[i] == 'x' or self.SWC_Applied[i] == 'no') and self.df['Applied'][i] == 'O':
                        num = 'E'+str(i+4)
                        a1 = ws.range(num)
                        a1.value = 'O'
                        a1.api.Font.Color = rgb_to_int((255,0,0)) #red
            workbook.save(filesave[0])
            workbook.close()
            app.kill()
            QMessageBox.about(self, "Notice", "Check 완료, 다른 사항이 {}개 발견되었습니다.".format(chk_diff))

if __name__ == "__main__" :
    app = QApplication(sys.argv)
    myWindow = SubWindow2()
    myWindow.show()
    app.exec_()


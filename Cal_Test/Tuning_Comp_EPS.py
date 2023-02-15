import os
import sys

import numpy as np
import pandas as pd

import xlsxwriter
import xlrd

from ReportGenerator_EPS import CdfxExporter
from ReportGenerator_EPS import PrmXmlImpoter
from ReportGenerator_EPS import ReportGenrCDFX
from ReportGenerator_EPS import ReportGenr
from ReportGenerator_EPS import GetVariant

from MenuWindow import SubWindow
from MenuWindow import *
from MenuWindow2 import SubWindow2
from MenuWindow2 import *

from tkinter import messagebox
from tkinter.filedialog import *
from tkinter.dialog import *

import sys
from PyQt5.QtWidgets import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import QCoreApplication

import openpyxl as op

#사내 문서는 보안처리 되어있어서 openpyxl이나 xlsxwriter로 excel파일 못연다.
import xlwings as xw

param_exporter = CdfxExporter()
param_importer1 = PrmXmlImpoter()
param_importer2 = PrmXmlImpoter()
param_importer3 = PrmXmlImpoter()

fileCDFX = ""
fileCDFX1 = ""
fileCDFX2 = ""
fileCDFX3 = ""

from PyQt5.QtWidgets import *
from PyQt5 import uic

from collections import defaultdict
# pyrcc5 resources.qrc -o resources.py 로 리소스 파일을 _rc.py화해준 다음 import
import resources_rc

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
form_class, QtBaseClass = uic.loadUiType(BASE_DIR + r'\ui\main.ui')


class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.ExportCDFXPath = ''
        self.ExportSysPath = ''
        self.BeforeBaseCDFXPath = ''
        self.SWCImportPath = ''

        self.OriginMapNr = ''
        self.OriginMapNr1 = ''
        self.OriginMapNr2 = ''
        self.BuildMapNr = ''
        self.BuildMapNr1 = ''
        self.BuildMapNr2 = ''
        self.BeforeBaseMapNr1 = ''
        self.BeforeBaseMapNr2 = ''
        self.BeforeBaseMapNr3 = ''

        self.SWCname = []

        self.setupUi(self)

        self.OEMDrawFileImport.triggered.connect(self.OemDrawImport)
        self.actionOpen_CDM_Studio.triggered.connect(self.CDMopen)
        self.AnlysisCheckAction.triggered.connect(self.CheckAnalysisFile)

        self.ExportCDFXBtn.clicked.connect(self.NewBaseFileCDFX)
        self.ExportSysBtn.clicked.connect(self.TuningFileCDFX)
        self.BeforeBaseCDFXBtn.clicked.connect(self.BeforeBaseLineFileCDFX)

        self.DefaultMap.clicked.connect(self.Mapname)

        self.CDFXAnalysis.clicked.connect(self.AnlyCDFX)
        self.ExitBtn.clicked.connect(QCoreApplication.instance().quit)

        #SW Component Import 하기 위해..
        #1. 기존처럼 GEN 선택하면 자동으로 SWC import하는 방법
        self.UseAutoBtn.clicked.connect(self.SWCFileImport1)

        self.Gen3ABtn.clicked.connect(self.Gen3AImport)
        self.Gen3Btn.clicked.connect(self.Gen3Import)
        self.Gen4BBtn.clicked.connect(self.Gen4BBtnImport)

        #2. SW Analysis File로 SWC import하는 방법
        self.UsetSWAnlysBtn.clicked.connect(self.Click_UseFile)
        self.ImportSWCBtn.clicked.connect(self.SWCFileImport2)

        # SW Analysis Excel 파일 경로
        self.CopyFile = ''

    def NewBaseFileCDFX(self):

        fileCDFX = QFileDialog.getOpenFileName(self, "BaseLine 파일 선택",'', "CDF20 files(*.cdfx);; 모든 파일(*)")
        NewFile = fileCDFX[0]
        self.ExportCDFXPath.setText('')

        if not fileCDFX[0]:
            QMessageBox.about(self, 'Notice', '\n파일을 업데이트 하세요')
            pass

        else:
            self.ExportCDFXPath.setText(NewFile)
            param_exporter.__load_from_cdfx__(NewFile)
            param_exporter.__export_to_xml__(NewFile + ".xml")
            param_importer1.__load_xml__(NewFile + ".xml")

        return

    def TuningFileCDFX(self):

        fileCDFX1 = QFileDialog.getOpenFileName(self, "Tuning 파일 선택",'', "CDF20 files(*.cdfx);; 모든 파일(*)")
        NewFile = fileCDFX1[0]
        self.ExportSysPath.setText('')

        if not fileCDFX1[0]:
            QMessageBox.about(self, 'Notice', '\n파일을 업데이트 하세요')
            pass

        else:
            self.ExportSysPath.setText(NewFile)
            param_exporter.__load_from_cdfx__(NewFile)
            param_exporter.__export_to_xml__(NewFile + ".xml")
            param_importer2.__load_xml__(NewFile + ".xml")

        return

    def BeforeBaseLineFileCDFX(self):

        fileCDFX2 = QFileDialog.getOpenFileName(self, "Previous BaseLine 파일 선택", '', "CDF20 files(*.cdfx);; 모든 파일(*)")
        NewFile = fileCDFX2[0]
        self.BeforeBaseCDFXPath.setText('')

        if not fileCDFX2[0]:
            QMessageBox.about(self, 'Notice', '\n파일을 업데이트 하세요')
            pass

        else:
            self.BeforeBaseCDFXPath.setText(NewFile)
            param_exporter.__load_from_cdfx__(NewFile)
            param_exporter.__export_to_xml__(NewFile + ".xml")
            param_importer3.__load_xml__(NewFile + ".xml")

        return

    def OemDrawImport(self):
        self.w = SubWindow()
        self.w.show()

    def Mapname(self):
        if self.DefaultMap.isChecked():
            self.OriginMapNr.setText('A00')
            self.OriginMapNr1.setText('A01')
            self.OriginMapNr2.setText('A02')
            self.BuildMapNr.setText('A00')
            self.BuildMapNr1.setText('A01')
            self.BuildMapNr2.setText('A02')
            self.BeforeBaseMapNr1.setText('A00')
            self.BeforeBaseMapNr2.setText('A01')
            self.BeforeBaseMapNr3.setText('A02')

        else:
            self.OriginMapNr.setText('')
            self.OriginMapNr1.setText('')
            self.OriginMapNr2.setText('')
            self.BuildMapNr.setText('')
            self.BuildMapNr1.setText('')
            self.BuildMapNr2.setText('')
            self.BeforeBaseMapNr1.setText('')
            self.BeforeBaseMapNr2.setText('')
            self.BeforeBaseMapNr3.setText('')

    def CDMopen(self):
        path = os.getenv('CDMPath')
        os.popen(path + "CDMStudio32.exe")

    def AnlyCDFX(self):
        YPos = 6
        YPos2 = 6
        filesave = QFileDialog.getSaveFileName(self, "저장", '.xlsx', "excel files(*.xlsx);; 모든 파일(*)")

        if not filesave[0]:
            QMessageBox.about(self, "Notice", "\n파일을 저장할 경로를 업데이트 하세요")
        else:

            workbook = xlsxwriter.Workbook(filesave[0])  # +".xlsx")
            Report_gen = ReportGenrCDFX(param_importer1, param_importer2, param_importer3, workbook)

            NewBase = self.ExportCDFXPath.toPlainText()
            TuningFile = self.ExportSysPath.toPlainText()
            BeforeBase = self.BeforeBaseCDFXPath.toPlainText()

            Report_gen.__createsheet__("Summary")
            Report_gen.__summarytext__("Summary", NewBase, TuningFile, BeforeBase)

            Orignal = self.OriginMapNr.toPlainText()
            Build = self.BuildMapNr.toPlainText()
            BeBase = self.BeforeBaseMapNr1.toPlainText()

            for i in range(len(self.SWCname)):
                AnlySWC = self.SWCname[i]
                # 공백 제거
                if ' ' in AnlySWC:
                    AnlySWC = AnlySWC.replace(' ', '')
                Report_gen.__createsheet__("{}_Internal".format(AnlySWC))
                Report_gen.__setvariant__("{}".format(AnlySWC), "{}".format(AnlySWC), "{}".format(AnlySWC))

                # 해당 SWC의 튜닝값이 local, global에 둘 다 있음
                if AnlySWC in param_importer1.__local_prms__.keys() and AnlySWC in param_importer1.__global_prms__.keys():
                    result_ = Report_gen.__drawMapSingle__()
                    result = Report_gen.__drawMapSingle2__(result_[2], result_[3]+12)
                # 해당 SWC의 튜닝값이 global에만 있음
                elif AnlySWC in param_importer1.__global_prms__.keys():
                    result = Report_gen.__drawMapSingle2__(2,2)
                # 해당 SWC의 튜닝값이 local에만 있거나 해당 SWC가 존재하지 않음
                else:
                    result_ = Report_gen.__drawMapSingle__()
                    result = result_[:2]

                YPos = Report_gen.__drawdecision__("{}_Internal".format(AnlySWC), result[0], YPos)
                YPos2 = Report_gen.__drawdecision2__("{}_Internal".format(AnlySWC), result[1], YPos2)

                # (AnlySWC in param_importer1.__local_prms__.keys() or AnlySWC in param_importer1.__global_prms__.keys())
                # Layer2가 있는 SWC들은 Layer1,2 비교도 할 수 있도록
                if AnlySWC[-1] == '2' and result[0] != 2:
                    New_name = "New_" + AnlySWC[0:len(AnlySWC) - 1] + "1_2"
                    Report_gen.__createsheet__("{}".format(New_name))
                    Report_gen.__setvariantNew__("{}".format(AnlySWC[0:len(AnlySWC) - 1]), "{}".format(AnlySWC))

                    # 글로벌, 로컬 튜닝변수 모두 있음
                    if AnlySWC in param_importer1.__local_prms__.keys() and AnlySWC in param_importer1.__global_prms__.keys():
                        result_ = Report_gen.__drawMapSingleLay2__(2,2,"__local_prms__")
                        # result_ = [result, x좌표, y좌표]
                        result_2 = Report_gen.__drawMapSingleLay2__(result_[1],result_[2],"__global_prms__")
                        result = result_2[0]

                    # 글로벌만
                    elif AnlySWC in param_importer1.__global_prms__.keys():
                        result_ = Report_gen.__drawMapSingleLay2__(2,2,"__global_prms__")
                        result = result_[0]
                    # 로컬만 or 없음
                    else:
                        result_ = Report_gen.__drawMapSingleLay2__(2,2,"__local_prms__")
                        result = result_[0]

                    YPos = Report_gen.__drawdecision__("{}".format(New_name), result, YPos)
                    YPos2 = Report_gen.__drawdecision2__("{}".format(New_name), -1, YPos2)

            if self.SWCname == []:
                QMessageBox.about(self, "Notice", "SW Component를 Import하세요")

            if Orignal == '':
                pass

            elif Build == '':
                pass
            elif BeBase == '':
                pass

            else:
                Report_gen.__createsheet__("New_" + Orignal + "_Tuning_" + Build)
                Report_gen.__setvariant__(Orignal + "_TunVrnt", Build + "_TunVrnt", BeBase + "_TunVrnt")
                result = Report_gen.__drawMap__()
                YPos = Report_gen.__drawdecision__("New_" + Orignal + "_Tuning_" + Build, result[0], YPos)
                YPos2 = Report_gen.__drawdecision2__("New_" + Orignal + "_Tuning_" + Build, result[1], YPos2)

            Orignal1 = self.OriginMapNr1.toPlainText()
            Build1 = self.BuildMapNr1.toPlainText()
            BeBase1 = self.BeforeBaseMapNr2.toPlainText()

            if Orignal1 == '':
                pass

            elif Build1 == '':
                pass

            elif BeBase1 == '':
                pass
            else:
                Report_gen.__createsheet__("New_" + Orignal1 + "_Tuning_" + Build1)
                Report_gen.__setvariant__(Orignal1 + "_TunVrnt", Build1 + "_TunVrnt", BeBase1 + "_TunVrnt")
                result = Report_gen.__drawMap__()
                YPos = Report_gen.__drawdecision__("New_" + Orignal1 + "_Tuning_" + Build1, result[0], YPos)
                YPos2 = Report_gen.__drawdecision2__("New_" + Orignal1 + "_Tuning_" + Build1, result[1], YPos2)

            Orignal2 = self.OriginMapNr2.toPlainText()
            Build2 = self.BuildMapNr2.toPlainText()
            BeBase2 = self.BeforeBaseMapNr3.toPlainText()

            if Orignal2 == '':
                pass

            elif Build2 == '':
                pass

            elif BeBase2 == '':
                pass
            else:
                Report_gen.__createsheet__("New_" + Orignal2 + "_Tuning_" + Build2)
                Report_gen.__setvariant__(Orignal2 + "_TunVrnt", Build2 + "_TunVrnt", BeBase2 + "_TunVrnt")
                result = Report_gen.__drawMap__()
                YPos = Report_gen.__drawdecision__("New_" + Orignal2 + "_Tuning_" + Build2, result[0], YPos)
                YPos2 = Report_gen.__drawdecision2__("New_" + Orignal2 + "_Tuning_" + Build2, result[1], YPos2)

            workbook.close()

            if self.SWCImportPath.toPlainText() != '':
                app = xw.App(visible=False)

                wb_from = xw.Book(r"{}".format(self.CopyFile))
                wb_to = xw.Book(r"{}".format(filesave[0]))
                ws_from = wb_from.sheets[0]
                ws_to = wb_to.sheets[0]
                ws_from.api.Copy(Before=ws_to.api)
                wb_to.save(filesave[0])

                wb_from.close()
                wb_to.close()

                app.kill()

            QMessageBox.about(self, "Notice", "\n비교 자료생성 완료!")

    def SWCFileImport1(self):
        self.SWCImportPath.setText('')
        self.Gen3ABtn.setEnabled(True)
        self.Gen3Btn.setEnabled(True)
        self.Gen4BBtn.setEnabled(True)
        self.SWAnlysisFileImport.setDisabled(True)
        self.SWCImportPath.setDisabled(True)
        self.ImportSWCBtn.setDisabled(True)

        return

    def Gen3AImport(self):
        self.Gen3Btn.setCheckState(0)
        self.Gen4BBtn.setCheckState(0)
        self.SWCname = []

        self.SWCname = ['CLBaseFunctions', 'CLBaseFunctions2','FinalTqCmd','FinalTqCmd2','LkaCtrl','LkaCtrl2',
                        'AWLAgCtrl','CLSigProc','JudCompCtrl', 'CacCtrl', 'CacPreCdn', 'CLCoCtrlPreCdn', 'OverLoadProtn',
                        'PinionAgEstimr','PullCmp','ModArbn','ShmCompCtrl', 'VsmCtrl',"LoaMitCtrl",'LogicCorrlnDiagc','LoTCmp',
                        'SoftEndStop','Haptic', 'StrWhlAgTrackingCtrl']
        self.SWCname.sort()

        return

    def Gen3Import(self):
        self.Gen3ABtn.setCheckState(0)
        self.Gen4BBtn.setCheckState(0)
        self.SWCname = []

        self.SWCname = ['Assist', 'Assist2', 'Boost', 'DampgCtrl', 'DampgCtrl2', 'FbRetCtrl', 'StatFricCmp', 'HysTqCtrl', 'HysTqCtrl2',
                        'RetCtrl', 'RetCtrl2', 'PhaCmp', 'PhaCmp2', 'ADASPreCdn', 'EACCtrl', 'HptcFb', 'FinalTqCmd', 'FinalTqCmd2',
                        'LkaCtrl', 'LkaCtrl2', 'AWLAgCtrl', 'JudCompCtrl', 'CacCtrl', 'CacPreCdn', 'CLCoCtrlPreCdn',
                        'OverLoadProtn', 'PinionAgEstimr','PullCmp','ModArbn','ShmCompCtrl', 'VsmCtrl',"LoaMitCtrl",'LogicCorrlnDiagc','LoTCmp',
                        'SoftEndStop','Haptic', 'StrWhlAgTrackingCtrl']
        self.SWCname.sort()
        return

    def Gen4BBtnImport(self):
        self.Gen3ABtn.setCheckState(0)
        self.Gen3Btn.setCheckState(0)
        self.SWCname = []

        self.SWCname = ['FinalTqCmd','FinalTqCmd2','LkaCtrl','LkaCtrl2','PhaCmp','PhaCmp2',
                        'AWLAgCtrl','CLSigProc','JudCompCtrl','CacCtrl','CacPreCdn', 'CLBaseFunctions', 'CLBaseFunctions2',
                        'CLCoCtrlPreCdn','CLVAFPreCdn','MotTqCmdSeln','OverLoadProtn','PinionAgEstimr',
                        'PullCmp','ModArbn','ShmCompCtrl','VsmCtrl','LoaMitCtrl','LogicCorrlnDiagc','LoTCmp','SoftEndStop','Haptic',
                        'StrWhlAgTrackingCtrl']

        self.SWCname.sort()
        return

    def Click_UseFile(self):
        if self.Gen3ABtn.isChecked():
            self.Gen3ABtn.toggle()
        if self.Gen3Btn.isChecked():
            self.Gen3Btn.toggle()
        if self.Gen4BBtn.isChecked():
            self.Gen4BBtn.toggle()
        self.Gen3ABtn.setDisabled(True)
        self.Gen3Btn.setDisabled(True)
        self.Gen4BBtn.setDisabled(True)

        self.SWAnlysisFileImport.setEnabled(True)
        self.SWCImportPath.setEnabled(True)
        self.ImportSWCBtn.setEnabled(True)

    def SWCFileImport2(self):
        SWAnlyFile = QFileDialog.getOpenFileName(self, "SW Analysis File 선택",'', "Exel files(*.xlsx);; 모든 파일(*)")
        self.CopyFile = SWAnlyFile[0]
        self.Gen3ABtn.setDisabled(True)
        self.Gen3Btn.setDisabled(True)
        self.Gen4BBtn.setDisabled(True)

        self.SWCname = []
        self.SWCImportPath.setText('')

        if not SWAnlyFile[0]:
            QMessageBox.about(self, 'Notice', '\n파일을 업데이트 하세요')

        else:
            self.SWCImportPath.setText(self.CopyFile)
            df = defaultdict(list)

            app = xw.App(visible=False)
            workbook = xw.Book(self.CopyFile)
            ws = workbook.sheets(1)

            df['SWC'] = ws.range('D4').expand('down').value
            df['Applied'] = ws.range('E4').expand('down').value

            SWC_list = []
            SWC_Applied = []

            for i in df.keys():
                if i == "SWC":
                    SWC_list = df[i]
                else:
                    SWC_Applied = df[i]

            for i in range(len(SWC_list)):
                #해당 SWC가 O 표시 되어 있을 때
                if SWC_Applied[i] == 'O' or SWC_Applied[i] == 'o' or SWC_Applied[i] == '0' or SWC_Applied[i] == 'ㅇ':
                    self.SWCname.append(SWC_list[i])
            workbook.close()
            app.kill()

    def CheckAnalysisFile(self):
        self.w2 = SubWindow2()
        self.w2.show()

if __name__=="__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()



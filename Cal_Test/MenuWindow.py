import os
import sys
#import numpy as np
#import pandas as pd

import xlsxwriter
from ReportGenerator_EPS import CdfxExporter
from ReportGenerator_EPS import PrmXmlImpoter
from ReportGenerator_EPS import ReportGenrCDFX
from ReportGenerator_EPS import ReportGenr

import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic

import resources_rc

param_exporter = CdfxExporter()
param_importer1 = PrmXmlImpoter()
param_importer2 = PrmXmlImpoter()
param_importer3 = PrmXmlImpoter()

fileCDFX = ""
fileCDFX1 = ""
fileCDFX2 = ""
fileCDFX3 = ""

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
form_class, QtBaseClass = uic.loadUiType(BASE_DIR + r'\ui\menu_window.ui')

class SubWindow(QMainWindow, form_class) :
    def __init__(self) :
        super().__init__()
        self.setupUi(self)

        self.OEMDrawFileImport.clicked.connect(self.importOEM)
        self.OEMDraw.clicked.connect(self.OemDrawAnaly)

    def importOEM(self):
        fileImportCDFX = QFileDialog.getOpenFileName(self, "승인도 파일 선택", '', "CDF20 files(*.cdfx);; 모든 파일(*)")
        NewFile = fileImportCDFX[0]
        self.ImportCDFXPath.setText('')

        if not fileImportCDFX[0]:
            QMessageBox.about(self, 'Notice', '파일을 업데이트 하세요')
            pass

        else:
            self.ImportCDFXPath.setText(NewFile)
            param_exporter.__load_from_cdfx__(NewFile)
            param_exporter.__export_to_xml__(NewFile + ".xml")
            param_importer2.__load_xml__(NewFile + ".xml")

    def OemDrawAnaly(self):

        if self.OEMDrawStNr.toPlainText() == '':
            QMessageBox.about(self, "Notice", "맵 번호를 넣으세요")

        elif self.OEMDrawEndNr.toPlainText() == '':
            QMessageBox.about(self, "Notice", "맵 번호를 넣으세요")

        else:
            filesave = QFileDialog.getSaveFileName(self, "저장", '', "excel files(*.xlsx);; 모든 파일(*)")
            workbook1 = xlsxwriter.Workbook(filesave[0])

            report_gen1 = ReportGenr(param_importer2, param_importer2, workbook1)

            OEMMapStNr1 = self.OEMDrawStNr.toPlainText()
            OEMMapEndNr1 = self.OEMDrawEndNr.toPlainText()
            OEMMapNrOfVCNr1 = self.MapNrOfVCNr.toPlainText()

            OEMMapStNr = int(OEMMapStNr1[1:3])
            OEMMapEndNr = int(OEMMapEndNr1[1:3])
            OEMMapNrOfVCNr = int(OEMMapNrOfVCNr1)
            # a= ShimmyCompOEM.get()

            VCCode = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E"]
            Cnt = 0

            for i in range(OEMMapStNr, OEMMapEndNr, OEMMapNrOfVCNr):
                Cnt += 1
                if i < 10:
                    report_gen1.__createsheet__("VC0" + VCCode[Cnt - 1])
                    report_gen1.__setvariant__("A0" + str(i) + "_TunVrnt", "A0" + str(i + 1) + "_TunVrnt")
                else:
                    report_gen1.__createsheet__("VC0" + VCCode[Cnt - 1])
                    report_gen1.__setvariant__("A" + str(i) + "_TunVrnt", "A" + str(i + 1) + "_TunVrnt")

                LastPosn = 7
                MapWidth = 0

                LastPosn, MapWidth = report_gen1.__drawMap__('AssiGain', LastPosn, MapWidth,
                                                             {'GradientEna': False, 'PlausibilityCheckEna': False,
                                                              'Title': "Assist Control",
                                                              'Xlabel': "Steering Torque[Nm]",
                                                              'Ylabel': "Assist Torque[Nm]",
                                                              'GRADTitle': "Gradient of Assist Map",
                                                              'GRADXlabel': "Steering Torque[Nm]",
                                                              'GRADYlabel': "Gradient of Assist"})


                LastPosn, MapWidth = report_gen1.__drawCur__('StatFricVehSpdFac', LastPosn, MapWidth,
                                                             {'GradientEna': False, 'PlausibilityCheckEna': False,
                                                              'Title': "Static Friction Control",
                                                              'Xlabel': "Vehicle Speed[kph]",
                                                              'Ylabel': "Gain[%]",
                                                              'GRADTitle': "Gradient of RackF Map",
                                                              'GRADXlabel': "STW Angle[deg]",
                                                              'GRADYlabel': "Gradient of RackF"})

                LastPosn, MapWidth = report_gen1.__drawCur__('BoostSteerWhlTqHPFCutOffFreqCur1', LastPosn, MapWidth,
                                                             {'GradientEna': False, 'PlausibilityCheckEna': False,
                                                              'Title': "Boost Control1 HPF Fc",
                                                              'Xlabel': "Vehicle Speed[kph]",
                                                              'Ylabel': "Boost Control Fc[Hz]",
                                                              'GRADTitle': "Gradient of RackF Map",
                                                              'GRADXlabel': "STW Angle[deg]",
                                                              'GRADYlabel': "Gradient of RackF"})

                LastPosn, MapWidth = report_gen1.__drawCur__('BoostSteerWhlTqHPFCutOffFreqCur2', LastPosn, MapWidth,
                                                             {'GradientEna': False, 'PlausibilityCheckEna': False,
                                                              'Title': "Boost Control2 HPF Fc",
                                                              'Xlabel': "Vehicle Speed[kph]",
                                                              'Ylabel': "Boost Control Fc[Hz]",
                                                              'GRADTitle': "Gradient of RackF Map",
                                                              'GRADXlabel': "STW Angle[deg]",
                                                              'GRADYlabel': "Gradient of RackF"})
                LastPosn, MapWidth = report_gen1.__drawMap__('BoostTqMap1', LastPosn, MapWidth,
                                                             {'GradientEna': False, 'PlausibilityCheckEna': False,
                                                              'Title': "Boost Control1",
                                                              'Xlabel': "Steering Torque[Nm]",
                                                              'Ylabel': "Boost Torque[Nm]",
                                                              'GRADTitle': "Gradient of Bosst Map",
                                                              'GRADXlabel': "HPF Steering Torque[Nm]",
                                                              'GRADYlabel': "Gradient of Boost"})
                LastPosn, MapWidth = report_gen1.__drawMap__('BoostTqMap2', LastPosn, MapWidth,
                                                             {'GradientEna': False, 'PlausibilityCheckEna': False,
                                                              'Title': "Boost Control2",
                                                              'Xlabel': "Steering Torque[Nm]",
                                                              'Ylabel': "Boost Torque[Nm]",
                                                              'GRADTitle': "Gradient of Bosst Map",
                                                              'GRADXlabel': "HPF Steering Torque[Nm]",
                                                              'GRADYlabel': "Gradient of Boost"})

                LastPosn, MapWidth = report_gen1.__drawMap__('RetTqMap', LastPosn, MapWidth,
                                                             {'GradientEna': False, 'PlausibilityCheckEna': False,
                                                              'Title': "Return Cotrol",
                                                              'Xlabel': " Steering Angle[deg]",
                                                              'Ylabel': "Return Torque[Nm]",
                                                              'GRADTitle': "Gradient of Bosst Map",
                                                              'GRADXlabel': "HPF Steering Torque[Nm]",
                                                              'GRADYlabel': "Gradient of Boost"})

                LastPosn, MapWidth = report_gen1.__drawMap__('FbRetTarSteerWhlAgSpd', LastPosn, MapWidth,
                                                             {'GradientEna': False, 'PlausibilityCheckEna': False,
                                                              'Title': "Return Velocity Control",
                                                              'Xlabel': "Steering Angle[deg]",
                                                              'Ylabel': "Target Velocity[deg/sec]",
                                                              'GRADTitle': "Gradient of Bosst Map",
                                                              'GRADXlabel': "HPF Steering Torque[Nm]",
                                                              'GRADYlabel': "Gradient of Boost"})

                LastPosn, MapWidth = report_gen1.__drawMap__('DampgTqOut', LastPosn, MapWidth,
                                                             {'GradientEna': False, 'PlausibilityCheckEna': False,
                                                              'Title': "Damping Control",
                                                              'Xlabel': "Steering Velocity[deg/sec]",
                                                              'Ylabel': "Damping Torque[Nm]",
                                                              'GRADTitle': "Gradient of Damping Map",
                                                              'GRADXlabel': "Steer Wheel Angle[deg]",
                                                              'GRADYlabel': "Gradient of Damping"})

                LastPosn, MapWidth = report_gen1.__drawMap__('DampgSteerWhlAgVehSpdFac', LastPosn, MapWidth,
                                                             {'GradientEna': False, 'PlausibilityCheckEna': False,
                                                              'Title': "Damping Control(Steering Angle)",
                                                              'Xlabel': "Steering Angle[deg]",
                                                              'Ylabel': "Damping Gain[%]",
                                                              'GRADTitle': "Gradient of Damping Map",
                                                              'GRADXlabel': "Steer Wheel Angle[deg]",
                                                              'GRADYlabel': "Gradient of Damping"})

                LastPosn, MapWidth = report_gen1.__drawCur__('InerCmpVehSpdMomJ', LastPosn, MapWidth,
                                                             {'GradientEna': False, 'PlausibilityCheckEna': False,
                                                              'Title': "Inertia Contorl",
                                                              'Xlabel': "Vehicle Speed[kph]",
                                                              'Ylabel': "Gain[%]",
                                                              'GRADTitle': "Gradient of RackF Map",
                                                              'GRADXlabel': "STW Angle[deg]",
                                                              'GRADYlabel': "Gradient of RackF"})

                LastPosn, MapWidth = report_gen1.__drawCur__('HysTqCtrlVehSpdFac', LastPosn, MapWidth,
                                                             {'GradientEna': False, 'PlausibilityCheckEna': False,
                                                              'Title': "Hysteresis Control",
                                                              'Xlabel': "Vehicle Speed[kph]",
                                                              'Ylabel': "Gain[%]",
                                                              'GRADTitle': "Gradient of RackF Map",
                                                              'GRADXlabel': "STW Angle[deg]",
                                                              'GRADYlabel': "Gradient of RackF"})


                if self.OEMDShimmy.isChecked():
                    LastPosn, MapWidth = report_gen1.__drawCur__('ShmCompCtrlVehGain', LastPosn, MapWidth,
                                                                 {'GradientEna': False, 'PlausibilityCheckEna': False,
                                                                  'Title': "Shimmy Control",
                                                                  'Xlabel': "Vehicle Speed[kph]",
                                                                  'Ylabel': "Gain",
                                                                  'GRADTitle': "Gradient of RackF Map",
                                                                  'GRADXlabel': "STW Angle[deg]",
                                                                  'GRADYlabel': "Gradient of RackF"})
                else:
                    pass

                MaxXSize = 33
                report_gen1.__drawHeader__(1, 1, MaxXSize + 2, 2, 'CalVersion', 36)

            workbook1.close()

            QMessageBox.about(self,"Notice", "승인도 생성 완료")

        return

if __name__ == "__main__" :
    app = QApplication(sys.argv)
    myWindow = SubWindow()
    myWindow.show()
    app.exec_()
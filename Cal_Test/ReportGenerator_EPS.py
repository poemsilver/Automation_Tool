import xlsxwriter
import os
import sys
import openpyxl as op

import pandas as pd
import numpy as np
from pandas import DataFrame

from tkinter import messagebox

from Plausibility_EPS import PlausibilityCheck
from cdfxexporter_EPS import CdfxExporter
from cdfxexporter_EPS import PrmXmlImpoter
import xlsxwriter.utility

import os
import getpass

import xml.etree.ElementTree as ET

class ReportGenrCDFX:
    def __init__(self, importer1, importer2,importer3,workbook):
        self.importer1 = importer1
        self.importer2 = importer2
        self.importer3 = importer3
        self.workbook = workbook
        self.current_sheet = None

    def __setvariant__(self, variant_name1,variant_name2,variant_name3):
        self.variant_name1 = variant_name1
        self.variant_name2 = variant_name2
        self.variant_name3 = variant_name3

    def __setvariantNew__(self, variant_name1,variant_name2):
        self.variant_name1 = variant_name1
        self.variant_name2 = variant_name2


    def __setvariantTuning__(self,variant_name1,variant_name2):
        self.variant_name1 = variant_name1
        self.variant_name2 = variant_name2


    def __createsheet__(self, sheet_name):
        workbook = self.workbook
        self.sheet_name = sheet_name
        self.current_sheet = workbook.get_worksheet_by_name(sheet_name)

        if self.current_sheet is None:
            self.current_sheet = workbook.add_worksheet(sheet_name)

    def __drawMap__(self):

        workbook = self.workbook
        worksheet = self.current_sheet
        worksheet.set_column('A:FH', 5)
        StartPoint1_x = 2
        StartPoint1_Y = 2
        result = [0, 0]

        data_format_data = workbook.add_format({'border': 1})
        data_format_axis = workbook.add_format({'border': 1, 'bg_color': '#DCE6F1'})
        WriteTitleFormat = workbook.add_format({'font_color': '#000000', 'bold': True, 'font_size': 11})
        WrongFormat = workbook.add_format({'bg_color':'#FF0000'})

        # 조건부 서식 Format
        FormatForAfter_Change = workbook.add_format({'bold': True,'bg_color': 'gray'})
        FormatForAfter_NotEqual = workbook.add_format({'bold': True,'bg_color': 'red'})
        #FormatForAfter_NotEqual = workbook.add_format({'font_color': '#000000', 'bold': True})
        FormatForAfter_Greater = workbook.add_format({'font_color': '#000000', 'bg_color': '#F2DCDB', 'bold': True})
        FormatForAfter_Equal = workbook.add_format({'font_color': '#000000', 'bg_color': '#FFFFFF'})
        FormatForAfter_AxisEqual = workbook.add_format({'font_color': '#000000', 'bg_color': '#DCE6F1'})



        #VALUE 챠트 Layer1/Layer2

        try:
            ConstValue1 = self.importer1.__global_prms__[self.variant_name1]["VALUE"]
            ConstValue2 = self.importer2.__global_prms__[self.variant_name2]["VALUE"]
            ConstValue3 = self.importer3.__global_prms__[self.variant_name3]["VALUE"]
            Nr = max(len(ConstValue1),len(ConstValue2),len(ConstValue3))
            worksheet.write(StartPoint1_Y,StartPoint1_x+5,"New",WriteTitleFormat)
            worksheet.write(StartPoint1_Y, StartPoint1_x + 6, "Tuning", WriteTitleFormat)
            worksheet.write(StartPoint1_Y, StartPoint1_x + 7, "Before", WriteTitleFormat)
            index=0

            for i in ConstValue1.keys():
                index = index+1
                p=0
                try:
                    p=1
                    VariableValue1 = self.importer1.__global_prms__[self.variant_name1]["VALUE"][i].value
                    p=2
                    VariableValue2 = self.importer2.__global_prms__[self.variant_name2]["VALUE"][i].value
                    if VariableValue1 != VariableValue2:
                        result[0] = 1

                    p = 3
                    VariableValue3 = self.importer3.__global_prms__[self.variant_name3]["VALUE"][i].value

                except KeyError:
                    if p == 1:
                        VariableValue2 = self.importer2.__global_prms__[self.variant_name2]["VALUE"][i].value
                        VariableValue3 = self.importer3.__global_prms__[self.variant_name3]["VALUE"][i].value
                        VariableValue1 = 'NA'
                        if VariableValue1 != VariableValue2[0]:
                            result[0] = 1
                        if VariableValue1 != VariableValue3[0]:
                            result[1] = 1
                    elif p == 2:
                        VariableValue1 = self.importer1.__global_prms__[self.variant_name1]["VALUE"][i].value
                        VariableValue3 = self.importer3.__global_prms__[self.variant_name3]["VALUE"][i].value
                        VariableValue2 = 'NA'
                        if VariableValue1[0] != VariableValue2:
                            result[0] = 1
                        if VariableValue1 != VariableValue3:
                            result[1] = 1
                    else:
                        VariableValue1 = self.importer1.__global_prms__[self.variant_name1]["VALUE"][i].value
                        VariableValue2 = self.importer2.__global_prms__[self.variant_name2]["VALUE"][i].value
                        VariableValue3 = 'NA'
                        if VariableValue1[0] != VariableValue3:
                            result[1] = 1


                worksheet.write(StartPoint1_Y+index, StartPoint1_x-1, i, WriteTitleFormat)
                worksheet.write(StartPoint1_Y+index, StartPoint1_x+5, VariableValue1, data_format_data)
                worksheet.write(StartPoint1_Y+index, StartPoint1_x+6, VariableValue2, data_format_data)
                worksheet.write(StartPoint1_Y+index, StartPoint1_x+7, VariableValue3, data_format_data)

            Layer1_data_sp = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y+1, StartPoint1_x+5)
            Layer2_data_sp = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y+1, StartPoint1_x + 6)
            Layer3_data_sp = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y+1, StartPoint1_x + 7)

            after_data_range = xlsxwriter.utility.xl_range_abs(StartPoint1_Y+1, StartPoint1_x + 6 , StartPoint1_Y+Nr ,StartPoint1_x+6)
            after_data_range1 = xlsxwriter.utility.xl_range_abs(StartPoint1_Y+1, StartPoint1_x + 7 , StartPoint1_Y+Nr ,StartPoint1_x+7)

            worksheet.conditional_format(after_data_range, {'type': 'formula',
                                                            'criteria': Layer1_data_sp + ' <> ' + Layer2_data_sp,
                                                            'format': FormatForAfter_NotEqual})

            worksheet.conditional_format(after_data_range1, {'type': 'formula',
                                                            'criteria': Layer1_data_sp + ' <> ' + Layer3_data_sp,
                                                            'format': FormatForAfter_NotEqual})



            # Curve Layer1/Layer2

            StartPoint1_x = StartPoint1_x
            StartPoint1_Y = StartPoint1_Y + Nr

            CurValue = self.importer1.__global_prms__[self.variant_name1]["CURVE"]
            CurValue_lay2 = self.importer1.__global_prms__[self.variant_name2]["CURVE"]
            CurValue_lay3 = self.importer1.__global_prms__[self.variant_name3]["CURVE"]
            CurIndex = 0


            for j in CurValue.keys():
                CurIndex = CurIndex + 1

                p = 0
                try:
                    p = 1
                    CurveValue1 = self.importer1.__global_prms__[self.variant_name1]["CURVE"][j]  # Curve 파라미터명
                    p = 2
                    CurveValue2_lay2 = self.importer2.__global_prms__[self.variant_name2]["CURVE"][j]

                    p = 3
                    CurveValue3 = self.importer3.__global_prms__[self.variant_name3]["CURVE"][j]

                    p = 1
                    CurveValue1_X = self.importer1.__global_prms__[self.variant_name1]["COM_AXIS"][CurveValue1.__xaxis__].value
                    # Curve 파라미터의 x축
                    p = 2
                    CurveValue2_X_lay2 = self.importer2.__global_prms__[self.variant_name2]["COM_AXIS"][CurveValue2_lay2.__xaxis__].value
                    p = 3
                    CurveValue3_X = self.importer3.__global_prms__[self.variant_name3]["COM_AXIS"][CurveValue3.__xaxis__].value

                    CurveValue1_Value = self.importer1.__global_prms__[self.variant_name1]["CURVE"][j].value
                    # Curve 파라미터의 값 (비교대상)
                    CurveValue2_Value_lay2 = self.importer2.__global_prms__[self.variant_name2]["CURVE"][j].value
                    CurveValue3_Value = self.importer3.__global_prms__[self.variant_name3]["CURVE"][j].value


                except KeyError:

                    if p == 1:
                        CurveValue2_lay2 = self.importer2.__global_prms__[self.variant_name2]["CURVE"][j]
                        CurveValue3 = self.importer3.__global_prms__[self.variant_name3]["CURVE"][j]

                        CurveValue2_X_lay2 = self.importer2.__global_prms__[self.variant_name2]["COM_AXIS"][CurveValue2_lay2.__xaxis__].value
                        CurveValue3_X = self.importer3.__global_prms__[self.variant_name3]["COM_AXIS"][CurveValue3.__xaxis__].value
                        CurveValue1_X = ['NA']* len(CurveValue2_X_lay2)

                        CurveValue2_Value_lay2 = self.importer2.__global_prms__[self.variant_name2]["CURVE"][j].value
                        CurveValue3_Value = self.importer3.__global_prms__[self.variant_name3]["CURVE"][j].value
                        CurveValue1_Value = ['NA']* len(CurveValue1_X)

                    elif p == 2:
                        CurveValue1 = self.importer1.__global_prms__[self.variant_name1]["CURVE"][j]  # Curve 파라미터명
                        CurveValue3 = self.importer3.__global_prms__[self.variant_name3]["CURVE"][j]

                        CurveValue1_X = self.importer1.__global_prms__[self.variant_name2]["COM_AXIS"][CurveValue1.__xaxis__].value
                        CurveValue3_X = self.importer3.__global_prms__[self.variant_name3]["COM_AXIS"][CurveValue3.__xaxis__].value
                        CurveValue2_X_lay2 = ['NA']*len(CurveValue1_X)

                        CurveValue1_Value = self.importer1.__global_prms__[self.variant_name1]["CURVE"][j].value
                        CurveValue3_Value = self.importer3.__global_prms__[self.variant_name3]["CURVE"][j].value
                        CurveValue2_Value_lay2 = ['NA']*len(CurveValue1_Value)

                    else:
                        CurveValue1 = self.importer1.__global_prms__[self.variant_name1]["CURVE"][j]  # Curve 파라미터명
                        CurveValue2_lay2 = self.importer2.__global_prms__[self.variant_name2]["CURVE"][j]

                        CurveValue1_X = self.importer1.__global_prms__[self.variant_name1]["COM_AXIS"][CurveValue1.__xaxis__].value
                        CurveValue2_X_lay2 = self.importer2.__global_prms__[self.variant_name2]["COM_AXIS"][CurveValue2_lay2.__xaxis__].value
                        CurveValue3_X = ['NA']*len(CurveValue1_X)

                        CurveValue1_Value = self.importer1.__global_prms__[self.variant_name1]["CURVE"][j].value
                        CurveValue2_Value_lay2 = self.importer2.__global_prms__[self.variant_name2]["CURVE"][j].value
                        CurveValue3_Value = ['NA']*len(CurveValue1_Value)


                if CurIndex >= 2:
                    StartPoint1_Y = StartPoint1_Y+4
                else:
                    StartPoint1_Y = StartPoint1_Y

                worksheet.write(StartPoint1_Y+2, StartPoint1_x-1, j, WriteTitleFormat)

                leng = max(len(CurveValue1_X),len(CurveValue2_X_lay2),len(CurveValue3_Value))

                if len(CurveValue1_X) < leng:
                    CurveValue1_X = CurveValue2_X_lay2.tolist()
                    CurveValue1_Value = CurveValue2_Value_lay2.tolist()
                    for i in range(leng - len(CurveValue1_X)):
                        CurveValue1_X.append('NA')
                        CurveValue1_Value.append('NA')
                if len(CurveValue2_X_lay2) < leng:
                    CurveValue2_X_lay2 = CurveValue2_X_lay2.tolist()
                    CurveValue2_Value_lay2 = CurveValue2_Value_lay2.tolist()
                    for i in range(leng - len(CurveValue2_X_lay2)):
                        CurveValue2_X_lay2.append('NA')
                        CurveValue2_Value_lay2.append('NA')
                if len(CurveValue3_X) < leng:
                    CurveValue3_X = CurveValue2_X_lay2.tolist()
                    CurveValue3_Value = CurveValue2_Value_lay2.tolist()
                    for i in range(leng - len(CurveValue3_X)):
                        CurveValue3_X.append('NA')
                        CurveValue3_Value.append('NA')

                for i in range(leng):
                    # Axis
                    worksheet.write(StartPoint1_Y+3, StartPoint1_x+i-1, CurveValue1_X[i], data_format_axis)
                    worksheet.write(StartPoint1_Y + 3, StartPoint1_x + i + 1 + len(CurveValue1_X)+2,CurveValue2_X_lay2[i], data_format_axis)
                    worksheet.write(StartPoint1_Y + 3, StartPoint1_x + i + 1 + 4+(len(CurveValue1_X)*2),CurveValue3_X[i], data_format_axis)

                    if CurveValue1_X[i] != CurveValue2_X_lay2[i]:
                        result[0] = 1
                    else:
                        pass

                    if CurveValue1_X[i] != CurveValue3_X[i]:
                        result[1] = 1
                    else:
                        pass

                    #Curve
                    worksheet.write(StartPoint1_Y+4, StartPoint1_x+i-1, CurveValue1_Value[i], data_format_data)
                    worksheet.write(StartPoint1_Y + 4, StartPoint1_x + i + 1+len(CurveValue1_X)+2, CurveValue2_Value_lay2[i], data_format_data)
                    worksheet.write(StartPoint1_Y + 4, StartPoint1_x + i + 1+4+(len(CurveValue1_X)*2), CurveValue3_Value[i], data_format_data)

                    if CurveValue1_Value[i] != CurveValue2_Value_lay2[i]:
                        result[0] = 1
                    else:
                        pass

                    if CurveValue1_Value[i] != CurveValue3_Value[i]:
                        result[1] = 1
                    else:
                        pass

                Layer1_data_CurAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y+3, StartPoint1_x-1)
                Layer2_data_CurAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y+3, StartPoint1_x + 2 + len(CurveValue1_X)+2-1)
                Before_data_CurAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y+3, StartPoint1_x + 1 + 4+(len(CurveValue1_X)*2))

                Layer1_data_CurData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y+4, StartPoint1_x-1)
                Layer2_data_CurData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y+4, StartPoint1_x + 2 + len(CurveValue1_X)+2-1)
                Before_data_CurData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y+4, StartPoint1_x + 1+4+(len(CurveValue1_X)*2))

                after_data_Currange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y+4,StartPoint1_x + 2 + len(CurveValue1_X)+2-1,
                                                                      StartPoint1_Y+4,StartPoint1_x + 2 + len(CurveValue1_X)+1+i)

                after_data_CurrAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y+3,StartPoint1_x + 2 + len(CurveValue1_X)+2-1,
                                                                          StartPoint1_Y+3,StartPoint1_x + 2 + len(CurveValue1_X)+1+i)

                Before_data_CurrAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 3,
                                                                          StartPoint1_x + 1 + 4+(len(CurveValue1_X)*2),
                                                                          StartPoint1_Y + 3,
                                                                          StartPoint1_x + 1 + 4+(len(CurveValue1_X)*2) + i)

                Before_data_Currange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 4,
                                                                      StartPoint1_x + 1+4+(len(CurveValue1_X)*2),
                                                                      StartPoint1_Y + 4,
                                                                      StartPoint1_x + 1+4+(len(CurveValue1_X)*2) + i)



                worksheet.conditional_format(after_data_CurrAxisange, {'type': 'formula',
                                                                       'criteria': Layer1_data_CurAxis + ' <> ' + Layer2_data_CurAxis,
                                                                        'format': FormatForAfter_NotEqual})
                worksheet.conditional_format(after_data_Currange, {'type': 'formula',
                                                                  'criteria': Layer1_data_CurData + ' <> ' + Layer2_data_CurData,
                                                                  'format': FormatForAfter_NotEqual})

                worksheet.conditional_format(Before_data_CurrAxisange, {'type': 'formula',
                                                                       'criteria': Layer1_data_CurAxis + ' <> ' + Before_data_CurAxis,
                                                                        'format': FormatForAfter_Change})

                worksheet.conditional_format(Before_data_Currange, {'type': 'formula',
                                                                    'criteria': Layer1_data_CurData + ' <> ' + Before_data_CurData,
                                                                    'format': FormatForAfter_Change})

            # MAP Lay1/2
            StartPoint1_x = StartPoint1_x
            StartPoint1_Y = StartPoint1_Y + 6


            MapValue = self.importer1.__global_prms__[self.variant_name1]["MAP"]
            MapValue_lay2 = self.importer2.__global_prms__[self.variant_name2]["MAP"]
            MapIndex = 0
            MapValue1_Y = 0


            for k in MapValue.keys():
                MapValue1_y_old = MapValue1_Y
                MapIndex = MapIndex + 1

                p=0
                try:
                    p = 1
                    MapValue1 = self.importer1.__global_prms__[self.variant_name1]["MAP"][k]
                    p = 2
                    MapValue2_lay2 = self.importer2.__global_prms__[self.variant_name2]["MAP"][k]
                    p = 3
                    MapValue3 = self.importer3.__global_prms__[self.variant_name3]["MAP"][k]

                    p = 1
                    MapValue1_X = self.importer1.__global_prms__[self.variant_name1]["COM_AXIS"][
                        MapValue1.__xaxis__].value
                    p = 2
                    MapValue2_X_lay2 = self.importer2.__global_prms__[self.variant_name2]["COM_AXIS"][
                        MapValue2_lay2.__xaxis__].value
                    p = 3
                    MapValue3_X = self.importer3.__global_prms__[self.variant_name3]["COM_AXIS"][
                        MapValue3.__xaxis__].value

                    p = 1
                    MapValue1_Y = self.importer1.__global_prms__[self.variant_name1]["COM_AXIS"][
                        MapValue1.__yaxis__].value
                    p = 2
                    MapValue2_Y_lay2 = self.importer2.__global_prms__[self.variant_name2]["COM_AXIS"][
                        MapValue2_lay2.__yaxis__].value
                    p = 3
                    MapValue3_Y = self.importer3.__global_prms__[self.variant_name3]["COM_AXIS"][
                        MapValue3.__yaxis__].value

                    p = 1
                    MapValue1_Value = self.importer1.__global_prms__[self.variant_name1]["MAP"][k].value
                    p = 2
                    MapValue1_Value_lay2 = self.importer2.__global_prms__[self.variant_name2]["MAP"][k].value
                    p = 3
                    MapValue3_Value = self.importer3.__global_prms__[self.variant_name3]["MAP"][k].value



                except KeyError:

                    if p == 1:
                        MapValue2_lay2 = self.importer2.__global_prms__[self.variant_name2]["MAP"][k]
                        MapValue3 = self.importer3.__global_prms__[self.variant_name3]["MAP"][k]

                        MapValue2_X_lay2 = self.importer2.__global_prms__[self.variant_name2]["COM_AXIS"][
                            MapValue2_lay2.__xaxis__].value
                        MapValue3_X = self.importer3.__global_prms__[self.variant_name3]["COM_AXIS"][
                            MapValue3.__xaxis__].value
                        MapValue1_X = ['NA']*len(MapValue2_X_lay2)

                        MapValue2_Y_lay2 = self.importer2.__global_prms__[self.variant_name2]["COM_AXIS"][
                            MapValue2_lay2.__yaxis__].value
                        MapValue3_Y = self.importer3.__global_prms__[self.variant_name3]["COM_AXIS"][
                            MapValue3.__yaxis__].value
                        MapValue1_Y = ['NA']*len(MapValue2_X_lay2)

                        MapValue1_Value_lay2 = self.importer2.__global_prms__[self.variant_name2]["MAP"][k].value
                        MapValue3_Value = self.importer3.__global_prms__[self.variant_name3]["MAP"][k].value
                        MapValue1_Value = [['NA']*len(MapValue2_X_lay2)]*len(MapValue2_Y_lay2)


                    elif p == 2:
                        MapValue1 = self.importer1.__global_prms__[self.variant_name1]["MAP"][k]
                        MapValue3 = self.importer3.__global_prms__[self.variant_name3]["MAP"][k]

                        MapValue1_X = self.importer1.__global_prms__[self.variant_name1]["COM_AXIS"][
                            MapValue1.__xaxis__].value
                        MapValue3_X = self.importer3.__global_prms__[self.variant_name3]["COM_AXIS"][
                            MapValue3.__xaxis__].value
                        MapValue2_X_lay2 = ['NA']*len(MapValue1_X)

                        MapValue1_Y = self.importer1.__global_prms__[self.variant_name1]["COM_AXIS"][
                            MapValue1.__yaxis__].value
                        MapValue3_Y = self.importer3.__global_prms__[self.variant_name3]["COM_AXIS"][
                            MapValue3.__yaxis__].value
                        MapValue2_Y_lay2 = ['NA']*len(MapValue1_Y)

                        MapValue1_Value = self.importer1.__global_prms__[self.variant_name1]["MAP"][k].value
                        MapValue3_Value = self.importer3.__global_prms__[self.variant_name3]["MAP"][k].value
                        MapValue1_Value_lay2 = [['NA']*len(MapValue1_X)]*len(MapValue1_Y)

                    else:
                        MapValue1 = self.importer1.__global_prms__[self.variant_name1]["MAP"][k]
                        MapValue2_lay2 = self.importer2.__global_prms__[self.variant_name2]["MAP"][k]

                        MapValue1_X = self.importer1.__global_prms__[self.variant_name1]["COM_AXIS"][
                            MapValue1.__xaxis__].value
                        MapValue2_X_lay2 = self.importer2.__global_prms__[self.variant_name2]["COM_AXIS"][
                            MapValue2_lay2.__xaxis__].value
                        MapValue3_X = ['NA']*len(MapValue1_X)

                        MapValue1_Y = self.importer1.__global_prms__[self.variant_name1]["COM_AXIS"][
                            MapValue1.__yaxis__].value
                        MapValue2_Y_lay2 = self.importer2.__global_prms__[self.variant_name2]["COM_AXIS"][
                            MapValue2_lay2.__yaxis__].value
                        MapValue3_Y = ['NA']*len(MapValue1_Y)

                        MapValue1_Value = self.importer1.__global_prms__[self.variant_name1]["MAP"][k].value
                        MapValue1_Value_lay2 = self.importer2.__global_prms__[self.variant_name2]["MAP"][k].value
                        MapValue3_Value = [['NA']*len(MapValue1_X)]*len(MapValue1_Y)



                if MapIndex >= 2:
                    StartPoint1_Y = StartPoint1_Y+len(MapValue1_y_old)+3
                else:
                    StartPoint1_Y = StartPoint1_Y

                worksheet.write(StartPoint1_Y, StartPoint1_x - 1, k, WriteTitleFormat)

                lengX = max(len(MapValue1_X),len(MapValue2_X_lay2),len(MapValue3_X))
                lengY = max(len(MapValue1_Y),len(MapValue2_Y_lay2),len(MapValue3_Y))

                #X축 크기 맞추기
                if len(MapValue1_X) < lengX:
                    MapValue1_X = MapValue1_X.tolist()
                    MapValue1_Value = MapValue1_Value.tolist()
                    for i in range(lengX-len(MapValue1_X)):
                        MapValue1_X.append('NA')
                        for j in range(len(MapValue1_Y)):
                            for k in range(lengX-len(MapValue1_X)):
                                MapValue1_Value[j].append('NA')
                if len(MapValue2_X_lay2) < lengX:
                    MapValue2_X_lay2 = MapValue2_X_lay2.tolist()
                    MapValue1_Value_lay2 = MapValue1_Value_lay2.tolist()
                    for i in range(lengX-len(MapValue2_X_lay2)):
                        MapValue2_X_lay2.append('NA')
                        for j in range(len(MapValue2_Y_lay2)):
                            for k in range(lengX-len(MapValue2_X_lay2)):
                                MapValue1_Value_lay2[j].append('NA')
                if len(MapValue3_X) < lengX:
                    MapValue3_X = MapValue3_X.tolist()
                    MapValue3_Value = MapValue3_Value.tolist()
                    for i in range(lengX-len(MapValue3_X)):
                        MapValue3_X.append('NA')
                        for j in range(len(MapValue3_Y)):
                            for k in range(lengX-len(MapValue3_X)):
                                MapValue3_Value[j].append('NA')

                #Y축 크기 맞추기
                if len(MapValue1_Y) < lengY:
                    MapValue1_Y = MapValue1_Y.tolist()
                    MapValue1_Value = MapValue1_Value.tolist()
                    for i in range(lengY - len(MapValue1_Y)):
                        MapValue1_Y.append('NA')
                        temp = ['NA'] * len(MapValue1_X)
                        MapValue1_Value.append(temp)
                if len(MapValue2_Y_lay2) < lengY:
                    MapValue2_Y_lay2 = MapValue2_Y_lay2.tolist()
                    MapValue1_Value_lay2 = MapValue1_Value_lay2.toslit()
                    for i in range(lengY - len(MapValue2_Y_lay2)):
                        MapValue2_Y_lay2.append('NA')
                        temp = ['NA'] * len(MapValue2_X_lay2)
                        MapValue1_Value_lay2.append(temp)
                if len(MapValue3_Y) < lengY:
                    MapValue3_Y = MapValue3_Y.tolist()
                    MapValue3_Value = MapValue3_Value.tolist()
                    for i in range(lengY - len(MapValue3_Y)):
                        MapValue3_Y.append('NA')
                        temp = ['NA'] * len(MapValue3_X)
                        MapValue3_Value.append(temp)

                for l in range(lengX):
                    worksheet.write(StartPoint1_Y+1, l + StartPoint1_x, MapValue1_X[l], data_format_axis)
                    worksheet.write(StartPoint1_Y+1, l + StartPoint1_x + len(MapValue1_X)+3, MapValue2_X_lay2[l],data_format_axis)
                    worksheet.write(StartPoint1_Y + 1, l + StartPoint1_x +3+(len(MapValue1_X)*2)+3, MapValue3_X[l], data_format_axis)

                    if MapValue1_X[l] != MapValue2_X_lay2[l]:
                        result[0] = 1
                    else:
                        pass

                    if MapValue1_X[l] != MapValue3_X[l]:
                        result[1] = 1
                    else:
                        pass

                Layer1_data_MapXAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 1, StartPoint1_x)

                Layer2_data_MapXAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 1,
                                                                            StartPoint1_x + len(MapValue1_X) + 3)

                Before_data_MapXAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 1,
                                                                            StartPoint1_x +3+(len(MapValue1_X)*2)+3)

                after_data_MapXAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 1,
                                                                          StartPoint1_x + len(MapValue1_X) + 3,
                                                                          StartPoint1_Y + 1,
                                                                          StartPoint1_x + len(MapValue1_X) + 3 + l)

                Before_data_MapXAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 1,
                                                                          StartPoint1_x +3+(len(MapValue1_X)*2)+3,
                                                                          StartPoint1_Y + 1,
                                                                          StartPoint1_x +3+(len(MapValue1_X)*2)+3 + l)

                worksheet.conditional_format(after_data_MapXAxisange, {'type': 'formula',
                                                                       'criteria': Layer1_data_MapXAxis + ' <> ' + Layer2_data_MapXAxis,
                                                                         'format': FormatForAfter_NotEqual})

                worksheet.conditional_format(Before_data_MapXAxisange, {'type': 'formula',
                                                                       'criteria': Layer1_data_MapXAxis + ' <> ' + Before_data_MapXAxis,
                                                                       'format': FormatForAfter_Change})


                for m in range(len(MapValue1_Y)):
                    worksheet.write(StartPoint1_Y+m+2,StartPoint1_x-1,MapValue1_Y[m],data_format_axis)
                    worksheet.write(StartPoint1_Y + m + 2, StartPoint1_x - 1+len(MapValue1_X)+3, MapValue2_Y_lay2[m], data_format_axis)
                    worksheet.write(StartPoint1_Y + m + 2, StartPoint1_x - 1+3+(len(MapValue1_X)*2) + 3, MapValue3_Y[m], data_format_axis)

                    if MapValue1_Y[m] != MapValue2_Y_lay2[m]:
                        result[0] = 1
                    else:
                        pass

                    if MapValue1_Y[m] != MapValue3_Y[m]:
                        result[1] = 1
                    else:
                        pass

                Layer1_data_MapYAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y+2,StartPoint1_x-1)
                Layer2_data_MapYAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y+2,
                                                                            StartPoint1_x - 1+len(MapValue1_X)+3)

                Before_data_MapYAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 2,
                                                                            StartPoint1_x - 1+3+(len(MapValue1_X)*2) + 3)


                after_data_MapYAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y+2,
                                                                          StartPoint1_x - 1+len(MapValue1_X)+3,
                                                                          StartPoint1_Y+2+m,
                                                                          StartPoint1_x - 1+len(MapValue1_X)+3)

                Before_data_MapYAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y+2,
                                                                          StartPoint1_x - 1+3+(len(MapValue1_X)*2) + 3,
                                                                          StartPoint1_Y+2+m,
                                                                          StartPoint1_x - 1+3+(len(MapValue1_X)*2) + 3)




                worksheet.conditional_format(after_data_MapYAxisange, {'type': 'formula',
                                                                       'criteria': Layer1_data_MapYAxis + ' <> ' + Layer2_data_MapYAxis,
                                                                       'format': FormatForAfter_NotEqual})

                worksheet.conditional_format(Before_data_MapYAxisange, {'type': 'formula',
                                                                       'criteria': Layer1_data_MapYAxis + ' <> ' + Before_data_MapYAxis,
                                                                       'format': FormatForAfter_Change})

                for p in range(len(MapValue1_Y)):
                    for o in range(len(MapValue1_X)):
                        worksheet.write(StartPoint1_Y+2+p, o+StartPoint1_x,MapValue1_Value[p][o],data_format_data)
                        worksheet.write(StartPoint1_Y+2+p, o+StartPoint1_x+len(MapValue1_X)+3, MapValue1_Value_lay2[p][o], data_format_data)
                        worksheet.write(StartPoint1_Y+2+p, o + StartPoint1_x + (len(MapValue1_X)*2) + 6, MapValue3_Value[p][o], data_format_data)

                        if MapValue1_Value[p][o] != MapValue1_Value_lay2[p][o]:
                            result[0] = 1
                        else:
                            pass

                        if MapValue1_Value[p][o] != MapValue3_Value[p][o]:
                            result[1] = 1
                        else:
                            pass

                Layer1_data_MapData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y+2,StartPoint1_x)
                Layer2_data_MapData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y+2,
                                                                            StartPoint1_x+len(MapValue1_X)+3)

                Before_data_MapData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 2,
                                                                           StartPoint1_x + (len(MapValue1_X)*2) + 6)

                after_data_MapDatasange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y+2,
                                                                          StartPoint1_x+len(MapValue1_X)+3,
                                                                          StartPoint1_Y+2+p,
                                                                          o+StartPoint1_x+len(MapValue1_X)+3)

                Before_data_MapDatasange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 2,
                                                                          StartPoint1_x + (len(MapValue1_X)*2) + 6,
                                                                          StartPoint1_Y + 2 + p,
                                                                          o + StartPoint1_x + (len(MapValue1_X)*2) + 6)

                worksheet.conditional_format(after_data_MapDatasange, {'type': 'formula',
                                                                       'criteria': Layer1_data_MapData + ' <> ' + Layer2_data_MapData,
                                                                       'format': FormatForAfter_NotEqual})

                worksheet.conditional_format(Before_data_MapDatasange, {'type': 'formula',
                                                                       'criteria': Layer1_data_MapData + ' <> ' + Before_data_MapData,
                                                                       'format': FormatForAfter_Change})

        except KeyError:
            result[0] = 2
            result[1] = 2

        return result

    def __drawMapSingle__(self):
        workbook = self.workbook
        worksheet = self.current_sheet
        worksheet.set_column('A:FH',5)

        StartPoint1_x = 2
        StartPoint1_Y = 2

        result = [0, 0]

        data_format_data = workbook.add_format({'border': 1})
        data_format_axis = workbook.add_format({'border': 1, 'bg_color': '#DCE6F1'})
        WriteTitleFormat = workbook.add_format({'font_color': '#000000', 'bold': True, 'font_size': 11})
        WrongFormat = workbook.add_format({'bg_color': '#FF0000'})

        # 조건부 서식 Format
        FormatForAfter_NotEqual = workbook.add_format({'bold': True,'bg_color': 'red'})
        FormatForAfter_Change = workbook.add_format({'bold': True,'bg_color': 'gray'})
        #FormatForAfter_NotEqual = workbook.add_format({'font_color': '#000000', 'bold': True})
        FormatForAfter_Greater = workbook.add_format({'font_color': '#000000', 'bg_color': '#F2DCDB', 'bold': True})
        FormatForAfter_Equal = workbook.add_format({'font_color': '#000000', 'bg_color': '#FFFFFF'})
        FormatForAfter_AxisEqual = workbook.add_format({'font_color': '#000000', 'bg_color': '#DCE6F1'})


        #VALUE 챠트 Layer1/Layer2
        try :
            ConstValue1 = self.importer1.__local_prms__[self.variant_name1]["VALUE"]
            Nr = len(ConstValue1)
            worksheet.write(StartPoint1_Y,StartPoint1_x+5,"New",WriteTitleFormat)
            worksheet.write(StartPoint1_Y, StartPoint1_x + 6, "Tuning", WriteTitleFormat)
            worksheet.write(StartPoint1_Y, StartPoint1_x + 7, "Before", WriteTitleFormat)
            index=0

            for i in ConstValue1.keys():
                p = 0
                try:
                    index = index + 1
                    p=1
                    VariableValue1 = self.importer1.__local_prms__[self.variant_name1]["VALUE"][i].value
                    p=2
                    VariableValue2 = self.importer2.__local_prms__[self.variant_name2]["VALUE"][i].value

                    if VariableValue1 != VariableValue2:
                        result[0] = 1

                    p=3
                    VariableValue3 = self.importer3.__local_prms__[self.variant_name3]["VALUE"][i].value

                    if VariableValue1 != VariableValue3:
                        result[1] = 1


                except KeyError:

                    if p == 1:
                        VariableValue2 = self.importer2.__local_prms__[self.variant_name2]["VALUE"][i].value
                        VariableValue3 = self.importer3.__local_prms__[self.variant_name3]["VALUE"][i].value
                        VariableValue1 = 'NA'
                        if VariableValue1 != VariableValue2[0]:
                            result[0] = 1
                        if VariableValue1 != VariableValue3[0]:
                            result[1] = 1
                    elif p == 2:
                        VariableValue1 = self.importer1.__local_prms__[self.variant_name1]["VALUE"][i].value
                        VariableValue3 = self.importer3.__local_prms__[self.variant_name3]["VALUE"][i].value
                        VariableValue2 = 'NA'
                        if VariableValue1[0] != VariableValue2:
                            result[0] = 1
                        if VariableValue1 != VariableValue3:
                            result[1] = 1
                    else:
                        VariableValue1 = self.importer1.__local_prms__[self.variant_name1]["VALUE"][i].value
                        VariableValue2 = self.importer2.__local_prms__[self.variant_name2]["VALUE"][i].value
                        VariableValue3 = 'NA'
                        if VariableValue1[0] != VariableValue3:
                            result[1] = 1

                worksheet.write(StartPoint1_Y + index, StartPoint1_x - 1, i, WriteTitleFormat)
                worksheet.write(StartPoint1_Y + index, StartPoint1_x + 5, VariableValue1, data_format_data)
                worksheet.write(StartPoint1_Y + index, StartPoint1_x + 6, VariableValue2, data_format_data)
                worksheet.write(StartPoint1_Y + index, StartPoint1_x + 7, VariableValue3, data_format_data)


            Layer1_data_sp = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y+1, StartPoint1_x+5)
            Layer2_data_sp = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y+1, StartPoint1_x + 6)
            Layer3_data_sp = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 1, StartPoint1_x + 7)


            after_data_range = xlsxwriter.utility.xl_range_abs(StartPoint1_Y+1, StartPoint1_x + 6 , StartPoint1_Y+Nr ,StartPoint1_x+6)
            after_data_range1 = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 1, StartPoint1_x + 7, StartPoint1_Y + Nr,
                                                                StartPoint1_x + 7)

            worksheet.conditional_format(after_data_range, {'type': 'formula',
                                                            'criteria': Layer1_data_sp + ' <> ' + Layer2_data_sp,
                                                            'format': FormatForAfter_NotEqual})

            worksheet.conditional_format(after_data_range1, {'type': 'formula',
                                                            'criteria': Layer1_data_sp + ' <> ' + Layer3_data_sp,
                                                            'format': FormatForAfter_Change})



            # Curve Layer1/Layer2

            StartPoint1_x = StartPoint1_x
            StartPoint1_Y = StartPoint1_Y + Nr

            CurValue = self.importer1.__local_prms__[self.variant_name1]["CURVE"]
            CurValue_lay2 = self.importer2.__local_prms__[self.variant_name2]["CURVE"]
            CurIndex = 0

            for j in CurValue.keys():
                CurIndex = CurIndex + 1
                p = 0

                try:
                    p = 1
                    CurveValue1 = self.importer1.__local_prms__[self.variant_name1]["CURVE"][j]  # Curve 파라미터명
                    p = 2
                    CurveValue2_lay2 = self.importer2.__local_prms__[self.variant_name2]["CURVE"][j]

                    p = 3
                    CurveValue3 = self.importer3.__local_prms__[self.variant_name3]["CURVE"][j]

                    p = 1
                    CurveValue1_X = self.importer1.__local_prms__[self.variant_name1]["COM_AXIS"][
                        CurveValue1.__xaxis__].value
                    # Curve 파라미터의 x축
                    p = 2
                    CurveValue2_X_lay2 = self.importer2.__local_prms__[self.variant_name2]["COM_AXIS"][
                        CurveValue2_lay2.__xaxis__].value
                    p = 3
                    CurveValue3_X = self.importer3.__local_prms__[self.variant_name3]["COM_AXIS"][
                        CurveValue3.__xaxis__].value

                    CurveValue1_Value = self.importer1.__local_prms__[self.variant_name1]["CURVE"][j].value
                    # Curve 파라미터의 값 (비교대상)
                    CurveValue2_Value_lay2 = self.importer2.__local_prms__[self.variant_name2]["CURVE"][j].value
                    CurveValue3_Value = self.importer3.__local_prms__[self.variant_name3]["CURVE"][j].value


                except KeyError:
                    if p == 1:
                        CurveValue2_lay2 = self.importer2.__local_prms__[self.variant_name2]["CURVE"][j]
                        CurveValue3 = self.importer3.__local_prms__[self.variant_name3]["CURVE"][j]

                        CurveValue2_X_lay2 = self.importer2.__local_prms__[self.variant_name2]["COM_AXIS"][CurveValue2_lay2.__xaxis__].value
                        CurveValue3_X = self.importer3.__local_prms__[self.variant_name3]["COM_AXIS"][
                            CurveValue3.__xaxis__].value
                        CurveValue1_X = ['NA'] * len(CurveValue2_X_lay2)

                        CurveValue2_Value_lay2 = self.importer2.__local_prms__[self.variant_name2]["CURVE"][j].value
                        CurveValue3_Value = self.importer3.__local_prms__[self.variant_name3]["CURVE"][j].value
                        CurveValue1_Value = ['NA'] * len(CurveValue1_X)


                    elif p == 2 and j in self.importer3.__local_prms__[self.variant_name3]["CURVE"]:
                        CurveValue1 = self.importer1.__local_prms__[self.variant_name1]["CURVE"][j]  # Curve 파라미터명
                        CurveValue3 = self.importer3.__local_prms__[self.variant_name3]["CURVE"][j]

                        CurveValue1_X = self.importer1.__local_prms__[self.variant_name2]["COM_AXIS"][
                            CurveValue1.__xaxis__].value
                        CurveValue3_X = self.importer3.__local_prms__[self.variant_name3]["COM_AXIS"][
                            CurveValue3.__xaxis__].value
                        CurveValue2_X_lay2 = ['NA'] * len(CurveValue1_X)

                        CurveValue1_Value = self.importer1.__local_prms__[self.variant_name1]["CURVE"][j].value
                        CurveValue3_Value = self.importer3.__local_prms__[self.variant_name3]["CURVE"][j].value
                        CurveValue2_Value_lay2 = ['NA'] * len(CurveValue1_X)

                    # 2,3 모두 없음
                    elif p == 2:
                        CurveValue1 = self.importer1.__local_prms__[self.variant_name1]["CURVE"][j]  # Curve 파라미터명

                        CurveValue1_X = self.importer1.__local_prms__[self.variant_name2]["COM_AXIS"][
                            CurveValue1.__xaxis__].value
                        CurveValue3_X = ['NA'] * len(CurveValue1_X)
                        CurveValue2_X_lay2 = ['NA'] * len(CurveValue1_X)

                        CurveValue1_Value = self.importer1.__local_prms__[self.variant_name1]["CURVE"][j].value
                        CurveValue3_Value = ['NA'] * len(CurveValue1_Value)
                        CurveValue2_Value_lay2 = ['NA'] * len(CurveValue1_Value)

                    else:
                        CurveValue1 = self.importer1.__local_prms__[self.variant_name1]["CURVE"][j]  # Curve 파라미터명
                        CurveValue2_lay2 = self.importer2.__local_prms__[self.variant_name2]["CURVE"][j]

                        CurveValue1_X = self.importer1.__local_prms__[self.variant_name1]["COM_AXIS"][CurveValue1.__xaxis__].value
                        CurveValue2_X_lay2 = self.importer2.__local_prms__[self.variant_name2]["COM_AXIS"][CurveValue2_lay2.__xaxis__].value
                        CurveValue3_X = ['NA'] * len(CurveValue1_X)

                        CurveValue1_Value = self.importer1.__local_prms__[self.variant_name1]["CURVE"][j].value
                        CurveValue2_Value_lay2 = self.importer2.__local_prms__[self.variant_name2]["CURVE"][j].value
                        CurveValue3_Value = ['NA'] * len(CurveValue1_X)


                if CurIndex >= 2:
                    StartPoint1_Y = StartPoint1_Y + 4
                else:
                    StartPoint1_Y = StartPoint1_Y

                worksheet.write(StartPoint1_Y + 2, StartPoint1_x - 1, j, WriteTitleFormat)

                leng = max(len(CurveValue1_X), len(CurveValue2_X_lay2), len(CurveValue3_X))

                if len(CurveValue1_X) < leng:
                    CurveValue1_X = CurveValue2_X_lay2.tolist()
                    CurveValue1_Value = CurveValue2_Value_lay2.tolist()
                    for i in range(leng - len(CurveValue1_X)):
                        CurveValue1_X.append('NA')
                        CurveValue1_Value.append('NA')
                if len(CurveValue2_X_lay2) < leng:
                    CurveValue2_X_lay2 = CurveValue2_X_lay2.tolist()
                    CurveValue2_Value_lay2 = CurveValue2_Value_lay2.tolist()
                    for i in range(leng - len(CurveValue2_X_lay2)):
                        CurveValue2_X_lay2.append('NA')
                        CurveValue2_Value_lay2.append('NA')
                if len(CurveValue3_X) < leng:
                    CurveValue3_X = CurveValue2_X_lay2.tolist()
                    CurveValue3_Value = CurveValue2_Value_lay2.tolist()
                    for i in range(leng - len(CurveValue3_X)):
                        CurveValue3_X.append('NA')
                        CurveValue3_Value.append('NA')

                for i in range(leng):
                    # Axis
                    worksheet.write(StartPoint1_Y + 3, StartPoint1_x + i - 1, CurveValue1_X[i], data_format_axis)
                    worksheet.write(StartPoint1_Y + 3, StartPoint1_x + i + 1 + len(CurveValue1_X) + 2,
                                    CurveValue2_X_lay2[i], data_format_axis)
                    worksheet.write(StartPoint1_Y + 3, StartPoint1_x + i + 1 + 4 + (len(CurveValue1_X) * 2),
                                    CurveValue3_X[i], data_format_axis)

                    if CurveValue1_X[i] != CurveValue2_X_lay2[i]:
                        result[0] = 1
                    else:
                        pass

                    if CurveValue1_X[i] != CurveValue3_X[i]:
                        result[1] = 1
                    else:
                        pass

                    # Curve
                    worksheet.write(StartPoint1_Y + 4, StartPoint1_x + i - 1, CurveValue1_Value[i], data_format_data)
                    worksheet.write(StartPoint1_Y + 4, StartPoint1_x + i + 1 + len(CurveValue1_X) + 2,
                                    CurveValue2_Value_lay2[i], data_format_data)
                    worksheet.write(StartPoint1_Y + 4, StartPoint1_x + i + 1 + 4 + (len(CurveValue1_X) * 2),
                                    CurveValue3_Value[i], data_format_data)

                    if CurveValue1_Value[i] != CurveValue2_Value_lay2[i]:
                        result[0] = 1
                    else:
                        pass

                    if CurveValue1_Value[i] != CurveValue3_Value[i]:
                        result[1] = 1
                    else:
                        pass

                Layer1_data_CurAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 3, StartPoint1_x - 1)
                Layer2_data_CurAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 3,
                                                                           StartPoint1_x + 2 + len(
                                                                               CurveValue1_X) + 2 - 1)
                Before_data_CurAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 3,
                                                                           StartPoint1_x + 1 + 4 + (
                                                                                   len(CurveValue1_X) * 2))

                Layer1_data_CurData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 4, StartPoint1_x - 1)
                Layer2_data_CurData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 4,
                                                                           StartPoint1_x + 2 + len(
                                                                               CurveValue1_X) + 2 - 1)
                Before_data_CurData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 4,
                                                                           StartPoint1_x + 1 + 4 + (
                                                                                   len(CurveValue1_X) * 2))

                after_data_Currange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 4,
                                                                      StartPoint1_x + 2 + len(CurveValue1_X) + 2 - 1,
                                                                      StartPoint1_Y + 4,
                                                                      StartPoint1_x + 2 + len(CurveValue1_X) + 1 + i)

                after_data_CurrAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 3,
                                                                          StartPoint1_x + 2 + len(
                                                                              CurveValue1_X) + 2 - 1,
                                                                          StartPoint1_Y + 3,
                                                                          StartPoint1_x + 2 + len(
                                                                              CurveValue1_X) + 1 + i)

                Before_data_CurrAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 3,
                                                                           StartPoint1_x + 1 + 4 + (
                                                                                   len(CurveValue1_X) * 2),
                                                                           StartPoint1_Y + 3,
                                                                           StartPoint1_x + 1 + 4 + (
                                                                                   len(CurveValue1_X) * 2) + i)

                Before_data_Currange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 4,
                                                                       StartPoint1_x + 1 + 4 + (len(CurveValue1_X) * 2),
                                                                       StartPoint1_Y + 4,
                                                                       StartPoint1_x + 1 + 4 + (
                                                                               len(CurveValue1_X) * 2) + i)

                worksheet.conditional_format(after_data_CurrAxisange, {'type': 'formula',
                                                                       'criteria': Layer1_data_CurAxis + ' <> ' + Layer2_data_CurAxis,
                                                                       'format': FormatForAfter_NotEqual})
                worksheet.conditional_format(after_data_Currange, {'type': 'formula',
                                                                   'criteria': Layer1_data_CurData + ' <> ' + Layer2_data_CurData,
                                                                   'format': FormatForAfter_NotEqual})

                worksheet.conditional_format(Before_data_CurrAxisange, {'type': 'formula',
                                                                        'criteria': Layer1_data_CurAxis + ' <> ' + Before_data_CurAxis,
                                                                        'format': FormatForAfter_Change})

                worksheet.conditional_format(Before_data_Currange, {'type': 'formula',
                                                                    'criteria': Layer1_data_CurData + ' <> ' + Before_data_CurData,
                                                                    'format': FormatForAfter_Change})

            #VAL_BLK (Value Array)

            StartPoint1_x = StartPoint1_x
            StartPoint1_Y = StartPoint1_Y + 6

            ArrayValue = self.importer1.__local_prms__[self.variant_name1]["VAL_BLK"]
            ArrayIndex = 0
            ArrayValue1_Y = 0

            for v in ArrayValue.keys():
                ArrayValue1_Y_old = ArrayValue1_Y
                ArrayIndex = ArrayIndex + 1
                p = 0

                try:
                    p = 1
                    ArrayValue1 = self.importer1.__local_prms__[self.variant_name1]["VAL_BLK"][v]
                    p = 2
                    ArrayValue2 = self.importer2.__local_prms__[self.variant_name2]["VAL_BLK"][v]
                    p = 3
                    ArrayValue3 = self.importer3.__local_prms__[self.variant_name3]["VAL_BLK"][v]

                    p = 1
                    ArrayValue1_Value = self.importer1.__local_prms__[self.variant_name1]["VAL_BLK"][v].value
                    p = 2
                    ArrayValue2_Value = self.importer2.__local_prms__[self.variant_name2]["VAL_BLK"][v].value
                    p = 3
                    ArrayValue3_Value = self.importer3.__local_prms__[self.variant_name3]["VAL_BLK"][v].value

                    ArrayValue1_X = list(range(len(ArrayValue1_Value[0])))
                    ArrayValue2_X = list(range(len(ArrayValue2_Value[0])))
                    ArrayValue3_X = list(range(len(ArrayValue3_Value[0])))

                    ArrayValue1_Y = list(range(len(ArrayValue1_Value)))
                    ArrayValue2_Y = list(range(len(ArrayValue2_Value)))
                    ArrayValue3_Y = list(range(len(ArrayValue3_Value)))

                except KeyError:
                    if p == 1:
                        ArrayValue2 = self.importer2.__local_prms__[self.variant_name2]["VAL_BLK"][v]
                        ArrayValue3 = self.importer3.__local_prms__[self.variant_name3]["VAL_BLK"][v]

                        ArrayValue2_Value = self.importer2.__local_prms__[self.variant_name2]["VAL_BLK"][v].value
                        ArrayValue3_Value = self.importer3.__local_prms__[self.variant_name3]["VAL_BLK"][v].value

                        ArrayValue2_X = list(range(len(ArrayValue2_Value[0])))
                        ArrayValue3_X = list(range(len(ArrayValue3_Value[0])))
                        ArrayValue1_X = ['NA'] * len(ArrayValue2_X)

                        ArrayValue2_Y = list(range(len(ArrayValue2_Value)))
                        ArrayValue3_Y = list(range(len(ArrayValue3_Value)))
                        ArrayValue1_Y = ['NA'] * len(ArrayValue2_Y)

                        ArrayValue1_Value = [['NA'] * len(ArrayValue1_X)] * len(ArrayValue1_Y)

                    elif p == 2:
                        ArrayValue1 = self.importer1.__local_prms__[self.variant_name1]["VAL_BLK"][v]
                        ArrayValue3 = self.importer3.__local_prms__[self.variant_name3]["VAL_BLK"][v]

                        ArrayValue1_Value = self.importer1.__local_prms__[self.variant_name1]["VAL_BLK"][v].value
                        ArrayValue3_Value = self.importer3.__local_prms__[self.variant_name3]["VAL_BLK"][v].value

                        ArrayValue1_X = list(range(len(ArrayValue1_Value[0])))
                        ArrayValue3_X = list(range(len(ArrayValue3_Value[0])))
                        ArrayValue2_X = ['NA'] * len(ArrayValue1_X)

                        ArrayValue1_Y = list(range(len(ArrayValue1_Value)))
                        ArrayValue3_Y = list(range(len(ArrayValue3_Value)))
                        ArrayValue2_Y = ['NA'] * len(ArrayValue1_Y)

                        ArrayValue2_Value = [['NA'] * len(ArrayValue2_X)] * len(ArrayValue2_Y)

                    else:
                        ArrayValue1 = self.importer1.__local_prms__[self.variant_name1]["VAL_BLK"][v]
                        ArrayValue2 = self.importer2.__local_prms__[self.variant_name2]["VAL_BLK"][v]

                        ArrayValue1_Value = self.importer1.__local_prms__[self.variant_name1]["VAL_BLK"][v].value
                        ArrayValue2_Value = self.importer2.__local_prms__[self.variant_name2]["VAL_BLK"][v].value

                        ArrayValue1_X = list(range(len(ArrayValue1_Value[0])))
                        ArrayValue2_X = list(range(len(ArrayValue2_Value[0])))
                        ArrayValue3_X = ['NA'] * len(ArrayValue1_X)

                        ArrayValue1_Y = list(range(len(ArrayValue1_Value)))
                        ArrayValue2_Y = list(range(len(ArrayValue2_Value)))
                        ArrayValue3_Y = ['NA'] * len(ArrayValue1_Y)

                        ArrayValue3_Value = [['NA'] * len(ArrayValue3_X)] * len(ArrayValue3_Y)


                if ArrayIndex >= 2:
                    StartPoint1_Y = StartPoint1_Y + len(ArrayValue1_Y_old) + 3
                else:
                    StartPoint1_Y = StartPoint1_Y

                worksheet.write(StartPoint1_Y, StartPoint1_x - 1, v, WriteTitleFormat)

                lengX = max(len(ArrayValue1_X),len(ArrayValue2_X),len(ArrayValue3_X))
                lengY = max(len(ArrayValue1_Y),len(ArrayValue2_Y),len(ArrayValue3_Y))

                #X축 크기 맞추기
                if len(ArrayValue1_X) < lengX:
                    ArrayValue1_Value = ArrayValue1_Value.tolist()
                    for i in range(lengX - len(ArrayValue1_X)):
                        ArrayValue1_X.append('NA')
                        for j in range(len(ArrayValue1_Y)):
                            for k in range(lengX - len(ArrayValue1_X)):
                                ArrayValue1_Value[j].append('NA')
                if len(ArrayValue2_X) < lengX:
                    ArrayValue2_Value = ArrayValue2_Value.tolist()
                    for i in range(lengX - len(ArrayValue2_X)):
                        ArrayValue2_X.append('NA')
                        for j in range(len(ArrayValue2_Y)):
                            for k in range(lengX - len(ArrayValue2_X)):
                                ArrayValue2_Value[j].append('NA')
                if len(ArrayValue3_X) < lengX:
                    MapValue3_Value = ArrayValue3_Value.tolist()
                    for i in range(lengX - len(ArrayValue3_X)):
                        ArrayValue3_X.append('NA')
                        for j in range(len(ArrayValue3_Y)):
                            for k in range(lengX - len(ArrayValue3_X)):
                                ArrayValue3_Value[j].append('NA')
                #Y축 크기 맞추기
                if len(ArrayValue1_Y) < lengY:
                    ArrayValue1_Value = ArrayValue1_Value.tolist()
                    for i in range(lengY - len(ArrayValue1_Y)):
                        ArrayValue1_Y.append(self, 'NA')
                        temp = ['NA'] * len(ArrayValue1_X)
                        ArrayValue1_Value.append(self, temp)
                if len(ArrayValue2_Y) < lengY:
                    ArrayValue2_Value = ArrayValue2_Value.toslit()
                    for i in range(lengY - len(ArrayValue2_Y)):
                        ArrayValue2_Y.append(self, 'NA')
                        temp = ['NA'] * len(ArrayValue2_X)
                        ArrayValue2_Value.append(self, temp)
                if len(ArrayValue3_Y) < lengY:
                    ArrayValue3_Value = ArrayValue3_Value.tolist()
                    for i in range(lengY - len(ArrayValue3_Y)):
                        ArrayValue3_Y.append(self, 'NA')
                        temp = ['NA'] * len(ArrayValue3_X)
                        ArrayValue3_Value.append(self, temp)


                for l in range(len(ArrayValue1_X)):
                    worksheet.write(StartPoint1_Y + 1, l + StartPoint1_x, ArrayValue1_X[l], data_format_axis)
                    worksheet.write(StartPoint1_Y + 1, l + StartPoint1_x + len(ArrayValue1_X) + 3, ArrayValue2_X[l],
                                    data_format_axis)
                    worksheet.write(StartPoint1_Y + 1, l + StartPoint1_x + 3 + (len(ArrayValue1_X) * 2) + 3,
                                    ArrayValue3_X[l],
                                    data_format_axis)

                    if ArrayValue1_X[l] != ArrayValue2_X[l]:
                        result[0] = 1
                    else:
                        pass
                    if ArrayValue1_X[l] != ArrayValue3_X[l]:
                        result[1] = 1
                    else:
                        pass

                Layer1_data_ArrayXAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 1, StartPoint1_x)
                Layer2_data_ArrayXAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 1,
                                                                              StartPoint1_x + len(ArrayValue1_X) + 3)
                Before_data_ArrayXAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 1,
                                                                              StartPoint1_x + 3 + (
                                                                                          len(ArrayValue1_X) * 2) + 3)

                after_data_ArrayXAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 1,
                                                                            StartPoint1_x + len(ArrayValue1_X) + 3,
                                                                            StartPoint1_Y + 1,
                                                                            StartPoint1_x + len(ArrayValue1_X) + 3 + l)

                Before_data_ArrayXAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 1,
                                                                             StartPoint1_x + 3 + (
                                                                                         len(ArrayValue1_X) * 2) + 3,
                                                                             StartPoint1_Y + 1,
                                                                             StartPoint1_x + 3 + (
                                                                                         len(ArrayValue1_X) * 2) + 3 + l)

                worksheet.conditional_format(after_data_ArrayXAxisange, {'type': 'formula',
                                                                         'criteria': Layer1_data_ArrayXAxis + ' <> ' + Layer2_data_ArrayXAxis,
                                                                         'format': FormatForAfter_NotEqual})

                worksheet.conditional_format(Before_data_ArrayXAxisange, {'type': 'formula',
                                                                          'criteria': Layer1_data_ArrayXAxis + ' <> ' + Before_data_ArrayXAxis,
                                                                          'format': FormatForAfter_Change})

                for m in range(len(ArrayValue1_Y)):
                    worksheet.write(StartPoint1_Y + m + 2, StartPoint1_x - 1, ArrayValue1_Y[m], data_format_axis)
                    worksheet.write(StartPoint1_Y + m + 2, StartPoint1_x - 1 + len(ArrayValue1_X) + 3, ArrayValue2_Y[m],
                                    data_format_axis)
                    worksheet.write(StartPoint1_Y + m + 2, StartPoint1_x - 1 + 3 + (len(ArrayValue1_X) * 2) + 3,
                                    ArrayValue3_Y[m],
                                    data_format_axis)

                    if ArrayValue1_Y[m] != ArrayValue2_Y[m]:
                        result[0] = 1
                    else:
                        pass
                    if ArrayValue1_Y[m] != ArrayValue3_Y[m]:
                        result[1] = 1
                    else:
                        pass

                Layer1_data_ArrayYAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 2, StartPoint1_x - 1)
                Layer2_data_ArrayYAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 2,
                                                                              StartPoint1_x - 1 + len(
                                                                                  ArrayValue1_X) + 3)

                Before_data_ArrayYAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 2,
                                                                              StartPoint1_x - 1 + 3 + (
                                                                                      len(ArrayValue1_X) * 2) + 3)

                after_data_ArrayYAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 2,
                                                                            StartPoint1_x - 1 + len(ArrayValue1_X) + 3,
                                                                            StartPoint1_Y + 2 + m,
                                                                            StartPoint1_x - 1 + len(ArrayValue1_X) + 3)

                Before_data_ArrayYAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 2,
                                                                             StartPoint1_x - 1 + 3 + (
                                                                                         len(ArrayValue1_X) * 2) + 3,
                                                                             StartPoint1_Y + 2 + m,
                                                                             StartPoint1_x - 1 + 3 + (
                                                                                         len(ArrayValue1_X) * 2) + 3)

                worksheet.conditional_format(after_data_ArrayYAxisange, {'type': 'formula',
                                                                         'criteria': Layer1_data_ArrayYAxis + ' <> ' + Layer2_data_ArrayYAxis,
                                                                         'format': FormatForAfter_NotEqual})

                worksheet.conditional_format(Before_data_ArrayYAxisange, {'type': 'formula',
                                                                          'criteria': Layer1_data_ArrayYAxis + ' <> ' + Before_data_ArrayYAxis,
                                                                          'format': FormatForAfter_Change})

                for p in range(len(ArrayValue1_Y)):
                    for o in range(len(ArrayValue1_X)):
                        worksheet.write(StartPoint1_Y + 2 + p, o + StartPoint1_x, ArrayValue1_Value[p][o],
                                        data_format_data)
                        worksheet.write(StartPoint1_Y + 2 + p, o + StartPoint1_x + len(ArrayValue1_X) + 3,
                                        ArrayValue1_Value[p][o],
                                        data_format_data)
                        worksheet.write(StartPoint1_Y + 2 + p, o + StartPoint1_x + (len(ArrayValue1_X) * 2) + 6,
                                        ArrayValue3_Value[p][o], data_format_data)

                        if ArrayValue1_Value[p][o] != ArrayValue1_Value[p][o]:
                            result[0] = 1
                        else:
                            pass

                        if ArrayValue1_Value[p][o] != ArrayValue3_Value[p][o]:
                            result[1] = 1
                        else:
                            pass

                Layer1_data_ArrayData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 2, StartPoint1_x)
                Layer2_data_ArrayData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 2,
                                                                             StartPoint1_x + len(ArrayValue1_X) + 3)
                Before_data_ArrayData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 2,
                                                                             StartPoint1_x + (
                                                                                         len(ArrayValue1_X) * 2) + 6)

                after_data_ArrayDatasange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 2,
                                                                            StartPoint1_x + len(ArrayValue1_X) + 3,
                                                                            StartPoint1_Y + 2 + p,
                                                                            o + StartPoint1_x + len(ArrayValue1_X) + 3)

                Before_data_ArrayDatasange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 2,
                                                                             StartPoint1_x + (
                                                                                         len(ArrayValue1_X) * 2) + 6,
                                                                             StartPoint1_Y + 2 + p,
                                                                             o + StartPoint1_x + (
                                                                                         len(ArrayValue1_X) * 2) + 6)

                worksheet.conditional_format(after_data_ArrayDatasange, {'type': 'formula',
                                                                         'criteria': Layer1_data_ArrayData + ' <> ' + Layer2_data_ArrayData,
                                                                         'format': FormatForAfter_NotEqual})

                worksheet.conditional_format(Before_data_ArrayDatasange, {'type': 'formula',
                                                                          'criteria': Layer1_data_ArrayData + ' <> ' + Before_data_ArrayData,
                                                                          'format': FormatForAfter_Change})

            # MAP Lay1/2
            StartPoint1_x = StartPoint1_x
            StartPoint1_Y = StartPoint1_Y + 6

            MapValue = self.importer1.__local_prms__[self.variant_name1]["MAP"]
            MapValue_lay2 = self.importer2.__local_prms__[self.variant_name2]["MAP"]
            MapIndex = 0
            MapValue1_Y = 0
            for k in MapValue.keys():
                MapValue1_y_old = MapValue1_Y
                MapIndex = MapIndex + 1
                p=0
                try:
                    p = 1
                    MapValue1 = self.importer1.__local_prms__[self.variant_name1]["MAP"][k]
                    p = 2
                    MapValue2_lay2 = self.importer2.__local_prms__[self.variant_name2]["MAP"][k]
                    p = 3
                    MapValue3 = self.importer3.__local_prms__[self.variant_name3]["MAP"][k]

                    p = 1
                    MapValue1_X = self.importer1.__local_prms__[self.variant_name1]["COM_AXIS"][
                        MapValue1.__xaxis__].value
                    p = 2
                    MapValue2_X_lay2 = self.importer2.__local_prms__[self.variant_name2]["COM_AXIS"][
                        MapValue2_lay2.__xaxis__].value
                    p = 3
                    MapValue3_X = self.importer3.__local_prms__[self.variant_name3]["COM_AXIS"][
                        MapValue3.__xaxis__].value

                    p = 1
                    MapValue1_Y = self.importer1.__local_prms__[self.variant_name1]["COM_AXIS"][
                        MapValue1.__yaxis__].value
                    p = 2
                    MapValue2_Y_lay2 = self.importer2.__local_prms__[self.variant_name2]["COM_AXIS"][
                        MapValue2_lay2.__yaxis__].value
                    p = 3
                    MapValue3_Y = self.importer3.__local_prms__[self.variant_name3]["COM_AXIS"][
                        MapValue3.__yaxis__].value

                    p = 1
                    MapValue1_Value = self.importer1.__local_prms__[self.variant_name1]["MAP"][k].value
                    p = 2
                    MapValue1_Value_lay2 = self.importer2.__local_prms__[self.variant_name2]["MAP"][k].value
                    p = 3
                    MapValue3_Value = self.importer3.__local_prms__[self.variant_name3]["MAP"][k].value

                except KeyError:

                    if p == 1:
                        MapValue2_lay2 = self.importer2.__local_prms__[self.variant_name2]["MAP"][k]
                        MapValue3 = self.importer3.__local_prms__[self.variant_name3]["MAP"][k]

                        MapValue2_X_lay2 = self.importer2.__local_prms__[self.variant_name2]["COM_AXIS"][
                            MapValue2_lay2.__xaxis__].value
                        MapValue3_X = self.importer3.__local_prms__[self.variant_name3]["COM_AXIS"][
                            MapValue3.__xaxis__].value
                        MapValue1_X = ['NA']*len(MapValue2_X_lay2)

                        MapValue2_Y_lay2 = self.importer2.__local_prms__[self.variant_name2]["COM_AXIS"][
                            MapValue2_lay2.__yaxis__].value
                        MapValue3_Y = self.importer3.__local_prms__[self.variant_name3]["COM_AXIS"][
                            MapValue3.__yaxis__].value
                        MapValue1_Y = ['NA'] * len(MapValue2_Y_lay2)

                        MapValue1_Value_lay2 = self.importer2.__local_prms__[self.variant_name2]["MAP"][k].value
                        MapValue3_Value = self.importer3.__local_prms__[self.variant_name3]["MAP"][k].value
                        MapValue1_Value = [['NA'] * len(MapValue2_X_lay2)] * len(MapValue2_Y_lay2)

                    elif p == 2:
                        MapValue1 = self.importer1.__local_prms__[self.variant_name1]["MAP"][k]
                        MapValue3 = self.importer3.__local_prms__[self.variant_name3]["MAP"][k]

                        MapValue1_X = self.importer1.__local_prms__[self.variant_name1]["COM_AXIS"][
                            MapValue1.__xaxis__].value
                        MapValue3_X = self.importer3.__local_prms__[self.variant_name3]["COM_AXIS"][
                            MapValue3.__xaxis__].value
                        MapValue2_X_lay2 = ['NA']*len(MapValue1_X)

                        MapValue1_Y = self.importer1.__local_prms__[self.variant_name1]["COM_AXIS"][
                            MapValue1.__yaxis__].value
                        MapValue3_Y = self.importer3.__local_prms__[self.variant_name3]["COM_AXIS"][
                            MapValue3.__yaxis__].value
                        MapValue2_Y_lay2 = ['NA'] * len(MapValue1_Y)

                        MapValue1_Value = self.importer1.__local_prms__[self.variant_name1]["MAP"][k].value
                        MapValue3_Value = self.importer3.__local_prms__[self.variant_name3]["MAP"][k].value
                        MapValue1_Value_lay2 = [['NA'] * len(MapValue1_X)] * len(MapValue1_Y)

                    else:
                        MapValue1 = self.importer1.__local_prms__[self.variant_name1]["MAP"][k]
                        MapValue2_lay2 = self.importer2.__local_prms__[self.variant_name2]["MAP"][k]

                        MapValue1_X = self.importer1.__local_prms__[self.variant_name1]["COM_AXIS"][
                            MapValue1.__xaxis__].value
                        MapValue2_X_lay2 = self.importer2.__local_prms__[self.variant_name2]["COM_AXIS"][
                            MapValue2_lay2.__xaxis__].value
                        MapValue3_X = ['NA']*len(MapValue1_X)

                        MapValue1_Y = self.importer1.__local_prms__[self.variant_name1]["COM_AXIS"][
                            MapValue1.__yaxis__].value
                        MapValue2_Y_lay2 = self.importer2.__local_prms__[self.variant_name2]["COM_AXIS"][
                            MapValue2_lay2.__yaxis__].value
                        MapValue3_Y = ['NA'] * len(MapValue1_Y)

                        MapValue1_Value = self.importer1.__local_prms__[self.variant_name1]["MAP"][k].value
                        MapValue1_Value_lay2 = self.importer2.__local_prms__[self.variant_name2]["MAP"][k].value
                        MapValue3_Value = [['NA'] * len(MapValue1_X)] * len(MapValue1_Y)

                if MapIndex >= 2:
                    StartPoint1_Y = StartPoint1_Y+len(MapValue1_y_old)+3
                else:
                    StartPoint1_Y = StartPoint1_Y

                worksheet.write(StartPoint1_Y, StartPoint1_x - 1, k, WriteTitleFormat)

                lengX = max(len(MapValue1_X), len(MapValue2_X_lay2), len(MapValue3_X))
                lengY = max(len(MapValue1_Y), len(MapValue2_Y_lay2), len(MapValue3_Y))

                #X축 크기 맞추기
                if len(MapValue1_X) < lengX:
                    MapValue1_X = MapValue1_X.tolist()
                    MapValue1_Value = MapValue1_Value.tolist()
                    for i in range(lengX-len(MapValue1_X)):
                        MapValue1_X.append('NA')
                        for j in range(len(MapValue1_Y)):
                            for k in range(lengX-len(MapValue1_X)):
                                MapValue1_Value[j].append('NA')
                if len(MapValue2_X_lay2) < lengX:
                    MapValue2_X_lay2 = MapValue2_X_lay2.tolist()
                    MapValue1_Value_lay2 = MapValue1_Value_lay2.tolist()
                    for i in range(lengX-len(MapValue2_X_lay2)):
                        MapValue2_X_lay2.append('NA')
                        for j in range(len(MapValue2_Y_lay2)):
                            for k in range(lengX-len(MapValue2_X_lay2)):
                                MapValue1_Value_lay2[j].append('NA')
                if len(MapValue3_X) < lengX:
                    MapValue3_X = MapValue3_X.tolist()
                    MapValue3_Value = MapValue3_Value.tolist()
                    for i in range(lengX-len(MapValue3_X)):
                        MapValue3_X.append('NA')
                        for j in range(len(MapValue3_Y)):
                            for k in range(lengX-len(MapValue3_X)):
                                MapValue3_Value[j].append('NA')


                #Y축 크기 맞추기
                if len(MapValue1_Y) < lengY:
                    MapValue1_Y = MapValue1_Y.tolist()
                    MapValue1_Value = MapValue1_Value.tolist()
                    for i in range(lengY - len(MapValue1_Y)):
                        MapValue1_Y.append(self, 'NA')
                        temp = ['NA'] * len(MapValue1_X)
                        MapValue1_Value.append(self, temp)
                if len(MapValue2_Y_lay2) < lengY:
                    MapValue2_Y_lay2 = MapValue2_Y_lay2.tolist()
                    MapValue1_Value_lay2 = MapValue1_Value_lay2.toslit()
                    for i in range(lengY - len(MapValue2_Y_lay2)):
                        MapValue2_Y_lay2.append(self, 'NA')
                        temp = ['NA'] * len(MapValue2_X_lay2)
                        MapValue1_Value_lay2.append(self, temp)
                if len(MapValue3_Y) < lengY:
                    MapValue3_Y = MapValue3_Y.tolist()
                    MapValue3_Value = MapValue3_Value.tolist()
                    for i in range(lengY - len(MapValue3_Y)):
                        MapValue3_Y.append(self, 'NA')
                        temp = ['NA'] * len(MapValue3_X)
                        MapValue3_Value.append(self, temp)

                for l in range(lengX):
                    worksheet.write(StartPoint1_Y+1, l + StartPoint1_x, MapValue1_X[l], data_format_axis)
                    worksheet.write(StartPoint1_Y+1, l + StartPoint1_x + len(MapValue1_X)+3, MapValue2_X_lay2[l],data_format_axis)
                    worksheet.write(StartPoint1_Y + 1, l + StartPoint1_x + 3 + (len(MapValue1_X) * 2) + 3, MapValue3_X[l],
                                    data_format_axis)

                    if MapValue1_X[l] != MapValue2_X_lay2[l]:
                        result[0] = 1
                    else:
                        pass
                    if MapValue1_X[l] != MapValue3_X[l]:
                        result[1] = 1
                    else:
                        pass


                Layer1_data_MapXAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 1, StartPoint1_x)
                Layer2_data_MapXAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 1,
                                                                            StartPoint1_x + len(MapValue1_X) + 3)
                Before_data_MapXAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 1,
                                                                            StartPoint1_x +3+(len(MapValue1_X)*2)+3)


                after_data_MapXAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 1,
                                                                          StartPoint1_x + len(MapValue1_X) + 3,
                                                                          StartPoint1_Y + 1,
                                                                          StartPoint1_x + len(MapValue1_X) + 3 + l)

                Before_data_MapXAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 1,
                                                                          StartPoint1_x +3+(len(MapValue1_X)*2)+3,
                                                                          StartPoint1_Y + 1,
                                                                          StartPoint1_x +3+(len(MapValue1_X)*2)+3 + l)


                worksheet.conditional_format(after_data_MapXAxisange, {'type': 'formula',
                                                                       'criteria': Layer1_data_MapXAxis + ' <> ' + Layer2_data_MapXAxis,
                                                                         'format': FormatForAfter_NotEqual})

                worksheet.conditional_format(Before_data_MapXAxisange, {'type': 'formula',
                                                                       'criteria': Layer1_data_MapXAxis + ' <> ' + Before_data_MapXAxis,
                                                                       'format': FormatForAfter_Change})


                for m in range(len(MapValue1_Y)):
                    worksheet.write(StartPoint1_Y+m+2,StartPoint1_x-1,MapValue1_Y[m],data_format_axis)
                    worksheet.write(StartPoint1_Y + m + 2, StartPoint1_x - 1+len(MapValue1_X)+3, MapValue2_Y_lay2[m], data_format_axis)
                    worksheet.write(StartPoint1_Y + m + 2, StartPoint1_x - 1+3+(len(MapValue1_X)*2) + 3, MapValue3_Y[m], data_format_axis)

                    if MapValue1_Y[m] != MapValue2_Y_lay2[m]:
                        result[0] = 1
                    else:
                        pass
                    if MapValue1_Y[m] != MapValue3_Y[m]:
                        result[1] = 1
                    else:
                        pass

                Layer1_data_MapYAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y+2,StartPoint1_x-1)
                Layer2_data_MapYAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y+2,
                                                                            StartPoint1_x - 1+len(MapValue1_X)+3)

                Before_data_MapYAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 2,
                                                                            StartPoint1_x - 1 + 3 + (
                                                                                        len(MapValue1_X) * 2) + 3)

                after_data_MapYAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y+2,
                                                                          StartPoint1_x - 1+len(MapValue1_X)+3,
                                                                          StartPoint1_Y+2+m,
                                                                          StartPoint1_x - 1+len(MapValue1_X)+3)

                Before_data_MapYAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y+2,
                                                                          StartPoint1_x - 1+3+(len(MapValue1_X)*2) + 3,
                                                                          StartPoint1_Y+2+m,
                                                                          StartPoint1_x - 1+3+(len(MapValue1_X)*2) + 3)


                worksheet.conditional_format(after_data_MapYAxisange, {'type': 'formula',
                                                                       'criteria': Layer1_data_MapYAxis + ' <> ' + Layer2_data_MapYAxis,
                                                                       'format': FormatForAfter_NotEqual})

                worksheet.conditional_format(Before_data_MapYAxisange, {'type': 'formula',
                                                                       'criteria': Layer1_data_MapYAxis + ' <> ' + Before_data_MapYAxis,
                                                                       'format': FormatForAfter_Change})

                for p in range(len(MapValue1_Y)):
                    for o in range(len(MapValue1_X)):
                        worksheet.write(StartPoint1_Y+2+p, o+StartPoint1_x,MapValue1_Value[p][o],data_format_data)
                        worksheet.write(StartPoint1_Y+2+p, o+StartPoint1_x+len(MapValue1_X)+3, MapValue1_Value_lay2[p][o], data_format_data)
                        worksheet.write(StartPoint1_Y+2+p, o + StartPoint1_x + (len(MapValue1_X)*2) + 6, MapValue3_Value[p][o], data_format_data)

                        if MapValue1_Value[p][o] != MapValue1_Value_lay2[p][o]:
                            result[0] = 1
                        else:
                            pass

                        if MapValue1_Value[p][o] != MapValue3_Value[p][o]:
                            result[1] = 1
                        else:
                            pass

                Layer1_data_MapData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y+2,StartPoint1_x)
                Layer2_data_MapData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y+2,
                                                                            StartPoint1_x+len(MapValue1_X)+3)
                Before_data_MapData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 2,
                                                                           StartPoint1_x + (len(MapValue1_X)*2) + 6)

                after_data_MapDatasange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y+2,
                                                                          StartPoint1_x+len(MapValue1_X)+3,
                                                                          StartPoint1_Y+2+p,
                                                                          o+StartPoint1_x+len(MapValue1_X)+3)

                Before_data_MapDatasange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 2,
                                                                          StartPoint1_x + (len(MapValue1_X)*2) + 6,
                                                                          StartPoint1_Y + 2 + p,
                                                                          o + StartPoint1_x + (len(MapValue1_X)*2) + 6)

                worksheet.conditional_format(after_data_MapDatasange, {'type': 'formula',
                                                                       'criteria': Layer1_data_MapData + ' <> ' + Layer2_data_MapData,
                                                                       'format': FormatForAfter_NotEqual})

                worksheet.conditional_format(Before_data_MapDatasange, {'type': 'formula',
                                                                        'criteria': Layer1_data_MapData + ' <> ' + Before_data_MapData,
                                                                        'format': FormatForAfter_Change})

        except KeyError:
            result[0] = 2
            result[1] = 2

        StartPoint1_x = StartPoint1_x
        StartPoint1_Y = StartPoint1_Y + 6
        result.append(StartPoint1_x)
        result.append(StartPoint1_Y)

        return result

    def __drawMapSingle2__(self,start_x,start_y):
        workbook = self.workbook
        worksheet = self.current_sheet
        worksheet.set_column('A:FH', 5)
        StartPoint1_x = start_x
        StartPoint1_Y = start_y
        result = [0, 0]

        data_format_data = workbook.add_format({'border': 1})
        data_format_axis = workbook.add_format({'border': 1, 'bg_color': '#DCE6F1'})
        WriteTitleFormat = workbook.add_format({'font_color': '#000000', 'bold': True, 'font_size': 11})
        WrongFormat = workbook.add_format({'bg_color': '#FF0000'})

        # 조건부 서식 Format
        FormatForAfter_NotEqual = workbook.add_format({'bold': True, 'bg_color': 'red'})
        FormatForAfter_Change = workbook.add_format({'bold': True, 'bg_color': 'gray'})
        # FormatForAfter_NotEqual = workbook.add_format({'font_color': '#000000', 'bold': True})
        FormatForAfter_Greater = workbook.add_format({'font_color': '#000000', 'bg_color': '#F2DCDB', 'bold': True})
        FormatForAfter_Equal = workbook.add_format({'font_color': '#000000', 'bg_color': '#FFFFFF'})
        FormatForAfter_AxisEqual = workbook.add_format({'font_color': '#000000', 'bg_color': '#DCE6F1'})

        # VALUE 챠트 Layer1/Layer2
        try:
            ConstValue1 = self.importer1.__global_prms__[self.variant_name1]["VALUE"]
            Nr = len(ConstValue1)
            worksheet.write(StartPoint1_Y, StartPoint1_x + 5, "New", WriteTitleFormat)
            worksheet.write(StartPoint1_Y, StartPoint1_x + 6, "Tuning", WriteTitleFormat)
            worksheet.write(StartPoint1_Y, StartPoint1_x + 7, "Before", WriteTitleFormat)
            index = 0
            for i in ConstValue1.keys():
                p = 0
                try:
                    index = index + 1
                    p = 1
                    VariableValue1 = self.importer1.__global_prms__[self.variant_name1]["VALUE"][i].value
                    p = 2
                    VariableValue2 = self.importer2.__global_prms__[self.variant_name2]["VALUE"][i].value

                    if VariableValue1 != VariableValue2:
                        result[0] = 1

                    p = 3
                    VariableValue3 = self.importer3.__global_prms__[self.variant_name3]["VALUE"][i].value

                    if VariableValue1 != VariableValue3:
                        result[1] = 1


                except KeyError:

                    if p == 1:
                        VariableValue2 = self.importer2.__global_prms__[self.variant_name2]["VALUE"][i].value
                        VariableValue3 = self.importer3.__global_prms__[self.variant_name3]["VALUE"][i].value
                        VariableValue1 = 'NA'
                        if VariableValue1 != VariableValue2[0]:
                            result[0] = 1
                        if VariableValue1 != VariableValue3[0]:
                            result[1] = 1
                    elif p == 2:
                        VariableValue1 = self.importer1.__global_prms__[self.variant_name1]["VALUE"][i].value
                        VariableValue3 = self.importer3.__global_prms__[self.variant_name3]["VALUE"][i].value
                        VariableValue2 = 'NA'
                        if VariableValue1[0] != VariableValue2:
                            result[0] = 1
                        if VariableValue1 != VariableValue3:
                            result[1] = 1
                    else:
                        VariableValue1 = self.importer1.__global_prms__[self.variant_name1]["VALUE"][i].value
                        VariableValue2 = self.importer2.__global_prms__[self.variant_name2]["VALUE"][i].value
                        VariableValue3 = 'NA'
                        if VariableValue1[0] != VariableValue3:
                            result[1] = 1

                worksheet.write(StartPoint1_Y + index, StartPoint1_x - 1, i, WriteTitleFormat)
                worksheet.write(StartPoint1_Y + index, StartPoint1_x + 5, VariableValue1, data_format_data)
                worksheet.write(StartPoint1_Y + index, StartPoint1_x + 6, VariableValue2, data_format_data)
                worksheet.write(StartPoint1_Y + index, StartPoint1_x + 7, VariableValue3, data_format_data)

            Layer1_data_sp = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 1, StartPoint1_x + 5)
            Layer2_data_sp = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 1, StartPoint1_x + 6)
            Layer3_data_sp = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 1, StartPoint1_x + 7)

            after_data_range = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 1, StartPoint1_x + 6, StartPoint1_Y + Nr,
                                                               StartPoint1_x + 6)
            after_data_range1 = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 1, StartPoint1_x + 7,
                                                                StartPoint1_Y + Nr,
                                                                StartPoint1_x + 7)

            worksheet.conditional_format(after_data_range, {'type': 'formula',
                                                            'criteria': Layer1_data_sp + ' <> ' + Layer2_data_sp,
                                                            'format': FormatForAfter_NotEqual})

            worksheet.conditional_format(after_data_range1, {'type': 'formula',
                                                             'criteria': Layer1_data_sp + ' <> ' + Layer3_data_sp,
                                                             'format': FormatForAfter_Change})

            # Curve Layer1/Layer2

            StartPoint1_x = StartPoint1_x
            StartPoint1_Y = StartPoint1_Y + Nr

            CurValue = self.importer1.__global_prms__[self.variant_name1]["CURVE"]
            CurValue_lay2 = self.importer2.__global_prms__[self.variant_name2]["CURVE"]
            CurIndex = 0

            for j in CurValue.keys():
                CurIndex = CurIndex + 1
                p = 0

                try:
                    p = 1
                    CurveValue1 = self.importer1.__global_prms__[self.variant_name1]["CURVE"][j]  # Curve 파라미터명
                    p = 2
                    CurveValue2_lay2 = self.importer2.__global_prms__[self.variant_name2]["CURVE"][j]

                    p = 3
                    CurveValue3 = self.importer3.__global_prms__[self.variant_name3]["CURVE"][j]

                    p = 1
                    CurveValue1_X = self.importer1.__global_prms__[self.variant_name1]["COM_AXIS"][
                        CurveValue1.__xaxis__].value
                    # Curve 파라미터의 x축
                    p = 2
                    CurveValue2_X_lay2 = self.importer2.__global_prms__[self.variant_name2]["COM_AXIS"][
                        CurveValue2_lay2.__xaxis__].value
                    p = 3
                    CurveValue3_X = self.importer3.__global_prms__[self.variant_name3]["COM_AXIS"][
                        CurveValue3.__xaxis__].value

                    CurveValue1_Value = self.importer1.__global_prms__[self.variant_name1]["CURVE"][j].value
                    # Curve 파라미터의 값 (비교대상)
                    CurveValue2_Value_lay2 = self.importer2.__global_prms__[self.variant_name2]["CURVE"][j].value
                    CurveValue3_Value = self.importer3.__global_prms__[self.variant_name3]["CURVE"][j].value


                except KeyError:
                    if p == 1:
                        CurveValue2_lay2 = self.importer2.__global_prms__[self.variant_name2]["CURVE"][j]
                        CurveValue3 = self.importer3.__global_prms__[self.variant_name3]["CURVE"][j]

                        CurveValue2_X_lay2 = self.importer2.__global_prms__[self.variant_name2]["COM_AXIS"][
                            CurveValue2_lay2.__xaxis__].value
                        CurveValue3_X = self.importer3.__global_prms__[self.variant_name3]["COM_AXIS"][
                            CurveValue3.__xaxis__].value
                        CurveValue1_X = ['NA'] * len(CurveValue2_X_lay2)

                        CurveValue2_Value_lay2 = self.importer2.__global_prms__[self.variant_name2]["CURVE"][j].value
                        CurveValue3_Value = self.importer3.__global_prms__[self.variant_name3]["CURVE"][j].value
                        CurveValue1_Value = ['NA'] * len(CurveValue1_X)

                    # 3은 있고 2번만 해당 튜닝값 없음
                    elif p == 2 and j in self.importer3.__global_prms__[self.variant_name3]["CURVE"]:
                        CurveValue1 = self.importer1.__global_prms__[self.variant_name1]["CURVE"][j]  # Curve 파라미터명
                        CurveValue3 = self.importer3.__global_prms__[self.variant_name3]["CURVE"][j]

                        CurveValue1_X = self.importer1.__global_prms__[self.variant_name2]["COM_AXIS"][
                            CurveValue1.__xaxis__].value
                        CurveValue3_X = self.importer3.__global_prms__[self.variant_name3]["COM_AXIS"][
                            CurveValue3.__xaxis__].value
                        CurveValue2_X_lay2 = ['NA'] * len(CurveValue1_X)

                        CurveValue1_Value = self.importer1.__global_prms__[self.variant_name1]["CURVE"][j].value
                        CurveValue3_Value = self.importer3.__global_prms__[self.variant_name3]["CURVE"][j].value
                        CurveValue2_Value_lay2 = ['NA'] * len(CurveValue1_X)

                    # 2,3 모두 없음
                    elif p == 2:
                        CurveValue1 = self.importer1.__global_prms__[self.variant_name1]["CURVE"][j]  # Curve 파라미터명

                        CurveValue1_X = self.importer1.__global_prms__[self.variant_name2]["COM_AXIS"][
                            CurveValue1.__xaxis__].value
                        CurveValue3_X = ['NA'] * len(CurveValue1_X)
                        CurveValue2_X_lay2 = ['NA'] * len(CurveValue1_X)

                        CurveValue1_Value = self.importer1.__global_prms__[self.variant_name1]["CURVE"][j].value
                        CurveValue3_Value = ['NA'] * len(CurveValue1_X)
                        CurveValue2_Value_lay2 = ['NA'] * len(CurveValue1_X)

                    else:
                        CurveValue1 = self.importer1.__global_prms__[self.variant_name1]["CURVE"][j]  # Curve 파라미터명
                        CurveValue2_lay2 = self.importer2.__global_prms__[self.variant_name2]["CURVE"][j]

                        CurveValue1_X = self.importer1.__global_prms__[self.variant_name1]["COM_AXIS"][
                            CurveValue1.__xaxis__].value
                        CurveValue2_X_lay2 = self.importer2.__global_prms__[self.variant_name2]["COM_AXIS"][
                            CurveValue2_lay2.__xaxis__].value
                        CurveValue3_X = ['NA'] * len(CurveValue1_X)

                        CurveValue1_Value = self.importer1.__global_prms__[self.variant_name1]["CURVE"][j].value
                        CurveValue2_Value_lay2 = self.importer2.__global_prms__[self.variant_name2]["CURVE"][j].value
                        CurveValue3_Value = ['NA'] * len(CurveValue1_X)

                if CurIndex >= 2:
                    StartPoint1_Y = StartPoint1_Y + 4
                else:
                    StartPoint1_Y = StartPoint1_Y

                worksheet.write(StartPoint1_Y + 2, StartPoint1_x - 1, j, WriteTitleFormat)

                leng = max(len(CurveValue1_X), len(CurveValue2_X_lay2), len(CurveValue3_X))

                if len(CurveValue1_X) < leng:
                    CurveValue1_X = CurveValue2_X_lay2.tolist()
                    CurveValue1_Value = CurveValue2_Value_lay2.tolist()
                    for i in range(leng - len(CurveValue1_X)):
                        CurveValue1_X.append('NA')
                        CurveValue1_Value.append('NA')
                if len(CurveValue2_X_lay2) < leng:
                    CurveValue2_X_lay2 = CurveValue2_X_lay2.tolist()
                    CurveValue2_Value_lay2 = CurveValue2_Value_lay2.tolist()
                    for i in range(leng - len(CurveValue2_X_lay2)):
                        CurveValue2_X_lay2.append('NA')
                        CurveValue2_Value_lay2.append('NA')
                if len(CurveValue3_X) < leng:
                    CurveValue3_X = CurveValue2_X_lay2.tolist()
                    CurveValue3_Value = CurveValue2_Value_lay2.tolist()
                    for i in range(leng - len(CurveValue3_X)):
                        CurveValue3_X.append('NA')
                        CurveValue3_Value.append('NA')

                for i in range(leng):
                    # Axis
                    worksheet.write(StartPoint1_Y + 3, StartPoint1_x + i - 1, CurveValue1_X[i], data_format_axis)
                    worksheet.write(StartPoint1_Y + 3, StartPoint1_x + i + 1 + len(CurveValue1_X) + 2,
                                    CurveValue2_X_lay2[i], data_format_axis)
                    worksheet.write(StartPoint1_Y + 3, StartPoint1_x + i + 1 + 4 + (len(CurveValue1_X) * 2),
                                    CurveValue3_X[i], data_format_axis)

                    if CurveValue1_X[i] != CurveValue2_X_lay2[i]:
                        result[0] = 1
                    else:
                        pass

                    if CurveValue1_X[i] != CurveValue3_X[i]:
                        result[1] = 1
                    else:
                        pass

                    # Curve
                    worksheet.write(StartPoint1_Y + 4, StartPoint1_x + i - 1, CurveValue1_Value[i], data_format_data)
                    worksheet.write(StartPoint1_Y + 4, StartPoint1_x + i + 1 + len(CurveValue1_X) + 2,
                                    CurveValue2_Value_lay2[i], data_format_data)
                    worksheet.write(StartPoint1_Y + 4, StartPoint1_x + i + 1 + 4 + (len(CurveValue1_X) * 2),
                                    CurveValue3_Value[i], data_format_data)

                    if CurveValue1_Value[i] != CurveValue2_Value_lay2[i]:
                        result[0] = 1
                    else:
                        pass

                    if CurveValue1_Value[i] != CurveValue3_Value[i]:
                        result[1] = 1
                    else:
                        pass

                Layer1_data_CurAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 3, StartPoint1_x - 1)
                Layer2_data_CurAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 3,
                                                                           StartPoint1_x + 2 + len(
                                                                               CurveValue1_X) + 2 - 1)
                Before_data_CurAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 3,
                                                                           StartPoint1_x + 1 + 4 + (
                                                                                   len(CurveValue1_X) * 2))

                Layer1_data_CurData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 4, StartPoint1_x - 1)
                Layer2_data_CurData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 4,
                                                                           StartPoint1_x + 2 + len(
                                                                               CurveValue1_X) + 2 - 1)
                Before_data_CurData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 4,
                                                                           StartPoint1_x + 1 + 4 + (
                                                                                   len(CurveValue1_X) * 2))

                after_data_Currange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 4,
                                                                      StartPoint1_x + 2 + len(CurveValue1_X) + 2 - 1,
                                                                      StartPoint1_Y + 4,
                                                                      StartPoint1_x + 2 + len(CurveValue1_X) + 1 + i)

                after_data_CurrAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 3,
                                                                          StartPoint1_x + 2 + len(
                                                                              CurveValue1_X) + 2 - 1,
                                                                          StartPoint1_Y + 3,
                                                                          StartPoint1_x + 2 + len(
                                                                              CurveValue1_X) + 1 + i)

                Before_data_CurrAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 3,
                                                                           StartPoint1_x + 1 + 4 + (
                                                                                   len(CurveValue1_X) * 2),
                                                                           StartPoint1_Y + 3,
                                                                           StartPoint1_x + 1 + 4 + (
                                                                                   len(CurveValue1_X) * 2) + i)

                Before_data_Currange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 4,
                                                                       StartPoint1_x + 1 + 4 + (len(CurveValue1_X) * 2),
                                                                       StartPoint1_Y + 4,
                                                                       StartPoint1_x + 1 + 4 + (
                                                                               len(CurveValue1_X) * 2) + i)

                worksheet.conditional_format(after_data_CurrAxisange, {'type': 'formula',
                                                                       'criteria': Layer1_data_CurAxis + ' <> ' + Layer2_data_CurAxis,
                                                                       'format': FormatForAfter_NotEqual})
                worksheet.conditional_format(after_data_Currange, {'type': 'formula',
                                                                   'criteria': Layer1_data_CurData + ' <> ' + Layer2_data_CurData,
                                                                   'format': FormatForAfter_NotEqual})

                worksheet.conditional_format(Before_data_CurrAxisange, {'type': 'formula',
                                                                        'criteria': Layer1_data_CurAxis + ' <> ' + Before_data_CurAxis,
                                                                        'format': FormatForAfter_Change})

                worksheet.conditional_format(Before_data_Currange, {'type': 'formula',
                                                                    'criteria': Layer1_data_CurData + ' <> ' + Before_data_CurData,
                                                                    'format': FormatForAfter_Change})

            # VAL_BLK (Value Array)

            StartPoint1_x = StartPoint1_x
            StartPoint1_Y = StartPoint1_Y + 6

            ArrayValue = self.importer1.__global_prms__[self.variant_name1]["VAL_BLK"]
            ArrayIndex = 0
            ArrayValue1_Y = 0

            for v in ArrayValue.keys():
                ArrayValue1_Y_old = ArrayValue1_Y
                ArrayIndex = ArrayIndex + 1
                p = 0

                try:
                    p = 1
                    ArrayValue1 = self.importer1.__global_prms__[self.variant_name1]["VAL_BLK"][v]
                    p = 2
                    ArrayValue2 = self.importer2.__global_prms__[self.variant_name2]["VAL_BLK"][v]
                    p = 3
                    ArrayValue3 = self.importer3.__global_prms__[self.variant_name3]["VAL_BLK"][v]

                    p = 1
                    ArrayValue1_Value = self.importer1.__global_prms__[self.variant_name1]["VAL_BLK"][v].value
                    p = 2
                    ArrayValue2_Value = self.importer2.__global_prms__[self.variant_name2]["VAL_BLK"][v].value
                    p = 3
                    ArrayValue3_Value = self.importer3.__global_prms__[self.variant_name3]["VAL_BLK"][v].value

                    ArrayValue1_X = list(range(len(ArrayValue1_Value[0])))
                    ArrayValue2_X = list(range(len(ArrayValue2_Value[0])))
                    ArrayValue3_X = list(range(len(ArrayValue3_Value[0])))

                    ArrayValue1_Y = list(range(len(ArrayValue1_Value)))
                    ArrayValue2_Y = list(range(len(ArrayValue2_Value)))
                    ArrayValue3_Y = list(range(len(ArrayValue3_Value)))

                except KeyError:
                    if p == 1:
                        ArrayValue2 = self.importer2.__global_prms__[self.variant_name2]["VAL_BLK"][v]
                        ArrayValue3 = self.importer3.__global_prms__[self.variant_name3]["VAL_BLK"][v]

                        ArrayValue2_Value = self.importer2.__global_prms__[self.variant_name2]["VAL_BLK"][v].value
                        ArrayValue3_Value = self.importer3.__global_prms__[self.variant_name3]["VAL_BLK"][v].value

                        ArrayValue2_X = list(range(len(ArrayValue2_Value[0])))
                        ArrayValue3_X = list(range(len(ArrayValue3_Value[0])))
                        ArrayValue1_X = ['NA'] * len(ArrayValue2_X)

                        ArrayValue2_Y = list(range(len(ArrayValue2_Value)))
                        ArrayValue3_Y = list(range(len(ArrayValue3_Value)))
                        ArrayValue1_Y = ['NA'] * len(ArrayValue2_Y)

                        ArrayValue1_Value = [['NA'] * len(ArrayValue1_X)] * len(ArrayValue1_Y)

                    elif p == 2:
                        ArrayValue1 = self.importer1.__global_prms__[self.variant_name1]["VAL_BLK"][v]
                        ArrayValue3 = self.importer3.__global_prms__[self.variant_name3]["VAL_BLK"][v]

                        ArrayValue1_Value = self.importer1.__global_prms__[self.variant_name1]["VAL_BLK"][v].value
                        ArrayValue3_Value = self.importer3.__global_prms__[self.variant_name3]["VAL_BLK"][v].value

                        ArrayValue1_X = list(range(len(ArrayValue1_Value[0])))
                        ArrayValue3_X = list(range(len(ArrayValue3_Value[0])))
                        ArrayValue2_X = ['NA'] * len(ArrayValue1_X)

                        ArrayValue1_Y = list(range(len(ArrayValue1_Value)))
                        ArrayValue3_Y = list(range(len(ArrayValue3_Value)))
                        ArrayValue2_Y = ['NA'] * len(ArrayValue1_Y)

                        ArrayValue2_Value = [['NA'] * len(ArrayValue2_X)] * len(ArrayValue2_Y)

                    else:
                        ArrayValue1 = self.importer1.__global_prms__[self.variant_name1]["VAL_BLK"][v]
                        ArrayValue2 = self.importer2.__global_prms__[self.variant_name2]["VAL_BLK"][v]

                        ArrayValue1_Value = self.importer1.__global_prms__[self.variant_name1]["VAL_BLK"][v].value
                        ArrayValue2_Value = self.importer2.__global_prms__[self.variant_name2]["VAL_BLK"][v].value

                        ArrayValue1_X = list(range(len(ArrayValue1_Value[0])))
                        ArrayValue2_X = list(range(len(ArrayValue2_Value[0])))
                        ArrayValue3_X = ['NA'] * len(ArrayValue1_X)

                        ArrayValue1_Y = list(range(len(ArrayValue1_Value)))
                        ArrayValue2_Y = list(range(len(ArrayValue2_Value)))
                        ArrayValue3_Y = ['NA'] * len(ArrayValue1_Y)

                        ArrayValue3_Value = [['NA'] * len(ArrayValue3_X)] * len(ArrayValue3_Y)

                if ArrayIndex >= 2:
                    StartPoint1_Y = StartPoint1_Y + len(ArrayValue1_Y_old) + 3
                else:
                    StartPoint1_Y = StartPoint1_Y

                worksheet.write(StartPoint1_Y, StartPoint1_x - 1, v, WriteTitleFormat)

                lengX = max(len(ArrayValue1_X), len(ArrayValue2_X), len(ArrayValue3_X))
                lengY = max(len(ArrayValue1_Y), len(ArrayValue2_Y), len(ArrayValue3_Y))

                # X축 크기 맞추기
                if len(ArrayValue1_X) < lengX:
                    ArrayValue1_Value = ArrayValue1_Value.tolist()
                    for i in range(lengX - len(ArrayValue1_X)):
                        ArrayValue1_X.append('NA')
                        for j in range(len(ArrayValue1_Y)):
                            for k in range(lengX - len(ArrayValue1_X)):
                                ArrayValue1_Value[j].append('NA')
                if len(ArrayValue2_X) < lengX:
                    ArrayValue2_Value = ArrayValue2_Value.tolist()
                    for i in range(lengX - len(ArrayValue2_X)):
                        ArrayValue2_X.append('NA')
                        for j in range(len(ArrayValue2_Y)):
                            for k in range(lengX - len(ArrayValue2_X)):
                                ArrayValue2_Value[j].append('NA')
                if len(ArrayValue3_X) < lengX:
                    MapValue3_Value = ArrayValue3_Value.tolist()
                    for i in range(lengX - len(ArrayValue3_X)):
                        ArrayValue3_X.append('NA')
                        for j in range(len(ArrayValue3_Y)):
                            for k in range(lengX - len(ArrayValue3_X)):
                                ArrayValue3_Value[j].append('NA')
                # Y축 크기 맞추기
                if len(ArrayValue1_Y) < lengY:
                    ArrayValue1_Value = ArrayValue1_Value.tolist()
                    for i in range(lengY - len(ArrayValue1_Y)):
                        ArrayValue1_Y.append(self, 'NA')
                        temp = ['NA'] * len(ArrayValue1_X)
                        ArrayValue1_Value.append(self, temp)
                if len(ArrayValue2_Y) < lengY:
                    ArrayValue2_Value = ArrayValue2_Value.toslit()
                    for i in range(lengY - len(ArrayValue2_Y)):
                        ArrayValue2_Y.append(self, 'NA')
                        temp = ['NA'] * len(ArrayValue2_X)
                        ArrayValue2_Value.append(self, temp)
                if len(ArrayValue3_Y) < lengY:
                    ArrayValue3_Value = ArrayValue3_Value.tolist()
                    for i in range(lengY - len(ArrayValue3_Y)):
                        ArrayValue3_Y.append(self, 'NA')
                        temp = ['NA'] * len(ArrayValue3_X)
                        ArrayValue3_Value.append(self, temp)

                for l in range(len(ArrayValue1_X)):
                    worksheet.write(StartPoint1_Y + 1, l + StartPoint1_x, ArrayValue1_X[l], data_format_axis)
                    worksheet.write(StartPoint1_Y + 1, l + StartPoint1_x + len(ArrayValue1_X) + 3, ArrayValue2_X[l],
                                    data_format_axis)
                    worksheet.write(StartPoint1_Y + 1, l + StartPoint1_x + 3 + (len(ArrayValue1_X) * 2) + 3,
                                    ArrayValue3_X[l],
                                    data_format_axis)

                    if ArrayValue1_X[l] != ArrayValue2_X[l]:
                        result[0] = 1
                    else:
                        pass
                    if ArrayValue1_X[l] != ArrayValue3_X[l]:
                        result[1] = 1
                    else:
                        pass

                Layer1_data_ArrayXAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 1, StartPoint1_x)
                Layer2_data_ArrayXAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 1,
                                                                              StartPoint1_x + len(ArrayValue1_X) + 3)
                Before_data_ArrayXAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 1,
                                                                              StartPoint1_x + 3 + (
                                                                                      len(ArrayValue1_X) * 2) + 3)

                after_data_ArrayXAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 1,
                                                                            StartPoint1_x + len(ArrayValue1_X) + 3,
                                                                            StartPoint1_Y + 1,
                                                                            StartPoint1_x + len(ArrayValue1_X) + 3 + l)

                Before_data_ArrayXAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 1,
                                                                             StartPoint1_x + 3 + (
                                                                                     len(ArrayValue1_X) * 2) + 3,
                                                                             StartPoint1_Y + 1,
                                                                             StartPoint1_x + 3 + (
                                                                                     len(ArrayValue1_X) * 2) + 3 + l)

                worksheet.conditional_format(after_data_ArrayXAxisange, {'type': 'formula',
                                                                         'criteria': Layer1_data_ArrayXAxis + ' <> ' + Layer2_data_ArrayXAxis,
                                                                         'format': FormatForAfter_NotEqual})

                worksheet.conditional_format(Before_data_ArrayXAxisange, {'type': 'formula',
                                                                          'criteria': Layer1_data_ArrayXAxis + ' <> ' + Before_data_ArrayXAxis,
                                                                          'format': FormatForAfter_Change})

                for m in range(len(ArrayValue1_Y)):
                    worksheet.write(StartPoint1_Y + m + 2, StartPoint1_x - 1, ArrayValue1_Y[m], data_format_axis)
                    worksheet.write(StartPoint1_Y + m + 2, StartPoint1_x - 1 + len(ArrayValue1_X) + 3, ArrayValue2_Y[m],
                                    data_format_axis)
                    worksheet.write(StartPoint1_Y + m + 2, StartPoint1_x - 1 + 3 + (len(ArrayValue1_X) * 2) + 3,
                                    ArrayValue3_Y[m],
                                    data_format_axis)

                    if ArrayValue1_Y[m] != ArrayValue2_Y[m]:
                        result[0] = 1
                    else:
                        pass
                    if ArrayValue1_Y[m] != ArrayValue3_Y[m]:
                        result[1] = 1
                    else:
                        pass

                Layer1_data_ArrayYAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 2, StartPoint1_x - 1)
                Layer2_data_ArrayYAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 2,
                                                                              StartPoint1_x - 1 + len(
                                                                                  ArrayValue1_X) + 3)

                Before_data_ArrayYAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 2,
                                                                              StartPoint1_x - 1 + 3 + (
                                                                                      len(ArrayValue1_X) * 2) + 3)

                after_data_ArrayYAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 2,
                                                                            StartPoint1_x - 1 + len(ArrayValue1_X) + 3,
                                                                            StartPoint1_Y + 2 + m,
                                                                            StartPoint1_x - 1 + len(ArrayValue1_X) + 3)

                Before_data_ArrayYAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 2,
                                                                             StartPoint1_x - 1 + 3 + (
                                                                                     len(ArrayValue1_X) * 2) + 3,
                                                                             StartPoint1_Y + 2 + m,
                                                                             StartPoint1_x - 1 + 3 + (
                                                                                     len(ArrayValue1_X) * 2) + 3)

                worksheet.conditional_format(after_data_ArrayYAxisange, {'type': 'formula',
                                                                         'criteria': Layer1_data_ArrayYAxis + ' <> ' + Layer2_data_ArrayYAxis,
                                                                         'format': FormatForAfter_NotEqual})

                worksheet.conditional_format(Before_data_ArrayYAxisange, {'type': 'formula',
                                                                          'criteria': Layer1_data_ArrayYAxis + ' <> ' + Before_data_ArrayYAxis,
                                                                          'format': FormatForAfter_Change})

                for p in range(len(ArrayValue1_Y)):
                    for o in range(len(ArrayValue1_X)):
                        worksheet.write(StartPoint1_Y + 2 + p, o + StartPoint1_x, ArrayValue1_Value[p][o],
                                        data_format_data)
                        worksheet.write(StartPoint1_Y + 2 + p, o + StartPoint1_x + len(ArrayValue1_X) + 3,
                                        ArrayValue1_Value[p][o],
                                        data_format_data)
                        worksheet.write(StartPoint1_Y + 2 + p, o + StartPoint1_x + (len(ArrayValue1_X) * 2) + 6,
                                        ArrayValue3_Value[p][o], data_format_data)

                        if ArrayValue1_Value[p][o] != ArrayValue1_Value[p][o]:
                            result[0] = 1
                        else:
                            pass

                        if ArrayValue1_Value[p][o] != ArrayValue3_Value[p][o]:
                            result[1] = 1
                        else:
                            pass

                Layer1_data_ArrayData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 2, StartPoint1_x)
                Layer2_data_ArrayData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 2,
                                                                             StartPoint1_x + len(ArrayValue1_X) + 3)
                Before_data_ArrayData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 2,
                                                                             StartPoint1_x + (
                                                                                     len(ArrayValue1_X) * 2) + 6)

                after_data_ArrayDatasange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 2,
                                                                            StartPoint1_x + len(ArrayValue1_X) + 3,
                                                                            StartPoint1_Y + 2 + p,
                                                                            o + StartPoint1_x + len(ArrayValue1_X) + 3)

                Before_data_ArrayDatasange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 2,
                                                                             StartPoint1_x + (
                                                                                     len(ArrayValue1_X) * 2) + 6,
                                                                             StartPoint1_Y + 2 + p,
                                                                             o + StartPoint1_x + (
                                                                                     len(ArrayValue1_X) * 2) + 6)

                worksheet.conditional_format(after_data_ArrayDatasange, {'type': 'formula',
                                                                         'criteria': Layer1_data_ArrayData + ' <> ' + Layer2_data_ArrayData,
                                                                         'format': FormatForAfter_NotEqual})

                worksheet.conditional_format(Before_data_ArrayDatasange, {'type': 'formula',
                                                                          'criteria': Layer1_data_ArrayData + ' <> ' + Before_data_ArrayData,
                                                                          'format': FormatForAfter_Change})

            # MAP Lay1/2
            StartPoint1_x = StartPoint1_x
            StartPoint1_Y = StartPoint1_Y + 6

            MapValue = self.importer1.__global_prms__[self.variant_name1]["MAP"]
            MapValue_lay2 = self.importer2.__global_prms__[self.variant_name2]["MAP"]
            MapIndex = 0
            MapValue1_Y = 0

            for k in MapValue.keys():
                MapValue1_y_old = MapValue1_Y
                MapIndex = MapIndex + 1

                p = 0
                try:
                    p = 1
                    MapValue1 = self.importer1.__global_prms__[self.variant_name1]["MAP"][k]
                    p = 2
                    MapValue2_lay2 = self.importer2.__global_prms__[self.variant_name2]["MAP"][k]
                    p = 3
                    MapValue3 = self.importer3.__global_prms__[self.variant_name3]["MAP"][k]

                    p = 1
                    MapValue1_X = self.importer1.__global_prms__[self.variant_name1]["COM_AXIS"][
                        MapValue1.__xaxis__].value
                    p = 2
                    MapValue2_X_lay2 = self.importer2.__global_prms__[self.variant_name2]["COM_AXIS"][
                        MapValue2_lay2.__xaxis__].value
                    p = 3
                    MapValue3_X = self.importer3.__global_prms__[self.variant_name3]["COM_AXIS"][
                        MapValue3.__xaxis__].value

                    p = 1
                    MapValue1_Y = self.importer1.__global_prms__[self.variant_name1]["COM_AXIS"][
                        MapValue1.__yaxis__].value
                    p = 2
                    MapValue2_Y_lay2 = self.importer2.__global_prms__[self.variant_name2]["COM_AXIS"][
                        MapValue2_lay2.__yaxis__].value
                    p = 3
                    MapValue3_Y = self.importer3.__global_prms__[self.variant_name3]["COM_AXIS"][
                        MapValue3.__yaxis__].value

                    p = 1
                    MapValue1_Value = self.importer1.__global_prms__[self.variant_name1]["MAP"][k].value
                    p = 2
                    MapValue1_Value_lay2 = self.importer2.__global_prms__[self.variant_name2]["MAP"][k].value
                    p = 3
                    MapValue3_Value = self.importer3.__global_prms__[self.variant_name3]["MAP"][k].value

                except KeyError:

                    if p == 1:
                        MapValue2_lay2 = self.importer2.__global_prms__[self.variant_name2]["MAP"][k]
                        MapValue3 = self.importer3.__global_prms__[self.variant_name3]["MAP"][k]

                        MapValue2_X_lay2 = self.importer2.__global_prms__[self.variant_name2]["COM_AXIS"][
                            MapValue2_lay2.__xaxis__].value
                        MapValue3_X = self.importer3.__global_prms__[self.variant_name3]["COM_AXIS"][
                            MapValue3.__xaxis__].value
                        MapValue1_X = ['NA'] * len(MapValue2_X_lay2)

                        MapValue2_Y_lay2 = self.importer2.__global_prms__[self.variant_name2]["COM_AXIS"][
                            MapValue2_lay2.__yaxis__].value
                        MapValue3_Y = self.importer3.__global_prms__[self.variant_name3]["COM_AXIS"][
                            MapValue3.__yaxis__].value
                        MapValue1_Y = ['NA'] * len(MapValue2_Y_lay2)

                        MapValue1_Value_lay2 = self.importer2.__global_prms__[self.variant_name2]["MAP"][k].value
                        MapValue3_Value = self.importer3.__global_prms__[self.variant_name3]["MAP"][k].value
                        MapValue1_Value = [['NA'] * len(MapValue2_X_lay2)] * len(MapValue2_Y_lay2)

                    elif p == 2:
                        MapValue1 = self.importer1.__global_prms__[self.variant_name1]["MAP"][k]
                        MapValue3 = self.importer3.__global_prms__[self.variant_name3]["MAP"][k]

                        MapValue1_X = self.importer1.__global_prms__[self.variant_name1]["COM_AXIS"][
                            MapValue1.__xaxis__].value
                        MapValue3_X = self.importer3.__global_prms__[self.variant_name3]["COM_AXIS"][
                            MapValue3.__xaxis__].value
                        MapValue2_X_lay2 = ['NA'] * len(MapValue1_X)

                        MapValue1_Y = self.importer1.__global_prms__[self.variant_name1]["COM_AXIS"][
                            MapValue1.__yaxis__].value
                        MapValue3_Y = self.importer3.__global_prms__[self.variant_name3]["COM_AXIS"][
                            MapValue3.__yaxis__].value
                        MapValue2_Y_lay2 = ['NA'] * len(MapValue1_Y)

                        MapValue1_Value = self.importer1.__global_prms__[self.variant_name1]["MAP"][k].value
                        MapValue3_Value = self.importer3.__global_prms__[self.variant_name3]["MAP"][k].value
                        MapValue1_Value_lay2 = [['NA'] * len(MapValue1_X)] * len(MapValue1_Y)

                    else:
                        MapValue1 = self.importer1.__global_prms__[self.variant_name1]["MAP"][k]
                        MapValue2_lay2 = self.importer2.__global_prms__[self.variant_name2]["MAP"][k]

                        MapValue1_X = self.importer1.__global_prms__[self.variant_name1]["COM_AXIS"][
                            MapValue1.__xaxis__].value
                        MapValue2_X_lay2 = self.importer2.__global_prms__[self.variant_name2]["COM_AXIS"][
                            MapValue2_lay2.__xaxis__].value
                        MapValue3_X = ['NA'] * len(MapValue1_X)

                        MapValue1_Y = self.importer1.__global_prms__[self.variant_name1]["COM_AXIS"][
                            MapValue1.__yaxis__].value
                        MapValue2_Y_lay2 = self.importer2.__global_prms__[self.variant_name2]["COM_AXIS"][
                            MapValue2_lay2.__yaxis__].value
                        MapValue3_Y = ['NA'] * len(MapValue1_Y)

                        MapValue1_Value = self.importer1.__global_prms__[self.variant_name1]["MAP"][k].value
                        MapValue1_Value_lay2 = self.importer2.__global_prms__[self.variant_name2]["MAP"][k].value
                        MapValue3_Value = [['NA'] * len(MapValue1_X)] * len(MapValue1_Y)

                if MapIndex >= 2:
                    StartPoint1_Y = StartPoint1_Y + len(MapValue1_y_old) + 3
                else:
                    StartPoint1_Y = StartPoint1_Y

                worksheet.write(StartPoint1_Y, StartPoint1_x - 1, k, WriteTitleFormat)

                lengX = max(len(MapValue1_X), len(MapValue2_X_lay2), len(MapValue3_X))
                lengY = max(len(MapValue1_Y), len(MapValue2_Y_lay2), len(MapValue3_Y))

                # X축 크기 맞추기
                if len(MapValue1_X) < lengX:
                    MapValue1_X = MapValue1_X.tolist()
                    MapValue1_Value = MapValue1_Value.tolist()
                    for i in range(lengX - len(MapValue1_X)):
                        MapValue1_X.append('NA')
                        for j in range(len(MapValue1_Y)):
                            for k in range(lengX - len(MapValue1_X)):
                                MapValue1_Value[j].append('NA')
                if len(MapValue2_X_lay2) < lengX:
                    MapValue2_X_lay2 = MapValue2_X_lay2.tolist()
                    MapValue1_Value_lay2 = MapValue1_Value_lay2.tolist()
                    for i in range(lengX - len(MapValue2_X_lay2)):
                        MapValue2_X_lay2.append('NA')
                        for j in range(len(MapValue2_Y_lay2)):
                            for k in range(lengX - len(MapValue2_X_lay2)):
                                MapValue1_Value_lay2[j].append('NA')
                if len(MapValue3_X) < lengX:
                    MapValue3_X = MapValue3_X.tolist()
                    MapValue3_Value = MapValue3_Value.tolist()
                    for i in range(lengX - len(MapValue3_X)):
                        MapValue3_X.append('NA')
                        for j in range(len(MapValue3_Y)):
                            for k in range(lengX - len(MapValue3_X)):
                                MapValue3_Value[j].append('NA')

                # Y축 크기 맞추기
                if len(MapValue1_Y) < lengY:
                    MapValue1_Y = MapValue1_Y.tolist()
                    MapValue1_Value = MapValue1_Value.tolist()
                    for i in range(lengY - len(MapValue1_Y)):
                        MapValue1_Y.append(self, 'NA')
                        temp = ['NA'] * len(MapValue1_X)
                        MapValue1_Value.append(self, temp)
                if len(MapValue2_Y_lay2) < lengY:
                    MapValue2_Y_lay2 = MapValue2_Y_lay2.tolist()
                    MapValue1_Value_lay2 = MapValue1_Value_lay2.toslit()
                    for i in range(lengY - len(MapValue2_Y_lay2)):
                        MapValue2_Y_lay2.append(self, 'NA')
                        temp = ['NA'] * len(MapValue2_X_lay2)
                        MapValue1_Value_lay2.append(self, temp)
                if len(MapValue3_Y) < lengY:
                    MapValue3_Y = MapValue3_Y.tolist()
                    MapValue3_Value = MapValue3_Value.tolist()
                    for i in range(lengY - len(MapValue3_Y)):
                        MapValue3_Y.append(self, 'NA')
                        temp = ['NA'] * len(MapValue3_X)
                        MapValue3_Value.append(self, temp)

                for l in range(lengX):
                    worksheet.write(StartPoint1_Y + 1, l + StartPoint1_x, MapValue1_X[l], data_format_axis)
                    worksheet.write(StartPoint1_Y + 1, l + StartPoint1_x + len(MapValue1_X) + 3, MapValue2_X_lay2[l],
                                    data_format_axis)
                    worksheet.write(StartPoint1_Y + 1, l + StartPoint1_x + 3 + (len(MapValue1_X) * 2) + 3,
                                    MapValue3_X[l],
                                    data_format_axis)

                    if MapValue1_X[l] != MapValue2_X_lay2[l]:
                        result[0] = 1
                    else:
                        pass
                    if MapValue1_X[l] != MapValue3_X[l]:
                        result[1] = 1
                    else:
                        pass

                Layer1_data_MapXAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 1, StartPoint1_x)
                Layer2_data_MapXAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 1,
                                                                            StartPoint1_x + len(MapValue1_X) + 3)
                Before_data_MapXAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 1,
                                                                            StartPoint1_x + 3 + (
                                                                                        len(MapValue1_X) * 2) + 3)

                after_data_MapXAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 1,
                                                                          StartPoint1_x + len(MapValue1_X) + 3,
                                                                          StartPoint1_Y + 1,
                                                                          StartPoint1_x + len(MapValue1_X) + 3 + l)

                Before_data_MapXAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 1,
                                                                           StartPoint1_x + 3 + (
                                                                                       len(MapValue1_X) * 2) + 3,
                                                                           StartPoint1_Y + 1,
                                                                           StartPoint1_x + 3 + (
                                                                                   len(MapValue1_X) * 2) + 3 + l)

                worksheet.conditional_format(after_data_MapXAxisange, {'type': 'formula',
                                                                       'criteria': Layer1_data_MapXAxis + ' <> ' + Layer2_data_MapXAxis,
                                                                       'format': FormatForAfter_NotEqual})

                worksheet.conditional_format(Before_data_MapXAxisange, {'type': 'formula',
                                                                        'criteria': Layer1_data_MapXAxis + ' <> ' + Before_data_MapXAxis,
                                                                        'format': FormatForAfter_Change})

                for m in range(len(MapValue1_Y)):
                    worksheet.write(StartPoint1_Y + m + 2, StartPoint1_x - 1, MapValue1_Y[m], data_format_axis)
                    worksheet.write(StartPoint1_Y + m + 2, StartPoint1_x - 1 + len(MapValue1_X) + 3,
                                    MapValue2_Y_lay2[m],
                                    data_format_axis)
                    worksheet.write(StartPoint1_Y + m + 2, StartPoint1_x - 1 + 3 + (len(MapValue1_X) * 2) + 3,
                                    MapValue3_Y[m], data_format_axis)

                    if MapValue1_Y[m] != MapValue2_Y_lay2[m]:
                        result[0] = 1
                    else:
                        pass
                    if MapValue1_Y[m] != MapValue3_Y[m]:
                        result[1] = 1
                    else:
                        pass

                Layer1_data_MapYAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 2, StartPoint1_x - 1)
                Layer2_data_MapYAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 2,
                                                                            StartPoint1_x - 1 + len(MapValue1_X) + 3)

                Before_data_MapYAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 2,
                                                                            StartPoint1_x - 1 + 3 + (
                                                                                    len(MapValue1_X) * 2) + 3)

                after_data_MapYAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 2,
                                                                          StartPoint1_x - 1 + len(MapValue1_X) + 3,
                                                                          StartPoint1_Y + 2 + m,
                                                                          StartPoint1_x - 1 + len(MapValue1_X) + 3)

                Before_data_MapYAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 2,
                                                                           StartPoint1_x - 1 + 3 + (
                                                                                   len(MapValue1_X) * 2) + 3,
                                                                           StartPoint1_Y + 2 + m,
                                                                           StartPoint1_x - 1 + 3 + (
                                                                                   len(MapValue1_X) * 2) + 3)

                worksheet.conditional_format(after_data_MapYAxisange, {'type': 'formula',
                                                                       'criteria': Layer1_data_MapYAxis + ' <> ' + Layer2_data_MapYAxis,
                                                                       'format': FormatForAfter_NotEqual})

                worksheet.conditional_format(Before_data_MapYAxisange, {'type': 'formula',
                                                                        'criteria': Layer1_data_MapYAxis + ' <> ' + Before_data_MapYAxis,
                                                                        'format': FormatForAfter_Change})

                for p in range(len(MapValue1_Y)):
                    for o in range(len(MapValue1_X)):
                        worksheet.write(StartPoint1_Y + 2 + p, o + StartPoint1_x, MapValue1_Value[p][o],
                                        data_format_data)
                        worksheet.write(StartPoint1_Y + 2 + p, o + StartPoint1_x + len(MapValue1_X) + 3,
                                        MapValue1_Value_lay2[p][o], data_format_data)
                        worksheet.write(StartPoint1_Y + 2 + p, o + StartPoint1_x + (len(MapValue1_X) * 2) + 6,
                                        MapValue3_Value[p][o], data_format_data)

                        if MapValue1_Value[p][o] != MapValue1_Value_lay2[p][o]:
                            result[0] = 1
                        else:
                            pass

                        if MapValue1_Value[p][o] != MapValue3_Value[p][o]:
                            result[1] = 1
                        else:
                            pass

                Layer1_data_MapData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 2, StartPoint1_x)
                Layer2_data_MapData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 2,
                                                                           StartPoint1_x + len(MapValue1_X) + 3)
                Before_data_MapData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 2,
                                                                           StartPoint1_x + (len(MapValue1_X) * 2) + 6)

                after_data_MapDatasange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 2,
                                                                          StartPoint1_x + len(MapValue1_X) + 3,
                                                                          StartPoint1_Y + 2 + p,
                                                                          o + StartPoint1_x + len(MapValue1_X) + 3)

                Before_data_MapDatasange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 2,
                                                                           StartPoint1_x + (len(MapValue1_X) * 2) + 6,
                                                                           StartPoint1_Y + 2 + p,
                                                                           o + StartPoint1_x + (
                                                                                       len(MapValue1_X) * 2) + 6)

                worksheet.conditional_format(after_data_MapDatasange, {'type': 'formula',
                                                                       'criteria': Layer1_data_MapData + ' <> ' + Layer2_data_MapData,
                                                                       'format': FormatForAfter_NotEqual})

                worksheet.conditional_format(Before_data_MapDatasange, {'type': 'formula',
                                                                        'criteria': Layer1_data_MapData + ' <> ' + Before_data_MapData,
                                                                        'format': FormatForAfter_Change})

        except KeyError:
            result[0] = 2
            result[1] = 2

        return result

    def __drawMapSingleLay2__(self, startx, starty, envir):
        workbook = self.workbook
        worksheet = self.current_sheet
        worksheet.set_column('A:FH', 5)
        StartPoint1_x = startx
        StartPoint1_Y = starty
        result = 0

        data_format_data = workbook.add_format({'border': 1})
        data_format_axis = workbook.add_format({'border': 1, 'bg_color': '#DCE6F1'})
        WriteTitleFormat = workbook.add_format({'font_color': '#000000', 'bold': True, 'font_size': 11})
        WrongFormat = workbook.add_format({'bg_color': '#FF0000'})

        # 조건부 서식 Format
        FormatForAfter_NotEqual = workbook.add_format({'bold': True, 'bg_color': 'red'})
        # FormatForAfter_NotEqual = workbook.add_format({'font_color': '#000000', 'bold': True})
        FormatForAfter_Greater = workbook.add_format({'font_color': '#000000', 'bg_color': '#F2DCDB', 'bold': True})
        FormatForAfter_Equal = workbook.add_format({'font_color': '#000000', 'bg_color': '#FFFFFF'})
        FormatForAfter_AxisEqual = workbook.add_format({'font_color': '#000000', 'bg_color': '#DCE6F1'})

        import1_prm = getattr(self.importer1, envir)

        # VALUE 챠트 Layer1/Layer2
        try:

            ConstValue1 = import1_prm[self.variant_name1]["VALUE"]
            ConstValue2 = import1_prm[self.variant_name2]["VALUE"]
            Nr = len(ConstValue1)
            worksheet.write(StartPoint1_Y, StartPoint1_x + 5, "Layer1", WriteTitleFormat)
            worksheet.write(StartPoint1_Y, StartPoint1_x + 6, "Layer2", WriteTitleFormat)

            index = 0
            for i in ConstValue1.keys():
                index = index + 1

                VariableValue1 = import1_prm[self.variant_name1]["VALUE"][i].value

                if ConstValue2 == {} or i not in ConstValue2.keys():
                    VariableValue2 = 'NA'
                    result = 1
                else:
                    VariableValue2 = import1_prm[self.variant_name2]["VALUE"][i].value
                    if VariableValue1 != VariableValue2:
                        result = 1
                    else:
                        pass

                worksheet.write(StartPoint1_Y + index, StartPoint1_x - 1, i, WriteTitleFormat)
                worksheet.write(StartPoint1_Y + index, StartPoint1_x + 5, VariableValue1, data_format_data)
                worksheet.write(StartPoint1_Y + index, StartPoint1_x + 6, VariableValue2, data_format_data)

            Layer1_data_sp = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 1, StartPoint1_x + 5)
            Layer2_data_sp = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 1, StartPoint1_x + 6)

            after_data_range = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 1, StartPoint1_x + 6, StartPoint1_Y + Nr,
                                                               StartPoint1_x + 6)
            after_data_range1 = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 1, StartPoint1_x + 7,
                                                                StartPoint1_Y + Nr,
                                                                StartPoint1_x + 7)

            worksheet.conditional_format(after_data_range, {'type': 'formula',
                                                            'criteria': Layer1_data_sp + ' <> ' + Layer2_data_sp,
                                                            'format': FormatForAfter_NotEqual})

            # Curve Layer1/Layer2

            StartPoint1_x = StartPoint1_x
            StartPoint1_Y = StartPoint1_Y + Nr

            CurValue = import1_prm[self.variant_name1]["CURVE"]
            CurValue_lay2 = import1_prm[self.variant_name2]["CURVE"]
            CurIndex = 0

            for j in CurValue.keys():
                CurIndex = CurIndex + 1
                CurveValue1 = import1_prm[self.variant_name1]["CURVE"][j]
                CurveValue1_X = import1_prm[self.variant_name1]["COM_AXIS"][CurveValue1.__xaxis__].value
                CurveValue1_Value = import1_prm[self.variant_name1]["CURVE"][j].value

                if CurValue_lay2 == {} or j not in CurValue_lay2.keys():
                    CurveValue2_X_lay2 = ['NA'] * len(CurveValue1_X)
                    CurveValue2_Value_lay2 = ['NA'] * len(CurveValue1_Value)
                else:
                    CurveValue2_lay2 = import1_prm[self.variant_name2]["CURVE"][j]
                    CurveValue2_X_lay2 = import1_prm[self.variant_name2]["COM_AXIS"][
                        CurveValue2_lay2.__xaxis__].value
                    CurveValue2_Value_lay2 = import1_prm[self.variant_name2]["CURVE"][j].value

                if CurIndex >= 2:
                    StartPoint1_Y = StartPoint1_Y + 4
                else:
                    StartPoint1_Y = StartPoint1_Y

                worksheet.write(StartPoint1_Y + 2, StartPoint1_x - 1, j, WriteTitleFormat)

                for i in range(len(CurveValue1_X)):
                    # Axis
                    worksheet.write(StartPoint1_Y + 3, StartPoint1_x + i - 1, CurveValue1_X[i], data_format_axis)
                    worksheet.write(StartPoint1_Y + 3, StartPoint1_x + i + 1 + len(CurveValue1_X) + 2,
                                    CurveValue2_X_lay2[i], data_format_axis)

                    if CurveValue1_X[i] != CurveValue2_X_lay2[i]:
                        result = 1
                    else:
                        pass

                    # Curve
                    worksheet.write(StartPoint1_Y + 4, StartPoint1_x + i - 1, CurveValue1_Value[i], data_format_data)
                    worksheet.write(StartPoint1_Y + 4, StartPoint1_x + i + 1 + len(CurveValue1_X) + 2,
                                    CurveValue2_Value_lay2[i], data_format_data)

                    if CurveValue1_Value[i] != CurveValue2_Value_lay2[i]:
                        result = 1
                    else:
                        pass

                Layer1_data_CurAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 3, StartPoint1_x - 1)
                Layer2_data_CurAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 3,
                                                                           StartPoint1_x + 2 + len(
                                                                               CurveValue1_X) + 2 - 1)

                Layer1_data_CurData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 4, StartPoint1_x - 1)
                Layer2_data_CurData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 4,
                                                                           StartPoint1_x + 2 + len(
                                                                               CurveValue1_X) + 2 - 1)

                after_data_Currange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 4,
                                                                      StartPoint1_x + 2 + len(CurveValue1_X) + 2 - 1,
                                                                      StartPoint1_Y + 4,
                                                                      StartPoint1_x + 2 + len(CurveValue1_X) + 1 + i)
                after_data_CurrAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 3,
                                                                          StartPoint1_x + 2 + len(
                                                                              CurveValue1_X) + 2 - 1,
                                                                          StartPoint1_Y + 3,
                                                                          StartPoint1_x + 2 + len(
                                                                              CurveValue1_X) + 1 + i)
                Before_data_CurrAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 3,
                                                                           StartPoint1_x + 1 + 4 + (
                                                                                   len(CurveValue1_X) * 2),
                                                                           StartPoint1_Y + 3,
                                                                           StartPoint1_x + 1 + 4 + (
                                                                                   len(CurveValue1_X) * 2) + i)

                Before_data_Currange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 4,
                                                                       StartPoint1_x + 1 + 4 + (len(CurveValue1_X) * 2),
                                                                       StartPoint1_Y + 4,
                                                                       StartPoint1_x + 1 + 4 + (
                                                                               len(CurveValue1_X) * 2) + i)
                worksheet.conditional_format(after_data_CurrAxisange, {'type': 'formula',
                                                                       'criteria': Layer1_data_CurAxis + ' <> ' + Layer2_data_CurAxis,
                                                                       'format': FormatForAfter_NotEqual})
                worksheet.conditional_format(after_data_Currange, {'type': 'formula',
                                                                   'criteria': Layer1_data_CurData + ' <> ' + Layer2_data_CurData,
                                                                   'format': FormatForAfter_NotEqual})

            # Layer2에는 있는데 1에는 없음
            for j in CurValue_lay2.keys():
                if j in CurValue.keys():
                    pass

                CurIndex = CurIndex + 1

                CurveValue2_lay2 = import1_prm[self.variant_name2]["CURVE"][j]
                CurveValue2_X_lay2 = import1_prm[self.variant_name2]["COM_AXIS"][
                    CurveValue2_lay2.__xaxis__].value
                CurveValue2_Value_lay2 = import1_prm[self.variant_name2]["CURVE"][j].value

                if CurValue == {} or j not in CurValue.keys():
                    CurveValue1_X = ['NA'] * len(CurveValue2_X_lay2)
                    CurveValue1_Value = ['NA'] * len(CurveValue2_Value_lay2)
                else:
                    CurveValue1 = import1_prm[self.variant_name1]["CURVE"][j]
                    CurveValue1_X = import1_prm[self.variant_name1]["COM_AXIS"][
                        CurveValue1.__xaxis__].value
                    CurveValue1_Value = import1_prm[self.variant_name1]["CURVE"][j].value

                if CurIndex >= 2:
                    StartPoint1_Y = StartPoint1_Y + 4
                else:
                    StartPoint1_Y = StartPoint1_Y

                worksheet.write(StartPoint1_Y + 2, StartPoint1_x - 1, j, WriteTitleFormat)

                for i in range(len(CurveValue1_X)):
                    # Axis
                    worksheet.write(StartPoint1_Y + 3, StartPoint1_x + i - 1, CurveValue1_X[i], data_format_axis)
                    worksheet.write(StartPoint1_Y + 3, StartPoint1_x + i + 1 + len(CurveValue1_X) + 2,
                                    CurveValue2_X_lay2[i], data_format_axis)

                    if CurveValue1_X[i] != CurveValue2_X_lay2[i]:
                        result = 1
                    else:
                        pass

                    # Curve
                    worksheet.write(StartPoint1_Y + 4, StartPoint1_x + i - 1, CurveValue1_Value[i], data_format_data)
                    worksheet.write(StartPoint1_Y + 4, StartPoint1_x + i + 1 + len(CurveValue1_X) + 2,
                                    CurveValue2_Value_lay2[i], data_format_data)

                    if CurveValue1_Value[i] != CurveValue2_Value_lay2[i]:
                        result = 1
                    else:
                        pass

                Layer1_data_CurAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 3, StartPoint1_x - 1)
                Layer2_data_CurAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 3,
                                                                           StartPoint1_x + 2 + len(
                                                                               CurveValue1_X) + 2 - 1)

                Layer1_data_CurData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 4, StartPoint1_x - 1)
                Layer2_data_CurData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 4,
                                                                           StartPoint1_x + 2 + len(
                                                                               CurveValue1_X) + 2 - 1)

                after_data_Currange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 4,
                                                                      StartPoint1_x + 2 + len(CurveValue1_X) + 2 - 1,
                                                                      StartPoint1_Y + 4,
                                                                      StartPoint1_x + 2 + len(CurveValue1_X) + 1 + i)
                after_data_CurrAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 3,
                                                                          StartPoint1_x + 2 + len(
                                                                              CurveValue1_X) + 2 - 1,
                                                                          StartPoint1_Y + 3,
                                                                          StartPoint1_x + 2 + len(
                                                                              CurveValue1_X) + 1 + i)
                Before_data_CurrAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 3,
                                                                           StartPoint1_x + 1 + 4 + (
                                                                                   len(CurveValue1_X) * 2),
                                                                           StartPoint1_Y + 3,
                                                                           StartPoint1_x + 1 + 4 + (
                                                                                   len(CurveValue1_X) * 2) + i)

                Before_data_Currange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 4,
                                                                       StartPoint1_x + 1 + 4 + (len(CurveValue1_X) * 2),
                                                                       StartPoint1_Y + 4,
                                                                       StartPoint1_x + 1 + 4 + (
                                                                               len(CurveValue1_X) * 2) + i)
                worksheet.conditional_format(after_data_CurrAxisange, {'type': 'formula',
                                                                       'criteria': Layer1_data_CurAxis + ' <> ' + Layer2_data_CurAxis,
                                                                       'format': FormatForAfter_NotEqual})
                worksheet.conditional_format(after_data_Currange, {'type': 'formula',
                                                                   'criteria': Layer1_data_CurData + ' <> ' + Layer2_data_CurData,
                                                                   'format': FormatForAfter_NotEqual})

            # MAP Lay1/2
            StartPoint1_x = StartPoint1_x
            StartPoint1_Y = StartPoint1_Y + 6

            MapValue = import1_prm[self.variant_name1]["MAP"]
            MapValue_lay2 = import1_prm[self.variant_name2]["MAP"]
            MapIndex = 0
            MapValue1_Y = 0
            for k in MapValue.keys():
                MapValue1_y_old = MapValue1_Y
                MapIndex = MapIndex + 1

                MapValue1 = import1_prm[self.variant_name1]["MAP"][k]
                MapValue2_lay2 = import1_prm[self.variant_name2]["MAP"][k]

                MapValue1_X = import1_prm[self.variant_name1]["COM_AXIS"][MapValue1.__xaxis__].value
                MapValue2_X_lay2 = import1_prm[self.variant_name2]["COM_AXIS"][
                    MapValue2_lay2.__xaxis__].value

                MapValue1_Y = import1_prm[self.variant_name1]["COM_AXIS"][MapValue1.__yaxis__].value
                MapValue2_Y_lay2 = import1_prm[self.variant_name2]["COM_AXIS"][
                    MapValue2_lay2.__yaxis__].value

                MapValue1_Value = import1_prm[self.variant_name1]["MAP"][k].value
                MapValue1_Value_lay2 = import1_prm[self.variant_name2]["MAP"][k].value

                if MapIndex >= 2:
                    StartPoint1_Y = StartPoint1_Y + len(MapValue1_y_old) + 3
                else:
                    StartPoint1_Y = StartPoint1_Y

                worksheet.write(StartPoint1_Y, StartPoint1_x - 1, k, WriteTitleFormat)

                for l in range(len(MapValue1_X)):
                    worksheet.write(StartPoint1_Y + 1, l + StartPoint1_x, MapValue1_X[l], data_format_axis)
                    worksheet.write(StartPoint1_Y + 1, l + StartPoint1_x + len(MapValue1_X) + 3, MapValue2_X_lay2[l],
                                    data_format_axis)

                    if MapValue1_X[l] != MapValue2_X_lay2[l]:
                        result = 1
                    else:
                        pass

                Layer1_data_MapXAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 1, StartPoint1_x)
                Layer2_data_MapXAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 1,
                                                                            StartPoint1_x + len(MapValue1_X) + 3)

                after_data_MapXAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 1,
                                                                          StartPoint1_x + len(MapValue1_X) + 3,
                                                                          StartPoint1_Y + 1,
                                                                          StartPoint1_x + len(MapValue1_X) + 3 + l)

                Before_data_MapXAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 1,
                                                                           StartPoint1_x + 3 + (
                                                                                   len(MapValue1_X) * 2) + 3,
                                                                           StartPoint1_Y + 1,
                                                                           StartPoint1_x + 3 + (
                                                                                   len(MapValue1_X) * 2) + 3 + l)

                worksheet.conditional_format(after_data_MapXAxisange, {'type': 'formula',
                                                                       'criteria': Layer1_data_MapXAxis + ' <> ' + Layer2_data_MapXAxis,
                                                                       'format': FormatForAfter_NotEqual})

                for m in range(len(MapValue1_Y)):
                    worksheet.write(StartPoint1_Y + m + 2, StartPoint1_x - 1, MapValue1_Y[m], data_format_axis)
                    worksheet.write(StartPoint1_Y + m + 2, StartPoint1_x - 1 + len(MapValue1_X) + 3,
                                    MapValue2_Y_lay2[m],
                                    data_format_axis)

                    if MapValue1_Y[m] != MapValue2_Y_lay2[m]:
                        result = 1
                    else:
                        pass

                Layer1_data_MapYAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 2, StartPoint1_x - 1)
                Layer2_data_MapYAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 2,
                                                                            StartPoint1_x - 1 + len(MapValue1_X) + 3)

                Before_data_MapYAxis = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 2,
                                                                            StartPoint1_x - 1 + 3 + (
                                                                                    len(MapValue1_X) * 2) + 3)

                after_data_MapYAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 2,
                                                                          StartPoint1_x - 1 + len(MapValue1_X) + 3,
                                                                          StartPoint1_Y + 2 + m,
                                                                          StartPoint1_x - 1 + len(MapValue1_X) + 3)

                Before_data_MapYAxisange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 2,
                                                                           StartPoint1_x - 1 + 3 + (
                                                                                   len(MapValue1_X) * 2) + 3,
                                                                           StartPoint1_Y + 2 + m,
                                                                           StartPoint1_x - 1 + 3 + (
                                                                                   len(MapValue1_X) * 2) + 3)

                worksheet.conditional_format(after_data_MapYAxisange, {'type': 'formula',
                                                                       'criteria': Layer1_data_MapYAxis + ' <> ' + Layer2_data_MapYAxis,
                                                                       'format': FormatForAfter_NotEqual})

                for p in range(len(MapValue1_Y)):
                    for o in range(len(MapValue1_X)):
                        worksheet.write(StartPoint1_Y + 2 + p, o + StartPoint1_x, MapValue1_Value[p][o],
                                        data_format_data)
                        worksheet.write(StartPoint1_Y + 2 + p, o + StartPoint1_x + len(MapValue1_X) + 3,
                                        MapValue1_Value_lay2[p][o], data_format_data)

                        if MapValue1_Value[p][o] != MapValue1_Value_lay2[p][o]:
                            result = 1
                        else:
                            pass

                Layer1_data_MapData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 2, StartPoint1_x)
                Layer2_data_MapData = xlsxwriter.utility.xl_rowcol_to_cell(StartPoint1_Y + 2,
                                                                           StartPoint1_x + len(MapValue1_X) + 3)

                after_data_MapDatasange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 2,
                                                                          StartPoint1_x + len(MapValue1_X) + 3,
                                                                          StartPoint1_Y + 2 + p,
                                                                          o + StartPoint1_x + len(MapValue1_X) + 3)

                Before_data_MapDatasange = xlsxwriter.utility.xl_range_abs(StartPoint1_Y + 2,
                                                                           StartPoint1_x + (len(MapValue1_X) * 2) + 6,
                                                                           StartPoint1_Y + 2 + p,
                                                                           o + StartPoint1_x + (
                                                                                   len(MapValue1_X) * 2) + 6)

                worksheet.conditional_format(after_data_MapDatasange, {'type': 'formula',
                                                                       'criteria': Layer1_data_MapData + ' <> ' + Layer2_data_MapData,
                                                                       'format': FormatForAfter_NotEqual})
        except KeyError:
            result = 2

        return [result, StartPoint1_x, StartPoint1_Y]

    def __drawdecision__(self,target_unit,result,YPos):

        workbook = self.workbook
        self.current_sheet = workbook.get_worksheet_by_name("Summary")
        worksheet = self.current_sheet

        WriteTitleFormat = workbook.add_format({'font_color': '#000000', 'bold': True, 'font_size': 11,'border': 1})
        WriteNotEqualFormat = workbook.add_format({'bold': True, 'bg_color': 'red','align': 'center','valign': 'vcenter','border': 1})
        WriteEmptyFormat = workbook.add_format({'bold': True, 'bg_color': 'gray','align': 'center','valign': 'vcenter','border': 1})
        WriteEqualFormat = workbook.add_format({'bold': True, 'bg_color': 'green','align': 'center','valign': 'vcenter','border': 1})
        WriteLinkFormat = workbook.add_format({'font_color': 'red', 'bold': True, 'font_size': 11,'border': 1})

        # 해당 unit sheet로 이동 하이퍼링크 코드 구현
        worksheet.write_url(YPos+1, 1, 'internal:{}!A1'.format(target_unit))

        if result == 1:
            worksheet.write(YPos+1, 1, target_unit, WriteTitleFormat)
            worksheet.write(YPos+1, 2,"FAIL",WriteNotEqualFormat)
            worksheet.write(YPos+1, 3,None,WriteTitleFormat)
        elif result == 2:
            worksheet.write(YPos + 1, 1, target_unit, WriteTitleFormat)
            worksheet.write(YPos + 1, 2, "N/A", WriteEmptyFormat)
            worksheet.write(YPos + 1, 3, None, WriteTitleFormat)
        elif result == 3:
            worksheet.write(YPos + 1, 1, target_unit, WriteTitleFormat)
            worksheet.write(YPos + 1, 2, "Not Matching", WriteEmptyFormat)
            worksheet.write(YPos + 1, 3, None, WriteTitleFormat)

        else:
            worksheet.write(YPos+ 1, 1, target_unit, WriteTitleFormat)
            worksheet.write(YPos+ 1, 2, "PASS", WriteEqualFormat)
            worksheet.write(YPos + 1, 3, None, WriteTitleFormat)

        #해당 worksheet에서 Summary로 이동 하이퍼링크 생성
        self.current_ws = workbook.get_worksheet_by_name("{}".format(target_unit))
        ws = self.current_ws
        ws.write_url(0, 0, 'internal:{}!A1'.format('Summary'))
        ws.write(0, 0, 'Return To Summary', WriteLinkFormat)

        return YPos+1

    def __drawdecision2__(self,target_unit,result,YPos2):

        workbook = self.workbook

        self.current_sheet = workbook.get_worksheet_by_name("Summary")
        worksheet = self.current_sheet
        WriteTitleFormat = workbook.add_format({'font_color': '#000000', 'bold': True, 'font_size': 11,'border': 1})
        WriteNotChangeFormat = workbook.add_format({'bold': True, 'bg_color': '#CCCC99','align': 'center','valign': 'vcenter','border': 1})
        WriteEmptyFormat = workbook.add_format({'bold': True, 'bg_color': 'gray','align': 'center','valign': 'vcenter','border': 1})
        WriteChangeFormat = workbook.add_format({'bold': True, 'bg_color': 'yellow','align': 'center','valign': 'vcenter','border': 1})
        WriteNotCompareFormat = workbook.add_format({'bold': True, 'bg_color': '#FFFFFF','align': 'center','valign': 'vcenter','border': 1})


        if result == 1: #다르면
            worksheet.write(YPos2 + 1, 1,target_unit, WriteTitleFormat)
            worksheet.write(YPos2 + 1, 3,"Change",WriteChangeFormat)
            worksheet.write(YPos2 + 1, 4,None,WriteTitleFormat)
        elif result == 2:
            worksheet.write(YPos2 + 1, 1, target_unit, WriteTitleFormat)
            worksheet.write(YPos2 + 1, 3, "N/A", WriteEmptyFormat)
            worksheet.write(YPos2 + 1, 4, None, WriteTitleFormat)
        elif result == 3:
            worksheet.write(YPos2 + 1, 1, target_unit, WriteTitleFormat)
            worksheet.write(YPos2 + 1, 3, "N/A", WriteEmptyFormat)
            worksheet.write(YPos2 + 1, 4, None, WriteTitleFormat)
        elif result == -1:
            worksheet.write(YPos2 + 1, 3, "Not Compare", WriteNotCompareFormat)
            worksheet.write(YPos2 + 1, 4, None, WriteTitleFormat)

        else:
            worksheet.write(YPos2 + 1, 1, target_unit, WriteTitleFormat)
            worksheet.write(YPos2 + 1, 3, "Not Change", WriteNotChangeFormat)
            worksheet.write(YPos2 + 1, 4, None, WriteTitleFormat)

        return YPos2+1

    def __summarytext__(self,sheet_name,NewBase,TuningFile,BeforeBase):

        Newpath = ""

        workbook = self.workbook
        self.current_sheet = workbook.get_worksheet_by_name(sheet_name)
        worksheet = self.current_sheet

        WriteTitleFormat = workbook.add_format({'font_color': '#000000', 'bold': True, 'font_size': 12,'align': 'center',
            'valign': 'vcenter','border': 1})

        worksheet.set_column('B:B',40)
        worksheet.set_column('C:C',50)
        worksheet.set_column('D:D',50)
        worksheet.set_column('E:E',30)
        worksheet.write(1,1,'Author',WriteTitleFormat)
        worksheet.write(1,2, getpass.getuser(),WriteTitleFormat)
        worksheet.write(2,1,'New BaseLine File',WriteTitleFormat)
        worksheet.write(2,2,NewBase,WriteTitleFormat)
        worksheet.write(3,1,'Tuning File',WriteTitleFormat)
        worksheet.write(3,2, TuningFile, WriteTitleFormat)
        worksheet.write(4,1,'Before BaseLine File',WriteTitleFormat)
        worksheet.write(4,2, BeforeBase, WriteTitleFormat)
        worksheet.write(6,4,'비    고',WriteTitleFormat)
        worksheet.write(6,2,'New Baseline vs Tuning', WriteTitleFormat)
        worksheet.write(6,3,'New Baseline vs Before Baseline', WriteTitleFormat)


class ReportGenr:
    def __init__(self, importer1, importer2, workbook):
        self.importer1 = importer1
        self.importer2 = importer2
        self.workbook = workbook
        self.current_sheet = None

    def __setvariant__(self, variant_name1, variant_name2):
        self.variant_name1 = variant_name1
        self.variant_name2 = variant_name2

    def __createsheet__(self, sheet_name):
        workbook = self.workbook
        self.sheet_name = sheet_name
        self.current_sheet = workbook.get_worksheet_by_name(sheet_name)
        if self.current_sheet is None:
            self.current_sheet = workbook.add_worksheet(sheet_name)

    def __drawHeader__(self, x, y, xsize, ysize, Version, zoom):

        workbook = self.workbook
        worksheet = self.current_sheet
        worksheet.hide_gridlines(2)
        worksheet.set_zoom(zoom)

        temprange1 = xlsxwriter.utility.xl_range(y, x, y + ysize, x + xsize)
        temprange2 = xlsxwriter.utility.xl_range(y, x + xsize + 2, y + ysize, x + xsize + 2 + xsize)
        merge_format1 = workbook.add_format({
            'bold': 1,
            'border': 2,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': '#D7E4BC',
            'font_size': 20})

        worksheet.merge_range(temprange1, "Normal", merge_format1)
        worksheet.merge_range(temprange2, "Sports", merge_format1)

        worksheet.set_column(0, 0, 4)
        worksheet.set_column(1, 1 + xsize - 2, 8.3)
        worksheet.set_column(1 + xsize - 2 + 3, 1 + xsize - 2 + 3, 0.8)
        worksheet.set_column(1 + xsize - 2 + 3 + 1, 1 + xsize - 2 + 3 + 1 + xsize + 1, 8.3)

        LeftBorder_format = workbook.add_format({'left': 2})
        worksheet.set_column(1, 1, 7.25, LeftBorder_format)
        worksheet.set_column(1 + xsize - 2 + 3, 1 + xsize - 2 + 3, 0.8, LeftBorder_format)
        worksheet.set_column(1 + xsize - 2 + 3 + 1, 1 + xsize - 2 + 3 + 1, 7.25, LeftBorder_format)
        worksheet.set_column(1 + xsize - 2 + 3 + 1 + 1+ xsize, 1 + xsize - 2 + 3 + 1 + 1+ xsize, 7.25, LeftBorder_format)

    def __drawMap__(self, target_mapname, Lastposn, MapWidth, draw_info):

        GRADENA = draw_info['GradientEna']
        PlauENA = draw_info['PlausibilityCheckEna']
        ChartName = draw_info['Title']
        XaxisName = draw_info['Xlabel']
        YaxisName = draw_info['Ylabel']
        if GRADENA is True:
            GRADChartName = draw_info['GRADTitle']
            GRADXaxisName = draw_info['GRADXlabel']
            GRADYaxisName = draw_info['GRADYlabel']

        workbook = self.workbook
        worksheet = self.current_sheet

        MapValue1 = self.importer1.__global_prms__[self.variant_name1]["MAP"][target_mapname]
        MapValue1_X = self.importer1.__global_prms__[self.variant_name1]["COM_AXIS"][MapValue1.__xaxis__]
        MapValue1_Y = self.importer1.__global_prms__[self.variant_name1]["COM_AXIS"][MapValue1.__yaxis__]

        MapValue2 = self.importer2.__global_prms__[self.variant_name2]["MAP"][target_mapname]
        MapValue2_X = self.importer2.__global_prms__[self.variant_name2]["COM_AXIS"][MapValue2.__xaxis__]
        MapValue2_Y = self.importer2.__global_prms__[self.variant_name2]["COM_AXIS"][MapValue2.__yaxis__]

        x_axis1 = MapValue1_X.value
        y_axis1 = MapValue1_Y.value
        values1 = MapValue1.value
        x_size1 = len(x_axis1)
        y_size1 = len(y_axis1)

        x_axis2 = MapValue2_X.value
        y_axis2 = MapValue2_Y.value
        values2 = MapValue2.value
        x_size2 = len(x_axis2)
        y_size2 = len(y_axis2)

        startpoint1_y = Lastposn
        startpoint1_x = 2
        pos1 = [startpoint1_y, startpoint1_x]


        startpoint2_y = startpoint1_y
        startpoint2_x = 2 + x_size1 + 4

        startpoint2_x_old = MapWidth
        if startpoint2_x_old >= startpoint2_x:
            startpoint2_x = startpoint2_x_old

        else:
            startpoint2_x = startpoint2_x

        MapWidth = startpoint2_x


        pos2 = [startpoint1_y, startpoint2_x]

        WriteTitleFormat = workbook.add_format({'font_color': '#000000', 'bold': True, 'font_size': 20})
        worksheet.write(startpoint1_y - 2, startpoint1_x, target_mapname, WriteTitleFormat)
        worksheet.write(startpoint2_y - 2, startpoint2_x, target_mapname, WriteTitleFormat)


        # MAP1
        data_format_data = workbook.add_format({'border': 1})
        data_format_axis = workbook.add_format({'border': 1, 'bg_color': '#DCE6F1'})
        for col in range(x_size1):
            worksheet.write(startpoint1_y, col + 1 + startpoint1_x, x_axis1[col], data_format_axis)

        for row in range(y_size1):
            worksheet.write(row + 1 + startpoint1_y, startpoint1_x, y_axis1[row], data_format_axis)

        for row in range(y_size1):
            for col in range(x_size1):
                worksheet.write(row + startpoint1_y + 1, col + startpoint1_x + 1, values1[row][col], data_format_data)

        # MAP2
        for col in range(x_size2):
            worksheet.write(startpoint2_y, col + 1 + startpoint2_x, x_axis2[col], data_format_axis)

        for row in range(y_size2):
            worksheet.write(row + 1 + startpoint2_y, startpoint2_x, y_axis2[row], data_format_axis)

        for row in range(y_size2):
            for col in range(x_size2):
                worksheet.write(row + 1 + startpoint2_y, col + startpoint2_x + 1, values2[row][col], data_format_data)

        # 조건부서식
        FormatForAfter_Less = workbook.add_format({'font_color': '#000000', 'bg_color': '#EBF1DE', 'bold': True})
        FormatForAfter_Greater = workbook.add_format({'font_color': '#000000', 'bg_color': '#F2DCDB', 'bold': True})
        FormatForAfter_Equal = workbook.add_format({'font_color': '#000000', 'bg_color': '#FFFFFF'})
        FormatForAfter_AxisEqual = workbook.add_format({'font_color': '#000000', 'bg_color': '#DCE6F1'})

        # 조건부서식-MAP
        before_data_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y + 1, startpoint1_x + 1)

        before_xaxis_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y, startpoint1_x + 1)
        before_yaxis_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y + 1, startpoint1_x)

        before_data_range = xlsxwriter.utility.xl_range_abs(startpoint1_y + 1, startpoint1_x + 1, startpoint1_y + y_size1, startpoint1_x + x_size1)

        before_xaxis_range = xlsxwriter.utility.xl_range_abs(startpoint1_y, startpoint1_x + 1, startpoint1_y, startpoint1_x + x_size2)
        before_yaxis_range = xlsxwriter.utility.xl_range_abs(startpoint1_y + 1, startpoint1_x, startpoint1_y + y_size1, startpoint1_x)

        after_data_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint2_y + 1, startpoint2_x + 1)

        after_xaxis_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint2_y, startpoint2_x + 1)
        after_yaxis_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint2_y + 1, startpoint2_x)
        after_data_range = xlsxwriter.utility.xl_range_abs(startpoint2_y + 1, startpoint2_x + 1,
                                                           startpoint2_y + y_size2, startpoint2_x + x_size2)
        after_xaxis_range = xlsxwriter.utility.xl_range_abs(startpoint2_y, startpoint2_x + 1, startpoint2_y,
                                                            startpoint2_x + x_size2)
        after_yaxis_range = xlsxwriter.utility.xl_range_abs(startpoint2_y + 1, startpoint2_x, startpoint2_y + y_size2,
                                                            startpoint2_x)

        worksheet.conditional_format(after_data_range, {'type': 'formula',
                                                        'criteria': before_data_sp + ' > ' + after_data_sp,
                                                        'format': FormatForAfter_Less})
        worksheet.conditional_format(after_data_range, {'type': 'formula',
                                                        'criteria': before_data_sp + ' < ' + after_data_sp,
                                                        'format': FormatForAfter_Greater})
        worksheet.conditional_format(after_data_range, {'type': 'formula',
                                                        'criteria': before_data_sp + ' = ' + after_data_sp,
                                                        'format': FormatForAfter_Equal})

        worksheet.conditional_format(after_xaxis_range, {'type': 'formula',
                                                         'criteria': before_xaxis_sp + ' > ' + after_xaxis_sp,
                                                         'format': FormatForAfter_Less})
        worksheet.conditional_format(after_xaxis_range, {'type': 'formula',
                                                         'criteria': before_xaxis_sp + ' < ' + after_xaxis_sp,
                                                         'format': FormatForAfter_Greater})
        worksheet.conditional_format(after_xaxis_range, {'type': 'formula',
                                                         'criteria': before_xaxis_sp + ' = ' + after_xaxis_sp,
                                                         'format': FormatForAfter_AxisEqual})

        worksheet.conditional_format(after_yaxis_range, {'type': 'formula',
                                                         'criteria': before_yaxis_sp + ' > ' + after_yaxis_sp,
                                                         'format': FormatForAfter_Less})
        worksheet.conditional_format(after_yaxis_range, {'type': 'formula',
                                                         'criteria': before_yaxis_sp + ' < ' + after_yaxis_sp,
                                                         'format': FormatForAfter_Greater})
        worksheet.conditional_format(after_yaxis_range, {'type': 'formula',
                                                         'criteria': before_yaxis_sp + ' = ' + after_yaxis_sp,
                                                         'format': FormatForAfter_AxisEqual})

        # MAP차트 구성
        chart1 = workbook.add_chart(dict(type='scatter', subtype='straight_with_markers'))
        for i in range(y_size1):
            row = i + 1
            chart1.add_series({
                'name': [self.sheet_name, pos1[0] + row, pos1[1]],
                'categories': [self.sheet_name, pos1[0], pos1[1] + 1, pos1[0], pos1[1] + x_size1],
                'values': [self.sheet_name, pos1[0] + row, pos1[1] + 1, pos1[0] + row, pos1[1] + x_size1],
                'marker': {'type': 'automatic', 'size': 1, },
                'line': {'width': 2, }
            })
        chart1.set_title({'name': ChartName, 'font_size': 20})
        chart1.set_x_axis({'name': XaxisName, 'name_font': {'size': 18},
                           'major_gridlines': {'visible': True, 'line': {'width': 1, 'dash_type': 'dash'}}})
        chart1.set_y_axis({'name': YaxisName, 'name_font': {'size': 18}})

        chart2 = workbook.add_chart(dict(type='scatter', subtype='straight_with_markers'))
        for i in range(y_size2):
            row = i + 1
            chart2.add_series({
                'name': [self.sheet_name, pos2[0] + row, pos2[1]],
                'categories': [self.sheet_name, pos2[0], pos2[1] + 1, pos2[0], pos2[1] + x_size1],
                'values': [self.sheet_name, pos2[0] + row, pos2[1] + 1, pos2[0] + row, pos2[1] + x_size1],
                'marker': {'type': 'automatic', 'size': 1, },
                'line': {'width': 2, }
            })
        chart2.set_title({'name': ChartName, 'font_size': 20})
        chart2.set_x_axis({'name': XaxisName, 'name_font': {'size': 18},
                           'major_gridlines': {'visible': True, 'line': {'width': 1, 'dash_type': 'dash'}}})
        chart2.set_y_axis({'name': YaxisName, 'name_font': {'size': 18}})
        '''
        chart2.set_x_axis({'name': 'STW Angle[deg]', 'name_font': {'size': 18}, 'max': 300,
                           'major_gridlines': {'visible': True, 'line': {'width': 1, 'dash_type': 'dash'}}})
        '''

        # 차트 배치 및 사이즈
        if GRADENA is True:
            before_chart1_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y + 2 * (y_size1 + 2), startpoint1_x)
            after_chart2_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y + 2 * (y_size1 + 2), startpoint2_x)
        else:
            before_chart1_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y + (y_size1 + 2), startpoint1_x)
            after_chart2_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y + (y_size1 + 2), startpoint2_x)

        worksheet.insert_chart(before_chart1_sp, chart1)
        chart1.set_size({'width': 1324, 'height': 400})
        worksheet.insert_chart(after_chart2_sp, chart2)
        chart2.set_size({'width': 1323, 'height': 400})

        # GRADIENTMAP1
        GRADvalue1 = np.zeros([y_size1, x_size1])
        for row in range(y_size1):
            for col in range(x_size1):
                if target_mapname == 'KinematMap':
                    if col < x_size1 - 1:
                        GRADvalue1[row][col] = (values1[row][col + 1] - values1[row][col]) / (x_axis1[col + 1] - x_axis1[col]) * 360
                else:
                    if col == 0:
                        GRADvalue1[row][col] = 0
                    else:
                        GRADvalue1[row][col] = (values1[row][col] - values1[row][col-1]) / (x_axis1[col] - x_axis1[col-1])
        # GRADIENTMAP2
        GRADvalue2 = np.zeros([y_size2, x_size2])
        for row in range(y_size2):
            for col in range(x_size2):
                if target_mapname == 'KinematMap':
                    if col < x_size2 - 1:
                        GRADvalue2[row][col] = (values2[row][col + 1] - values2[row][col]) / ( x_axis2[col + 1] - x_axis2[col]) * 360
                else:
                    if col == 0:
                        GRADvalue2[row][col] = 0
                    else:
                        GRADvalue2[row][col] = (values2[row][col] - values2[row][col-1]) / (x_axis2[col] - x_axis2[col-1])
        if GRADENA is True:
           # GRADIENTMAP1
           for col in range(x_size1):
               worksheet.write(startpoint1_y + y_size1 + 2, col + 1 + startpoint1_x, x_axis1[col], data_format_axis)
           for row in range(y_size1):
               worksheet.write(row + 1 + startpoint1_y + y_size1 + 2, startpoint1_x, y_axis1[row], data_format_axis)

           for row in range(y_size1):
               for col in range(x_size1):
                   worksheet.write(row + 1 + startpoint1_y + y_size1 + 2, col + startpoint1_x + 1, GRADvalue1[row][col], data_format_data)
          # GRADIENTMAP2
           for col in range(x_size2):
               worksheet.write(startpoint2_y + y_size1 + 2, col + 1 + startpoint2_x, x_axis2[col], data_format_axis)

           for row in range(y_size2):
               worksheet.write(row + 1 + startpoint2_y + y_size1 + 2, startpoint2_x, y_axis2[row], data_format_axis)

           for row in range(y_size2):
               for col in range(x_size2):
                   worksheet.write(row + 1 + startpoint2_y + y_size1 + 2, col + startpoint2_x + 1, GRADvalue2[row][col], data_format_data)

            # 조건부서식-GRADIENT
           before_GRADdata_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y + 1 + y_size1 + 2, startpoint1_x + 1)
           before_GRADxaxis_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y + y_size1 + 2, startpoint1_x + 1)
           before_GRADyaxis_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y + 1 + y_size1 + 2, startpoint1_x)
           before_GRADdata_range = xlsxwriter.utility.xl_range_abs(startpoint1_y + 1 + y_size1 + 2, startpoint1_x + 1, startpoint1_y + y_size2 + y_size1 + 2, startpoint1_x + x_size2)
           before_GRADxaxis_range = xlsxwriter.utility.xl_range_abs(startpoint1_y + y_size2 + 2, startpoint1_x + 1, startpoint1_y + y_size1 + 2, startpoint1_x + x_size2)
           before_GRADyaxis_range = xlsxwriter.utility.xl_range_abs(startpoint1_y + 1 + y_size2 + 2, startpoint1_x, startpoint1_y + y_size2 + y_size1 + 2,
                                                                   startpoint1_x)

           after_GRADdata_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint2_y + 1 + y_size1 + 2, startpoint2_x + 1)
           after_GRADxaxis_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint2_y + y_size1 + 2, startpoint2_x + 1)
           after_GRADyaxis_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint2_y + 1 + y_size1 + 2, startpoint2_x)
           after_GRADdata_range = xlsxwriter.utility.xl_range_abs(startpoint2_y + 1 + y_size1 + 2, startpoint2_x + 1, startpoint2_y + y_size1 + y_size2 + 2, startpoint2_x + x_size2)
           after_GRADxaxis_range = xlsxwriter.utility.xl_range_abs(startpoint2_y + y_size1 + 2, startpoint2_x + 1, startpoint2_y + y_size1 + 2, startpoint2_x + x_size2)
           after_GRADyaxis_range = xlsxwriter.utility.xl_range_abs(startpoint2_y + 1 + y_size1 + 2, startpoint2_x, startpoint2_y + y_size1 + y_size2 + 2,
                                                                   startpoint2_x)

           worksheet.conditional_format(after_GRADdata_range, {'type': 'formula',
                                                               'criteria': before_GRADdata_sp + ' > ' + after_GRADdata_sp,
                                                               'format': FormatForAfter_Less})
           worksheet.conditional_format(after_GRADdata_range, {'type': 'formula',
                                                                'criteria': before_GRADdata_sp + ' < ' + after_GRADdata_sp,
                                                                'format': FormatForAfter_Greater})
           worksheet.conditional_format(after_GRADdata_range, {'type': 'formula',
                                                                'criteria': before_GRADdata_sp + ' = ' + after_GRADdata_sp,
                                                                'format': FormatForAfter_Equal})

           worksheet.conditional_format(after_GRADxaxis_range, {'type': 'formula',
                                                                 'criteria': before_GRADxaxis_sp + ' > ' + after_GRADxaxis_sp,
                                                                 'format': FormatForAfter_Less})
           worksheet.conditional_format(after_GRADxaxis_range, {'type': 'formula',
                                                                 'criteria': before_GRADxaxis_sp + ' < ' + after_GRADxaxis_sp,
                                                                 'format': FormatForAfter_Greater})
           worksheet.conditional_format(after_GRADxaxis_range, {'type': 'formula',
                                                                 'criteria': before_GRADxaxis_sp + ' = ' + after_GRADxaxis_sp,
                                                                 'format': FormatForAfter_AxisEqual})
           worksheet.conditional_format(after_GRADyaxis_range, {'type': 'formula',
                                                                 'criteria': before_GRADyaxis_sp + ' > ' + after_GRADyaxis_sp,
                                                                 'format': FormatForAfter_Less})
           worksheet.conditional_format(after_GRADyaxis_range, {'type': 'formula',
                                                                 'criteria': before_GRADyaxis_sp + ' < ' + after_GRADyaxis_sp,
                                                                 'format': FormatForAfter_Greater})
           worksheet.conditional_format(after_GRADyaxis_range, {'type': 'formula',
                                                                 'criteria': before_GRADyaxis_sp + ' = ' + after_GRADyaxis_sp,
                                                                 'format': FormatForAfter_AxisEqual})
            # GRAD차트 구성
           GRADpos1 = [startpoint1_y + y_size1 + 2, startpoint1_x]
           GRADpos2 = [startpoint1_y + y_size1 + 2, startpoint2_x]

           GRADchart1 = workbook.add_chart(dict(type='scatter', subtype='straight_with_markers'))
           for i in range(y_size1):
                row = i + 1
                GRADchart1.add_series({
                    'name': [self.sheet_name, GRADpos1[0] + row, GRADpos1[1]],
                    'categories': [self.sheet_name, GRADpos1[0], GRADpos1[1] + 1, GRADpos1[0], GRADpos1[1] + x_size1],
                    'values': [self.sheet_name, GRADpos1[0] + row, GRADpos1[1] + 1, GRADpos1[0] + row,
                               GRADpos1[1] + x_size1],
                    'marker': {'type': 'automatic', 'size': 1, },
                    'line': {'width': 2, }
                })
           GRADchart1.set_title({'name': GRADChartName, 'font_size': 20})
           GRADchart1.set_x_axis({'name': GRADXaxisName, 'name_font': {'size': 18},
                                   'major_gridlines': {'visible': True, 'line': {'width': 1, 'dash_type': 'dash'}}})
           GRADchart1.set_y_axis({'name': GRADYaxisName, 'name_font': {'size': 18}})

           GRADchart2 = workbook.add_chart(dict(type='scatter', subtype='straight_with_markers'))
           for i in range(y_size2):
                row = i + 1
                GRADchart2.add_series({
                    'name': [self.sheet_name, GRADpos2[0] + row, GRADpos2[1]],
                    'categories': [self.sheet_name, GRADpos2[0], GRADpos2[1] + 1, GRADpos2[0], GRADpos2[1] + x_size1],
                    'values': [self.sheet_name, GRADpos2[0] + row, GRADpos2[1] + 1, GRADpos2[0] + row,
                               GRADpos2[1] + x_size1],
                    'marker': {'type': 'automatic', 'size': 1, },
                    'line': {'width': 2, }
                })
           GRADchart2.set_title({'name': GRADChartName, 'font_size': 20})
           GRADchart2.set_x_axis({'name': GRADXaxisName, 'name_font': {'size': 18},
                                   'major_gridlines': {'visible': True, 'line': {'width': 1, 'dash_type': 'dash'}}})
           GRADchart2.set_y_axis({'name': GRADYaxisName, 'name_font': {'size': 18}})

            # 차트 배치 및 사이즈
           before_GRADchart1_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y + 2 * (y_size1 + 2) + 21,
                                                                        startpoint1_x)
           after_GRADchart2_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y + 2 * (y_size1 + 2) + 21,
                                                                       startpoint2_x)
           worksheet.insert_chart(before_GRADchart1_sp, GRADchart1)
           GRADchart1.set_size({'width': 1324, 'height': 400})
           worksheet.insert_chart(after_GRADchart2_sp, GRADchart2)
           GRADchart2.set_size({'width': 1323, 'height': 400})
        else:
           before_GRADdata_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y + 1 + y_size1 + 2,
                                                                      startpoint1_x + 1)
           before_GRADdata_range = xlsxwriter.utility.xl_range_abs(startpoint1_y + 1 + y_size1 + 2, startpoint1_x + 1,
                                                                    startpoint1_y + y_size2 + y_size1 + 2,
                                                                    startpoint2_x + x_size2)
           after_GRADdata_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint2_y + 1 + y_size2 + 2, startpoint2_x + 1)
           after_GRADdata_range = xlsxwriter.utility.xl_range_abs(startpoint2_y + 1 + y_size2 + 2, startpoint2_x + 1,
                                                                   startpoint2_y + y_size2 + y_size1 + 2,
                                                                   startpoint2_x + x_size2)

        merge_format2 = workbook.add_format({
            'bold': 1,
            'border': 2,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 20})


        PC = PlausibilityCheck()
        if PlauENA is True:
            PC.__PlauCheck__(target_mapname, xlsxwriter, workbook, worksheet, GRADENA, startpoint1_x, startpoint1_y, startpoint2_x, startpoint2_y, x_size1, y_size1, x_size2, y_size2)

        if GRADENA is True:
            Commentrange1 = xlsxwriter.utility.xl_range(startpoint1_y + 2 * (y_size1 + 2) + 42, startpoint1_x - 1,
                                                        startpoint1_y + 2 * (y_size1 + 2) + 47, startpoint1_x - 1 + MapWidth - 4)
            Commentrange2 = xlsxwriter.utility.xl_range(startpoint1_y + 2 * (y_size1 + 2) + 42, startpoint2_x - 1,
                                                        startpoint1_y + 2 * (y_size1 + 2) + 47, startpoint2_x - 1 + 4)
            Commentrange3 = xlsxwriter.utility.xl_range(startpoint1_y + 2 * (y_size1 + 2) + 42, startpoint2_x - 1 + 5,
                                                        startpoint1_y + 2 * (y_size1 + 2) + 47, startpoint2_x - 1 + MapWidth - 4)
            worksheet.merge_range(Commentrange1, '', merge_format2)
            worksheet.merge_range(Commentrange2, 'Comment', merge_format2)
            worksheet.merge_range(Commentrange3, '', merge_format2)
        else:
            Commentrange1 = xlsxwriter.utility.xl_range(startpoint1_y + (y_size1 + 2) + 21, startpoint1_x - 1,
                                                        startpoint1_y + (y_size1 + 2) + 26, startpoint1_x - 1 + MapWidth - 4)
            Commentrange2 = xlsxwriter.utility.xl_range(startpoint1_y + (y_size1 + 2) + 21, startpoint2_x - 1,
                                                        startpoint1_y + (y_size1 + 2) + 26, startpoint2_x - 1 + 4)
            Commentrange3 = xlsxwriter.utility.xl_range(startpoint1_y + (y_size1 + 2) + 21, startpoint2_x - 1 + 5,
                                                        startpoint1_y + (y_size1 + 2) + 26, startpoint2_x - 1 + MapWidth - 4)
            worksheet.merge_range(Commentrange1, '', merge_format2)
            worksheet.merge_range(Commentrange2, 'Comment', merge_format2)
            worksheet.merge_range(Commentrange3, '', merge_format2)

        if GRADENA is True:
            LastYposn = startpoint1_y + 2 * (y_size1 + 2 + 21) + 9
        else:
            LastYposn = startpoint1_y + (y_size1 + 2 + 21) + 9
        return LastYposn, MapWidth

    def __drawCur__(self, target_curname, Lastposn, MapWidth, draw_info):
        workbook = self.workbook
        worksheet = self.current_sheet

        ChartName = draw_info['Title']
        XaxisName = draw_info['Xlabel']
        YaxisName = draw_info['Ylabel']

        CurValue1 = self.importer1.__global_prms__[self.variant_name1]["CURVE"][target_curname]
        CurValue1_X = self.importer1.__global_prms__[self.variant_name1]["COM_AXIS"][CurValue1.__xaxis__]

        CurValue2 = self.importer2.__global_prms__[self.variant_name2]["CURVE"][target_curname]
        CurValue2_X = self.importer2.__global_prms__[self.variant_name2]["COM_AXIS"][CurValue2.__xaxis__]

        x_axis1 = CurValue1_X.value
        values1 = CurValue1.value
        x_size1 = len(x_axis1)
        x_axis2 = CurValue2_X.value
        values2 = CurValue2.value
        x_size2 = len(x_axis2)

        startpoint1_y = Lastposn
        startpoint1_x = 2
        pos1 = [startpoint1_y, startpoint1_x]
        startpoint2_y = startpoint1_y
        startpoint2_x = 26

        startpoint2_x_old = MapWidth
        if startpoint2_x_old >= startpoint2_x:
            startpoint2_x = startpoint2_x_old
        else:
            startpoint2_x = startpoint2_x

        MapWidth = startpoint2_x

        pos2 = [startpoint1_y, startpoint2_x]

        WriteTitleFormat = workbook.add_format({'font_color': '#000000', 'bold': True, 'font_size': 20})
        worksheet.write(startpoint1_y - 2, startpoint1_x, target_curname,WriteTitleFormat)
        worksheet.write(startpoint2_y - 2, startpoint2_x, target_curname,WriteTitleFormat)

        # CURVE1
        data_format_data = workbook.add_format({'border': 1})
        data_format_axis = workbook.add_format({'border': 1, 'bg_color': '#DCE6F1'})

        for col in range(x_size1):
            worksheet.write(startpoint1_y, col  + startpoint1_x, x_axis1[col], data_format_axis)
        for col in range(x_size1):
            worksheet.write(startpoint1_y + 1, col + startpoint1_x, values1[col], data_format_data)
        # CURVE2
        for col in range(x_size2):
            worksheet.write(startpoint2_y, col + startpoint2_x, x_axis2[col], data_format_axis)
        for col in range(x_size2):
            worksheet.write(startpoint2_y + 1, col + startpoint2_x, values2[col], data_format_data)

        # 조건부서식
        FormatForAfter_Less = workbook.add_format({'font_color': '#000000', 'bg_color': '#EBF1DE', 'bold': True})
        FormatForAfter_Greater = workbook.add_format({'font_color': '#000000', 'bg_color': '#F2DCDB', 'bold': True})
        FormatForAfter_Equal = workbook.add_format({'font_color': '#000000', 'bg_color': '#FFFFFF'})
        FormatForAfter_AxisEqual = workbook.add_format({'font_color': '#000000', 'bg_color': '#DCE6F1'})

        # 조건부서식-CURVE
        before_data_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y + 1, startpoint1_x )
        before_xaxis_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y, startpoint1_x)

        after_data_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint2_y + 1, startpoint2_x )
        after_xaxis_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint2_y, startpoint2_x)

        after_data_range = xlsxwriter.utility.xl_range_abs(startpoint2_y + 1, startpoint2_x,
                                                           startpoint2_y + 1, startpoint2_x + x_size2 -1)
        after_xaxis_range = xlsxwriter.utility.xl_range_abs(startpoint2_y, startpoint2_x, startpoint2_y,
                                                            startpoint2_x + x_size2 -1)


        worksheet.conditional_format(after_data_range, {'type': 'formula',
                                                        'criteria': before_data_sp + ' > ' + after_data_sp,
                                                        'format': FormatForAfter_Less})
        worksheet.conditional_format(after_data_range, {'type': 'formula',
                                                        'criteria': before_data_sp + ' < ' + after_data_sp,
                                                        'format': FormatForAfter_Greater})
        worksheet.conditional_format(after_data_range, {'type': 'formula',
                                                        'criteria': before_data_sp + ' = ' + after_data_sp,
                                                        'format': FormatForAfter_Equal})

        worksheet.conditional_format(after_xaxis_range, {'type': 'formula',
                                                         'criteria': before_xaxis_sp + ' > ' + after_xaxis_sp,
                                                         'format': FormatForAfter_Less})
        worksheet.conditional_format(after_xaxis_range, {'type': 'formula',
                                                         'criteria': before_xaxis_sp + ' < ' + after_xaxis_sp,
                                                         'format': FormatForAfter_Greater})
        worksheet.conditional_format(after_xaxis_range, {'type': 'formula',
                                                         'criteria': before_xaxis_sp + ' = ' + after_xaxis_sp,
                                                         'format': FormatForAfter_AxisEqual})


        # CURVE차트 구성
        Curchart1 = workbook.add_chart(dict(type='scatter', subtype='straight_with_markers'))

        row = 1
        Curchart1.add_series({
            'name': [self.sheet_name, pos1[0] + row, pos1[1]],
            'categories': [self.sheet_name, pos1[0], pos1[1] + 1, pos1[0], pos1[1] + x_size1],
            'values': [self.sheet_name, pos1[0] + row, pos1[1] + 1, pos1[0] + row, pos1[1] + x_size1],
            'marker': {'type': 'automatic', 'size': 1, },
            'line': {'width': 2, }
        })
        Curchart1.set_title({'name': ChartName, 'font_size': 20})
        Curchart1.set_x_axis({'name': XaxisName, 'name_font': {'size': 18},
                           'major_gridlines': {'visible': True, 'line': {'width': 1, 'dash_type': 'dash'}}})
        Curchart1.set_y_axis({'name': YaxisName, 'name_font': {'size': 18}})
        Curchart1.set_legend({'none': True})

        Curchart2 = workbook.add_chart(dict(type='scatter', subtype='straight_with_markers'))
        Curchart2.add_series({
            'name': [self.sheet_name, pos2[0] + row, pos2[1]],
            'categories': [self.sheet_name, pos2[0], pos2[1] + 1, pos2[0], pos2[1] + x_size2],
            'values': [self.sheet_name, pos2[0] + row, pos2[1] + 1, pos2[0] + row, pos2[1] + x_size2],
            'marker': {'type': 'automatic', 'size': 1, },
            'line': {'width': 2, }
        })
        Curchart2.set_title({'name': ChartName, 'font_size': 20})
        Curchart2.set_x_axis({'name': XaxisName, 'name_font': {'size': 18},
                           'major_gridlines': {'visible': True, 'line': {'width': 1, 'dash_type': 'dash'}}})
        Curchart2.set_y_axis({'name': YaxisName, 'name_font': {'size': 18}})
        Curchart2.set_legend({'none': True})
        before_Curchart1_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y + 3, startpoint1_x)
        after_Curchart2_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y + 3, startpoint2_x)

        worksheet.insert_chart(before_Curchart1_sp, Curchart1)
        Curchart1.set_size({'width': 1324, 'height': 200})
        worksheet.insert_chart(after_Curchart2_sp, Curchart2)
        Curchart2.set_size({'width': 1323, 'height': 200})

        merge_format2 = workbook.add_format({
            'bold': 1,
            'border': 2,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 20})
        Commentrange1 = xlsxwriter.utility.xl_range(startpoint1_y + (2) + 12, startpoint1_x - 1,
                                                    startpoint1_y + (2) + 17, startpoint1_x - 1 + MapWidth - 4)
        Commentrange2 = xlsxwriter.utility.xl_range(startpoint1_y + (2) + 12, startpoint2_x - 1,
                                                    startpoint1_y + (2) + 17, startpoint2_x - 1 + 4)
        Commentrange3 = xlsxwriter.utility.xl_range(startpoint1_y + (2) + 12, startpoint2_x - 1 + 5,
                                                    startpoint1_y + (2) + 17, startpoint2_x - 1 + MapWidth - 4)
        worksheet.merge_range(Commentrange1, '', merge_format2)
        worksheet.merge_range(Commentrange2, 'Comment', merge_format2)
        worksheet.merge_range(Commentrange3, '', merge_format2)

        LastYposn = startpoint1_y + 23
        return LastYposn, MapWidth

class GetVariant:
    def __init__(self, importer):
        self.importer = importer

    def __setvariant__(self, variant_name1):
        self.variant_name1 = variant_name1

    def __getname__(self):

        if self.variant_name1 in self.importer.__local_prms__.keys():
            return self.variant_name1

        if self.variant_name1 in self.importer.__global_prms__.keys():
            return self.variant_name1

        return 0
        # try:
        #     try3 = self.importer1.__local_prms__[self.variant_name1]["MAP"]
        #     result = self.variant_name1
        #     return result
        # except KeyError:
        #     result = 0
        #
        # try:
        #     try3 = self.importer1.__global_prms__[self.variant_name1]["MAP"]
        #     result = self.variant_name1
        #     return result
        # except KeyError:
        #     result = 0
        #
        # try:
        #     try1 = self.importer1.__local_prms__[self.variant_name1]["VALUE"]
        #     result = self.variant_name1
        #     return result
        # except KeyError:
        #     result = 0
        #
        # try:
        #     try1 = self.importer1.__global_prms__[self.variant_name1]["VALUE"]
        #     result = self.variant_name1
        #     return result
        # except KeyError:
        #     result = 0
        #
        # try:
        #     try1 = self.importer1.__etc_prms__[self.variant_name1]["VALUE"]
        #     result = self.variant_name1
        #     return result
        # except KeyError:
        #     result = 0
        #
        # try:
        #     try2 = self.importer1.__local_prms__[self.variant_name1]["CURVE"]
        #     result = self.variant_name1
        #     return result
        # except KeyError:
        #     result = 0
        #
        # try:
        #     try2 = self.importer1.__global_prms__[self.variant_name1]["CURVE"]
        #     result = self.variant_name1
        #     return result
        # except KeyError:
        #     result = 0
        #
        # return result








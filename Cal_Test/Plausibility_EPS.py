import numpy as np

##Definition
DampingTq_Max = '10'
DampingTq_Min = '0'
DampingTq_GradMax = '0.2'

EffortTq_Max = '10'
EffortTq_Min = '0'
EffortTq_GradMax = '0.6'

RackFTq_Max = '10'
RackFTq_Min = '0'
RackFTq_GradMax = '0.167/80*2'

ReturnAgSpd_Max = '1000'
ReturnAgSpd_Min = '0'
ReturnAgSpd_GradMax = '35'

StaticFricTq_Max = '10'
StaticFricTq_Min = '0'
StaticFricTq_GradMax = '0.1'

RackPosn_Max = '74'
RackPosn_Min = '0'
CFactor_Max = '200'
CFactorAg_GradMax = '-20'
CFactorVehSpd_GradMax = '-8'

VLC_Max = '1'
VLC_Min = '0'
VLC_GradMax = '1.5'


class PlausibilityCheck:
    def __init__(self):
        self.Max = '0'
        self.Min = '0'
        self.GradMax = '0'

    def __PlauCheck__(self, target_mapname, xlsxwriter, workbook, worksheet, GRADENA, startpoint1_x, startpoint1_y, startpoint2_x, startpoint2_y, x_size1, y_size1, x_size2, y_size2):
        if target_mapname is 'EFC_EffortMap':
            self.Max = EffortTq_Max
            self.Min = EffortTq_Min
            self.GradMax = EffortTq_GradMax
        elif target_mapname is 'RFC_RackFMap':
            self.Max = RackFTq_Max
            self.Min = RackFTq_Min
            self.GradMax = RackFTq_Max
        elif target_mapname is 'DPC_DampgMap':
            self.Max = DampingTq_Max
            self.Min = DampingTq_Min
            self.GradMax = DampingTq_GradMax
        elif target_mapname is 'SFC_StatFricAgMap':
            self.Max = StaticFricTq_Max
            self.Min = StaticFricTq_Min
            self.GradMax = StaticFricTq_GradMax
        elif target_mapname is 'ARC_ActvRetMap':
            self.Max = ReturnAgSpd_Max
            self.Min = ReturnAgSpd_Min
            self.GradMax = ReturnAgSpd_GradMax
        elif target_mapname is 'KinematMap':
            self.Max = RackPosn_Max
            self.Min = RackPosn_Min
            self.GradMax = CFactor_Max
            self.Ag_GradMax = CFactorAg_GradMax
            self.VehSpd_GradMax = CFactorVehSpd_GradMax
        elif target_mapname is 'VLC_OCPosnMap':
            self.Max = VLC_Max
            self.Min = VLC_Min
            self.GradMax = VLC_GradMax
        else:
            return

        ##MinMax/Zero Check 조건부서식
        FormatForMinMaxCheck = workbook.add_format({'border': 5, 'border_color': '#FF00FF'})
        FormatForZeroCheck = FormatForMinMaxCheck

        before_data_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y + 1, startpoint1_x + 1)
        after_data_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint2_y + 1, startpoint2_x + 1)
        before_data_range = xlsxwriter.utility.xl_range_abs(startpoint1_y + 1, startpoint1_x + 1, startpoint1_y + y_size1, startpoint1_x + x_size1)
        after_data_range = xlsxwriter.utility.xl_range_abs(startpoint2_y + 1, startpoint2_x + 1, startpoint2_y + y_size2, startpoint2_x + x_size2)

        before_zero_range = xlsxwriter.utility.xl_range_abs(startpoint1_y + 1, startpoint1_x + 1, startpoint1_y + y_size1, startpoint1_x + 1)
        after_zero_range = xlsxwriter.utility.xl_range_abs(startpoint2_y + 1, startpoint2_x + 1, startpoint2_y + y_size2, startpoint2_x + 1)

        worksheet.conditional_format(before_data_range, {'type': 'formula',
                                                            'criteria': before_data_sp + ' > ' + self.Max,
                                                            'format': FormatForMinMaxCheck})
        worksheet.conditional_format(before_data_range, {'type': 'formula',
                                                       'criteria': before_data_sp + ' < ' + self.Min,
                                                       'format': FormatForMinMaxCheck})
        worksheet.conditional_format(after_data_range, {'type': 'formula',
                                                            'criteria': after_data_sp + ' > ' + self.Max,
                                                            'format': FormatForMinMaxCheck})
        worksheet.conditional_format(after_data_range, {'type': 'formula',
                                                            'criteria': after_data_sp + ' < ' + self.Min,
                                                            'format': FormatForMinMaxCheck})
        ##ZeroChekch 수행 할 맵 결정 (Friction은 Pass)
        if target_mapname is 'KinematMap' or target_mapname is 'EFC_EffortMap' or target_mapname is 'ARC_ActvRetMap' or target_mapname is 'RFC_RackFMap' or target_mapname is 'DPC_DampgMap' or target_mapname is 'VLC_OCPosnMap':
            worksheet.conditional_format(before_zero_range, {'type': 'formula',
                                                             'criteria': before_data_sp + ' > ' + '0',
                                                             'format': FormatForZeroCheck})
            worksheet.conditional_format(after_zero_range, {'type': 'formula',
                                                             'criteria': after_data_sp + ' > ' + '0',
                                                             'format': FormatForZeroCheck})

        



        FormatForGradCheck = workbook.add_format({'border': 5, 'border_color': '#FF0000'})
        FormatForGradAgCheck = workbook.add_format({'border': 5, 'border_color': '#FF9900'})
        FormatForGradSpdCheck = workbook.add_format({'border': 5, 'border_color': '#CC0000'})

        ##GRADIENT 표시 여부에 따라 GRADIENT CHECK 다르게
        if GRADENA is False:
            before_GRADdata_range = xlsxwriter.utility.xl_range_abs(startpoint1_y + 1, startpoint1_x + 2, startpoint1_y + y_size1, startpoint1_x + x_size1)
            after_GRADdata_range = xlsxwriter.utility.xl_range_abs(startpoint2_y + 1, startpoint2_x + 2, startpoint2_y + y_size2, startpoint2_x + x_size2)

            BeforeData1 = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y + 1, startpoint1_x + 1)
            BeforeData2 = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y + 1, startpoint1_x + 2)
            BeforeXAxis1 = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y, startpoint1_x + 1, row_abs = True)
            BeforeXAxis2 = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y, startpoint1_x + 2, row_abs = True)
            AfterData1 = xlsxwriter.utility.xl_rowcol_to_cell(startpoint2_y + 1, startpoint2_x + 1)
            AfterData2 = xlsxwriter.utility.xl_rowcol_to_cell(startpoint2_y + 1, startpoint2_x + 2)
            AfterXAxis1 = xlsxwriter.utility.xl_rowcol_to_cell(startpoint2_y, startpoint2_x + 1, row_abs = True)
            AfterXAxis2 = xlsxwriter.utility.xl_rowcol_to_cell(startpoint2_y, startpoint2_x + 2, row_abs = True)

            worksheet.conditional_format(before_GRADdata_range, {'type': 'formula',
                                                               'criteria': '('+ BeforeData2 + '-' + BeforeData1 + ')' + '/' + '(' + BeforeXAxis2 + '-' + BeforeXAxis1 + ')' + ' > ' + self.GradMax,
                                                               'format': FormatForGradCheck})
            worksheet.conditional_format(before_GRADdata_range, {'type': 'formula',
                                                                 'criteria': '(' + BeforeData2 + '-' + BeforeData1 + ')' + '/' + '(' + BeforeXAxis2 + '-' + BeforeXAxis1 + ')' + ' < ' + '0',
                                                                 'format': FormatForGradCheck})

            worksheet.conditional_format(after_GRADdata_range, {'type': 'formula',
                                                              'criteria': '('+ AfterData2 + '-' + AfterData1 + ')' + '/' + '(' + AfterXAxis2 + '-' + AfterXAxis1 + ')' + ' > ' + self.GradMax,
                                                              'format': FormatForGradCheck})
            worksheet.conditional_format(after_GRADdata_range, {'type': 'formula',
                                                                'criteria': '(' + AfterData2 + '-' + AfterData1 + ')' + '/' + '(' + AfterXAxis2 + '-' + AfterXAxis1 + ')' + ' < ' + '0',
                                                                'format': FormatForGradCheck})


        else:
            if target_mapname is 'KinematMap':
                before_GRADdata_range_forMinMax = xlsxwriter.utility.xl_range_abs(startpoint1_y + 1 + y_size1 + 2, startpoint1_x + 1, startpoint1_y + y_size1 + y_size1 + 2, startpoint1_x + x_size1)
                before_GRADdata_range_forAg = xlsxwriter.utility.xl_range_abs(startpoint1_y + 1 + y_size1 + 2, startpoint1_x + 2, startpoint1_y + y_size1 + y_size1 + 2, startpoint1_x + x_size1)
                before_GRADdata_range_forSpd = xlsxwriter.utility.xl_range_abs(startpoint1_y + 1 + y_size1 + 3, startpoint1_x + 1, startpoint1_y + y_size1 + y_size1 + 2, startpoint1_x + x_size1)
                before_GRADdata_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y + 1 + y_size1 + 2, startpoint1_x + 1)
                after_GRADdata_range_forMinMax = xlsxwriter.utility.xl_range_abs(startpoint2_y + 1 + y_size2 + 2, startpoint2_x + 1, startpoint2_y + y_size2 + y_size2 + 2, startpoint2_x + x_size2)
                after_GRADdata_range_forAg = xlsxwriter.utility.xl_range_abs(startpoint2_y + 1 + y_size2 + 2, startpoint2_x + 2, startpoint2_y + y_size2 + y_size2 + 2, startpoint2_x + x_size2)
                after_GRADdata_range_forSpd = xlsxwriter.utility.xl_range_abs(startpoint2_y + 1 + y_size2 + 3, startpoint2_x + 1, startpoint2_y + y_size2 + y_size2 + 2, startpoint2_x + x_size2)
                after_GRADdata_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint2_y + 1 + y_size2 + 2, startpoint2_x + 1)
                before_data_range_forMinMax = xlsxwriter.utility.xl_range_abs(startpoint1_y + 1, startpoint1_x + 1, startpoint1_y + y_size1, startpoint1_x + x_size1)
                after_data_range_forMinMax = xlsxwriter.utility.xl_range_abs(startpoint2_y + 1, startpoint2_x + 1, startpoint2_y + y_size2, startpoint2_x + x_size2)
                before_data_range_forAg = xlsxwriter.utility.xl_range_abs(startpoint1_y + 1, startpoint1_x + 2, startpoint1_y + y_size1, startpoint1_x + x_size1)
                before_data_range_forSpd = xlsxwriter.utility.xl_range_abs(startpoint1_y + 2, startpoint1_x + 1, startpoint1_y + y_size1, startpoint1_x + x_size1)
                after_data_range_forAg = xlsxwriter.utility.xl_range_abs(startpoint2_y + 1, startpoint2_x + 2, startpoint2_y + y_size2, startpoint2_x + x_size2)
                after_data_range_forSpd = xlsxwriter.utility.xl_range_abs(startpoint2_y + 2, startpoint2_x + 1, startpoint2_y + y_size2, startpoint2_x + x_size2)

                worksheet.conditional_format(before_GRADdata_range_forMinMax, {'type': 'formula',
                                                                               'criteria': before_GRADdata_sp + ' > ' + self.GradMax,
                                                                               'format': FormatForGradCheck})
                worksheet.conditional_format(before_GRADdata_range_forMinMax, {'type': 'formula',
                                                                               'criteria': before_GRADdata_sp + ' < ' + '0',
                                                                               'format': FormatForGradCheck})
                worksheet.conditional_format(after_GRADdata_range_forMinMax, {'type': 'formula',
                                                                              'criteria': after_GRADdata_sp + ' > ' + self.GradMax,
                                                                              'format': FormatForGradCheck})
                worksheet.conditional_format(after_GRADdata_range_forMinMax, {'type': 'formula',
                                                                              'criteria': after_GRADdata_sp + ' < ' + '0',
                                                                              'format': FormatForGradCheck})


                worksheet.conditional_format(before_data_range_forMinMax, {'type': 'formula',
                                                                               'criteria': before_GRADdata_sp + ' > ' + self.GradMax,
                                                                               'format': FormatForGradCheck})
                worksheet.conditional_format(before_data_range_forMinMax, {'type': 'formula',
                                                                               'criteria': before_GRADdata_sp + ' < ' + '0',
                                                                               'format': FormatForGradCheck})
                worksheet.conditional_format(after_data_range_forMinMax, {'type': 'formula',
                                                                              'criteria': after_GRADdata_sp + ' > ' + self.GradMax,
                                                                              'format': FormatForGradCheck})
                worksheet.conditional_format(after_data_range_forMinMax, {'type': 'formula',
                                                                              'criteria': after_GRADdata_sp + ' < ' + '0',
                                                                              'format': FormatForGradCheck})


                BeforeData1 = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y + 1 + y_size1 + 2, startpoint1_x + 1)
                BeforeData2forAg = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y + 1 + y_size1 + 2, startpoint1_x + 2)
                BeforeData2forSpd = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y + 1 + y_size1 + 3, startpoint1_x + 1)
                BeforeXAxis1 = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y + 1 + y_size1 + 1, startpoint1_x + 1, row_abs=True)
                BeforeXAxis2 = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y + 1 + y_size1 + 1, startpoint1_x + 2, row_abs=True)
                BeforeYAxis1 = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y + 1 + y_size1 + 2, startpoint1_x , col_abs=True)
                BeforeYAxis2 = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y + 1 + y_size1 + 3, startpoint1_x , col_abs=True)
                
                AfterData1 = xlsxwriter.utility.xl_rowcol_to_cell(startpoint2_y + 1 + y_size2 + 2, startpoint2_x + 1)
                AfterData2forAg = xlsxwriter.utility.xl_rowcol_to_cell(startpoint2_y + 1 + y_size2 + 2, startpoint2_x + 2)
                AfterData2forSpd = xlsxwriter.utility.xl_rowcol_to_cell(startpoint2_y + 1 + y_size2 + 3, startpoint2_x + 1)
                AfterXAxis1 = xlsxwriter.utility.xl_rowcol_to_cell(startpoint2_y + 1 + y_size2 + 1, startpoint2_x + 1, row_abs=True)
                AfterXAxis2 = xlsxwriter.utility.xl_rowcol_to_cell(startpoint2_y + 1 + y_size2 + 1, startpoint2_x + 2, row_abs=True)
                AfterYAxis1 = xlsxwriter.utility.xl_rowcol_to_cell(startpoint2_y + 1 + y_size2 + 2, startpoint2_x , col_abs=True)
                AfterYAxis2 = xlsxwriter.utility.xl_rowcol_to_cell(startpoint2_y + 1 + y_size2 + 3, startpoint2_x , col_abs=True)

                worksheet.conditional_format(before_GRADdata_range_forAg, {'type': 'formula',
                                                                     'criteria': '(' + BeforeData2forAg + '-' + BeforeData1 + ')' + '/' + '(' + BeforeXAxis2 + '-' + BeforeXAxis1 + ')' + ' < ' + self.Ag_GradMax,
                                                                     'format': FormatForGradAgCheck})
                worksheet.conditional_format(before_GRADdata_range_forAg, {'type': 'formula',
                                                                           'criteria': '(' + BeforeData2forAg + '-' + BeforeData1 + ')' + '/' + '(' + BeforeXAxis2 + '-' + BeforeXAxis1 + ')' + ' > ' + '0.001',
                                                                           'format': FormatForGradAgCheck})

                worksheet.conditional_format(before_GRADdata_range_forSpd, {'type': 'formula',
                                                                     'criteria': '(' + BeforeData2forSpd + '-' + BeforeData1 + ')' + '/' + '(' + BeforeYAxis2 + '-' + BeforeYAxis1 + ')' + ' < ' + self.VehSpd_GradMax,
                                                                     'format': FormatForGradSpdCheck})
                worksheet.conditional_format(before_GRADdata_range_forSpd, {'type': 'formula',
                                                                            'criteria': '(' + BeforeData2forSpd + '-' + BeforeData1 + ')' + '/' + '(' + BeforeYAxis2 + '-' + BeforeYAxis1 + ')' + ' > ' + '0.001',
                                                                            'format': FormatForGradSpdCheck})

                worksheet.conditional_format(after_GRADdata_range_forAg, {'type': 'formula',
                                                                     'criteria': '(' + AfterData2forAg + '-' + AfterData1 + ')' + '/' + '(' + AfterXAxis2 + '-' + AfterXAxis1 + ')' + ' < ' + self.Ag_GradMax,
                                                                     'format': FormatForGradAgCheck})
                worksheet.conditional_format(after_GRADdata_range_forAg, {'type': 'formula',
                                                                          'criteria': '(' + AfterData2forAg + '-' + AfterData1 + ')' + '/' + '(' + AfterXAxis2 + '-' + AfterXAxis1 + ')' + ' > ' + '0.001',
                                                                          'format': FormatForGradAgCheck})

                worksheet.conditional_format(after_GRADdata_range_forSpd, {'type': 'formula',
                                                                     'criteria': '(' + AfterData2forSpd + '-' + AfterData1 + ')' + '/' + '(' + AfterYAxis2 + '-' + AfterYAxis1 + ')' + ' < ' + self.VehSpd_GradMax,
                                                                     'format': FormatForGradSpdCheck})
                worksheet.conditional_format(after_GRADdata_range_forSpd, {'type': 'formula',
                                                                           'criteria': '(' + AfterData2forSpd + '-' + AfterData1 + ')' + '/' + '(' + AfterYAxis2 + '-' + AfterYAxis1 + ')' + ' > ' + '0.001',
                                                                           'format': FormatForGradSpdCheck})



                worksheet.conditional_format(before_data_range_forAg, {'type': 'formula',
                                                                     'criteria': '(' + BeforeData2forAg + '-' + BeforeData1 + ')' + '/' + '(' + BeforeXAxis2 + '-' + BeforeXAxis1 + ')' + ' < ' + self.Ag_GradMax,
                                                                     'format': FormatForGradAgCheck})
                worksheet.conditional_format(before_data_range_forAg, {'type': 'formula',
                                                                           'criteria': '(' + BeforeData2forAg + '-' + BeforeData1 + ')' + '/' + '(' + BeforeXAxis2 + '-' + BeforeXAxis1 + ')' + ' > ' + '0.001',
                                                                           'format': FormatForGradAgCheck})

                worksheet.conditional_format(before_data_range_forSpd, {'type': 'formula',
                                                                     'criteria': '(' + BeforeData2forSpd + '-' + BeforeData1 + ')' + '/' + '(' + BeforeYAxis2 + '-' + BeforeYAxis1 + ')' + ' < ' + self.VehSpd_GradMax,
                                                                     'format': FormatForGradSpdCheck})
                worksheet.conditional_format(before_data_range_forSpd, {'type': 'formula',
                                                                            'criteria': '(' + BeforeData2forSpd + '-' + BeforeData1 + ')' + '/' + '(' + BeforeYAxis2 + '-' + BeforeYAxis1 + ')' + ' > ' + '0.001',
                                                                            'format': FormatForGradSpdCheck})

                worksheet.conditional_format(after_data_range_forAg, {'type': 'formula',
                                                                     'criteria': '(' + AfterData2forAg + '-' + AfterData1 + ')' + '/' + '(' + AfterXAxis2 + '-' + AfterXAxis1 + ')' + ' < ' + self.Ag_GradMax,
                                                                     'format': FormatForGradAgCheck})
                worksheet.conditional_format(after_data_range_forAg, {'type': 'formula',
                                                                          'criteria': '(' + AfterData2forAg + '-' + AfterData1 + ')' + '/' + '(' + AfterXAxis2 + '-' + AfterXAxis1 + ')' + ' > ' + '0.001',
                                                                          'format': FormatForGradAgCheck})

                worksheet.conditional_format(after_data_range_forSpd, {'type': 'formula',
                                                                     'criteria': '(' + AfterData2forSpd + '-' + AfterData1 + ')' + '/' + '(' + AfterYAxis2 + '-' + AfterYAxis1 + ')' + ' < ' + self.VehSpd_GradMax,
                                                                     'format': FormatForGradSpdCheck})
                worksheet.conditional_format(after_data_range_forSpd, {'type': 'formula',
                                                                           'criteria': '(' + AfterData2forSpd + '-' + AfterData1 + ')' + '/' + '(' + AfterYAxis2 + '-' + AfterYAxis1 + ')' + ' > ' + '0.001',
                                                                           'format': FormatForGradSpdCheck})



            else:
                before_GRADdata_range = xlsxwriter.utility.xl_range_abs(startpoint1_y + 1 + y_size1 + 2, startpoint1_x + 1, startpoint1_y + y_size1 + y_size1 + 2, startpoint1_x + x_size1)
                before_GRADdata_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint1_y + 1 + y_size1 + 2, startpoint1_x + 1)
                after_GRADdata_range = xlsxwriter.utility.xl_range_abs(startpoint2_y + 1 + y_size1 + 2, startpoint2_x + 1, startpoint2_y + y_size1 + y_size2 + 2, startpoint2_x + x_size2)
                after_GRADdata_sp = xlsxwriter.utility.xl_rowcol_to_cell(startpoint2_y + 1 + y_size1 + 2, startpoint2_x + 1)


                worksheet.conditional_format(before_GRADdata_range, {'type': 'formula',
                                                                  'criteria': before_GRADdata_sp + ' > ' + self.GradMax,
                                                                  'format': FormatForGradCheck})
                worksheet.conditional_format(before_GRADdata_range, {'type': 'formula',
                                                                     'criteria': before_GRADdata_sp + ' < ' + '0',
                                                                     'format': FormatForGradCheck})
                worksheet.conditional_format(before_data_range, {'type': 'formula',
                                                                     'criteria': before_GRADdata_sp + ' > ' + self.GradMax,
                                                                     'format': FormatForGradCheck})
                worksheet.conditional_format(before_data_range, {'type': 'formula',
                                                                 'criteria': before_GRADdata_sp + ' < ' + '0',
                                                                 'format': FormatForGradCheck})

                worksheet.conditional_format(after_GRADdata_range, {'type': 'formula',
                                                                    'criteria': after_GRADdata_sp + ' > ' + self.GradMax,
                                                                    'format': FormatForGradCheck})
                worksheet.conditional_format(after_GRADdata_range, {'type': 'formula',
                                                                    'criteria': after_GRADdata_sp + ' < ' + '0',
                                                                    'format': FormatForGradCheck})
                worksheet.conditional_format(after_data_range, {'type': 'formula',
                                                                    'criteria': after_GRADdata_sp + ' > ' + self.GradMax,
                                                                    'format': FormatForGradCheck})
                worksheet.conditional_format(after_data_range, {'type': 'formula',
                                                                'criteria': after_GRADdata_sp + ' < ' + '0',
                                                                'format': FormatForGradCheck})



# xml import용도의 코드
# 파싱한 xml 내용을 불러온다.

import re
import xml.etree.ElementTree as ET
import numpy as np

class PrmXmlImpoter:
    category_names = ["MAP", "CURVE", "CURVE_AXIS", "COM_AXIS", "VAL_BLK", "VALUE", "ASCII"]

    def __init__(self):
        self.__filename__ = ""
        self.__etc_prms__ = {}
        self.__local_prms__ = {}
        self.__global_prms__ = {}

    def __load_xml__(self, filename):
        self.__filename__ = filename
        tree = ET.parse(self.__filename__)
        root = tree.getroot()

        self.__etc_prms__.clear()
        self.__load_prm_data(self.__etc_prms__, root.find("ETC"))

        self.__local_prms__.clear()
        self.__load_prm_data(self.__local_prms__, root.find("Local"))

        self.__global_prms__.clear()
        self.__load_prm_data(self.__global_prms__, root.find("Global"))

    @staticmethod
    def __load_prm_data(prm_lst, root):
        prm_lst.clear()

        for swc_elem in root.findall("*"):
            swc_name = swc_elem.tag

            prm_lst[swc_name] = {}
            for category_name in PrmXmlImpoter.category_names:
                prm_lst[swc_name][category_name] = {}
                cat_elem = swc_elem.find(category_name)
                if cat_elem is not None:
                    for param_elem in cat_elem.findall("*"):
                        param_name = param_elem.tag
                        data = None
                        if category_name == "MAP":
                            data = PrmMap(param_name, param_elem)
                        elif category_name == "CURVE" or category_name == "CURVE_AXIS":
                            category_name = "CURVE"
                            data = PrmCurve(param_name, param_elem)
                        elif category_name == "COM_AXIS":
                            data = PrmComAxis(param_name, param_elem)
                        elif category_name == "VAL_BLK":
                            data = PrmValBlk(param_name, param_elem)
                        elif category_name == "VALUE":
                            data = PrmValue(param_name, param_elem)
                        elif category_name == "ASCII":
                            data = PrmASCII(param_name, param_elem)

                        if data is not None:
                            data.__initialize_from_xml__()

                        prm_lst[swc_name][category_name][param_name] = data

import re
import xml.etree.ElementTree as ET
import numpy as np
import matplotlib.pyplot as plt


class PrmBase:
    def __init__(self, name, root_elem):
        self.__name__ = name
        self.root_elem = root_elem
        self.pv_root_elem = None
        self.axis_root_elem = None
        self.vc_elem = None
        self.category = "BASE"
        self.value_str = ""
        self.value = None
        self.isnumeric = True

    def __initialize_from_cdfx__(self):
        temp_elem = self.root_elem.find("SW-INSTANCE-PROPS-VARIANTS")
        if temp_elem is None:
            self.pv_root_elem=self.root_elem
        else:
            self.pv_root_elem=temp_elem.find('SW-INSTANCE-PROPS-VARIANT')
        self.axis_root_elem = self.pv_root_elem.find('SW-AXIS-CONTS')
        self.vc_elem = self.pv_root_elem.find('SW-VALUE-CONT')

    def __initialize_from_xml__(self):
        self.value_str = self.root_elem.find("VALUE").text


class PrmMap(PrmBase):
    def __init__(self, name, root_elem):
        super().__init__(name, root_elem)
        self.__xsize__ = 0
        self.__ysize__ = 0
        self.__xaxis__ = ""
        self.__yaxis__ = ""
        self.category = "MAP"

    def __initialize_from_cdfx__(self):
        super().__initialize_from_cdfx__()

        size_elms = []
        value_elms = []
        axis_elms = []
        for child in self.vc_elem.find('SW-ARRAYSIZE'):
            size_elms.append(child)

        for child in self.vc_elem.find('SW-VALUES-PHYS'):
            value_elms.append(child.text)

        for child in self.axis_root_elem:
            axis_elms.append(child)

        self.__xsize__ = int(size_elms[0].text)
        self.__ysize__ = int(size_elms[1].text)

        self.__xaxis__ = re.split("\.", axis_elms[0].find('SW-INSTANCE-REF').text)[1]
        self.__yaxis__ = re.split("\.", axis_elms[1].find('SW-INSTANCE-REF').text)[1]

        y = 0

        self.value_str = "["

        while y < self.__ysize__:
            line_data = "["
            x = 0
            while x < self.__xsize__:
                value = value_elms[x + y * self.__xsize__]
                line_data = line_data + value
                x = x + 1
                if x < self.__xsize__:
                    line_data = line_data + ","
            y = y + 1
            self.value_str = self.value_str + line_data + "]"
            if y < self.__ysize__:
                self.value_str = self.value_str + ";"

        self.value_str = self.value_str + "]"

    def __initialize_from_xml__(self):
        super().__initialize_from_xml__()
        self.__xsize__ = int(self.root_elem.get("X_SIZE"))
        self.__ysize__ = int(self.root_elem.get("Y_SIZE"))
        self.__xaxis__ = self.root_elem.get("X_AXIS")
        self.__yaxis__ = self.root_elem.get("Y_AXIS")
        self.value = np.zeros((self.__ysize__, self.__xsize__))
        row_data_list = re.split(";", self.value_str.replace("[", "").replace("]", ""))
        for row in range(len(row_data_list)):
            val_data_list = re.split(",", row_data_list[row])
            for col in range(len(val_data_list)):
                self.value[row][col] = float(val_data_list[col])


class PrmCurve(PrmBase):
    def __init__(self, name, root_elem):
        super().__init__(name, root_elem)
        self.category = "MAP"

    def __initialize_from_cdfx__(self):
        super().__initialize_from_cdfx__()
        size_elms = []
        value_elms = []
        axis_elms = []
        for child in self.vc_elem.find('SW-ARRAYSIZE'):
            size_elms.append(child)

        for child in self.vc_elem.find('SW-VALUES-PHYS'):
            value_elms.append(child.text)

        for child in self.axis_root_elem:
            axis_elms.append(child)

        self.__xsize__ = int(size_elms[0].text)

        self.__xaxis__ = re.split("\.", axis_elms[0].find('SW-INSTANCE-REF').text)[1]

        self.value_str = "["
        x = 0
        while x < self.__xsize__:
            value = value_elms[x]
            self.value_str = self.value_str + value
            x = x + 1
            if x < self.__xsize__:
                self.value_str = self.value_str + ","
        self.value_str = self.value_str + "]"

    def __initialize_from_xml__(self):
        super().__initialize_from_xml__()
        self.__xsize__ = int(self.root_elem.get("X_SIZE"))
        self.__xaxis__ = self.root_elem.get("X_AXIS")
        self.value = np.zeros(self.__xsize__)
        data_list = re.split(",", self.value_str.replace("[", "").replace("]", ""))
        for col in range(len(data_list)):
            self.value[col] = float(data_list[col])

class PrmASCII(PrmBase):
    def __init__(self, name, root_elem):
        super().__init__(name, root_elem)
        self.category = "ASCII"

    def __initialize_from_cdfx__(self):
        super().__initialize_from_cdfx__()
        version_elms = []
        for child in self.vc_elem.find('SW-VALUES-PHYS'):
            version_elms.append(child.text)
        self.value_str = version_elms[0]


    def __initialize_from_xml__(self):
        super().__initialize_from_xml__()
        self.value = np.zeros(1)




class PrmComAxis(PrmBase):
    def __init__(self, name, root_elem):
        super().__init__(name, root_elem)

    def __initialize_from_cdfx__(self):
        super().__initialize_from_cdfx__()
        value_elms = []
        for child in self.vc_elem.find('SW-VALUES-PHYS'):
            value_elms.append(child.text)
        self.__xsize__ = len(value_elms)
        self.value_str = "["
        x = 0
        while x < self.__xsize__:
            value = value_elms[x]
            self.value_str = self.value_str + value
            x = x + 1
            if x < self.__xsize__:
                self.value_str = self.value_str + ","
        self.value_str = self.value_str + "]"

    def __initialize_from_xml__(self):
        super().__initialize_from_xml__()
        self.__xsize__ = int(self.root_elem.get("X_SIZE"))
        self.value = np.zeros(self.__xsize__)
        data_list = re.split(",", self.value_str.replace("[", "").replace("]", ""))
        for col in range(len(data_list)):
            self.value[col] = float(data_list[col])


class PrmValBlk(PrmBase):
    def __init__(self, name, root_elem):
        super().__init__(name, root_elem)

    def __initialize_from_cdfx__(self):
        super().__initialize_from_cdfx__()

        size_elms = []
        value_elms = []
        for child in self.vc_elem.find('SW-ARRAYSIZE'):
            size_elms.append(child)

        for child in self.vc_elem.find('SW-VALUES-PHYS'):
            value_elms.append(child.text)

        self.__xsize__ = int(size_elms[0].text)
        self.__ysize__ = int(size_elms[1].text)
        y = 0
        self.value_str = "["

        while y < self.__ysize__:
            line_data = "["
            x = 0
            while x < self.__xsize__:
                value = value_elms[x + y * self.__xsize__]
                line_data = line_data + value
                x = x + 1
                if x < self.__xsize__:
                    line_data = line_data + ","
            y = y + 1
            self.value_str = self.value_str + line_data + "]"
            if y < self.__ysize__:
                self.value_str = self.value_str + ";"
        self.value_str = self.value_str + "]"

    def __initialize_from_xml__(self):
        super().__initialize_from_xml__()
        self.__xsize__ = int(self.root_elem.get("X_SIZE"))
        self.__ysize__ = int(self.root_elem.get("Y_SIZE"))
        self.value = np.zeros((self.__ysize__, self.__xsize__))
        row_data_list = re.split(";", self.value_str.replace("[", "").replace("]", ""))
        for row in range(len(row_data_list)):
            val_data_list = re.split(",", row_data_list[row])
            for col in range(len(val_data_list)):
                self.value[row][col] = float(val_data_list[col])


class PrmValue(PrmBase):
    def __init__(self, name, root_elem):
        super().__init__(name, root_elem)

    def __initialize_from_cdfx__(self):
        super().__initialize_from_cdfx__()
        value_elms = []

        for child in self.vc_elem.find('SW-VALUES-PHYS'):
            value_elms.append(child.text)
        self.value_str = value_elms[0]

    def __initialize_from_xml__(self):

        super().__initialize_from_xml__()
        self.value = np.zeros(1)

        try:
            self.value[0] = float(self.value_str)
        except ValueError:
            self.value = self.value_str
            self.isnumeric = False


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
        # swc_name = SWC명
        # swc_elem = Local, ETC, Global
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


class CdfxExporter:

    def __init__(self):
        self.__cdfx_filename = ""
        self.__xml_filename = ""
        self.etc_prms = {}
        self.local_prms = {}
        self.global_prms = {}

    def __load_from_cdfx__(self, cdfx_filename):
        self.__cdfx_filename = cdfx_filename
        tree = ET.parse(self.__cdfx_filename)
        self.etc_prms.clear()
        self.local_prms.clear()
        self.global_prms.clear()

        root = tree.getroot()
        root = root.find("SW-SYSTEMS")
        root = root.find("SW-SYSTEM")
        root = root.find("SW-INSTANCE-SPEC")
        root = root.find("SW-INSTANCE-TREE")

        for param_elem in root.findall("SW-INSTANCE"):
            name_elem = param_elem.find("SHORT-NAME")
            cat_elem = param_elem.find("CATEGORY")
            vldflg = name_elem is not None
            vldflg = vldflg and cat_elem is not None
            if not vldflg:
                continue

            param_fullname = name_elem.text
            category = cat_elem.text
            param_datas = re.split("\.", param_fullname)
            if len(param_datas) < 2:
                selected_data_dic = self.etc_prms
                swc_name = param_datas[0]
                param_name = param_datas[0]
            elif len(param_datas) < 3:
                selected_data_dic = self.local_prms
                swc_name = param_datas[0]
                param_name = param_datas[1]
            elif len(param_datas) >= 3:
                selected_data_dic = self.global_prms
                swc_name = param_datas[0]
                param_name = param_datas[1]


            if swc_name not in selected_data_dic:
                selected_data_dic[swc_name] = {}

            param_data = None
            if category == "MAP":
                param_data = PrmMap(param_name, param_elem)
            elif category == "CURVE"  or category == "CURVE_AXIS":
                category = "CURVE"
                param_data = PrmCurve(param_name, param_elem)
            elif category == "COM_AXIS":
                param_data = PrmComAxis(param_name, param_elem)
            elif category == "VAL_BLK":
                param_data = PrmValBlk(param_name, param_elem)
            elif category == "VALUE":
                param_data = PrmValue(param_name, param_elem)
            elif category == "ASCII":
                param_data = PrmASCII(param_name, param_elem)

            if category not in selected_data_dic[swc_name]:
                selected_data_dic[swc_name][category] = {}

            if param_name not in selected_data_dic[swc_name][category]:
                param_data.__initialize_from_cdfx__()
                selected_data_dic[swc_name][category][param_name] = param_data

    def __export_to_xml__(self, xml_filename):

        self.__xml_filename = xml_filename
        dest_root = ET.Element("Root")

        etc_keys = self.etc_prms.keys()
        local_keys = self.local_prms.keys()
        global_keys = self.global_prms.keys()

        etc_root = ET.Element("ETC")
        etc_root.set("Count", str(len(etc_keys)))
        CdfxExporter.__export_xml_data(self.etc_prms, etc_root)

        local_root = ET.Element("Local")
        local_root.set("Count", str(len(local_keys)))
        CdfxExporter.__export_xml_data(self.local_prms, local_root)

        global_root = ET.Element("Global")
        global_root.set("Count", str(len(global_keys)))
        CdfxExporter.__export_xml_data(self.global_prms, global_root)

        dest_root.append(etc_root)
        dest_root.append(global_root)
        dest_root.append(local_root)
        self.__indentcheck(dest_root)
        ET.ElementTree(dest_root).write(self.__xml_filename, encoding="utf-8", xml_declaration=True)

    @staticmethod
    def __indentcheck(elem, level=0):
        i = "\n" + level * "  "
        if len(elem):
            if not elem.text or not elem.text.strip():
                elem.text = i + "  "
            if not elem.tail or not elem.tail.strip():
                elem.tail = i
            for elem in elem:
                CdfxExporter.__indentcheck(elem, level + 1)
            if not elem.tail or not elem.tail.strip():
                elem.tail = i
        else:
            if level and (not elem.tail or not elem.tail.strip()):
                elem.tail = i

    def __export_xml_data(prm_lst, root):
        for swc_name in prm_lst.keys():
            swc_elem = ET.Element(swc_name)
            for category in prm_lst[swc_name].keys():
                c_elem = ET.Element(category)
                swc_elem.append(c_elem)
                for param_name in prm_lst[swc_name][category].keys():
                    param_elem = ET.Element(param_name)
                    c_elem.append(param_elem)
                    data = prm_lst[swc_name][category][param_name]
                    value_elem = ET.Element('VALUE')
                    if category == "MAP":
                        param_elem.set('X_SIZE', str(data.__xsize__))
                        param_elem.set('Y_SIZE', str(data.__ysize__))
                        param_elem.set('X_AXIS', str(data.__xaxis__))
                        param_elem.set('Y_AXIS', str(data.__yaxis__))
                    elif category == "CURVE":
                        param_elem.set('X_SIZE', str(data.__xsize__))
                        param_elem.set('X_AXIS', str(data.__xaxis__))
                    elif category == "COM_AXIS" or category == "CURVE_AXIS":
                        param_elem.set('X_SIZE', str(data.__xsize__))
                    elif category == "VAL_BLK":
                        param_elem.set('X_SIZE', str(data.__xsize__))
                        param_elem.set('Y_SIZE', str(data.__ysize__))

                    value_elem.text = data.value_str
                    param_elem.append(value_elem)
            root.append(swc_elem)

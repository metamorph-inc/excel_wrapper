from openmdao.api import Component
import os
import win32com.client
from xml.etree import ElementTree as ET
import json


class ExcelWrapper(Component):
    """ An Excel Wrapper """

    def __init__(self, excelFile, varFile):
        super(ExcelWrapper, self).__init__()
        self.Var_dict = None
        self.json_True = False
        if varFile.endswith('.json'):
            self.json_true = True
        elif varFile.endswith('.xml'):
            self.json_True = False

        if not self.json_true:
            self.xmlFile = varFile
            self.create_xml_dict()
        else:
            self.jsonFile = varFile
            self.create_json_dict()

        for key, value in self.Var_dict.items():
            if key == "params":
                for z in value:
                    self.add_param(**z)
            elif key == "unknowns":
                for z in value:
                    print z["name"]
                    self.add_output(**z)

        self.excelFile = excelFile
        self.xlInstance = None
        self.workbook = None
        self.xl_sheet = None
        self.ExcelConnectionIsValid = True
        if not os.path.exists(self.excelFile):
            print "Invalid file given"
            self.ExcelConnectionIsValid = False

        else:
            self.excelFile = os.path.abspath(self.excelFile)
            xl = self.openExcel()
            if xl is None:
                print "Connection to Excel failed."
                self.ExcelConnectionIsValid = False

            else:
                self.xlInstance = xl
                self.workbook = xl.Workbooks.Open(self.excelFile)
                self.workbook = xl.ActiveWorkbook

    # End __init__

    def __del__(self):
        if self.workbook is not None:
            self.workbook.Close(SaveChanges=False)

        if self.xlInstance is not None:
            del(self.xlInstance)
            self.xlInstance = None
    # End __del__

    def create_xml_dict(self):
        try:
            tree = ET.parse(self.xmlFile)
        except:
            if not os.path.exists(self.xmlFile):
                print 'Cannot find the xml file at ' + self.xmlFile

        self.Var_dict = {
            "unknowns": [],
            "params": []
        }
        variables = tree.findall("Variable")
        for v in variables:
            name = v.attrib['name']
            kwargs = dict([(key, v.attrib[key]) for key in ('name', 'val', 'iotype', 'desc', 'units', 'row', 'column', 'sheet', 'type') if key in v.attrib])
            if v.attrib['iotype'] == 'in':
                self.Var_dict["params"].append(kwargs)
            elif v.attrib['iotype'] == 'out':
                self.Var_dict["unknowns"].append(kwargs)

    def create_json_dict(self):
        try:
            with open(self.jsonFile) as jh:
                self.Var_dict = json.load(jh)
        except:
            if not os.path.exists(self.jsonFile):
                print 'Cannot find the json file at ' + self.jsonFileFile

    def openExcel(self):
        try:
            xl = win32com.client.Dispatch("Excel.Application")

        except:
            return None

        return xl

    def letter2num(self, letters, zbase=False):
        letters = str(letters)
        letters_up = str(letters.upper())
        res = 0
        weight = len(letters_up) - 1
        for i, c in enumerate(letters_up):
            res += (ord(c) - 64) * 26**(weight - i)
        if not zbase:
            return res
        return res - 1

    def solve_nonlinear(self, params, unknowns, resids):

        if not self.ExcelConnectionIsValid or \
            self.xlInstance is None or \
                self.workbook is None:
            print "Aborted Execution of Bad ExcelWrapper Component Instance"
            return

        wb = self.workbook
        namelist = [x.name for x in wb.Names]

        data_x = self.Var_dict

        for key, value in data_x.items():
            if key == "params":
                for z in value:
                    name = z["name"]
                    if 'row' and 'column' in z:
                        if 'sheet' in z:
                            xl_sheet = self.xlInstance.Sheets(z['sheet'])
                        else:
                            xl_sheet = self.xlInstance.Sheets(1)

                        xl_sheet.Select()
                        xl_sheet.Cells(z["row"], self.letter2num(z["column"])).value = params[name]
                    else:
                        self.xlInstance.Range(wb.Names(name).RefersToLocal).Value = params[name]

        for key, value in data_x.items():
            if key == "unknowns":
                for z in value:

                    name = z["name"]
                    if "row" and "column" in z:
                        if "sheet" in z:
                            xl_sheet = self.xlInstance.Sheets(z['sheet'])
                        else:
                            xl_sheet = self.xlInstance.Sheets(1)

                        xl_sheet.Select()
                        excel_value = xl_sheet.Cells(z["row"], self.letter2num(z["column"])).value
                    else:
                        excel_value = self.xlInstance.Range(wb.Names(name).RefersToLocal).Value
                    # print excel_value
                    if z["type"] == 'Float':
                        unknowns[name] = float(excel_value)
                    elif z["type"] == 'Int':
                        unknowns[name] = int(excel_value)

                    elif z["type"] == 'Bool':
                        unknowns[name] = excel_value
                    elif z["type"] == 'Str':
                        unknowns[name] = str(excel_value)

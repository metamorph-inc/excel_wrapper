from openmdao.api import Component
import os
import win32com.client
import ast
from xml.etree import ElementTree as ET

class ExcelWrapper(Component):
    """ An Excel Wrapper """

    def __init__(self, excelFile, xmlFile):
        super(ExcelWrapper, self).__init__()

        self.xmlFile = xmlFile
        try:
            tree = ET.parse(self.xmlFile)
        except:
            if not os.path.exists(self.xmlFile):
                print 'Cannot find the xml file at ' + self.xmlFile

        self.variables = tree.findall("Variable")
        for v in self.variables:
            name = v.attrib['name']
            kwargs = dict([(key, v.attrib[key]) for key in ('iotype', 'desc', 'units') if key in v.attrib])
            print v.attrib['name']
            if v.attrib['iotype'] == 'in':

                if v.attrib['type'] == 'Float':
                    print v.attrib['name']
                    self.add_param(v.attrib['name'],float(v.attrib['value']), **kwargs)
                elif v.attrib['type'] == 'Int':
                    self.add_param(v.attrib['name'], int(v.attrib['value']), **kwargs)
                elif v.attrib['type'] == 'Bool':
                    self.add_param(v.attrib['name'], ast.literal_eval(v.attrib['value']), **kwargs)
                elif v.attrib['type'] == 'Str':
                    self.add_param(v.attrib['name'], v.attrib['value'], **kwargs)

            else:
                if v.attrib['type'] == 'Float':
                    self.add_output(v.attrib['name'], 1.0)
                elif v.attrib['type'] == 'Int':
                    self.add_output(v.attrib['name'],  1)
                elif v.attrib['type'] == 'Bool':
                    self.add_output(v.attrib['name'], True)
                elif v.attrib['type'] == 'Str':
                    self.add_output(v.attrib['name'], "abc")

        self.excelFile = excelFile
        self.xlInstance = None
        self.workbook = None
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
    # End __init__

    def __del__(self):
        if self.workbook is not None:
            self.workbook.Close(SaveChanges=False)

        if self.xlInstance is not None:
            del(self.xlInstance)
            self.xlInstance = None
    # End __del__

    def openExcel(self):
        try:
            xl = win32com.client.Dispatch("Excel.Application")

        except:
            return None

        return xl
    # End openExcel

    def solve_nonlinear(self, params, unknowns, resids):

        if not self.ExcelConnectionIsValid or \
            self.xlInstance is None or \
                self.workbook is None:
            print "Aborted Execution of Bad ExcelWrapper Component Instance"
            return

        wb = self.workbook
        namelist = [x.name for x in wb.Names]

        for v in self.variables:
            name = v.attrib['name']

            if v.attrib['iotype'] == 'in':
                    self.xlInstance.Range(wb.Names(name).RefersToLocal).Value = params[name]
            else:
                try:
                    excel_value = self.xlInstance.Range(wb.Names(name).RefersToLocal).Value
                except:
                    print 'Cannot retrieve values from the Excel file'
                    if name not in namelist:
                        print 'Error: ' + name + ' is not defined in ' + self.excelFile

                if v.attrib['type'] == 'Float':
                    unknowns[name] = float(excel_value)
                elif v.attrib['type'] == 'Int':
                    unknowns[name] = int(excel_value)
                    print int(excel_value)
                elif v.attrib['type'] == 'Bool':
                   unknowns[name] = excel_value
                elif v.attrib['type'] == 'Str':
                    unknowns[name] = excel_value

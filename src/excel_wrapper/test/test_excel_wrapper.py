import unittest
import nose
import os
from openmdao.api import Group, Problem
from excel_wrapper import ExcelWrapper


class ExcelWrapperTestCase(unittest.TestCase):

    def setUp(self):
        if os.name != 'nt':
            raise nose.SkipTest('Currently, excel_wrapper works only on Windows.')
        if os.name == 'posix':
            raise nose.SkipTest('Currently, excel_wrapper works only on Windows.')

    def tearDown(self):
        pass

    def test_ExcelWrapper(self):
        prob = Problem()
        root = prob.root = Group()
        excelFile = r"excel_wrapper_test.xlsx"
        jsonFile = r"testjson_1.json"
        root.add('ew', ExcelWrapper(excelFile, jsonFile, True), promotes=['*'])
        prob.setup()
        prob.run()

        self.assertEqual((2.1 * float(prob['x'])), prob['y'], "Excel Wrapper failed for FLoat values")
        self.assertEqual((bool(prob['b'])), prob['bout'], "Excel Wrapper failed for Boolean Values")
        self.assertEqual(prob['s'].lower(), prob['sout'], "Excel Wrapper failed for String values")
        self.assertEqual(float(prob['sheet1_in']) + 100, prob['sheet2_out'], "Excel wrapper fails in multiple sheets")

if __name__ == "__main__":
    unittest.main()

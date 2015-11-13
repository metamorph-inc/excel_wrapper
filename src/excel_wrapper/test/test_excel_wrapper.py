import unittest
import glob
import nose
import sys
import logging
import os
from openmdao.api import IndepVarComp, Group, Problem


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
        xmlFile = r"excel_wrapper_test.xml"
        root.add('ew', ExcelWrapper(excelFile, xmlFile),promotes=['*'])
        prob.setup()
        prob.run()

        self.assertEqual((2.1* prob['x']),prob['y'],"Excel Wrapper failed for FLoat values")
        self.assertEqual((not prob['b']),prob['bout'],"Excel Wrapper failed for Boolean Values")
        self.assertEqual(prob['s'].lower(),prob['sout'],"Excel Wrapper failed for String values")
        self.assertEqual(prob['sheet1_in']+100,prob['sheet2_out'],"Excel wrapper fails in multiple sheets")





        
if __name__ == "__main__":
    unittest.main()

import unittest
import nose
import os
import os.path
from openmdao.api import Group, Problem, IndepVarComp
from excel_wrapper.excel_wrapper import ExcelWrapper


class ExcelWrapperTestCase(unittest.TestCase):

    def setUp(self):
        if os.name != 'nt':
            raise nose.SkipTest('Currently, excel_wrapper works only on Windows.')

    def tearDown(self):
        pass

    def _test_ExcelWrapper(self, varFile, inputs={'x': 10, 'b': True, 's': u'aSdF'}):
        prob = Problem()
        root = prob.root = Group()
        this_dir = os.path.dirname(os.path.abspath(__file__))
        excelFile = os.path.join(this_dir, "excel_wrapper_test.xlsx")
        jsonFile = os.path.join(this_dir, varFile)
        root.add('ew', ExcelWrapper(excelFile, jsonFile), promotes=['*'])
        varComp = IndepVarComp(((name, val) for name, val in inputs.iteritems()))
        root.add('vc', varComp)
        root.connect('vc.x', 'x')
        root.connect('vc.b', 'b')
        root.connect('vc.s', 's')
        prob.setup()
        prob.run()

        self.assertEqual((2.1 * float(prob['x'])), prob['y'], "Excel Wrapper failed for Float values")
        self.assertEqual((2.1 * float(inputs['x'])), prob['y'], "Excel Wrapper failed for Float values")
        self.assertEqual(inputs['b'], prob['b'])
        self.assertEqual(bool(prob['b']), not prob['bout'])
        self.assertEqual(inputs['s'], prob['s'])
        self.assertEqual(prob['s'].lower(), prob['sout'], "Excel Wrapper failed for String values")
        self.assertEqual(float(prob['sheet1_in']) + 100, prob['sheet2_out'], "Excel wrapper fails in multiple sheets")

    def test_ExcelWrapperJson(self):
        return self._test_ExcelWrapper("testjson_1.json")

    def test_ExcelWrapperJson2(self):
        return self._test_ExcelWrapper("testjson_1.json", inputs={'x': -10, 'b': False, 's': u'TEST'})

    def test_ExcelWrapperXml(self):
        return self._test_ExcelWrapper("excel_wrapper_test.xml")


if __name__ == "__main__":
    unittest.main()
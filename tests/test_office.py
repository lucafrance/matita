import unittest

from vipera.office import excel, powerpoint, word

class TestApplicationOpenClose(unittest.TestCase):
    
    def test_excel(self):
        xl_app = excel.Application().new()
        self.assertIs(type(xl_app), excel.Application)
        xl_app.Visible = True
        self.assertTrue(xl_app.Visible)
        xl_app.Quit()

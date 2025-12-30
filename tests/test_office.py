import unittest

from vipera.office import access, excel, powerpoint, word

class TestApplicationOpenClose(unittest.TestCase):
    
    def test_access(self):
        acc_app = access.Application().new()
        self.assertIs(type(acc_app), access.Application)
        acc_app.Visible = True
        self.assertTrue(acc_app.Visible)
        acc_app.Quit()
    
    def test_excel(self):
        xl_app = excel.Application().new()
        self.assertIs(type(xl_app), excel.Application)
        xl_app.Visible = True
        self.assertTrue(xl_app.Visible)
        xl_app.Quit()

    def test_powerpoint(self):
        pp_app = powerpoint.Application().new()
        self.assertIs(type(pp_app), powerpoint.Application)
        pp_app.Visible = True
        self.assertTrue(pp_app.Visible)
        pp_app.Quit()

    def test_word(self):
        wd_app = word.Application().new()
        self.assertIs(type(wd_app), word.Application)
        wd_app.Visible = True
        self.assertTrue(wd_app.Visible)
        wd_app.Quit()
        
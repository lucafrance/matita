import unittest

import win32com.client

from matita.office import access, excel, outlook, powerpoint, word

class TestExcel(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.xl_app = excel.Application().new()
        cls.xl_app.Visible = True

    @classmethod
    def tearDownClass(cls):
        cls.xl_app.Quit()

    def test_excel_types(self):
        wkb = self.xl_app.Workbooks.Add()
        wks = wkb.Worksheets(1)

        cell_str = wks.Range("A1")
        cell_int = wks.Range("A2")
        cell_float = wks.Range("A3")
        cell_bool = wks.Range("A4")

        cell_str.Value = "ciao"
        cell_int.Value = 123
        cell_float.Value = 3.14159
        cell_bool.Value = True

        self.assertEqual(cell_str.Value, "ciao")
        self.assertEqual(cell_int.Value, 123)
        self.assertAlmostEqual(cell_float.Value, 3.14159)
        self.assertEqual(cell_bool.Value, True)

        wkb.Close(SaveChanges=False)

    def test_range_address(self):
        wkb = self.xl_app.Workbooks.Add()
        wks = wkb.Worksheets(1)

        r = wks.Range("B2:D4")
        self.assertEqual(r.Address(), "$B$2:$D$4")
        self.assertEqual(r.address(), "$B$2:$D$4")
        self.assertEqual(r.Address(ReferenceStyle=excel.xlR1C1), "R2C2:R4C4")
        self.assertEqual(r.address(ReferenceStyle=excel.xlR1C1), "R2C2:R4C4")

        wkb.Close(SaveChanges=False)

    def test_excel_aliases(self):
        wkb = self.xl_app.Workbooks.Add()

        self.assertIs(type(wkb.Worksheets.Add()), excel.Worksheet)
        self.assertIs(type(wkb.Worksheets.add()), excel.Worksheet)

        wkb.Close(SaveChanges=False)

    def test_excel_constants(self):
        self.assertEqual(excel.xlAscending, 1)
        self.assertEqual(excel.xlDescending, 2)

    def test_excel_com_object(self):
        wkb = self.xl_app.Workbooks.Add()
        self.assertIs(type(wkb), excel.Workbook)
        self.assertIs(type(wkb.com_object), win32com.client.CDispatch)
        wkb.Close(SaveChanges=False)

class TestPowerPoint(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.pp_app = powerpoint.Application().new()
        cls.pp_app.Visible = True

    @classmethod
    def tearDownClass(cls):
        cls.pp_app.Quit()

    def test_powerpoint(self):
        self.assertIs(type(self.pp_app), powerpoint.Application)
        self.assertTrue(self.pp_app.Visible)

        prs = self.pp_app.Presentations.Add()
        self.assertIs(type(prs), powerpoint.Presentation)

        prs = self.pp_app.Presentations.add()
        self.assertIs(type(prs), powerpoint.Presentation)

class TestOffice(unittest.TestCase):
    
    def test_access(self):
        acc_app = access.Application().new()
        self.assertIs(type(acc_app), access.Application)
        acc_app.Visible = True
        self.assertTrue(acc_app.Visible)
        acc_app.Quit()
    
    def test_outlook(self):
        ol_app = outlook.Application().new()
        self.assertIs(type(ol_app), outlook.Application)
        ol_app.Visible = True
        self.assertTrue(ol_app.Visible)
        ol_app.Quit()

    def test_word(self):
        wd_app = word.Application().new()
        self.assertIs(type(wd_app), word.Application)
        wd_app.Visible = True
        self.assertTrue(wd_app.Visible)
        wd_app.Quit()
        
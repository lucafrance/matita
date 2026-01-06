import unittest

from matita.office import access as ac, excel as xl, outlook as ol, powerpoint as pp, word as wd

class TestExcel(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.xl_app = xl.Application().new()
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
        self.assertEqual(r.Address(ReferenceStyle=xl.xlR1C1), "R2C2:R4C4")
        self.assertEqual(r.address(ReferenceStyle=xl.xlR1C1), "R2C2:R4C4")

        wkb.Close(SaveChanges=False)

    def test_excel_aliases(self):
        wkb = self.xl_app.Workbooks.Add()

        self.assertIs(type(wkb.Worksheets.Add()), xl.Worksheet)
        self.assertIs(type(wkb.Worksheets.add()), xl.Worksheet)

        wkb.Close(SaveChanges=False)

    def test_excel_constants(self):
        self.assertEqual(xl.xlAscending, 1)
        self.assertEqual(xl.xlDescending, 2)

    def test_excel_com_object(self):
        wkb = self.xl_app.Workbooks.Add()
        self.assertIs(type(wkb), xl.Workbook)
        self.assertIn("win32", str(type(wkb.com_object)))
        wkb.Close(SaveChanges=False)

    def test_range_operations(self):
        wkb = self.xl_app.Workbooks.Add()
        wks = wkb.worksheets(1)
        rng = wks.cells(2,3)
        self.assertEqual(rng.resize(4,5).address(), "$C$2:$G$5")


class TestPowerPoint(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.pp_app = pp.Application().new()
        cls.pp_app.Visible = True

    @classmethod
    def tearDownClass(cls):
        cls.pp_app.Quit()

    def test_powerpoint(self):
        self.assertTrue(self.pp_app.Visible)
        prs = self.pp_app.presentations.add()
        sld = pp.Slide(prs.slides.com_object.Add(1, pp.ppLayoutBlank))
        shp = sld.shapes.addshape(pp.msoShapeRectangle, 30, 30 , 30, 30)
        eff = sld.timeline.mainsequence.addeffect(
            Shape=shp,
            effectId=pp.msoAnimEffectFly,
            Level=pp.msoAnimateLevelNone,
            trigger=pp.msoAnimTriggerAfterPrevious,
        )
        prs.close()
        
    def test_powerpoint_com_object(self):
        prs = self.pp_app.presentations.add()
        self.assertIs(type(prs), pp.Presentation)
        self.assertIn("win32", str(type(prs.com_object)))
        prs.close()


class TestOffice(unittest.TestCase):
    
    def test_access(self):
        acc_app = ac.Application().new()
        self.assertIs(type(acc_app), ac.Application)
        acc_app.Visible = True
        self.assertTrue(acc_app.Visible)
        acc_app.Quit()
    
    def test_outlook(self):
        ol_app = ol.Application().new()
        self.assertIs(type(ol_app), ol.Application)
        ol_app.Visible = True
        self.assertTrue(ol_app.Visible)
        ol_app.Quit()

    def test_word(self):
        wd_app = wd.Application().new()
        self.assertIs(type(wd_app), wd.Application)
        wd_app.Visible = True
        self.assertTrue(wd_app.Visible)
        wd_app.Quit()
        
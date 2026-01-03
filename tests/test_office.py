import unittest

from matita.office import access, excel, outlook, powerpoint, word

class TestOffice(unittest.TestCase):
    
    def test_access(self):
        acc_app = access.Application().new()
        self.assertIs(type(acc_app), access.Application)
        acc_app.Visible = True
        self.assertTrue(acc_app.Visible)
        acc_app.Quit()
    
    def test_excel(self):
        # Create new excel Application
        xl_app = excel.Application().new()
        self.assertIs(type(xl_app), excel.Application)
        xl_app.Visible = True
        self.assertTrue(xl_app.Visible)

        # Add workbook
        wkb = xl_app.Workbooks.Add()
        self.assertIs(type(xl_app.Workbooks), excel.Workbooks)
        self.assertIs(type(wkb), excel.Workbook)
        wks = wkb.Worksheets(1)
        self.assertIs(type(wks), excel.Worksheet)

        # Write value to cell and read it back
        cell = wks.Range("A1")
        self.assertIs(type(cell), excel.Range)
        cell.Value = "Lorem Ipsum"
        self.assertEqual(cell.Value, "Lorem Ipsum")
        self.assertEqual(cell.Value2, "Lorem Ipsum")

        # Write value to another cell differently and read it back
        cell2 = wks.Cells(2,1)
        self.assertIs(type(cell2), excel.Range)
        cell2.Value = 12345.678
        self.assertEqual(cell2.Value, 12345.678)
        self.assertEqual(cell2.Value2, 12345.678)
        \
        # Test that `Range.Address` behaves as expected
        used_range = wks.UsedRange
        self.assertIs(type(used_range), excel.Range)
        self.assertEqual(used_range.Address(), "$A$1:$A$2")
        self.assertEqual(used_range.Address(ReferenceStyle=excel.xlR1C1), "R1C1:R2C1")

        # Test Range.Cells
        cell3 = used_range.Cells(2,1)
        self.assertEqual(cell3.Value2, 12345.678)

        # Add worksheet, change name, and read it back
        wks_ciao = wkb.Worksheets.Add()
        self.assertIs(type(wks_ciao), excel.Worksheet)
        wks_ciao.Name = "ciao"
        self.assertEqual(wks_ciao.Name, "ciao")
        wks_ciao = None
        wks_ciao = wkb.Worksheets("ciao")
        self.assertIs(type(wks_ciao), excel.Worksheet)
        
        # Close workbook without saving
        wkb.Close(SaveChanges=False)
        xl_app.Quit()

    def test_outlook(self):
        ol_app = outlook.Application().new()
        self.assertIs(type(ol_app), outlook.Application)
        ol_app.Visible = True
        self.assertTrue(ol_app.Visible)
        ol_app.Quit()

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
        
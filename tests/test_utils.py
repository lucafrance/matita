import unittest
import logging

from vipera.reference import MarkdownTree

test_text = open("tests/test_doc_page.md", "rt").read()
test_tree = MarkdownTree(test_text)

class TestParser(unittest.TestCase):

    def test_front_matter(self):
        fm = test_tree.front_matter.variables
        self.assertEqual(fm["title"], "Lorem Ipsum")
        self.assertEqual(fm["tags"], "consectetaur adipisicing")
        self.assertEqual(fm["color"], "sed do eiusmod tempor")
    
    def test_sections(self):
        self.assertEqual(test_tree.sections[0].title, "")
        self.assertEqual(test_tree.sections[0].paragraphs[0].txt, "At vero eos et accusamus.")
        self.assertEqual(test_tree.sections[0].paragraphs[1].txt, "Praesentium voluptatum.")
        self.assertEqual(test_tree.sections[1].title, "Section A")
        self.assertEqual(test_tree.sections[1].paragraphs[0].txt, "Lorem ipsum\ndolor sit amet.")
        self.assertEqual(test_tree.sections[1].level, 1)
        self.assertEqual(test_tree.sections[2].title, "Section B")
        self.assertEqual(test_tree.sections[2].paragraphs[0].txt, "Excepteur sint occaecat.")
        self.assertEqual(test_tree.sections[2].level, 2)
        self.assertEqual(test_tree.sections[6].title, "Section F")
        self.assertEqual(test_tree.sections[6].paragraphs[0].txt, "Sed quia consequuntur.")
        self.assertEqual(test_tree.sections[6].level, 2)
    
    def test_sections_by_level(self):
        self.assertEqual(test_tree.sections_by_level(2)[1].title, "Section C")
    
    def test_sections_by_title(self):
        self.assertEqual(test_tree.sections_by_title("on C")[0].title, "Section C")

    def test_table_parsing(self):
        table_paragraph = test_tree.sections_by_title("Section G")[0].paragraphs[1]
        self.assertTrue(table_paragraph.is_table)
        rows = table_paragraph.table.rows
        self.assertEqual(rows[0][0], "Lorem")
        self.assertEqual(rows[2][1], "e")
        self.assertEqual(rows[3][2], "i")


with open("pyvba/Excel.py", "rt") as f:
    excel_src = f.read()

class TestExcelModule(unittest.TestCase):

    def test_application_class(self):
        self.assertIn("    def new(self):\n        self.application = pyvba.genmodules.Excel.Application()\n        return self", excel_src)

    def test_collections(self):
        self.assertIn("def __call__(self, item):\n        return Worksheet(self.worksheets(item))", excel_src)

    def test_properties(self):
        self.assertIn("def __init__(self, aboveaverage=None):", excel_src)
        self.assertIn("@property\n    def Name(self):\n        return self.worksheet.Name", excel_src)
        self.assertIn("def Range(self, *args, Cell1=None, Cell2=None):\n        arguments = {\"Cell1\": Cell1, \"Cell2\": Cell2}\n        arguments = {key: value for key, value in arguments.items() if value is not None}\n        return Range(self.worksheet.Range(*args, **arguments))", excel_src)
        self.assertIn("@Visible.setter\n    def Visible(self, value):\n        self.application.Visible = value", excel_src)
        self.assertIn("@property\n    def Value(self):\n        return self.range.Value\n\n    @Value.setter\n    def Value(self, value):\n        self.range.Value = value", excel_src)
        self.assertIn("def Columns(self):\n        return Range(self.range.Columns)", excel_src)

    def test_methods(self):
        self.assertIn("def Quit(self):\n        self.application.Quit()", excel_src)
        self.assertIn("def Add(self, *args, Template=None):\n        arguments = {\"Template\": Template}\n        arguments = {key: value for key, value in arguments.items() if value is not None}\n        return Workbook(self.workbooks.Add(*args, **arguments))", excel_src)

import pyvba.Excel
import pyvba.genmodules.Excel

def is_ignored_page(txt):
    with open("ignored_pages.log", "rt") as f:
        if txt.lower() in f.read():
            return True
    return False

class TestModulesConsistency(unittest.TestCase):

    def test_compare_excel_modules(self):
        object_names = dir(pyvba.Excel)
        object_names.remove("pyvba")
        for object_name in object_names.copy():
            if object_name.startswith("_"):
                object_names.remove(object_name)
        # Check that all classes of `pyvba.Excel` also exist in `pyvba.genmodules.Excel` 
        for object_name in object_names:
            self.assertIn(object_name, dir(pyvba.genmodules.Excel))
        # Check that all properties and methods of `pyvba.genmodules.Excel` also exist in `pyvba.Excel`.
        # Only focus on classes which exist in `pyvba.Excel`, therefore reuse the list `object_names`.
        # If a class is not in the documentation, then it also not in `pyvba.Excel` and is not in scope.
        for object_name in object_names:
            # Ignore Excel Graph objects, not (yet) supported
            # https://learn.microsoft.com/en-us/office/vba/api/overview/excel/graph-visual-basic-reference
            excel_graph_objects =  ["ChartColorFormat", "ChartFillFormat"]
            # Ignore PivotLayout, as most methods of the COM are unavailable in VBA (e.g. `AddFields`, `GetColumnFields`)
            other_ignored_objects = ["ChartObjects", "PivotLayout"]
            if object_name in excel_graph_objects + other_ignored_objects:
                continue
            props_meths = dir(getattr(pyvba.genmodules.Excel, object_name))
            # Ignore properties and methods which are undefined in the documentation
            # Ignore properties and methods unavailable in VBA (but maybe available in the COM)
            undefined_props_meths = [
                ("Connections", "Add2"), ("PivotFilters", "Add2"), ("PivotTable", "ApplyLayout"),
                ("Range", "RefreshLinkedDataType"), ("SlicerCaches", "Add2"), ("WorksheetFunction", "ArrayToText"),
                ("WorksheetFunction", "Concat"), ("WorksheetFunction", "CoupPcd"), ("WorksheetFunction", "FieldValue"),
                ("WorksheetFunction", "Index"), ("WorksheetFunction", "MaxIfs"), ("WorksheetFunction", "MinIfs"),
                ("WorksheetFunction", "RandArray"), ("WorksheetFunction", "Sequence"), ("WorksheetFunction", "Single"),
                ("WorksheetFunction", "Sort"), ("WorksheetFunction", "SortBy"), ("WorksheetFunction", "StockHistory"),
                ("WorksheetFunction", "TextJoin"), ("WorksheetFunction", "Unique"), ("WorksheetFunction", "ValueToText"),
                ("WorksheetFunction", "XLookup"), ("WorksheetFunction", "XMatch"),
            ]
            unavailable_props_meths = [
                ("FillFormat", "Background"), ("LegendKey", "Select"), ("ListObject", "UpdateChanges"),
                ("OLEObjects", "Group"), ("PivotCaches", "Add"), ("PivotTable", "Dummy15"), ("PivotTable", "Dummy2"),
                ("PivotTable", "Format"), ("Range", "AutoFormat"), ("Range", "CreatePublisher"),
                ("Range", "GoalSeek"), ("Series", "ApplyCustomType"), ("Shape", "CanvasCropBottom"),
                ("Shape", "CanvasCropTop"), ("Shape", "CanvasCropLeft"), ("Shape", "CanvasCropRight"),
                ("Shapes", "AddCanvas"), ("Shapes", "AddDiagram"), ("WorksheetFunction", "Dummy19"),
                ("WorksheetFunction", "Dummy21"), ("WorksheetFunction", "IsThaiDigit"),
                ("WorksheetFunction", "RoundBahtDown"), ("WorksheetFunction", "RoundBahtUp"),
                ("WorksheetFunction", "ThaiDayOfWeek"), ("WorksheetFunction", "ThaiDigit"),
                ("WorksheetFunction", "ThaiMonthOfYear"), ("WorksheetFunction", "ThaiNumSound"),
                ("WorksheetFunction", "ThaiNumString"), ("WorksheetFunction", "ThaiStringLength"),
                ("WorksheetFunction", "ThaiYear"),
            ]
            for obj, prop_meth in undefined_props_meths + unavailable_props_meths:
                if object_name == obj:
                    props_meths.remove(prop_meth)
            # Ignore properties and methods specific to the module generated with pywin32
            for pm in ["CLSID", "GetProperty", "SetProperty"]:
                if pm in props_meths:
                    props_meths.remove(pm)
            for pm in props_meths.copy():
                if "_" in pm or pm.startswith("coclass") or pm.startswith("Set"):
                    props_meths.remove(pm)
            for pm in props_meths:
                object_attr = dir(getattr(pyvba.Excel, object_name))
                try:
                    self.assertTrue((pm in object_attr) or (pm.removeprefix("Get") in object_attr))
                except AssertionError as e:
                    full_pm_name = f"{object_name}.{pm}"
                    if is_ignored_page(full_pm_name):
                        logging.warning(f"{full_pm_name} exists in genomodules, but the doc page has been ignored.")
                    else:
                        print(f"Property/method {object_name}.{pm} from genmodules not found in pyvba.")
                        print(f"object_attr == {object_attr}")
                        raise(e)
                except Exception as e:
                    raise(e)

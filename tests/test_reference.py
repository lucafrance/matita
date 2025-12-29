import unittest

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


with open("src/vipera/office/excel.py", "rt") as f:
    excel_src = f.read()

class TestExcelModule(unittest.TestCase):

    def test_application_class(self):
        self.assertTrue("    def new(self):\n        self.application = win32com.client.Dispatch(\"Excel.Application\")\n        return self" in excel_src, msg="def new... not found")

    def test_collections(self):
        self.assertTrue("def __call__(self, item):\n        return Worksheet(self.worksheets(item))" in excel_src, msg="__call__... not found")

    def test_properties(self):
        self.assertTrue("def __init__(self, aboveaverage=None):" in excel_src, msg="__init__... not found")
        self.assertTrue("@property\n    def Name(self):\n        return self.worksheet.Name" in excel_src, msg="@property Name... not found")
        self.assertTrue("def Range(self, *args, Cell1=None, Cell2=None):\n        arguments = {\"Cell1\": Cell1, \"Cell2\": Cell2}\n        arguments = {key: value for key, value in arguments.items() if value is not None}\n        return Range(self.worksheet.Range(*args, **arguments))" in excel_src, msg="def Range... not found")
        self.assertTrue("@Visible.setter\n    def Visible(self, value):\n        self.application.Visible = value" in excel_src, msg="@Visible.setter... not found")
        self.assertTrue("@property\n    def Value(self):\n        return self.range.Value\n\n    @Value.setter\n    def Value(self, value):\n        self.range.Value = value" in excel_src, msg="@property Value... not found")
        self.assertTrue("def Columns(self):\n        return Range(self.range.Columns)" in excel_src, msg="def Columns... not found")

    def test_methods(self):
        self.assertTrue("def Quit(self):\n        self.application.Quit()" in excel_src, msg="def Quit... not found")
        self.assertTrue("def Add(self, *args, Template=None):\n        arguments = {\"Template\": Template}\n        arguments = {key: value for key, value in arguments.items() if value is not None}\n        return Workbook(self.workbooks.Add(*args, **arguments))" in excel_src, msg="def Add... not found")

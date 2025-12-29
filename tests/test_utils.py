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


with open("src/vipera/excel.py", "rt") as f:
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

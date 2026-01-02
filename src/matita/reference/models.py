import logging
import os

from .markdown import MarkdownTree

def reset_ignored_pages():
    f = open("logs/ignored_pages.log", "wt")
    f.close()

def log_ignored_page(page_key, page):
    with open("logs/ignored_pages.log", "at") as f:
        if page.title is not None:
            f.write(page.title.lower())
        else:
            f.write(page_key.lower())
        f.write("\n")

def page_filename_to_key(filename):
    key = filename.removesuffix(".md")
    # Remove portions of the filename which are not part of the api name
    # e.g. "(Object)" from "Excel.Application(Object)"
    key = key.split("(", 1)[0]
    return key

class ComAppCache:
    _instances = {}

    @classmethod
    def get(cls, module_name):
        if module_name not in cls._instances:
            progid = f"{module_name}.Application"
            cls._instances[module_name] = win32com.client.Dispatch(progid)
        return cls._instances[module_name]

class DocPage:

    def __init__(self, markdown_src):
        self.md = markdown_src
        self.md_tree = None
        self.title = None
        self.api_name = None

        self.module_name = None
        self.object_name = None
        self.property_name = None
        self.method_name = None
        
        self.is_object = None
        self.is_collection = None
        self.is_property = None
        self.is_method = None
        self.is_enumeration = None

        self.is_read_only_property = None
        self.has_return_value = None
        self.property_class = None
        self.return_value_class = None

        self.enumeration_values = None

        self.properties = []
        self.parameters = []
        self.methods = []
    
    def process_page(self):
        try:
            self.md_tree  = MarkdownTree(self.md)
        except Exception as e:
            logging.warning("Failed building MarkdownTree, raising an exception.")
            raise(e)
        
        self.title = self.md_tree .front_matter.variables["title"]
        if "api_name" in self.md_tree.front_matter.variables:
            self.api_name = self.md_tree.front_matter.variables["api_name"]
        self.process_title()
        self.process_api_name()

        if self.is_enumeration:
            self.process_enumeration()

        # Retrieve information from the opening paragraph
        # Examples
        # -------------------------------------------
        # A collection of all the **[Worksheet](Excel.Worksheet.md)** objects in the specified or active workbook. Each **Worksheet** object represents a worksheet.
        # Returns or sets a **String** value that represents the object name.
        # Returns a **[Range](Excel.Range(object).md)** object that represents all the cells on the worksheet (not just the cells that are currently in use).
        # Returns a **Range** object that represents the columns in the specified range.
        # -------------------------------------------
        p = self.md_tree.sections_by_level(1)[0].paragraphs[0].txt
        # Check whether the object is a collection
        self.is_collection = p.startswith("A collection")

        # Get the return type of the property (if any)
        if "Returns" in p:
            property_class = None
            if "**[" in p:
                property_class = p.split("**[", 1)[1].split("]", 1)[0]
            elif "**" in p:
                property_class = p.split("**", 1)[1].split("**", 1)[0]
            self.property_class = property_class

        # Check whether the property is read only
        if "Returns or sets" in p or "Read/write" in p:
            self.is_read_only_property = False
        else:
            self.is_read_only_property = True

        # Find the parameters of a property.
        # The "Syntax" section should look like this.
        # -------------------------------------------
        # ## Syntax
        # _expression_.**Range** (_Cell1_, _Cell2_)
        # -------------------------------------------
        sections = self.md_tree.sections_by_title("Syntax")
        if len(sections) > 0:
            line = sections[0].paragraphs[0].txt.replace("()", "")
            if "(" in line:
                parameters = line.split("(", 1)[1].split(")", 1)[0]
                parameters = parameters.split(", ")
                parameters = [p.strip(" _`[]*\\").replace("\\_", "_") for p in parameters]
                self.parameters = parameters
        
        # Find the return value of a property. The section looks like this:
        # Example
        # -------------------------------------------
        # ## Return value
        # A **[Workbook](Excel.Workbook.md)** object that represents the new workbook.
        # -------------------------------------------
        self.has_return_value = False
        sections = self.md_tree.sections_by_title("Return value")
        if len(sections) > 0:
            self.has_return_value = True
            s = sections[0]
            line = s.paragraphs[0].txt.splitlines()[0]
            if "**[" in line:
                self.return_value_class = line.split("**[", 1)[1].split("](")[0]

    def process_title(self):
        if self.title is None:
            self.is_object = None
            self.is_property = None
            self.is_method = None
            self.is_enumeration = None
            return
        else:
            self.is_object = False
            self.is_property = False
            self.is_method = False
            self.is_enumeration = False

        if "object" in self.title:
            self.is_object = True
        elif "property" in self.title:
            self.is_property = True
        elif "method" in self.title:
            self.is_method = True
        elif "enumeration" in self.title:
            self.is_enumeration = True
    
    def process_api_name(self):
        if self.api_name is None:
            return
        api_name_split = self.api_name.split(".")
        self.module_name = api_name_split[0]
        self.object_name = api_name_split[1]
        if len(api_name_split) > 2:
            suffix = api_name_split[2].strip()
            if self.is_property:
                self.property_name = suffix
            elif self.is_method:
                self.method_name = suffix
    
    def process_enumeration(self):
        """Parse enumeration information. Assumes that the page refers to an enumeration.

        Example: [XlReferenceType enumeration (Excel)](https://learn.microsoft.com/en-gb/office/vba/api/excel.xlreferencetype)
        """

        # Find table paragraph in the first section
        table = None
        for p in self.md_tree.sections_by_level(1)[0].paragraphs:
            if p.is_table:
                try:
                    table = p.table
                except Exception as e:
                    logging.WARNING(f"Failed parsing values table for enumeration {self.api_name}. {e}")
                break
        
        if "Name" not in table.rows[0] or "Value" not in table.rows[0]:
            logging.WARNING(f"Unexpected header for values table for enumeration {self.api_name}.")
            return
        
        self.enumeration_values = {}
        for r in table.rows[1:]:
            key = r[0].strip("*")
            value = r[1]
            if value.lstrip("-").isnumeric():
                if int(value) == float(value):
                    value = int(value)
                else:
                    value = float(value)
            self.enumeration_values[key] = value
    
    def parent_object_key(self):
        if (self.module_name and self.object_name) is None:
            return None
        return ".".join([self.module_name, self.object_name]).lower()
    
    def to_dict(self):
        return {
            "title": self.title,
            "module_name": self.module_name,
            "object_name": self.object_name,
            "property_name": self.property_name,
            "method_name": self.method_name,
            "is_object": self.is_object,
            "is_collection": self.is_collection,
            "is_method": self.is_method,
            "is_property": self.is_property,
            "is_enumeration": self.is_enumeration,
            "api_name": self.api_name,
            "is_collection": self.is_collection,
            "is_read_only_property": self.is_read_only_property,
            "parent object key": self.parent_object_key(),
            "property_class": self.property_class,
            "return_value_class": self.return_value_class,
            "enumeration_values": self.enumeration_values,
            "properties": [page.title for page in self.properties],
            "parameters": self.parameters,
            "methods": [page.title for page in self.methods],
        }
    
    def to_python(self):
        """Return source code of a python class based on the page"""

        if self.is_enumeration:
            code = []
            code.append(f"# {self.object_name} enumeration")
            for key, value in self.enumeration_values.items():
                if type(value) == str:
                    value = f"\"{value}\""
                code.append(f"{key} = {value}")
            code.append("")
            return "\n".join(code)

        if not self.is_object:
            return ""

        code = []
        code.append("class " + self.object_name + ":")
        code.append("")
        code.append("    def __init__(self, " + self.object_name.lower() + "=None):")
        code.append("        self." + self.object_name.lower() + " = " + self.object_name.lower())
        code.append("")

        # New method for Application objects
        if self.object_name == "Application":
            code.append(f"    def new(self):")
            code.append(f"        self.application = win32com.client.Dispatch(\"{self.module_name}.Application\")")
            code.append(f"        return self")
            code.append(f"")

        # Call method for collections
        if self.is_collection:
            code.append(f"    def __call__(self, item):")
            code.append(f"        return {self.object_name[:-1]}(self.{self.object_name.lower()}(item))")
            code.append("")

        code += self.to_python_properties()
        code += self.to_python_methods()
        code.append("")
        return "\n".join(code)

    def parameters_code(self):
        """Return code to include parameters as arguments of a property or method"""
        return "=None, ".join(self.parameters) + "=None"

    def to_python_arguments_expansion(self):
        """Return code for non-None arguments extraction to an 'arguments' dictionary

        Example output
        -------------------------------------------
        arguments = {"SaveChanges": SaveChanges, "FileName": FileName, "RouteWorkbook": RouteWorkbook}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        -------------------------------------------
        """
        code = []
        code_line = [f"\"{arg}\": {arg}" for arg in self.parameters]
        code_line = "arguments = {" + ", ".join(code_line) + "}"
        code.append(" "*8 + code_line)
        code.append("        arguments = {key: value for key, value in arguments.items() if value is not None}")
        return code

    def to_python_properties(self):
        """Return python code for all properties of the object"""
        code = []
        for p in self.properties:
            if p.property_name is None:
                logging.info("Property '{}' ignored when exporting '{}', because property_name is None.".format(p.title, self.title))
                continue

            # Getter method
            if len(p.parameters) == 0 or not p.is_read_only_property:
                if p.property_class is not None:
                    code.append("    @property")
                    code.append(f"    def {p.property_name}(self):")
                    code.append(f"        return {p.property_class}(self.{self.object_name.lower()}.{p.property_name})")
                    code.append(f"")
                else:
                    code.append("    @property")
                    code.append(f"    def {p.property_name}(self):")
                    code.append(f"        return self.{self.object_name.lower()}.{p.property_name}")
                    code.append(f"")
            else:
                code.append(f"    def {p.property_name}(self, *args, {p.parameters_code()}):")
                code += p.to_python_arguments_expansion()
                if p.property_class is not None:
                    code.append(f"        if callable(self.{self.object_name.lower()}.{p.property_name}):")
                    code.append(f"            return {p.property_class}(self.{self.object_name.lower()}.{p.property_name}(*args, **arguments))")
                    code.append(f"        else:")
                    code.append(f"            return {p.property_class}(self.{self.object_name.lower()}.Get{p.property_name}(*args, **arguments))")
                    code.append(f"")
                else:
                    code.append(f"        if callable(self.{self.object_name.lower()}.{p.property_name}):")
                    code.append(f"            return self.{self.object_name.lower()}.{p.property_name}(*args, **arguments)")
                    code.append(f"        else:")
                    code.append(f"            return self.{self.object_name.lower()}.Get{p.property_name}(*args, **arguments)")
                    code.append(f"")

            # Setter method
            # If the property is editable, it must have a setter method.
            # If so, no argument can be used in the setter method.
            if not p.is_read_only_property:
                code.append(f"    @{p.property_name}.setter")
                code.append(f"    def {p.property_name}(self, value):")
                code.append(f"        self.{self.object_name.lower()}.{p.property_name} = value")
                code.append("")

        return code
    
    def to_python_methods(self):
        """Return python code for all methods of the object"""
        code = []
        for m in self.methods:
            if len(m.parameters) == 0:
                code.append(f"    def {m.method_name}(self):")
                code_line = f"self.{self.object_name.lower()}.{m.method_name}()"
            else:
                code.append(f"    def {m.method_name}(self, *args, {m.parameters_code()}):")
                code += m.to_python_arguments_expansion()
                # Actual method invocation
                code_line = f"self.{self.object_name.lower()}.{m.method_name}(*args, **arguments)"
            if m.has_return_value:
                if m.return_value_class is not None:
                    code_line = f"{m.return_value_class}({code_line})"
                # Certain methods of a collection can only return certain types
                # e.g. `Worksheets.Add`` returns a `Worksheet`
                # `.startswith("Open") is for methods like `Workbooks.OpenText`
                elif self.is_collection and \
                    (m.method_name == "Add" or m.method_name.startswith("Open")):
                    code_line = f"{self.object_name[:-1]}({code_line})"
                code_line = "return " + code_line
            code_line = " "*8 + code_line
            code.append(code_line)
            code.append("")

        return code

class VbaDocs:

    def __init__(self):
        self.pages = dict()

    def read_directory(self, path):
        with os.scandir(path) as it:
            for entry in it:
                if entry.is_file():
                    print(f"Parsing {entry.name}...")
                    page_key = page_filename_to_key(entry.name).lower()
                    if page_key in ["word.break", "word.global"]:
                        logging.info("Ignoring page '{}', because it conflicts with a Python keyword.".format(entry.name))
                    elif page_key in ["access.report.circle", "access.report.line"]:
                        logging.info("Ignoring page '{}', because non-scalar arguments are not implemented.".format(entry.name))
                    elif "-" in page_key:
                        logging.info("Ignoring page '{}', because the object name includes a dash.".format(entry.name))
                    else:
                        self.pages[page_key] = DocPage(open(entry, "rt", encoding="utf8").read())
    
    def process_pages(self):
        reset_ignored_pages()
        pages_to_remove = []
        for page_key, page in self.pages.items():
            try:
                page.process_page()
                # Remove pages without `api_name`
                if page.api_name is None:
                    logging.warning(f"Attribute `api_name` not found for {page_key}, ignoring.")
                    log_ignored_page(page_key, page)
                    pages_to_remove.append(page_key)
            except Exception as e:
                logging.error(f"Failed processing page: '{page_key}'. {e}")
        for key in pages_to_remove:
            del self.pages[key]
        # Populate properties and methods of objects
        for page_key, page in self.pages.items():
            parent_object_key = page.parent_object_key()
            if parent_object_key is not None:
                if page.is_property:
                    if parent_object_key in self.pages:
                        self.pages[parent_object_key].properties.append(page)
                    else:
                        logging.warning(f"Parent object '{parent_object_key}' for property '{page_key}' not found.")
                elif page.is_method:
                    if parent_object_key in self.pages:
                        self.pages[parent_object_key].methods.append(page)
                    else:
                        logging.warning(f"Parent object '{parent_object_key}' for method '{page_key}' not found.")
            else:
                if page.is_property or page.is_method:
                    logging.warning(f"Page'{page_key}' is a property or method, but the key of the parent object of is None.")
        # Remove invalid class types
        for page in self.pages.values():
            if page.property_class is not None:
                if f"{page.module_name}.{page.property_class}".lower() not in self.pages:
                    page.property_class = None
        self.apply_manual_adjustments()
    
    def to_dict(self):
        dictionaries = [page.to_dict() for page in self.pages.values()]
        keys_and_values = zip(self.pages.keys(), dictionaries)
        return {key: value for key, value in keys_and_values}

    def apply_manual_adjustments(self):
        # Add parameters for Cell properties, whose parameters are not properly imported
        for p in self.pages.values():
            if p.property_name == "Cells" and len(p.parameters) == 0:
                p.parameters = ["RowIndex", "ColumnIndex"]

    def to_python(self, application):
        code = ["import win32com.client", ""]
        for page_key, page in self.pages.items():
            if page.module_name is None:
                continue
            if page.module_name.lower() != application.lower():
                continue
            try:
                page_code = page.to_python()
            except Exception as e:
                logging.warning(f"Can't export '{page_key}' to python code. {e}")
            else:
                if page_code is not None:
                    code.append(page_code)
        return "\n".join(code)
    
    def apply_manual_adjustments(self):
        # Add parameters for Cell properties, whose parameters are not properly imported
        for p in self.pages.values():
            if p.property_name == "Cells" and len(p.parameters) == 0:
                p.parameters = ["RowIndex", "ColumnIndex"]

import logging
import os
import traceback

from utils.markdown_parser import MarkdownTree

def reset_ignored_pages():
    f = open("ignored_pages.log", "wt")
    f.close()

def log_ignored_page(page_key, page):
    with open("ignored_pages.log", "at") as f:
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

class VbaDocs:

    def __init__(self):
        self.pages = dict()

    def read_directory(self, path):
        with os.scandir(path) as it:
            for entry in it:
                if entry.is_file() and entry.name.startswith("Excel."):
                    page_key = page_filename_to_key(entry.name).lower()
                    if "-" in page_key:
                        logging.info("Ignoring page '{}', because the object name includes a dash.".format(entry.name))
                    else:
                        self.pages[page_key] = DocPage(open(entry, "rt", encoding="utf8").read())
    
    def process_pages(self):
        reset_ignored_pages()
        for page_key, page in self.pages.items():
            try:
                page.process_page()
                # Remove pages without `api_name`
                if page.api_name is None:
                    logging.warning(f"Attribute `api_name` not found for {page_key}, ignoring.")
                    log_ignored_page(page_key, page)
                    self.pages.pop[page_key]
            except Exception as e:
                logging.error("Failed processing page: '{}'. {}".format(page_key, e))
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

class DocPage:

    def __init__(self, markdown_src):
        self.md = markdown_src
        self.title = None
        self.api_name = None

        self.module_name = None
        self.object_name = None
        self.property_name = None
        self.method_name = None

        self.is_read_only_property = None
        self.has_return_value = None
        self.property_class = None
        self.return_value_class = None

        self.properties = []
        self.parameters = []
        self.methods = []
        
        self.is_object = None
        self.is_collection = None
        self.is_property = None
        self.is_method = None
    
    def process_page(self):
        try:
            t = MarkdownTree(self.md)
        except Exception as e:
            logging.warning("Failed building MarkdownTree, raising an exception.")
            raise(e)
        
        self.title = t.front_matter.variables["title"]
        if "api_name" in t.front_matter.variables:
            self.api_name = t.front_matter.variables["api_name"]

        # Retrieve information from the opening paragraph
        # Examples
        # -------------------------------------------
        # A collection of all the **[Worksheet](Excel.Worksheet.md)** objects in the specified or active workbook. Each **Worksheet** object represents a worksheet.
        # Returns or sets a **String** value that represents the object name.
        # Returns a **[Range](Excel.Range(object).md)** object that represents all the cells on the worksheet (not just the cells that are currently in use).
        # Returns a **Range** object that represents the columns in the specified range.
        # -------------------------------------------
        p = t.sections_by_level(1)[0].paragraphs[0].txt
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
        if "Returns or sets" in p:
            self.is_read_only_property = False
        else:
            self.is_read_only_property = True

        # Find the parameters of a property.
        # The "Syntax" section should look like this.
        # -------------------------------------------
        # ## Syntax
        # _expression_.**Range** (_Cell1_, _Cell2_)
        # -------------------------------------------
        sections = t.sections_by_title("Syntax")
        if len(sections) > 0:
            line = sections[0].paragraphs[0].txt.replace("()", "")
            if "(" in line:
                parameters = line.split("(", 1)[1].split(")", 1)[0]
                parameters = parameters.split(", ")
                parameters = [p.strip("_") for p in parameters]
                self.parameters = parameters
        
        # Find the return value of a property. The section looks like this:
        # Example
        # -------------------------------------------
        # ## Return value
        # A **[Workbook](Excel.Workbook.md)** object that represents the new workbook.
        # -------------------------------------------
        self.has_return_value = False
        sections = t.sections_by_title("Return value")
        if len(sections) > 0:
            self.has_return_value = True
            s = sections[0]
            line = s.paragraphs[0].txt.splitlines()[0]
            if "**[" in line:
                self.return_value_class = line.split("**[", 1)[1].split("](")[0]

        self.process_title()
        self.process_api_name()

    def process_title(self):
        if self.title is None:
            self.is_object = None
            self.is_property = None
            self.is_method = None
            return
        else:
            self.is_object = False
            self.is_property = False
            self.is_method = False

        if "object" in self.title:
            self.is_object = True
        elif "property" in self.title:
            self.is_property = True
        elif "method" in self.title:
            self.is_method = True
    
    def process_api_name(self):
        if self.api_name is None:
            return
        api_name_split = self.api_name.split(".")
        self.module_name = api_name_split[0]
        self.object_name = api_name_split[1]
        if len(api_name_split) > 2:
            suffix = api_name_split[2]
            if self.is_property:
                self.property_name = suffix
            elif self.is_method:
                self.method_name = suffix
    
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
            "api_name": self.api_name,
            "is_collection": self.is_collection,
            "is_read_only_property": self.is_read_only_property,
            "parent object key": self.parent_object_key(),
            "property_class": self.property_class,
            "return_value_class": self.return_value_class,
            "properties": [page.title for page in self.properties],
            "parameters": self.parameters,
            "methods": [page.title for page in self.methods],
        }

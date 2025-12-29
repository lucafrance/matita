import logging
import os

from .models import DocPage


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
                logging.error("Failed processing page: '{}'. {}".format(page_key, e))
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

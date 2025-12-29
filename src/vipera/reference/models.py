import logging

from .markdown import MarkdownTree


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
    
    def to_python(self):
        """Return source code of a python class based on the page"""

        if not self.is_object:
            return "\n"

        code = []
        code.append("class " + self.object_name + ":")
        code.append("")
        code.append("    def __init__(self, " + self.object_name.lower() + "=None):")
        code.append("        self." + self.object_name.lower() + " = " + self.object_name.lower())
        code.append("")

        # New method for Application objects
        if self.object_name == "Application" and self.is_object:
            code.append("    def new(self):")
            code.append(f"        self.application = pyvba.genmodules.{self.module_name}.Application()")
            code.append("        return self")
            code.append("")

        # Call method for collections
        if self.is_collection:
            code.append("    def __call__(self, item):")
            code.append(f"        return {self.object_name[:-1]}(self.{self.object_name.lower()}(item))")
            code.append("")

        code += self.to_python_properties()
        code += self.to_python_methods()
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
            # If the property is not read only, it must have a setter method. If so, no argument can be used in the setter method.
            # E.g. `Range.Value(RangeValueDataType)`
            if len(p.parameters) == 0 or not p.is_read_only_property:
                code.append("    @property")
                code.append(f"    def {p.property_name}(self):")
                code_line = f"self.{self.object_name.lower()}.{p.property_name}"
            else:
                code.append(f"    def {p.property_name}(self, *args, {p.parameters_code()}):")
                code += p.to_python_arguments_expansion()
                # If the genmodule includes a Get... method for a property, use that one.
                # E.g. `Range.GetAddress` for `Range.Address`
                if f"Get{p.property_name}" in dir(self.genmodule_object()):
                    code_line = f"self.{self.object_name.lower()}.Get{p.property_name}(*args, **arguments)"
                else:
                    code_line = f"self.{self.object_name.lower()}.{p.property_name}(*args, **arguments)"
            # If there is a class for the property, wrap it
            if p.property_class is not None:
                code_line = f"{p.property_class}({code_line})"
            code.append("        return " + code_line)
            code.append("")

            if p.is_read_only_property:
                continue
            
            # Setter method
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

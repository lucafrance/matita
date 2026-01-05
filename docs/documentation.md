# Matita ✏️ - Documentation

## Get started

You need to create an object for the application you need.

Unlike starting Microsoft Office normally, all application objects are by default invisible.
I recommend making the application visible as long as you are developing.
Once your code is stable, you may decide to keep the application invisible unless an exception aris

```python
from matita.office import access as ac, excel as xl, outlook as ol, powerpoint as pp, word as wd

ac_app = ac.Application().new()
ac_app.visible = True

xl_app = xl.Application().new()
xl_app.visible = True

ol_app = ol.Application().new()
ol_app.visible = True

pp_app = pp.Application().new()
pp_app.visible = True

wd_app = wd.Application().new()
wd_app.visible = True
```

With the application object created, you can start creating [documents](https://learn.microsoft.com/en-gb/office/vba/api/word.documents), [workbooks](https://learn.microsoft.com/en-gb/office/vba/api/excel.workbooks), [presentations](https://learn.microsoft.com/en-gb/office/vba/api/powerpoint.presentations), emails, and so on.

```python
# Create a new Excel workbook
wkb = xl_app.workbooks.add()

# Create a new PowerPoint presentation
ppt = pp_app.presentations.add()

# Create a new Word document
doc = wd_app.documents.add()
```

You can also open existing files.

```python
# Open an existing Access database
ac_db = ac_app.databases.OpenCurrentDatabase("C:\\path\\to\\your\\database.accdb")

# Open an existing Excel workbook
wkb = xl_app.workbooks.open("C:\\path\\to\\your\\workbook.xlsx")   

# Open an existing PowerPoint presentation
ppt = pp_app.presentations.open("C:\\path\\to\\your\\presentation.pptx")

# Open an existing Word document
doc = wd_app.documents.open("C:\\path\\to\\your\\document.docx")
```

You have access to all objects, methods, and properties of the Office VBA Object Library.
Consult the [Office VBA Reference](https://learn.microsoft.com/en-us/office/vba/api/overview) for details.

## Comparison with other Python packages


`Matita` wraps Microsoft Office [COM](https://learn.microsoft.com/en-us/windows/win32/com/the-component-object-model) objects created with [`pywin32`](https://pypi.org/project/pywin32/) and provides a Pythonic interface that closely matches the VBA syntax.

Every `matita.office` class includes an underlying COM object, accessible via the `com_object` property.

```python
from matita.office import word as wd

wd_app = wd.Application().new()
print(type(wd_app)) # <class 'matita.office.word.Application'>
print(type(wd_app.com_object)) # <class 'win32com.client.CDispatch'>
wd_app.Quit()
```

For Excel specifically [`xlwings`](https://docs.xlwings.org/en/latest/missing_features.html) follows a similar approach.

This is different from other popular Python packages for Office automation, such as [`openpyxl`](https://openpyxl.readthedocs.io) for Excel, [`python-docx`](https://python-docx.readthedocs.io) for Word, or [`python-pptx`](https://pypi.org/project/python-pptx/) for PowerPoint, which implement their own object models and do not use the Office VBA Object Library.

## Parser for the Office VBA Reference

This project is based on the [Office VBA Reference](https://learn.microsoft.com/en-us/office/vba/api/overview) by Microsoft Corporation, [licensed](https://github.com/MicrosoftDocs/VBA-Docs/blob/main/LICENSE) under [Creative Commons Attribution 4.0 International](https://creativecommons.org/licenses/by/4.0/).

The subpackage `matita.reference`:
- parses of the [Office VBA Reference](https://learn.microsoft.com/en-us/office/vba/api/overview),
- saves the object model to `data/office-vba-api.json`,
- creates the subpackage `matita.office`.

### Advantages of `matita` over `pywin32` directly

Lower case aliases for all objects, methods, and properties. Consistent with Python naming conventions.

```python
import win32com.client

xl_app = win32com.client.Dispatch("Excel.Application")
# Does not work
# xl_app.visible = True
xl_app.Visible = True

# Does not work
# wkb = xl_app.Workbooks.add()
wkb = xl_app.Workbooks.Add()
```

```python
from matita.office import excel as xl

xl_app = xl.Application().new()

# Equally valid
xl_app.visible = True
xl_app.Visible = True

# Equally valid
wkb1 = xl_app.Workbooks.add()
wkb2 = xl_app.Workbooks.Add()
```

For compatibility reasons `pywin32` sometimes uses different names for methods and properties compared to the Office VBA Object Library. All names in `matita` match the VBA reference.

Enumerations values can be retrieved with `pywin32`, but they are easier to access with `matita`.

```python
import win32com.client

xl_app = win32com.client.gencache.EnsureDispatch("Excel.Application")
constants = win32com.client.constants
xl_app.Visible = True
wkb = xl_app.Workbooks.Add()
wks = wkb.Worksheets(1)
c = wks.Cells(1,1)
# Fails, Range.Address is just a string
# print(c.Address(ReferenceStyle=constants.xlR1C1))
print(c.GetAddress(ReferenceStyle=constants.xlR1C1)) #R1C1
```

```python
from matita.office import excel as xl

xl_app = xl.Application().new()
xl_app.visible = True
wkb = xl_app.Workbooks.add()
wks = wkb.Worksheets(1)
c = wks.Range("A1")

# Works, Range.Address is a method with arguments
print(c.Address(ReferenceStyle=xl.xlR1C1)) #R1C1
```

## Limitations of `matita`

The following objects are unsupported, because their name conflicts with reserved keywords in Python.
- [Break object (Word)](https://learn.microsoft.com/en-us/office/vba/api/word.break)
- [Global object (Word)](https://learn.microsoft.com/en-us/office/vba/api/word.global)

The following objects are unsupported, because non-scalar arguments are not implemented.
- [Report.Circle method (Access)](https://learn.microsoft.com/en-gb/office/vba/api/access.report.circle)
- [Report.Line method (Access)](https://learn.microsoft.com/en-gb/office/vba/api/access.report.line)

# Matita ✏️ - Documentation

## Get started

You need to create an object for the application you need.

Unlike starting Microsoft Office normally, all application objects are by default invisible.
I recommend making the application visible as long as you are developing.
Once your code is stable, you may decide to keep the application invisible unless an exception aris

```python
from matita.office import access, excel, outlook, powerpoint, word

ac_app = access.application.new()
ac_app.visible = True

xl_app = excel.application.new()
xl_app.visible = True

ol_app = outlook.application.new()
ol_app.visible = True

pp_app = powerpoint.application.new()
pp_app.visible = True

wd_app = word.application.new()
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

## Parser for the Office VBA Reference

This project is based on the [Office VBA Reference](https://learn.microsoft.com/en-us/office/vba/api/overview) by Microsoft Corporation, [licensed](https://github.com/MicrosoftDocs/VBA-Docs/blob/main/LICENSE) under [Creative Commons Attribution 4.0 International](https://creativecommons.org/licenses/by/4.0/).

The subpackage `matita.reference`:
- parses of the [Office VBA Reference](https://learn.microsoft.com/en-us/office/vba/api/overview),
- saves the object model to `data/office-vba-api.json`,
- creates the subpackage `matita.office`.

## Limitations

The following objects are unsupported, because their name conflicts with reserved keywords in Python.
- [Break object (Word)](https://learn.microsoft.com/en-us/office/vba/api/word.break)
- [Global object (Word)](https://learn.microsoft.com/en-us/office/vba/api/word.global)

The following objects are unsupported, because non-scalar arguments are not implemented.
- [Report.Circle method (Access)](https://learn.microsoft.com/en-gb/office/vba/api/access.report.circle)
- [Report.Line method (Access)](https://learn.microsoft.com/en-gb/office/vba/api/access.report.line)

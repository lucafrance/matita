# Matita ✏️

*Matita* is Python wrapper for the [Office VBA Object library](https://learn.microsoft.com/en-us/office/vba/api/overview/).
It is designed to  match the VBA syntax as much as possible.
There are modules for Microsoft Access, Excel, Outlook, PowerPoint, Word.
It can be used for Microsoft Office automation.

```python
from matita.office import excel

xl_app = excel.Application().new()
xl_app.Visible = True

wkb = xl_app.Workbooks.Add()
wks = wkb.Worksheets(1)
wks.Cells(1,1).Value = "Hello World"
```

```vba
Option Explicit

Sub example()
    Dim xl_app As Excel.Application
    Set xl_app = New Excel.Application
    xl_app.Visible = True
    
    Dim wkb As Workbook
    Set wkb = xl_app.Workbooks.Add()
    
    Dim wks As Worksheet
    Set wks = wkb.Worksheets(1)
    
    wks.Cells(1, 1).Value = "Hello World"
End Sub
```

## Installation

Install the package with:

```powershell
python -m pip install matita
```

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

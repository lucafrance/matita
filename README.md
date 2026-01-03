# Matita ✏️

*Matita* is a Python wrapper for the [Office VBA Object library](https://learn.microsoft.com/en-us/office/vba/api/overview/).
It is designed to match the VBA syntax as much as possible.
There are modules for Microsoft Access, Excel, Outlook, PowerPoint, Word.
It can be used for Microsoft Office automation.

```python
from matita.office import excel

xl_app = excel.Application().new()
xl_app.Visible = True

wkb = xl_app.Workbooks.Add()
wks = wkb.Worksheets(1)
c = wks.Cells(1,1)

c.Value = "Hello World"
print(c.Address(excel.xlR1C1))
```

VBA equivalent:

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
    
    Dim c As Range
    Set c = wks.Cells(1, 1)
    
    c.Value = "Hello World"
    Debug.Print c.Address(ReferenceStyle:=xlR1C1)
End Sub

```

See the [documentation](./docs/documentation.md) for more details.

## Installation

Install the package with:

```powershell
python -m pip install matita
```

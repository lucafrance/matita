# Matita - Full Microsoft Office automation in Python ✏️

*Matita* is a Python wrapper for the [VBA Object library](https://learn.microsoft.com/en-us/office/vba/api/overview/).

```python
from matita.office import excel as xl

xl_app = xl.Application()
xl_app.visible = True

wkb = xl_app.workbooks.add()
wks = wkb.worksheets(1)
c = wks.cells(1,1)

c.value = "Hello World"
```

There are modules for Microsoft Access, Excel, Outlook, PowerPoint, Word.

See the [documentation](./docs/documentation.md) for more details.

## Installation

```powershell
python -m pip install matita
```

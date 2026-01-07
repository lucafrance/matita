from matita.office import excel as xl

xl_app = xl.Application().new()
xl_app.Visible = True

wkb = xl_app.Workbooks.Add()
wks = wkb.Worksheets(1)
c = wks.Cells(1,1)

c.Value = "Hello World"
print(c.Address(ReferenceStyle=xl.xlR1C1))

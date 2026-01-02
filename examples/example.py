from matita.office import excel

xl_app = excel.Application().new()
xl_app.Visible = True

wkb = xl_app.Workbooks.Add()
wks = wkb.Worksheets(1)
c = wks.Cells(1,1)

c.Value = "Hello World"
xlR1C1 = -4150
print(c.Address(None, None, xlR1C1))

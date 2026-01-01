from matita.office import excel

xl_app = excel.Application().new()
xl_app.Visible = True

wkb = xl_app.Workbooks.Add()
wks = wkb.Worksheets(1)
wks.Cells(1,1).Value = "Hello World"

# TODO adapt vba code

# Sub GenerateReport()
#     Const xlFormatCommas As Long = 2
#     Dim filePath As String
#     filePath = ThisWorkbook.Path & "/" & "tps00001__custom_19561911_linear_2_0.csv"
    
#     Dim dataWkb As Workbook
#     Set dataWkb = Workbooks.Open(filePath, ReadOnly:=True, Format:=xlFormatCommas)
    
#     Dim dataWks As Worksheet
#     Set dataWks = dataWkb.Worksheets(1)
    
#     Dim dataTbl As ListObject
#     Set dataTbl = dataWks.ListObjects.Add(Source:=dataWks.UsedRange, XlListObjectHasHeaders:=xlYes)
    
#     Dim countriesClm As ListColumn
#     Dim populationClm As ListColumn
#     Set countriesClm = dataTbl.ListColumns("Geopolitical entity (reporting)")
#     Set populationClm = dataTbl.ListColumns("OBS_VALUE")
    
#     Dim euCountries As Variant
#     euCountries = Application.WorksheetFunction.Sort( _
#                     Application.WorksheetFunction.Unique(countriesClm.DataBodyRange))
    
#     Dim startYear As Long, endYear As Long, numRows As Long
#     startYear = WorksheetFunction.Min(dataTbl.ListColumns("TIME_PERIOD").DataBodyRange)
#     endYear = WorksheetFunction.Max(dataTbl.ListColumns("TIME_PERIOD").DataBodyRange)
#     numRows = endYear - startYear + 1
    
#     Dim reportWkb As Workbook
#     Set reportWkb = Workbooks.Add
    
#     Dim firstWks As Worksheet
#     Set firstWks = reportWkb.Worksheets(1)
    
#     Dim i As Long
#     For i = LBound(euCountries) To UBound(euCountries)
#         Dim country As String
#         country = euCountries(i, 1)
        
#         Dim countryWks As Worksheet
#         Set countryWks = reportWkb.Worksheets.Add(After:=reportWkb.Worksheets(reportWkb.Worksheets.Count))
        
#         With countryWks
#             .Name = country
#             .Cells(2, 2).Value = "Population of " & country
#             .Rows(2).Style = "Heading 1"
#             .Columns(1).ColumnWidth = 3
            
#             .Cells(4, 2).Value = "Year"
#             .Cells(4, 3).Value = "Population"
#             .Cells(5, 2).Value = startYear
#             .Cells(5, 2).DataSeries Rowcol:=xlColumns, Type:=xlLinear, Step:=1, Stop:=endYear
            
#             dataTbl.DataBodyRange.AutoFilter field:=countriesClm.Index, Criteria1:=country
#             .Cells(5, 3).Resize(numRows).Value = populationClm.DataBodyRange.SpecialCells(xlCellTypeVisible).Value
            
#             Dim countryTbl As ListObject
#             Set countryTbl = .ListObjects.Add(Source:=.Cells(4, 2).CurrentRegion, XlListObjectHasHeaders:=xlYes)
#             countryTbl.ListColumns("Population").DataBodyRange.NumberFormat = "#,##0"
            
#             Dim shp As Shape
#             Set shp = .Shapes.AddChart2( _
#                 XlChartType:=xlLineMarkers, _
#                 Left:=.Cells(4, 5).Left, _
#                 Top:=.Cells(4, 5).Top)
            
#             Dim c As Chart
#             Set c = shp.Chart
#             c.HasTitle = False
#             c.FullSeriesCollection(1).Name = .Cells(4, 2).Address
#             c.FullSeriesCollection(1).Values = "'" & country & "'!" & countryTbl.ListColumns("Population").DataBodyRange.Address
#             c.FullSeriesCollection(1).XValues = "'" & country & "'!" & countryTbl.ListColumns("Year").DataBodyRange.Address
#         End With
        
#         dataWks.ShowAllData
#     Next i
    
#     dataWkb.Close False
    
#     Application.DisplayAlerts = False
#     firstWks.Delete
#     Application.DisplayAlerts = True
    
#     reportWkb.Worksheets(1).Activate
# End Sub

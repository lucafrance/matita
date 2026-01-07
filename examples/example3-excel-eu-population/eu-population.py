import os

from matita.office import excel as xl

def generate_report():
    xlFormatCommas = 2
    file_path = os.path.dirname(os.path.abspath(__file__)) + "/tps00001__custom_19561911_linear_2_0.csv"

    xl_app = xl.Application().new()
    xl_app.visible = True
    
    data_wkb = xl_app.Workbooks.Open(file_path, ReadOnly=True, Format=xlFormatCommas)
    data_wks = data_wkb.worksheets(1)
    data_tbl = data_wks.ListObjects.add(
        SourceType=xl.xlSrcRange,
        Source=data_wks.usedrange,
        XlListObjectHasHeaders=xl.xlYes,
    )
    countries_clm = data_tbl.ListColumns("Geopolitical entity (reporting)")
    population_clm = data_tbl.ListColumns("OBS_VALUE")
    eu_countries = sorted(set(countries_clm.databodyrange.value))
    eu_countries = [x[0] for x in eu_countries.copy()]
    print(eu_countries)

    data_wkb.close(False)

if __name__ == "__main__":
    generate_report()

# TODO adapt vba code

# Sub GenerateReport()

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

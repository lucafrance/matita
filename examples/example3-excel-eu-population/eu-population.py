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
    countries_clm = data_tbl.listcolumns("Geopolitical entity (reporting)")
    population_clm = data_tbl.listcolumns("OBS_VALUE")

    eu_countries = sorted(set(countries_clm.databodyrange.value))
    eu_countries = [x[0] for x in eu_countries.copy()]

    years = [x[0] for x in data_tbl.listcolumns("TIME_PERIOD").databodyrange.value]
    start_year = int(min(years))
    end_year = int(max(years))
    num_rows = end_year - start_year + 1

    report_wkb = xl_app.workbooks.add()
    first_wks = report_wkb.worksheets(1)
    for country in eu_countries:
        last_wks = report_wkb.worksheets(report_wkb.worksheets.count)
        country_wks = report_wkb.worksheets.add(After=last_wks)

        country_wks.name = country
        country_wks.cells(2, 2).value = f"Population of {country}"
        country_wks.rows.item(2).style = "Heading 1"
        country_wks.columns.item(1).ColumnWidth = 3

        country_wks.cells(4, 2).Value = "Year"
        country_wks.cells(4, 3).Value = "Population"
        country_wks.cells(5, 2).Value = start_year
        country_wks.cells(5, 2).DataSeries(Rowcol=xl.xlColumns, Type=xl.xlLinear, Step=1, Stop=end_year)

    first_wks.delete()

    data_wkb.close(False)

if __name__ == "__main__":
    generate_report()

# TODO adapt vba code

# Sub GenerateReport()
#     For i = LBound(euCountries) To UBound(euCountries)
        
#         With countryWks
            
#             .cells(4, 2).Value = "Year"
#             .cells(4, 3).Value = "Population"
#             .cells(5, 2).Value = start_year
#             .cells(5, 2).DataSeries Rowcol=xlColumns, Type=xlLinear, Step=1, Stop=end_year
            
#             dataTbl.DataBodyRange.AutoFilter field=countriesClm.Index, Criteria1=country
#             .cells(5, 3).Resize(numRows).Value = populationClm.DataBodyRange.SpecialCells(xlCellTypeVisible).Value
            
#             Dim countryTbl As ListObject
#             Set countryTbl = .ListObjects.Add(Source=.cells(4, 2).CurrentRegion, XlListObjectHasHeaders=xlYes)
#             countryTbl.ListColumns("Population").DataBodyRange.NumberFormat = "#,##0"
            
#             Dim shp As Shape
#             Set shp = .Shapes.AddChart2( _
#                 XlChartType=xlLineMarkers, _
#                 Left=.cells(4, 5).Left, _
#                 Top=.cells(4, 5).Top)
            
#             Dim c As Chart
#             Set c = shp.Chart
#             c.HasTitle = False
#             c.FullSeriesCollection(1).Name = .cells(4, 2).Address
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

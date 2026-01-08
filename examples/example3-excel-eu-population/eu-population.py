import os

from matita.office import excel as xl

def generate_report():
    xlFormatCommas = 2
    file_path = os.path.dirname(os.path.abspath(__file__)) + "/tps00001__custom_19561911_linear_2_0.csv"

    xl_app = xl.Application()
    xl_app.visible = True
    
    data_wkb = xl_app.Workbooks.Open(file_path, ReadOnly=True, Format=xlFormatCommas)
    data_wks = data_wkb.worksheets(1)
    data_tbl = data_wks.listobjects.add(
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

        # Add heading
        country_wks.name = country
        country_wks.cells(2, 2).value = f"Population of {country}"
        country_wks.rows.item(2).style = "Heading 1"
        country_wks.columns.item(1).ColumnWidth = 3

        # Add table
        country_wks.cells(4, 2).Value = "Year"
        country_wks.cells(4, 3).Value = "Population"
        country_wks.cells(5, 2).Value = start_year
        country_wks.cells(5, 2).DataSeries(Rowcol=xl.xlColumns, Type=xl.xlLinear, Date=xl.xlDay, Step=1, Stop=end_year)

        data_tbl.databodyrange.autofilter(Field=countries_clm.Index, Criteria1=country)
        country_wks.cells(5, 3).Resize(num_rows).Value = population_clm.databodyrange.specialcells(xl.xlCellTypeVisible).Value

        country_tbl = country_wks.listobjects.add(
            SourceType=xl.xlSrcRange,
            Source=country_wks.cells(4, 2).currentregion,
            XlListObjectHasHeaders=xl.xlYes
        )
        country_tbl.listcolumns("Population").databodyrange.numberformat = "#,##0"

        shp = country_wks.shapes.addchart2( 
            XlChartType=xl.xlLineMarkers,
            Left=country_wks.cells(4, 5).left,
            Top=country_wks.cells(4, 5).top,
        )
        c = shp.chart
        c.hastitle = False
        chart_series = c.fullseriescollection().item(1)
        chart_series.name = country_wks.cells(4, 2).address()
        chart_series.values  = f"'{country}'!{country_tbl.listcolumns("Population").databodyrange.address()}"
        chart_series.xvalues = f"'{country}'!{country_tbl.listcolumns("Year").databodyrange.address()}"

    first_wks.delete()
    data_wkb.close(False)
    report_wkb.worksheets(1).activate()

if __name__ == "__main__":
    generate_report()

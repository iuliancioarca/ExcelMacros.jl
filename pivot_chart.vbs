Option Explicit

PivotCharts
Sub PivotCharts()
	Dim xlApp 
	Dim xlBook 
	Dim pvtTable
	Dim pvtChart
	Const xlLineMarkers = 65
	Set xlApp = CreateObject("Excel.Application") 
	Set xlBook = xlApp.Workbooks.Open("C:\Projects\Excel_Report\pivot_chart.xlsx")
    xlBook.CheckCompatibility = False
	Const xlDatabase = 1
    Set pvtTable = xlBook.PivotCaches.Create(xlDatabase, "FirstSheet!R1C1:R101C4").CreatePivotTable("FirstSheet!R2C7", "PivotTable1")
	xlBook.Sheets("FirstSheet").PivotTables("PivotTable1").HasAutoFormat= False
    xlBook.Sheets("FirstSheet").Select
    xlBook.Sheets("FirstSheet").Cells(2,7).Select
	
	Set pvtChart = xlBook.Sheets("FirstSheet").Shapes.AddChart()
	xlBook.Sheets("FirstSheet").ChartObjects(1).Chart.ChartType = 65 'https://docs.microsoft.com/en-us/office/vba/api/excel.xlcharttype
	Call xlBook.Sheets("FirstSheet").ChartObjects(1).Chart.SetSourceData(xlBook.Sheets("FirstSheet").Range("$G$3:$K$10"))
	xlBook.Sheets("FirstSheet").ChartObjects(1).Chart.HasTitle = True 
	xlBook.Sheets("FirstSheet").ChartObjects(1).Chart.ChartTitle.Text = "My first title"
	
	Call xlBook.Sheets("FirstSheet").PivotTables("PivotTable1").AddDataField(xlBook.Sheets("FirstSheet").PivotTables( _
        "PivotTable1").PivotFields("OUT1"), "Avg of OUT1",-4106)
		' mind the function codes at:https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlconsolidationfunction?view=excel-pia
		
	xlBook.Save
	xlBook.Close
	xlApp.Quit 

  Set xlBook = Nothing 
  Set xlApp = Nothing 
End Sub
msgbox("Excel: Done")
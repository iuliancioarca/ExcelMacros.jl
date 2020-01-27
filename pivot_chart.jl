cd(raw"C:\Projects\Excel_Report")
using XLSX
using FileIO

function send2xls()
    XLSX.openxlsx("pivot_chart.xlsx", mode="w") do xf
        sheet = xf[1]
        XLSX.rename!(sheet, "FirstSheet")
        # write header
        sheet["A1"] = "IN1"
        sheet["B1"] = "IN2"
        sheet["C1"] = "OUT1"
        sheet["D1"] = "OUT2"
        # write some data
        IN1  = 1:100
        IN2  = range(1, step=5, length=100)
        OUT1 = randn(100)
        OUT2 = randn(100)
        sheet["A2:D101"] = hcat(IN1,IN2,OUT1,OUT2)
    end
end

#generate vbs
function gen_vbs()
    # generate the vbscript to create pivot chart based on data written above;
    # for the moment all params ar hardcoded: cell refferences, names
    open("pivot_chart.txt","w") do vbfile
    # define vars, constants etc
    write(vbfile,"Option Explicit\nPivotCharts\nSub PivotCharts()\nDim xlApp
        \nDim xlBook\nDim pvtTable\nDim pvtChart\nConst xlDatabase = 1\n")
    # open excel file
    write(vbfile,"""Set xlApp = CreateObject("Excel.Application")\n""")
    write(vbfile,"""Set xlBook = xlApp.Workbooks.Open("C:\\Projects\\Excel_Report\\pivot_chart.xlsx")\n""")
    write(vbfile,"xlBook.CheckCompatibility = False\n")
    # set a pivot table
    write(vbfile,"""Set pvtTable = xlBook.PivotCaches.Create(xlDatabase, "FirstSheet!R1C1:R101C4").CreatePivotTable("FirstSheet!R2C7", "PivotTable1")\n""")
    write(vbfile,"""xlBook.Sheets("FirstSheet").PivotTables("PivotTable1").HasAutoFormat= False\n""")
    write(vbfile,"""xlBook.Sheets("FirstSheet").Select\n""")
    write(vbfile,"""xlBook.Sheets("FirstSheet").Cells(2,7).Select\n""")
    # set a pivot chart
    write(vbfile,"""Set pvtChart = xlBook.Sheets("FirstSheet").Shapes.AddChart()\n""")
    write(vbfile,"""xlBook.Sheets("FirstSheet").ChartObjects(1).Chart.ChartType = 65\n""")
    write(vbfile,"""Call xlBook.Sheets("FirstSheet").ChartObjects(1).Chart.SetSourceData(xlBook.Sheets("FirstSheet").Range("\$G\$3:\$K\$10"))\n""")
    write(vbfile,"""xlBook.Sheets("FirstSheet").ChartObjects(1).Chart.HasTitle = True\n""")
    write(vbfile,"""xlBook.Sheets("FirstSheet").ChartObjects(1).Chart.ChartTitle.Text = "My first title"\n""")
    write(vbfile,"""Call xlBook.Sheets("FirstSheet").PivotTables("PivotTable1").AddDataField(xlBook.Sheets("FirstSheet").PivotTables( _\n "PivotTable1").PivotFields("OUT1"), "Avg of OUT1",-4106)\n""")
    # save and quit
    write(vbfile,"xlBook.Save\n")
    write(vbfile,"xlBook.Close\n")
    write(vbfile,"xlApp.Quit\n")
    write(vbfile,"Set xlBook = Nothing\n")
    write(vbfile,"Set xlApp = Nothing\n")
    write(vbfile,"End Sub\n")
    write(vbfile,"""msgbox("Excel: Done")\n""")
    end
    # not rename to .vbs
    mv("pivot_chart.txt", "pivot_chart.vbs",force = true)
end

### USAGE
# call send2xls and generate excel file with data inside
send2xls()
# generate vbs
gen_vbs()
# call vbs
mycommand = `WScript.exe "pivot_chart.vbs"`
wait(run(mycommand))

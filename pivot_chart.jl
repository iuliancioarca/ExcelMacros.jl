cd(raw"C:\Projects\Excel_Report")
using XLSX

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
# call send2xls and generate excel file with data inside
send2xls()
# call vbs
mycommand = `WScript.exe "pivot_chart.vbs"`
wait(run(mycommand))

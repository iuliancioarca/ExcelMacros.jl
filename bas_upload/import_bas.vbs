'Target Excel file to import BAS file to
Set Fso = WScript.CreateObject("Scripting.FileSystemObject")

Filepath = "C:\Users\IulianCioarca\Desktop\a4.xlsx"

Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open(filepath)

objExcel.Visible = True

objExcel.DisplayAlerts = False
'Imports BAS module, but using a filepath

objExcel.VBE.ActiveVBProject.VBComponents.Import "C:\Users\IulianCioarca\Desktop\Module2.bas"

objExcel.Run "Macro1"

objWorkbook.Save 
objWorkbook.Close 
objExcel.Quit 
 
'WScript.Echo "Finished." 
WScript.Quit
Attribute VB_Name = "Module2"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "=PRODUCT(R[-2]C:R[-1]C)"
    Range("A4").Select
End Sub

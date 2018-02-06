Attribute VB_Name = "Module7"
Sub CorrectDataType()
Attribute CorrectDataType.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CorrectDataType Macro
'

'
    Windows("310310310.xlsx").Activate
    ActiveWindow.SmallScroll Down:=-28
    Range("C10").Select
    Selection.NumberFormat = "h:mm;@"
    Selection.AutoFill Destination:=Range("C10:D10"), Type:=xlFillDefault
    Range("C10:D10").Select
    Selection.Copy
    Range("C11:D38").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveWindow.SmallScroll Down:=-12
    Range("E10").Select
    ActiveCell.FormulaR1C1 = "=MOD(RC[-2]-R[1]C[-2],1)"
    ActiveCell.FormulaR1C1 = "=MOD(R[1]C[-2]-RC[-2],1)"
    Range("E10").Select
    Selection.AutoFill Destination:=Range("E10:E38"), Type:=xlFillDefault
    Range("E10:E38").Select
    ActiveWindow.SmallScroll Down:=-4
    Range("E27").Select
    ActiveWindow.SmallScroll Down:=-8
    Range("E16").Select
    ActiveWindow.SmallScroll Down:=24
    Range("F42").Select
    ActiveWindow.SmallScroll Down:=-8
    Range("E27").Select
    ActiveWindow.SmallScroll Down:=-8
    Range("E21").Select
    ActiveWindow.SmallScroll Down:=-12
    Range("E10").Select
    ActiveCell.FormulaR1C1 = "=MOD(RC[-1]-RC[-2],1)"
    Range("E10").Select
    Selection.AutoFill Destination:=Range("E10:E38"), Type:=xlFillDefault
    Range("E10:E38").Select
    ActiveWindow.SmallScroll Down:=-16
    Sheets("2").Select
    Range("J8").Select
    Sheets("Summary").Select
End Sub

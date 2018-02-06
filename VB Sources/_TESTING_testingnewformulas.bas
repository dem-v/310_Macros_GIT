Attribute VB_Name = "Module12"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    ActiveWindow.SmallScroll Down:=-8
    Application.CommandBars("Research").Visible = False
    ActiveCell.FormulaR1C1 = "=RC[-1]/R[22]C[-1]"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]/R24C5"
    Range("F2").Select
    Selection.AutoFill Destination:=Range("F2:F23"), Type:=xlFillDefault
    Range("F2:F23").Select
    Selection.Style = "Percent"
    Selection.FormatConditions.AddTop10
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .TopBottom = xlTop10Top
        .Rank = 10
        .Percent = True
    End With
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub

Sub testestest()
i = 1
                        'lister = Application.ConvertFormula("'" + CStr(i) + "'!J4:J28", xlA1, xlR1C1, , Range("E" & CStr(i + 8)))
                        'arrayer = Application.ConvertFormula("'" + CStr(i) + "'!K4:K28", xlA1, xlR1C1, , Range("E" & CStr(i + 8)))
                        lister1 = "'" + CStr(i) + "'!J4:J28"
                        
                        
                        Range("AA" & CStr(i + 8)).FormulaArray = "=IFERROR(SMALL(IF(" + lister1 + "=""x"",ROW(" + lister1 + ")-MIN(ROW(" + lister1 + "))+1),2), -1)" '22
                        Range("AB" & CStr(i + 8)).FormulaArray = "=IFERROR(SMALL(IF(" + lister1 + "=""x"",ROW(" + lister1 + ")-MIN(ROW(" + lister1 + "))+1),3), -1)" '23
                        Range("AC" & CStr(i + 8)).FormulaArray = "=IFERROR(SMALL(IF(" + lister1 + "=""x"",ROW(" + lister1 + ")-MIN(ROW(" + lister1 + "))+1),4), -1)" '24
                        Range("AD" & CStr(i + 8)).FormulaArray = "=SMALL(IF(" + lister1 + "=""x"",ROW(" + lister1 + ")-MIN(ROW(" + lister1 + "))+1),1)" '25
                        Range("AE" & CStr(i + 8)).FormulaArray = "=SMALL(IF(" + lister1 + "=""x"",ROW(" + lister1 + ")-MIN(ROW(" + lister1 + "))+1),2)" '26
                        Range("AF" & CStr(i + 8)).FormulaArray = "=SMALL(IF(" + lister1 + "=""x"",ROW(" + lister1 + ")-MIN(ROW(" + lister1 + "))+1),3)" '27
End Sub

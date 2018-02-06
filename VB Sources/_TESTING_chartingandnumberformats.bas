Attribute VB_Name = "Module9"
Sub ChartExpansion()
Attribute ChartExpansion.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ChartExpansion Macro
'

'
    Sheets("TOTAL").Select
    ActiveSheet.ChartObjects("Chart 2").Activate
    ActiveWindow.SmallScroll Down:=4
    ActiveChart.SetSourceData Source:=Range("TOTAL!$D$2:$E$24")
    ActiveChart.SetSourceData Source:=Range("TOTAL!$D$2:$E$23")
End Sub
Sub hh()
Attribute hh.VB_ProcData.VB_Invoke_Func = " \n14"
'
' hh Macro
'

'
    ActiveChart.ChartArea.Select
    Range("F6").Select
    ActiveSheet.ChartObjects("PieCharty").Activate
    ActiveChart.SetSourceData Source:=Range("TOTAL!$D$2:$E$22")
    ActiveChart.SetSourceData Source:=Range("TOTAL!$D$3:$E$23")
End Sub
Sub one()
i = 2
TaxiService = 3

                With ActiveSheet.Range("F" & CStr(i + 8))
                    .Select
                    .Formula = "=IF(OR(D" & CStr(i + 8) & "<=0.25,D" & CStr(i + 8) & ">=0.9166),IF(ISTEXT(C" & CStr(i + 8) & "),0," & TaxiService & "),0)"
                    .NumberFormat = "0.00;@"
                End With

End Sub

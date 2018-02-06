Attribute VB_Name = "Module3"
Sub sfafa()
Attribute sfafa.VB_ProcData.VB_Invoke_Func = " \n14"
'
' sfafa Macro
'

'
    Range("F3:I3").Select
    ActiveCell.FormulaR1C1 = "666f"
    Range("F3:I3").Select
    ActiveCell.FormulaR1C1 = "666fds"
    Range("F4").Select
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    Sheets("MASTER TOTAL").Select
    With ActiveWorkbook.Sheets("MASTER TOTAL").Tab
        .Color = 12611584
        .TintAndShade = 0
    End With
    Sheets("MASTER TOTAL").Select
    With ActiveWorkbook.Sheets("MASTER TOTAL").Tab
        .Color = 5287936
        .TintAndShade = 0
    End With
    Sheets("MASTER TOTAL").Select
    With ActiveWorkbook.Sheets("MASTER TOTAL").Tab
        .Color = 65535
        .TintAndShade = 0
    End With
End Sub
Sub NamedRange()
Attribute NamedRange.VB_ProcData.VB_Invoke_Func = " \n14"
'
' NamedRange Macro
'

'
    With ActiveWorkbook.Names("Subject1")
        .Name = "Category"
        .RefersToR1C1 = "=MAIN!R78C15:R98C15"
        .Comment = ""
    End With
    Range("O76").Select
    ActiveCell.FormulaR1C1 = "O78:O98"
    Range("O77").Select
End Sub
Sub TotalUpdate()
Attribute TotalUpdate.VB_ProcData.VB_Invoke_Func = " \n14"
'
' TotalUpdate Macro
'

'
    Range("A2").Select
    ActiveCell.FormulaR1C1 = _
        "='1'!R[29]C[12]+'2'!R[29]C[12]+'3'!R[29]C[12]+'4'!R[29]C[12]+'5'!R[29]C[12]+'6'!R[29]C[12]+'7'!R[29]C[12]+'8'!R[29]C[12]+'9'!R[29]C[12]+'10'!R[29]C[12]+'11'!R[29]C[12]+'12'!R[29]C[12]+'13'!R[29]C[12]+'14'!R[29]C[12]+'15'!R[29]C[12]+'16'!R[29]C[12]+'17'!R[29]C[12]+'18'!R[29]C[12]+'19'!R[29]C[12]+'20'!R[29]C[12]+'21'!R[29]C[12]+'22'!R[29]C[12]+'23'!R[29]C[12]+'24'!" & _
        "R[29]C[12]+'25'!R[29]C[12]+'26'!R[29]C[12]+'27'!R[29]C[12]+'28'!R[29]C[12]+'29'!R[29]C[12]+'30'!R[29]C[12]+'31'!R[29]C[12]" & _
        ""
    Range("A3").Select
    ActiveSheet.ChartObjects("Chart 2").Activate
    Range("E2").Select
    ActiveCell.FormulaR1C1 = _
        "=SUMIF('1'!C[-2],RC[-1],'1'!C[8])+SUMIF('2'!C[-2],RC[-1],'2'!C[8])+SUMIF('3'!C[-2],RC[-1],'3'!C[8])+SUMIF('4'!C[-2],RC[-1],'4'!C[8])+SUMIF('5'!C[-2],RC[-1],'5'!C[8])+SUMIF('6'!C[-2],RC[-1],'6'!C[8])+SUMIF('7'!C[-2],RC[-1],'7'!C[8])+SUMIF('8'!C[-2],RC[-1],'8'!C[8])+SUMIF('9'!C[-2],RC[-1],'9'!C[8])+SUMIF('10'!C[-2],RC[-1],'10'!C[8])+SUMIF('11'!C[-2],RC[-1],'11'!" & _
        "C[8])+SUMIF('12'!C[-2],RC[-1],'12'!C[8])+SUMIF('13'!C[-2],RC[-1],'13'!C[8])+SUMIF('14'!C[-2],RC[-1],'14'!C[8])+SUMIF('15'!C[-2],RC[-1],'15'!C[8])+SUMIF('16'!C[-2],RC[-1],'16'!C[8])+SUMIF('17'!C[-2],RC[-1],'17'!C[8])+SUMIF('18'!C[-2],RC[-1],'18'!C[8])+SUMIF('19'!C[-2],RC[-1],'19'!C[8])+SUMIF('20'!C[-2],RC[-1],'20'!C[8])+SUMIF('21'!C[-2],RC[-1],'21'!C[8])+SUMIF('22'!C" & _
        "[-2],RC[-1],'22'!C[8])+SUMIF('23'!C[-2],RC[-1],'23'!C[8])+SUMIF('24'!C[-2],RC[-1],'24'!C[8])+SUMIF('25'!C[-2],RC[-1],'25'!C[8])+SUMIF('26'!C[-2],RC[-1],'26'!C[8])+SUMIF('27'!C[-2],RC[-1],'27'!C[8])+SUMIF('28'!C[-2],RC[-1],'28'!C[8])+SUMIF('29'!C[-2],RC[-1],'29'!C[8])+SUMIF('30'!C[-2],RC[-1],'30'!C[8])+ SUMIF('31'!C[-2],RC[-1],'31'!C[8])" & _
        ""
    Range("E2").Select
    Selection.AutoFill Destination:=Range("E2:E20"), Type:=xlFillDefault
    Range("E2:E20").Select
    Range("A3").Select
End Sub
Sub EditingCategories()
Attribute EditingCategories.VB_ProcData.VB_Invoke_Func = " \n14"
'
' EditingCategories Macro
'

'
    Range("D2").Select
    Application.Goto Reference:="Category"
    Selection.Copy
    Sheets("TOTAL").Select
    ActiveSheet.Paste
    Range("D23").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Total this month"
    Range("E23").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-21]C:R[-1]C)"
    Range("E2").Select
    Selection.Copy
    Range("E23").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
End Sub

Sub ttt()

    daovto = "=SUM("
     tohothwe = "=SUM("
    For j = 1 To 5
        daovto = daovto + "'" + CStr(j) + "'!" + "M33,"
        tohothwe = tohothwe + "'" + CStr(j) + "'!" + "M31,"
    Next j
    daovto = daovto + "0)"
    tohothwe = tohothwe + "0)"
    
    
End Sub
Sub Fill()
Attribute Fill.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Fill Macro
'

'
    Range("E2").Select
    Selection.Copy
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("E2:E22"), Type:=xlFillDefault
    Range("E2:E22").Select
End Sub

Attribute VB_Name = "Module10"
Sub UpdateSumSource()
Attribute UpdateSumSource.VB_ProcData.VB_Invoke_Func = " \n14"
'
' UpdateSumSource Macro
'

'
Range("A8:F8").WrapText = True

End Sub
Sub UrgentPatch()
Attribute UrgentPatch.VB_ProcData.VB_Invoke_Func = " \n14"
'
' UrgentPatch Macro
'

'
    
    Sheets("Categories").Visible = True
    Range("I12").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveCell.FormulaR1C1 = "12. HR California"
    Range("I13").FormulaR1C1 = "13. Insurance"
    Range("I14").FormulaR1C1 = "14. Inventory maintenance"
    Range("I15").FormulaR1C1 = "15. Management"
    Range("I16").FormulaR1C1 = "16. Marketing"
    Range("I17").FormulaR1C1 = "17. Meeting"
    Range("I18").FormulaR1C1 = "18. Odessa office"
    Range("I19").FormulaR1C1 = "19. Payroll"
    Range("I20").FormulaR1C1 = "20. Purchase technology"
    Range("I21").FormulaR1C1 = "21. Recruitment"
    Range("I22").FormulaR1C1 = "22. Reports"
    Range("I23").FormulaR1C1 = "23. Research"
    Range("I24").FormulaR1C1 = "24. Scanning"
    Range("I25").FormulaR1C1 = "24.1. Filing"
    Range("I26").FormulaR1C1 = "25. Scheduling"
    Range("I27").FormulaR1C1 = "26. Software"
    Range("I28").FormulaR1C1 = "27. Technical maintenance"
    Range("I29").FormulaR1C1 = "28. Technical support"
    Range("I30").FormulaR1C1 = "29. Training"
    
    Range("I69").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveCell.FormulaR1C1 = "18.3. HR"
    
    ActiveWorkbook.Names("CA").RefersTo = "=Categories!I1:I30"
    ActiveWorkbook.Names("TX").RefersTo = "=Categories!I33:I71"
    
    Sheets("Summary").Select

    Sheets("Categories").Visible = False
    
    Sheets("Summary").Select
    Range("A8:F8").WrapText = True
    Range("F9").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-2]<=0.25,RC[-2]>=0.9166),IF(ISTEXT(RC[-3]),0,3),0)"
    Selection.Copy
    Range("F10:F38").Select
    ActiveSheet.Paste
        
    For i = 1 To 30
        Sheets(CStr(i)).Select
        Range("L15").Select
        Selection.Copy
        Range("L4:L28").Select
        ActiveSheet.Paste
        
        Range("B4").Select
    Next i
    
    Sheets("TOTAL").Select
    Regionel = "TX"
    
    If Range("D2").Text = "1. 3000 System" Then Regionel = "CA"

Range(Regionel).Copy
Range("D2").Select
ActiveSheet.Paste

CatLen = 2 + Range(Regionel).Rows.Count
'Add total count
Range("D" + CStr(CatLen)).Select
ActiveCell.FormulaR1C1 = "Total this month"
'Fill formulas
Range("E2").Select
pss = "E2:E" + CStr(CatLen - 1)
Selection.AutoFill Destination:=Range(pss), Type:=xlFillDefault
'Add sum formula
Range("E" + CStr(CatLen)).Select
ActiveCell.Formula = "=SUM(" + pss + ")"
Range("E2").Copy
Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False

'update graph
ActiveSheet.ChartObjects("PieCharty").Activate
ActiveChart.SetSourceData Source:=Range("TOTAL!$D$2:$E$" + CStr(CatLen - 1))

Range("A3").Formula = "=(SUM(E2:E" + CStr(CatLen - 1) + "))*24"

Range("A1").Select

Sheets("1").Select

End Sub

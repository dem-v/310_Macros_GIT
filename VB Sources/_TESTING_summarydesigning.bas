Attribute VB_Name = "Module6"
Sub SummaryGen()
Attribute SummaryGen.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SummaryGen Macro
'

'
    Sheets("BatchCreate").Select
    Sheets.Add
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "Summary"
    Windows("Book1").Activate
    ActiveWindow.SmallScroll Down:=-12
    Windows("test 2.xlsm [Last saved by user]").Activate
    Range("C7:E7").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Time sheet October 2017"
    Range("B8").Select
    ActiveCell.FormulaR1C1 = "Day of month"
    Range("C8").Select
    ActiveCell.FormulaR1C1 = "Start*"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "End*"
    Range("E8").Select
    ActiveCell.FormulaR1C1 = "Total hours"
    Range("F8").Select
    ActiveCell.FormulaR1C1 = "Taxi service"
    Range("B9").Select
    ActiveCell.FormulaR1C1 = "1"
    Windows("310_Demi Valantsevich_Oct.xlsx [Last saved by user]").Activate
    ActiveWindow.SmallScroll Down:=28
    Windows("test 2.xlsm [Last saved by user]").Activate
    Range("B7:F39").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Font
        .Name = "Times New Roman"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Range("B8:F8").Select
    Selection.Font.Bold = True
    ActiveWindow.SmallScroll Down:=24
    Range("B42:E45").Select
    Selection.Merge True
    Range("B42:E42").Select
    ActiveCell.FormulaR1C1 = "Total Worked Hours"
    With ActiveCell.Characters(Start:=1, Length:=13).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With ActiveCell.Characters(Start:=14, Length:=5).Font
        .Name = "Calibri"
        .FontStyle = "Bold"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Range("B42:F45").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    Range("B7:F39").Select
    Range("F39").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    Range("F42").Select
    Windows("310_Demi Valantsevich_Oct.xlsx [Last saved by user]").Activate
    Range("F43").Select
    Windows("test 2.xlsm [Last saved by user]").Activate
    ActiveCell.FormulaR1C1 = "=SUM(R[-33]C[-1]:R[-3]C[-1])"
    Range("F43").Select
    ActiveCell.FormulaR1C1 = "=COUNT(R[-34]C[-1]:R[-4]C[-1])"
    Range("F44").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C"
    Range("F45").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-36]C:R[-6]C)"
    Range("B42:F45").Select
    With Selection.Font
        .Name = "Times New Roman"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("B44:E44").Select
End Sub

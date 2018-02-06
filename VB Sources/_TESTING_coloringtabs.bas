Attribute VB_Name = "Module2"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet2").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("MASTER").Select
    ActiveWindow.SmallScroll Down:=-24
    Sheets("MASTER").Select
    Sheets("MASTER").Copy After:=Sheets(4)
    Sheets("MASTER (2)").Select
    Sheets("MASTER (2)").Name = "1"
    Sheets("1").Select
    With ActiveWorkbook.Sheets("1").Tab
        .Color = 5287936
        .TintAndShade = 0
    End With
    Range("C17").Select
    Sheets("1").Select
    Range("C12").Select
    Sheets("MASTER TOTAL").Select
    Sheets("MASTER TOTAL").Copy After:=Sheets(5)
    Sheets("MASTER TOTAL (2)").Select
    Sheets("MASTER TOTAL (2)").Name = "2"
    Range("C18").Select
    Sheets("2").Select
    With ActiveWorkbook.Sheets("2").Tab
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = 0
    End With
    Range("C22").Select
    Sheets("2").Select
    With ActiveWorkbook.Sheets("2").Tab
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.499984740745262
    End With
    Range("C16").Select
    Sheets("2").Select
    With ActiveWorkbook.Sheets("2").Tab
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.249977111117893
    End With
    Range("C19").Select
    Sheets("2").Select
End Sub

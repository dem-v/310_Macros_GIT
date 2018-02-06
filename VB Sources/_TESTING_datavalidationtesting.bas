Attribute VB_Name = "Module4"
Sub ChangingCats()
Attribute ChangingCats.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ChangingCats Macro
'

'
    Range("H5").Select
    Sheets("MASTER").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=CA"
        .IgnoreBlank = False
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Selection.Copy
    Range("C5:C28").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=-4
    Range("C4").Select
    Sheets("MASTER TOTAL").Select
    Range("C4").Select
    ActiveWindow.SmallScroll Down:=16
    Range("C4:C28").Select
    ActiveSheet.Paste
    Range("C4").Select
End Sub
Sub ChangingCats2()
Attribute ChangingCats2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ChangingCats2 Macro
'

'
    Sheets("TOTAL").Select
    Range("D2").Select
    Application.Goto Reference:="CA"
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("TOTAL").Select
    Range("D2").Select
    ActiveSheet.Paste
    Range("D23").Select
End Sub
Sub Visibility()
Attribute Visibility.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Visibility Macro
'

'
    Sheets("TOTAL").Select
    Sheets("MASTER").Visible = True
    Sheets("MASTER").Select
    Sheets("MASTER TOTAL").Visible = True
    Sheets("MASTER TOTAL").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("MASTER").Select
    ActiveWindow.SelectedSheets.Visible = False
End Sub
Sub hidingrows()
Attribute hidingrows.VB_ProcData.VB_Invoke_Func = " \n14"
'
' hidingrows Macro
'

'
    Rows("76:137").Select
    ActiveWindow.SmallScroll Down:=-20
    Selection.EntireRow.Hidden = True
    Rows("75:138").Select
    Selection.EntireRow.Hidden = False
End Sub

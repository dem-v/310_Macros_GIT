Attribute VB_Name = "Module5"
Sub NewNamedRangeManuallyDefined()
Attribute NewNamedRangeManuallyDefined.VB_ProcData.VB_Invoke_Func = " \n14"
'
' NewNamedRangeManuallyDefined Macro
'

'
    ActiveWorkbook.Names.Add Name:="test", RefersToR1C1:= _
        "=alfa,beta,theta,gamma"
    ActiveWorkbook.Names("test").Comment = ""
    With ActiveWorkbook.Names("test")
        .Name = "test"
        .RefersToR1C1 = "=alfa,beta,theta,gamma"
        .Comment = ""
    End With
    Range("C4").Select
    ActiveWorkbook.Names("test").RefersToR1C1 = _
        "=""{""""alfa"""";""""beta"""";""""theta"""";""""gamma""""}"""
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="""alpha"";""beta"";""gamma"""
        .IgnoreBlank = False
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="alpha,beta,gammma"
        .IgnoreBlank = False
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Sub
Sub MoveToLast()
Attribute MoveToLast.VB_ProcData.VB_Invoke_Func = " \n14"
'
' MoveToLast Macro
'

'
    Sheets("Categories").Select
    Sheets("Categories").Move After:=Sheets(5)
End Sub

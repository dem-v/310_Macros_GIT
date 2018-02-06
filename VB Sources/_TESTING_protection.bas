Attribute VB_Name = "Module8"
Sub Protection()
Attribute Protection.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Protection Macro
'

'
    Windows("310 - testingLine5 - Nov_2017.xls").Activate
    Sheets("TOTAL").Select
    ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
        False
    Sheets("Summary").Select
    ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
        False
    Sheets("1").Select
    ActiveWindow.SmallScroll Down:=0
    Range("B4").Select
    ActiveSheet.Protection.AllowEditRanges.Add Title:="Range2", Range:=Range( _
        "B4:K28")
    ActiveSheet.Protection.AllowEditRanges(1).Delete
    ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
        False
    ActiveWindow.SmallScroll Down:=-40
End Sub

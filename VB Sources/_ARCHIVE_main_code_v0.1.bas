Attribute VB_Name = "main_restored1"
Sub DeleteAllDailyPages1()
For i = 1 To 31
    Sheets(CStr(i)).Select
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets(CStr(i)).Delete
    Application.DisplayAlerts = True
Next i
End Sub

Sub Update3101()

End Sub
Sub Create3101(Optional InputName As String, Optional DayOfWeek As Integer, Optional MaxDaysM As Integer, _
                Optional MonthNum As Integer, Optional YearNum As Integer, Optional Region As String)

'Made by Demi Valantsevich(c) deistyny@gmail.com. Free to use under GPL/GNU license.

'Here we focus on getting data from MAIN
Sheets("MAIN").Select
Range("F3:I3").Select
Name = ActiveCell.Text

'InvisibilityMode
Application.ScreenUpdating = False

'This is the color table
Dim Colors(1 To 5) As Integer
Colors(1) = 19
Colors(2) = 17
Colors(3) = 20
Colors(4) = 31
Colors(5) = 53

ColN = 0

Rows("76:137").Select
Selection.EntireRow.Hidden = True
Selection.EntireRow.Hidden = False

'We are getting values
WeekDays = Range("H77").Value
WeekDays = WeekDays - 1
DayCnt = 0
MaxDays = Range("F78").Value
MonthNum = Range("D77").Value
YearNum = Range("F5").Value

'TX or CA
Regionel = Range("H5").Value

Sheets("MASTER").Visible = False
Sheets("MASTER TOTAL").Visible = False

Sheets("MASTER").Visible = True
Sheets("MASTER TOTAL").Visible = True
Sheets("MASTER").Select
Range("C4").Select
With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=" + Regionel
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
Range("B4").Select
Sheets("MASTER TOTAL").Select
Range("C4:C28").Select
ActiveSheet.Paste
Range("B4").Select

Application.CutCopyMode = False
    
'Main Loop
For i = 1 To MaxDays

'Some Counting
WeekDays = WeekDays + 1
If WeekDays > 7 Then WeekDays = 1
DayCnt = DayCnt + 1

'What do we do for Sundays!
If WeekDays = 7 Then
    ColN = ColN + 1
    'Copy Master Sheet and Rename IT
    Sheets("MASTER TOTAL").Copy Before:=Sheets("Categories")
    Sheets("MASTER TOTAL (2)").Name = CStr(DayCnt)
    Sheets(CStr(DayCnt)).Select
    'Put all necessary information
    Range("Q1").Value = Name
    Range("B2").Value = CStr(MonthNum) + "/" + CStr(DayCnt) + "/" + CStr(YearNum)
    'Counts for coloring
    st = 1
    If DayCnt > 7 Then
        st = DayCnt - 6
    End If
    'Coloring week
    For j = st To DayCnt - 2
        ActiveWorkbook.Sheets(CStr(j)).Tab.ColorIndex = Colors(ColN)
    Next j
    'Workaround for bug, about month starting from Sat or Sun
    If DayCnt < 3 Then
        ColN = ColN - 1
    End If
    
    'Fixing Daily Overtime Total
     daovto = "=SUM("
     tohothwe = "=SUM("
    For j = st To DayCnt
        daovto = daovto + "'" + CStr(j) + "'!" + "M33,"
        tohothwe = tohothwe + "'" + CStr(j) + "'!" + "M31,"
    Next j
    daovto = daovto + "0)"
    tohothwe = tohothwe + "0)"
    
    With Range("M35")
    .Select
    .Formula = daovto
    End With
    With Range("J29")
    .Select
    .Formula = tohothwe
    
    End With
    
'Any other Day
Else
    'Copy Master
    Sheets("MASTER").Copy Before:=Sheets("Categories")
    Sheets("MASTER (2)").Name = CStr(DayCnt)
    Sheets(CStr(DayCnt)).Select
    'Put in necessary fields
    Range("Q1").Value = Name
    Range("B2").Value = CStr(MonthNum) + "/" + CStr(DayCnt) + "/" + CStr(YearNum)

End If
Next i

'Post-Loop processing

'Calculate final week
A = DayCnt - WeekDays + 1
If WeekDays = 6 Then
B = DayCnt - 1
Else
If WeekDays = 7 Then
B = DayCnt - 2
Else
B = DayCnt
End If
End If

'Color it
For i = A To B
    ActiveWorkbook.Sheets(CStr(i)).Tab.ColorIndex = Colors(5)
        Next i

'Updating TOTAL
Sheets("TOTAL").Select

'updating formula of total
formmm = "=0"
grap = "="
For i = 1 To DayCnt
    formmm = formmm + "+'" + CStr(i) + "'!M31"
    grap = grap + "SUMIF('" + CStr(i) + "'!C:C,D2,'" + CStr(i) + "'!M:M)+"
Next i
Range("A2").Select
ActiveCell.Formula = formmm
grap = grap + "0"

'updating formula of graphs
Range("E2").Select
ActiveCell.Formula = grap

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

'Clean-up mess in case
Range("D" + CStr(CatLen + 1) + ":E200").Clear

'Clean-up!
Sheets("MASTER").Visible = False
Sheets("MASTER TOTAL").Visible = False
Sheets("Main").Select
Rows("76:137").Select
Selection.EntireRow.Hidden = True
Range("A1").Select

'InvisibilityMode off
Application.ScreenUpdating = True

'Select Main
Sheets("1").Select
Range("B4").Select

End Sub

Sub SingleCreate1()
'Get Current Directory and FileName
masterpath = ActiveWorkbook.Path & "\"
mastername = ActiveWorkbook.Name
'Create New Excel File
Workbooks.Add

sheetToDelete = ActiveSheet.Name
childname = Application.GetSaveAsFilename(InitialFileName:=InitialName, Title:="Save New Document", FileFilter:="Excel Files (*.xls), *.xls")
'Saving new workbook
'ActiveWorkbook.SaveAs FileFormat = 51
If childname <> False Then
ActiveWorkbook.SaveAs Filename:=childname
End If
'Get New FileName
childname = ActiveWorkbook.Name

'Copy Sheets
Workbooks(mastername).Activate
Sheets("MASTER").Visible = True
Sheets("MASTER TOTAL").Visible = True
Sheets.Copy Before:=Workbooks(childname).Sheets(sheetToDelete)
'Prepare workbook
Workbooks(childname).Activate
Sheets(sheetToDelete).Delete
Sheets("MAIN").Select

'Generating
Create310

Sheets("MASTER").Visible = True
Sheets("MASTER TOTAL").Visible = True
Sheets(Array("MAIN", "MASTER", "MASTER TOTAL", "BatchCreate")).Delete

Sheets("Categories").Visible = False

ActiveWorkbook.Save

End Sub
Sub BatchCreate1()



Sheets("BatchCreate").Visible = False


Sheets("BatchCreate").Visible = True
End Sub


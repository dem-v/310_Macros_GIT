Attribute VB_Name = "main_initial"

Public InName As String
Public Response As String
Public InToChange As String

Public InExistingFileAddress As String
Public RenamingResponse As String

Sub DeleteAllDailyPages()
    For i = 1 To 31
        Sheets(CStr(i)).Select
        On Error Resume Next
        Application.DisplayAlerts = False
        Sheets(CStr(i)).Delete
        Application.DisplayAlerts = True
    Next i
End Sub

Sub Update310()

End Sub
Function Create310(Optional InputName As String, Optional DayOfWeek As Integer, Optional MaxDaysM As Integer, _
                Optional MonthNumM As Integer, Optional YearNumM As Integer, Optional Region As String, Optional Taxi As Integer)

'Made by Demi Valantsevich(c) deistyny@gmail.com. Free to use under GPL/GNU license.

'Here we focus on getting data from MAIN
Sheets("MAIN").Select
Range("F3:I3").Select
If Not InputName = "" Then
    Name = InputName
    Else
    Name = ActiveCell.Text
End If

If Taxi > 0 Then
    TAXIRATE = Taxi
Else
    TAXIRATE = Range("I8").Value
End If
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
If Not DayOfWeek = 0 Then
    WeekDays = DayOfWeek
    Else
    WeekDays = Range("H77").Value
End If

WeekDays = WeekDays - 1
DayCnt = 0

If Not MaxDaysM = 0 Then
    MaxDays = MaxDaysM
    Else
    MaxDays = Range("F78").Value
End If

If Not MonthNumM = 0 Then
    MonthNum = MonthNumM
    Else
    MonthNum = Range("D77").Value
End If

If Not YearNumM = 0 Then
    YearNum = YearNumM
    Else
    YearNum = Range("F5").Value
End If



'TX or CA
If Not Region = "" Then
    Regionel = Region
    Else
    Regionel = Range("H5").Value
End If


Sheets("MASTER").Visible = False
Sheets("MASTER TOTAL").Visible = False

Sheets("MASTER").Visible = True
Sheets("MASTER TOTAL").Visible = True
Sheets("MASTER").Select

cc = Names(Regionel).RefersTo
If InStr(cc, "]") > 0 Then Names(Regionel).RefersTo = "=" & Right(cc, Len(cc) - InStr(cc, "]"))

Range("C4").Select
With Sheets("MASTER").Cells(4, 3).Validation
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
    Sheets("MASTER TOTAL").Copy After:=Sheets(ActiveWorkbook.Sheets.Count)
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
    
    'Fixing Daily Overtime Total and Total Weekly Hours
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
    Sheets("MASTER").Copy After:=Sheets(ActiveWorkbook.Sheets.Count)
    Sheets("MASTER (2)").Name = CStr(DayCnt)
    Sheets(CStr(DayCnt)).Select
    'Put in necessary fields
    Range("Q1").Value = Name
    Range("B2").Value = CStr(MonthNum) + "/" + CStr(DayCnt) + "/" + CStr(YearNum)
        
End If
Range("B4").Select
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

''''''''
'update graph
'ActiveSheet.ChartObjects("PieCharty").Activate
'ActiveChart.SetSourceData Source:=Range("TOTAL!$D$2:$E$" + CStr(CatLen - 1))
'
'Range("A3").Formula = "=(SUM(E2:E" + CStr(CatLen - 1) + "))*24"
''''''''

'update numeric representation

Range("F2").Formula = "=IFERROR(E2/$E$" + CStr(CatLen) + ","""")"
Range("F2").Select
pss1 = "F2:F" + CStr(CatLen - 1)
Selection.AutoFill Destination:=Range(pss1), Type:=xlFillDefault
Range("F2").Copy
Range("F3:F" + CStr(CatLen - 1)).Select
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

    Sheets("Categories").Select
    Sheets("Categories").Move After:=Sheets(Sheets.Count)
    Sheets("Categories").Visible = False

t = SummaryGenerate(MonthName(MonthNum), Val(YearNum), Val(MaxDays), Val(TAXIRATE))

'Protection
Sheets("TOTAL").Select
ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False
Sheets("Summary").Select
ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False

'InvisibilityMode off
Application.ScreenUpdating = True

'Select Main
Sheets("1").Select
Range("B4").Select

End Function

Function SingleCreate(Optional InitialName As String)
'Get Current Directory and FileName
masterpath = ActiveWorkbook.Path & "\"
mastername = ActiveWorkbook.Name
'Create New Excel File
Workbooks.Add

sheetToDelete = ActiveSheet.Name

childname = Application.GetSaveAsFilename(InitialFileName:=InitialName, Title:="Save New Document", FileFilter:="Excel Files (*.xlsx), *.xlsx")
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

ActiveWorkbook.Save

End Function

Function SingleCreateOverride_AndClose(Optional InitialName As String, Optional InputName As String, Optional DayOfWeek As Integer, Optional MaxDaysM As Integer, _
                          Optional MonthNumM As Integer, Optional YearNumM As Integer, Optional Region As String, Optional TaxiService As Integer)

'Get Current Directory and FileName
masterpath = ActiveWorkbook.Path & "\"
mastername = ActiveWorkbook.Name
'Create New Excel File
Workbooks.Add

sheetToDelete = ActiveSheet.Name
If Not InitialName = "" Then
    childname = InitialName
Else
childname = Application.GetSaveAsFilename(InitialFileName:=InitialName, Title:="Save New Document", FileFilter:="Excel Files (*.xlsx), *.xlsx")
End If
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
ResCode = Create310(InputName, DayOfWeek, MaxDaysM, MonthNumM, YearNumM, Region, TaxiService)

Sheets("MASTER").Visible = True
Sheets("MASTER TOTAL").Visible = True
Sheets(Array("MAIN", "MASTER", "MASTER TOTAL", "BatchCreate")).Delete

ActiveWorkbook.Save
ActiveWorkbook.Close

End Function

Function SummaryGenerate(MonStr As String, Year As Integer, MaxDays As Integer, Optional TaxiService As Integer)
    
    Sheets.Add Before:=Worksheets(1)
    ActiveSheet.Name = "Summary"
    
    'put b7
    
    With Range("C7:G7")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
        .FormulaR1C1 = "Time sheet " & MonStr & " " & CStr(Year)
    End With
    
    Range("B8").FormulaR1C1 = "Day of month"
    Range("C8").FormulaR1C1 = "Start*"
    Range("D8").FormulaR1C1 = "End*"
    Range("E8").FormulaR1C1 = "Break Time"
    Range("F8").FormulaR1C1 = "Total hours"
    Range("G8").FormulaR1C1 = "Taxi service"
    
    Range("A8:G8").WrapText = True
    
    Columns("AA:AF").Hidden = True
    
    For i = 1 To MaxDays
        Sheets(CStr(i)).Select
        StartTime = Range("K4").Value
        
        DayOfWeek = Weekday(CStr(i) & "-" & MonStr & "-" & CStr(Year), vbMonday)
        
        With Range("J4:J35")
            Set Rng = .Find(What:="x", _
                            After:=.Cells(.Cells.Count), _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
        End With
        
        If Not Rng Is Nothing Then
            EndTime = Range("K" & CStr(Rng.Row)).Value
        
            Sheets("Summary").Select
            Range("B" & CStr(i + 8)).Value = i
            With Range("C" & CStr(i + 8))
                .Formula = "=IF(AND(OR(WEEKDAY('" & i & "'!B2,11)=6,WEEKDAY('" & i & "'!B2,11)=7),ISBLANK('" & i & "'!K4)),""Weekend"",'" & i & "'!K4)"
                .NumberFormat = "hh:mm;@"
            End With
            ''' UPD 11/27/2017 Formula with attention to LAST occurence of "x"
            With Range("D" & CStr(i + 8))
                .Formula = "=IF(ISTEXT(C" & CStr(i + 8) & "),"""",IFERROR(LOOKUP(2,1/('" & i & "'!J:J=""x""),'" & i & "'!K:K),0))"
                .NumberFormat = "hh:mm;@"
            End With
            ''' UPD 1/31/2018 Formula added to count break times
            With Range("E" & CStr(i + 8))
                lister = Application.ConvertFormula("'" + CStr(i) + "'!J4:J28", xlA1, xlR1C1, , Range("E" & CStr(i + 8)))
                arrayer = Application.ConvertFormula("'" + CStr(i) + "'!K4:K28", xlA1, xlR1C1, , Range("E" & CStr(i + 8)))
                lister1 = "'" + CStr(i) + "'!J4:J28"
                
                Range("AA" & CStr(i + 8)).FormulaArray = "=IFERROR(SMALL(IF(" + lister1 + "=""x"",ROW(" + lister1 + ")-MIN(ROW(" + lister1 + "))+1),2), -1)" '22
                Range("AB" & CStr(i + 8)).FormulaArray = "=IFERROR(SMALL(IF(" + lister1 + "=""x"",ROW(" + lister1 + ")-MIN(ROW(" + lister1 + "))+1),3), -1)" '23
                Range("AC" & CStr(i + 8)).FormulaArray = "=IFERROR(SMALL(IF(" + lister1 + "=""x"",ROW(" + lister1 + ")-MIN(ROW(" + lister1 + "))+1),4), -1)" '24
                Range("AD" & CStr(i + 8)).FormulaArray = "=SMALL(IF(" + lister1 + "=""x"",ROW(" + lister1 + ")-MIN(ROW(" + lister1 + "))+1),1)" '25
                Range("AE" & CStr(i + 8)).FormulaArray = "=SMALL(IF(" + lister1 + "=""x"",ROW(" + lister1 + ")-MIN(ROW(" + lister1 + "))+1),2)" '26
                Range("AF" & CStr(i + 8)).FormulaArray = "=SMALL(IF(" + lister1 + "=""x"",ROW(" + lister1 + ")-MIN(ROW(" + lister1 + "))+1),3)" '27
                
                .FormulaR1C1 = _
        "=SUM(IF(RC[22] <> -1,IF(RC[23] <> -1, IF(RC[24] <> -1,MOD(INDEX(" + arrayer + ",RC[27]+1)-INDEX(" + arrayer + ",RC[27]),1) + MOD(INDEX(" + arrayer + ",RC[26]+1)-INDEX(" + arrayer + "," & _
        "RC[26]),1) + MOD(INDEX(" + arrayer + ",RC[25]+1)-INDEX(" + arrayer + ",RC[25]),1),MOD(INDEX(" + arrayer + ",RC[26]+1)-INDEX(" + arrayer + ",RC[26]),1) + MOD(INDEX(" + arrayer + ",RC[25]+1)" & _
        "-INDEX(" + arrayer + ",RC[25]),1)),MOD(INDEX(" + arrayer + ",RC[25]+1)-INDEX(" + arrayer + ",RC[25]),1)),0))"
                '.FormulaArray = .FormulaR1C1
                .NumberFormat = "hh:mm;@"
                
            End With
            With Range("F" & CStr(i + 8))
                .FormulaR1C1 = "=MOD(RC[-2]-RC[-3],1)-RC[-1]"
                .NumberFormat = "hh:mm;@"
            End With
            With Range("G" & CStr(i + 8))
                .Value = "=IF(OR(D" & CStr(i + 8) & "<=0.25,D" & CStr(i + 8) & ">=0.9583),IF(ISTEXT(C" & CStr(i + 8) & "),0," & TaxiService & "),0)"
                .NumberFormat = "0.00;@"
            End With
        Else
            Sheets("Summary").Select
            Range("B" & CStr(i + 8)).Value = i
            If DayOfWeek = 6 Or DayOfWeek = 7 Then
                With Range("C" & CStr(i + 8))
                    .Formula = "=IF(AND(OR(WEEKDAY('" & i & "'!B2,11)=6,WEEKDAY('" & i & "'!B2,11)=7),ISBLANK('" & i & "'!K4)),""Weekend"",'" & i & "'!K4)"
                    .NumberFormat = "hh:mm;@"
                End With
                ''' UPD 11/27/2017 Formula with attention to LAST occurence of "x"
                With Range("D" & CStr(i + 8))
                    .Formula = "=IF(ISTEXT(C" & CStr(i + 8) & "),"""",IFERROR(LOOKUP(2,1/('" & i & "'!J:J=""x""),'" & i & "'!K:K),0))"
                    .NumberFormat = "hh:mm;@"
                End With
                ''' UPD 1/31/2018 Formula added to count break times
                With Range("E" & CStr(i + 8))
                    lister = Application.ConvertFormula("'" + CStr(i) + "'!J4:J28", xlA1, xlR1C1, , Range("E" & CStr(i + 8)))
                    arrayer = Application.ConvertFormula("'" + CStr(i) + "'!K4:K28", xlA1, xlR1C1, , Range("E" & CStr(i + 8)))
                    lister1 = "'" + CStr(i) + "'!J4:J28"
                    
                    Range("AA" & CStr(i + 8)).FormulaArray = "=IFERROR(SMALL(IF(" + lister1 + "=""x"",ROW(" + lister1 + ")-MIN(ROW(" + lister1 + "))+1),2), -1)" '22
                    Range("AB" & CStr(i + 8)).FormulaArray = "=IFERROR(SMALL(IF(" + lister1 + "=""x"",ROW(" + lister1 + ")-MIN(ROW(" + lister1 + "))+1),3), -1)" '23
                    Range("AC" & CStr(i + 8)).FormulaArray = "=IFERROR(SMALL(IF(" + lister1 + "=""x"",ROW(" + lister1 + ")-MIN(ROW(" + lister1 + "))+1),4), -1)" '24
                    Range("AD" & CStr(i + 8)).FormulaArray = "=SMALL(IF(" + lister1 + "=""x"",ROW(" + lister1 + ")-MIN(ROW(" + lister1 + "))+1),1)" '25
                    Range("AE" & CStr(i + 8)).FormulaArray = "=SMALL(IF(" + lister1 + "=""x"",ROW(" + lister1 + ")-MIN(ROW(" + lister1 + "))+1),2)" '26
                    Range("AF" & CStr(i + 8)).FormulaArray = "=SMALL(IF(" + lister1 + "=""x"",ROW(" + lister1 + ")-MIN(ROW(" + lister1 + "))+1),3)" '27
                    
                    .FormulaR1C1 = _
        "=SUM(IF(RC[22] <> -1,IF(RC[23] <> -1, IF(RC[24] <> -1,MOD(INDEX(" + arrayer + ",RC[27]+1)-INDEX(" + arrayer + ",RC[27]),1) + MOD(INDEX(" + arrayer + ",RC[26]+1)-INDEX(" + arrayer + "," & _
        "RC[26]),1) + MOD(INDEX(" + arrayer + ",RC[25]+1)-INDEX(" + arrayer + ",RC[25]),1),MOD(INDEX(" + arrayer + ",RC[26]+1)-INDEX(" + arrayer + ",RC[26]),1) + MOD(INDEX(" + arrayer + ",RC[25]+1)" & _
        "-INDEX(" + arrayer + ",RC[25]),1)),MOD(INDEX(" + arrayer + ",RC[25]+1)-INDEX(" + arrayer + ",RC[25]),1)),0))"
                    '.FormulaArray = .FormulaR1C1
                    .NumberFormat = "hh:mm;@"
                    
                End With
                With Range("F" & CStr(i + 8))
                    .Formula = "=IF(ISTEXT(C" & CStr(i + 8) & "),"""",MOD(D" & CStr(i + 8) & "-C" & CStr(i + 8) & ",1) - E" & CStr(i + 8) & ")"
                    .NumberFormat = "hh:mm;@"
                End With
                With Range("G" & CStr(i + 8))
                    .Formula = "=IF(OR(D" & CStr(i + 8) & "<=0.25,D" & CStr(i + 8) & ">=0.9583),IF(ISTEXT(C" & CStr(i + 8) & "),0," & TaxiService & "),0)"
                    .NumberFormat = "0.00;@"
                End With
                
            Else
                With Range("C" & CStr(i + 8))
                    .Formula = "=IF(AND(NOT(OR(WEEKDAY('" & i & "'!B2,11)=6,WEEKDAY('" & i & "'!B2,11)=7)),ISBLANK('" & i & "'!K4)),""Day off"",'" & i & "'!K4)"
                    .NumberFormat = "hh:mm;@"
                End With
                ''' UPD 11/27/2017 Formula with attention to LAST occurence of "x"
                With Range("D" & CStr(i + 8))
                    .Formula = "=IF(ISTEXT(C" & CStr(i + 8) & "),"""",IFERROR(LOOKUP(2,1/('" & i & "'!J:J=""x""),'" & i & "'!K:K),0))"
                    .NumberFormat = "hh:mm;@"
                End With
                    ''' UPD 1/31/2018 Formula added to count break times
                    With Range("E" & CStr(i + 8))
                        lister = Application.ConvertFormula("'" + CStr(i) + "'!J4:J28", xlA1, xlR1C1, , Range("E" & CStr(i + 8)))
                        arrayer = Application.ConvertFormula("'" + CStr(i) + "'!K4:K28", xlA1, xlR1C1, , Range("E" & CStr(i + 8)))
                        lister1 = "'" + CStr(i) + "'!J4:J28"
                        
                        Range("AA" & CStr(i + 8)).FormulaArray = "=IFERROR(SMALL(IF(" + lister1 + "=""x"",ROW(" + lister1 + ")-MIN(ROW(" + lister1 + "))+1),2), -1)" '22
                        Range("AB" & CStr(i + 8)).FormulaArray = "=IFERROR(SMALL(IF(" + lister1 + "=""x"",ROW(" + lister1 + ")-MIN(ROW(" + lister1 + "))+1),3), -1)" '23
                        Range("AC" & CStr(i + 8)).FormulaArray = "=IFERROR(SMALL(IF(" + lister1 + "=""x"",ROW(" + lister1 + ")-MIN(ROW(" + lister1 + "))+1),4), -1)" '24
                        Range("AD" & CStr(i + 8)).FormulaArray = "=SMALL(IF(" + lister1 + "=""x"",ROW(" + lister1 + ")-MIN(ROW(" + lister1 + "))+1),1)" '25
                        Range("AE" & CStr(i + 8)).FormulaArray = "=SMALL(IF(" + lister1 + "=""x"",ROW(" + lister1 + ")-MIN(ROW(" + lister1 + "))+1),2)" '26
                        Range("AF" & CStr(i + 8)).FormulaArray = "=SMALL(IF(" + lister1 + "=""x"",ROW(" + lister1 + ")-MIN(ROW(" + lister1 + "))+1),3)" '27
                        
                        
                        .FormulaR1C1 = _
        "=SUM(IF(RC[22] <> -1,IF(RC[23] <> -1, IF(RC[24] <> -1,MOD(INDEX(" + arrayer + ",RC[27]+1)-INDEX(" + arrayer + ",RC[27]),1) + MOD(INDEX(" + arrayer + ",RC[26]+1)-INDEX(" + arrayer + "," & _
        "RC[26]),1) + MOD(INDEX(" + arrayer + ",RC[25]+1)-INDEX(" + arrayer + ",RC[25]),1),MOD(INDEX(" + arrayer + ",RC[26]+1)-INDEX(" + arrayer + ",RC[26]),1) + MOD(INDEX(" + arrayer + ",RC[25]+1)" & _
        "-INDEX(" + arrayer + ",RC[25]),1)),MOD(INDEX(" + arrayer + ",RC[25]+1)-INDEX(" + arrayer + ",RC[25]),1)),0))"
                        '.FormulaArray = .FormulaR1C1
                        .NumberFormat = "hh:mm;@"
                                                
                End With
                With Range("F" & CStr(i + 8))
                    .Formula = "=IF(ISTEXT(C" & CStr(i + 8) & "),"""",MOD(D" & CStr(i + 8) & "-C" & CStr(i + 8) & ",1) - E" & CStr(i + 8) & ")"
                    .NumberFormat = "hh:mm;@"
                End With
                With Range("G" & CStr(i + 8))
                    .Formula = "=IF(OR(D" & CStr(i + 8) & "<=0.25,D" & CStr(i + 8) & ">=0.9583),IF(ISTEXT(C" & CStr(i + 8) & "),0," & TaxiService & "),0)"
                    .NumberFormat = "0.00;@"
                End With
                
            End If
        End If
    Next i
    
        Range("B7:G39").Select
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
    
    Range("B8:G8").Select
    Selection.Font.Bold = True
    Range("B42:F45").Select
    Selection.Merge True
    
    Range("B42:F42").Select
    ActiveCell.FormulaR1C1 = "Total Worked Hours"
    With ActiveCell.Characters(Start:=1, Length:=13).Font
        .Name = "Times New Roman"
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
        .ThemeFont = xlThemeFontNone
    End With
    With ActiveCell.Characters(Start:=14, Length:=5).Font
        .Name = "Times New Roman"
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
        .ThemeFont = xlThemeFontNone
    End With
    
    Range("B43:F43").Select
    ActiveCell.FormulaR1C1 = "Total Worked Days"
    With ActiveCell.Characters(Start:=1, Length:=13).Font
        .Name = "Times New Roman"
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
        .ThemeFont = xlThemeFontNone
    End With
    With ActiveCell.Characters(Start:=14, Length:=4).Font
        .Name = "Times New Roman"
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
        .ThemeFont = xlThemeFontNone
    End With
    
    Range("B44:F44").Select
    ActiveCell.FormulaR1C1 = "Total Days Taxi"
    With ActiveCell.Characters(Start:=1, Length:=6).Font
        .Name = "Times New Roman"
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
        .ThemeFont = xlThemeFontNone
    End With
    With ActiveCell.Characters(Start:=7, Length:=9).Font
        .Name = "Times New Roman"
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
        .ThemeFont = xlThemeFontNone
    End With
    
    Range("B45:F45").Select
    ActiveCell.FormulaR1C1 = "Total Cost Taxi"
    With ActiveCell.Characters(Start:=1, Length:=6).Font
        .Name = "Times New Roman"
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
        .ThemeFont = xlThemeFontNone
    End With
    With ActiveCell.Characters(Start:=7, Length:=9).Font
        .Name = "Times New Roman"
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
        .ThemeFont = xlThemeFontNone
    End With
    
    Range("B42:G45").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    
    Range("G42:G45").Font.Bold = True
    
    With Range("G42")
        .FormulaR1C1 = "=SUM(R[-33]C[-1]:R[-3]C[-1])"
        .NumberFormat = "[h]:mm;@"
    End With
    Range("G43").FormulaR1C1 = "=COUNT(R[-34]C[-1]:R[-4]C[-1])"
    Range("G44").FormulaR1C1 = "=R[-1]C"
    With Range("G45")
        .FormulaR1C1 = "=SUM(R[-36]C:R[-6]C)"
        .NumberFormat = "$#,##0.00"
    End With
    
    Range("G42:G45").Select
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
    Range("B42:G45").Select
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
    
End Function

Sub Test()
    MsgBox ("result = " & SummaryGenerate("October", 2017, 31, 3))
End Sub

Function SetName(str As String)
    InName = str
End Function

Function GetName()
    GetName = InName
End Function

Function SetResponse(str As String)
    Response = str
End Function

Sub BatchCreate()

    'part 1: reading data
Sheets("BatchCreate").Visible = True
Sheets("BatchCreate").Select

Set tbl = ActiveSheet.ListObjects("BatchList")

Dim INP() As Range
ReDim INP(tbl.ListRows.Count)
i = 0

'pick folders
For i = 1 To tbl.ListRows.Count
    
    If tbl.DataBodyRange(i, 10).Value = "" Then
        InName = tbl.DataBodyRange(i, 1).Value
        UserForm1.Show
        
        tbl.DataBodyRange(i, 10).Value = Response
        
        Unload UserForm1
    End If
    
    
        
    Set INP(i) = tbl.ListRows(i).Range
Next i

'MsgBox CStr(INP.Count), vbOKOnly
'part 2: processing
Sheets("BatchCreate").Visible = False

Application.DisplayAlerts = False

For i = 1 To tbl.ListRows.Count
    With INP(i)
        SingleCreateOverride_AndClose .Cells(1, 13).Text, _
                                      .Cells(1, 1).Text, _
                                      Val(.Cells(1, 8).Value), _
                                      Val(.Cells(1, 9).Value), _
                                      Val(.Cells(1, 5).Value), _
                                      Val(.Cells(1, 3).Value), _
                                      .Cells(1, 4).Text, _
                                      Val(.Cells(1, 12).Value)
        
    End With
Next i

Application.DisplayAlerts = True

Sheets("BatchCreate").Visible = True
Sheets("BatchCreate").Select
End Sub

Sub SelectDirectoriesForBatch()

Set tbl = ActiveSheet.ListObjects("BatchList")

Dim INP() As Range
ReDim INP(tbl.ListRows.Count)
i = 0

'pick folders
For i = 1 To tbl.ListRows.Count
    
    'If tbl.DataBodyRange(i, 10).Value = "" Then
        InName = tbl.DataBodyRange(i, 1).Value
        InToChange = tbl.DataBodyRange(i, 10).Value
        
        UserForm1.Show
        
        tbl.DataBodyRange(i, 10).Value = Response
        
        Unload UserForm1
        
        InToChange = ""
    'End If

Next i

End Sub

Sub SingleCreate_Click()
    t = Sheets("MAIN").Range("C11:F11").Cells(1, 1).Value
    SingleCreate (t)
End Sub


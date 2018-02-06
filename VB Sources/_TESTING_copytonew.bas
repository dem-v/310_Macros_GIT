Attribute VB_Name = "Module1"
Sub copytonew()
Attribute copytonew.VB_ProcData.VB_Invoke_Func = " \n14"
'
' copytonew Macro
'

'
    Workbooks.Add
    Windows("test 2.xlsm").Activate
    Sheets("MAIN").Select
    ActiveSheet.Buttons.Add(96, 98.25, 144, 48.75).Select
    ActiveSheet.Buttons.Add(49.5, 553.5, 144, 48.75).Select
    Sheets("MAIN").Copy Before:=Workbooks("Book1").Sheets(1)
    Windows("test 2.xlsm").Activate
    Sheets("MAIN").Select
    Cells.Select
    Sheets(Array("MAIN", "TOTAL")).Select
    Sheets("MAIN").Activate
    ActiveSheet.Buttons.Add(96, 98.25, 144, 48.75).Select
    ActiveSheet.Buttons.Add(49.5, 553.5, 144, 48.75).Select
    Sheets(Array("MAIN", "TOTAL")).Copy Before:=Workbooks("Book1").Sheets(1)
    Windows("test 2.xlsm").Activate
    Windows("Book1").Activate
    ChDir "F:\downloads_torrent\_DOWNLOADS"
    ActiveWorkbook.SaveAs Filename:="F:\downloads_torrent\_DOWNLOADS\310310.xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
End Sub

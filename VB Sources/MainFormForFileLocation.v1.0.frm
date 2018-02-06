VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Select Working Folder."
   ClientHeight    =   1635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9315
   OleObjectBlob   =   "MainFormForFileLocation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public Input and Response Variable

Private Sub Cancel_Click()
    Response = "C:\"
    Me.Hide
End Sub

Private Sub Label1_Click()
    Shell "explorer.exe" & " " & Label1.Caption, vbNormalFocus
End Sub

Private Sub Leave_Click()
    Response = Label1.Caption
    Me.Hide
End Sub

Private Sub SelectDir_Click()
    Root = Application.Path & "\"
        
    Set pick = Application.FileDialog(msoFileDialogFolderPicker)
    With pick
        .Title = "Select folder for " & main_initial.InName & "..."
        .AllowMultiSelect = False
        .InitialFileName = Root
        If .Show = -1 Then
            main_initial.Response = .SelectedItems(1)
            Else
            main_initial.Response = Root
        End If
    End With
    Label1.Caption = main_initial.Response
    Me.Hide
End Sub

Private Sub UserForm_Activate()
    If main_initial.InToChange = "" Then
    Label1.Caption = Application.Path & "\"
    Else
    Label1.Caption = InToChange
    End If
    Me.Caption = "Select Working Folder for " & main_initial.InName
End Sub

'Private Sub UserForm_Initialize()

'End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If main_initial.Response = "" Then main_initial.Response = Application.Path & "\"
    Me.Hide
End Sub

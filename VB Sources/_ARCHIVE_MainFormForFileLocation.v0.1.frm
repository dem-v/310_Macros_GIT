VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FileExists 
   Caption         =   "ERROR: File Exists!"
   ClientHeight    =   1620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6735
   OleObjectBlob   =   "ARCHIVE_MainFormForFileLocation.v0.1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FileExists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub overwrite_Click()
    Resp = MsgBox("This can not be undone. Continue?", vbYesNo, "Overwrite?")
    If Resp = vbYes Then DeleteFile (Label1.Caption)
    Me.Hide
End Sub

Sub DeleteFile(ByVal FileToDelete As String)
   If FileExists(FileToDelete) Then 'See above
      ' First remove readonly attribute, if set
      SetAttr FileToDelete, vbNormal
      ' Then delete the file
      Kill FileToDelete
   End If
End Sub

Private Sub rename_Click()
  
    NewName = Label1.Caption & "_bckp" & Now("mmddyy-hhmmss")
    Name Label1.Caption As NewName
    Me.Hide
End Sub

Private Sub skip_Click()
    main_initial.RenamingResponse = "SKIP"
    Me.Hide
End Sub



Private Sub UserForm_Activate()
    Label1.Caption = main_initial.InExistingFileAddress
End Sub


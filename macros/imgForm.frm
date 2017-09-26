VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} imgForm 
   Caption         =   "capture images"
   ClientHeight    =   4284
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   6300
   OleObjectBlob   =   "imgForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "imgForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 


Private Sub bClose_Click()
        imgForm.hide
End Sub

Private Sub bExport_Click()
        extractImages
End Sub

Private Sub bReview_Click()
    'LCase (Environ("SystemRoot"))
    'explorerPath = Environ("SystemRoot") & "\Explorer.exe" & " " & imgLocation()
    'Shell pathname:=explorerPath
    On Error GoTo errlab
    ActiveDocument.FollowHyperlink imgLocation()
    Exit Sub
errlab:
    MsgBox Err.Description & "<" & imgLocation() & ">"
    
End Sub


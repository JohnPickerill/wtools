VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} anchorForm 
   Caption         =   "UserForm1"
   ClientHeight    =   1656
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   4584
   OleObjectBlob   =   "anchorForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "anchorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rng As Range

Private Sub bCancel_Click()
    anchorForm.hide
End Sub

Private Sub bOK_Click()
    If Len(tName.text) < 3 Then
        MsgBox ("invalid anchor name")
        Exit Sub
    End If
    Set rng = rng.Paragraphs(1).Range
    rng.Collapse (wdCollapseStart)
    rng.Select
    rng.text = "!!(" + tName.text + ") "
    muEdit.wrapExtra rng
    anchorForm.hide
End Sub

Private Sub UserForm_Activate()
    tName.text = ""
    Set rng = Selection.Range.Duplicate
 
End Sub


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ghostForm 
   Caption         =   "Ghost Hunt"
   ClientHeight    =   4968
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   8316
   OleObjectBlob   =   "ghostForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ghostForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rng As Range
Dim app As New CcApparition
Dim ghost As CcApparition


Public Function getGhost() As CcApparition
    Set getGhost = ghost
End Function


Private Sub cbAcquire_Click()
    If app.isValid Then
        Set ghost = app
        ghostForm.hide
    Else
        MsgBox "this ghost is not valid"
    End If
End Sub

Private Sub cbCancel_Click()
    Set ghost = Nothing
    ghostForm.hide
End Sub

Private Sub cbNext_Click()
   If app.getApparition("") Then
        ctGhost.text = app.getHash
        ctContents.text = app.getContent.text
   Else
        ctGhost.text = ""
        ctContents.text = ""
   End If
End Sub


Private Sub cbRestart_Click()
    app.clear ActiveDocument
    cbNext_Click
End Sub


Private Sub cbRestore_Click()
    If app.isValid Then
        app.restore
        cbNext_Click
    Else
        MsgBox "this ghost is not valid"
    End If
End Sub

Private Sub UserForm_Initialize()
    cbNext_Click
End Sub

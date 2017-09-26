VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} hdForm 
   Caption         =   "UserForm1"
   ClientHeight    =   2316
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   3552
   OleObjectBlob   =   "hdForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "hdForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private reTest As New RegExp
Private prefix As String
Private str As String
Private r As Range


Private Sub bCancel_Click()
    hdForm.hide
End Sub

Private Sub bOK_Click()
    If (Len(tAnchor.text) < 2) Then
        MsgBox "invalid anchor id"
    End If
    
    If r.Find.Execute(findtext:=tAnchor.text) Then
        MsgBox ("Anchor already exists")
        Exit Sub
    End If
    
    muEdit.createDrop Selection.Range, tAnchor.text
    hdForm.hide
End Sub

Private Sub tAnchor_Change()
    Dim m As MatchCollection
    Set m = reTest.Execute(tAnchor.text)
    If m.count = 1 Then
        str = prefix + m(0).SubMatches(0)
    End If
    tAnchor.text = str
End Sub

Private Sub UserForm_Activate()
    
    Dim kmj As Object
    Dim p As String
    Set r = muEdit.expandArticle(ActiveDocument, Selection.Range, kmj)
    
    If r Is Nothing Then
        MsgBox "selection not in article"
        hdForm.hide
        Exit Sub
    End If
    
    prefix = "_" + kmj("id") + "_"

    
    p = "^" + prefix + "([a-z_0-9]*)$"
    
    With reTest
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = p
    End With
    
    str = prefix
    tAnchor.text = str
    
End Sub


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} resultsForm 
   Caption         =   "Results"
   ClientHeight    =   9888
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   9216
   OleObjectBlob   =   "resultsForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "resultsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private content As New StringBuilder


Public Sub setString(str As String)
    content.clear
    content.Append str
    ctResults.value = content.text
    
End Sub

Public Sub Append(str As String)
    content.Append str
    ctResults.value = content.text
    ctResults.SetFocus
    ctResults.SelStart = content.Length - Len(str)
    ctResults.SelLength = 1
    resultsForm.Repaint
End Sub

 
Public Function text()
    text = content.text
End Function
 
 
 


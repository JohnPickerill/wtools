VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} outlineForm 
   Caption         =   "Paragraphs"
   ClientHeight    =   3408
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   9972
   OleObjectBlob   =   "outlineForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "outlineForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
   
 paraBox.value = muEdit.walkDoc()
 
End Sub

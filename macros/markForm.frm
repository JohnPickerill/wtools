VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} markForm 
   Caption         =   "Guide Markup"
   ClientHeight    =   3156
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   6888
   OleObjectBlob   =   "markForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "markForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbNext_Click()
    uINext
    ctContent = Selection.text
End Sub

 

Private Sub cbPrev_Click()
    uIPrev
    ctContent = Selection.text
End Sub


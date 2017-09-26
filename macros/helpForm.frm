VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} helpForm 
   Caption         =   "Help"
   ClientHeight    =   4656
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   9432
   OleObjectBlob   =   "helpForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "helpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 

Private Sub UserForm_Initialize()
    clVersion.Caption = "Guide Plugin Version : " & kmVer
    ctbLicence = "LICENCE" & _
    "This software includes the following components" & _
    "TODO - include list of components and licences here"
End Sub

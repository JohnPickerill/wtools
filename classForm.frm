VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} classForm 
   Caption         =   "UserForm1"
   ClientHeight    =   5748
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   12516
   OleObjectBlob   =   "classForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "classForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public id As String
Public aType As String
 

Private Sub cbCancel_Click()
    classForm.hide
End Sub

Private Sub cbClear_Click()
    id = ""
    For i = 0 To clClass.ListCount - 1
             clClass.Selected(i) = False
    Next i
    classForm.hide
End Sub

Private Sub cbOk_Click()
       For i = 0 To clClass.ListCount - 1
            If clClass.Selected(i) Then
                id = clClass.List(i, 0)
            End If
       Next i
       classForm.hide
End Sub

Private Sub UserForm_Activate()
    Caption = aType
    Dim classes As Object
    Set classes = getAssociates(aType)
    
    If Not classes Is Nothing Then
        For Each c In classes("articles")
            clClass.AddItem c("id")(1)
            clClass.List(clClass.ListCount - 1, 1) = c("title")(1)
            If id = c("id")(1) Then clClass.Selected(clClass.ListCount - 1) = True
        Next c
    End If
End Sub

 

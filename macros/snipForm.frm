VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} snipForm 
   Caption         =   "Snippets"
   ClientHeight    =   2484
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   4104
   OleObjectBlob   =   "snipForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "snipForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rng As Range
Dim snippets As Object



Private Sub bCancel_Click()
    snipForm.hide
End Sub

Private Sub bOK_Click()
    If Len(cName.text) < 3 Then
        MsgBox ("invalid snippet name")
        Exit Sub
    End If
    rng.text = "[[@" & setupForm.getCode("snippet", cbPurpose.text) & ":" & cName.text & "]]"
    muEdit.wrapExtra rng
    snipForm.hide
End Sub


Private Sub UserForm_Activate()
    'take para mark out of selection
    bExists = False
    Set rng = muEdit.targetMarkup(Selection.Range)
    If Not rng Is Nothing Then
        Dim mc As MatchCollection
        Dim ReLink As New RegExp
        ReLink.MultiLine = True
        ReLink.pattern = "\[\[@([a-z]{3}):(\S+?)\]\]"
        ReLink.Global = True
        Set mc = ReLink.Execute(rng.text)
        ' should at most be one
        If mc.count > 0 Then
            cbPurpose.text = setupForm.entityTypes("snippet")(mc(0).SubMatches(0))
            cName = mc(0).SubMatches(1)
            Exit Sub
        End If
        GoTo InvalidLab
    End If
    
    If Len(Selection.Range.text) = 0 Then
        Set rng = Selection.Range
        cName.text = ""
        Exit Sub
    End If
    
    
InvalidLab:
    MsgBox "Current selection is invalid for snippet insertion modification"
    hide
    Exit Sub
End Sub


Private Sub UserForm_Initialize()
    For Each p In setupForm.entityTypes("snippet")
        cbPurpose.AddItem setupForm.entityTypes("snippet")(p)
    Next p
End Sub

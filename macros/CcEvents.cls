VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CcEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public WithEvents wdApp As Word.Application
Attribute wdApp.VB_VarHelpID = -1
 

Public Sub wdApp_DocumentBeforeClose(ByVal doc As Document, Cancel As Boolean)
    'undo all changes protect document and then save
    CloseHandler doc
End Sub

Public Sub wdApp_DocumentBeforeSave(ByVal doc As Document, SaveAsUI As Boolean, Cancel As Boolean)
    'If isGuide(doc) Then
        Cancel = BeforeSaveHandler(doc)
    'End If
End Sub



Public Sub wdApp_DocumentOpen(ByVal doc As Document)
        openhandler doc
End Sub

Private Sub wdApp_WindowActivate(ByVal doc As Document, ByVal Wn As Window)
    On Error GoTo exitLab
    Dim repo As String
    repo = ""
    If checkRepo(doc.path) Then
        Application.StatusBar = "Guidance document : state = " & getProp(doc, "guide")
    End If
exitLab:
End Sub

 

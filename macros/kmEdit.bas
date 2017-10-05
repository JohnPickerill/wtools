Attribute VB_Name = "kmEdit"

Public markup As New CcMarkup
Public muEdit As New CcEdit
Public evnts As New CcEvents
Public cctxt As New CcText

Public Cfg As New Cfg
Public Const kmVer = "3.0.0"
Public Const lockkey = "ixdkkaspdd"




Function checkSelection(kmj As Object) As Boolean
     Dim r As Range
    
     Set r = muEdit.expandArticle(ActiveDocument, Selection.Range, kmj)
     
     If r Is Nothing Then
        checkSelection = False
     Else
        checkSelection = True
        r.Select
     End If
        
End Function





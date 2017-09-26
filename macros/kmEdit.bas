Attribute VB_Name = "kmEdit"

Public markup As New CcMarkup
Public muEdit As New CcEdit
Public evnts As New CcEvents
Public cctxt As New CcText

Public Cfg As New Cfg
Public Const kmVer = "2.0.9"
Public Const lockkey = "ixdkkaspddwatsrrtcmtm"


fff

Function checkSelection(kmj As Object) As Boolean
     Dim lineType As String
     Dim action As String
     Dim r As Range
    
     Set r = muEdit.expandArticle(ActiveDocument, Selection.Range, kmj)
     
     If r Is Nothing Then
        checkSelection = False
     Else
        checkSelection = True
        r.Select
     End If
        
End Function





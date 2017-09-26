Attribute VB_Name = "Test"
 
Dim spanStyles() As String
Dim blockStyles() As String
Dim scnt As Integer
Dim bcnt As Integer
Dim kmStyles As Object

Sub ccc()
Dim x As String
Dim Y As String
x = "111"
a = Y
Y = x
x = "333"
Y = "222"
End Sub




Sub testjson()
    j = "{""j\\\\s"":""\\\\hair""}"

    Set x = JsonDecode(j)
End Sub

Sub t()
    s = muEdit.stripSpan("hhhh-!j!-fdksfkl")
End Sub

Sub test()
'Create a ribbon instance for use in this project
 
Dim ba() As String
Dim sa() As String
 
Dim s As Object
Dim k As Variant

If kmStyles Is Nothing Then
    Set kmStyles = getStyles()
End If
ReDim ba(0 To kmStyles.count - 1) As String
ReDim sa(0 To kmStyles.count - 1) As String
scnt = 0
bcnt = 0
For Each k In kmStyles.keys()
     Set s = kmStyles(k)
     If s("block") <> "" Then
             blockStyles(bcnt) = k
             bcnt = bcnt + 1
     End If
     If s("span") <> "" Then
             spanStyles(scnt) = k
              scnt = scnt + 1
     End If
Next k

End Sub
 
 
Public Sub testcluster()
    'On Error GoTo exitLab
    On Error Resume Next
    Dim cluster As String
    Dim doc As Document
    Set doc = ActiveDocument
    cluster = "no"
    cluster = doc.CustomDocumentProperties("km_cluster")
    MsgBox "1" & cluster
    cluster = ActiveDocument.CustomDocumentProperties("km_cluster")
    ' SET DEFAULTS FOR km DOCUMENTS
    MsgBox "This document contains guidance articles belonging to the cluster <" & cluster & ">"
    doc.TrackRevisions = False
    With doc.ActiveWindow.View
        .ShowRevisionsAndComments = False
        .RevisionsView = wdRevisionsViewFinal
        .FieldShading = wdFieldShadingAlways
    End With
    
    showMeta True
    Exit Sub
exitLab:
 
End Sub
 
 
Sub cfgtest()
    Cfg.cfgRead
End Sub
 
Sub tt()
    Dim fn As String
    Dim d As Document
    Set d = ActiveDocument
    fn = d.FullName
    While d.Undo
    Wend
    While Not d.Saved
        d.Undo
    Wend
End Sub
 

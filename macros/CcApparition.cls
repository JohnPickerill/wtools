VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CcApparition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim rngGhost As Range
Dim rngHunt As Range
Dim hash As String
Dim reGhost As New RegExp
Dim doc As Document

Sub setHash(str As String)
    hash = BASE64SHA1(str)
End Sub

Function getHash() As String
    getHash = hash
End Function

Function getContent() As Range
    Set getContent = doc.Range(start:=rngGhost.Paragraphs(2).Range.start, _
                                            End:=rngGhost.Paragraphs(rngGhost.Paragraphs.count - 1).Range.End)
End Function

 


Function isValid() As Boolean
isValid = False
If rngGhost.Paragraphs.count < 2 Then Exit Function
If Not reGhost.test(rngGhost.Paragraphs(1).Range.text) Then Exit Function
If Not reGhost.test(rngGhost.Paragraphs(rngGhost.Paragraphs.count).Range.text) Then Exit Function
isValid = True
End Function

Sub restore()
    For Each para In rngGhost.Paragraphs
        muEdit.unhide (para.Range)
    Next para
    rngGhost.Paragraphs(1).Range.delete
    rngGhost.Paragraphs(rngGhost.Paragraphs.count).Range.delete
End Sub


Sub setRange(r As Range)
    Set rngGhost = r.Duplicate
    setHash rngGhost.text
End Sub


Private Sub parseGhost(r As Range)
    Dim m As MatchCollection
    Set m = reGhost.Execute(r.Paragraphs(1).Range.text)
    Set rngGhost = r
    hash = m(0).SubMatches(0)
End Sub

Public Sub clear(ByVal dc As Document)
    Set doc = dc
    hash = ""
    Set rngGhost = Nothing
    Set rngHunt = Nothing
End Sub

Public Sub delete()
    rngGhost.text = ""
    rngGhost.Collapse (wdCollapseEnd)
    Set rngHunt = rnghost.Duplicate
End Sub



Public Function getApparition(Optional ghost As String = "") As Boolean
    'Dim rng As Range
    If rngHunt Is Nothing Then
        Set rngHunt = doc.Range
    End If
    getApparition = False
    With rngHunt.Find
       .ClearFormatting
       .MatchWildcards = False
       .text = "*ghost=" & ghost
       .Execute Forward:=True, Wrap:=False
    End With
    If rngHunt.Find.Found Then
        s = rngHunt.Paragraphs(1).Range.start
    Else
        Exit Function
    End If
    rngHunt.Find.Execute Forward:=True, Wrap:=False
    If rngHunt.Find.Found Then
        f = rngHunt.Paragraphs(rngHunt.Paragraphs.count).Range.End
        parseGhost doc.Range(start:=s, End:=f)
        getApparition = True
    End If
End Function


Private Sub Class_Initialize()
    With reGhost
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "\*ghost=(.{5})"
    End With
End Sub

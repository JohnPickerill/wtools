VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFixDrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private re_start As New RegExp
Private re_end As New RegExp
Private hn As String
Private hind As Integer
Private hdid As String
Private hdopen As Boolean


Private Sub Class_Initialize()
    hdid = "_hdid"
    hind = 0
    
    With re_start
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        'this is the pattern I use in python but VBA can't do positive look behind so ..
        '.Pattern = "^-<:<{{(drop\!?(?:(?<=\!)([_a-z0-9]+))?):[ \t]*([a-zA-Z0-9 \t]*):>:>-"
        .pattern = "((-<:<{{drop)(\!?)(.*):(.*>:>-.*))\x0D"
    End With
    
    With re_end
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "((-<\+<drop)(\!?)(.*)(}}>:>-.*))\x0D"
    End With
    
 
    
    
End Sub

 

Private Function new_hdid() As String
    hind = hind + 1
    hdid = hn + "_" + CStr(hind)
End Function



Public Function fixHot(para As Paragraph) As Boolean
    Dim m As MatchCollection
    Dim txt As String
    para.Range.Select
    Set m = re_start.Execute(para.Range.text)
        If m.count = 1 Then
            If hdopen Then
                Dim fu As New cFixDrop
                Set para = fu.walk(para, hdid)
                Set fu = Nothing
                Exit Function
            End If
            new_hdid
            hdopen = True
            If m(0).SubMatches(2) = "!" Then
                 If Left((m(0).SubMatches(3)), 1) = "_" Then
                    If (Left((m(0).SubMatches(3)), 5) <> "_drop") Then
                        Exit Function
                    End If
                 End If
                 
            End If
            txt = m(0).SubMatches(1) + "!" + hdid + ":" + m(0).SubMatches(4)
            Selection.Range.Font.Hidden = False
            Selection.TypeText text:=txt
            muEdit.markMarkup para.Range, True
            Exit Function
        End If
        

        
        
        Set m = re_end.Execute(para.Range.text)
        If m.count = 1 Then
            If Not hdopen Then
                Set para = para.Previous
                fixHot = True
                Exit Function
            End If
            hdopen = False
            If m(0).SubMatches(2) = "!" Then
                If Left(m(0).SubMatches(3), 1) = "_" Then
                    If (Left((m(0).SubMatches(3)), 5) <> "_drop") Then
                        Exit Function
                    End If
                End If
            End If
            txt = m(0).SubMatches(1) + "!" + hdid + m(0).SubMatches(4)
            Selection.Range.Font.Hidden = False
            Selection.TypeText text:=txt
            muEdit.markMarkup para.Range, True
            Exit Function
        End If

End Function


Public Function walk(p As Paragraph, Optional uid As String) As Paragraph
    Dim para As Paragraph
    Dim m As MatchCollection
    If (uid <> "") Then
        hn = uid
    End If
    hdopen = False
    Set para = p
    Do While Not (para Is Nothing)
        If muEdit.analMeta(para.Range.text, m) Then
            hn = "_" + m(0).SubMatches(1)
            hind = 0
            Exit Do
        End If

        If fixHot(para) Then
            Exit Do
        End If
        Set para = para.Next
    Loop
    Set walk = para
End Function



Public Sub convert()
    Dim p As Paragraph

    Set p = ActiveDocument.Paragraphs(1)
    p.Range.Select
    Do While Not (p Is Nothing)
        Set p = walk(p)
        If (Not p Is Nothing) Then
            Set p = p.Next
        End If
    Loop
    
    Set cfdx = Nothing
End Sub

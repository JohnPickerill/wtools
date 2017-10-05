VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CcEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private ReFormat As New RegExp
Private ReMarkup As New RegExp
Private ReArticle As New RegExp
Private ReOld As New RegExp
Private ReSpan As New RegExp
Private log As StringBuilder


Enum muType
    tTrove
    tGuide
    tBlock
    tAction
    tComment
    tGhost
End Enum


Const c_mu_sh = "-<%s-<" ' trove only
Const c_mu_sb = "-<%s:<" ' block start
Const c_mu_se = "-<%s+<" ' Guide only
Const c_mu_sa = "-<%s&<" ' action/command e.g. apparate
Const c_mu_sc = "-<%s*<" ' comment
Const c_mu_sg = "-<%s~<" ' ghost
Const c_mu_f = ">:%s>-"
' markdown strings
Const c_Para_f = vbLf & vbLf

'word meta strings
Const c_article_s = "{article:{"
Const c_meta_f = "<--"

' word formating
Const cHighlight = wdYellow
Const cLowlight = wdGreen


Private Sub Class_Initialize()

    ReFormat.Global = True
    ReFormat.MultiLine = True
    ReFormat.pattern = "%s"
        
    With ReMarkup
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "(-*)<(?:([-+:&*~])([a-z]*?)<)(.*?)(?:#|>:\3)>(-*)"
    End With
    With ReArticle
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = c_article_s & "([s,f]+),(.+?)}}(.*)<(-+)"
        
    End With
    With ReOld
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "{""meta"":\s*{""(start|finish)"":""(.+)""}}(.*)<-+"
    End With
    With ReSpan
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "(-*)(!\S!)(-*)"
    End With
   
End Sub

Public Sub setLog(Lg As StringBuilder)
    Set log = Lg
End Sub


Public Function stripSpan(str As String) As String
    stripSpan = ReSpan.replace(str, "$2")
End Function


Function isMeta(linestr) As Boolean
    isMeta = ReArticle.test(linestr)
End Function

Function analMeta(linestr, m As MatchCollection) As Boolean

 
 
    Set m = ReArticle.Execute(linestr)
    
    analMeta = False
    If (m.count < 1) Then Exit Function
    If (m(0).SubMatches.count < 4) Then Exit Function
    
    analMeta = True
    End Function
    
Function splitMeta(linestr, _
        Optional id As String, Optional action As String, Optional smeta As Object _
        ) As Boolean
    Dim m As MatchCollection
    splitMeta = False
    If Not analMeta(linestr, m) Then Exit Function
        
    If Not IsMissing(id) Then id = m(0).SubMatches(1)
    If Not IsMissing(action) Then action = m(0).SubMatches(0)
    'TODO should have exception handling
    If Not IsMissing(smeta) Then Set smeta = JsonDecode(m(0).SubMatches(2))
    If smeta Is Nothing Then
        If log Is Nothing Then
            MsgBox "Invalid json at line " & linestr
        Else
            log.Append "ERROR: Invalid json at line " & linestr
        End If
        GoTo errorLab
    End If
    'if type hasn't been set it
    If smeta("type") = "" Then smeta("type") = setupForm.getType
    If smeta("purpose") = "" Then smeta("purpose") = setupForm.getPurpose
    
    'TODO check matches document type need to do this whether or not para passed.
    If smeta("type") <> setupForm.getType Then
        If log Is Nothing Then
            MsgBox ("meta data type:" & smeta("type") & " does not correspond to document type:" & setupForm.getType)
        Else
            log.Append "ERROR: meta data type:" & smeta("type") & " does not correspond to document type:" & setupForm.getType
        End If
    End If
    If smeta("purpose") <> setupForm.getPurpose Then
        If log Is Nothing Then
            MsgBox ("meta data purpose:" & smeta("purpose") & " does not correspond to document type:" & setupForm.getPurpose)
        Else
            log.Append "ERROR: meta data type:" & smeta("purpose") & " does not correspond to document type:" & setupForm.getPurpose
        End If
    End If
    
    splitMeta = True
errorLab:
    End Function
    
    
Function tagMeta(linestr, _
        Optional m_start As Integer, Optional m_length As Integer, Optional leadout As Integer _
        ) As Boolean

    Dim m As MatchCollection
    tagMeta = False
    If Not analMeta(linestr, m) Then Exit Function
    
    If Not IsMissing(m_start) Then m_start = Len(c_article_s) + Len(m(0).SubMatches(0)) + Len(m(0).SubMatches(1)) + 3
    If Not IsMissing(m_length) Then m_length = Len(m(0).SubMatches(2))
    If Not IsMissing(leadout) Then leadout = Len(m(0).SubMatches(3)) + 1  'include end of para mark
    
    tagMeta = True
End Function

' TODO this is a placeholder and is wrong
Function splitMarkup(linestr) As MatchCollection
    Set splitMarkup = ReMarkup.Execute(linestr)
End Function

Function testMarkup(linestr) As Boolean
    testMarkup = ReMarkup.test(linestr)
End Function




' Wrap markup
Private Function wrapMu(rng As Range, Optional mt As muType = tGuide, Optional name As String = "") As Boolean
    If rng.End = rng.Paragraphs(1).Range.End Then
        ' Do not want end of para mark
        Set rng = ActiveDocument.Range(start:=rng.start, End:=rng.End - 1)
    End If
    Dim t As String
    t = rng.text
    If Len(t) = 0 Then
        wrapMu = False
        Exit Function
    End If
    Select Case mt
        Case tGuide
            rng.InsertBefore ReFormat.replace(c_mu_se, "")
        Case tBlock
            rng.InsertBefore ReFormat.replace(c_mu_sb, "")
        Case tAction
            rng.InsertBefore ReFormat.replace(c_mu_sa, "")
        Case tTrove
            rng.InsertBefore ReFormat.replace(c_mu_sh, "")
        Case tComment
            rng.InsertBefore ReFormat.replace(c_mu_sc, "")
        Case tGhost
            rng.InsertBefore ReFormat.replace(c_mu_sg, "")
    End Select
 
    rng.InsertAfter ReFormat.replace(c_mu_f, "")
    muEdit.markMarkup rng.Paragraphs(1).Range, True
    wrapMu = True
End Function

' TODO refactor these 2 for duplicate code
'Guide Only
Public Function wrapExtra(rng As Range) As Range
    Set wrapExtra = Nothing
    If rng.Paragraphs.count > 1 Then
        MsgBox "Selection to big"
        Exit Function
    End If
    If Not muEdit.targetMarkup(rng) Is Nothing Then
        MsgBox "selected text already contains markup"
        Exit Function
    End If
    If rng.Fields.count > 0 Then
        'TODO hide field code text in markdown conversion
        MsgBox "Error : Selection that contains field codes "
        Exit Function
    End If
    
    If wrapMu(rng, tGuide) Then
        Set wrapExtra = rng.Duplicate
    End If
    
End Function


' Trove Only
Public Function wrapHidden(rng As Range) As Range
    Set wrapHidden = Nothing
    If rng.Paragraphs.count > 1 Then
        MsgBox "Selection to big"
        Exit Function
    End If
    If Not muEdit.targetMarkup(rng) Is Nothing Then
        MsgBox "selected text already contains markup"
        Exit Function
    End If

    If wrapMu(rng, tTrove) Then
        Set wrapHidden = rng.Duplicate
    End If
End Function








Public Function createKmjObj() As Object
        Dim keywords As New Collection
        Dim facets As New Collection
        Dim extLinks As New Collection
        Dim cluster As New Collection
        Dim clusters As New Collection
        Dim intlinks As New Collection
        Dim items As New Collection

            
        Set o = CreateObject("Scripting.Dictionary")

        On Error Resume Next
        o("id") = "Enter Unique ID"
        
        o("title") = ""
        o("type") = setupForm.getType
        o("purpose") = setupForm.getPurpose
        If o("purpose") = "legislation" Then
            o("scope") = "The full text of the legislation"
        Else
            o("scope") = ""
        End If
        
        
        'o("lastupdate") = Format(Now, "yyyy-mm-ddThh:mm:ss.000Z")
        'o("popularity") = "3"
        
        'old schema
        cluster.Add setupForm.getCluster() ' first cluster should be the parent document
        o.Add "cluster", cluster
        'new schema
        Set c = CreateObject("Scripting.Dictionary")
        c("cluster") = setupForm.getCluster() ' first cluster should be the parent document
        c("priority") = 9999 ' default priority
        clusters.Add c 'add object to array
        o.Add "clusters", clusters ' add array to structure
        
        o("author") = ""
        o("owner") = ""
        o("expert") = ""
        o("sensitivity") = "normal"
        o("archive") = "false"
        
        Set m = CreateObject("Scripting.Dictionary")
        m("change") = ""
        m("significance") = "low"
        o.Add "master", m
        
        'o.Add "items", items
        'o.Add "keywords", keywords
        'o.Add "facets", facets
        'o.Add "kmlinks", intlinks
        'o.Add "extlinks", extLinks
        
        'o("content") = ""
        'o("markup") = ""


 
        
       
        Set createKmjObj = o
End Function



Public Function createSnippetObj() As Object
        Dim clusters As New Collection

        On Error Resume Next
        Set o = CreateObject("Scripting.Dictionary")
        
        o("id") = "Enter Unique ID"
        o("type") = setupForm.getType
        o("purpose") = setupForm.getPurpose
        
        o("author") = ""
        o("owner") = ""
        o("expert") = ""
        
        'o("markup") = ""
        
        Set m = CreateObject("Scripting.Dictionary")
        o.Add "master", m
        
        o.Add "clusters", clusters ' add array to structure
        

        Set createSnippetObj = o
End Function



Public Function setPeople(kmj As Object) As Object
        On Error Resume Next
        ' TODO check this is what we want if value is set or key doesn't exist
        kmj("author") = ""
        kmj("expert") = ""
        kmj("owner") = ""
        
        kmj("author") = ActiveDocument.BuiltInDocumentProperties("Last Author")
        kmj("expert") = ActiveDocument.BuiltInDocumentProperties("Author")
        kmj("owner") = ActiveDocument.Variables("AuthorEmail")
        Set updateAuthor = kmj
End Function


Sub markArticle(p As Paragraph, show As Boolean)
    Dim rng As Range
    Dim m_start As Integer
    Dim m_length As Integer
    Dim leadout As Integer
    Dim m As MatchCollection
    

    
    Set rng = p.Range
    rng.TextRetrievalMode.IncludeHiddenText = True
    If Not tagMeta(rng.text, m_start, m_length, leadout) Then
        Exit Sub
    End If
  
 
    'TODO need to look at performance as this sub is could be called in a loop
    Dim blnTrackChangesOn As Boolean
    blnTrackChangesOn = ActiveDocument.TrackRevisions
    ' make sure we swith track changes on
    On Error GoTo errlab
    ActiveDocument.TrackRevisions = False
 
 

    

    rng.Style = ActiveDocument.styles("Normal")
    rng.HighlightColorIndex = wdNoHighlight
    rng.Font.Hidden = Not show
    
    Set rng = ActiveDocument.Range(start:=rng.start, End:=rng.End - leadout)
    rng.HighlightColorIndex = cHighlight
    rng.Font.ColorIndex = wdBlue
    s = rng.start + m_start
    
    Set rng = ActiveDocument.Range(start:=s, End:=s + m_length)
    rng.HighlightColorIndex = cLowlight
    rng.Font.ColorIndex = wdRed
    rng.Font.Hidden = True

errlab:
    ActiveDocument.TrackRevisions = blnTrackChangesOn
     
    'content controls not available in word 97
    'Selection.Range.ContentControls.add (wdContentControlRichText)
    'Selection.ParentContentControl.Title = "end:" & kmj("id")
    'Selection.ParentContentControl.Tag = "metaTag"
    'Selection.ParentContentControl.LockContentControl = True
    'Selection.ParentContentControl.LockContents = True
 
End Sub


Public Function unhide(r As Range) As Boolean
    Dim rng As Range
    Dim b As Boolean
    b = False
    Set rng = targetMarkup(r)
    If Not rng Is Nothing Then
       Dim m As Match
       'there must only be one
       For Each m In splitMarkup(rng.text)
           leadout = Len(m.SubMatches(4))
           leadin = Len(m.SubMatches(0))
           If m.SubMatches(1) = "-" Or m.SubMatches(1) = "~" Then
               b = True
               Set st = ActiveDocument.Range(start:=rng.start, End:=rng.start + leadin + 3 + Len(m.SubMatches(2)))
               Set fn = ActiveDocument.Range(start:=rng.End - 3 - leadout - Len(m.SubMatches(2)), End:=rng.End)
               fn.delete
               st.delete
           End If
       Next m
       rng.Font.Hidden = False
       rng.HighlightColorIndex = wdNoHighlight
       'rng.Font.ColorIndex = wdAuto
    End If
    unhide = b
End Function

Function padPara(rng As Range) As Range
' this is to overcome a word/trove feature that causes problems if the first character of a para is hidden
Dim r As Range
On Error GoTo errlab
Set padPara = rng.Duplicate
If rng.start = 0 Then Exit Function
Set r = ActiveDocument.Range(start:=rng.start - 1, End:=rng.start)
If r.text = vbCr Then
        r.MoveEnd wdCharacter, 1
        r.start = r.End - 1
        r.InsertBefore " "
        padPara.MoveStart wdCharacter, 1
End If
Exit Function
errlab:
    MsgBox Err.Description & " paragraph start not found "
End Function


 




Private Sub markit(r As Range, show As Boolean)
    
    ' the range passed in here must always be a single markup
    ' TODO there is no need for this loop as there will only be one
    For Each m In splitMarkup(r.text)
        's = r.start + m.FirstIndex
        leadout = Len(m.SubMatches(4))
        leadin = Len(m.SubMatches(0))
        
        
        'hide/show including leading and leadout
        'TODO need to check if end of para mark is included
        If Not show Then
            Dim rng As Range

            Set rng = padPara(r)
            rng.Font.Hidden = True
            'Trove only
            If m.SubMatches(1) = "-" Or m.SubMatches(1) = "~" Then
                st = rng.start + leadin + 3 + Len(m.SubMatches(2))
                fn = rng.End - 3 - leadout - Len(m.SubMatches(2))
                Set rng = ActiveDocument.Range(start:=st, End:=fn)
                rng.Font.Hidden = False
                rng.HighlightColorIndex = wdNoHighlight
                'rng.Font.ColorIndex = wdAuto
            End If
            'if line is hidden except for para mark then hide para mark
            Set rng = r.Paragraphs(1).Range.Duplicate
            rng.TextRetrievalMode.IncludeHiddenText = False

            'TODO there has got to be a more elegant way of doing this
            ' If Mid(rng.text, 1, 1) = vbCr Then
            '    rng.Font.Hidden = True
            'TODO not sure if this is more elegant but need to add padded text.
            If cctxt.isWhitespace(rng.text) Then
                rng.Font.Hidden = True
            End If
      
        Else
            'Set rng = ActiveDocument.Range(start:=s, End:=s + m.Length)
            Set rng = r.Duplicate
            rng.Font.Hidden = False
            rng.HighlightColorIndex = wdNoHighlight
            'highlight except for leading and trailing -
            Set rng = ActiveDocument.Range(start:=rng.start + leadin, End:=rng.End - leadout)
            rng.HighlightColorIndex = cHighlight
            'rng.Font.ColorIndex = wdBlue
            'if we just have a hidden paramark left before the end of line then unhide it
            'rng.Collapse wdCollapseEnd
            'rng.MoveStart unit:=wdCharacter, count:=leadout
            rng.MoveEnd unit:=wdParagraph, count:=1
            rng.Font.Hidden = False
        End If
    Next m
End Sub

Private Sub markSpan(r As Range, show As Boolean)
        'Note we shouldn't have to worry about paramarks as and range passed here should match the pattern !.!
        Dim rng As Range
        
        If Not show Then
            Set rng = padPara(r)
            rng.Font.Hidden = True
        Else
            Set rng = r.Duplicate
            rng.Font.Hidden = False
            Dim m As MatchCollection
            Set m = ReSpan.Execute(r.text)
            rng.MoveStart wdCharacter, Len(m(0).SubMatches(0))
            rng.MoveEnd wdCharacter, -Len(m(0).SubMatches(2))
            rng.HighlightColorIndex = cHighlight
        End If
End Sub



Sub markMarkup(r As Range, show As Boolean)
    Dim rng As Range
    Dim b As Boolean
    
 
    Dim blnTrackChangesOn As Boolean
    blnTrackChangesOn = ActiveDocument.TrackRevisions
    
    'make sure we set trackrevisions back on
    On Error GoTo errlab
    ActiveDocument.TrackRevisions = False
    
    Set rng = r.Duplicate

    b = True
    Do While b
         With rng.Find
            .ClearFormatting
            .MatchWildcards = True
            .text = "-@\<[:\-+&*\~]\<*\>:\>-@"

            .Execute Forward:=True, Wrap:=False
         End With
         b = False
         If rng.Find.Found Then
            If rng.InRange(r) Then
                b = True
                markit rng, show
            End If
         End If
         Loop

    
   Set rng = r.Duplicate

    b = True
    Do While b
         With rng.Find
            .ClearFormatting
            .MatchWildcards = True
            .text = "-@\!?\!-@"

            .Execute Forward:=True, Wrap:=False
         End With
         b = False
         If rng.Find.Found Then
            If rng.InRange(r) Then
                b = True
                markSpan rng, show
                
            End If
         End If
         Loop

    
    
    
    
errlab:
    ActiveDocument.TrackRevisions = blnTrackChangesOn
 
 
    
     'content controls not available in word 97
     'Selection.Range.ContentControls.add (wdContentControlRichText)
     'Selection.ParentContentControl.Title = "end:" & kmj("id")
     'Selection.ParentContentControl.Tag = "metaTag"
     'Selection.ParentContentControl.LockContentControl = True
     'Selection.ParentContentControl.LockContents = True
     
End Sub






Function targetMarkup(r As Range, Optional mode As Integer = 0) As Range

    Dim rng As Range
    Dim b As Boolean
    Dim dir As Boolean
    
    Set targetMarkup = Nothing
    If mode = 0 Then
        Set rng = r.Paragraphs(1).Range.Duplicate
        cwrap = wdFindStop
    Else
        Set rng = r.Duplicate
        cwrap = wdFindAsk
    End If
     
   
    
    b = True
    Do While b
         With rng.Find
            .ClearFormatting
            .MatchWildcards = True
            .text = "-@\<[:\-+&*~]\<*\>:\>-@"
            .Execute Forward:=(mode >= 0), Wrap:=cwrap
         End With
         b = False
         Select Case mode
            Case 0
                If rng.Find.Found Then
                       If rng.start < r.End Then
                           If rng.End > r.start Then
                               Set targetMarkup = rng
                           Else
                               b = True
                           End If
                       End If
                End If
             Case 1, -1
                Set targetMarkup = rng
         End Select
         Loop
    Exit Function
End Function


Public Function expandArticle(ByVal doc As Document, r As Range, kmj As Object) As Range
    ' expand selection for whole article and check syntax
    ' if error return nothing and set kmj = nothing
    ' if no metadata return original range and set kmj = nothing
    
    Dim action As String
    Dim start_id As String
    Dim end_id As String
    Dim id As String
    Dim article_s As Paragraph
    Dim article_f As Paragraph
    Dim o As Object
    Dim asWas As Boolean
    
    Set kmj = Nothing
 
     'endfound = False
     Set expandArticle = r
     Set st = r.Paragraphs(1)
     siz = r.Paragraphs.count
     Set article_s = Nothing
     Set article_f = Nothing
     cnt = 1
     Set para = st
     
    ' show hidden text so we can find article tags
    asWas = doc.ActiveWindow.View.ShowHiddenText
    If Not asWas Then
        doc.ActiveWindow.View.ShowHiddenText = True
    End If
     
     ' find start by looking before selection
     Do While Not (para Is Nothing)
        Set o = readMeta(para.Range.text, start_id, action)
        If Not o Is Nothing Then
            Select Case action
                Case "f", "x"
                    Set para = Nothing
                Case "s"
                    Set article_s = para
                    If cnt = 1 Then  ' TODO neater way of testing if first para in range
                        Set st = st.Next
                        siz = siz - 1
                    End If
                    Set para = Nothing
                    Set kmjObj = o
                Case Else
                    Set para = para.Previous ' TODO don't understand  why this isn't an error
            End Select
        Else
            Set para = para.Previous
        End If
        cnt = cnt + 1
     Loop
     

      
 
      ' look forward through selection (except last line)
     cnt = 1
 
     Set para = st
     Do While (cnt < siz)
        Set o = readMeta(para.Range.text, id, action) 'checks if selection contains meta data
        If o Is Nothing Then
                    Set para = para.Next
                    cnt = cnt + 1
        Else
                    If log Is Nothing Then
                        MsgBox "unexpected article marker: structure error: type = " & action
                    Else
                        log.Append "ERROR: unexpected article marker: structure error: type = " & action
                    End If
                    
                    Set expandArticle = Nothing
                    GoTo endLabel
        End If
      Loop
            
 
     
     ' find end of article
     Do While Not (para Is Nothing)
        Set o = readMeta(para.Range.text, end_id, action)
        If o Is Nothing Then
            Set para = para.Next
        Else
            Select Case action
                Case "s", "x"
                    If log Is Nothing Then
                        MsgBox "unexpected metadata marker  found"
                    Else
                        log.Append "ERROR: unexpected metadata marker found"
                    End If
                    Set expandArticle = Nothing
                    GoTo endLabel
                Case "f"
                    If Not (start_id = end_id) Then
                        Set expandArticle = Nothing
                        If log Is Nothing Then
                            MsgBox "unmatched article marker found"
                        Else
                            log.Append "ERROR: unmatched article marker found"
                        End If
                        GoTo endLabel
                    Else
                        Set article_f = para
                        Set para = Nothing
                    End If
                Case Else
                    Set para = para.Next ' TODO don't undesratnad why this isn't an error
            End Select
        End If
        
        If article_s Is Nothing Then
            GoTo endLabel ' no need to search further
        End If
     Loop
     ' TODO think about why this is inside the loop ?
     If article_f Is Nothing Then
        Set expandArticle = Nothing
        If log Is Nothing Then
            MsgBox "matching finish article marker not found"
        Else
            log.Append "ERROR matching finish article marker not found"
        End If
        GoTo endLabel
     End If
     
     Set kmj = kmjObj
     Set expandArticle = doc.Range(start:=article_s.Range.start, End:=article_f.Range.End)
endLabel:
    If doc.ActiveWindow.View.ShowHiddenText <> asWas Then
        doc.ActiveWindow.View.ShowHiddenText = asWas
    End If

   
End Function

Function readMeta(linestr As String, id As String, action As String) As Object
    Dim m As MatchCollection
    Dim tmeta As String
    Dim rm As Object
    On Error GoTo errlab:
    Set readMeta = Nothing
    If Not splitMeta(linestr, id, action, rm) Then
        Exit Function
    End If
    Set readMeta = rm
    ' change to new schema for clusters
    rm("id") = cctxt.cleanUID(rm("id"))
    If action = "s" Then
        'ensure km uids conform to convention
        
        'If rm.Exists("kmlinks") Then
        '    For Each lnk In rm("kmlinks")
         '       lnk("id") = cctxt.cleanUID(lnk("id"))
        '    Next lnk
        'End If
                      

        'If Not rm.Exists("clusters") Then
        '    Dim clusters As New Collection
        '    rm.Add "clusters", clusters
        '    For Each c In rm("cluster")
        '        Set cs = CreateObject("Scripting.Dictionary")
        '        cs("cluster") = c
        '        cs("priority") = 9999 ' default priority
        '        rm("clusters").Add cs 'add object to array
        '    Next c
        'End If
        
        'If Not rm.Exists("items") Then
        '    Dim items As New Collection
        '    rm.Add "items", items
        'End If

        If rm Is Nothing Then GoTo errlab:
        On Error GoTo errlab
        If Not rm("master").Exists("significance") Then
            rm("master")("significance") = "low"
        End If

        If Not rm.Exists("archive") Then
            rm("archive") = "false"
        End If
            
        
        If Not rm("master").Exists("change") Then
            rm("master")("change") = ""
        End If
        
 
        If rm("type") <> "snippet" Then
            If Not rm.Exists("sensitivity") Then  'TODO temporary fix until ES reloaded
                rm("sensitivity") = "unknown"
            End If
        End If
    End If
    Exit Function
errlab:
    action = "x"
    If log Is Nothing Then
        MsgBox "Error: Invalid meta data format"
    Else
        log.Append "Error: Invalid meta data format"
    End If

End Function



Public Function walkDoc() As String
    Dim txt As String
    Dim pre As String
    Dim mdStr As New StringBuilder
    Dim id As String
    Dim articleId As String
    Dim action As String
    Dim rng As Range

    
    pre = ""
    txt = ""
    articleId = ""


    Dim para As Paragraph
    asWas = ActiveWindow.View.ShowHiddenText
 
    ActiveWindow.View.ShowHiddenText = True
    
    
    
  
    ActiveDocument.Background.Fill.Visible = msoTrue
    ActiveDocument.Background.Fill.ForeColor.RGB = RGB(102, 153, 255)
    ActiveDocument.Background.Fill.Solid
    ActiveDocument.Background.Fill.Transparency = 0#
    ActiveDocument.Background.Fill.PresetTextured msoTextureDenim
    ActiveWindow.View.DisplayBackgrounds = True
    
 
 For Each para In ActiveDocument.Paragraphs
            Set rng = para.Range
            Set o = readMeta(rng.text, id, action)
            
            'TODO check for duplicate articles
            If Not o Is Nothing Then
                mdStr.Append vbLf
                Select Case action
                Case "s"
                    If articleId <> "" Then mdStr.Append "ERROR=>"
                    mdStr.Append "Article:" & action & ":" & id & vbLf
                    articleId = id
                Case "f"
                    If articleId <> id Then mdStr.Append "ERROR=>"
                    mdStr.Append "Article:" & action & ":" & id & vbLf
                    articleId = ""
                End Select
            End If
            
' TODO take out hotdrops

                'If para.Range.Information(wdWithInTable) Then
                '    txt = "Table"
                'Else
                '    txt = rng.FormattedText.text
                '    pre = markup.headLevel(para)
                'End If
                'If ((pre <> "")) Then
                '    markup.cleanStr txt
                '    mdStr.Append vbLf & pre & txt & "  " & vbLf
                'Else
                    mdStr.Append "."
                'End If


continueloop:
        pre = ""
        txt = ""
     
 
 Next para
 walkDoc = mdStr.text
 
 ActiveWindow.View.ShowHiddenText = asWas
 ActiveWindow.View.DisplayBackgrounds = False
End Function

Public Function bracketSection(r As Range, bracketType As String, Optional bComment As Boolean = False)
Dim ct As muType

pcount = r.Paragraphs.count
Set startMeta = r.Paragraphs(1).Range
Set endMeta = r.Paragraphs(pcount).Range
startMeta.Select
Selection.HomeKey unit:=wdLine
Selection.TypeParagraph
Selection.MoveUp unit:=wdParagraph, count:=1
Selection.Style = ActiveDocument.styles("Normal")
Selection.TypeText text:="{{" & bracketType & ":"
'TODO this is really crude - need to do something more elegant
If bComment Then
    ct = tComment
Else
    ct = tBlock
End If
wrapMu Selection.Paragraphs(1).Range, ct


endMeta.Select
Selection.EndKey unit:=wdLine
Selection.TypeParagraph
Selection.TypeText text:=bracketType & "}}"
Selection.Style = ActiveDocument.styles("Normal")
If bComment Then
    ct = tComment
Else
    ct = tGuide
End If
wrapMu Selection.Paragraphs(1).Range, ct

r.Select
End Function

Public Sub setTableForm(cls As String)
Dim str As String
str = "table=" & cls
Selection.HomeKey unit:=wdLine
Selection.TypeParagraph
Selection.MoveUp unit:=wdParagraph, count:=1
Selection.Style = ActiveDocument.styles("Normal")
Selection.TypeText text:=str
wrapMu Selection.Paragraphs(1).Range, tAction
End Sub


Public Function createDrop(r As Range, anchor As String)
' TODO check that this is a valid range for a hotdrop
' TODO different hitdrop types
' Add markup
bracketSection r, "drop!" & anchor
End Function

Public Function createGhost(r As Range) As CcApparition
'TODO do we really need an apparition cost or do we just remember the hash, what is the copy method for ?
' actually I think I overcomplicated this when I was playing with copy and paste
Dim a As New CcApparition
a.setRange r
Dim p As Paragraph
For Each p In r.Paragraphs
    wrapMu p.Range, tGhost
Next p
bracketSection r, "*ghost=" & a.getHash, True
Set createGhost = a
End Function
 
Public Function postApparition(rng As Range, a As CcApparition)
rng.Select
Selection.EndKey unit:=wdLine
Selection.TypeParagraph
Selection.TypeText "apparition=" & a.getHash
wrapMu Selection.Paragraphs(1).Range, tAction
End Function




Sub addKmjMeta(kmj As Object, article As Range)
     Dim lineType As String
     Dim action As String
     Dim place As Range
     Dim startMeta As Range
     Dim endMeta As Range
     
     Set article = Selection.Range
     
     
     article.Select
     
     pcount = article.Paragraphs.count

     'TODO check this works with track changes possibly need to delete the old and then insert the new
     Set startMeta = article.Paragraphs(1).Range
     Set endMeta = article.Paragraphs(pcount).Range
     
     If Not isMeta(endMeta.text) Then
         endMeta.Select
         Selection.EndKey unit:=wdLine
         Selection.TypeParagraph
         Selection.EndKey unit:=wdLine
         Set endMeta = Selection.Range
     End If

     
     If Not isMeta(startMeta.text) Then
        startMeta.Select
        Selection.HomeKey unit:=wdLine
        Selection.TypeParagraph
        Selection.MoveUp unit:=wdParagraph, count:=1
        Set startMeta = Selection.Paragraphs(1).Range
     End If
     
     'This is to overcome a feature in word that messes things up if the next para is in a table
     If (article.Paragraphs.count) > 1 Then
        If article.Paragraphs(2).Range.Information(wdWithInTable) Then
           startMeta.MoveEnd wdCharacter, -1
           startMeta.InsertParagraphAfter
        End If
     End If
        
     'This is to overcome a feature in word that messes things up if the next para is in a table
     Dim x As Range
     Set x = endMeta.Next
     If (Not x Is Nothing) Then
        If (x.Information(wdWithInTable)) Then
            endMeta.MoveEnd wdCharacter, -1
            endMeta.InsertParagraphAfter
        End If
     End If
        
        
     ' write start meta data
     startMeta.Select
     jsonstr = JsonEncode(kmj)
     Selection.Style = ActiveDocument.styles("Normal")
     Selection.TypeText text:=c_article_s & "s," & CStr(kmj("id")) & "}}" & jsonstr & c_meta_f
     muEdit.markArticle startMeta.Paragraphs(1), True
     
     ' write finish meta data
     endMeta.Select
     Selection.Style = ActiveDocument.styles("Normal")
     Selection.TypeText text:=c_article_s & "f," & CStr(kmj("id")) & "}}{""id"":""" & CStr(kmj("id")) & """}" & c_meta_f
     muEdit.markArticle endMeta.Paragraphs(1), True
              
     Selection.setRange start:=startMeta.start, End:=endMeta.End

End Sub


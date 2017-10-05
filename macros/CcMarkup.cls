VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CcMarkup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ReHeading As New RegExp

Private ReCrLf As New RegExp
Private ReCell As New RegExp
Private ReClean As New RegExp
Private ReMark As New RegExp
Private ReImage As New RegExp
Private ReImageMark As New RegExp
Private ReEscapedLf As New RegExp
Private ReHide As New RegExp
Private reAmp As New RegExp
Private reLt As New RegExp
Private headOffset As Integer
Private ReCre As New RegExp
Private ReBullet As New RegExp
Private ReAction As New RegExp
Private ReSpace As New RegExp
Private ReLeadSpace As New RegExp
Public reMdLink As New RegExp
Public reMdxLink As New RegExp ' for use in cleanup
Private reRelative As New RegExp

Private log As StringBuilder

Private tableForm As String
      
'table
Const c_col_delm = "|"
Const c_Row_mid = "| -- "
Const c_Row_end = "|" & vbLf
Const lineend = "    " & vbLf

Private Sub Class_Initialize()
' rules
    With ReSpace
        .Global = True
        .IgnoreCase = False
        .pattern = " "
    End With


    With ReHeading
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = "Heading ([0-9])"
    End With

    With ReCell
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "\x0D\x07"
    End With

    With ReMark
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "[\x15]"
    End With

 

    With ReClean
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "([\s\S]+)[\x0D\x0A]*$"
    End With

    With ReImage
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
       ' .Pattern = "\[\[\s*([\S]+?)(?:\s*(?:\|| )([\s\S]+?))?\]\]"
       ' only accept | as deliminator and not space because filename could contain space
       .pattern = "\[\[\s*([\S ]+?)(?:\s*(?:\|)([\s\S]+?))?\]\]"
    End With

    With ReEscapedLf
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "/\x0D"
    End With
 
    ' need to allow for CREs and infils that are surrounded by literal * but allow for lists and rules
    With ReCre
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "\*([^\*\n]*)\*"
    End With
 
    With ReBullet
         .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "( {0,6}\*)?([^\x0D\x0A]*([\x0D\x0A]|$))"
    End With
     
     With ReAction
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "(.*)=(.*)"
    End With
 
     With reAmp
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "&(?!amp|lt)"
    End With
 
    With reLt
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "<(?!br\/>)"
    End With
  
  
     With ReLeadSpace
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "^(?: (?=-+<))?(.*)"
    End With
 
    With reMdLink
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        'look for markdown link syntax that indicates a link
        .pattern = "(.*?)\[\[([^\| ]+?)((?:\]\])|(?:[\| ].*?\]\]))(.*)"
    End With
    
    With reMdxLink
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        'look for markdown link syntax that indicates a link
        .pattern = "\[.*\]\(.*\)"
    End With
    
    
     With reRelative
        .Global = False
        .MultiLine = False
        .IgnoreCase = True
        'look for markdown link syntax that indicates a link
        .pattern = "^(\.\.\/)+(.*)"
    End With
 
    tableForm = "NORMAL"
End Sub

Function escHtml(txt As String) As String
'TODO I don't think that > needs to be escaped I'm not sure about " but I think that \ might need to be
'escaping quote will mess up the image markup as this contains quotes and may upset other stuff
    txt = reAmp.replace(txt, "&amp;")
    txt = reLt.replace(txt, "&lt;")
    escHtml = txt
End Function


Function cleanStr(ByRef txt As String) As String
        'Do While ((Right(txt, 1) = vbCr) Or (Right(txt, 1) = vbLf))
        '    txt = Left(txt, (Len(txt) - 1))
        'Loop
        'txt = RTrim(txt)
        'cleanStr = "ok"
        txt = ReMark.replace(txt, "")
        'JOHNP can't remember what this bit does looks like it strips LF CR off the end.
        'TODO do I need to get rid of special chanacters
        Dim m As MatchCollection
        Set m = ReClean.Execute(txt)
        If m.count > 0 Then
            txt = m(0).SubMatches(0)
        End If
End Function


Function stripCell(str) As String
'TODO we could replace the end of cell marker (/x13/x07) with a cell deliminator
'we could use /x13 in the middle of cells to guess at merged cells
    stripCell = ReCell.replace(str, "")
    stripCell = replace(stripCell, Chr(11), "<br/>")
    stripCell = replace(stripCell, Chr(13), "<br/>")
End Function


Function get_apparition(para As String) As Range
    Dim a As New CcApparition
    If Not a.getApparition(para) Then
        Set get_apparition = ActiveDocument.Range(start:=0, End:=0)
        If log Is Nothing Then
            MsgBox "Warning: could not find ghost: " & para
        Else
            log.Append "ERROR: could not find ghost: " & para
        End If
    Else
        Set get_apparition = a.getContent()
    End If
End Function


Sub extract_command(markup, com, para)
    com = ""
    para = ""
    Dim mc As MatchCollection
    Set mc = ReAction.Execute(markup)
    If mc.count > 0 Then
        com = mc(0).SubMatches(0)
        para = mc(0).SubMatches(1)
    End If
End Sub



Function extractMu(doc As Document, linestr As String, Optional appset As Boolean = False) As String
    Dim m As Match
    Dim act As String
    Dim str As New StringBuilder
     
    pos = 0 ' zero based index
        
    'Strip a leading space before a markup
    linestr = ReLeadSpace.Execute(linestr)(0).SubMatches(0)
    'strip dashes from span markup
    'TODO the regex is a bit crude
    linestr = muEdit.stripSpan(linestr)
 

        
    For Each m In muEdit.splitMarkup(linestr)
        lead = m.FirstIndex - pos
        If lead > 0 Then
            str.Append Mid(linestr, pos + 1, lead)
        End If
        pos = m.FirstIndex + m.Length
        act = m.SubMatches(1)
        Select Case act
            Case "-" 'trove only
            Case ":" ' This is for block starts like dropdowns
                str.Append vbLf
                str.Append m.SubMatches(3)
                'need to prevent rest of line being appended (especially the CR
                pos = Len(linestr)
            Case "+" 'guide only
                Dim intlink As MatchCollection
                Dim lnk As String
                Dim rm As String
                rm = m.SubMatches(3)
                Do
                   Set intlink = reMdLink.Execute(rm)
                   If intlink.count = 1 Then
                           str.Append intlink(0).SubMatches(0)
                           lnk = intlink(0).SubMatches(1)
                           If (Left(lnk, 1) <> "@") And (InStr(lnk, "://") = 0) Then
                               lnk = cctxt.cleanUID(lnk)
                           Else
                               lnk = replace(lnk, " ", "%20") 'todo CREATE A FUNCTION TO CLEAN UP LINKS PROERLY AND ALSO APPLY AS VALIDATORS ON RELEVANT FORMS
                           End If
                           str.Append "[["
                           str.Append lnk
                           str.Append intlink(0).SubMatches(2)
                           rm = intlink(0).SubMatches(3)
                   Else
                       str.Append rm
                   End If
                Loop While (intlink.count = 1)
 
            Case "~" ' ghost
                If appset Then str.Append m.SubMatches(3)
            Case "&" ' insert command (apparition)
                'TODO get hash a bit more elegantly
                Dim com As String
                Dim para As String
                extract_command m.SubMatches(3), com, para
                
                'Dim com As String
                'Dim para As String
                'Dim mc As MatchCollection
                'Set mc = ReAction.Execute(m.SubMatches(3))
                'If mc.count > 0 Then
                '    com = mc(0).SubMatches(0)
                '    para = mc(0).SubMatches(1)
                'End If
                If com = "table" Then tableForm = para ' insert formating instruction for next table
                If com = "apparition" Then ' apparate a ghost at this point
                    Dim apps As String
                    apps = markup(doc, get_apparition(para), True)
                    ' TODO find a more elegant way - markup will append an CR because we are calling it recursively at this point I need to get rid of it
                    If Len(apps) > 5 Then
                            apps = Left(apps, Len(apps) - 5)
                    End If
                    str.Append apps
                    'Dim app As Range
                    'Dim apps As String
                    'Dim a As New CcApparition
                    'If Not a.getApparition(para) Then
                    '    If Log Is Nothing Then
                    '        MsgBox "Warning: could not find ghost: " & para
                    '    Else
                    '        Log.Append "ERROR: could not find ghost: " & para
                    '    End If
                    'Else
                    '    apps = markup(a.getContent(), True)
                    '    'TODO find a more elegant way - markup will append an CR because we are calling it recursively at this point I need to get rid of it
                    '    If Len(apps) > 5 Then
                    '        apps = Left(apps, Len(apps) - 5)
                    '    End If
                    '    str.Append apps
                    'End If
                    
                    'need to prevent rest of line being appended (especially the CR
                    pos = Len(linestr)
                 End If
            Case "*" ' defines start of somthing
                

        End Select
    Next m
    
    If pos = 0 Then
        ' no markup found

        extractMu = linestr
    Else
        ' add rest of line
        If (pos) < Len(linestr) Then
            str.Append Right(linestr, Len(linestr) - pos)
        End If

        extractMu = str.text
    End If
    'extractMu = nl & ReMarkup.Replace(lineStr, vbLf & "$2")
End Function


Function escCre(linestr As String) As String
        Dim mcre As MatchCollection
        Set mcre = ReBullet.Execute(linestr)
        escCre = ""
        For i = 0 To mcre.count - 1
            escCre = escCre + mcre(i).SubMatches(0) + ReCre.replace(mcre(i).SubMatches(1), "\*$1\*")
        Next i
End Function
 

 
Function extractHL(r As Range) As String
   ' so at this point we probably want hidden text to be hidden but before we want it unhidden
   ' assume markup isn;t hidden
     
    'hide Fields
    If r.Fields.count > 0 Then
        For Each fld In r.Fields
            Select Case fld.Type
                Case wdFieldRef
                    fld.Result.Font.Hidden = True
            End Select
        Next fld
    End If

    
    If r.Hyperlinks.count = 0 Then
        r.TextRetrievalMode.IncludeHiddenText = False
        extractHL = r.text
    Else
        Dim str As New StringBuilder
        Dim s As Long
        Dim hl As Hyperlink
        Dim rng As Range
        s = r.start
        For Each hl In r.Hyperlinks
            
            Set rng = ActiveDocument.Range(start:=s, End:=hl.Range.start)
            rng.TextRetrievalMode.IncludeHiddenText = False
            str.Append rng.text
           
            s = hl.Range.End
            str.Append "["
            str.Append hl.TextToDisplay
            str.Append "]("
            If hl.Address <> "" Then
                Dim mc_url As MatchCollection
                Set mc_url = reRelative.Execute(hl.Address)
                If mc_url.count > 0 Then
                    url = Cfg.getVar("sharepoint") & mc_url(0).SubMatches(1)
                Else
                    url = hl.Address
                End If
                str.Append ReSpace.replace(url, "%20")
            Else
                str.Append "Bookmark." & hl.SubAddress
            End If
            str.Append " """
            str.Append hl.TextToDisplay
            str.Append """)"
        Next hl
        Set rng = ActiveDocument.Range(start:=s, End:=r.End)
        rng.TextRetrievalMode.IncludeHiddenText = False
        str.Append rng.text
        extractHL = str.text
    End If
    
    ' expose fields
    If r.Fields.count > 0 Then
        For Each fld In r.Fields
            If fld.Type = wdFieldRef Then
                 fld.Result.Font.Hidden = False
            End If
        Next fld
    End If
    
End Function






Private Function makeTable(doc As Document, para As Paragraph) As String
'TODO need to change this so it does not step outside the range being converted
    Dim md As New StringBuilder
    Dim line2 As New StringBuilder
    Dim tbl As Table
    Dim r As row
    Dim c As Cell
    Dim ltype As Long
    Dim x As String
    
    If tableForm = "NONE" Then
        makeTable = ""
        tableForm = "NORMAL"
        Exit Function
    End If
    
    Set tbl = para.Range.Tables(1)
    
    ccnt = tbl.Columns.count
    rcnt = tbl.Rows.count
    For i = 1 To ccnt
        line2.Append c_Row_mid
    Next i
    line2.Append c_Row_end
    ltype = 1
    
 
    For Each r In tbl.Rows
 
        Select Case ltype
            Case 3
            Case 2
                md.Append line2.text
                ltype = 3
            Case 1
                ltype = 2
        End Select
        
        cnt = 0
        For Each c In r.Cells
            cnt = cnt + 1
            md.Append c_col_delm
            Dim txt As String
            txt = extractHL(c.Range)
            txt = extractMu(doc, txt, False)
            md.Append stripCell(txt)
        Next c
        Do While cnt < ccnt
            md.Append c_col_delm
            cnt = cnt + 1
        Loop
        md.Append c_Row_end
    Next r
  
    md.Append vbLf
    makeTable = md.text
End Function


Public Function headLevel(para) As String
    Set m = ReHeading.Execute(para.Style.NameLocal)

    Dim lev As Integer
    
    
    If m.count > 0 Then
        lev = CInt(m.item(0).SubMatches.item(0))
        If headOffset = 99 Then
            headOffset = 2 - lev
        End If
        
        lev = lev + headOffset
        Select Case lev
        Case 1
            headLevel = "## "
        Case 2
            headLevel = "## "
        Case 3
            headLevel = "### "
        Case 4
            headLevel = "#### "
        Case Else
            headLevel = "##### "
        End Select
    Else
        headLevel = ""
   End If
End Function


Private Function mdPara(para) As String

    mdPara = headLevel(para)

    If mdPara <> "" Then
        Exit Function
    End If
        
        If para.Range.ListFormat.ListType = WdListType.wdListBullet Then
            Select Case para.Range.ListFormat.ListLevelNumber
            Case 1
                mdPara = "* "
            Case 2
                mdPara = " * "
            Case 3
                mdPara = "  * "
            Case 4
                mdPara = "   * "
            Case Else
                mdPara = "    * "
            End Select
            Exit Function
        End If
        
      If para.Range.ListFormat.ListType = WdListType.wdListSimpleNumbering Then
            Select Case para.Range.ListFormat.ListLevelNumber
            Case 1
                mdPara = "1. "
            Case 2
                mdPara = " 1. "
            Case 3
                mdPara = "  1. "
            Case Else
                mdPara = "   1. "
            End Select
            Exit Function
      End If
End Function

Public Sub setLog(Lg As StringBuilder)
    Set log = Lg
End Sub
    
 
 

Public Function markup(doc As Document, r As Range, Optional appset As Boolean = False) As String
    Dim lineType As String
    Dim action As String
    Dim name As String
    Dim id As String
    Dim md As New StringBuilder
    Dim para As Paragraph
    Dim txt As String
    Dim kmj As Object
    Dim pre As String
    Dim inTab As Boolean
    Dim shp As InlineShape
        
    inTab = False
  
    ind = 0
    
    FldCodes = doc.ActiveWindow.View.ShowFieldCodes
    doc.ActiveWindow.View.ShowFieldCodes = False
    hide = doc.ActiveWindow.View.ShowHiddenText
    If Not hide Then
        doc.ActiveWindow.View.ShowHiddenText = True
    End If
    
    headOffset = 99
    
    For Each para In r.Paragraphs
    
        'this needs to be first because if in tab need to ingore other options

        If para.Range.Information(wdWithInTable) Then
            If Not inTab Then
                md.Append makeTable(doc, para)
                inTab = True
            End If
            GoTo nextPara
        End If
        inTab = False
        
        If muEdit.isMeta(para.Range.text) Then
            GoTo nextPara
        End If
  
            
        ' check for images
        missingCnt = 0
        Dim imageStr As String
        If para.Range.InlineShapes.count > 0 Then
            Dim s As String
            s = para.Range.start
            f = para.Range.End
            For Each img In para.Range.InlineShapes
                Dim mi As MatchCollection
                Set mi = ReImage.Execute(img.AlternativeText)
                imageStr = ""
                If s < f Then
                    Set trng = ActiveDocument.Range(start:=s, End:=img.Range.start)
                    md.Append trng.text
                    s = img.Range.End
                End If
                If mi.count > 0 Then
                    title = "Untitled Picture"
                    'TODO could make this slightly more efficient
                    If mi(0).SubMatches.count > 1 Then
                        If mi(0).SubMatches(1) <> "" Then
                            title = mi(0).SubMatches(1)
                        End If
                    End If
                    url = ReSpace.replace(mi(0).SubMatches(0), "%20")
                    If Left(url, 1) <> "_" Then
                        url = cleanFilename() & "/" & url
                    Else
                        url = "_common/" & url
                    End If
                    
                    imageStr = " ![" & title & "](" & url & " """ & title & """) " & vbLf
                End If
                If imageStr = "" Then
                        imageStr = " ![missing image](/missing.png "" Image has not been marked up properly in source document "") " & vbLf
                        missingCnt = missingCnt + 1
                End If
                    
                md.Append imageStr
            Next img
            If s < f Then
                Set trng = ActiveDocument.Range(start:=s, End:=f)
                md.Append trng.text
            End If
            md.Append lineend
            GoTo nextPara
        End If
        
        Dim rx As Range
        Set rx = para.Range.Duplicate
        txt = extractHL(rx)
        txt = extractMu(doc, txt, appset)
        If Not appset Then
            txt = escCre(txt)
        End If
        
        cleanStr txt ' get rid of field marks etc
        'whitespace whith a pre produces spurious lines in teh HTML
        If cctxt.isWhitespace(txt) Then GoTo nextPara
        pre = mdPara(para)
        md.Append pre
        md.Append txt
 
        'New para in markup  NB issue with recursive calls via extractMu
        md.Append lineend
 
      
nextPara:
    Next para
    markup = escHtml(md.text)
    
    If doc.ActiveWindow.View.ShowHiddenText <> hide Then
        doc.ActiveWindow.View.ShowHiddenText = hide
    End If
    doc.ActiveWindow.View.ShowFieldCodes = FldCodes
End Function


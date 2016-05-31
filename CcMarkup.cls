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

Private reHide As New RegExp
 
Private headOffset As Integer



       
'table
Const c_col_delm = "|"
Const c_Row_mid = "| -- "
Const c_Row_end = "|" & vbLf


Private Sub Class_Initialize()
' rules
    With ReHeading
        .Global = True
        .MultiLine = True
        .ignorecase = False
        .Pattern = "Heading ([0-9])"
    End With

    With ReCell
        .Global = True
        .MultiLine = True
        .ignorecase = True
        .Pattern = "[\x0D\x0A\x07]"
    End With

    With ReMark
        .Global = True
        .MultiLine = True
        .ignorecase = True
        .Pattern = "[\x15]"
    End With

    With ReClean
        .Global = True
        .MultiLine = True
        .ignorecase = True
        .Pattern = "([\s\S]+)[\x0D\x0A]*$"
    End With

    With ReImage
        .Global = True
        .MultiLine = True
        .ignorecase = True
        .Pattern = "\[\[\s*([\S]+?)(?:\s*(?:\|| )([\s\S]+?))?\]\]"
    End With

 
     With ReEscapedLf
        .Global = True
        .MultiLine = True
        .ignorecase = True
        .Pattern = "/\x0D"
    End With
 
End Sub

Function escHtml(txt As String) As String
'TODO I don't think that > needs to be escaped I'm not sure about " but I think that \ might need to be
    txt = Replace(txt, "&", "&amp;")
    txt = Replace(txt, "<", "&lt;")
    txt = Replace(txt, """", "&quot;")
    escHtml = txt
End Function


Function cleanStr(ByRef txt As String) As String
        'Do While ((Right(txt, 1) = vbCr) Or (Right(txt, 1) = vbLf))
        '    txt = Left(txt, (Len(txt) - 1))
        'Loop
        'txt = RTrim(txt)
        'cleanStr = "ok"
        txt = ReMark.Replace(txt, "")
        Dim m As MatchCollection
        Set m = ReClean.Execute(txt)
        If m.count > 0 Then
            txt = m(0).SubMatches(0)
        End If
        
End Function


Function stripCell(str) As String
'TODO we could replace the end of cell marker (/x13/x07) with a cell deliminator
'we could use /x13 in the middle of cells to guess at merged cells
    stripCell = ReCell.Replace(str, "")
End Function


 
 
 
' TODO think this can go
Sub convertOld(para As Paragraph)
    Dim m As MatchCollection
    linestr = para.Range.text
    
    Set m = ReOld.Execute(linestr)
    If (m.count > 0) Then
        If (m(0).SubMatches.count > 2) Then
            meta = m(0).SubMatches(2)
            action = m(0).SubMatches(0)
            id = m(0).SubMatches(1)
            linestr = c_article_s & Left(action, 1) & "," & id & "}}" & meta & "<--" & vbCrLf
            para.Range.text = linestr
            muEdit.markArticle para.Previous, True
        End If
    End If
End Sub


Function extractMu(linestr As String) As String
    Dim m As Match
    Dim str As New StringBuilder
     
    pos = 0 ' zero based index
        
    For Each m In muEdit.splitMarkup(linestr)
        lead = m.FirstIndex - pos
        If lead > 0 Then
            str.Append Mid(linestr, pos + 1, lead)
        End If
        pos = m.FirstIndex + m.Length
        If m.SubMatches(1) <> "-" Then
            If m.SubMatches(1) = ":" Then str.Append vbLf
            str.Append m.SubMatches(3)
        End If
    Next m
    
    If pos = 0 Then
        ' no markup found
        extractMu = linestr
    Else
        ' add rest of line
        If (pos + 1) < Len(linestr) Then
            str.Append Right(linestr, Len(linestr) - pos)
        End If
        extractMu = str.text
    End If
    'extractMu = nl & ReMarkup.Replace(lineStr, vbLf & "$2")
End Function


Function extractHL(r As Range) As String
    'hide Fields
    If r.Fields.count > 0 Then
        For Each fld In r.Fields
            If fld.Type = wdFieldRef Then
                 fld.Result.Font.Hidden = True
            End If
        Next fld
    End If

     
    'TODO so at this point we probably want hidden text to be hidden but before we want it unhidden
    'TODO however we need to be sure that the markup isn't hidden
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
            str.Append "[["
            If hl.Address <> "" Then
                str.Append hl.Address
            Else
                str.Append "Bookmark." & hl.SubAddress
            End If
            str.Append "|"
            str.Append hl.TextToDisplay
            str.Append "]]"
        Next hl
        str.Append ActiveDocument.Range(start:=s, End:=r.End).text
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






Private Function makeTable(para As Paragraph) As String
'TODO need to change this so it does not step outside the range being converted
    Dim md As New StringBuilder
    Dim line2 As New StringBuilder
    Dim tbl As Table
    Dim r As row
    Dim c As Cell
    Dim ltype As Long
    Dim x As String
    
    Set tbl = para.Range.Tables(1)
    
    ccnt = tbl.Columns.count
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
            x = c.Range.text
            md.Append stripCell(c.Range.text)
        Next c
        Do While cnt < ccnt
            md.Append c_col_delm
            cnt = cnt + 1
        Loop
        md.Append c_Row_end
    Next r
  
    makeTable = md.text
End Function


Public Function headLevel(para) As String
    Set m = ReHeading.Execute(para.Style.NameLocal)

    Dim lev As Integer
    
    
    If m.count > 0 Then
        lev = CInt(m.Item(0).SubMatches.Item(0))
        If headOffset = 99 Then
            headOffset = 2 - lev
        End If
        
        lev = lev + headOffset
        Select Case lev
        Case 1
            headLevel = "# "
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


Public Function markup(r As Range) As String
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
    
    hide = ActiveWindow.View.ShowHiddenText
    ActiveWindow.View.ShowHiddenText = True
    headOffset = 99
    
    For Each para In r.Paragraphs
    
        'this needs to be first because if in tab need to ingore other options

        If para.Range.Information(wdWithInTable) Then
            If Not inTab Then
                md.Append makeTable(para)
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
            For Each Img In para.Range.InlineShapes
                Dim mi As MatchCollection
                Set mi = ReImage.Execute(Img.AlternativeText)
                imageStr = ""
                If s < f Then
                    Set trng = ActiveDocument.Range(start:=s, End:=Img.Range.start)
                    md.Append trng.text
                    s = Img.Range.End
                End If
                If mi.count > 0 Then
                    title = "Untitled Picture"
                    'TODO could make this slightly more efficient
                    If mi(0).SubMatches.count > 1 Then
                        If mi(0).SubMatches(1) <> "" Then
                            title = mi(0).SubMatches(1)
                        End If
                    End If
                    imageStr = " ![" & title & "](" & cleanFilename() & "/" & mi(0).SubMatches(0) & " """ & title & """) " & vbLf
                End If
                If imageStr = "" Then
                        imageStr = " ![missing image](/missing.png "" Image has not been marked up properly in source document "") " & vbLf
                        missingCnt = missingCnt + 1
                End If
                    
                md.Append imageStr
            Next Img
            If s < f Then
                Set trng = ActiveDocument.Range(start:=s, End:=f)
                md.Append trng.text
            End If
            md.Append "   " & vbLf
            GoTo nextPara
        End If
        
            
        txt = extractHL(para.Range)
        txt = extractMu(txt)
        pre = mdPara(para)
        cleanStr txt
        md.Append pre
        md.Append txt
        'TODO don't understand what the issue is here however appending more than one lf messes up hotdrops
        If pre = "" Then
            ' start new para
            md.Append "    "
        End If
        md.Append vbLf
      
 
 
nextPara:
    Next para
 
    markup = escHtml(md.text)
    ActiveWindow.View.ShowHiddenText = hide
End Function


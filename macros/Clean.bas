Attribute VB_Name = "Clean"
'word meta strings
Const c_article_s = "{article:{"
Const c_meta_f = "<--"
Dim flds() As String

Private tableForm As String

Sub testClean()
    flds = Split("keywords,kmlinks,extlinks,facets,content,markup,items,class,sdlt,fees,cluster", ",")
    cleanFile ActiveDocument
End Sub
 
 
 
Sub cleanFile(ByRef dc As Document)
    dc.Activate
    With dc.ActiveWindow.View.RevisionsFilter
        .markup = wdRevisionsMarkupNone
        .View = wdRevisionsViewFinal
    End With
    
    'delete hidden tables
    'naturalise apparitions
    'delete trove only
    'naturalise guide only
    'remove ghosts
    'remove unneeded metadata ?
    'remove trove links
    dc.TrackRevisions = False
    showMeta True
    dc.ActiveWindow.View.ShowHiddenText = True
    resultsForm.Append vbTab & "cleaning ..." & vbCrLf
    cleandoc dc
    resultsForm.Append vbTab & "ghost removal" & vbCrLf
    removeGhosts dc
End Sub

Sub cleandoc(dc As Document)
    tableForm = "NORMAL"
    Dim para As Paragraph
    Set para = dc.Range.Paragraphs(1)
    Do While Not para Is Nothing
        
        para.Range.Select
        If para.Range.Information(wdWithInTable) Then
            If tableForm = "NONE" Then
                 resultsForm.Append vbTab & "hidden tables" & vbCrLf
                 Dim tt As Table
                 Set tt = para.Range.Tables(1)
                 tt.delete
            End If
        End If
        para.Range.Select


        'meta data
        Dim kmj As Object
        Dim start_id As String
        Dim action As String
        If muEdit.isMeta(para.Range.text) Then
            resultsForm.Append vbTab & "meta data" & vbCrLf
            Set kmj = muEdit.readMeta(para.Range.text, start_id, action)
            If Not kmj Is Nothing And action = "s" Then
                kmj("master")("significance") = "low"
                kmj("master")("comment") = "Data cleanup and migration"
                kmj("commit_comment") = kmj("master")("change")
                
                For fld = 0 To UBound(flds)
                    If kmj.Exists(flds(fld)) Then
                           kmj.Remove flds(fld)
                    End If
                Next fld

                jsonstr = JsonEncode(kmj)
                Dim pr As Range
                Set pr = para.Range.Duplicate
                pr.MoveEnd wdCharacter, -1 ' exclude paragraph mark
                pr.text = c_article_s & "s," & CStr(kmj("id")) & "}}" & jsonstr & c_meta_f
                muEdit.markArticle para, True
            End If
            GoTo nextPara
        End If
    
        ' trove references
        delTroveFields para.Range

        'clean tags - need to check it handles images
        cleanRange dc, para.Range
        


nextPara:
        Set para = para.Next
        Loop
End Sub


Sub delTroveFields(r As Range)
    'hide Fields
    Dim fld As Field
    If r.Fields.count > 0 Then
        r.Select
        For Each fld In r.Fields
            Select Case fld.Type
                Case wdFieldRef, wdFieldGoToButton, wdFieldPrivate, wdFieldMacroButton
                    resultsForm.Append vbTab & "trove reference" & vbCrLf
                    fld.delete
            End Select
        Next fld
    End If
End Sub



Sub cleanRange(ByVal dc As Document, r As Range)
    Dim rng As Range
    Dim b As Boolean
    
    On Error GoTo errlab

    
    Set rng = r.Duplicate

    b = True
    firstMarkup = True
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

                'markit rng, True
                cleanMarkup dc, rng
            End If
         End If
         
         Loop
    Exit Sub
errlab:
    MsgBox "error"
    Resume Next
End Sub


Private Sub cleanMarkup(ByVal dc As Document, r As Range)
    For Each m In muEdit.splitMarkup(r.text)
        
        leadout = Len(m.SubMatches(4))
        leadin = Len(m.SubMatches(0))
        act = m.SubMatches(1)
        
        Select Case act
            Case "+"  ' Guide only, end of hotdrops and links
                guideOnly dc, r, m.SubMatches(3)
                'r.text = m.SubMatches(3)
                'r.Collapse wdCollapseEnd
                'r.Font.Hidden = False
                  
            Case "-"  ' Trove Only
                 unpadPara dc, r
                 r.text = ""
                 r.HighlightColorIndex = wdNoHighlight
                 r.Collapse wdCollapseEnd
            Case "&" ' commands e.g. Apparate, table
                Dim com As String
                Dim para As String
                markup.extract_command m.SubMatches(3), com, para
                'If com = "table" Then tableForm = para ' insert formating instruction for next table
                If com = "apparition" Then ' apparate a ghost at this point
                    Dim app As Range
                    Set app = get_apparition(dc, para)

                    unpadPara dc, r
                    If app.End > app.start Then
                        app.Copy
                        r.Paste
                    Else
                        r.text = ""
                    End If
                    r.Collapse wdCollapseStart
                End If
                If com = "table" Then
                    tableForm = para ' insert formating instruction for next table
                End If
            Case "~" ' Ghost contents
                 unpadPara dc, r
                 r.text = m.SubMatches(3)
                 r.Font.Hidden = False
                 r.Collapse wdCollapseEnd

                 
                
            Case ":" 'This is for block starts like dropdowns
            Case "*" 'start of block e.g. ghost
                  ' should find ghost id check if its apparated anywhere and if it isn't delete it
        End Select
    Next m
End Sub

Sub guideOnly(dc As Document, r As Range, content As String)
    If markup.reMdxLink.test(content) Then
        GoTo exitLab
    End If

    If Left(content, 5) = "drop!" Or Left(content, 4) = "blk!" Then  ' hotdrop or block
        GoTo exitLab
    End If
 
    Dim intlink As MatchCollection
    Set intlink = markup.reMdLink.Execute(content) ' old style link
    If intlink.count = 1 Then
        'convertLink r, intlink
        GoTo exitLab
    End If
    unpadPara dc, r
    r.text = content
    r.HighlightColorIndex = wdNoHighlight


exitLab:
   r.Font.Hidden = False
   r.Collapse wdCollapseEnd
End Sub


Sub convertLink(r As Range, intlink As MatchCollection)
        Dim str As New StringBuilder
        Dim lnk As String
        Dim lab As String
        lnk = intlink(0).SubMatches(1)
        If (Left(lnk, 1) <> "@") And (InStr(lnk, "://") = 0) Then
            lnk = cctxt.cleanUID(lnk)
        Else
            lnk = replace(lnk, " ", "%20") 'todo CREATE A FUNCTION TO CLEAN UP LINKS PROERLY AND ALSO APPLY AS VALIDATORS ON RELEVANT FORMS
        End If
 
        lab = replace(intlink(0).SubMatches(2), "|", "", 1, 1)
        lab = Left(lab, Len(lab) - 2)
        
        str.Append "-<+<["
        str.Append lab
        str.Append "]("
        str.Append lnk
        str.Append """link to guidance article"""
        str.Append ")>:>-"
        r.text = str.text
End Sub


Function unpadPara(dc As Document, rng As Range) As Range
' this is to overcome a word/trove feature that causes problems if the first character of a para is hidden
Dim p As Range
Set p = rng.Paragraphs(1).Range
If rng.start - p.start = 1 Then
    Dim r As Range
    Set r = dc.Range(p.start, p.start + 1)
    If r.text = " " Then
        r.Font.Hidden = False
        r.text = ""
    End If
End If
End Function

'====== Ghost stuff

Sub removeGhosts(ByRef doc As Document)
    Dim app As New CcApparition
    app.clear doc
    Do While app.getApparition("")
        app.delete
        Loop
End Sub

Function get_apparition(dc As Document, para As String) As Range
    Dim a As New CcApparition
    a.clear dc
    If Not a.getApparition(para) Then
        Set get_apparition = dc.Range(start:=0, End:=0)
    Else
        Set get_apparition = a.getContent()
    End If
End Function

'=========================

Sub cleanAllDirectories()
     Dim source As String
     Dim target As String
 
     source = Cfg.getVar("webDav") & "\" & "delivered"
     target = Cfg.getVar("webDav") & "\" & "content"
     flds = Split("keywords,kmlinks,extlinks,facets,content,markup,items,class,sdlt,fees", ",")
     
     Dim dirs() As String
     dirs = Split(",algorithms,snippets,items,help", ",")
     'dirs = Split("help", ",")

     For i = 0 To UBound(dirs)
        resultsForm.show
        resultsForm.Append "analysing directory " & dirs(i) & vbCrLf
        copyAllFiles source, target, dirs(i)
     Next i
     resultsForm.Append "complete"
End Sub



Sub copyAllFiles(source As String, target As String, fold As String)

   Dim file As String
   Dim fileFilter As String
   Dim filename As String

   Dim ad As Document

   Dim folder As String
   
   folder = "\" & fold & "\"

   Dim reFilter As New RegExp
    With reFilter
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        '.pattern = "^" + pattern + ".*docx?"
        .pattern = "banks.*"
    End With
    
    Set ad = ActiveDocument

    'For s = asc("0") To asc("Z")
    For s = asc("I") To asc("J")
        If (s <= asc("9") Or s >= asc("A")) Then
            fileFilter = source & folder & Chr(s) & "*.doc*"
            On Error GoTo fileError
            file = dir(fileFilter)
            On Error GoTo 0
            While (file <> "")
                 copyFile source, target, folder, file
                 file = dir
            Wend
        End If
    Next s
    ad.Activate
    Exit Sub
fileError:
    MsgBox ("Cannot find sharepoint folder, please open in File Explorer from Internet Explorer ")
    Err.Raise 555, "Sharepoint", "Error reading folder"
End Sub
    




Sub copyFile(source As String, target As String, folder As String, file As String)

    Dim dc As Document

    sourceName = source & folder & file
    targetName = target & "\" & file
    If Right(targetName, 1) <> "x" Then
         targetName = targetName & "x"
    End If
        
    
    resultsForm.Append "Cleaning .." & vbCrLf
    On Error GoTo fileErr
    Set dc = Application.Documents.Open(sourceName, AddToRecentFiles:=False, Visible:=True, ReadOnly:=True)
    On Error GoTo 0
  
    resultsForm.Append file & vbCrLf

    
    dc.SaveAs2 filename:= _
        targetName, FileFormat:=wdFormatXMLDocument, _
        LockComments:=False, Password:="", AddToRecentFiles:=True, _
        WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, CompatibilityMode:=wdWord2013

    cleanFile dc
    guard dc
    setSpProp dc, "guide", "OK"
    setProp dc, "hash", BASE64SHA1(dc.name & "_" & Format(Now, "yyyy-mm-ddThh:mm:ss"))
    setSpProp dc, "Cluster Start", "1"
    setSpProps dc
    'dc.Activate
    dc.ActiveWindow.View.ShowHiddenText = False

    If dc.CanCheckin Then
        dc.CheckInWithVersion SaveChanges:=True, MakePublic:=True, _
           Comments:="Migrate to SPO content management", VersionType:=wdCheckInMajorVersion
     Else
        resultsForm.Append "File checkin failed" & vbCrLf
        dc.Close SaveChanges:=wdDoNotSaveChanges
     End If
     GoTo endfunction
fileErr:
    If dc Is Nothing Then
        resultsForm.Append "File failed to open" & vbCrLf
    Else
        dc.Close SaveChanges:=wdDoNotSaveChanges
    End If
endfunction:
    Set dc = Nothing
    resultsForm.show
    Application.ScreenUpdating = False
End Sub
 



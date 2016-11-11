Attribute VB_Name = "UserIf"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim displayMode As Boolean
Dim xMode As Boolean
Dim apparition As CcApparition
Public Log As StringBuilder



Sub setBgnd()
    ActiveDocument.Background.Fill.ForeColor.ObjectThemeColor = _
        wdThemeColorAccent6
    ActiveDocument.Background.Fill.ForeColor.TintAndShade = 0.7
    ActiveDocument.Background.Fill.Visible = msoTrue
    ActiveDocument.Background.Fill.Solid
    ActiveDocument.ActiveWindow.View.DisplayBackgrounds = True
End Sub

Sub showAll(dMode As Boolean)
    displayMode = dMode
    ActiveWindow.View.ShowHiddenText = displayMode
    ActiveWindow.View.ShowBookmarks = displayMode
    ActiveWindow.View.ShowFieldCodes = displayMode
    ActiveWindow.ActivePane.View.showAll = displayMode
End Sub

Sub showArticleMeta(show As Boolean) ' show tags for just one article
    Dim para As Paragraph
    Dim kmj As Object
    Dim r As Range
    ActiveWindow.View.ShowHiddenText = True
    Set r = muEdit.expandArticle(Selection.Range, kmj)
    
    If r Is Nothing Then
        MsgBox ("Cursor is not currently inside an article")
    End If
    
    muEdit.markMarkup r, show
    
    'TODO make the interaction of the hide/show tags and show/hide all buttons better
    If show Then
        showAll displayMode
        ActiveWindow.View.ShowBookmarks = True
        ActiveWindow.View.ShowHighlight = True
        'ActiveWindow.View.ShowFieldCodes = True
    Else
        ActiveWindow.View.ShowHiddenText = False
        ActiveWindow.View.ShowBookmarks = True
    End If
    
End Sub


Sub showMeta(show As Boolean)
    Dim para As Paragraph
    Dim Sel As Range
    Set Sel = Selection.Range
    'TODO should be using set retrieval mode rather than this as its probably quicker
    ActiveWindow.View.ShowHiddenText = True
    For Each para In ActiveDocument.Paragraphs
        muEdit.markArticle para, show
    Next para
    muEdit.markMarkup ActiveDocument.Range, show
    
    'TODO make the interaction of the hide/show tags and show/hide all buttons better
    If show Then
        showAll displayMode
        ActiveWindow.View.ShowBookmarks = True
        ActiveWindow.View.ShowHighlight = True
        'ActiveWindow.View.ShowFieldCodes = True
    Else
        ActiveWindow.View.ShowHiddenText = False
        ActiveWindow.View.ShowBookmarks = True
        ActiveWindow.View.ShowFieldCodes = False
        ActiveWindow.ActivePane.View.showAll = False
    End If
 
    
    Sel.Collapse wdCollapseEnd
    Sel.Select
    
End Sub

Private Function saveArticle(rng As Range, Optional replace As Boolean = False, Optional seq As Integer = 999) As Object
     Dim kmj As Object
     Set rng = muEdit.expandArticle(rng, kmj)
     
     If rng Is Nothing Then
        Set res = CreateObject("Scripting.Dictionary")
        res("status") = 444
        res("statusText") = "Error: Selected range does not contain an article"
        res("responseText") = ""
        Set saveArticle = res
        Exit Function
     End If

     If kmj Is Nothing Then
        Set res = CreateObject("Scripting.Dictionary")
        res("status") = 444
        res("statusText") = "Error: Invalid article structure found - not saved"
        res("responseText") = "{}"
        Set saveArticle = res
        Exit Function
     Else

        
        kmj("markup") = markup.markup(rng)

        Select Case kmj("type")

        Case Is = "article"
            Select Case kmj("purpose")
                Case Is = "landing"
                Case Is = "legislation"
                Case Else
                    kmj("purpose") = "article"
            End Select
        Case Is = "class"
            kmj("type") = "item"
            kmj("purpose") = "class"
       
        Case Is = "SDLT"
            kmj("type") = "item"
            kmj("purpose") = "sdlt"
         
        Case Is = "item"
            kmj("purpose") = "item"
            
        Case Is = "snippet"
            
            
        Case Else
            MsgBox "invalid type major problem <" & kmj("type") & ">"
            res("status") = 444
            res("statusText") = "Error: invalid type major problem <" & kmj("type") & ">"
            res("responseText") = ""
            Set saveArticle = res
            Exit Function
     
        End Select
        
        
        kmj("id") = cctxt.cleanUID(kmj("id"))
        kmj("master")("version") = "0"
        On Error Resume Next ' in case variable doesn't exist
        kmj("master")("lastmodified") = Format(ActiveDocument.BuiltInDocumentProperties("Last save time"), "yyyy-mm-ddThh:mm:ss.000Z")
        kmj("master")("version") = ActiveDocument.Variables("VersionId")
        On Error GoTo 0
        
        pri = seq

        For Each c In kmj("clusters")
           c("priority") = pri
           pri = 999 ' only give first cluster a priority
        Next c

        ' remove old schema remnants if there
        If kmj.Exists("cluster") Then
            kmj.Remove "cluster"
        End If
        
        
        'update if not already set
        If kmj("owner") = "" Then
            muEdit.setPeople kmj
        End If
        'save2Repo kmj
        Set saveArticle = save2Repo(kmj)
        'Select Case fileSave(kmj, replace)
        '    Case 0
        '        saveArticle = "Error: save failed <" & kmj("id") & ".kmj>"
        '    Case 1
        '        saveArticle = "<" & kmj("id") & ".kmj> saved"
        '    Case 2
        '        saveArticle = "<" & kmj("id") & ".kmj> replaced"
        'End Select
     End If
End Function



Public Sub saveJson()
    Dim r As Range
    Dim res As Object
    
    Set r = Selection.Range
    Set res = saveArticle(r)
    If Not (res Is Nothing) Then
        MsgBox res("statusText") & vbCrLf & res("responseText")
    End If
    
    If Not (r Is Nothing) Then
        r.Select
    End If
End Sub

Public Sub saveAllArticles(Optional Log = Nothing)
    Dim rng As Range
    Dim b As Boolean
    Dim results As New StringBuilder
    Dim replace As Boolean
    Dim seq As Integer
    Dim res As Object

    
    
    ' TODO this was from when it wrote to files I don't think this works with GIT
    'replace = (vbNo = MsgBox("do you wish to be prompted before replacing existing articles", vbYesNo))
     replace = True
     
    Set rng = ActiveDocument.Range

    b = True
    seq = 1
    
    If Log Is Nothing Then
        resultsForm.setString "Export all articles" & vbCrLf & vbCrLf
        resultsForm.show
    End If
    
    Do While b
         With rng.Find
            .ClearFormatting
            .MatchWildcards = True
            .text = "\{article:\{s"
            .Execute Forward:=True, Wrap:=False
         End With
         If rng.Find.Found Then
            Set res = saveArticle(rng.Duplicate, replace, seq)
            If Log Is Nothing Then
                resultsForm.Append res("statusText") & vbCrLf & res("responseText")
            Else
                resultsForm.Append "--" & res("id") & ":" & res("statusText")
                results.Append res("statusText") & vbCrLf & res("responseText")
            End If
            seq = seq + 1
            resultsForm.Append vbCrLf
         Else
            b = False
        End If
        Loop
   If Log Is Nothing Then
       resultsForm.Append "Export complete !!" & vbCrLf & vbCrLf
       resultsForm.show
       saveLog cleanFilename(), resultsForm.text
   Else
       resultsForm.show
       saveLog cleanFilename(), results.text
       Log.Append results.text 'todo this should write to the file rather than cache in a string
       Log.Append "Export complete !!" & vbCrLf & vbCrLf
   End If
   
 
   ' ActiveDocument.Close savechanges:=wdDoNotSaveChanges
   'MsgBox "export complete"
End Sub


Sub SaveAllFiles()
   Dim file As String
   Dim fileFilter As String
   Dim path As String
   Dim ad As Document
   Dim dc As Document
   
   Dim reFilter As New RegExp
    With reFilter
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = "^[a-z].*"
    End With
 
    Dim Log As New logger
    'Log.init Cfg.getVar("repo") & "\logs\" & Format(Now, "yyyy-mm-dd-hh-mm-ss") & ".log"
    Log.init "c:\dev\delivered\logs\" & Format(Now, "yyyy-mm-dd-hh-mm-ss") & ".log"
    markup.setLog Log
 
    resultsForm.setString "Export all files" & vbCrLf & vbCrLf
    resultsForm.show
    
    
    Set ad = ActiveDocument
    'path = Cfg.getVar("trove") & "\delivered\"
    path = "c:\dev\delivered\"
    'path = "c:\dev\delivered\snippets\"
For s = asc("a") To asc("z")
    
    fileFilter = path & Chr(s) & "*.doc"
    file = dir(fileFilter)
    While (file <> "")
         If reFilter.test(file) Then
            fildate = CDate(FileDateTime(path & file))
            lastdate = CDate("28/09/2016") ' date will be from midnight so will include this day
            If (fildate > lastdate) Then
                
                'ad.Activate
                'Application.ScreenUpdating = True
                'Application.ScreenRefresh
                Sleep (1000)
                'resultsForm.Append Format(Now(), "yyyy-MM-dd hh:mm:ss") & path & file & vbCrLf
                'Log.Append Format(Now(), "yyyy-MM-dd hh:mm:ss") & path & file & vbCrLf
                Set dc = ad.Application.Documents.Open(path & file, addtorecentfiles:=False, Visible:=True, ReadOnly:=False)
                Sleep (5000)
                'resultsForm.Append Format(Now(), "yyyy-MM-dd hh:mm:ss") & vbCrLf
                dc.Activate
                ActiveDocument.ActiveWindow.View.ShowHiddenText = True
                dc.ActiveWindow.View.ShowHiddenText = True
               
                'saveAllArticles Log

                dc.Close savechanges:=wdDoNotSaveChanges
                Sleep (5000)
                'Set dc = Nothing
                'resultsForm.Append vbCrLf
            End If
            'Log.Flush
         End If
         file = dir
         
    Wend
Next s

   ad.Activate
   resultsForm.Append vbCrLf & "DONE"
   'saveLog "_batch", Log.text
   
   'todo some error handling to make sure this happens
   Set Log = Nothing



End Sub


Sub uImarkup()
    Dim markup As New CcMarkup

    mdForm.tbMd = markup.markup(Selection.Range)
    mdForm.show

End Sub


 
Sub exportImages()
    imgForm.show
End Sub


Sub previewJson()
    Dim markup As New CcMarkup
    Dim kmj As Object
    
    If Not checkSelection(kmj) Or (kmj Is Nothing) Then
        MsgBox "selected range is not an article"
        Exit Sub
    End If
    
    kmj("markup") = markup.markup(Selection.Range)
    For Each el In kmj("extlinks")
        el("display") = el("name")
        el("extlink") = el("url")
    Next el
    

    If Not kmj Is Nothing Then
        doPreview (JsonEncode(kmj))
    End If
End Sub

Sub uiSharepoint()
    Dim markup As New CcMarkup
    Dim kmj As Object
    
    If Not checkSelection(kmj) Or (kmj Is Nothing) Then
        MsgBox "selected range is not an article"
        Exit Sub
    End If
    
    kmj("markup") = markup.markup(Selection.Range)

    If Not kmj Is Nothing Then
        doSharepoint kmj
    End If
End Sub




Sub createDrop()
    hdForm.show
End Sub



Sub outline()
    outlineForm.show
End Sub

Sub displayMeta()
    If setupForm.getType = "snippet" Then
        entityForm.show    'add snippet
    Else
        metaForm.show 'add article
    End If
End Sub

Sub uILink()
    articlelinkForm.show
End Sub



'TODO refactor common code from the next two functions
'Guide Only
Sub extraText()
    Dim rng As Range
    Set rng = Selection.Range
    Set rng = muEdit.wrapExtra(rng)
    If Not rng Is Nothing Then
        rng.Collapse wdCollapseEnd
        rng.Select
    End If
    
End Sub

'Trove Only
Sub hideText()
    Dim rng As Range
    Set rng = Selection.Range
    Set rng = muEdit.wrapHidden(rng)
    If Not rng Is Nothing Then
        rng.Collapse wdCollapseEnd
        rng.Select
    End If
End Sub
 
Sub uIunhide()
    If Not muEdit.unhide(Selection.Range) Then
        MsgBox "Selection does not contain hidden text"
    End If
End Sub

Sub cfgServer()
    serverForm.show
End Sub
 
Sub markInfo(r As Range)
    markForm.ctContent = r.text
    markForm.show
End Sub

 
Sub uISel()
    Dim s As Range
    Set s = muEdit.targetMarkup(Selection.Range)
    If Not s Is Nothing Then
        s.Select
    Else
        Set s = muEdit.targetMarkup(Selection.Range, 1)
    End If
    If Not s Is Nothing Then
        markInfo s
    End If
End Sub

Sub uINext()
    Dim s As Range
    Set s = muEdit.targetMarkup(Selection.Range, 1)
    If Not s Is Nothing Then
        s.Select
    End If
End Sub

Sub uIPrev()
    Dim s As Range
    Set s = muEdit.targetMarkup(Selection.Range, -1)
    If Not s Is Nothing Then
        s.Select
    End If
End Sub
 
 Sub uIHelp()
    helpForm.show
 End Sub
 
 Sub uISetup()
    setupForm.show
 End Sub



Sub uIGhost()
     Dim lineType As String
     Dim Action As String
     Dim r As Range
     Dim kmj As Object
    
     'CHECK VALID RANGE I.E. DOES NOT OVERLAP ARTICLE MARKERS
     Set apparition = Nothing
     Set r = muEdit.expandArticle(Selection.Range, kmj)
     
     If r Is Nothing Then
        MsgBox ("Invalid range could not acquire")
        Exit Sub
     End If

     'MARK APPARITION
     Set apparition = muEdit.createGhost(Selection.Range)
End Sub

Sub uIApparate()
    If apparition Is Nothing Then
        MsgBox "Nothing is currently acquired for apparition"
        Exit Sub
    End If
    muEdit.postApparition Selection.Range, apparition
    
End Sub

Sub uIAcquire()
    Dim a As CcApparition
    ghostForm.show
    Set a = ghostForm.getGhost()
    If Not a Is Nothing Then
        Set apparition = a
    End If
End Sub

Sub uiBlock()
    Dim r As Range
    Set r = Selection.Range
    muEdit.bracketSection r, "blk!" & kmStyles(blockIndex)("block")
End Sub

Sub uiSpan()
    Dim r As Range
    Set r = Selection.Range
    If r.Paragraphs.count > 1 Then
        MsgBox ("Selection spans paragraphs please use block formating")
        Exit Sub
    End If
    r.InsertBefore "-!" & kmStyles(spanIndex)("span") & "!-"
    r.InsertAfter "-!:!-"
    r.Select
    muEdit.markMarkup r, True
End Sub

Sub uiNew()
    Dim dlgOpen As FileDialog
    Dim sourceName As String
    Dim targetName As String
    On Error GoTo errlabPath
    Set dlgOpen = Application.FileDialog( _
    FileDialogType:=msoFileDialogOpen)
    dlgOpen.InitialFileName = Cfg.getVar("trove") & "\virgin\"
    dlgOpen.AllowMultiSelect = False
    dlgOpen.Filters.Add "Guidance", "*.doc", 1
    dlgOpen.FilterIndex = 1
    dlgOpen.ButtonName = "Get"
    On Error GoTo errlabPath
    If dlgOpen.show = -1 Then
         On Error GoTo errlabCopy
         nam = getFileName(dlgOpen.SelectedItems(1))
         sourceName = Cfg.getVar("trove") & "\virgin\" & nam
         targetName = Cfg.getVar("trove") & "\wip\" & nam
         'TODO error handling and check if file already in WIP
         Set fso = CreateObject("Scripting.FileSystemObject")
         If fso.FileExists(targetName) Then
            MsgBox "File <" & targetName & "<is already in WIP"
         Else
            fso.CopyFile sourceName, targetName
            If Documents.CanCheckOut(filename:=targetName) Then
                Documents.CheckOut filename:=targetName
            End If
            MsgBox "Your file has been copied to the author repository and checked out to you. " & _
            "Please use the open button to open the file for editing"
         End If
    End If
    Exit Sub
errlabPath:
    MsgBox "error : " & Err.Description & "<" & Cfg.getVar("trove") & "\virgin\" & ">"
    Exit Sub
errlabCopy:
    MsgBox "error : " & Err.Description & "<" & nam & ">"
    Exit Sub
 
End Sub

Sub uiOpen()
'https://technet.microsoft.com/en-us/library/ee692906.aspx
    Dim dlgOpen As FileDialog
    'Application.ChangeFileOpenDirectory Cfg.getVar("wip")
    Set dlgOpen = Application.FileDialog( _
    FileDialogType:=msoFileDialogOpen)
    dlgOpen.AllowMultiSelect = False
    dlgOpen.InitialFileName = Cfg.getVar("trove") & "\wip\" & "*.doc"
    dlgOpen.Filters.Add "Guidance", "*.doc", 1
    dlgOpen.FilterIndex = 1
    dlgOpen.ButtonName = "Edit"
    If dlgOpen.show = -1 Then
        TheCurrentfilename = Cfg.getVar("trove") & "\wip\" & getFileName(dlgOpen.SelectedItems(1))
        On Error GoTo errLab
        If Documents.CanCheckOut(filename:=TheCurrentfilename) Then
            Documents.CheckOut filename:=TheCurrentfilename
        End If
        Dim openedDoc As Document
        
        Set openedDoc = Application.Documents.Open(TheCurrentfilename)
        
        If openedDoc.CanCheckin Then
            If openedDoc.ReadOnly Then
                MsgBox ("Warning document is readonly. Edit at your own risk")
            End If
        Else
             MsgBox ("Warning document is not checked out. Edit at your own risk")
        End If
        openedDoc.Activate
        
    End If

    Exit Sub
errLab:
    MsgBox "Document did not open properly " & Err.Description
End Sub

 
Sub uiReview()
'https://technet.microsoft.com/en-us/library/ee692906.aspx
    Dim dlgOpen As FileDialog
    'Application.ChangeFileOpenDirectory Cfg.getVar("wip")
    Set dlgOpen = Application.FileDialog( _
    FileDialogType:=msoFileDialogOpen)
    dlgOpen.AllowMultiSelect = False
    dlgOpen.InitialFileName = Cfg.getVar("trove") & "\review\" & "*.doc"
    dlgOpen.Filters.Add "Guidance", "*.doc", 1
    dlgOpen.FilterIndex = 1
    dlgOpen.ButtonName = "Review"
    If dlgOpen.show = -1 Then
        TheCurrentfilename = Cfg.getVar("trove") & "\review\" & getFileName(dlgOpen.SelectedItems(1))
        On Error GoTo errLab
        If Documents.CanCheckOut(filename:=TheCurrentfilename) Then
            Documents.CheckOut filename:=TheCurrentfilename
        Else
            MsgBox "Warning: the file could not be checked out possibly because its already checked out by someone else. Any changes you do could be lost."
        End If
        Application.Documents.Open TheCurrentfilename
        setBgnd
    End If
    Exit Sub
errLab:
    MsgBox Err.Description
End Sub
 
 

Function checkRepo(repo As String, pth As String) As Boolean
    nod = getFileName(Cfg.getVar("trove"))
    Dim reRepo As New RegExp
    With reRepo
        .IgnoreCase = True
        .Global = True
        .MultiLine = True
        .Pattern = ".*" & nod & "[\/\\](.*)"
    End With
    Dim m As MatchCollection
    Set m = reRepo.Execute(pth)
    If m.count > 0 Then
        If repo <> "" Then
            If LCase(repo) = LCase(m(0).SubMatches(0)) Then
                checkRepo = True
            Else
                checkRepo = False
            End If
        Else
            checkRepo = True
            repo = m(0).SubMatches(0)
        End If
    Else
        checkRepo = False
    End If
End Function



Sub uiSave()
    On Error GoTo errLab
    Dim repo As String
    Dim pth As String
    pth = ActiveDocument.path

    If Not checkRepo(repo, pth) Then
           MsgBox "currently active file not valid for checkin" & vbCrLf & ActiveDocument.FullName
           Exit Sub
    End If
    
    cluster = setupForm.getCluster()
    If cluster = "" Then
        MsgBox "You cannot checkin until a cluster name has been set"
        Exit Sub
    End If
    
    If repo = "wip" Then
        setProp ActiveDocument, "km_state", "Authoring"
    End If
    
    On Error GoTo errLab
          
    If ActiveDocument.CanCheckin Then
        TheCurrentfilename = Cfg.getVar("trove") & "\" & repo & "\" & ActiveDocument.name
        ActiveDocument.CheckInWithVersion savechanges:=True, MakePublic:=False, _
        Comments:="Guidance Tool", VersionType:=wdCheckInMinorVersion
        
        If Documents.CanCheckOut(filename:=TheCurrentfilename) Then
            Documents.CheckOut filename:=TheCurrentfilename
        Else
            MsgBox "Error checking in document"
        End If
        
        Dim openedDoc As Document
        
        Set openedDoc = Documents.Open(filename:=TheCurrentfilename)
        
        If openedDoc.CanCheckin Then
            If openedDoc.ReadOnly Then
                MsgBox ("Warning document is readonly. Edit at your own risk")
            End If
        Else
             MsgBox ("Warning document is not checked out. Edit at your own risk")
        End If
        openedDoc.Activate
        
        Application.StatusBar = "Document has been checked in and then checked out : " & repo
    
    Else
        ActiveDocument.Save
        Application.StatusBar = "Document has been saved but not checked in : " & repo
        MsgBox "This file was saved but not checked in. This may beecause the file is checked out to someone else. To avoid losing your changes please save a local copy"
    End If
    Exit Sub
errLab:
    MsgBox Err.Description
End Sub


Sub uiCompare()
   Dim ad As Document

   On Error GoTo errLab
   Set ad = ActiveDocument
   nam = ad.name
   'nam = "johnp1.doc" ' ttt
   wipDoc = Cfg.getVar("trove") & "\wip\" & nam
   'Documents.Open wipDoc ' ttt
   'Set ad = ActiveDocument 'ttt
   On Error GoTo 0
   If Not ad.Saved Then
        MsgBox "Document not saved. You need to checkin before you can compare"
        Exit Sub
   End If
   'ad.Close
   Set virgin = Documents.Open(Cfg.getVar("trove") & "\virgin\" & nam)
   virgin.Compare name:=wipDoc, IgnoreAllComparisonWarnings:=True, DetectFormatChanges:=False
   virgin.Close
   ad.Windows(1).WindowState = wdWindowStateMinimize

   
    'Set wip = Documents.Open((Cfg.getVar("trove") & "\wip\" & nam))
   'name:=(Cfg.getVar("trove") & "\wip\" & nam), _
   ' comparetarget:=wdCompareTargetSelected, _
   ' detectFormatchanges:=False
    
   'Application.CompareDocuments OriginalDocument:=virgin, RevisedDocument:=wip
   
   'Application.CompareDocuments OriginalDocument:=Documents(virgin), _
   '     RevisedDocument:=Documents(wip), Destination:= _
   '     wdCompareDestinationNew, Granularity:=wdGranularityWordLevel, _
   '     CompareFormatting:=True, CompareCaseChanges:=True, CompareWhitespace:= _
   '     True, CompareTables:=True, CompareHeaders:=True, CompareFootnotes:=True, _
   '     CompareTextboxes:=True, CompareFields:=True, CompareComments:=True, _
   '     CompareMoves:=True, RevisedAuthor:="John Pickerill", _
   '     IgnoreAllComparisonWarnings:=False
   'ActiveWindow.ShowSourceDocuments = wdShowSourceDocumentsBoth
    Exit Sub
errLab:
    MsgBox "comparison aborted"
End Sub
 
Sub ciComment(TheCurrentfilename, comment As String)
             Dim ad As Document
            If Documents.CanCheckOut(filename:=TheCurrentfilename) Then
                Documents.CheckOut filename:=TheCurrentfilename
            End If
            Set ad = Documents.Open(TheCurrentfilename)
            If comment = "Reviewing" Then
               setBgnd
            End If
            
            setProp ad, "km_state", comment
            ad.Saved = False
            If ad.CanCheckin Then
                ad.CheckInWithVersion savechanges:=True, MakePublic:=False, _
                Comments:=comment, VersionType:=wdCheckInMinorVersion
            Else
                ad.Save
                ad.Close
            End If
End Sub
 
Sub setProp(doc As Document, Prop As String, val As String)
    On Error GoTo addLab:
    doc.CustomDocumentProperties(Prop) = val
addLab:
    On Error Resume Next
    doc.CustomDocumentProperties.Add name:=Prop, value:=val, _
                LinkToContent:=False, Type:=msoPropertyTypeString
End Sub
 
 
Sub uiSubmit()
   'submit for review
    Dim wipDoc As Document
    Dim state As String
    On Error GoTo errLab
    'nam = "johnp1.doc" 'ttt
    'Documents.Open Cfg.getVar("trove") & "\wip\" & nam 'ttt
    Set wipDoc = ActiveDocument

    
    Dim repo As String
    repo = "wip"
    If Not checkRepo(repo, wipDoc.path) Then
           MsgBox "currently active file not valid for review" & vbCrLf & ActiveDocument.FullName
           Exit Sub
    End If
    
    nam = wipDoc.name
    If vbNo = MsgBox("Are you sure you have finished editing and want to submit this document for review?", vbYesNo) Then
        Exit Sub
    End If
    
    TheCurrentfilename = Cfg.getVar("trove") & "\wip\" & nam
    '    Set wipDoc = Documents.Open(filename:=TheCurrentfilename)
    revdoc = Cfg.getVar("trove") & "\review\" & nam
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(revdoc) Then
            If vbNo = MsgBox("File <" & revdoc & ">is already in review. do you wish to resubmit", vbYesNo) Then
                Exit Sub
            End If
            
            If Documents.CanCheckOut(revdoc) Then
                Documents.CheckOut revdoc
            End If
            Documents.Open revdoc
            If ActiveDocument.CanCheckin Then
                ActiveDocument.CheckInWithVersion savechanges:=True, MakePublic:=False, _
                Comments:="Review Aborted", VersionType:=wdCheckInMinorVersion
                Documents.CheckOut revdoc
            Else
                ActiveDocument.Close
                MsgBox "document checked out to someone else, cannot resubmit"
                Exit Sub
            End If
    End If
    

 
    state = "Review"
    If wipDoc.CanCheckin Then
        setProp wipDoc, "km_state", state
        wipDoc.CheckInWithVersion savechanges:=True, MakePublic:=False, _
        Comments:=state, VersionType:=wdCheckInMinorVersion
    Else
        wipDoc.Save
        wipDoc.Close
        ciComment TheCurrentfilename, state
    End If
         
    If Documents.CanCheckOut(filename:=TheCurrentfilename) Then
            Documents.CheckOut filename:=TheCurrentfilename
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile TheCurrentfilename, revdoc
    ciComment revdoc, "Reviewing"
    MsgBox "Your file has been copied to the review repository and the review workflow will be started. "
    Exit Sub
errLab:
    MsgBox Err.Description
End Sub

Sub uiTable()
    muEdit.setTableForm "NONE"
End Sub



Sub uiTodo()
    todofile = Cfg.getVar("trove") & "\ToDo\" & ActiveDocument.name
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(todofile) Then
        Template = Cfg.getVar("trove") & "\ToDo\_todo.doc"
        fso.CopyFile Template, todofile
    End If
    Documents.Open todofile
End Sub

Sub uiSnippet()
    snipForm.show
End Sub

Sub uiAnchor()
    anchorForm.show
End Sub

Public Sub hdConvert()
    Dim cfdx As New cFixDrop
    showMeta True
    cfdx.convert
End Sub

Attribute VB_Name = "UI"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim displayMode As Boolean
Dim xMode As Boolean
Dim apparition As CcApparition
Public log As StringBuilder


Sub uiCleanDoc()
    cleanFile ActiveDocument
End Sub
 
Sub uiUnlock()
    docUnlock ActiveDocument
End Sub


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
    Set r = muEdit.expandArticle(ActiveDocument, Selection.Range, kmj)
    
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




Sub uImarkup()
    Dim markup As New CcMarkup

    mdForm.tbMd = markup.markup(ActiveDocument, Selection.Range)
    mdForm.show

End Sub


 
Sub exportImages()
    imgForm.show
End Sub



 



Sub previewJson()
    Dim markup As New CcMarkup
    Dim kmj As Object
    ActiveWindow.View.ShowHiddenText = True
    
    If Not checkSelection(kmj) Or (kmj Is Nothing) Then
        MsgBox "selected range is not an article"
        Exit Sub
    End If
 
    muEdit.markMarkup Selection.Range, True
    kmj("markup") = markup.markup(ActiveDocument, Selection.Range)
    'For Each el In kmj("extlinks")
    '    el("display") = el("name")
    '    el("extlink") = el("url")
    'Next el
    

    If Not kmj Is Nothing Then
        doPreview (JsonEncode(kmj))
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
     Dim action As String
     Dim r As Range
     Dim kmj As Object
    
     'CHECK VALID RANGE I.E. DOES NOT OVERLAP ARTICLE MARKERS
     Set apparition = Nothing
     Set r = muEdit.expandArticle(ActiveDocument, Selection.Range, kmj)
     
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
    dlgOpen.InitialFileName = Cfg.getVar("docs") & "virgin/"
    dlgOpen.AllowMultiSelect = False
    dlgOpen.Filters.Add "Guidance", "*.doc", 1
    dlgOpen.FilterIndex = 1
    dlgOpen.ButtonName = "Get"
    On Error GoTo errlabPath
    If dlgOpen.show = -1 Then
         On Error GoTo errlabCopy
         nam = getFileName(dlgOpen.SelectedItems(1))
         sourceName = Cfg.getVar("docs") & "virgin/" & nam
         targetName = Cfg.getVar("docs") & "wip/" & nam
         'TODO error handling and check if file already in WIP
         
        On Error Resume Next
        Dim openedDoc As Document
        Application.ScreenUpdating = False
        Set openedDoc = Application.Documents.Open(targetName, ReadOnly:=True)
        If Not openedDoc Is Nothing Then
            openedDoc.Close SaveChanges:=False
            Application.ScreenUpdating = True
            GoTo errlabExists
        End If
        On Error GoTo errlabCopy
        Set openedDoc = Application.Documents.Open(sourceName, ReadOnly:=True)
        openedDoc.SaveAs2 filename:=targetName, FileFormat:=wdFormatDocument
        Application.ScreenUpdating = True
        openedDoc.Activate
        MsgBox "Your file has been copied to the author repository and checked out to you."
        
         
         ' below is for on prem sp2010
         'Set fso = CreateObject("Scripting.FileSystemObject")
         'If fso.FileExists(targetName) Then
         '   MsgBox "File <" & targetName & "<is already in WIP"
         'Else
         '   fso.CopyFile sourceName, targetName
         '   If Documents.CanCheckOut(filename:=targetName) Then
         '       Documents.CheckOut filename:=targetName
         '   End If
         '   MsgBox "Your file has been copied to the author repository and checked out to you. " & _
         '   "Please use the open button to open the file for editing"
         'End If
    End If
    Exit Sub
errlabPath:
    
    MsgBox "error : " & Err.Description & "<" & Cfg.getVar("docs") & "virgin/" & ">"
    Exit Sub
errlabCopy:
    Application.ScreenUpdating = True
    MsgBox "error : " & Err.Description & "<" & nam & ">"
    Exit Sub
errlabExists:
    Application.ScreenUpdating = True
    MsgBox "error : " & "Target file already exists <" & nam & ">"
    Exit Sub
End Sub

Sub uiOpen()
'https://technet.microsoft.com/en-us/library/ee692906.aspx
    Dim dlgOpen As FileDialog
    'Application.ChangeFileOpenDirectory Cfg.getVar("wip")
    Set dlgOpen = Application.FileDialog( _
    FileDialogType:=msoFileDialogOpen)
    dlgOpen.AllowMultiSelect = False
    dlgOpen.InitialFileName = Cfg.getVar("docs") & "wip/" & "*.doc"
    dlgOpen.Filters.Add "Guidance", "*.doc", 1
    dlgOpen.FilterIndex = 1
    dlgOpen.ButtonName = "Edit"
    If dlgOpen.show = -1 Then
        TheCurrentfilename = Cfg.getVar("docs") & "wip/" & getFileName(dlgOpen.SelectedItems(1))
        On Error GoTo errlab
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
errlab:
    MsgBox "Document did not open properly " & Err.Description
End Sub

 
Sub uiReview()
'https://technet.microsoft.com/en-us/library/ee692906.aspx
    Dim dlgOpen As FileDialog
    'Application.ChangeFileOpenDirectory Cfg.getVar("wip")
    Set dlgOpen = Application.FileDialog( _
    FileDialogType:=msoFileDialogOpen)
    dlgOpen.AllowMultiSelect = False
    dlgOpen.InitialFileName = Cfg.getVar("docs") & "review/" & "*.doc"
    dlgOpen.Filters.Add "Guidance", "*.doc", 1
    dlgOpen.FilterIndex = 1
    dlgOpen.ButtonName = "Review"
    If dlgOpen.show = -1 Then
        TheCurrentfilename = Cfg.getVar("docs") & "review/" & getFileName(dlgOpen.SelectedItems(1))
        On Error GoTo errlab
        If Documents.CanCheckOut(filename:=TheCurrentfilename) Then
            Documents.CheckOut filename:=TheCurrentfilename
        Else
            MsgBox "Warning: the file could not be checked out possibly because its already checked out by someone else. Any changes you do could be lost."
        End If
        Application.Documents.Open TheCurrentfilename
        setBgnd
    End If
    Exit Sub
errlab:
    MsgBox Err.Description
End Sub
 
Function checkLibrary(pth As String)
    If InStr(1, pth, "wtools") Then
        checkLibrary = True
        Exit Function
    End If
    checkLibrary = checkRepo(Cfg.getVar("library"), pth)
End Function

Function checkExport(pth As String)
    If InStr(1, pth, "wtools") Then
        checkExport = True
        Exit Function
    End If
    checkExport = checkRepo(Cfg.getVar("export"), pth)
End Function

Function checkRepo(repo As String, pth As String) As Boolean
    ' nod = getFileName(Cfg.getVar("docs"))' this is the 2010 prem version

    nod = Cfg.getVar("docs")
    Dim reRepo As New RegExp
    With reRepo
        .IgnoreCase = True
        .Global = True
        .MultiLine = False
        '.Pattern = ".*" & nod & "[\/\\](.*)" ' this is the 2010 prem version
        .pattern = nod & "(.*)"
    End With
    Dim m As MatchCollection
    Set m = reRepo.Execute(pth)
    If m.count > 0 Then
        If repo <> "" Then
            repo = replace(repo, " ", "%20")
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
    doSave ActiveDocument
End Sub


Sub uiCompare()
   Dim ad As Document

   On Error GoTo errlab
   Set ad = ActiveDocument
   nam = ad.name
   'nam = "johnp1.doc" ' ttt
   wipDoc = Cfg.getVar("docs") & "wip/" & nam
   'Documents.Open wipDoc ' ttt
   'Set ad = ActiveDocument 'ttt
   On Error GoTo 0
   If Not ad.Saved Then
        MsgBox "Document not saved. You need to checkin before you can compare"
        Exit Sub
   End If
   'ad.Close
   Set virgin = Documents.Open(Cfg.getVar("docs") & "virgin/" & nam)
   virgin.Compare name:=wipDoc, IgnoreAllComparisonWarnings:=True, DetectFormatChanges:=False
   virgin.Close
   ad.Windows(1).WindowState = wdWindowStateMinimize

   
    'Set wip = Documents.Open((Cfg.getVar("docs") & "\wip\" & nam))
   'name:=(Cfg.getVar("docs") & "\wip\" & nam), _
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
errlab:
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
                ad.CheckInWithVersion SaveChanges:=True, MakePublic:=False, _
                Comments:=comment, VersionType:=wdCheckInMinorVersion
            Else
                ad.save
                ad.Close
            End If
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



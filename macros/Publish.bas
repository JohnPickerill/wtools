Attribute VB_Name = "Publish"
Private Const wip_branch = "wip"
Private Const adv_branch = "advance"
Private Const live_branch = "live"
Dim filecount As Integer
Dim record  As Object
Const checkpoint_interval = 20

Dim errcodes As Object

Sub setErrcodes()
    Set errcodes = CreateObject("Scripting.Dictionary")
    errcodes.Add 100, "no change"
    errcodes.Add 200, "Updated OK"
    errcodes.Add 201, "Created OK"
    errcodes.Add 441, "Invalid data sent to server"
    errcodes.Add 444, "File failed to open"
    errcodes.Add 445, "Error selecting article"
    errcodes.Add 446, "Guidance flag in wrong state"
    errcodes.Add 447, "Invalid Metadata"
    errcodes.Add 448, "Invalid Object Type"
    errcodes.Add 449, "Invalid Object ID"
End Sub

Function errString(code) As String
    errString = code & " ==> " & errcodes(i)
    
End Function




Sub writeErrorKey()
    For Each i In errcodes
        resultsForm.Append errString(1) & vbCrLf
    Next i

End Sub

Sub checkpoint()
    filecount = filecount + 1
    If filecount > checkpoint_interval Then
        resultsForm.Append vbCRLR & "===== Checkpoint =====" & vbCrLf
        putCheckpoint JsonEncode(record)
        'pushRecord wip_branch, record
        filecount = 0
    End If
End Sub


Sub testit()
    saveAllDirectories False
End Sub


Function syncWip()
    syncWip = syncES(wip_branch)
End Function
    
Function promoteWip()
    promoteWip = promote(wip_branch, adv_branch)
End Function
    
Function promoteAdvance()
    promoteAdvance = promote(adv_branch, live_branch)
End Function
 


Function logDir() As String
    logDir = Cfg.getVar("webDav") & "\logs\"
End Function

Function getNode(rec As Object, node, Optional status) As Object
    If Not rec.Exists(node) Then
        Dim obj As Object
        Set obj = CreateObject("Scripting.Dictionary")

        rec.Add node, obj
    End If
    
    If Not IsMissing(status) Then
        rec(node)("status") = status
    End If
    Set getNode = rec(node)
End Function

Sub setStatus(rec, node)
    If rec(node)("status") > rec("status") Then
        rec("status") = rec(node)("status")
    End If
End Sub
 
Function testFile(fileRecord, lastmodified, filesize)
    testFile = True
    If Not fileRecord.Exists("lastmodified") Then
        fileRecord("lastmodified") = lastmodified
        testFile = False
        Exit Function
    Else
        If fileRecord("lastmodified") <> lastmodified Then
            testFile = False
            Exit Function
        End If
    End If
    
    If Not fileRecord.Exists("size") Then
        fileRecord("size") = filesize
        testFile = False
        Exit Function
    Else
        'the filesize seems to change for no reason ?
        'If fileRecord("size") <> filesize Then
        '    testFile = False
        'End If
    End If
    If fileRecord("status") >= 300 Then
        testFile = False
    End If
        
End Function



Function pretty(js As String) As String
    Dim p As New StringBuilder
    Dim ch As String
    Dim indent As String
    p.clear
    indent = ""
    For Index = 1 To Len(js)
        ch = Mid(js, Index, 1)
        Select Case ch
            Case "{"
                indent = indent & "     "
                p.Append vbCrLf & indent
            Case "}"
                p.Append indent
                indent = Left(indent, Len(indent) - 5)
            Case ","
                p.Append vbCrLf & indent
            Case """"
            Case ":"
                p.Append vbTab & " : "
            Case Else
                p.Append ch
        End Select
    Next Index
    pretty = p.text
End Function

Function flatten(record As Object) As Object
     Set folders = record("folders")
     For Each fold In folders
     If fold <> "status" And fold <> "dir_" Then
            Dim sFold As Object
            Set sFold = getNode(record, fold)
            For Each fil In folders(fold)("files")
                If fil <> "status" Then
                    Set folders("dir_")("files")(fil) = folders(fold)("files")(fil)
                Else
                    folders("dir_")("files")(fil) = folders(fold)("files")(fil)
                End If
            Next fil
            folders.Remove fold
     End If
     
     Next fold
End Function



Function summarise(record As Object, test As Boolean) As Object
     Dim summary As Object
     Set summary = CreateObject("Scripting.Dictionary")
     push_start = record("push_start")
     summary("status") = record("status")
     Set folders = record("folders")
     For Each fold In folders
        If fold = "status" Then
            If Not test Then
                summary("status") = "Last check (" & record("push_start") & ") "
            Else
                summary("status") = "Last publication (" & record("push_start") & ") "
                summary("status") = "Update "
            End If
            If folders("status") >= 300 Then
                summary("status") = summary("status") & "errored " & errString(folders("status"))
            Else
                summary("status") = summary("status") & "ok " & errString(folders("status"))
            End If
        Else
            Dim sFold As Object
            Set sFold = getNode(summary, fold)
            num_docs = 0
            num_errors = 0
            num_updates = 0
            For Each fil In folders(fold)("files")
                If fil <> "status" Then
                    num_docs = num_docs + 1
                    node_exists = False
                    Dim sFile As Object
                    If folders(fold)("files")(fil)("last_checked") < push_start Then
                        num_errors = num_errors + 1
                        Set sFile = getNode(sFold, fil)
                        sFile("last_checked") = folders(fold)("files")(fil)("last_checked")
                        sFile("warning") = "File may be an invalid guidance document or no longer exist"
                    End If
                    
                    Dim filobj As Object
                    Set filobj = folders(fold)("files")(fil)
                    
                    If filobj("export_needed") Then

                        Set sFile = getNode(sFold, fil)
                        sFile("status") = folders(fold)("files")(fil)("status")
                        If folders(fold)("files")(fil)("status") >= 300 Then
                            num_errors = num_errors + 1
                        End If
                        num_updates = num_updates + 1
                        
                        If Not test Then
                                sFile("export_needed") = folders(fold)("files")(fil)("export_needed")
                        End If
                        
                        Set sFile("objects") = CreateObject("Scripting.Dictionary")
                            
                        err_found = False
                        If filobj.Exists("objects") Then
                            For Each obj In folders(fold)("files")(fil)("objects")
                                If obj <> "status" Then
                                    Set thisobj = folders(fold)("files")(fil)("objects")(obj)
                                    If thisobj("status") >= 300 Then
                                        err_found = True
                                        sFile("objects")(obj) = thisobj("status") & " " & thisobj("statusText")
                                    Else
                                        If Not test And thisobj("status") >= 200 Then
                                            sFile("objects")(obj) = thisobj("status") & " " & thisobj("statusText")
                                        End If
                                    End If
                                End If
                            Next obj
                        Else
                            err_found = False
                            sFile("objects")("warning") = "No objects found"
                        End If
                           
                        If test And Not err_found Then
                            sFile.Remove "objects"
                        End If
                           
                    End If
                End If
                    

            Next fil
            sFold("number_of_docs") = num_docs
            sFold("number_of_errors") = num_errors
            sFold("number_of_updates") = num_updates
        End If
     Next fold
     Set summarise = summary
End Function


Function Remove_deletes(record As Object, Optional Remove As Boolean = False) As String
     Dim del_list As New StringBuilder
     Set folders = record("folders")
     push_start = record("push_start")
     For Each fold In folders
        If fold <> "status" Then
            For Each fil In folders(fold)("files")
                If fil <> "status" Then
                    If folders(fold)("files")(fil)("last_checked") < push_start Then
                        If Remove Then
                            folders(fold)("files").Remove fil
                        End If
                        del_list.Append vbCrLf + fil
                    End If
                End If
            Next fil
        End If
     Next fold
     Remove_deletes = del_list.text
End Function



Sub saveAllDirectories(Optional test As Boolean = False)
     Dim path As String
     Dim lastdate As Date
  
     filecount = 1
     
     If Not canDo("publisher") Then
        MsgBox ("You do not have sufficient priviledge for this action")
        Exit Sub
     End If
     
     this_name = ActiveDocument.name
     mgr_name = "Managing Practice Guidance"
     If (InStr(this_name, "kmj.dotm") = 0) And (InStr(this_name, mgr_name) = 0) Then
        MsgBox ("Please start export from document:" & vbCrLf & mgr_name)
        Exit Sub
     End If
     
     
     
     If test Then
        resultsForm.setString "Starting Word Document Check " & vbCrLf
     Else
        resultsForm.setString "Starting Export" & vbCrLf
     End If
     resultsForm.show
     resultsForm.Append "Reading checkpoint file. This may take a few moments ...." & vbCrLf
     'Set record = pullRecord(wip_branch)
     Set record = getCheckpoint()
     'If Not record.Exists("push_start") Then
     '   Set record = flatten(pullRecord(wip_branch))
     'End If
     'putCheckpoint JsonEncode(record)
      
     If Not record.Exists("push_start") Then
        If (vbNo = MsgBox("ERROR getting last checkpoint. Continuing will attempt to load all documents and may take a very very long time. Do you wish to continue ?", vbYesNo)) Then
            Exit Sub
        End If
     End If
     
     
     'Set Record = CreateObject("Scripting.Dictionary")
     If Not test Then
         If record.Exists("push_start") Then
            del_list = Remove_deletes(record, False)
            If del_list <> "" Then
                If vbYes = MsgBox("Do you wish to remove the following deleted files from the checkpoint file ?" & vbCrLf & del_list, vbYesNo) Then
                   Remove_deletes record, True
                End If
            End If
         End If
         record("push_start") = Format(Now(), "yyyy-mm-ddThh:mm:ss.000Z")
     End If
     record("status") = 0
     
     'path = "c:\dev\delivered"
     path = Cfg.getVar("webDav") & "\" & Cfg.getVar("library") ' delivered"
     'path = Cfg.getVar("webDav") & "\devLibrary"
     
     Dim folders As Object
     Set folders = getNode(record, "folders", 0)

     
     'Dim dirs() As String
     'dirs = Split("algorithms,,snippets,items,help", ",")
     'dirs = Split(",,algorithms,snippets,items", ",")
     Dim dirs(0 To 0) As String
     dirs(0) = ""

     For i = 0 To UBound(dirs)
        resultsForm.show
        resultsForm.setString "analysing directory " & dirs(i) & vbCrLf
        SaveAllFiles path, dirs(i), lastdate, getNode(folders, "dir_" + dirs(i), 0), test
        setStatus folders, "dir_" + dirs(i)
        setStatus record, "folders"
     Next i
 
     
     If Not test Then
        pushRecord wip_branch, record
        putCheckpoint JsonEncode(record)
     End If
     
     
     If test Then
        resultsForm.setString "Word Document Check " & vbCrLf
     Else
        resultsForm.setString "Export complete" & vbCrLf
     End If
     
     Dim Summarystr As String
     Dim SummaryJson As Object
     Set SummaryJson = summarise(record, test)
     Summarystr = pretty(JsonEncode(SummaryJson))
     putExportLog Summarystr
     resultsForm.Append Summarystr
     resultsForm.Append vbCrLf & "== Complete ==" & vbCrLf
     writeErrorKey
End Sub




Sub SaveAllFiles(path As String, fold As String, lastdate As Date, dirRecord As Object, _
                Optional test As Boolean = False, Optional pattern As String = "[0-9,a-z]")

   Dim file As String
   Dim fileFilter As String
   Dim filesRecord As Object
   Dim fileRecord As Object
   
   Dim filename As String

   Dim ad As Document

   Dim folder As String
   
   folder = "\" & fold & "\"

   Set filesRecord = getNode(dirRecord, "files", 0)
 
   If Not canDo("publisher") Then
        MsgBox ("You do not have sufficient priviledge for this action")
        Exit Sub
   End If
   
   
   Dim reFilter As New RegExp
    With reFilter
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "^" + pattern + ".*docx?"
        '.Pattern = "exped.*"
    End With
    
    Set ad = ActiveDocument

    For s = asc("0") To asc("Z")
        If (s <= asc("9") Or s >= asc("A")) Then
            fileFilter = path & folder & Chr(s) & "*.doc*"
            On Error GoTo fileError
            file = dir(fileFilter)
            On Error GoTo 0
            While (file <> "")
                 Set fileRecord = getNode(filesRecord, file) ' don't reset status because need to test ' note json gets into issues with \ in data
                 saveFile path, folder, file, fileRecord, test
                 ad.ActiveWindow.Visible = True
                 Application.ScreenUpdating = True
                 setStatus filesRecord, file
                 file = dir
            Wend
        End If
    Next s
    setStatus dirRecord, "files"
    ad.Activate
    Exit Sub
fileError:
    MsgBox ("Cannot find sharepoint folder, please open in File Explorer from Internet Explorer ")
    Err.Raise 555, "Sharepoint", "Error reading folder"
End Sub

 
Sub test_saveFile()
    Dim test_rec As Object
    Set test_rec = JsonDecode("{}")
    Dim fileRecord As Object
    Dim file As String
    Dim folder As String
    Dim path As String
    
    path = Cfg.getVar("webDav") & "\devLibrary"
    folder = "\"
    file = "johnp.docx"
    Set fileRecord = getNode(test_rec, file)
    saveFile path, folder, file, fileRecord, False
End Sub
 


Sub saveFile(path As String, folder As String, file As String, _
    fileRecord As Object, Optional test As Boolean = False)

    Dim dc As Document

    filename = path & folder & file
    'fildate = Format(Format(FileDateTime(filename), "yyyy-mm-ddThh:mm:ss.000Z"))
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filename)
    fildate = Format(f.DateLastModified, "yyyy-mm-ddThh:mm:ss.000Z")
    filsize = FileLen(filename)
    export_needed = Not testFile(fileRecord, fildate, filsize)

    Application.ScreenUpdating = True
    resultsForm.Append fildate & vbTab & folder & file
    If export_needed Then
         resultsForm.Append vbTab & " =>export needed " & vbCrLf
    Else
        resultsForm.Append vbCrLf
    End If
    
    fileRecord("export_needed") = export_needed
 
    fileRecord("last_checked") = Format(Now(), "yyyy-mm-ddThh:mm:ss.000Z")
        
    If Not test Then
        fileRecord("status") = 0
        If export_needed Then
            checkpoint

            resultsForm.Append vbTab & "Document has been updated - opening" & vbCrLf
            Application.ScreenRefresh
            On Error GoTo fileErr
            Set dc = Application.Documents.Open(filename, AddToRecentFiles:=False, Visible:=True, ReadOnly:=True)
            On Error GoTo 0
            ' ActiveDocument.TrackRevisions = False ' file is readonly so can't do this
            With ActiveWindow.View.RevisionsFilter
                .markup = wdRevisionsMarkupNone
                .View = wdRevisionsViewFinal
            End With

            dc.ActiveWindow.Visible = False
            Dim guide As String
            guide = getProp(dc, "guide")
            If guide = "OK" Then
                resultsForm.Append vbTab & "Pushing objects" & vbCrLf
                Sleep (1000)
                saveAllArticles dc, fileRecord, fildate
                fileRecord("lastmodified") = fildate
                fileRecord("size") = filsize
                fileRecord("last_analysed") = Format(Now(), "yyyy-mm-ddThh:mm:ss.000Z")
            Else
                resultsForm.Append vbTab & "File ignored : Guidance flag in wrong state = " & guide & vbCrLf
                Sleep (1000)
                fileRecord("status") = 446
           End If
fileErr:
            
            If dc Is Nothing Then
                fileRecord("status") = 444
                resultsForm.Append vbTab & "File failed to open" & vbCrLf
            Else
                dc.Close SaveChanges:=wdDoNotSaveChanges
            End If
            Set dc = Nothing
            Sleep (1000)
        End If
    End If
    resultsForm.show
    Application.ScreenUpdating = False
End Sub







Sub saveAllArticles(dc As Document, fileRecord As Object, lastmodified)
    Dim rng As Range
    Dim b As Boolean
    Dim results As New StringBuilder
    Dim seq As Integer
    Dim res As Object
    Dim s As String
    Dim objectsRecord As Object
    Dim objRecord As Object
    Dim errorRecord As Object
    
    If Not canDo("publisher") Then
        MsgBox ("You do not have sufficient priviledge for this action")
        Exit Sub
    End If
    

    
    
    
    Set objectsRecord = getNode(fileRecord, "objects", 0)
    On Error GoTo errorLab
        
    For Each obj In objectsRecord    ' remove any existing parsing errors from record
        If obj <> "status" Then
            Select Case objectsRecord(obj)("status")
            Case Is = 441, 442, 443, 444, 445, 446, 447, 448, 449
                objectsRecord.Remove obj
            End Select
        End If
    Next obj
        
        
    markup.setLog results
    muEdit.setLog results
    
    dc.Activate
    dc.ActiveWindow.Activate
    
    ' comment
    With ActiveWindow.View.RevisionsFilter
        .markup = wdRevisionsMarkupNone
        .View = wdRevisionsViewFinal
    End With
    ' we need to be sure that the markup isn't hidden

     
    ' need to show hidded text so that we can find article tags
    asWas = ActiveWindow.View.ShowHiddenText
    If Not asWas Then
        ActiveWindow.View.ShowHiddenText = True
    End If
    
    Set rng = dc.Range
    muEdit.markMarkup rng, True
    
    b = True
    seq = 1
    
    Do While b
         With rng.Find
            .ClearFormatting
            .MatchWildcards = True
            .text = "\{article:\{s"
            .Execute Forward:=True, Wrap:=False
         End With
         If rng.Find.Found Then
            Set res = saveArticle(dc, rng.Duplicate, seq)
            Set objRecord = getNode(objectsRecord, res("id"), 0)
            If res("status") <> 200 Then
                objRecord("status") = res("status")
                objRecord("statusText") = res("statusText")
                stat = res("status")
                setStatus objectsRecord, res("id")
            Else
                objRecord("type") = res("type")
                stat = res("response")("status")
                objRecord("status") = res("response")("status")
                objRecord("statusText") = res("statusText")
                setStatus objectsRecord, res("id")
                If res("response")("status") < 200 Then
                    If Not objRecord.Exists("lastmodified") Then
                        objRecord("lastmodified") = lastmodified ' need to get rid of this
                    End If
                    If objRecord.Exists("errors") Then
                        objRecord.Remove "errors"
                    End If
                ElseIf res("response")("status") < 300 Then
                    objRecord("lastmodified") = lastmodified
                    If objRecord.Exists("errors") Then
                        objRecord.Remove "errors"
                    End If
                Else
                    Set errorRecord = getNode(objRecord, "errors", 0)
                    For Each obj In res("response")("result")
                            'If obj("status") >= 300 Then
                                       errorRecord.Add "d", obj
                            'End If
                    Next obj
                End If
            End If
            setStatus objectsRecord, res("id")
            resultsForm.Append vbTab & vbTab & res("id") & " - " & objRecord("statusText") & " - " & objRecord("status") & vbCrLf

            seq = seq + 1
         Else
            b = False
        End If
        Loop

GoTo cleanupLab
errorLab:
    MsgBox "unexpected error in " & dc.name & " press OK to continue"

cleanupLab:
   setStatus fileRecord, "objects"
   resultsForm.show
       
   If ActiveWindow.View.ShowHiddenText <> asWas Then
        ActiveWindow.View.ShowHiddenText = asWas
    End If
End Sub


Private Function saveArticle(dc As Document, rng As Range, seq As Integer) As Object
    If Not canDo("publisher") Then
        MsgBox ("You do not have sufficient priviledge for this action")
        Exit Function
    End If
     
    Dim kmj As Object
    rng_start = rng.start
    Set rng = muEdit.expandArticle(dc, rng, kmj)
     
    If rng Is Nothing Then
        Set res = CreateObject("Scripting.Dictionary")
        res("id") = "error_expanding_article at " & str(rng_start)
        res("type") = "unknown"
        res("status") = 445
        res("statusText") = "Error: Selected range does not contain an article"
        res("response") = "{}"
        Set saveArticle = res
        Exit Function
    End If

    If kmj Is Nothing Then
        Set res = CreateObject("Scripting.Dictionary")
        res("id") = "error_at_position" + str(rng_start)
        res("type") = "unknown"
        res("status") = 447
        res("statusText") = "Error: Invalid article structure found - not saved"
        res("response") = "{}"
        Set saveArticle = res
        Exit Function
    End If

    
 
    If kmj("id") = "enteruniqueid" Or kmj("id") = "" Then
        Set res = CreateObject("Scripting.Dictionary")
        res("id") = "error_at_position" + str(rng_start)
        res("type") = "unknown"
        res("status") = 449
        res("statusText") = "Error: Invalid article id <" + kmj("id") + ">"
        res("response") = "{}"
        Set saveArticle = res
        Exit Function
    End If
    
    kmj("markup") = markup.markup(rng)
        Select Case kmj("type")

        Case Is = "article", "Article"
            kmj("type") = "article"
            Select Case kmj("purpose")
                Case Is = "landing"
                Case Is = "legislation"
                Case Is = "help"
                Case Is = "glossary"
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
            Set res = CreateObject("Scripting.Dictionary")
            res("id") = "error_at_position" + str(rng_start)
            res("type") = "unknown"
            res("status") = 448
            res("statusText") = "Error: invalid type major problem <" & kmj("type") & ">"
            res("response") = "{}"
            Set saveArticle = res
            Exit Function
     
        End Select
        
        
        kmj("id") = cctxt.cleanUID(kmj("id"))
        kmj("master")("version") = "0"
        On Error Resume Next ' in case variable doesn't exist
        kmj("master")("lastmodified") = Format(dc.BuiltInDocumentProperties("Last save time"), "yyyy-mm-ddThh:mm:ss.000Z") ' time in ISO 8601 format
        'kmj("master")("version") = dc.Variables("VersionId")
        kmj("master")("version") = dc.BuiltInDocumentProperties("Revision number")
        On Error GoTo 0
        
        pri = seq
        For Each c In kmj("clusters")
           c("priority") = pri
           pri = 1000000 + seq ' secondary clusters have a lower priority
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
        kmj("commit_comment") = kmj("master")("change")
        kmj("master")("where") = dc.path
        kmj("master")("filename") = ActiveDocument.name
        Set res = pushObject(wip_branch, kmj)
        'res("markup") = JsonEncode(kmj)
         res("type") = kmj("type")
    
        Set saveArticle = res
        

End Function


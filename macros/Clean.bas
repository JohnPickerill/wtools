Attribute VB_Name = "Clean"
Sub testClean()
    cleanfile ActiveDocument
End Sub
 
 
 
Sub cleanfile(dc As Document)
    dc.Activate
    With ActiveWindow.View.RevisionsFilter
        .markup = wdRevisionsMarkupNone
        .View = wdRevisionsViewFinal
    End With
    
    'delete hidden tables
    'naturalise apparitions
    'delete trove only
    'naturalise guide only
    'remove ghosts
    'remove unneeded metadata ?

    ActiveWindow.View.ShowHiddenText = True
    cleanMarkup dc.Range
    removeGhosts dc
End Sub

Sub removeGhosts(ByVal doc As Document)
    Dim app As New CcApparition
    app.clear doc
    Do While app.getApparition("")
        app.delete
        Loop
End Sub



Sub cleanMarkup(r As Range)
    Dim rng As Range
    Dim b As Boolean
    
    On Error GoTo errlab
    ActiveDocument.TrackRevisions = True
    
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
                rng.Select
                'markit rng, True
                cleanRange rng
            End If
         End If
         
         Loop
    Exit Sub
errlab:
    Resume Next
End Sub

Private Sub cleanRange(ByVal r As Range)
    For Each m In muEdit.splitMarkup(r.text)
        
        leadout = Len(m.SubMatches(4))
        leadin = Len(m.SubMatches(0))
        act = m.SubMatches(1)
        
        Select Case act
            Case "+"  ' Guide only, end of hotdrops and links
                 Dim str As String
                 str = m.SubMatches(3)
                 If Left(str, 5) = "drop!" Then   ' hotdrop
                    r.text = replace(r.text, "<*<", "<:<", 1, 1)
                 Else
                    r.text = m.SubMatches(3)
                    r.Collapse wdCollapseEnd
                    r.Font.Hidden = False
                 End If
            Case "-"  ' Trove Only
                 r.text = ""
                 r.Collapse wdCollapseEnd
            Case "&" ' commands e.g. Apparate, table
                Dim com As String
                Dim para As String
                markup.extract_command m.SubMatches(3), com, para
                'If com = "table" Then tableForm = para ' insert formating instruction for next table
                If com = "apparition" Then ' apparate a ghost at this point
                    Dim app As Range
                    Set app = markup.get_apparition(para)
                    app.Copy
                    r.Paste
                    r.Collapse wdCollapseStart
                End If
            Case "~" ' Ghost contents
                 r.text = m.SubMatches(3)
                 r.Font.Hidden = False
                 r.Collapse wdCollapseEnd
                 
                
            Case ":" 'This is for block starts like dropdowns
            Case "*" 'start of block e.g. ghost
                  ' should find ghost id check if its apparated anywhere and if it isn't delete it
        End Select
    Next m
End Sub




Sub cleanAllDirectories()
     Dim source As String
     Dim target As String
 
     source = Cfg.getVar("webDav") & "\" & "delivered"
     target = Cfg.getVar("webDav") & "\" & "content"
     
     Dim dirs() As String
     dirs = Split(",algorithms,snippets,items,help", ",")
 

     For i = 0 To UBound(dirs)
        resultsForm.show
        resultsForm.setString "analysing directory " & dirs(i) & vbCrLf
        copyAllFiles source, target, dirs(i)
     Next i
 
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
        .pattern = "^" + pattern + ".*docx?"
        '.Pattern = "exped.*"
    End With
    
    Set ad = ActiveDocument

    For s = asc("0") To asc("Z")
        If (s <= asc("9") Or s >= asc("A")) Then
            fileFilter = source & folder & Chr(s) & "*.doc*"
            On Error GoTo fileError
            file = dir(fileFilter)
            On Error GoTo 0
            While (file <> "")
                 saveFile source, target, folder, file
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
    




Sub saveFile(source As String, target As String, folder As String, file As String)

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
  
    cleanfile dc

    resultsForm.Append file & vbCrLf
    dc.SaveAs2 filename:= _
        targetName, FileFormat:=wdFormatXMLDocument, _
        LockComments:=False, Password:="", AddToRecentFiles:=True, _
        WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, CompatibilityMode:=15
    setSpProps dc
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
 

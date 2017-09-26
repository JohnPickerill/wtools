Attribute VB_Name = "kmExport"
Function CheckpointName() As String
    url = Cfg.getVar("appURL")
    If InStr(url, "//") Then
        url = Mid(url, InStr(url, "//") + 2)
    End If
    server = Split(url, ".")(0)
    CheckpointName = Cfg.getVar("webDav") & "\" & Cfg.getVar("checkpoint") & "\" & "_" & server & "_checkpoint.txt"
End Function

Function ExportLogName() As String
    url = Cfg.getVar("appURL")
    If InStr(url, "//") Then
        url = Mid(url, InStr(url, "//") + 2)
    End If
    server = Split(url, ".")(0)
    ExportLogName = Cfg.getVar("webDav") & "\" & Cfg.getVar("checkpoint") & "\" & "_" & server & "_export_log.txt"
End Function

Sub putExportLog(log)
    Dim fso As New FileSystemObject
    On Error GoTo errorLabel
    Set fso = CreateObject("Scripting.FileSystemObject")
    fn = ExportLogName
    Set chkpt = fso.CreateTextFile(fn, True)
    chkpt.WriteLine (log)
    chkpt.Close
    Exit Sub
errorLabel:
    MsgBox Err.Description & " : <" & fn & ">"
End Sub


Function putCheckpoint(checkpoint As String) As Boolean
    Dim fso As New FileSystemObject
    On Error GoTo errorLabel
    Set fso = CreateObject("Scripting.FileSystemObject")
    fn = CheckpointName
    Set chkpt = fso.CreateTextFile(fn, True)
    chkpt.WriteLine (checkpoint)
    chkpt.Close
    putCheckpoint = True
    Exit Function
errorLabel:
    putCheckpoint = False
    MsgBox Err.Description & " : <" & fn & ">"
 End Function
 
Sub testget()
    Set x = getCheckpoint
End Sub

 
Function getCheckpoint() As Object
    Dim fso As New FileSystemObject
    On Error GoTo errorLabel
    Set fso = CreateObject("Scripting.FileSystemObject")
    fn = CheckpointName
    Set chkpt = fso.OpenTextFile(fn, ForReading)
    record = chkpt.ReadLine
    Set getCheckpoint = JsonDecode(record)
    chkpt.Close
    Exit Function
errorLabel:
    Set getCheckpoint = JsonDecode("{}")
 End Function



 
 Function saveSharepoint(name As String, txt As String) As Boolean
    Dim fso As New FileSystemObject
    Dim fld As folder
    Dim txts As TextStream
    Dim DD As String
    saveSharepoint = True
    
    On Error GoTo errorLabel
    Set fso = CreateObject("Scripting.FileSystemObject")
    fff = Cfg.getVar("SPsite") & "\SiteAssets\" & name & ".html"


          
    Set txts = fso.CreateTextFile(fff, True)
    txts.WriteLine (txt)
    txts.Close
    Set txts = Nothing
    Set fso = Nothing
    Exit Function
errorLabel:
    saveSharepoint = False
    MsgBox Err.Description & " : <" & fff & ">"
End Function
 
 
 
 



Public Function imgLocation() As String
    'construct web
    imgLocation = Cfg.getVar("images") & "images/" & cleanFilename()
    ' convert to webdav
    Dim rexp As New RegExp
    With rexp
        .Global = True
        .IgnoreCase = True
        .pattern = "\S*https?:\/\/([^\/]*)\/(\S*)"
    End With
    Dim m As MatchCollection
    
    Set m = rexp.Execute(imgLocation)
    If m.count = 1 Then
        imgLocation = "\\" & m.item(0).SubMatches(0) & "@SSL\DavWWWRoot\" & replace(m.item(0).SubMatches(1), "/", "\")
    End If

End Function



Sub extractImages()
    Dim strFileName As String
    Dim strBasePath As String
    Dim strFolderName As String
    Dim strDocname As String
    

    Dim doc As Document
    Dim fso As Object
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    
    Set doc = Nothing
 
    strDocname = cleanFilename()
  
    ' Test if Activedocument has previously been saved
    If ActiveDocument.path <> "" Then
        Set doc = Documents.Add(ActiveDocument.FullName)
    End If

     If doc Is Nothing Then
        MsgBox "The current document must be saved at least once."
        Exit Sub
     End If
 
    
    strBasePath = "C:\Temp\"

    strFolderName = strBasePath & "KM_img_" & strDocname
     'Delete the folder if it exists
    On Error Resume Next
    Kill strFolderName & "_files\*"  'Delete all files
    RmDir strFolderName & "_files" 'Delete folder
    On Error GoTo 0

    'Save in HTML format
    tempname = strFolderName & ".html"
    doc.SaveAs2 filename:=tempname, FileFormat:=wdFormatHTML
    doc.Close
    
    
    On Error Resume Next
    Kill strFolderName & ".html"
    Kill strFolderName & "_files\*.xml"
    Kill strFolderName & "_files\*.html"
    Kill strFolderName & "_files\*.thmx"
    Kill strFolderName & "_files\*.mso"
    
    On Error GoTo errlab
    fso.CopyFolder source:=strFolderName & "_files", Destination:=imgLocation()

    GoTo cleanup
errlab:
    MsgBox Err.Description & "Source:<" & strFolderName & "_files> Destination:<" & imgLocation() & ">"
cleanup:
    On Error Resume Next
    Kill strFolderName & "_files\*"  'Delete all files
    RmDir strFolderName & "_files" 'Delete folder
End Sub

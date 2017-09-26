VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private config As Object

Private Sub Class_Initialize()
    ' the following will happen when word starts
    cfgRead
    setErrcodes
End Sub

Public Function getVar(p As String) As String
    If config Is Nothing Then Class_Initialize
    If config.Exists(p) Then
        getVar = Trim(config(p))
    End If
End Function

Public Sub setVar(p As String, v As String)
    config(p) = Trim(v)
End Sub


Public Function serverRead(server) As Boolean

    If server = "" Then GoTo clearLab
    Set s_config = getConfig(server)
 
    On Error GoTo clearLab
    config("sharepoint") = s_config("sharepoint")
    config("site") = s_config("site")
    config("library") = s_config("library")
    config("images") = s_config("static")
    config("checkpoint") = s_config("checkpoint")
    serverRead = True
    Exit Function
clearLab:
    serverRead = False
End Function



Function cfgRead() As Boolean
    Dim fso As New FileSystemObject
    Dim kmjStream As TextStream
    cfgRead = True
    
    On Error Resume Next
    cfgfile = LCase(Environ("APPDATA")) + "\curiouscrab\guidance\guidance.txt"
      
    On Error GoTo errorLabel
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.GetFile(cfgfile).Size > 0 Then
        Set kmjStream = fso.OpenTextFile(cfgfile)
        jsonstr = kmjStream.ReadAll
        kmjStream.Close
        Set config = JsonDecode(jsonstr)
        If config Is Nothing Then
            GoTo errorLabel
        End If
    Else
        GoTo errorLabel
    End If
    

    If config.Exists("cfgURL") Then
        If Not serverRead(config("cfgURL")) Then GoTo errorLabel
    Else
        GoTo errorLabel
    End If

    
    Set kmjStream = Nothing
    Set fso = Nothing
    Exit Function
errorLabel:
    Set kmjStream = Nothing
    Set fso = Nothing
    Set config = CreateObject("Scripting.Dictionary")
    cfgRead = False
    MsgBox "Please provide a valid server configuration for the guidance toolbar"
End Function




Function cfgSave() As Boolean
    Dim fso As New FileSystemObject
    Dim kmjStream As TextStream
    Dim cfgfile As String
    cfgSave = True
    
    On Error Resume Next
    cfgfile = LCase(Environ("APPDATA")) + "\curiouscrab"
    MkDir cfgfile
    cfgfile = cfgfile + "\guidance"
    MkDir cfgfile
    cfgfile = cfgfile + "\guidance.txt"
   
    
    On Error GoTo errorLabel

    jsonstr = JsonEncode(config)
    Set fso = CreateObject("Scripting.FileSystemObject")
        
    Set kmjStream = fso.CreateTextFile(cfgfile, True)
    kmjStream.WriteLine jsonstr
    kmjStream.Close
    Set kmjStream = Nothing
    Set fso = Nothing
    Exit Function
errorLabel:
    cfgSave = False
    MsgBox Err.Description & " : <" & cfgfile & ">"
End Function

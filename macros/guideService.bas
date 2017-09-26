Attribute VB_Name = "guideService"

 
Private Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long




Private rbac As Object

Sub testrbac()
    canDo "publisher"
End Sub

Function canDo(action) As Boolean
    On Error GoTo errLabel
    canDo = False
    If rbac Is Nothing Then
        Dim MyRequest As Object
        Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
        url = Cfg.getVar("cfgURL") & "/valid/rbac?wtool=" & kmVer
        MyRequest.Open "GET", url
        MyRequest.setRequestHeader "X-Api-Key", "defaultapikey"
        MyRequest.Option(4) = &H3300  ' ignore certicicate error
        
        ' Send Request.
        MyRequest.send
        'And we get this response
        jsonstr = MyRequest.responseText
        If MyRequest.status <> 200 Then GoTo noauthLabel
        Set rbac = JsonDecode(jsonstr)(LCase(Environ("UserName")))
    End If
    For Each a In rbac
        If action = a Then
             canDo = True
             Exit Function
        End If
    Next a
    Exit Function
noauthLabel:
    MsgBox MyRequest.responseText
    Exit Function
errLabel:
    MsgBox "No Authorization for this action "
End Function


Sub test_ga()
    gaExists "banksxxx"
End Sub

Function gaExists(id) As Boolean
    On Error GoTo errLabel
    gaExists = True

    Dim MyRequest As Object
    Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    url = Cfg.getVar("cfgURL") & "/branches/wip/ga/" & id
    MyRequest.Open "GET", url
    MyRequest.setRequestHeader "X-Api-Key", "defaultapikey"
    MyRequest.Option(4) = &H3300  ' ignore certicicate error
    
    ' Send Request.
    MyRequest.send
    'And we get this response
    jsonstr = MyRequest.responseText
    If MyRequest.status = 404 Then
        gaExists = False
    'Set ga = JsonDecode(jsonstr)
    ElseIf MyRequest.status <> 200 Then
        GoTo errLabel
    End If
    Exit Function
errLabel:
    MsgBox "Could not check whether " & id & " exists. Please try later"

End Function


Function getFacets(purpose) As Object
    Dim facets As Object
    Dim MyRequest As Object
    On Error GoTo errLabel
    
    Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    MyRequest.Open "GET", Cfg.getVar("cfgURL") & "/valid/facets/" & purpose
    MyRequest.setRequestHeader "X-Api-Key", "defaultapikey"
    MyRequest.Option(4) = &H3300  ' ignore certicicate error
    
    ' Send Request.
    MyRequest.send
    'And we get this response
    jsonstr = MyRequest.responseText
    If MyRequest.status <> 200 Then GoTo errLabel
    Set getFacets = JsonDecode(jsonstr)
    Exit Function
errLabel:
    MsgBox "ERROR connecting to server or decoding facet list"
        Set getFacets = JsonDecode("{}")
End Function



Function getEntityTypes() As Object
    Dim MyRequest As Object
    On Error GoTo errLabel
    
    Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    MyRequest.Open "GET", Cfg.getVar("cfgURL") & "/valid/purposes"
    MyRequest.setRequestHeader "X-Api-Key", "defaultapikey"
    MyRequest.Option(4) = &H3300  ' ignore certicicate error
    
    ' Send Request.
    MyRequest.send
    'And we get this response
    jsonstr = MyRequest.responseText
    If MyRequest.status <> 200 Then GoTo errLabel
    Set getEntityTypes = JsonDecode(jsonstr)
    Exit Function
errLabel:
    MsgBox "ERROR connecting to server or decoding type & purpose list"
    Set getEntityTypes = JsonDecode("{}")
End Function


Function getConfig(server) As Object
    Dim styles As Object
    Dim MyRequest As Object
    On Error GoTo errLabel
    
    url = server & "/word/config"
    Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    MyRequest.Open "GET", url
    MyRequest.setRequestHeader "X-Api-Key", "defaultapikey"
    MyRequest.Option(4) = &H3300  ' ignore certicicate error
    
    ' Send Request.
    MyRequest.send
    'And we get this response
    jsonstr = MyRequest.responseText
     
    If MyRequest.status <> 200 Then GoTo errLabel
    Set getConfig = JsonDecode(jsonstr)
    
    
    Exit Function
errLabel:
    'TODO this test should be removed once server updated
    MsgBox "GuideTools: ERROR getting configuration from server"
    Set getConfig = Nothing
End Function




Function getStyles() As Object
    Dim styles As Object
    Dim MyRequest As Object
    On Error GoTo errLabel
    
    Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    MyRequest.Open "GET", Cfg.getVar("cfgURL") & "/valid/styles"
    MyRequest.setRequestHeader "X-Api-Key", "defaultapikey"
    MyRequest.Option(4) = &H3300  ' ignore certicicate error
    
    ' Send Request.
    MyRequest.send
    'And we get this response
    jsonstr = MyRequest.responseText
    If MyRequest.status <> 200 Then GoTo errLabel
    Set getStyles = JsonDecode(jsonstr)
    Exit Function
errLabel:
    MsgBox "GuideTools: ERROR connecting to server or decoding style list"
    jsonstr = "{""styles"":    {""l_italic"":{""span"":""!l!"",""block"":""legal_italic"",""class"":""g_legal_italic"",""purpose"":""Italics required to match prescribed formatting in legislation, documents or forms""},""l_bold"":{""span"":""!b!"",""block"":""legal_bold"",""class"":""g_legal_bold"",""purpose"":""Bold required to match prescribed formatting in legislation, documents or forms""}}}"
    Set getStyles = JsonDecode(jsonstr)
End Function

Function getAssociates(aType As String) As Object
    Dim styles As Object
    Dim MyRequest As Object
    On Error GoTo errLabel
    
    Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    MyRequest.Open "GET", Cfg.getVar("cfgURL") & "/" & aType
    MyRequest.setRequestHeader "X-Api-Key", "defaultapikey"
    MyRequest.Option(4) = &H3300  ' ignore certicicate error
    
    ' Send Request.
    MyRequest.send
    'And we get this response
    jsonstr = MyRequest.responseText
    If MyRequest.status <> 200 Then GoTo errLabel
    Set getAssociates = JsonDecode(jsonstr)
    Exit Function
errLabel:
    MsgBox "GuideTools: ERROR connecting to server or decoding style list"
    jsonstr = "{""count"": ""1"",""articles"":[""id"":""classid"",""title"":""class tittle""]}"
    Set getClasses = JsonDecode(jsonstr)
End Function

Sub test_ss()
    syncStatus "wip", "content"
End Sub


Sub syncStatus(channel, Optional pub_type = "content")
   r = ShellExecute(0, "open", Cfg.getVar("guideMgr") & "/spy/publish/" & channel & "/" & pub_type & "/status", Chr(32), 0, 1)
End Sub

Sub pubStatus()
   r = ShellExecute(0, "open", Cfg.getVar("guideMgr") & "/spy/status", Chr(32), 0, 1)
End Sub




Function doPreview(kmj As String) As Object
    Dim MyRequest As Object
    Dim info As String
    On Error GoTo errLabel
 
    ' JohnP - VBA doesn't seem to be able to post data and trigger a browser window so need to store th post temporarily
    ' in the server and send a seperate follow link
    ' That is the following line will post the data but not display the result
    ' ActiveDocument.FollowHyperlink Address:=URL, NewWindow:=False, Method:=msoMethodPost, ExtraInfo:=info
    

    'TODO need to add identifier to prevent getting someone elses page back- RACF prob best to avoid having complex
    'server side garbage collection
    pv_id = LCase(Environ("UserName"))
    info = "kmj=" & URLEncode(kmj) & "&name=" & pv_id
    Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    MyRequest.Open "POST", Cfg.getVar("appURL") & "/preview"
    MyRequest.setRequestHeader "X-Api-Key", "defaultapikey"
    MyRequest.Option(4) = &H3300  ' ignore certicicate error
    
    MyRequest.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
 
    MyRequest.send info
    
    ' an alternative might be to store the responce in a file and just open that. However relative links wouldn't work
    'chr(32) as the 4th parameter stops chrome opening an extra tab TODO test this with other browsers
    r = ShellExecute(0, "open", Cfg.getVar("appURL") & "/preview" & "?name=" & pv_id, Chr(32), 0, 1)
    'follow hyperlink setnds the get result twivw
    'ActiveDocument.FollowHyperlink Address:=appURL & "/preview?name=johnp", NewWindow:=False

    Exit Function
errLabel:
    MsgBox Err.Description
    'jsonStr = "{""Service"":{""multiple"":""false"",""foci"":[""Registration"",""Information services""] }}"

End Function

Function doSharepoint(kmj As Object) As Object
    Dim MyRequest As Object
    Dim info As String
    Dim s As String
    On Error GoTo errLabel
 
    ' JohnP - VBA doesn't seem to be able to post data and trigger a browser window so need to store th post temporarily
    ' in the server and send a seperate follow link
    ' That is the following line will post the data but not display the result
    ' ActiveDocument.FollowHyperlink Address:=URL, NewWindow:=False, Method:=msoMethodPost, ExtraInfo:=info
    

    'TODO need to add identifier to prevent getting someone elses page back- RACF prob best to avoid having complex
    'server side garbage collection
    pv_id = LCase(Environ("UserName"))
    s = JsonEncode(kmj)
    info = "kmj=" & URLEncode(s) & "&name=" & pv_id
    Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    MyRequest.Open "POST", Cfg.getVar("appURL") & "/prestatic"
    MyRequest.setRequestHeader "X-Api-Key", "defaultapikey"
    MyRequest.Option(4) = &H3300  ' ignore certicicate error
    
    MyRequest.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
 
    MyRequest.send info
    
    If (MyRequest.status = 200) Then
        saveSharepoint kmj("id"), MyRequest.responseText
    End If

    Exit Function
errLabel:
    MsgBox Err.Description
    'jsonStr = "{""Service"":{""multiple"":""false"",""foci"":[""Registration"",""Information services""] }}"

End Function








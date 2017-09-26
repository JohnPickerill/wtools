Attribute VB_Name = "guideWord"
'TODO error handling
Function pushObject(branch As String, kmjObj As Object) As Object
    Dim MyRequest As Object
    Dim info As String
    Dim res  As Object
    Set res = CreateObject("Scripting.Dictionary")
    On Error GoTo errLabel
    url = Cfg.getVar("guideMgr") + "/branches/" + branch + "/kmj/" + kmjObj("id")
    Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    MyRequest.SetTimeouts 10000, 10000, 10000, 90000
    MyRequest.Open "POST", url, False
    MyRequest.setRequestHeader "Content-type", "application/json"
    MyRequest.setRequestHeader "X-Api-Key", "defaultapikey"
    MyRequest.Option(4) = &H3300  ' ignore certicicate error
    MyRequest.send JsonEncode(kmjObj)
        
    
    If (MyRequest.status >= 300) Then
        Set res("response") = CreateObject("Scripting.Dictionary")
        res("status") = MyRequest.status
        res("statusText") = MyRequest.StatusText
    Else
        Set res("response") = JsonDecode(MyRequest.responseText)
        res("status") = MyRequest.status
        res("statusText") = MyRequest.StatusText
    End If
       
    res("id") = kmjObj("id")
    Set pushObject = res
    Set MyRequest = Nothing
    Exit Function
errLabel:
    res("status") = 500
    res("statusText") = "Error VBA:" & Err.Description
    res("id") = kmjObj("id")
    Set res("response") = CreateObject("Scripting.Dictionary")
    Set pushObject = res
    Set MyRequest = Nothing
End Function


Function pushRecord(branch As String, rec As Object)
    Dim MyRequest As Object
    Dim info As String
    Dim res  As Object
    Set res = CreateObject("Scripting.Dictionary")
    On Error GoTo errLabel
    url = Cfg.getVar("guideMgr") + "/branches/" + branch + "/record"
    Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    MyRequest.SetTimeouts 10000, 10000, 10000, 90000
    MyRequest.Open "POST", url, False
    MyRequest.setRequestHeader "Content-type", "application/json"
    MyRequest.setRequestHeader "X-Api-Key", "defaultapikey"
    MyRequest.Option(4) = &H3300  ' ignore certicicate error
    content = JsonEncode(rec)
    MyRequest.send content
        
    Set pushRecord = MyRequest
    Exit Function
    
errLabel:
    res("status") = 500
    res("statusText") = "Error VBA:" & Err.Description
    Set pushRecord = res
    Set MyRequest = Nothing
End Function


Function pullRecord(branch As String) As Object
    Dim facets As Object
    Dim MyRequest As Object
    Dim url As String
    On Error GoTo errLabel
    
    Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    url = Cfg.getVar("guideMgr") + "/branches/" + branch + "/record"
    MyRequest.Open "GET", url
    MyRequest.setRequestHeader "X-Api-Key", "defaultapikey"
    MyRequest.Option(4) = &H3300  ' ignore certicicate error
    
    ' Send Request.
    MyRequest.send
    'And we get this response
    jsonstr = MyRequest.responseText
    If MyRequest.status <> 200 Then GoTo errLabel
    Set pullRecord = JsonDecode(jsonstr)
    Exit Function
errLabel:
    Set pullRecord = JsonDecode("{}")
End Function

 

Function syncES(channel)
    Dim MyRequest As Object
    Dim url As String
    On Error GoTo errLabel
    syncES = True
    resultsForm.setString "Requesting publication to channel " + channel + " ...."
    resultsForm.show
       
    If Not canDo("publisher") Then
        MsgBox ("You do not have sufficient priviledge for this action")
        resultsForm.hide
        Exit Function
    End If
    
    If vbNo = MsgBox("Are you sure to wish publish to channel " + channel, vbYesNo) Then
        resultsForm.hide
        Exit Function
    End If
    
     
    Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    url = Cfg.getVar("guideMgr") + "/publish/" + channel + "/_publish?user=" + LCase(Environ("UserName"))
    MyRequest.Open "PUT", url
    MyRequest.setRequestHeader "X-Api-Key", "defaultapikey"
    MyRequest.Option(4) = &H3300  ' ignore certicicate error

    ' Send Request.
    MyRequest.send
    'And we get this response
    resultsForm.hide
    If MyRequest.status = 200 Then
        MsgBox "Synchronisation of channel <" & channel & "> requested" & vbCrLf & MyRequest.responseText
        syncStatus channel
    Else
        MsgBox "Synchronisation of channel <" & channel & "> failed" & vbCrLf & MyRequest.StatusText
    End If
    Exit Function
errLabel:
    MsgBox "Error contacting server"
End Function

Function promote(source_channel, target_channel)
    punlish = True
    Dim MyRequest As Object
    Dim url As String
    On Error GoTo errLabel
    resultsForm.setString "Starting promotion of " + source_channel + " to " + target_channel + vbCrLf & "...."
    resultsForm.show
    
    If Not canDo("publisher") Then
        MsgBox ("You do not have sufficient priviledge for this action")
        Exit Function
    End If
    
    If vbNo = MsgBox("Are you sure you wish to publish to channel " + target_channel, vbYesNo) Then
        resultsForm.hide
        Exit Function
    End If
    
    
    Set request = CreateObject("Scripting.Dictionary")
    request("from") = source_channel
    request("to") = target_channel
    request("user") = LCase(Environ("UserName"))
    request("msg") = "from MS Word toolbar"
    
    Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    MyRequest.Option(4) = &H3300  ' ignore certicicate error
    url = Cfg.getVar("guideMgr") + "/promote?user=" + LCase(Environ("UserName"))
    MyRequest.SetTimeouts 10000, 10000, 10000, 90000
    MyRequest.Open "PUT", url
    MyRequest.setRequestHeader "Content-type", "application/json"
    MyRequest.setRequestHeader "X-Api-Key", "defaultapikey"
    MyRequest.send JsonEncode(request)

    resultsForm.hide
    If MyRequest.status = 200 Then
        MsgBox "Publication to channel <" & target_channel & "> requested" & vbLrcf & MyRequest.responseText
        syncStatus target_channel
    Else
        MsgBox "Publication to channel <" & target_channel & "> request failed"
    End If
    Exit Function
errLabel:
    MsgBox "Error contacting server"
End Function



Function pushES(channel)
    Dim MyRequest As Object
    Dim url As String
    On Error GoTo errLabel
    pushES = True
    resultsForm.setString "Requesting metadata push to channel " & channel & " ...."
    resultsForm.show
       
    If Not canDo("publisher") Then
        MsgBox ("You do not have sufficient priviledge for this action")
        resultsForm.hide
        Exit Function
    End If
    
    If vbNo = MsgBox("Are you sure you wish to do a metadata push to channel " & channel, vbYesNo) Then
        resultsForm.hide
        Exit Function
    End If
    
     
    Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    url = Cfg.getVar("guideMgr") + "/publish/" + channel + "/_publish_metadata?user=" + LCase(Environ("UserName"))
    MyRequest.Open "PUT", url
    MyRequest.setRequestHeader "X-Api-Key", "defaultapikey"
    MyRequest.Option(4) = &H3300  ' ignore certicicate error

    ' Send Request.
    MyRequest.send
    'And we get this response
    resultsForm.hide
    If MyRequest.status = 200 Then
        MsgBox "Metadata push to channel <" & channel & "> requested" & vbCrLf & MyRequest.responseText
        syncStatus channel, "metadata"
    Else
        MsgBox "Metadata push to channel <" & channel & "> failed" & vbCrLf & MyRequest.StatusText
    End If
    Exit Function
errLabel:
    MsgBox "Error contacting server"
End Function


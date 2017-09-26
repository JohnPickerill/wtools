VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CcText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ReClean As New RegExp
Private ReNoprint As New RegExp
Private ReWhitespace As New RegExp
Private ReTrimWhite As New RegExp
Private ReNotASCII As New RegExp
Private ReIFN As New RegExp
Private reUrl As New RegExp
Private reAmp As New RegExp
Private reCom As New RegExp
Private reWebdav As New RegExp

Private Sub Class_Initialize()
    
    'match whitespace
    ReClean.Global = True
    ReClean.MultiLine = True
    ReClean.pattern = "[ \t\r\n\v\f]"

    ' match strings only containing whitespace
    ReWhitespace.pattern = "^[ \t\r\n\v\f]+$"
    ReWhitespace.MultiLine = False ' TODO revisit the setting of this flag on all regex
    ReWhitespace.Global = True
    
    'trim whitespace
    ReTrimWhite.pattern = "^[ \t\r\n\v\f]*(.*?)[ \t\r\n\v\f]*$"
    ReTrimWhite.MultiLine = False
    

    'match characters outside normal ascii range
    With ReNotASCII
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "[^\x00-\x7F]"
    End With
    'Match Ascii non printable
    With ReNoprint
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "[^ -~]"
    End With
    
    'Invalid file characters (gitlab is most restrictive letters, numbers .-_ and no leading -)
    ' windows / unix [\^<>:""'/\\!?*\|]
    With ReIFN
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "[^\-\.\_0-9A-Za-z]"
    End With
    
    With reAmp
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "&|_&_"
    End With
    
'used @imme_emosol (54 chars) but this failed on english-heritage
'"https?:\/\/(-\.)?([^\s\/?\.#-]+\.?)+(\/[^\s]*)?$"
' switchto @stephenhay (38 chars) this passes more that should fail but this is better
'https?:\/\/[^\s\/$.?#].[^\s]*$
'https://mathiasbynens.be/demo/url-regex
'must trim before
    
    
    With reUrl
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = "https?:\/\/[^\s\/$.?#].[^\s]*$"
    End With
    
     
    With reCom
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = "^com[0-9].*"
    End With
    
    
    With reWebdav
        .Global = False
        .MultiLine = False
        .IgnoreCase = True
        .pattern = "^https?:(\/\/[^/]*)(.*)"
    End With
    
End Sub

Public Function webdav(url) As String
    webdav = "Could not create webdav address for sharepoint"
    Dim mc As MatchCollection
    Set mc = reWebdav.Execute(url)
    If mc.count > 0 Then
        webdav = mc(0).SubMatches(0) + "@SSL\DavWWWRoot" + mc(0).SubMatches(1)
        webdav = replace(webdav, "/", "\")
        If Right(webdav, 1) = "\" Then
            webdav = Left(webdav, Len(webdav) - 1)
        End If
    End If
End Function


Public Function testURL(url As String) As Boolean
    testURL = reUrl.test(url)
End Function

Public Function TrimWhite(str As String) As String
    Dim m As MatchCollection
    Set m = ReTrimWhite.Execute(str)
    If m.count = 1 Then
        TrimWhite = m.item(0).SubMatches(0)
    Else
        TrimWhite = str
    End If
     
End Function




Public Function cleanUID(str) As String
'makes it a valid file body and removes spaces and sets maximum length
'github only accepts letters numbers - and _ and .
    'cleanUID = str
    'cleanUID = ReClean.replace(str, "")
    'cleanUID = ReNoprint.replace(cleanUID, "")
    'cleanUID = Left(ReIFN.replace(cleanUID, ""), 20)
    
    
    cleanUID = Left(reAmp.replace(str, "_"), 20)
    If reCom.test(cleanUID) Then
        cleanUID = "_" + cleanUID
    End If
    cleanUID = Left(ReIFN.replace(cleanUID, ""), 20)
    cleanUID = LCase(cleanUID)
End Function

Public Function cleanText(str, Optional sLen As Integer = 0) As String
'normal ascii characters and normal controls
    cleanText = ReNotASCII.replace(str, "")
    If sLen > 0 Then cleanText = Left(cleanText, sLen)
End Function


Public Function isWhitespace(str As String) As Boolean
    isWhitespace = (Len(str) = 0) Or ReWhitespace.test(str)
End Function









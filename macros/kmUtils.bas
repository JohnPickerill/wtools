Attribute VB_Name = "kmUtils"
Option Explicit
 
 






Function cleanStr(ByRef txt As String) As String
        Do While ((Right(txt, 1) = vbCr) Or (Right(txt, 1) = vbLf))
            txt = Left(txt, (Len(txt) - 1))
        Loop
        txt = RTrim(txt)
        cleanStr = "ok"
End Function

Public Function getFileName(filnam As String) As String
    Dim reExp As New RegExp
    Dim m As MatchCollection
    'getfilename from path
    
    With reExp
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = "^((.+)(?=([\/\\]))[\/\\](..*)$)"
    End With
    Set m = reExp.Execute(filnam)

    If m.count > 0 Then
        getFileName = m(0).SubMatches(3)
    Else
        getFileName = "nofile"
        ' TODO get this warning lots of times so must be calling this routine redndantly
        MsgBox "Warnimg : Word document must be saved at least once to give it a name to attach articles and images to"
    End If
End Function



Public Function cleanFilename() As String
' TODO check that this isn't called in a loop anywhere . If it is we need to factor this into a class
    Dim reExp As New RegExp
    Dim m As MatchCollection
    
    'remove file extension and substitute for silly characters
    With reExp
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .pattern = "^(.+)(?=(\..*)$)"
    End With
    Set m = reExp.Execute(ActiveDocument.name)

    If m.count > 0 Then
        Dim reCFN As New RegExp
        With reCFN
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .pattern = "[\s()\[\]{}]"
        End With
        cleanFilename = reCFN.replace(m(0).SubMatches(0), "_")
    Else
        cleanFilename = "nofile"
        ' TODO get this warning lots of times so must be calling this routine redndantly
        MsgBox "Warnimg : Word document must be saved at least once to give it a name to attach articles and images to"
    End If
End Function


Function JsonDecode(jsonString) As Object 'This works, uses vba-json library
    Dim lib As New JSONLib 'Instantiate JSON class object
    Dim jsonParsedObj As Object
    Set jsonParsedObj = lib.parse(CStr(jsonString))
    Set JsonDecode = jsonParsedObj
    Set jsonParsedObj = Nothing
    Set lib = Nothing
End Function

Function JsonEncode(jsonObj) As String   'This works, uses vba-json library
    Dim lib As New JSONLib
    JsonEncode = lib.toString(jsonObj)
    Set lib = Nothing
End Function

Public Function URLEncode( _
   StringVal As String, _
   Optional SpaceAsPlus As Boolean = False _
) As String

  Dim StringLen As Long: StringLen = Len(StringVal)

  If StringLen > 0 Then
    ReDim Result(StringLen) As String
    Dim i As Long, CharCode As Integer
    Dim Char As String, Space As String

    If SpaceAsPlus Then Space = "+" Else Space = "%20"

    For i = 1 To StringLen
      Char = Mid$(StringVal, i, 1)
      CharCode = asc(Char)
      Select Case CharCode
        Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
          Result(i) = Char
        Case 32
          Result(i) = Space
        Case 0 To 15
          Result(i) = "%0" & Hex(CharCode)
        Case Else
          Result(i) = "%" & Hex(CharCode)
      End Select
    Next i
    URLEncode = Join(Result, "")
  End If
End Function



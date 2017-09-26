VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private fso As FileSystemObject
Private ts As TextStream
Private fn As String

 

Public Sub Append(txt As String)
    ts.WriteLine (txt)
End Sub
    
Public Sub init(filename As String)
    fn = filename
    Set ts = fso.CreateTextFile(fn, True)
End Sub

Public Sub Flush()
    If Not ts Is Nothing Then
        ts.Close
    End If
     
    Set ts = fso.OpenTextFile(fn, ForAppending)
End Sub


Private Sub Class_Initialize()
    Set fso = CreateObject("Scripting.FileSystemObject")
End Sub

Private Sub Class_Terminate()
    If Not ts Is Nothing Then
        ts.Close
    End If
End Sub

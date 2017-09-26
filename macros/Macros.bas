Attribute VB_Name = "Macros"
Option Explicit
Sub ShowEditor()
ShowVisualBasicEditor = True
End Sub
Sub Macro1()
Dim r As Range
Set r = ActiveDocument.Range(start:=Selection.Range.start - 1, End:=Selection.Range.start)
If r.text = vbCr Then
        r.MoveEnd wdCharacter, 1
        r.Select
        r.start = r.End - 1
        r.InsertBefore " "
End If
End Sub

Sub tmp()
    muEdit.padPara Selection.Range
End Sub


Sub Macro2()
MsgBox "Article 2"
End Sub
Sub Macro3()
MsgBox "Article 3"
End Sub

Sub insertLink()
MsgBox "link not implemented yet"
End Sub


 

Sub testlink()
    Dim L As Object
    
    Dim extLinks As New Collection
    Set L = CreateObject("Scripting.Dictionary")
    L("url") = "url"
    L("title") = "title"
    L("scope") = "scope this could be quite a lot of text descriping what is in the linked page it could go on and on and cover several lines so I don't know it we need to do something special"
    extLinks.Add L
    Set L = CreateObject("Scripting.Dictionary")
    L("url") = "1url"
    L("title") = "1title"
    L("scope") = "1scope this could be quite a lot of text descriping what is in the linked page it could go on and on and cover several lines so I don't know it we need to do something special"
    extLinks.Add L
    linkForm.load extLinks
    linkForm.show
End Sub

 



Sub expandtest()
    Dim s As Range
    Set s = muEdit.targetMarkup(Selection.Range)
    If Not s Is Nothing Then s.Select
End Sub



Sub markTest()
 
 
 
      muEdit.markMarkup Selection.Range.Paragraphs(1), True
  
    
     
     'content controls not available in word 97
     'Selection.Range.ContentControls.add (wdContentControlRichText)
     'Selection.ParentContentControl.Title = "end:" & kmj("id")
     'Selection.ParentContentControl.Tag = "metaTag"
     'Selection.ParentContentControl.LockContentControl = True
     'Selection.ParentContentControl.LockContents = True
     
End Sub

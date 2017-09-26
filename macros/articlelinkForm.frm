VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} articlelinkForm 
   Caption         =   "Insert Link to article"
   ClientHeight    =   3396
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   6696
   OleObjectBlob   =   "articlelinkForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "articlelinkForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rng As Range
Private bExists As Boolean
Private Enum linkEnum
    art = 1
    pic = 2
    rel = 3
    url = 4
End Enum
Dim linkType As linkEnum

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


Private Sub cbArticle_Click()
    linkType = art
    Label3.Caption = "Anchor"
    ctTitle.Width = 100
    'ctLink = cctxt.cleanUID(ctLink.text)
    ctLink_Change
End Sub

Private Sub cbImage_Click()
    ctTitle.Width = 200
    Label3.Caption = "Popup Description"

    linkType = pic
    ctLink_Change
End Sub

Private Sub cbRelative_Click()
    ctTitle.Width = 200
    Label3.Caption = "Popup Description"

    linkType = rel
    ctLink_Change
End Sub

Private Sub cbURL_Click()
    ctTitle.Width = 200
    Label3.Caption = "Popup Description"
    linkType = url
    ctLink_Change
End Sub


Private Sub cbCancel_Click()
    articlelinkForm.hide
End Sub

Private Function buildLink(link As String, text As String, title As String, Optional img As Boolean = False)
    Dim str As New StringBuilder
    If img Then str.Append ("!")
    str.Append ("[")
    str.Append (text)
    str.Append "]("
    str.Append link
    If (Len(link) > 0) Then
        str.Append " """
        str.Append title
        str.Append """)"
    Else
        str.Append ")"
    End If
    buildLink = str.text
End Function




 
Private Sub cbOk_Click()
    Dim str As String
    Dim lnk As String
    
    
    If (ctLink.text = "") Then
        MsgBox ("Invalid link")
        Exit Sub
    End If
    If ((linkType = url) And Not (cctxt.testURL(ctLink.text))) Then
        MsgBox ("Invalid link")
        Exit Sub
    End If
    ctText.text = Trim(ctText.text)
    If (ctText.text = "") Then
        MsgBox ("Invalid display label")
        Exit Sub
    End If
    
    ctTitle.text = Trim(ctTitle.text)
    If (linkType = art) Then
        If (ctTitle.text <> "") Then
            lnk = ctLink.text + "#" + ctTitle.text
        Else
            lnk = ctLink.text
        End If
        ctTitle.text = "link to guidance article"
    Else
        lnk = ctLink.text
    End If
    
    
    
    If (((linkType = rel) Or (linkType = pic))) Then
        If (ctLink.text = "/") Then
            MsgBox ("Invalid link")
            Exit Sub
        End If
        If (Left(ctLink.text, 1) <> "/") Then
            MsgBox ("Invalid relative link syntax - must begin with /")
            Exit Sub
        End If
    End If
        
    

  
    str = buildLink(lnk, ctText.text, ctTitle.text, linkType = pic)
    rng.text = str
    muEdit.wrapExtra rng

endLabel:
    articlelinkForm.hide
End Sub

Private Sub srch(ifn, indx)
    Set dlgOpen = Application.FileDialog(FileDialogType:=msoFileDialogFilePicker)
    dlgOpen.Filters.clear
    dlgOpen.Filters.Add "Images", "*.png;*.jpg;*.jpeg;*.gif", 1
    dlgOpen.Filters.Add "pdf", "*.pdf", 2
    dlgOpen.FilterIndex = indx
    dlgOpen.AllowMultiSelect = False
    dlgOpen.InitialFileName = ifn
    dlgOpen.ButtonName = "Select"
    If dlgOpen.show = -1 Then
        Dim str As String
        str = dlgOpen.SelectedItems(1)
        L = InStr(str, "assets") + Len("assets")
        ctLink.text = Right(str, Len(str) - L)
    End If
End Sub


Private Sub cbSearch_Click()
'https://technet.microsoft.com/en-us/library/ee692906.aspx
    Dim dlgOpen As FileDialog
    Dim staticDir As String
    
    staticDir = Cfg.getVar("images")
    
    Dim ifn As String
    If (linkType = pic) Then
        srch staticDir & "_static/images/*.*", 1
    ElseIf (linkType = rel) Then
        srch staticDir & "_static/docs/*.*", 2
    ElseIf (linkType = art) Then
        MsgBox ("Use search in metadata editor and then copy and paste article UID")
        s = Cfg.getVar("cfgURL") & "/search"
        r = ShellExecute(0, "open", Cfg.getVar("cfgURL") & "/search", Chr(32), 0, 1)
    Else
        MsgBox ("not implemented for this link type")
        Exit Sub
    End If

    Set dlgOpen = Nothing
End Sub



Private Sub ctLink_Change()
    Select Case linkType
        Case art
            ctLink.text = cctxt.cleanUID(ctLink.text)
        Case url
            ctLink.text = replace(cctxt.cleanText(ctLink.text), " ", "%20")
        Case rel
            If Left(ctLink.text, 1) <> "/" Then
                ctLink.text = "/" + ctLink.text
            End If
            ctLink.text = replace(cctxt.cleanText(ctLink.text), " ", "%20")
        Case pic
            If Left(ctLink.text, 1) <> "/" Then
                ctLink.text = "/" + ctLink.text
            End If
            ctLink.text = replace(cctxt.cleanText(ctLink.text), " ", "%20")
    End Select
End Sub

Private Sub ctText_Change()
    ctText.text = cctxt.cleanText(ctText.text)
End Sub

Private Function checkLink(str As String) As Object
    Dim mc As MatchCollection
    Dim wa() As String
    
    Set L = CreateObject("Scripting.Dictionary")
    
    
    ' check for markdown link
    Dim reMdLink As New RegExp
    'TODO should remove smartquotes on markdown conversion not here ?
    qt = "[" + Chr(147) + Chr(148) + Chr(34) + "]" 'quotes and smartquote
    With reMdLink
        .MultiLine = True
        .Global = True
        .pattern = "(!?)\[(.*)]\((\S+)\s*" + qt + "(.*)" + qt + "\)"
    End With
    Set mc = reMdLink.Execute(str)
    If mc.count > 0 Then
        L("ok") = True
        L("link") = mc(0).SubMatches(2)
        L("text") = mc(0).SubMatches(1)
        L("title") = mc(0).SubMatches(3)
        
        If mc(0).SubMatches(0) = "!" Then
            L("typ") = pic
        ElseIf (Left(mc(0).SubMatches(2), 1) = "/") Then
            L("typ") = rel
        ElseIf (cctxt.testURL(mc(0).SubMatches(2))) Then
            L("typ") = url
        Else
            L("typ") = art
              
            wa = Split(L("link"), "#")
            L("link") = wa(0)
            If UBound(wa) > 0 Then
                L("title") = wa(1)
            Else
                L("title") = ""
            End If
        End If
             
        Set checkLink = L
        Exit Function
    End If
    
    ' check for wikitext link
    Dim ReLink As New RegExp
    ReLink.MultiLine = True
    ReLink.pattern = "\[\[\s*([\S]+?)(?:\s*(?:\|| )([\s\S]+?))?\]\]"
    ReLink.Global = True
    Set mc = ReLink.Execute(rng.text)
    ' should at most be one
    If mc.count > 0 Then
        L("ok") = True
        L("link") = mc(0).SubMatches(0)
        L("text") = mc(0).SubMatches(1)
        L("title") = ""
        
        If (cctxt.testURL(mc(0).SubMatches(0))) Then
            L("typ") = url
        Else
            L("typ") = art
        End If
        
        Set checkLink = L
        Exit Function
    End If
    
    L("ok") = False
    Set checkLink = L
    
End Function


Private Sub UserForm_Activate()
    'take para mark out of selection

    Set L = CreateObject("Scripting.Dictionary")
    Set rng = muEdit.targetMarkup(Selection.Range)
    If rng Is Nothing Then
        Set rng = Selection.Range
        L("ok") = True
        L("link") = ""
        L("text") = ""
        L("title") = ""
        L("typ") = art
    Else
        Set L = checkLink(rng.text)
        
        If (Not L("ok")) Then
            MsgBox "Selected text already contains markup"
            articlelinkForm.hide
            Exit Sub
        End If
    End If
        
 
    If rng.End = rng.Paragraphs(1).Range.End Then
                    ActiveDocument.Range(start:=rng.start, End:=rng.End - 1).Select
    End If
    

    
    cbArticle.value = (L("typ") = art)
    cbURL.value = (L("typ") = url)
    cbRelative.value = (L("typ") = rel)
    cbImage.value = (L("typ") = pic)
    
    ctLink.text = L("link")
    ctText.text = L("text")
    ctTitle.text = L("title")
    
End Sub

 
Private Sub UserForm_Initialize()
    bArticle = True
    cbArticle = bArticle
End Sub


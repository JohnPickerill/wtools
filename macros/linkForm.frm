VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} linkForm 
   Caption         =   "Related External Link"
   ClientHeight    =   5652
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   6264
   OleObjectBlob   =   "linkForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "linkForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xlinks As Collection
'Dim reUrl As New RegExp
Private modInput As Boolean
 
Public Sub load(kmLinks)
    Set xlinks = kmLinks
    dispLinks
End Sub

Private Sub updateLinks()
    Do Until xlinks.count = 0
        xlinks.Remove 1
    Loop
    For i = 0 To ctLinks.ListCount - 1
        Set L = CreateObject("Scripting.Dictionary")
        L("name") = ctLinks.List(i, 0)
        L("url") = ctLinks.List(i, 1)
        L("scope") = ctLinks.List(i, 2)
        xlinks.Add L
    Next i
End Sub

Public Sub dispLinks()
    ctLinks.clear
    ctLinks.ColumnCount = 3
    For Each L In xlinks
       modLink L("name"), L("url"), L("scope")
    Next L
End Sub

Public Function modLink(ByVal title As String, ByVal url As String, ByVal scope As String, Optional Index As Integer = -1) As Boolean
        'this function is used to add or edit

        
        url = Trim(url)
        'If Not reUrl.test(url) Then
        If Not cctxt.testURL(url) Then
            MsgBox "URL is invalid"
            modLink = False
            Exit Function
        End If
        'this is to prevent duplicates
        For i = 0 To ctLinks.ListCount - 1

            If LCase(url) = LCase(ctLinks.Column(1, i)) Then
                If (Index = i) Then
                    'so this is a valid modify
                    Exit For
                Else
                    MsgBox "URL is already in list"
                    modLink = False
                    Exit Function
                End If
            End If
        Next i
        
        'prevent input fields being modified by event handler during this function
        modInput = False
        On Error GoTo errlab
        If Index = -1 Then
            'add
            ctLinks.AddItem title
            listInd = ctLinks.ListCount - 1
        Else
            'modify
            If (Index >= 0) And (Index < ctLinks.ListCount) Then
                listInd = Index
                ctLinks.List(listInd, 0) = title
            Else
                modLink = False
                Exit Function
            End If
        End If
        'TODO need to check that modification doesn't produce duplicate link or else do de-dup
        ctLinks.List(listInd, 1) = url
        ctLinks.List(listInd, 2) = scope
        modLink = True
errlab:
    'allow input fields to be changed by clicking list
    modInput = True
    
End Function






 

Private Sub cbAdd_Click()
    url = ctUrl

    If Not modLink(ctTitle.text, ctUrl.text, ctScope.text) Then
        
        Exit Sub
    End If
End Sub



Private Sub cbCancel_Click()
    linkForm.hide
End Sub




Private Sub cbOk_Click()
    updateLinks
    linkForm.hide
End Sub

Private Sub cbRemove_Click()
    ctLinks.RemoveItem ctLinks.ListIndex
End Sub

Private Sub ctLinks_Click()
    'this function is triggered by changing ctLinks.List(i,0) so we need to be careful
    If modInput Then
        ctTitle.text = ctLinks.List(ctLinks.ListIndex, 0)
        ctUrl.text = ctLinks.List(ctLinks.ListIndex, 1)
        ctScope.text = ctLinks.List(ctLinks.ListIndex, 2)
    End If
End Sub

Private Sub cbEdit_Click()

    If Not modLink(ctTitle.text, ctUrl.text, ctScope.text, ctLinks.ListIndex) Then
        MsgBox "Modification failed"
        Exit Sub
    End If
End Sub

 

Private Sub ctScope_Change()
ctScope.text = cctxt.cleanText(ctScope.text)
End Sub

Private Sub ctTitle_Change()
    ctTitle.text = cctxt.cleanText(ctTitle.text)
End Sub

Private Sub ctUrl_Change()
    ctUrl.text = replace(cctxt.cleanText(ctUrl.text), " ", "%20")
End Sub

Private Sub UserForm_Initialize()
'used @imme_emosol (54 chars) but this failed on english heritage
'"https?:\/\/(-\.)?([^\s\/?\.#-]+\.?)+(\/[^\s]*)?$"
'switched to @stephenhay (38 chars) this passes more that should fail but this is better
'https?:\/\/[^\s\/$.?#].[^\s]*$
'https://mathiasbynens.be/demo/url-regex
'must trim before


'         With reUrl
'        .Global = True
'        .MultiLine = True
'        .IgnoreCase = False
'        .Pattern = "https?:\/\/[^\s\/$.?#].[^\s]*$"
'    End With
    modInput = True
End Sub

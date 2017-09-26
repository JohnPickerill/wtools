VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} snippetForm 
   Caption         =   "Snippet"
   ClientHeight    =   3492
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   7044
   OleObjectBlob   =   "snippetForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "snippetForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 

Private Sub bCancel_Click()
    snippetForm.hide
End Sub

Private Sub bOK_Click()
    If ctId.text = "Enter Unique ID" Then
        MsgBox "id is invalid"
        Exit Sub
    End If
    
    Dim kmj As Object
    Set kmj = muEdit.createSnippetObj
    kmj("id") = ctId.text
    kmj("purpose") = cPurpose.text
    kmj("master")("where") = "Word"
    kmj("master")("filename") = ActiveDocument.name
    On Error Resume Next ' in case variable doesn't exist
    kmj("master")("version") = ActiveDocument.Variables("VersionId")
    On Error GoTo 0
    
    'TODO think about whether to allow multiple clusters
    Do While kmj("clusters").count > 0
         kmj("clusters").Remove 1
    Loop
    Set c = CreateObject("Scripting.Dictionary")
    c("cluster") = cClusters.text
    c("priority") = 0
    kmj("clusters").Add c
  
    muEdit.addKmjMeta kmj:=kmj, article:=Selection.Range
    snippetForm.hide
    
End Sub
 

Private Sub ctId_Change()
    ctId.text = cctxt.cleanUID(ctId.text)
End Sub

Private Sub UserForm_Activate()
     Dim kmj As Object
     
     Set kmj = Nothing
        
     If Not checkSelection(kmj) Then
        snippetForm.hide
        Exit Sub
     End If
               
     'initialise form
      If kmj Is Nothing Then
        Set kmj = muEdit.createSnippetObj
      Else
        If kmj("type") <> setupForm.getType Then
            MsgBox ("meta data type is not the same as document type")
            snippetForm.hide
        End If
      End If
      
      
      cPurpose.text = kmj("purpose")
      ctId.text = kmj("id")
      
      
  
      ' clusters
      tClusters = ""
      del = ""
      For Each prent In kmj("clusters")
           tClusters = tClusters & del & prent("cluster")
           del = vbCrLf
      Next prent
      cClusters.text = tClusters
      
      
End Sub

Private Sub UserForm_Initialize()
    cPurpose.AddItem "url"
    cPurpose.AddItem "cre"
    cPurpose.AddItem "snp"
End Sub

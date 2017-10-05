VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} entityForm 
   Caption         =   "UserForm1"
   ClientHeight    =   6816
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   7740
   OleObjectBlob   =   "entityForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "entityForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim sLen As Integer

Private Sub bCancel_Click()
        entityForm.hide
End Sub



Private Sub bOK_Click()
        If tID.text = "Enter Unique ID" Then
            MsgBox "id is invalid"
            Exit Sub
        End If
             
        Dim kmj As Object
        Set kmj = muEdit.createKmjObj
        kmj("id") = tID
        kmj("title") = tTitle
        'kmj("scope") = tScope
        kmj("author") = tAuthor.text
        kmj("expert") = tExpert.text
        kmj("owner") = tOwner.text
        kmj("master")("where") = "Word"
        kmj("master")("filename") = ActiveDocument.name
        On Error Resume Next ' in case variable doesn't exist
        kmj("master")("version") = ActiveDocument.Variables("VersionId")
        On Error GoTo 0
        
        'kmj("lastupdate") = Format(Now, "yyyy-mm-ddThh:mm:ss.000Z")
 
        kmj("sensitivity") = ccbSensitivity.text
        thisCluster = tParent
           
        ' Clusters
        'TODO clusters (a simple list of names) has been replaced by clusters a structure with sequence. need to remove the old stuff at some point
        Do While kmj("cluster").count > 0
            kmj("cluster").Remove 1
        Loop
        Do While kmj("clusters").count > 0
            kmj("clusters").Remove 1
        Loop

        kmj("cluster").Add thisCluster
        Set c = CreateObject("Scripting.Dictionary")
        c("cluster") = thisCluster
        c("priority") = 9999
        kmj("clusters").Add c
        clusters = Split(tClusters, vbCrLf)
        For Each cl In clusters
            cl = Trim(cl)
            If Len(cl) > 0 Then
                If cl <> thisCluster Then
                    kmj("cluster").Add cl ' old schema
                    Set c = CreateObject("Scripting.Dictionary")
                    c("cluster") = cl
                    c("priority") = 9999
                    kmj("clusters").Add c 'new schema
                End If
            End If
        Next cl
        
 
        muEdit.addKmjMeta kmj:=kmj, article:=Selection.Range
        entityForm.hide
End Sub

Private Sub clearPublic()
    Set facets = Nothing
    Set cfacet = Nothing
    switchFacet = False
End Sub
 
 











 
Private Sub tClusters_Change()
    tClusters.text = LCase(cctxt.cleanText(tClusters.text))
End Sub

 

Private Sub tID_Change()
    tID.text = cctxt.cleanUID(tID.text)
    tItem = tID.text
End Sub





Private Sub tTitle_Change()
    tTitle.text = cctxt.cleanText(tTitle, sLen)
End Sub


Private Sub UserForm_Activate()

     
     Dim st As Paragraph
     Dim fin As Paragraph
     
     Dim para As Paragraph
     Dim kmj As Object
     
     Set kmj = Nothing
        
     If Not checkSelection(kmj) Then
        entityForm.hide
        Exit Sub
     End If
     
          
     'initialise form
     entityForm.Caption = "Create/Update " & setupForm.getType & ":" & setupForm.getPurpose
      If kmj Is Nothing Then
        Set kmj = muEdit.createKmjObj
      Else
        If kmj("type") <> setupForm.getType Then
            'MsgBox ("meta data type is not the same as document type")
            entityForm.hide
            Exit Sub
        End If
        If kmj("purpose") <> setupForm.getPurpose Then
            'MsgBox ("meta data type is not the same as document type")
            entityForm.hide
            Exit Sub
        End If
      End If
     
     
    If kmj("type") = "snippet" Then
        sLen = 200
    Else
        sLen = 70
    End If
 

      tID.text = kmj("id")
'      tScope.text = kmj("scope")
      tTitle.text = kmj("title")


 
      
      'update if not already set
      If kmj("owner") = "" Then
            muEdit.setPeople kmj
      End If
      
      tAuthor.text = kmj("author")
      tExpert.text = kmj("expert")
      tOwner.text = kmj("owner")
 
        
      cVersion.text = "0"
      On Error Resume Next
      cVersion.text = ActiveDocument.Variables("VersionId")
      On Error GoTo 0
      
       
      Dim thisCluster As String
      
      thisCluster = setupForm.getCluster()
      
      ccbSensitivity.text = kmj("sensitivity")
      If ccbSensitivity.text = "" Then
        ccbSensitivity.text = ccbSensitivity.List(0)
      End If
      
      
      tParent = thisCluster
      ' clusters
      tClusters = ""
      del = ""
      For Each prent In kmj("cluster")
           If prent <> thisCluster Then
                tClusters = tClusters & del & prent
                del = vbCrLf
           End If
      Next prent
     
  
     
    Exit Sub
endLabel:
    entityForm.hide
End Sub

Private Sub UserForm_Initialize()
     load entityForm
     ccbSensitivity.AddItem "normal"
     ccbSensitivity.AddItem "sensitive"
End Sub


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} metaForm 
   Caption         =   "Guidance MetaData"
   ClientHeight    =   8832
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   7884
   OleObjectBlob   =   "metaForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "metaForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Public place As Range
Public allFacets As Object
Dim extLinks As Collection
Dim items As Collection
Private saveId As String

Private Sub bCancel_Click()
        metaForm.hide
End Sub


Private Sub bOK_Click()
        If tID.text = "enteruniqueid" Then
            MsgBox "Please enter a valid article ID"
            Exit Sub
        End If
        
        
        If tID <> saveId Then
            If gaExists(tID) Then
                 If vbNo = MsgBox("You have changed the article id to one that has already been published. " & vbCrLf & " Are you sure you wish to do this  ?", vbYesNo) Then
                    Exit Sub
                 End If
            Else
                If vbNo = MsgBox("Proceeding will result in a new article being created at the next publication." & vbCrLf & "Are you sure you wish to do this ?", vbYesNo) Then
                    Exit Sub
                End If
            End If
        End If
        
        
        Dim kmj As Object
        Set kmj = muEdit.createKmjObj
        kmj("id") = tID
        kmj("title") = tTitle
        kmj("scope") = tScope
        
        kmj("author") = tAuthor.text
        kmj("expert") = tExpert.text
        kmj("owner") = tOwner.text
        kmj("master")("where") = "Word"
        kmj("master")("filename") = ActiveDocument.name
        kmj("master")("change") = tChange.text
        If Not bSignificance Then
            kmj("master")("significance") = "low"
        Else
            kmj("master")("significance") = "high"
        End If
        
        If Not bArchive Then
            kmj("archive") = "false"
        Else
            kmj("archive") = "true"
        End If
        
        
        On Error Resume Next ' in case variable doesn't exist
        kmj("master")("version") = ActiveDocument.Variables("VersionId")
        On Error GoTo 0
        
        kmj("lastupdate") = Format(Now, "yyyy-mm-ddThh:mm:ss.000Z")


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
 
        metaForm.hide
End Sub


Private Sub btnClear_Click()
    tChange.text = ""
End Sub




Private Sub tbType_Change()

End Sub

 

Private Sub tClusters_Change()
    tClusters.text = LCase(cctxt.cleanText(tClusters.text))
End Sub


Private Sub tID_Change()
    tID.text = cctxt.cleanUID(tID.text)
    tItem = tID.text
End Sub




Private Sub tScope_Change()
    tScope.text = cctxt.cleanText(tScope, 200)
End Sub


Private Sub tTitle_Change()
    tTitle.text = cctxt.cleanText(tTitle, 75)
End Sub


Private Sub UserForm_Activate()

     
     Dim st As Paragraph
     Dim fin As Paragraph
     
     Dim para As Paragraph
     Dim kmj As Object
     
     Set kmj = Nothing
        
     If Not checkSelection(kmj) Then
        metaForm.hide
        Exit Sub
     End If
     
          
     'initialise form
      If kmj Is Nothing Then
        If vbNo = MsgBox("The cursor does not indicate an article;  are you sure you wish to a new create one", vbYesNo) Then
            metaForm.hide
            Exit Sub
        End If
        Set kmj = muEdit.createKmjObj
      Else
        If kmj("type") <> setupForm.getType Then
            MsgBox ("meta data type is not the same as document type")
        End If
      End If
     
      If (kmj("purpose") = "legislation") Then
            tScope.Enabled = False
      Else
            tScope.Enabled = True
      End If
     
      tID.text = kmj("id")
      saveId = kmj("id")

      tScope.text = kmj("scope")
      
      tTitle.text = kmj("title")


      'update if not already set
      If kmj("owner") = "" Then
            muEdit.setPeople kmj
      End If
      
      tAuthor.text = kmj("author")
      tExpert.text = kmj("expert")
      tOwner.text = kmj("owner")
      tType.text = kmj("type")
      tPurpose.text = kmj("purpose")
 
 
      tChange.text = kmj("master")("change")
      If kmj("master")("significance") = "low" Then
        bSignificance = False
      Else
        bSignificance = True
      End If
        
      If kmj("archive") = "false" Then
        bArchive = False
      Else
        bArchive = True
      End If
        
        
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
    metaForm.hide
End Sub

Private Sub UserForm_Initialize()
     load metaForm
     ccbSensitivity.AddItem "normal"
     ccbSensitivity.AddItem "sensitive"
End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} metaForm 
   Caption         =   "Guidance MetaData"
   ClientHeight    =   10248
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   7080
   OleObjectBlob   =   "metaForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "metaForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public facets As Object
Public cfacet As Object
Public switchFacet As Boolean
Public place As Range
Public allFacets As Object
Dim extLinks As Collection
Dim items As Collection

Private Sub bCancel_Click()
        metaForm.Hide
End Sub

Private Sub bExtLink_Click()
    linkForm.load extLinks
    linkForm.show
    disExtLinks
End Sub

Private Sub bItems_Click()
    itemForm.load items
    itemForm.show
    disItems
End Sub


Private Sub bFacet_Change()
    switchFacet = True
    bFoci.clear
    facetName = bFacet.text
   


   
    For Each foci In allFacets(facetName)("foci")
            bFoci.AddItem foci
    Next foci
     
    
    Set cfacet = Nothing
    For Each facet In facets
        If facet("name") = facetName Then
            Set cfacet = facet
            For Each foci In facet("foci")
                For i = 0 To bFoci.ListCount - 1
                    If bFoci.Column(0, i) = foci Then
                         bFoci.Selected(i) = True
                    End If
                Next i
            Next foci
        End If
    Next facet
    
    If cfacet Is Nothing Then
           Set cfacet = CreateObject("Scripting.Dictionary")
           cfacet.Add "name", facetName
           cfacet.Add "foci", New Collection
           facets.Add cfacet
    End If
    
        
    switchFacet = False
End Sub


Private Sub bFoci_Change()
       Dim str As String
       
       If switchFacet Then GoTo endLabel
           
       Do Until cfacet("foci").count = 0
            cfacet("foci").Remove 1
       Loop
       
       For i = 0 To bFoci.ListCount - 1
       If bFoci.Selected(i) Then
            cfacet("foci").Add bFoci.Column(0, i)
       End If
       Next i
endLabel:
tFacets.text = listFacets()
End Sub




Private Function listFacets() As String
    Dim str As String
    str = ""
    For Each facet In facets
        str = str + facet("name") + " : [ "
        delm = " "
        For Each foci In facet("foci")
            str = str + delm + foci
            delm = ","
        Next foci
        str = str + "]" + vbCrLf
    Next facet
     listFacets = str
End Function



 
 


Private Sub bOK_Click()
        If tID.text = "Enter Unique ID" Then
            MsgBox "id is invalid"
            Exit Sub
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
        On Error Resume Next ' in case variable doesn't exist
        kmj("master")("version") = ActiveDocument.Variables("VersionId")
        On Error GoTo 0
        
        kmj("lastupdate") = Format(Now, "yyyy-mm-ddThh:mm:ss.000Z")
        kmj("fees") = ctFees
        kmj("class") = ctClass.text
        kmj("sdlt") = ctSDLT.text
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
        
        ' Internal Links
        Do While kmj("kmlinks").count > 0
        'a bug was found here by inspection used to delete cluster rather than kmlinks, fixed but not properly tested
            kmj("kmlinks").Remove 1
        Loop
        
        intlinks = Split(ctIntLinks, vbCrLf)
        For Each lnk In intlinks
            Set l = CreateObject("Scripting.Dictionary")
            lnk = Trim(lnk)
            If Len(lnk > 0) Then
                l("id") = lnk
                kmj("kmlinks").Add l
            End If
        Next lnk
        
        ' External Links
        Set kmj("extlinks") = extLinks
        'Items
        Set kmj("items") = items
        
        ' Keywords
        Do While kmj("keywords").count > 0
            kmj("cluster").Remove 1
        Loop
        keywords = Split(tKeywords, vbCrLf)
        For Each kw In keywords
            Dim kwss As String
            kwss = kw
            kw = cctxt.TrimWhite(CStr(kw))
            If Len(kw) > 0 Then kmj("keywords").Add kw
        Next kw
        
        ' facets
        Set kmj("facets") = facets
        
        muEdit.addKmjMeta kmj:=kmj, article:=Selection.Range
 
        metaForm.Hide
End Sub

Private Sub clearPublic()
    Set facets = Nothing
    Set cfacet = Nothing
    switchFacet = False
End Sub
 
Private Sub disExtLinks()
     ctExtLink = ""
     For Each l In extLinks
        ctExtLink = ctExtLink + l("url") & vbLf
     Next l
End Sub

Private Sub disItems()
    tItems = ""
    For Each i In items
        tItems = tItems & i("item") & vbLf
    Next i
End Sub



Private Sub cbClass_Click()
    classForm.aType = "classes"
    classForm.id = ctClass.text
    classForm.show
    ctClass.text = classForm.id
End Sub

Private Sub cbSDLT_Click()
    classForm.aType = "sdlts"
    classForm.id = ctClass.text
    classForm.show
    ctClass.text = classForm.id
End Sub


Private Sub cbFacets_Click()
    facetsForm.load listFacets
    facetsForm.show
End Sub

Private Sub ctAuthor_Change()

End Sub

Private Sub ctClass_Change()
    ctClass.text = cctxt.cleanUID(ctClass.text)
End Sub

Private Sub ctSDLT_Change()
    ctSDLT.text = cctxt.cleanUID(ctSDLT.text)
End Sub

Private Sub ctFees_Change()
    ctFees.text = cctxt.cleanUID(ctFees.text)
End Sub

Private Sub ctIntLinks_Change()
        txt = ctIntLinks.text
        Dim intlinks() As String
        intlinks = Split(txt, vbCrLf)
        txt = ""
        'x = intlinks.Length
        For i = LBound(intlinks) To UBound(intlinks)
            txt = txt + cctxt.cleanUID(intlinks(i))
            If i <> UBound(intlinks) Then txt = txt + vbCrLf
        Next
        ctIntLinks.text = txt
End Sub

 

Private Sub tClusters_Change()
    tClusters.text = LCase(cctxt.cleanText(tClusters.text))
End Sub

 

Private Sub tID_Change()
    tID.text = cctxt.cleanUID(tID.text)
    tItem = tID.text
End Sub


Private Sub tKeywords_Change()
    tKeywords = cctxt.cleanText(tKeywords.text)
End Sub

Private Sub tScope_Change()
    tScope.text = cctxt.cleanText(tScope, 200)
End Sub


Private Sub tTitle_Change()
    tTitle.text = cctxt.cleanText(tTitle, 70)
End Sub


Private Sub UserForm_Activate()

     
     Dim st As Paragraph
     Dim fin As Paragraph
     
     Dim para As Paragraph
     Dim kmj As Object
     
     Set kmj = Nothing
        
     If Not checkSelection(kmj) Then
        metaForm.Hide
        Exit Sub
     End If
     
          
     'initialise form
      If kmj Is Nothing Then
        Set kmj = muEdit.createKmjObj
      Else
        If kmj("type") <> setupForm.getType Then
            MsgBox ("meta data type is not the same as document type")
            metaForm.Hide
            Exit Sub
        End If
      End If
     
      Set facets = kmj("facets")
      tID.text = kmj("id")
      tFacets.text = listFacets()
      tScope.text = kmj("scope")
      tTitle.text = kmj("title")
      ctClass.text = kmj("class")
      ctSDLT.text = kmj("sdlt")
      ctFees.text = kmj("fees")
      
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
     
     'items
     Set items = kmj("items")
     disItems
     
    
     ' eternal links
     Set extLinks = kmj("extlinks")
     disExtLinks
     
     
     
     ' internal links
     ctIntLinks.text = ""
     del = ""
     For Each lnk In kmj("kmlinks")
        ctIntLinks.text = ctIntLinks.text & del & lnk("id")
        del = vbCrLf
     Next lnk
     
     'keywords
     tKeywords = ""
     del = ""
     For Each kw In kmj("keywords")
        tKeywords = tKeywords & del & kw
        del = vbCrLf
     Next kw
     
    Exit Sub
endLabel:
    metaForm.Hide
End Sub

Private Sub UserForm_Initialize()
     load metaForm
     If (allFacets Is Nothing) Then
        Set allFacets = getFacets()
     End If
     For Each facet In allFacets
        bFacet.AddItem facet
     Next facet
     ccbSensitivity.AddItem "normal"
     ccbSensitivity.AddItem "sensitive"
    
End Sub

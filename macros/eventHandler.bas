Attribute VB_Name = "eventHandler"

Sub testOpen()
    Dim e As New CcEvents
    Dim d As Document
    Set d = ActiveDocument
    e.wdApp_DocumentOpen d
End Sub

Sub testClose()
    Dim e As New CcEvents
    Dim d As Document
    Set d = ActiveDocument
    e.wdApp_DocumentBeforeClose d, False
End Sub

Sub testBeforeSave()
    Dim e As New CcEvents
    Dim d As Document
    Set d = ActiveDocument
    e.wdApp_DocumentBeforeSave d, False, False
End Sub

Sub testSave()
    doSave ActiveDocument
End Sub


'==================

Public Sub autoexec()
    Set evnts.wdApp = Word.Application
End Sub

Public Sub openhandler(doc As Document)
    On Error GoTo errlab
    If Not checkLibrary(doc.path) Then GoTo ngLab
    If InStr(1, doc.name, "kmj.dotm") Then
        docUnlock doc
    Else
        doUnlock doc
    End If
    
    setKmDefaults doc
    doc.Saved = True ' to prevent spurious save on close
    CustomizationContext = doc.AttachedTemplate
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyS, wdKeyControl), KeyCategory:= _
        wdKeyCategoryCommand, Command:="uiSave"
    Application.StatusBar = "Practice guidance document : state = " & getProp(doc, "guide") _
                    & "  cluster = " & setupForm.getCluster _
                    & " type = " & setupForm.getType()
    Exit Sub
ngLab:
    If isGuide(doc) Then
        If checkExport(doc.path) Then Exit Sub
        doWrongLoc doc
        Application.StatusBar = "Guidance document : Invalid location"
    End If
    
    Exit Sub
errlab:
    Application.StatusBar = "Guidance setup failed to complete"
    Exit Sub
End Sub

Sub CloseHandler(doc As Document)

End Sub

Function BeforeSaveHandler(ByVal doc As Document) As Boolean
        BeforeSaveHandler = False
        If isGuide(doc) Then
            If InStr(1, doc.name, "kmj.dotm") Then
                Exit Function
            End If
            Select Case getProp(doc, "guide")
                Case Is = "_LIBR"
                    Exit Function
                Case Is = "OK"
                    Exit Function
                Case Else
                    MsgBox "Please use the save button in the Guidance toolbar to save this document"
                    BeforeSaveHandler = True
                    Exit Function
            End Select
        End If
End Function

'==============




Sub setKmDefaults(doc As Document)
    ' SET DEFAULTS FOR km DOCUMENTS
    doc.TrackRevisions = False
    With doc.ActiveWindow.View
        .ShowRevisionsAndComments = False
        .RevisionsView = wdRevisionsViewFinal
        .FieldShading = wdFieldShadingAlways
    End With
    doc.Activate
    showMeta True
End Sub

Function doSave(ByVal doc As Document) As Boolean
    doSave = False
    
    If isGuide(doc) Then
        If checkLibrary(doc.path) Then
            setProp doc, "hash", BASE64SHA1(doc.name & "_" & Format(Now, "yyyy-mm-ddThh:mm:ss"))
        Else
            MsgBox "You can only save this document to the correct Practice Guidance Library"
            Exit Function ' Cancel Save
        End If
        
        If getProp(doc, "guide") <> "_EDIT" Then
            MsgBox getProp(doc, "guide") & ": Guidance document cannot be saved"
            Exit Function ' Cancel Save
        End If
        
        guard doc
        setSpProp doc, "guide", "OK"
        setSpProps doc
        doc.save
        doUnlock doc
        

    Else
        doc.save
    End If
End Function







 




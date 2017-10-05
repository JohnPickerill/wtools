Attribute VB_Name = "locking"
Function isGuide(doc As Document) As Boolean
    isGuide = (Len(getProp(doc, "guide")) > 0)
End Function
 

 
Function guard(doc As Document)
    On Error GoTo errlab
    guard = True
    If doc.ProtectionType <> wdNoProtection Then
        doc.Unprotect lockkey
    End If
    
    doc.protect Password:=lockkey, NoReset:=False, Type:= _
        wdAllowOnlyReading, UseIRM:=False, EnforceStyleLock:=False

    GoTo endLab
errlab:
    guard = False
endLab:
    CommandBars("Restrict Editing").Visible = False
End Function


Function unguard(doc As Document) As Boolean
    On Error GoTo errlab
    unguard = True
    If doc.ProtectionType = wdAllowOnlyReading Then
        doc.Unprotect lockkey
    Else
        GoTo errlab
    End If
    GoTo endLab
errlab:
    unguard = False
endLab:
    CommandBars("Restrict Editing").Visible = False
End Function
 

Sub docUnlock(ByVal doc As Document)
    ' this function is to allow administrator to unlock any document wherever
    Dim repo As String
    Dim dguide As String
    Set doc = ActiveDocument
    doc.Unprotect lockkey
    setSpProp doc, "guide", "_EDIT"
    doc.Saved = True ' to prevent spurious save on close
End Sub

Sub doUnlock(ByVal doc As Document)
    On Error Resume Next
    dguide = getProp(doc, "guide")
    If Left$(dguide, 1) = "_" Then
         guard doc
         MsgBox ("WARNING:" & dguide & " Practice Document may have data integrity issue" & vbCrLf _
                & "Please consult with the OCDS team before continuing")
    Else
        If unguard(doc) Then
            setSpProp doc, "guide", "_EDIT"
            doc.Saved = True ' to prevent spurious save on close
        Else
            setSpProp doc, "guide", "_LOCK"
            guard doc
            MsgBox ("WARNING Practice Document may have been edited offline" & vbCrLf _
                    & "Please consult with the OCDS team before continuing")
        End If
    End If
End Sub

Sub doWrongLoc(ByVal doc As Document)
    On Error Resume Next
    doc.Unprotect lockkey
    setSpProp doc, "guide", "_LIBR"
    guard doc
    doc.save
    MsgBox "WARNING:" & dguide & " This Guidance document is not in the appropriate practice guidance library." & vbCrLf & _
        "This can cause data integrity issues. Please consult with the ODCS team before continuing"
End Sub

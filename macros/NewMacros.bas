Attribute VB_Name = "NewMacros"
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro3"
'
' Macro3 Macro
'
'
    ActiveDocument.TrackRevisions = Not ActiveDocument.TrackRevisions
    With ActiveWindow.View.RevisionsFilter
        .Markup = wdRevisionsMarkupNone
        .View = wdRevisionsViewFinal
    End With
    With ActiveWindow.View.RevisionsFilter
        .Markup = wdRevisionsMarkupSimple
        .View = wdRevisionsViewFinal
    End With
    With ActiveWindow.View.RevisionsFilter
        .Markup = wdRevisionsMarkupAll
        .View = wdRevisionsViewFinal
    End With
End Sub
Sub Macro5()
Attribute Macro5.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro5"
'
' Macro5 Macro
'
'
    RecentFiles(15).Open
    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    Else
        ActiveWindow.View.Type = wdPrintView
    End If
    ActiveDocument.SaveAs2 FileName:="alg_leases_concur.docx", FileFormat:= _
        wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, CompatibilityMode:=15
End Sub
Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro4"
'
' Macro4 Macro
'
'
    ChangeFileOpenDirectory "C:\Temp\"
    ActiveDocument.SaveAs2 FileName:="alg_leases_concur.docx", FileFormat:= _
        wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, CompatibilityMode:=15
    ActiveWindow.Close
    ActiveWindow.Close
End Sub
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro1"
'
' Macro1 Macro
'
'
    Application.Move Left:=898, Top:=66
    ActiveDocument.Protect Password:="fred", NoReset:=False, Type:= _
        wdAllowOnlyReading, UseIRM:=False, EnforceStyleLock:=False
    CommandBars("Restrict Editing").Visible = False
    Application.Move Left:=1031, Top:=69
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro2"
'
' Macro2 Macro
'
'
End Sub
Sub Macro6()
Attribute Macro6.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Macro6"
'
' Macro6 Macro
'
'
    Selection.Tables(1).Select
    Selection.Tables(1).Delete
End Sub

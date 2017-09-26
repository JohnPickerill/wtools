Attribute VB_Name = "NewMacros"
Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = "JohnPWordTools.NewMacros.Macro4"
'
' Macro4 Macro
'
'
    RecentFiles(23).Open
    ChangeFileOpenDirectory _
        "C:\Users\CS062JP\AppData\Roaming\Microsoft\Windows\Network Shortcuts\"
    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    Else
        ActiveWindow.View.Type = wdPrintView
    End If
    ActiveDocument.SaveAs2 filename:= _
        "stock-letters (2016-07-22 15-26-30).docx", FileFormat:= _
        wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, CompatibilityMode:=15
    ActiveWindow.Close
End Sub

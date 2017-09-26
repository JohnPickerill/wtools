Attribute VB_Name = "Module1"
 
Sub ExportMods()
' reference to extensibility library
 

Set thisProj = Application.VBE.ActiveVBProject
For Each MODULE In thisProj.vbcomponents
    MODULE.Export "C:\DEV\WTOOLS\" & MODULE.name
Next MODULE

 
End Sub

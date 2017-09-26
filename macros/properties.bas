Attribute VB_Name = "properties"
Option Explicit

Sub proptest()
 Dim x As String
 x = getProp(ActiveDocument, "guide")
End Sub
Public Sub setSpProps(doc As Document)
    On Error GoTo notGuide:
    setSpProp doc, "guide", getProp(doc, "guide")
    setSpProp doc, "Cluster", setupForm.getCluster
    setSpProp doc, "Entity Type", setupForm.getType
    setSpProp doc, "Entity Purpose", setupForm.getPurpose
notGuide:
End Sub


Public Function getProp(doc As Document, name As String)
    On Error GoTo errlab
        getProp = doc.CustomDocumentProperties(name)
        Exit Function
errlab:
        getProp = ""
End Function


Public Sub setProp(doc As Document, Prop As String, val As String)
    On Error GoTo addLab:
    doc.CustomDocumentProperties(Prop) = val
addLab:
    On Error Resume Next
    doc.CustomDocumentProperties.Add name:=Prop, value:=val, _
                LinkToContent:=False, Type:=msoPropertyTypeString
End Sub


Sub setSpProp(doc As Document, name As String, value As String)
    On Error GoTo errlab
    setProp doc, name, value
    If Not doc.ContentTypeProperties Is Nothing Then
        If Not doc.ContentTypeProperties(name) Is Nothing Then
            doc.ContentTypeProperties(name).value = getProp(doc, name)
        End If
    End If
errlab:
    
End Sub

Public Function getSpProp(doc As Document, name As String)
    getSpProp = ""
    On Error GoTo errlab
    If Not doc.ContentTypeProperties Is Nothing Then
        If Not doc.ContentTypeProperties(name) Is Nothing Then
            getSpProp = doc.ContentTypeProperties(name).value
        End If
    End If
    Exit Function
errlab:
End Function



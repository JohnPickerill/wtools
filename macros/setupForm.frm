VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} setupForm 
   Caption         =   "km setup"
   ClientHeight    =   4080
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   6924
   OleObjectBlob   =   "setupForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "setupForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ReSpace As New RegExp

Public entityTypes As Object




Private Sub cbType_Change()
    switchType = True
    cbPurpose.clear
    entityType = cbType.text
    For Each p In entityTypes(entityType)
        cbPurpose.AddItem entityTypes(entityType)(p)
    Next p
End Sub


Private Sub cbCancel_Click()
    setupForm.hide
End Sub

Private Sub setCluster()
    On Error GoTo clusterUp
    ActiveDocument.CustomDocumentProperties.Add name:="km_cluster", value:=ctCluster.text, _
          LinkToContent:=False, Type:=msoPropertyTypeString
    Exit Sub
clusterUp:
    ActiveDocument.CustomDocumentProperties("km_cluster") = ctCluster.text
End Sub

Private Sub setType()
    On Error GoTo classUpd
    ActiveDocument.CustomDocumentProperties.Add name:="km_type", value:=cbType.text, _
          LinkToContent:=False, Type:=msoPropertyTypeString
    Exit Sub
classUpd:
    ActiveDocument.CustomDocumentProperties("km_type") = cbType.text
End Sub

Public Function getCode(typ As String, purp As String) As String
    getCode = ""
    For Each p In entityTypes(typ)
        If (entityTypes(typ)(p) = purp) Then
            getCode = p
            Exit For
        End If
    Next p
End Function


Private Sub setPurpose()
    On Error GoTo classUpd
    purpose = getCode(cbType.text, cbPurpose.text)
    ActiveDocument.CustomDocumentProperties.Add name:="km_purpose", value:=purpose, _
          LinkToContent:=False, Type:=msoPropertyTypeString
    Exit Sub
classUpd:
    ActiveDocument.CustomDocumentProperties("km_purpose") = purpose
End Sub

Private Sub cbOk_Click()
    setCluster
    setType
    setPurpose
exitLab:
    ActiveDocument.Saved = False
    setupForm.hide
End Sub
 

Private Sub ctCluster_Change()
 
     ctCluster.text = cctxt.cleanUID(ctCluster.text)
 
End Sub

Public Function getCluster() As String
    getCluster = ""
getLab:
    On Error GoTo errlab
    getCluster = ActiveDocument.CustomDocumentProperties("km_cluster")
    Exit Function
errlab:
    ctCluster.text = cctxt.cleanUID(cleanFilename())
    ActiveDocument.CustomDocumentProperties.Add name:="km_cluster", value:=ctCluster.text, _
          LinkToContent:=False, Type:=msoPropertyTypeString
    'If vbYes = MsgBox("There is no cluster id. Do you want to create one now", vbYesNo) Then
    '    setupForm.show
    '    GoTo getLab
    'End If
End Function

Public Function getType() As String
    Dim str As String
    On Error Resume Next
    str = cbType.List(0)
    str = ActiveDocument.CustomDocumentProperties("km_type")
    'On Error GoTo 0
    getType = str
    cbType.text = str
End Function


Public Function getPurpose() As String
    Dim str As String
    On Error Resume Next
    str = cbPurpose.List(0)
    str = ActiveDocument.CustomDocumentProperties("km_purpose")
    'On Error GoTo 0
    getPurpose = str
    cbPurpose.text = str
End Function


Private Sub UserForm_Activate()
    On Error GoTo errClusLab
    ctCluster.text = ActiveDocument.CustomDocumentProperties("km_cluster")
    clWarning.Caption = ""
    GoTo classLab
errClusLab:
    ctCluster.text = cctxt.cleanUID(cleanFilename())
    clWarning.Caption = "WARNING: The cluster name has not been set yet"
classLab:
    cbType.text = getType()
    cbPurpose.text = getPurpose()
End Sub


Private Sub UserForm_Initialize()
    On Error GoTo errlab
    ReSpace.Global = True
    ReSpace.MultiLine = True
    ReSpace.pattern = "\s"
    load setupForm
    Set entityTypes = getEntityTypes()
    For Each tp In entityTypes
       cbType.AddItem tp
    Next tp
    Exit Sub
errlab:
    MsgBox "error initialising type and purpose lists"
End Sub

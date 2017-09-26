VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} serverForm 
   Caption         =   "Configuration"
   ClientHeight    =   6324
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   8364
   OleObjectBlob   =   "serverForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "serverForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Function fixslash(url As String)
    If (Right(url, 1) <> "/") Then
        fixslash = url + "/"
    Else
        fixslash = url
    End If
End Function


Private Function checkServer()
    
End Function


 
Private Function update() As Boolean
    Dim guide As String
    Dim sharepoint As String
    Dim preview As String
    Dim site As String
    Dim library As String
    
    update = False
    guide = fixslash(ctGuide.text)
    If Not cctxt.testURL(guide) Then
        MsgBox (guide + "is not a valid url")
        Exit Function
    End If
      
    preview = ctPreview.text

    Cfg.serverRead (guide + "api")
    
    sharepoint = fixslash(Cfg.getVar("sharepoint"))

    
    Cfg.setVar "guide", guide
    Cfg.setVar "preview", preview
 
    site = Cfg.getVar("site")
    library = Cfg.getVar("library")
     
    Cfg.setVar "appURL", guide + "ui/" + preview
    Cfg.setVar "cfgURL", guide + "api"
    Cfg.setVar "guideMgr", guide + "api"
    Cfg.setVar "trove", sharepoint + site + "/"
    Cfg.setVar "webDav", cctxt.webdav(sharepoint + site)

    update = True
End Function



Private Sub cbCancel_Click()
    serverForm.hide
End Sub



Private Sub cbReset_Click()
    Cfg.cfgRead
    UserForm_Activate
End Sub

Private Sub cbSave_Click()
    If update() Then
        Cfg.cfgSave
        setView
    End If
End Sub


Private Sub Label4_Click()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub setView()
    tbState.text = "sharepoint: " + Cfg.getVar("sharepoint") + vbCrLf
    tbState.text = tbState.text + "site:    " + Cfg.getVar("site") + vbCrLf
    tbState.text = tbState.text + "library: " + Cfg.getVar("library") + vbCrLf
    tbState.text = tbState.text + "checkpoint: " + Cfg.getVar("checkpoint") + vbCrLf + vbCrLf
    tbState.text = tbState.text + "config:  " + Cfg.getVar("cfgURL") + vbCrLf
    tbState.text = tbState.text + "manager: " + Cfg.getVar("guideMgr") + vbCrLf
    tbState.text = tbState.text + "assets:  " + Cfg.getVar("images") + vbCrLf
    tbState.text = tbState.text + "docs:    " + Cfg.getVar("trove") + vbCrLf
    tbState.text = tbState.text + "webdav:  " + Cfg.getVar("webDav") + vbCrLf
End Sub


Private Sub UserForm_Activate()
    ctGuide.text = Cfg.getVar("guide")
    ctPreview.text = Cfg.getVar("preview")
    Cfg.serverRead (Cfg.getVar("cfgURL"))
    setView
    tbDoc.text = getProp(ActiveDocument, "guide") & vbCrLf _
                & ActiveDocument.path & vbCrLf _
                & ActiveDocument.name
End Sub

 

Attribute VB_Name = "RibbonControl"
Option Explicit
Public KmRibbon As IRibbonUI
Private bAll As Boolean
Private bTags As Boolean

Private spanStyles() As String
Private blockStyles() As String
Private scnt As Integer
Private bcnt As Integer
Public spanIndex As String
Public blockIndex As String
Public kmStyles  As Object




Sub setStyles()
Dim k As Variant
Dim s As Object

If kmStyles Is Nothing Then
    Set kmStyles = getStyles()
    ReDim blockStyles(0 To kmStyles.count - 1) As String
    ReDim spanStyles(0 To kmStyles.count - 1) As String
    scnt = 0
    bcnt = 0
    For Each k In kmStyles.keys()
         Set s = kmStyles(k)
         If s("block") <> "" Then
                 blockStyles(bcnt) = k
                 bcnt = bcnt + 1
         End If
         If s("span") <> "" Then
                 spanStyles(scnt) = k
                 scnt = scnt + 1
         End If
    Next k
End If
End Sub

Sub Onload(ribbon As IRibbonUI)
'Create a ribbon instance for use in this project
Set KmRibbon = ribbon
setStyles
End Sub
'Callback for DropDown onAction


 
'Callback for DropDown GetItemCount
Sub GetItemCount(ByVal Control As IRibbonControl, ByRef count)
  'Tell the ribbon to show 4 items in the dropdown
  setStyles
  If Control.id = "KmSpanDD" Then
      count = scnt
  Else
      count = bcnt
  End If
End Sub

'Callback for DropDown GetItemLabel
Sub GetItemLabel(ByVal Control As IRibbonControl, Index As Integer, ByRef label)
  'This procedure fires once for each item in the dropdown. Index is _
   received as 0, 1, 2, etc. and label is returned.
    If Control.id = "KmSpanDD" Then
        label = spanStyles(Index)
    Else
        label = blockStyles(Index)
    End If
End Sub

'Callback DropDown GetSelectedIndex
Sub GetSelectedItemIndex(ByVal Control As IRibbonControl, ByRef Index)
  'This procedure is used to ensure the first item in the dropdown is selected _
   when the control is displayed
    Index = 0
    If Control.id = "KmSpanDD" Then
        spanIndex = spanStyles(0)
    Else
        blockIndex = blockStyles(0)
    End If
End Sub

'Callback for DropDown onAction
Sub myDDMacro(ByVal Control As IRibbonControl, selectedID As String, selectedIndex As Integer)
  Select Case Control.id
    Case Is = "KmSpanDD"
      spanIndex = spanStyles(selectedIndex)
    Case Else
      blockIndex = blockStyles(selectedIndex)
  End Select
  'KmRibbon.InvalidateControl control.id
End Sub


'Callback for Button onAction
Sub CcBtn(ByVal Control As IRibbonControl)
Select Case Control.id
  Case Is = "kmClean"
        uiCleanDoc
  Case Is = "kmUnlock"
        uiUnlock
  Case Is = "kmMetaWip"
        pushES "wip"
  Case Is = "kmMetaAdvance"
        pushES "advance"
  Case Is = "kmMetaLive"
        pushES "live"
  Case Is = "kmAdd"
        displayMeta
  Case Is = "kmImages"
        imgForm.show
  Case Is = "kmPubStatus"
        pubStatus
  Case Is = "kmPubList"
        saveAllDirectories True
  Case Is = "kmPubExport"
        saveAllDirectories False
  Case Is = "kmPubSync"
        syncWip
  Case Is = "kmPubAdvance"
        promoteWip
  Case Is = "kmPubLive"
        promoteAdvance
  Case Is = "kmSave"
        uiSave
  Case Is = "kmOpen"
        uiOpen
  Case Is = "kmNew"
        uiNew
  Case Is = "kmPreview"
        previewJson
  Case Is = "kmMarkup"
        uImarkup
  Case Is = "kmOutline"
        outline
  Case Is = "kmDrop"
        createDrop
  Case Is = "kmLink"
        uILink
  Case Is = "kmHideText"
        hideText
  Case Is = "kmExtraText"
        extraText
  Case Is = "kmUnhideText"
        uIunhide
  Case Is = "kmServers"
        cfgServer
  Case Is = "kmSel"
        uISel
  Case Is = "kmNext"
        uINext
  Case Is = "kmPrev"
        uIPrev
  Case Is = "kmDefaults"
        uISetup
  Case Is = "kmHelp"
        uIHelp
  Case Is = "kmApparate"
        uIApparate
  Case Is = "kmAquire"
        uIAcquire
  Case Is = "kmGhost"
        uIGhost
  Case Is = "kmSpan"
        uiSpan
  Case Is = "kmBlock"
        uiBlock
  Case Is = "kmCompare"
    uiCompare
 
  Case Is = "kmReview"
    uiReview
 
 
  Case Is = "kmSnippet"
    uiSnippet
  Case Is = "kmAnchor"
    uiAnchor
  Case Is = "kmArticleNoTags"
    showArticleMeta False
  Case Is = "kmArticleTags"
    showArticleMeta True
  Case Is = "kmHdConvert"
    hdConvert
  
  End Select
End Sub
'Callback for Toogle onAction
Sub ToggleonAction(Control As IRibbonControl, pressed As Boolean)
Select Case Control.id
  Case Is = "kmShowAll"
    bAll = pressed
    showAll pressed
  Case Is = "kmShowTags"
    bTags = pressed
    showMeta Not bTags
End Select
'Force the ribbon to redefine the control wiht correct image and label
On Error Resume Next
KmRibbon.InvalidateControl (Control.id)
End Sub
'Callback for togglebutton getLabel
Sub getLabel(Control As IRibbonControl, ByRef returnedVal)
Select Case Control.id

  Case Is = "kmShowAll"
    If Not bAll Then
      returnedVal = "Show All"
    Else
      returnedVal = "Hide All"
    End If
  Case Is = "kmShowTags"
    If Not bTags Then
      returnedVal = "Hide Tags"
    Else
      returnedVal = "Show Tags"
    End If
End Select
End Sub

'Callback for togglebutton getImage
Sub GetImage(Control As IRibbonControl, ByRef returnedVal)
Select Case Control.id


  Case Is = "kmShowAll"
   If bAll Then
      returnedVal = "_3DTiltRightClassic"
    Else
      returnedVal = "_3DTiltLeftClassic"
   End If
  Case Is = "kmShowTags"
    If Not bTags Then
      returnedVal = "SlideShowInAWindow"
    Else
      returnedVal = "WebControlHidden"
    End If
End Select
End Sub
'Callback for togglebutton getPressed
Sub buttonPressed(Control As IRibbonControl, ByRef toggleState)
'toggleState is used tp set the toggle state (i.e., true or false) and determine how the
'toggle appears on the ribbon (i.e., flusn or sunken).
Select Case Control.id
  Case Is = "kmShowAll"
      toggleState = bAll
 
  Case Is = "kmShowTags"
      toggleState = bTags
End Select
End Sub


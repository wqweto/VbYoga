VERSION 5.00
Begin VB.UserControl ctxFlexContainer 
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ScaleHeight     =   2880
   ScaleWidth      =   3840
   Begin Project1.ctxNineButton btnButton 
      Height          =   600
      Index           =   0
      Left            =   588
      TabIndex        =   1
      Top             =   336
      Visible         =   0   'False
      Width           =   1524
      _ExtentX        =   2688
      _ExtentY        =   1058
      Caption         =   "ctxNineButton1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label labLabel 
      Caption         =   "labLabel"
      Height          =   264
      Index           =   0
      Left            =   756
      TabIndex        =   0
      Top             =   1428
      Visible         =   0   'False
      Width           =   1272
   End
End
Attribute VB_Name = "ctxFlexContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' VbYoga (c) 2019 by wqweto@gmail.com
'
' Facebook's Yoga bindings for VB6. Implements CSS Flexbox layout
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "ctxFlexContainer"

'=========================================================================
' Events
'=========================================================================

Event Click(DomNode As cFlexDomNode)

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_oStyles               As Object
Private m_oRoot                 As cFlexDomNode
Private m_oYogaConfig           As cYogaConfig
Private m_lButtonCount          As Long
Private m_lLabelCount           As Long
Private m_cMapping              As Collection
'--- debug
#If DebugMode Then
    Private m_sDebugID          As String
#End If

'=========================================================================
' Properties
'=========================================================================

Property Get Styles() As Object
    Set m_oStyles = m_oStyles
End Property

Property Set Styles(oValue As Object)
    Set m_oStyles = oValue
    pvApplyStyles m_oRoot
End Property

Property Get Root() As cFlexDomNode
    Set Root = m_oRoot
End Property

Property Get frYogaConfig() As cYogaConfig
    Set frYogaConfig = m_oYogaConfig
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Reset()
    For m_lButtonCount = m_lButtonCount To 1 Step -1
        btnButton(m_lButtonCount).Visible = False
    Next
    For m_lLabelCount = m_lLabelCount To 1 Step -1
        labLabel(m_lLabelCount).Visible = False
    Next
    Set m_oRoot = New cFlexDomNode
    #If DebugMode Then
        Debug.Print "YogaNodeInstanceCount=" & YogaNodeInstanceCount(), Timer
    #End If
    Set m_oRoot.Layout = YogaNodeNew(m_oYogaConfig)
    Set m_oRoot.frFlexBox = Me
    m_oRoot.CssClass = "root container"
    Set m_cMapping = New Collection
End Function

'= friend ================================================================

Friend Function frLoadButton() As VBControlExtender
    m_lButtonCount = m_lButtonCount + 1
    If btnButton.UBound < m_lButtonCount Then
        Load btnButton(m_lButtonCount)
    End If
    Set frLoadButton = btnButton(m_lButtonCount)
End Function

Friend Function frLoadLabel() As VB.Label
    m_lLabelCount = m_lLabelCount + 1
    If labLabel.UBound < m_lLabelCount Then
        Load labLabel(m_lLabelCount)
    End If
    Set frLoadLabel = labLabel(m_lLabelCount)
End Function

Friend Sub frAddDomNodeMapping(oDomNode As cFlexDomNode, oCtl As Object)
    m_cMapping.Add oDomNode, "#" & ObjPtr(oCtl)
End Sub

'= private ===============================================================

Private Sub pvApplyStyles(oDomNode As cFlexDomNode)
    Dim oStyle          As Object
    Dim oItem           As cFlexDomNode
    Dim vKey            As Variant
    Dim vSplit          As Variant
    Dim vValue          As Variant
    
    Set oStyle = pvGetStyle(oDomNode.Name, oDomNode.CssClass, TypeName(oDomNode.Control))
    With oDomNode.Layout
        For Each vKey In oStyle.Keys
            vValue = oStyle.Item(vKey)
            Select Case LCase$(vKey)
            Case "width"
                .Width = pvToYogaValue(vValue)
            Case "height"
                .Height = pvToYogaValue(vValue)
            Case "min-width"
                .MinWidth = pvToYogaValue(vValue)
            Case "min-height"
                .MinHeight = pvToYogaValue(vValue)
            Case "max-width"
                .MinWidth = pvToYogaValue(vValue)
            Case "max-height"
                .MinHeight = pvToYogaValue(vValue)
            Case "direction"
                Select Case LCase$(vValue)
                Case "ltr"
                    .StyleDirection = yogaDirLTR
                Case "rtl"
                    .StyleDirection = yogaDirRTL
                Case Else
                    #If DebugMode Then
                        Debug.Print "Unknown value for '" & vKey & "': " & vValue
                    #End If
                End Select
            Case "position"
                Select Case LCase$(vValue)
                Case "absolute"
                    .PositionType = yogaPosAbsolute
                Case "relative"
                    .PositionType = yogaPosRelative
                Case Else
                    #If DebugMode Then
                        Debug.Print "Unknown value for '" & vKey & "': " & vValue
                    #End If
                End Select
            Case "display"
                Select Case LCase$(vValue)
                Case "none"
                    .Display = yogaDisplayNone
                Case "flex"
                    .Display = yogaDisplayFlex
                Case Else
                    #If DebugMode Then
                        Debug.Print "Unknown value for '" & vKey & "': " & vValue
                    #End If
                End Select
            Case "overflow"
                Select Case LCase$(vValue)
                Case "hidden"
                    .Overflow = yogaOverflowHidden
                Case "visible"
                    .Overflow = yogaOverflowVisible
                Case "scroll"
                    .Overflow = yogaOverflowScroll
                Case Else
                    #If DebugMode Then
                        Debug.Print "Unknown value for '" & vKey & "': " & vValue
                    #End If
                End Select
            Case "flex-direction"
                Select Case LCase$(vValue)
                Case "row"
                    .FlexDirection = yogaFlexRow
                Case "row-reverse"
                    .FlexDirection = yogaFlexRowReverse
                Case "column"
                    .FlexDirection = yogaFlexColumn
                Case "column-reverse"
                    .FlexDirection = yogaFlexColumnReverse
                Case Else
                    #If DebugMode Then
                        Debug.Print "Unknown value for '" & vKey & "': " & vValue
                    #End If
                End Select
            Case "flex-wrap"
                Select Case LCase$(vValue)
                Case "nowrap"
                    .Wrap = yogaWrapNoWrap
                Case "wrap"
                    .Wrap = yogaWrapWrap
                Case "wrap-reverse"
                    .Wrap = yogaWrapWrapReverse
                Case Else
                    #If DebugMode Then
                        Debug.Print "Unknown value for '" & vKey & "': " & vValue
                    #End If
                End Select
            Case "flex-flow"
                vSplit = Split(LCase$(vValue), " ")
                Select Case vSplit(0)
                Case "row"
                    .FlexDirection = yogaFlexRow
                Case "row-reverse"
                    .FlexDirection = yogaFlexRowReverse
                Case "column"
                    .FlexDirection = yogaFlexColumn
                Case "column-reverse"
                    .FlexDirection = yogaFlexColumnReverse
                Case Else
                    #If DebugMode Then
                        Debug.Print "Unknown value for '" & vKey & "': " & vValue
                    #End If
                End Select
                Select Case vSplit(1)
                Case "nowrap"
                    .Wrap = yogaWrapNoWrap
                Case "wrap"
                    .Wrap = yogaWrapWrap
                Case "wrap-reverse"
                    .Wrap = yogaWrapWrapReverse
                Case Else
                    #If DebugMode Then
                        Debug.Print "Unknown value for '" & vKey & "': " & vValue
                    #End If
                End Select
            Case "justify-content"
                Select Case LCase$(vValue)
                Case "flex-start"
                    .JustifyContent = yogaJustFlexStart
                Case "flex-end"
                    .JustifyContent = yogaJustFlexEnd
                Case "center"
                    .JustifyContent = yogaJustCenter
                Case "space-between"
                    .JustifyContent = yogaJustSpaceBetween
                Case "space-around"
                    .JustifyContent = yogaJustSpaceAround
                Case "space-evenly"
                    .JustifyContent = yogaJustSpaceEvenly
                    #If DebugMode Then
                        Debug.Print "Unknown value for '" & vKey & "': " & vValue
                    #End If
                End Select
            Case "align-items"
                Select Case LCase$(vValue)
                Case "flex-start"
                    .AlignItems = yogaAlignFlexStart
                Case "flex-end"
                    .AlignItems = yogaAlignFlexEnd
                Case "center"
                    .AlignItems = yogaAlignCenter
                Case "stretch"
                    .AlignItems = yogaAlignStretch
                Case "baseline"
                    .AlignItems = yogaAlignBaseline
                    #If DebugMode Then
                        Debug.Print "Unknown value for '" & vKey & "': " & vValue
                    #End If
                End Select
            Case "align-content"
                Select Case LCase$(vValue)
                Case "flex-start"
                    .AlignContent = yogaAlignFlexStart
                Case "flex-end"
                    .AlignContent = yogaAlignFlexEnd
                Case "center"
                    .AlignContent = yogaAlignCenter
                Case "stretch"
                    .AlignContent = yogaAlignStretch
                Case "space-between"
                    .AlignContent = yogaAlignSpaceBetween
                Case "space-around"
                    .AlignContent = yogaAlignSpaceAround
                Case Else
                    #If DebugMode Then
                        Debug.Print "Unknown value for '" & vKey & "': " & vValue
                    #End If
                End Select
            Case "align-self"
                Select Case LCase$(vValue)
                Case "auto"
                    .AlignSelf = yogaAlignAuto
                Case "flex-start"
                    .AlignSelf = yogaAlignFlexStart
                Case "flex-end"
                    .AlignSelf = yogaAlignFlexEnd
                Case "center"
                    .AlignSelf = yogaAlignCenter
                Case "baseline"
                    .AlignSelf = yogaAlignBaseline
                Case "stretch"
                    .AlignSelf = yogaAlignStretch
                Case Else
                    #If DebugMode Then
                        Debug.Print "Unknown value for '" & vKey & "': " & vValue
                    #End If
                End Select
            Case "flex"
                .Flex = Val(vValue)
            Case "flex-grow":
                .FlexGrow = Val(vValue)
            Case "flex-shrink":
                .FlexShrink = Val(vValue)
            Case "flex-basic":
                .FlexShrink = pvToYogaValue(vValue)
            Case "aspect-ratio"
                .AspectRatio = Val(vValue)
            '--- spacing
            Case "left"
                .Left = pvToYogaValue(vValue)
            Case "top"
                .Top = pvToYogaValue(vValue)
            Case "right"
                .Right = pvToYogaValue(vValue)
            Case "bottom"
                .Bottom = pvToYogaValue(vValue)
            Case "start"
                .Start = pvToYogaValue(vValue)
            Case "end"
                .End_ = pvToYogaValue(vValue)
            Case "margin"
                .Margin = pvToYogaValue(vValue)
            Case "margin-left"
                .MarginLeft = pvToYogaValue(vValue)
            Case "margin-top"
                .MarginTop = pvToYogaValue(vValue)
            Case "margin-right"
                .MarginRight = pvToYogaValue(vValue)
            Case "margin-bottom"
                .MarginBottom = pvToYogaValue(vValue)
            Case "margin-horizontal"
                .MarginHorizontal = pvToYogaValue(vValue)
            Case "margin-vertical"
                .MarginVertical = pvToYogaValue(vValue)
            Case "margin-start"
                .MarginStart = pvToYogaValue(vValue)
            Case "margin-end"
                .MarginEnd = pvToYogaValue(vValue)
            Case "padding"
                .Padding = pvToYogaValue(vValue)
            Case "padding-left"
                .PaddingLeft = pvToYogaValue(vValue)
            Case "padding-top"
                .PaddingTop = pvToYogaValue(vValue)
            Case "padding-right"
                .PaddingRight = pvToYogaValue(vValue)
            Case "padding-bottom"
                .PaddingBottom = pvToYogaValue(vValue)
            Case "padding-horizontal"
                .PaddingHorizontal = pvToYogaValue(vValue)
            Case "padding-vertical"
                .PaddingVertical = pvToYogaValue(vValue)
            Case "padding-start"
                .PaddingStart = pvToYogaValue(vValue)
            Case "padding-end"
                .PaddingEnd = pvToYogaValue(vValue)
            Case "border-left", "border-left-width"
                .BorderLeftWidth = Val(vValue)
            Case "border-top", "border-top-width"
                .BorderTopWidth = Val(vValue)
            Case "border-right", "border-right-width"
                .BorderRightWidth = Val(vValue)
            Case "border-bottom", "border-bottom-width"
                .BorderBottomWidth = Val(vValue)
            Case "border-start", "border-start-width"
                .BorderStartWidth = Val(vValue)
            Case "border-end", "border-end-width"
                .BorderEndWidth = Val(vValue)
            Case "border", "border-width"
                .BorderWidth = Val(vValue)
            Case Else
                #If DebugMode Then
                    Debug.Print "Unknown style: " & vKey
                #End If
            End Select
        Next
    End With
    If oDomNode.Count > 0 Then
        For Each oItem In oDomNode
            pvApplyStyles oItem
        Next
    End If
End Sub

Private Function pvToYogaValue(ByVal sValue As String) As Variant
    Select Case LCase$(sValue)
    Case "auto"
        pvToYogaValue = Array(yogaUnitAuto)
    Case "undefined"
        pvToYogaValue = Array(yogaUnitUndefined)
    Case Else
        If Right$(Trim$(sValue), 1) = "%" Then
            pvToYogaValue = Array(yogaUnitPercent, Val(sValue))
        Else
            pvToYogaValue = Val(sValue)
        End If
    End Select
End Function

Private Function pvGetStyle(CtlName As String, CssClass As String, CtlType As String) As Object
    Dim oRetVal         As Object
    Dim vElem           As Variant
    Dim oCache          As Object
    Dim sKey            As String
    
    If LenB(CtlName) <> 0 Then
        sKey = "#" & CtlName
    Else
        sKey = IIf(LenB(CtlType) <> 0, CtlType, vbNullString) & IIf(LenB(CssClass) <> 0, "." & CssClass, vbNullString)
    End If
    Set oRetVal = pvTryGetCache(sKey)
    If oRetVal Is Nothing Then
        Set oRetVal = CreateObject("Scripting.Dictionary")
        Set oCache = pvTryGetCache("#" & CtlName)
        If oCache Is Nothing Then
            If m_oStyles.Exists("#" & CtlName) Then
                Set oCache = m_oStyles.Item("#" & CtlName)
            Else
                Set oCache = pvEmptyStyle
            End If
            pvSetCache "#" & CtlName, oCache
        End If
        pvMergeStyle oRetVal, oCache
        For Each vElem In Split(CssClass)
            Set oCache = pvTryGetCache("." & vElem)
            If oCache Is Nothing Then
                If m_oStyles.Exists("." & vElem) Then
                    Set oCache = m_oStyles.Item("." & vElem)
                Else
                    Set oCache = pvEmptyStyle
                End If
                pvSetCache "." & vElem, oCache
            End If
            pvMergeStyle oRetVal, oCache
        Next
        Set oCache = pvTryGetCache(CtlType)
        If oCache Is Nothing Then
            If m_oStyles.Exists(CtlType) Then
                Set oCache = m_oStyles.Item(CtlType)
            Else
                Set oCache = pvEmptyStyle
            End If
            pvSetCache CtlType, oCache
        End If
        pvMergeStyle oRetVal, oCache
        pvSetCache sKey, oRetVal
    End If
    Set pvGetStyle = oRetVal
End Function

Private Function pvTryGetCache(sKey As String) As Object
    If m_oStyles.Exists("__cache__" & sKey) Then
        Set pvTryGetCache = m_oStyles.Item("__cache__" & sKey)
    End If
End Function

Private Function pvSetCache(sKey As String, oValue As Object)
    If oValue Is Nothing Then
        m_oStyles.Remove "__cache__" & sKey
    Else
        Set m_oStyles.Item("__cache__" & sKey) = oValue
    End If
End Function

Private Function pvEmptyStyle() As Object
    Static oEmpty       As Object
    
    If oEmpty Is Nothing Then
        Set oEmpty = CreateObject("Scripting.Dictionary")
    End If
    Set pvEmptyStyle = oEmpty
End Function

Private Sub pvMergeStyle(oDest As Object, oSrc As Object)
    Dim vElem           As Variant
    
    For Each vElem In oSrc.Keys
        If Not oDest.Exists(vElem) Then
            oDest.Item(vElem) = oSrc.Item(vElem)
        End If
    Next
End Sub

'=========================================================================
' Events
'=========================================================================

Private Sub btnButton_Click(Index As Integer)
    RaiseEvent Click(m_cMapping.Item("#" & ObjPtr(btnButton(Index))))
End Sub

Private Sub labLabel_Click(Index As Integer)
    RaiseEvent Click(m_cMapping.Item("#" & ObjPtr(labLabel(Index))))
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If Ambient.UserMode Then
        Set m_oYogaConfig = YogaConfigNew()
        m_oYogaConfig.PointScaleFactor = 1# / Screen.TwipsPerPixelX
        m_oYogaConfig.UseWebDefaults = True
        Set m_oRoot = New cFlexDomNode
        Set m_oRoot.Layout = YogaNodeNew(m_oYogaConfig)
        Set m_oRoot.frFlexBox = Me
        m_oRoot.CssClass = "root container"
        Set m_cMapping = New Collection
    End If
End Sub

Private Sub UserControl_Resize()
    If Not m_oRoot Is Nothing Then
        m_oRoot.Layout.CalculateLayout Width, Height
        m_oRoot.ApplyLayout
    End If
End Sub

'=========================================================================
' Base class events
'=========================================================================

#If DebugMode Then
    Private Sub UserControl_Initialize()
        DebugInstanceInit MODULE_NAME, m_sDebugID, Me
    End Sub
#End If

Private Sub UserControl_Terminate()
    Set m_oRoot = Nothing
    Set m_cMapping = Nothing
    Set m_oYogaConfig = Nothing
    #If DebugMode Then
        DebugInstanceTerm MODULE_NAME, m_sDebugID
    #End If
End Sub

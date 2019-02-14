VERSION 5.00
Begin VB.UserControl ctxFlexContainer 
   BackColor       =   &H8000000D&
   BackStyle       =   0  'Transparent
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   HitBehavior     =   0  'None
   KeyPreview      =   -1  'True
   ScaleHeight     =   2880
   ScaleWidth      =   3840
   Windowless      =   -1  'True
   Begin Project1.ctxNineButton btnButton 
      Height          =   684
      Index           =   0
      Left            =   672
      TabIndex        =   1
      Top             =   504
      Visible         =   0   'False
      Width           =   1356
      _ExtentX        =   2392
      _ExtentY        =   1207
      Caption         =   "btnButton"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label labLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "labLabel"
      Height          =   192
      Index           =   0
      Left            =   756
      TabIndex        =   0
      Top             =   1428
      Visible         =   0   'False
      Width           =   564
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
Private Const STR_MODULE_NAME As String = "ctxFlexContainer"

#Const ImplUseShared = VBYOGA_USE_SHARED <> 0

'=========================================================================
' Events
'=========================================================================

Event Click(DomNode As cFlexDomNode)
Event OwnerDraw(DomNode As cFlexDomNode, ByVal hGraphics As Long, ByVal hFont As Long, ByVal ButtonState As UcsNineButtonStateEnum, ClientLeft As Long, ClientTop As Long, ClientWidth As Long, ClientHeight As Long, Caption As String, ByVal hPicture As Long)
Event RegisterCancelMode(oCtl As Object, Handled As Boolean)
Event StyleCustomProperty(DomNode As cFlexDomNode, Key As String, Value As Variant)
Event AccessKeyPress(DomNode As cFlexDomNode, KeyAscii As Integer)
Event StartDrag(Button As Integer, Shift As Integer, X As Single, Y As Single, Handled As Boolean)
Event MouseDown(DomNode As cFlexDomNode, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(DomNode As cFlexDomNode, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(DomNode As cFlexDomNode, Button As Integer, Shift As Integer, X As Single, Y As Single)

'=========================================================================
' API
'=========================================================================

'Private Declare Function ApiUpdateWindow Lib "user32" Alias "UpdateWindow" (ByVal hWnd As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const FLOAT_UNDEFINED       As Single = 3.40282347E+38
Private Const DEF_AUTOAPPLYLAYOUT   As Boolean = False
Private Const DEF_ENABLED           As Boolean = True
Private Const LNG_DRAG_DISTANCE     As Long = 16

Private WithEvents m_oFont      As StdFont
Attribute m_oFont.VB_VarHelpID = -1
Private m_bAutoApplyLayout      As Boolean
'--- run-time
Private m_oStyles               As Object
Private m_oCache                As Object
Private m_oRoot                 As cFlexDomNode
Private m_oYogaConfig           As cYogaConfig
Private m_lButtonCount          As Long
Private m_lLabelCount           As Long
Private m_cMapping              As Collection
Private m_oCtlCancelMode        As Object
Private m_nDownButton           As Integer
Private m_nDownShift            As Integer
Private m_sngDownX              As Single
Private m_sngDownY              As Single
Private m_bDragging             As Boolean
'--- debug
Private m_sInstanceName         As String
#If DebugMode Then
    Private m_sDebugID          As String
#End If

'=========================================================================
' Error handling
'=========================================================================

Friend Function frInstanceName() As String
    frInstanceName = m_sInstanceName
End Function

Private Property Get MODULE_NAME() As String
#If ImplUseShared Then
    #If DebugMode Then
        MODULE_NAME = GetModuleInstance(STR_MODULE_NAME, frInstanceName, m_sDebugID)
    #Else
        MODULE_NAME = GetModuleInstance(STR_MODULE_NAME, frInstanceName)
    #End If
#Else
    MODULE_NAME = STR_MODULE_NAME
#End If
End Property

Private Function PrintError(sFunction As String) As VbMsgBoxResult
#If ImplUseShared Then
    PopPrintError sFunction, MODULE_NAME, PushError
#Else
    Debug.Print "Critical error: " & Err.Description & " [" & STR_MODULE_NAME & "." & sFunction & "]", Timer
#End If
End Function

'=========================================================================
' Properties
'=========================================================================

Property Get Font() As StdFont
Attribute Font.VB_UserMemId = -512
    Set Font = m_oFont
End Property

Property Set Font(oValue As StdFont)
    If Not m_oFont Is oValue Then
        Set m_oFont = oValue
        m_oFont_FontChanged vbNullString
        PropertyChanged
    End If
End Property

Property Get AutoApplyLayout() As Boolean
    AutoApplyLayout = m_bAutoApplyLayout
End Property

Property Let AutoApplyLayout(ByVal bValue As Boolean)
    If m_bAutoApplyLayout <> bValue Then
        m_bAutoApplyLayout = bValue
        PropertyChanged
        If Not m_oRoot Is Nothing And bValue Then
            ApplyLayout
        End If
    End If
End Property

Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Property Let Enabled(ByVal bValue As Boolean)
    If UserControl.Enabled <> bValue Then
        UserControl.Enabled = bValue
    End If
    PropertyChanged
End Property

'= run-time ==============================================================

Property Get Styles() As Object
    Set Styles = m_oStyles
End Property

Property Set Styles(oValue As Object)
    If oValue Is Nothing Then
        Set m_oStyles = CreateObject("Scripting.Dictionary")
    Else
        Set m_oStyles = oValue
    End If
    Set m_oCache = CreateObject("Scripting.Dictionary")
    pvApplyStyles m_oRoot
End Property

Property Get Root() As cFlexDomNode
    Set Root = m_oRoot
End Property

Property Get Dragging() As Boolean
    Dragging = m_bDragging
End Property

Property Get DownX() As Single
    DownX = m_sngDownX
End Property

Property Get DownY() As Single
    DownY = m_sngDownY
End Property

Property Get frYogaConfig() As cYogaConfig
    Set frYogaConfig = m_oYogaConfig
End Property

'=========================================================================
' Methods
'=========================================================================

Public Sub Reset()
    For m_lButtonCount = m_lButtonCount To 1 Step -1
        btnButton.Item(m_lButtonCount).Visible = False
    Next
    For m_lLabelCount = m_lLabelCount To 1 Step -1
        labLabel.Item(m_lLabelCount).Visible = False
    Next
    pvInitRoot
End Sub

Public Sub ApplyLayout( _
            Optional ByVal sngWidth As Single = FLOAT_UNDEFINED, _
            Optional ByVal sngHeight As Single = FLOAT_UNDEFINED, _
            Optional ByVal sngClipLeft As Single = FLOAT_UNDEFINED, _
            Optional ByVal sngClipTop As Single = FLOAT_UNDEFINED, _
            Optional ByVal sngClipWidth As Single = FLOAT_UNDEFINED, _
            Optional ByVal sngClipHeight As Single = FLOAT_UNDEFINED)
    m_oRoot.Layout.CalculateLayout sngWidth, sngHeight
    If sngClipLeft = FLOAT_UNDEFINED Or sngClipTop = FLOAT_UNDEFINED Then
        m_oRoot.ApplyLayout
    Else
        If sngClipWidth = FLOAT_UNDEFINED Then
            sngClipWidth = ScaleWidth
        End If
        If sngClipHeight = FLOAT_UNDEFINED Then
            sngClipHeight = ScaleHeight
        End If
        m_oRoot.ApplyLayout sngClipLeft, sngClipTop, sngClipWidth, sngClipHeight
    End If
End Sub

Public Sub CancelMode()
    If Not m_oCtlCancelMode Is Nothing Then
        m_oCtlCancelMode.CancelMode
        Set m_oCtlCancelMode = Nothing
    End If
End Sub

Public Sub Refresh()
    UserControl.Refresh
End Sub

Public Sub Repaint()
    Dim lIdx            As Long
    
    For lIdx = 1 To m_lButtonCount
        btnButton.Item(lIdx).Refresh
    Next
    For lIdx = 1 To m_lLabelCount
        labLabel.Item(lIdx).Refresh
    Next
    UserControl.Refresh
'    Call ApiUpdateWindow(ContainerHwnd)
End Sub

'= friend ================================================================

Friend Function frLoadButton() As VBControlExtender
    m_lButtonCount = m_lButtonCount + 1
    If btnButton.UBound < m_lButtonCount Then
        Load btnButton.Item(m_lButtonCount)
        Set btnButton.Item(m_lButtonCount).Font = m_oFont
    End If
    Set frLoadButton = btnButton.Item(m_lButtonCount)
    Set frLoadButton.Object.Font = m_oFont
    frLoadButton.ForeColor = ForeColor
End Function

Friend Function frLoadLabel() As VB.Label
    m_lLabelCount = m_lLabelCount + 1
    If labLabel.UBound < m_lLabelCount Then
        Load labLabel.Item(m_lLabelCount)
        Set labLabel.Item(m_lLabelCount).Font = m_oFont
    End If
    Set frLoadLabel = labLabel.Item(m_lLabelCount)
    Set frLoadLabel.Font = m_oFont
    frLoadLabel.ForeColor = ForeColor
    frLoadLabel.UseMnemonic = False
End Function

Friend Sub frAddDomNodeMapping(oDomNode As cFlexDomNode, oCtl As Object)
    m_cMapping.Add oDomNode, "#" & ObjPtr(oCtl)
End Sub

'= private ===============================================================

Private Sub pvInitRoot()
    Set m_oYogaConfig = YogaConfigNew()
    m_oYogaConfig.PointScaleFactor = 1# / ScreenTwipsPerPixelX
    Set m_oRoot = New cFlexDomNode
    Set m_oRoot.Layout = YogaNodeNew(m_oYogaConfig)
    Set m_oRoot.frFlexBox = Me
    m_oRoot.CssClass = "root container"
    Set m_cMapping = New Collection
End Sub

Private Sub pvApplyStyles(oDomNode As cFlexDomNode)
    Const FUNC_NAME     As String = "pvApplyStyles"
    Const STR_PROP_PREFIX As String = "property-"
    Dim oStyle          As Object
    Dim oItem           As cFlexDomNode
    Dim vKey            As Variant
    Dim vSplit          As Variant
    Dim sValue          As String
    
    On Error GoTo EH
    Set oStyle = pvGetStyle(oDomNode.Name, oDomNode.CssClass, TypeName(oDomNode.Control))
'    Debug.Print oDomNode.Name, oDomNode.CssClass, JsonDump(oStyle, Minimize:=True), Timer
    Set oDomNode.Style = oStyle
    With oDomNode.Layout
        For Each vKey In oStyle.Keys
            sValue = C_Str(oStyle.Item(vKey))
            Select Case LCase$(vKey)
            Case "width"
                .Width = pvToYogaValue(sValue)
            Case "height"
                .Height = pvToYogaValue(sValue)
            Case "min-width"
                .MinWidth = pvToYogaValue(sValue)
            Case "min-height"
                .MinHeight = pvToYogaValue(sValue)
            Case "max-width"
                .MaxWidth = pvToYogaValue(sValue)
            Case "max-height"
                .MaxHeight = pvToYogaValue(sValue)
            Case "direction"
                .StyleDirection = frToYogaEnum(LCase$(vKey), sValue)
            Case "position"
                .PositionType = frToYogaEnum(LCase$(vKey), sValue)
            Case "display"
                .Display = frToYogaEnum(LCase$(vKey), sValue)
            Case "overflow"
                .Overflow = frToYogaEnum(LCase$(vKey), sValue)
            Case "flex-direction"
                .FlexDirection = frToYogaEnum(LCase$(vKey), sValue)
            Case "flex-wrap"
                .Wrap = frToYogaEnum(LCase$(vKey), sValue)
            Case "flex-flow"
                vSplit = Split(sValue)
                If LenB(At(vSplit, 0)) <> 0 Then
                    .FlexDirection = frToYogaEnum(LCase$(vKey), At(vSplit, 0))
                End If
                If LenB(At(vSplit, 1)) <> 0 Then
                    .Wrap = frToYogaEnum(LCase$(vKey), At(vSplit, 1))
                End If
            Case "justify-content"
                .JustifyContent = frToYogaEnum(LCase$(vKey), sValue)
            Case "align-items"
                .AlignItems = frToYogaEnum(LCase$(vKey), sValue)
            Case "align-content"
                .AlignContent = frToYogaEnum(LCase$(vKey), sValue)
            Case "align-self"
                .AlignSelf = frToYogaEnum(LCase$(vKey), sValue)
            Case "flex"
                .Flex = Val(sValue)
            Case "flex-grow":
                .FlexGrow = Val(sValue)
            Case "flex-shrink":
                .FlexShrink = Val(sValue)
            Case "flex-basis":
                .FlexBasis = pvToYogaValue(sValue)
            Case "aspect-ratio"
                .AspectRatio = Val(sValue)
            '--- spacing
            Case "left"
                .Left = pvToYogaValue(sValue)
            Case "top"
                .Top = pvToYogaValue(sValue)
            Case "right"
                .Right = pvToYogaValue(sValue)
            Case "bottom"
                .Bottom = pvToYogaValue(sValue)
            Case "start"
                .Start = pvToYogaValue(sValue)
            Case "end"
                .End_ = pvToYogaValue(sValue)
            Case "margin"
                .Margin = pvToYogaValue(sValue)
            Case "margin-left"
                .MarginLeft = pvToYogaValue(sValue)
            Case "margin-top"
                .MarginTop = pvToYogaValue(sValue)
            Case "margin-right"
                .MarginRight = pvToYogaValue(sValue)
            Case "margin-bottom"
                .MarginBottom = pvToYogaValue(sValue)
            Case "margin-horizontal"
                .MarginHorizontal = pvToYogaValue(sValue)
            Case "margin-vertical"
                .MarginVertical = pvToYogaValue(sValue)
            Case "margin-start"
                .MarginStart = pvToYogaValue(sValue)
            Case "margin-end"
                .MarginEnd = pvToYogaValue(sValue)
            Case "padding"
                .Padding = pvToYogaValue(sValue)
            Case "padding-left"
                .PaddingLeft = pvToYogaValue(sValue)
            Case "padding-top"
                .PaddingTop = pvToYogaValue(sValue)
            Case "padding-right"
                .PaddingRight = pvToYogaValue(sValue)
            Case "padding-bottom"
                .PaddingBottom = pvToYogaValue(sValue)
            Case "padding-horizontal"
                .PaddingHorizontal = pvToYogaValue(sValue)
            Case "padding-vertical"
                .PaddingVertical = pvToYogaValue(sValue)
            Case "padding-start"
                .PaddingStart = pvToYogaValue(sValue)
            Case "padding-end"
                .PaddingEnd = pvToYogaValue(sValue)
            Case "border-left", "border-left-width"
                .BorderLeftWidth = Val(sValue)
            Case "border-top", "border-top-width"
                .BorderTopWidth = Val(sValue)
            Case "border-right", "border-right-width"
                .BorderRightWidth = Val(sValue)
            Case "border-bottom", "border-bottom-width"
                .BorderBottomWidth = Val(sValue)
            Case "border-start", "border-start-width"
                .BorderStartWidth = Val(sValue)
            Case "border-end", "border-end-width"
                .BorderEndWidth = Val(sValue)
            Case "border", "border-width"
                .BorderWidth = Val(sValue)
            Case Else
                '--- allow setting control properties from CSS
                If Left$(LCase$(vKey), Len(STR_PROP_PREFIX)) = STR_PROP_PREFIX Then
                    If InStr(vKey, ":") = 0 And Not oDomNode.Control Is Nothing Then
                        CallByName oDomNode.Control, Replace(Mid$(vKey, Len(STR_PROP_PREFIX) + 1), "-", vbNullString), VbLet, sValue
                    End If
                Else
                    RaiseEvent StyleCustomProperty(oDomNode, vKey & vbNullString, oStyle.Item(vKey))
                End If
            End Select
        Next
    End With
    If oDomNode.Count > 0 Then
        For Each oItem In oDomNode
            pvApplyStyles oItem
        Next
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
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

Friend Function frToYogaEnum(sProp As String, sValue As String) As Long
    Select Case sProp
    Case "direction"
        Select Case LCase$(sValue)
        Case "ltr"
            frToYogaEnum = yogaDirLTR
        Case "rtl"
            frToYogaEnum = yogaDirRTL
        Case Else
            #If DebugMode Then
                Debug.Print "Unknown value for '" & sProp & "': " & sValue, Timer
            #End If
        End Select
    Case "position"
        Select Case LCase$(sValue)
        Case "absolute"
            frToYogaEnum = yogaPosAbsolute
        Case "relative"
            frToYogaEnum = yogaPosRelative
        Case Else
            #If DebugMode Then
                Debug.Print "Unknown value for '" & sProp & "': " & sValue, Timer
            #End If
        End Select
    Case "display"
        Select Case LCase$(sValue)
        Case "none"
            frToYogaEnum = yogaDisplayNone
        Case "flex"
            frToYogaEnum = yogaDisplayFlex
        Case Else
            #If DebugMode Then
                Debug.Print "Unknown value for '" & sProp & "': " & sValue, Timer
            #End If
        End Select
    Case "overflow"
        Select Case LCase$(sValue)
        Case "hidden"
            frToYogaEnum = yogaOverflowHidden
        Case "visible"
            frToYogaEnum = yogaOverflowVisible
        Case "scroll"
            frToYogaEnum = yogaOverflowScroll
        Case Else
            #If DebugMode Then
                Debug.Print "Unknown value for '" & sProp & "': " & sValue, Timer
            #End If
        End Select
    Case "flex-direction"
        Select Case LCase$(sValue)
        Case "row"
            frToYogaEnum = yogaFlexRow
        Case "row-reverse"
            frToYogaEnum = yogaFlexRowReverse
        Case "column"
            frToYogaEnum = yogaFlexColumn
        Case "column-reverse"
            frToYogaEnum = yogaFlexColumnReverse
        Case Else
            #If DebugMode Then
                Debug.Print "Unknown value for '" & sProp & "': " & sValue, Timer
            #End If
        End Select
    Case "flex-wrap"
        Select Case LCase$(sValue)
        Case "nowrap"
            frToYogaEnum = yogaWrapNoWrap
        Case "wrap"
            frToYogaEnum = yogaWrapWrap
        Case "wrap-reverse"
            frToYogaEnum = yogaWrapWrapReverse
        Case Else
            #If DebugMode Then
                Debug.Print "Unknown value for '" & sProp & "': " & sValue, Timer
            #End If
        End Select
    Case "justify-content"
        Select Case LCase$(sValue)
        Case "flex-start"
            frToYogaEnum = yogaJustFlexStart
        Case "flex-end"
            frToYogaEnum = yogaJustFlexEnd
        Case "center"
            frToYogaEnum = yogaJustCenter
        Case "space-between"
            frToYogaEnum = yogaJustSpaceBetween
        Case "space-around"
            frToYogaEnum = yogaJustSpaceAround
        Case "space-evenly"
            frToYogaEnum = yogaJustSpaceEvenly
            #If DebugMode Then
                Debug.Print "Unknown value for '" & sProp & "': " & sValue, Timer
            #End If
        End Select
    Case "align-items"
        Select Case LCase$(sValue)
        Case "flex-start"
            frToYogaEnum = yogaAlignFlexStart
        Case "flex-end"
            frToYogaEnum = yogaAlignFlexEnd
        Case "center"
            frToYogaEnum = yogaAlignCenter
        Case "stretch"
            frToYogaEnum = yogaAlignStretch
        Case "baseline"
            frToYogaEnum = yogaAlignBaseline
            #If DebugMode Then
                Debug.Print "Unknown value for '" & sProp & "': " & sValue, Timer
            #End If
        End Select
    Case "align-content"
        Select Case LCase$(sValue)
        Case "flex-start"
            frToYogaEnum = yogaAlignFlexStart
        Case "flex-end"
            frToYogaEnum = yogaAlignFlexEnd
        Case "center"
            frToYogaEnum = yogaAlignCenter
        Case "stretch"
            frToYogaEnum = yogaAlignStretch
        Case "space-between"
            frToYogaEnum = yogaAlignSpaceBetween
        Case "space-around"
            frToYogaEnum = yogaAlignSpaceAround
        Case Else
            #If DebugMode Then
                Debug.Print "Unknown value for '" & sProp & "': " & sValue, Timer
            #End If
        End Select
    Case "align-self"
        Select Case LCase$(sValue)
        Case "auto"
            frToYogaEnum = yogaAlignAuto
        Case "flex-start"
            frToYogaEnum = yogaAlignFlexStart
        Case "flex-end"
            frToYogaEnum = yogaAlignFlexEnd
        Case "center"
            frToYogaEnum = yogaAlignCenter
        Case "baseline"
            frToYogaEnum = yogaAlignBaseline
        Case "stretch"
            frToYogaEnum = yogaAlignStretch
        Case Else
            #If DebugMode Then
                Debug.Print "Unknown value for '" & sProp & "': " & sValue, Timer
            #End If
        End Select
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
        For Each vElem In Split(StrReverse(CssClass))
            vElem = StrReverse(vElem)
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
    If m_oCache.Exists(sKey) Then
        Set pvTryGetCache = m_oCache.Item(sKey)
    End If
End Function

Private Function pvSetCache(sKey As String, oValue As Object)
    If oValue Is Nothing Then
        m_oCache.Remove sKey
    Else
        Set m_oCache.Item(sKey) = oValue
    End If
End Function

Private Function pvEmptyStyle() As Object
    Static oEmpty       As Object
    
    If oEmpty Is Nothing Then
        Set oEmpty = CreateObject("Scripting.Dictionary")
    End If
    Debug.Assert oEmpty.Count = 0
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

Private Function pvParentRegisterCancelMode(oCtl As Object) As Boolean
    Dim bHandled        As Boolean
    
    RaiseEvent RegisterCancelMode(oCtl, bHandled)
    If Not bHandled Then
        On Error GoTo QH
        Parent.RegisterCancelMode oCtl
        On Error GoTo 0
    End If
    '--- success
    pvParentRegisterCancelMode = True
QH:
End Function

Private Sub pvHandleMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "pvHandleMouseDown"
    
    On Error GoTo EH
    m_bDragging = False
    m_nDownButton = Button
    m_nDownShift = Shift
    m_sngDownX = X
    m_sngDownY = Y
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub pvHandleMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "pvHandleMouseMove"
#If Not ImplUseShared Then
    Const ucsPicHandCursor As Long = 1
#End If
    
    On Error GoTo EH
    #If Shift Then '--- touch args
    #End If
    m_nDownButton = m_nDownButton And Button
    If m_nDownButton <> 0 Then
        If Not m_bDragging Then
            If Abs(X - m_sngDownX) > LNG_DRAG_DISTANCE * ScreenTwipsPerPixelX Or Abs(Y - m_sngDownY) > LNG_DRAG_DISTANCE * ScreenTwipsPerPixelY Then
                RaiseEvent StartDrag(m_nDownButton, m_nDownShift, m_sngDownX, m_sngDownY, m_bDragging)
                If m_bDragging Then
                    MousePointer = vbCustom
                    Set MouseIcon = LoadStdPicture(ucsPicHandCursor)
                End If
            End If
        End If
    End If
    If m_bDragging Then
        If GetCapture() <> ContainerHwnd Or m_nDownButton = 0 Then
            m_bDragging = False
            MousePointer = vbDefault
        End If
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub pvHandleMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    #If Button And Shift And X And Y Then '--- touch args
    #End If
    m_nDownButton = 0
    m_bDragging = False
    MousePointer = vbDefault
End Sub

#If Not ImplUseShared Then
Private Function C_Str(ByVal Value As Variant) As String
    On Error GoTo QH
    C_Str = CStr(Value)
QH:
End Function

Private Function At(Data As Variant, ByVal Index As Long, Optional Default As String) As String
    On Error GoTo QH
    At = Default
    If IsArray(Data) Then
        If LBound(Data) <= Index And Index <= UBound(Data) Then
            At = Data(Index)
        End If
    End If
QH:
End Function

Private Function LoadStdPicture(ByVal eType As Long) As StdPicture
    #If eType Then '--- touch args
    #End If
End Function

Private Property Get ScreenTwipsPerPixelX() As Single
    ScreenTwipsPerPixelX = Screen.TwipsPerPixelX
End Property

Private Property Get ScreenTwipsPerPixelY() As Single
    ScreenTwipsPerPixelY = Screen.TwipsPerPixelY
End Property
#End If

'=========================================================================
' Events
'=========================================================================

Private Sub btnButton_AccessKeyPress(Index As Integer, KeyAscii As Integer)
    Const FUNC_NAME     As String = "btnButton_AccessKeyPress"
    
    On Error GoTo EH
    RaiseEvent AccessKeyPress(m_cMapping.Item("#" & ObjPtr(btnButton.Item(Index))), KeyAscii)
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub btnButton_Click(Index As Integer)
    Const FUNC_NAME     As String = "btnButton_Click"
    Dim bDragging       As Boolean
    
    On Error GoTo EH
    bDragging = m_bDragging
    pvHandleMouseUp m_nDownButton, m_nDownShift, m_sngDownX, m_sngDownY
    If Not bDragging Then
        RaiseEvent Click(m_cMapping.Item("#" & ObjPtr(btnButton.Item(Index))))
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub btnButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "btnButton_MouseDown"
    Dim sngX            As Single
    Dim sngY            As Single
    
    On Error GoTo EH
    sngX = btnButton.Item(Index).Left + X
    sngY = btnButton.Item(Index).Top + Y
    pvHandleMouseDown Button, Shift, sngX, sngY
    RaiseEvent MouseDown(m_cMapping.Item("#" & ObjPtr(btnButton.Item(Index))), Button, Shift, sngX, sngY)
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub btnButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "btnButton_MouseMove"
    Dim sngX            As Single
    Dim sngY            As Single
    
    On Error GoTo EH
    sngX = btnButton.Item(Index).Left + X
    sngY = btnButton.Item(Index).Top + Y
    pvHandleMouseMove Button, Shift, sngX, sngY
    RaiseEvent MouseMove(m_cMapping.Item("#" & ObjPtr(btnButton.Item(Index))), Button, Shift, sngX, sngY)
    If m_bDragging Then
        btnButton.Item(Index).CancelMode
        Button = -1
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub btnButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "btnButton_MouseUp"
    Dim sngX            As Single
    Dim sngY            As Single
    Dim bDragging       As Boolean
    
    On Error GoTo EH
    bDragging = m_bDragging
    sngX = btnButton.Item(Index).Left + X
    sngY = btnButton.Item(Index).Top + Y
    pvHandleMouseUp Button, Shift, sngX, sngY
    RaiseEvent MouseUp(m_cMapping.Item("#" & ObjPtr(btnButton.Item(Index))), Button, Shift, sngX, sngY)
    If bDragging Then
        btnButton.Item(Index).CancelMode
        Button = -1
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub btnButton_OwnerDraw(Index As Integer, ByVal hGraphics As Long, ByVal hFont As Long, ByVal ButtonState As UcsNineButtonStateEnum, ClientLeft As Long, ClientTop As Long, ClientWidth As Long, ClientHeight As Long, Caption As String, ByVal hPicture As Long)
    Const FUNC_NAME     As String = "btnButton_OwnerDraw"
    
    On Error GoTo EH
    RaiseEvent OwnerDraw(m_cMapping.Item("#" & ObjPtr(btnButton.Item(Index))), hGraphics, hFont, ButtonState, ClientLeft, ClientTop, ClientWidth, ClientHeight, Caption, hPicture)
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub btnButton_RegisterCancelMode(Index As Integer, oCtl As Object, Handled As Boolean)
    Const FUNC_NAME     As String = "btnButton_RegisterCancelMode"
    
    On Error GoTo EH
    pvParentRegisterCancelMode Me
    If Not m_oCtlCancelMode Is Nothing And Not m_oCtlCancelMode Is oCtl Then
        m_oCtlCancelMode.CancelMode
    End If
    Set m_oCtlCancelMode = oCtl
    Handled = True
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub labLabel_Click(Index As Integer)
    Const FUNC_NAME     As String = "labLabel_Click"
    Dim bDragging       As Boolean
    
    On Error GoTo EH
    bDragging = m_bDragging
    pvHandleMouseUp m_nDownButton, m_nDownShift, m_sngDownX, m_sngDownY
    If Not bDragging Then
        RaiseEvent Click(m_cMapping.Item("#" & ObjPtr(labLabel.Item(Index))))
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub labLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "labLabel_MouseDown"
    Dim sngX            As Single
    Dim sngY            As Single
    
    On Error GoTo EH
    sngX = labLabel.Item(Index).Left + X
    sngY = labLabel.Item(Index).Top + Y
    pvHandleMouseDown Button, Shift, sngX, sngY
    RaiseEvent MouseDown(m_cMapping.Item("#" & ObjPtr(labLabel.Item(Index))), Button, Shift, sngX, sngY)
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub labLabel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "labLabel_MouseMove"
    Dim sngX            As Single
    Dim sngY            As Single
    
    On Error GoTo EH
    sngX = labLabel.Item(Index).Left + X
    sngY = labLabel.Item(Index).Top + Y
    pvHandleMouseMove Button, Shift, sngX, sngY
    RaiseEvent MouseMove(m_cMapping.Item("#" & ObjPtr(labLabel.Item(Index))), Button, Shift, sngX, sngY)
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub labLabel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "labLabel_MouseUp"
    Dim sngX            As Single
    Dim sngY            As Single
    
    On Error GoTo EH
    sngX = labLabel.Item(Index).Left + X
    sngY = labLabel.Item(Index).Top + Y
'    pvHandleMouseUp Button, Shift, sngX, sngY
    RaiseEvent MouseUp(m_cMapping.Item("#" & ObjPtr(labLabel.Item(Index))), Button, Shift, sngX, sngY)
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub m_oFont_FontChanged(ByVal PropertyName As String)
    Const FUNC_NAME     As String = "m_oFont_FontChanged"
    Dim lIdx            As Long
    
    On Error GoTo EH
    For lIdx = 0 To m_lButtonCount
        btnButton.Item(lIdx).Font = m_oFont
    Next
    For lIdx = 0 To m_lLabelCount
        labLabel.Item(lIdx).Font = m_oFont
    Next
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    HitResult = vbHitResultHit
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "UserControl_MouseDown"
    
    On Error GoTo EH
    CancelMode
    pvHandleMouseDown Button, Shift, X, Y
    RaiseEvent MouseDown(Nothing, Button, Shift, ScaleX(X, ScaleMode, vbContainerPosition), ScaleY(Y, ScaleMode, vbContainerPosition))
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "UserControl_MouseMove"
    
    On Error GoTo EH
    CancelMode
    pvHandleMouseMove Button, Shift, X, Y
    RaiseEvent MouseMove(Nothing, Button, Shift, ScaleX(X, ScaleMode, vbContainerPosition), ScaleY(Y, ScaleMode, vbContainerPosition))
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "UserControl_MouseUp"
    
    On Error GoTo EH
    pvHandleMouseUp Button, Shift, X, Y
    RaiseEvent MouseUp(Nothing, Button, Shift, ScaleX(X, ScaleMode, vbContainerPosition), ScaleY(Y, ScaleMode, vbContainerPosition))
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_Resize()
    Const FUNC_NAME     As String = "UserControl_Resize"
    
    On Error GoTo EH
    If Not m_oRoot Is Nothing And m_bAutoApplyLayout Then
        ApplyLayout Width, Height
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_InitProperties()
    Const FUNC_NAME     As String = "UserControl_InitProperties"
    
    On Error GoTo EH
    If Ambient.UserMode Then
        pvInitRoot
    End If
    Set Font = Ambient.Font
    AutoApplyLayout = DEF_AUTOAPPLYLAYOUT
    Enabled = DEF_ENABLED
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Const FUNC_NAME     As String = "UserControl_ReadProperties"
    
    On Error GoTo EH
    If Ambient.UserMode Then
        pvInitRoot
    End If
    With PropBag
        Set Font = .ReadProperty("Font", Ambient.Font)
        AutoApplyLayout = .ReadProperty("AutoApplyLayout", DEF_AUTOAPPLYLAYOUT)
        Enabled = .ReadProperty("Enabled", DEF_ENABLED)
    End With
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Const FUNC_NAME     As String = "UserControl_WriteProperties"
    
    On Error GoTo EH
    With PropBag
        .WriteProperty "Font", Font, Ambient.Font
        .WriteProperty "AutoApplyLayout", AutoApplyLayout, DEF_AUTOAPPLYLAYOUT
        .WriteProperty "Enabled", Enabled, DEF_ENABLED
    End With
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
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
    If Not m_oRoot Is Nothing Then
        Set m_oRoot.frFlexBox = Nothing
        Set m_oRoot = Nothing
    End If
    Set m_cMapping = Nothing
    Set m_oYogaConfig = Nothing
    #If DebugMode Then
        DebugInstanceTerm MODULE_NAME, m_sDebugID
    #End If
End Sub

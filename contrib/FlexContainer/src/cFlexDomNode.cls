VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFlexDomNode"
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
Private Const MODULE_NAME As String = "cFlexDomNode"

#Const ImplUseShared = VBYOGA_USE_SHARED <> 0

'=========================================================================
' API
'=========================================================================

Private Const DT_WORDBREAK                  As Long = &H10
Private Const DT_CALCRECT                   As Long = &H400

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZEAPI) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Type SIZEAPI
    cx                  As Long
    cy                  As Long
End Type

Private Type RECT
    Left            As Long
    Top             As Long
    Right           As Long
    Bottom          As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_lFlexBoxWeakRef       As Long
Private m_lParentWeakRef        As Long
Private m_cChildren             As Collection
Private m_lControlWeakRef       As Long
Public Name                     As String
Public CssClass                 As String
Public Layout                   As cYogaNode
Public Style                    As Object
Public Tag                      As Variant
'--- debug
#If DebugMode Then
    Private m_sDebugID          As String
#End If

'=========================================================================
' Properties
'=========================================================================

Friend Property Get frFlexBox() As ctxFlexContainer
    If m_lFlexBoxWeakRef <> 0 Then
        Set frFlexBox = YogaWeakRefResurrectTarget(m_lFlexBoxWeakRef)
    ElseIf m_lParentWeakRef <> 0 Then
        Set frFlexBox = Parent.frFlexBox
    End If
End Property

Friend Property Set frFlexBox(oValue As ctxFlexContainer)
    m_lFlexBoxWeakRef = YogaWeakRefInit(oValue)
End Property

Property Get Parent() As cFlexDomNode
    Set Parent = YogaWeakRefResurrectTarget(m_lParentWeakRef)
End Property

Friend Property Set frParent(oValue As cFlexDomNode)
    m_lParentWeakRef = YogaWeakRefInit(oValue)
End Property

Property Get Root() As cFlexDomNode
    Dim oParent         As cFlexDomNode
    
    If m_lParentWeakRef <> 0 Then
        Set oParent = Parent
        Set Root = oParent.Root
        If Root Is Nothing Then
            Set Root = oParent
        End If
    End If
End Property

Property Get Item(Index As Variant) As cFlexDomNode
    Set Item = m_cChildren.Item(Index)
End Property

Property Get Count() As Long
    If Not m_cChildren Is Nothing Then
        Count = m_cChildren.Count
    End If
End Property

Public Function GetEnumerator() As IUnknown
Attribute GetEnumerator.VB_UserMemId = -4
Attribute GetEnumerator.VB_MemberFlags = "40"
    Set GetEnumerator = m_cChildren.[_NewEnum]
End Function

Property Get Control() As Object
    Set Control = YogaWeakRefResurrectTarget(m_lControlWeakRef)
End Property

Friend Property Set frControl(oValue As Object)
    m_lControlWeakRef = YogaWeakRefInit(oValue)
End Property

Property Get CtlLabel() As VB.Label
    Dim oRetVal         As Object
    
    Set oRetVal = Control
    If Not oRetVal Is Nothing Then
        If TypeOf oRetVal Is VB.Label Then
            Set CtlLabel = oRetVal
        End If
    End If
End Property

Property Get CtlButton() As ctxNineButton
    Dim oRetVal         As Object
    
    Set oRetVal = Control
    If Not oRetVal Is Nothing Then
        If TypeOf oRetVal Is ctxNineButton Then
            Set CtlButton = oRetVal
        End If
    End If
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function AddContainer(Optional Name As String, Optional CssClass As String) As cFlexDomNode
    Set AddContainer = AddControl(Nothing, Name, Trim$("container " & CssClass))
End Function

Public Function AddButton(Optional Name As String, Optional CssClass As String) As cFlexDomNode
    Set AddButton = AddControl(frFlexBox.frLoadButton, Name, Trim$("button " & CssClass))
End Function

Public Function AddLabel(Optional Name As String, Optional CssClass As String) As cFlexDomNode
    Set AddLabel = AddControl(frFlexBox.frLoadLabel, Name, Trim$("label " & CssClass))
    AddLabel.Layout.SetMeasureFunction AddLabel
End Function

Public Function AddControl(oCtl As Object, Optional Name As String, Optional CssClass As String) As cFlexDomNode
    Dim oDomNode        As cFlexDomNode
    Dim oFlexBox        As ctxFlexContainer
    
    Set oFlexBox = frFlexBox
    If m_cChildren Is Nothing Then
        Set m_cChildren = New Collection
    End If
    Set oDomNode = New cFlexDomNode
    With oDomNode
        Set .frControl = oCtl
        .Name = Name
        .CssClass = CssClass
        Set .frParent = Me
        Set .Layout = YogaNodeNew(oFlexBox.frYogaConfig)
        Layout.AddChild .Layout
    End With
    If Not oCtl Is Nothing Then
        oFlexBox.frAddDomNodeMapping oDomNode, oCtl
    End If
    If LenB(Name) <> 0 Then
        m_cChildren.Add oDomNode, Name
    Else
        m_cChildren.Add oDomNode
    End If
    Set AddControl = oDomNode
End Function

Public Sub ApplyLayout( _
            Optional ByVal sngClipLeft As Single, _
            Optional ByVal sngClipTop As Single, _
            Optional ByVal sngClipWidth As Single, _
            Optional ByVal sngClipHeight As Single)
    Dim oItem           As cFlexDomNode
    Dim oCtl            As Object
    
    Set oCtl = Control
    If Not oCtl Is Nothing Then
        With Layout
            oCtl.Visible = (.Display <> yogaDisplayNone)
            Debug.Assert .GetAbsoluteX = AlignTwipsToPix(.GetAbsoluteX)
            Debug.Assert .GetAbsoluteY = AlignTwipsToPix(.GetAbsoluteY)
'            Debug.Assert .LayoutWidth = AlignTwipsToPix(.LayoutWidth)
'            Debug.Assert .LayoutHeight = AlignTwipsToPix(.LayoutHeight)
            If sngClipWidth > 0 Or sngClipHeight > 0 Then
                If .GetAbsoluteX <= sngClipLeft + sngClipWidth And .GetAbsoluteX + .LayoutWidth >= sngClipLeft Then
                    If .GetAbsoluteY <= sngClipTop + sngClipHeight And .GetAbsoluteY + .LayoutHeight >= sngClipTop Then
                        MoveCtl oCtl, .GetAbsoluteX, .GetAbsoluteY, .LayoutWidth, .LayoutHeight
                    Else
                        oCtl.Left = -oCtl.Width
                    End If
                Else
                    oCtl.Left = -oCtl.Width
                End If
            Else
                MoveCtl oCtl, .GetAbsoluteX, .GetAbsoluteY, .LayoutWidth, .LayoutHeight
            End If
        End With
    End If
    If Not m_cChildren Is Nothing Then
        For Each oItem In m_cChildren
            oItem.ApplyLayout sngClipLeft, sngClipTop, sngClipWidth, sngClipHeight
        Next
    End If
End Sub

Public Function SelectItemByName(sName As String) As cFlexDomNode
    Dim vItem           As Variant
    Dim oItem           As cFlexDomNode
    
    If SearchCollection(m_cChildren, sName, RetVal:=vItem) Then
        Set SelectItemByName = vItem
    ElseIf Not m_cChildren Is Nothing Then
        For Each oItem In m_cChildren
            Set SelectItemByName = oItem.SelectItemByName(sName)
            If Not SelectItemByName Is Nothing Then
                Exit For
            End If
        Next
    End If
End Function

Public Sub MeasureCallback( _
            oNode As Object, _
            ByVal sngWidth As Single, _
            ByVal eWidthMode As YogaMeasureMode, _
            ByVal sngHeight As Single, _
            ByVal eHeightMode As YogaMeasureMode, _
            sngLayoutWidth As Single, _
            sngLayoutHeight As Single)
Attribute MeasureCallback.VB_MemberFlags = "40"
    Const LNG_PRIVATE   As Long = &HE1B6 '-- U+E000 to U+F8FF - Private Use Area (PUA)
    Dim hDC             As Long
    Dim pFont           As IFont
    Dim hPrevFont       As Long
    Dim sText           As String
    Dim sz              As SIZEAPI
    Dim rc              As RECT
    
    #If oNode And eWidthMode And sngHeight And eHeightMode Then '--- touch args (silence MZ-Tools->Review Source Code)
    #End If
    With CtlLabel
        If .AutoSize Then
            If YogaConstantsIsUndefined(sngWidth) Then
                rc.Right = Screen.Width \ Screen.TwipsPerPixelX
            Else
                rc.Right = sngWidth \ Screen.TwipsPerPixelX
            End If
            If YogaConstantsIsUndefined(sngHeight) Then
                rc.Bottom = Screen.Height \ Screen.TwipsPerPixelY
            Else
                rc.Bottom = sngHeight \ Screen.TwipsPerPixelY
            End If
            hDC = GetDC(0)
            Set pFont = .Font
            hPrevFont = SelectObject(hDC, pFont.hFont)
            sText = .Caption
            If .UseMnemonic Then
                sText = Replace(Replace(Replace(sText, _
                    "&&", ChrW$(LNG_PRIVATE)), _
                    "&", vbNullString), _
                    ChrW$(LNG_PRIVATE), "&")
            End If
            If InStr(sText, vbCrLf) > 0 Then
                GoTo MultiLine
            End If
            Call GetTextExtentPoint32(hDC, sText, Len(sText), sz)
            If sz.cx > rc.Right Then
                If .WordWrap Then
MultiLine:
                    Call DrawText(hDC, sText, Len(sText), rc, DT_CALCRECT Or DT_WORDBREAK)
                    sz.cx = rc.Right
                    sz.cy = rc.Bottom
                End If
            End If
            Call SelectObject(hDC, hPrevFont)
            Call ReleaseDC(0, hDC)
'            Debug.Print sText, .Font.Name, .Font.Size, sz.cx, sz.cy, Timer
            sngLayoutWidth = sz.cx * Screen.TwipsPerPixelX
            sngLayoutHeight = sz.cy * Screen.TwipsPerPixelY
        Else
            sngLayoutWidth = .Width
            sngLayoutHeight = .Height
        End If
    End With
End Sub

#If Not ImplUseShared Then
Private Function AlignTwipsToPix(ByVal sngTwips As Single) As Single
    AlignTwipsToPix = Int(sngTwips / Screen.TwipsPerPixelX + 0.5) * Screen.TwipsPerPixelX
End Function

Private Sub MoveCtl(oCtl As Object, ByVal Left As Single, ByVal Top As Variant, ByVal Width As Variant, ByVal Height As Variant)
    If 1440 \ Screen.TwipsPerPixelX = 1440 / Screen.TwipsPerPixelX Then
        oCtl.Move Left, Top, Width, Height
    Else
        oCtl.Move Left + Screen.TwipsPerPixelX, Top, Width, Height
        oCtl.Move Left
    End If
End Sub

Private Function SearchCollection(ByVal pCol As Object, Index As Variant, Optional RetVal As Variant) As Boolean
    On Error GoTo QH
    AssignVariant RetVal, pCol.Item(Index)
    SearchCollection = True
QH:
End Function

Private Sub AssignVariant(vDest As Variant, vSrc As Variant)
    If IsObject(vSrc) Then
        Set vDest = vSrc
    Else
        vDest = vSrc
    End If
End Sub
#End If

'=========================================================================
' Base class events
'=========================================================================

#If DebugMode Then
    Private Sub Class_Initialize()
        DebugInstanceInit MODULE_NAME, m_sDebugID, Me
    End Sub
#End If

Private Sub Class_Terminate()
    Dim oChild          As cFlexDomNode
    
    If Not m_cChildren Is Nothing Then
        For Each oChild In m_cChildren
            Set oChild.frParent = Nothing
        Next
        Set m_cChildren = Nothing
    End If
    Set Layout = Nothing
    Set Style = Nothing
    #If DebugMode Then
        DebugInstanceTerm MODULE_NAME, m_sDebugID
    #End If
End Sub

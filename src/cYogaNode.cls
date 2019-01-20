VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cYogaNode"
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

'=========================================================================
' Public Enums
'=========================================================================

Public Enum YogaAlign
    yogaAlignAuto
    yogaAlignFlexStart
    yogaAlignCenter
    yogaAlignFlexEnd
    yogaAlignStretch
    yogaAlignBaseline
    yogaAlignSpaceBetween
    yogaAlignSpaceAround
End Enum
    
Public Enum YogaDimension
    yogaDimWidth
    yogaDimHeight
End Enum

Public Enum YogaDirection
    yogaDirInherit
    yogaDirLTR
    yogaDirRTL
End Enum

Public Enum YogaDisplay
    yogaDisplayFlex
    yogaDisplayNone
End Enum

Public Enum YogaEdge
    yogaEdgeLeft
    yogaEdgeTop
    yogaEdgeRight
    yogaEdgeBottom
    yogaEdgeStart
    yogaEdgeEnd
    yogaEdgeHorizontal
    yogaEdgeVertical
    yogaEdgeAll
End Enum

Public Enum YogaExperimentalFeature
    yogaExpWebFlexBasis
End Enum

Public Enum YogaFlexDirection
    yogaFlexColumn
    yogaFlexColumnReverse
    yogaFlexRow
    yogaFlexRowReverse
End Enum

Public Enum YogaJustify
    yogaJustFlexStart
    yogaJustCenter
    yogaJustFlexEnd
    yogaJustSpaceBetween
    yogaJustSpaceAround
    yogaJustSpaceEvenly
End Enum

Public Enum YogaLogLevel
    yogaLogError
    yogaLogWarn
    yogaLogInfo
    yogaLogDebug
    yogaLogVerbose
    yogaLogFatal
End Enum

Public Enum YogaMeasureMode
    yogaMeasureUndefined
    yogaMeasureExactly
    yogaMeasureAtMost
End Enum

Public Enum YogaOverflow
    yogaOverflowVisible
    yogaOverflowHidden
    yogaOverflowScroll
End Enum

Public Enum YogaPositionType
    yogaPosRelative
    yogaPosAbsolute
End Enum

Public Enum YogaPrintOptions
    yogaProLayout = 1
    yogaProStyle = 2
    yogaProChildren = 4
End Enum

Public Enum YogaUnit
    yogaUnitUndefined
    yogaUnitPoint
    yogaUnitPercent
    yogaUnitAuto
End Enum

Public Enum YogaWrap
    yogaWrapNoWrap
    yogaWrapWrap
    yogaWrapWrapReverse
End Enum

'=========================================================================
' API
'=========================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function YGNodeNew Lib "yoga" Alias "_YGNodeNew@0" () As Long
Private Declare Function YGNodeNewWithConfig Lib "yoga" Alias "_YGNodeNewWithConfig@4" (ByVal lConfigPtr As Long) As Long
Private Declare Function YGNodeFree Lib "yoga" Alias "_YGNodeFree@4" (ByVal lNodePtr As Long) As Long
Private Declare Function YGNodeSetContext Lib "yoga" Alias "_YGNodeSetContext@8" (ByVal lNodePtr As Long, ByVal lContext As Long) As Long
Private Declare Function YGNodeReset Lib "yoga" Alias "_YGNodeReset@4" (ByVal lNodePtr As Long) As Long
Private Declare Function YGNodeCopyStyle Lib "yoga" Alias "_YGNodeCopyStyle@8" (ByVal lNodePtr As Long, ByVal lSrcPtr As Long) As Long
Private Declare Function YGNodeIsDirty Lib "yoga" Alias "_YGNodeIsDirty@4" (ByVal lNodePtr As Long) As Long
Private Declare Function YGNodeMarkDirty Lib "yoga" Alias "_YGNodeMarkDirty@4" (ByVal lNodePtr As Long) As Long
Private Declare Function YGNodeGetHasNewLayout Lib "yoga" Alias "_YGNodeGetHasNewLayout@4" (ByVal lNodePtr As Long) As Long
Private Declare Function YGNodeSetHasNewLayout Lib "yoga" Alias "_YGNodeSetHasNewLayout@8" (ByVal lNodePtr As Long, ByVal lValue As Long) As Long
Private Declare Function YGNodePrint Lib "yoga" Alias "_YGNodePrint@8" (ByVal lNodePtr As Long, ByVal lOptions As Long) As Long
Private Declare Function YGNodeInsertChild Lib "yoga" Alias "_YGNodeInsertChild@12" (ByVal lNodePtr As Long, ByVal lChildPtr As Long, ByVal lIndex As Long) As Long
Private Declare Function YGNodeRemoveChild Lib "yoga" Alias "_YGNodeRemoveChild@8" (ByVal lNodePtr As Long, ByVal lChildPtr As Long) As Long
Private Declare Function YGNodeSetMeasureFunc Lib "yoga" Alias "_YGNodeSetMeasureFunc@8" (ByVal lNodePtr As Long, ByVal pfn As Long) As Long
Private Declare Function YGNodeSetBaselineFunc Lib "yoga" Alias "_YGNodeSetBaselineFunc@8" (ByVal lNodePtr As Long, ByVal pfn As Long) As Long
Private Declare Function YGNodeCalculateLayout Lib "yoga" Alias "_YGNodeCalculateLayout@16" (ByVal lNodePtr As Long, ByVal sngWidth As Single, ByVal sngHeight As Single, ByVal lParentDirection As Long) As Long
Private Declare Function YGNodeStyleGetDirection Lib "yoga" Alias "_YGNodeStyleGetDirection@4" (ByVal lNodePtr As Long) As Long
Private Declare Function YGNodeStyleSetDirection Lib "yoga" Alias "_YGNodeStyleSetDirection@8" (ByVal lNodePtr As Long, ByVal lDirection As Long) As Long
Private Declare Function YGNodeStyleGetFlexDirection Lib "yoga" Alias "_YGNodeStyleGetFlexDirection@4" (ByVal lNodePtr As Long) As Long
Private Declare Function YGNodeStyleSetFlexDirection Lib "yoga" Alias "_YGNodeStyleSetFlexDirection@8" (ByVal lNodePtr As Long, ByVal lFlexDirection As Long) As Long
Private Declare Function YGNodeStyleGetJustifyContent Lib "yoga" Alias "_YGNodeStyleGetJustifyContent@4" (ByVal lNodePtr As Long) As Long
Private Declare Function YGNodeStyleSetJustifyContent Lib "yoga" Alias "_YGNodeStyleSetJustifyContent@8" (ByVal lNodePtr As Long, ByVal lJustifyContent As Long) As Long
Private Declare Function YGNodeStyleGetDisplay Lib "yoga" Alias "_YGNodeStyleGetDisplay@4" (ByVal lNodePtr As Long) As Long
Private Declare Function YGNodeStyleSetDisplay Lib "yoga" Alias "_YGNodeStyleSetDisplay@8" (ByVal lNodePtr As Long, ByVal lDisplay As Long) As Long
Private Declare Function YGNodeStyleGetAlignItems Lib "yoga" Alias "_YGNodeStyleGetAlignItems@4" (ByVal lNodePtr As Long) As Long
Private Declare Function YGNodeStyleSetAlignItems Lib "yoga" Alias "_YGNodeStyleSetAlignItems@8" (ByVal lNodePtr As Long, ByVal lAlignItems As Long) As Long
Private Declare Function YGNodeStyleGetAlignSelf Lib "yoga" Alias "_YGNodeStyleGetAlignSelf@4" (ByVal lNodePtr As Long) As Long
Private Declare Function YGNodeStyleSetAlignSelf Lib "yoga" Alias "_YGNodeStyleSetAlignSelf@8" (ByVal lNodePtr As Long, ByVal lAlignItems As Long) As Long
Private Declare Function YGNodeStyleGetAlignContent Lib "yoga" Alias "_YGNodeStyleGetAlignContent@4" (ByVal lNodePtr As Long) As Long
Private Declare Function YGNodeStyleSetAlignContent Lib "yoga" Alias "_YGNodeStyleSetAlignContent@8" (ByVal lNodePtr As Long, ByVal lAlignContent As Long) As Long
Private Declare Function YGNodeStyleGetPositionType Lib "yoga" Alias "_YGNodeStyleGetPositionType@4" (ByVal lNodePtr As Long) As Long
Private Declare Function YGNodeStyleSetPositionType Lib "yoga" Alias "_YGNodeStyleSetPositionType@8" (ByVal lNodePtr As Long, ByVal lPositionType As Long) As Long
Private Declare Function YGNodeStyleGetFlexWrap Lib "yoga" Alias "_YGNodeStyleGetFlexWrap@4" (ByVal lNodePtr As Long) As Long
Private Declare Function YGNodeStyleSetFlexWrap Lib "yoga" Alias "_YGNodeStyleSetFlexWrap@8" (ByVal lNodePtr As Long, ByVal lFlexWrap As Long) As Long
Private Declare Function YGNodeStyleSetFlex Lib "yoga" Alias "_YGNodeStyleSetFlex@8" (ByVal lNodePtr As Long, ByVal sngFlex As Single) As Long
Private Declare Function YGNodeStyleGetFlexGrow Lib "yoga" Alias "_YGNodeStyleGetFlexGrow@4" (ByVal lNodePtr As Long) As Single
Private Declare Function YGNodeStyleSetFlexGrow Lib "yoga" Alias "_YGNodeStyleSetFlexGrow@8" (ByVal lNodePtr As Long, ByVal sngFlexGrow As Single) As Long
Private Declare Function YGNodeStyleGetFlexShrink Lib "yoga" Alias "_YGNodeStyleGetFlexShrink@4" (ByVal lNodePtr As Long) As Single
Private Declare Function YGNodeStyleSetFlexShrink Lib "yoga" Alias "_YGNodeStyleSetFlexShrink@8" (ByVal lNodePtr As Long, ByVal sngFlexShrink As Single) As Long
Private Declare Function YGNodeStyleGetFlexBasis Lib "yoga" Alias "_YGNodeStyleGetFlexBasis@4" (ByVal lNodePtr As Long) As YogaValue
Private Declare Function YGNodeStyleSetFlexBasis Lib "yoga" Alias "_YGNodeStyleSetFlexBasis@8" (ByVal lNodePtr As Long, ByVal sngFlexBasis As Single) As Long
Private Declare Function YGNodeStyleSetFlexBasisPercent Lib "yoga" Alias "_YGNodeStyleSetFlexBasisPercent@8" (ByVal lNodePtr As Long, ByVal sngFlexBasisPercent As Single) As Long
Private Declare Function YGNodeStyleSetFlexBasisAuto Lib "yoga" Alias "_YGNodeStyleSetFlexBasisAuto@4" (ByVal lNodePtr As Long) As Long
Private Declare Function YGNodeStyleGetWidth Lib "yoga" Alias "_YGNodeStyleGetWidth@4" (ByVal lNodePtr As Long) As YogaValue
Private Declare Function YGNodeStyleSetWidth Lib "yoga" Alias "_YGNodeStyleSetWidth@8" (ByVal lNodePtr As Long, ByVal sngWidth As Single) As Long
Private Declare Function YGNodeStyleSetWidthPercent Lib "yoga" Alias "_YGNodeStyleSetWidthPercent@8" (ByVal lNodePtr As Long, ByVal sngWidthPercent As Single) As Long
Private Declare Function YGNodeStyleSetWidthAuto Lib "yoga" Alias "_YGNodeStyleSetWidthAuto@4" (ByVal lNodePtr As Long) As Long
Private Declare Function YGNodeStyleGetHeight Lib "yoga" Alias "_YGNodeStyleGetHeight@4" (ByVal lNodePtr As Long) As YogaValue
Private Declare Function YGNodeStyleSetHeight Lib "yoga" Alias "_YGNodeStyleSetHeight@8" (ByVal lNodePtr As Long, ByVal sngHeight As Single) As Long
Private Declare Function YGNodeStyleSetHeightPercent Lib "yoga" Alias "_YGNodeStyleSetHeightPercent@8" (ByVal lNodePtr As Long, ByVal sngHeightPercent As Single) As Long
Private Declare Function YGNodeStyleSetHeightAuto Lib "yoga" Alias "_YGNodeStyleSetHeightAuto@4" (ByVal lNodePtr As Long) As Long
Private Declare Function YGNodeStyleGetMaxWidth Lib "yoga" Alias "_YGNodeStyleGetMaxWidth@4" (ByVal lNodePtr As Long) As YogaValue
Private Declare Function YGNodeStyleSetMaxWidth Lib "yoga" Alias "_YGNodeStyleSetMaxWidth@8" (ByVal lNodePtr As Long, ByVal sngWidth As Single) As Long
Private Declare Function YGNodeStyleSetMaxWidthPercent Lib "yoga" Alias "_YGNodeStyleSetMaxWidthPercent@8" (ByVal lNodePtr As Long, ByVal sngWidthPercent As Single) As Long
Private Declare Function YGNodeStyleGetMaxHeight Lib "yoga" Alias "_YGNodeStyleGetMaxHeight@4" (ByVal lNodePtr As Long) As YogaValue
Private Declare Function YGNodeStyleSetMaxHeight Lib "yoga" Alias "_YGNodeStyleSetMaxHeight@8" (ByVal lNodePtr As Long, ByVal sngHeight As Single) As Long
Private Declare Function YGNodeStyleSetMaxHeightPercent Lib "yoga" Alias "_YGNodeStyleSetMaxHeightPercent@8" (ByVal lNodePtr As Long, ByVal sngHeightPercent As Single) As Long
Private Declare Function YGNodeStyleGetMinWidth Lib "yoga" Alias "_YGNodeStyleGetMinWidth@4" (ByVal lNodePtr As Long) As YogaValue
Private Declare Function YGNodeStyleSetMinWidth Lib "yoga" Alias "_YGNodeStyleSetMinWidth@8" (ByVal lNodePtr As Long, ByVal sngWidth As Single) As Long
Private Declare Function YGNodeStyleSetMinWidthPercent Lib "yoga" Alias "_YGNodeStyleSetMinWidthPercent@8" (ByVal lNodePtr As Long, ByVal sngWidthPercent As Single) As Long
Private Declare Function YGNodeStyleGetMinHeight Lib "yoga" Alias "_YGNodeStyleGetMinHeight@4" (ByVal lNodePtr As Long) As YogaValue
Private Declare Function YGNodeStyleSetMinHeight Lib "yoga" Alias "_YGNodeStyleSetMinHeight@8" (ByVal lNodePtr As Long, ByVal sngHeight As Single) As Long
Private Declare Function YGNodeStyleSetMinHeightPercent Lib "yoga" Alias "_YGNodeStyleSetMinHeightPercent@8" (ByVal lNodePtr As Long, ByVal sngHeightPercent As Single) As Long
Private Declare Function YGNodeStyleGetAspectRatio Lib "yoga" Alias "_YGNodeStyleGetAspectRatio@4" (ByVal lNodePtr As Long) As Single
Private Declare Function YGNodeStyleSetAspectRatio Lib "yoga" Alias "_YGNodeStyleSetAspectRatio@8" (ByVal lNodePtr As Long, ByVal sngAspectRatio As Single) As Long
Private Declare Function YGNodeLayoutGetLeft Lib "yoga" Alias "_YGNodeLayoutGetLeft@4" (ByVal lNodePtr As Long) As Single
Private Declare Function YGNodeLayoutGetTop Lib "yoga" Alias "_YGNodeLayoutGetTop@4" (ByVal lNodePtr As Long) As Single
Private Declare Function YGNodeLayoutGetWidth Lib "yoga" Alias "_YGNodeLayoutGetWidth@4" (ByVal lNodePtr As Long) As Single
Private Declare Function YGNodeLayoutGetHeight Lib "yoga" Alias "_YGNodeLayoutGetHeight@4" (ByVal lNodePtr As Long) As Single
Private Declare Function YGNodeLayoutGetDirection Lib "yoga" Alias "_YGNodeLayoutGetDirection@4" (ByVal lNodePtr As Long) As Long
Private Declare Function YGNodeStyleGetOverflow Lib "yoga" Alias "_YGNodeStyleGetOverflow@4" (ByVal lNodePtr As Long) As Long
Private Declare Function YGNodeStyleSetOverflow Lib "yoga" Alias "_YGNodeStyleSetOverflow@8" (ByVal lNodePtr As Long, ByVal lOverflow As Long) As Long
Private Declare Function YGNodeIsReferenceBaseline Lib "yoga" Alias "_YGNodeIsReferenceBaseline@4" (ByVal lNodePtr As Long) As Long
Private Declare Function YGNodeSetIsReferenceBaseline Lib "yoga" Alias "_YGNodeSetIsReferenceBaseline@8" (ByVal lNodePtr As Long, ByVal fBaseline As Long) As Long
Private Declare Function YGNodeStyleGetPosition Lib "yoga" Alias "_YGNodeStyleGetPosition@8" (ByVal lNodePtr As Long, ByVal lEdge As Long) As YogaValue
Private Declare Function YGNodeStyleSetPosition Lib "yoga" Alias "_YGNodeStyleSetPosition@12" (ByVal lNodePtr As Long, ByVal lEdge As Long, ByVal sngPosition As Single) As Long
Private Declare Function YGNodeStyleSetPositionPercent Lib "yoga" Alias "_YGNodeStyleSetPositionPercent@12" (ByVal lNodePtr As Long, ByVal lEdge As Long, ByVal sngPositionPercent As Single) As Long
Private Declare Function YGNodeStyleGetMargin Lib "yoga" Alias "_YGNodeStyleGetMargin@8" (ByVal lNodePtr As Long, ByVal lEdge As Long) As YogaValue
Private Declare Function YGNodeStyleSetMargin Lib "yoga" Alias "_YGNodeStyleSetMargin@12" (ByVal lNodePtr As Long, ByVal lEdge As Long, ByVal sngMargin As Single) As Long
Private Declare Function YGNodeStyleSetMarginAuto Lib "yoga" Alias "_YGNodeStyleSetMarginAuto@8" (ByVal lNodePtr As Long, ByVal lEdge As Long) As Long
Private Declare Function YGNodeStyleSetMarginPercent Lib "yoga" Alias "_YGNodeStyleSetMarginPercent@12" (ByVal lNodePtr As Long, ByVal lEdge As Long, ByVal sngMarginPercent As Single) As Long
Private Declare Function YGNodeStyleGetPadding Lib "yoga" Alias "_YGNodeStyleGetPadding@8" (ByVal lNodePtr As Long, ByVal lEdge As Long) As YogaValue
Private Declare Function YGNodeStyleSetPadding Lib "yoga" Alias "_YGNodeStyleSetPadding@12" (ByVal lNodePtr As Long, ByVal lEdge As Long, ByVal sngPadding As Single) As Long
Private Declare Function YGNodeStyleSetPaddingPercent Lib "yoga" Alias "_YGNodeStyleSetPaddingPercent@12" (ByVal lNodePtr As Long, ByVal lEdge As Long, ByVal sngPaddingPercent As Single) As Long
Private Declare Function YGNodeStyleGetBorder Lib "yoga" Alias "_YGNodeStyleGetBorder@8" (ByVal lNodePtr As Long, ByVal lEdge As Long) As Single
Private Declare Function YGNodeStyleSetBorder Lib "yoga" Alias "_YGNodeStyleSetBorder@12" (ByVal lNodePtr As Long, ByVal lEdge As Long, ByVal sngBorder As Single) As Long
Private Declare Function YGNodeLayoutGetMargin Lib "yoga" Alias "_YGNodeLayoutGetMargin@8" (ByVal lNodePtr As Long, ByVal lEdge As Long) As Single
Private Declare Function YGNodeLayoutGetPadding Lib "yoga" Alias "_YGNodeLayoutGetPadding@8" (ByVal lNodePtr As Long, ByVal lEdge As Long) As Single
Private Declare Function YGFloatIsUndefined Lib "yoga" Alias "_YGFloatIsUndefined@4" (ByVal sngValue As Single) As Long

#If False Then
Const Value = 1, Unit = 1, Config = 1, Parent = 1
#End If

Private Type YogaValue
    Value           As Single
    Unit            As YogaUnit
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const FLOAT_UNDEFINED       As Single = 3.40282347E+38
Private Const FLOAT_EPSILON         As Single = 0.000000001

Public Tag                      As Variant
Private m_oConfig               As cYogaConfig
Private m_lNodePtr              As Long
Private m_oMeasureFn            As Object
Private m_oBaselineFn           As Object
Private m_oParent               As cYogaNode
Private m_cChildren             As Collection
Private m_cLog                  As Collection

'=========================================================================
' Properties
'=========================================================================

Property Get Config() As cYogaConfig
    Set Config = m_oConfig
End Property

Property Get NodePtr() As Long
    If m_lNodePtr = 0 Then
        m_lNodePtr = YGNodeNew()
        Call YGNodeSetContext(m_lNodePtr, ObjPtr(Me))
    End If
    NodePtr = m_lNodePtr
End Property

Property Get IsDirty() As Boolean
    IsDirty = (YGNodeIsDirty(NodePtr) <> 0)
End Property

Public Sub MarkDirty()
    Call YGNodeMarkDirty(NodePtr)
End Sub

Property Get HasNewLayout() As Boolean
    HasNewLayout = (YGNodeGetHasNewLayout(NodePtr) <> 0)
End Property

Public Sub MarkHasNewLayout()
    Call YGNodeSetHasNewLayout(NodePtr, 1)
End Sub

Public Sub MarkLayoutSeen()
    Call YGNodeSetHasNewLayout(NodePtr, 0)
End Sub

Property Get Parent() As cYogaNode
    Set Parent = m_oParent
End Property

Property Set Parent(ByVal oValue As cYogaNode)
    Call CopyMemory(m_oParent, oValue, 4)
End Property

Property Get IsMeasureDefined() As Boolean
    IsMeasureDefined = Not m_oMeasureFn Is Nothing
End Property

Property Get IsBaselineDefined() As Boolean
    IsBaselineDefined = Not m_oBaselineFn Is Nothing
End Property

Property Get StyleDirection() As YogaDirection
    StyleDirection = YGNodeStyleGetDirection(NodePtr)
End Property

Property Let StyleDirection(ByVal eValue As YogaDirection)
    Call YGNodeStyleSetDirection(NodePtr, eValue)
End Property

Property Get flexDirection() As YogaFlexDirection
    flexDirection = YGNodeStyleGetFlexDirection(NodePtr)
End Property

Property Let flexDirection(ByVal eValue As YogaFlexDirection)
    Call YGNodeStyleSetFlexDirection(NodePtr, eValue)
End Property

Property Get JustifyContent() As YogaJustify
    JustifyContent = YGNodeStyleGetJustifyContent(NodePtr)
End Property

Property Let JustifyContent(ByVal eValue As YogaJustify)
    Call YGNodeStyleSetJustifyContent(NodePtr, eValue)
End Property

Property Get Display() As YogaDisplay
    Display = YGNodeStyleGetDisplay(NodePtr)
End Property

Property Let Display(ByVal eValue As YogaDisplay)
    Call YGNodeStyleSetDisplay(NodePtr, eValue)
End Property

Property Get AlignItems() As YogaAlign
    AlignItems = YGNodeStyleGetAlignItems(NodePtr)
End Property

Property Let AlignItems(ByVal eValue As YogaAlign)
    Call YGNodeStyleSetAlignItems(NodePtr, eValue)
End Property

Property Get AlignSelf() As YogaAlign
    AlignSelf = YGNodeStyleGetAlignSelf(NodePtr)
End Property

Property Let AlignSelf(ByVal eValue As YogaAlign)
    Call YGNodeStyleSetAlignSelf(NodePtr, eValue)
End Property

Property Get AlignContent() As YogaAlign
    AlignContent = YGNodeStyleGetAlignContent(NodePtr)
End Property

Property Let AlignContent(ByVal eValue As YogaAlign)
    Call YGNodeStyleSetAlignContent(NodePtr, eValue)
End Property

Property Get PositionType() As YogaPositionType
    PositionType = YGNodeStyleGetPositionType(NodePtr)
End Property

Property Let PositionType(ByVal eValue As YogaPositionType)
    Call YGNodeStyleSetPositionType(NodePtr, eValue)
End Property

Property Get Wrap() As YogaWrap
    Wrap = YGNodeStyleGetFlexWrap(NodePtr)
End Property

Property Let Wrap(ByVal eValue As YogaWrap)
    Call YGNodeStyleSetFlexWrap(NodePtr, eValue)
End Property

Property Get Flex() As Single
    Err.Raise vbObjectError, , "Write-only property"
End Property

Property Let Flex(ByVal sngValue As Single)
    Call YGNodeStyleSetFlex(NodePtr, sngValue)
End Property

Property Get FlexGrow() As Single
    FlexGrow = YGNodeStyleGetFlexGrow(NodePtr)
End Property

Property Let FlexGrow(ByVal sngValue As Single)
    Call YGNodeStyleSetFlexGrow(NodePtr, sngValue)
End Property

Property Get FlexShrink() As Single
    FlexShrink = YGNodeStyleGetFlexShrink(NodePtr)
End Property

Property Let FlexShrink(ByVal sngValue As Single)
    Call YGNodeStyleSetFlexShrink(NodePtr, sngValue)
End Property

Property Get FlexBasis() As Variant
    FlexBasis = FromYogaValue(YGNodeStyleGetFlexBasis(NodePtr))
End Property

Property Let FlexBasis(vValue As Variant)
    With ToYogaValue(vValue)
        Select Case .Unit
        Case yogaUnitPercent
            Call YGNodeStyleSetFlexBasisPercent(NodePtr, .Value)
        Case yogaUnitAuto
            Call YGNodeStyleSetFlexBasisAuto(NodePtr)
        Case Else
            Call YGNodeStyleSetFlexBasis(NodePtr, .Value)
        End Select
    End With
End Property

Property Get Width() As Variant
    Width = FromYogaValue(YGNodeStyleGetWidth(NodePtr))
End Property

Property Let Width(vValue As Variant)
    With ToYogaValue(vValue)
        Select Case .Unit
        Case yogaUnitPercent
            Call YGNodeStyleSetWidthPercent(NodePtr, .Value)
        Case yogaUnitAuto
            Call YGNodeStyleSetWidthAuto(NodePtr)
        Case Else
            Call YGNodeStyleSetWidth(NodePtr, .Value)
        End Select
    End With
End Property

Property Get Height() As Variant
    Height = FromYogaValue(YGNodeStyleGetHeight(NodePtr))
End Property

Property Let Height(vValue As Variant)
    With ToYogaValue(vValue)
        Select Case .Unit
        Case yogaUnitPercent
            Call YGNodeStyleSetHeightPercent(NodePtr, .Value)
        Case yogaUnitAuto
            Call YGNodeStyleSetHeightAuto(NodePtr)
        Case Else
            Call YGNodeStyleSetHeight(NodePtr, .Value)
        End Select
    End With
End Property

Property Get MaxWidth() As Variant
    MaxWidth = FromYogaValue(YGNodeStyleGetMaxWidth(NodePtr))
End Property

Property Let MaxWidth(vValue As Variant)
    With ToYogaValue(vValue)
        Select Case .Unit
        Case yogaUnitPercent
            Call YGNodeStyleSetMaxWidthPercent(NodePtr, .Value)
        Case Else
            Call YGNodeStyleSetMaxWidth(NodePtr, .Value)
        End Select
    End With
End Property

Property Get MaxHeight() As Variant
    MaxHeight = FromYogaValue(YGNodeStyleGetMaxHeight(NodePtr))
End Property

Property Let MaxHeight(vValue As Variant)
    With ToYogaValue(vValue)
        Select Case .Unit
        Case yogaUnitPercent
            Call YGNodeStyleSetMaxHeightPercent(NodePtr, .Value)
        Case Else
            Call YGNodeStyleSetMaxHeight(NodePtr, .Value)
        End Select
    End With
End Property

Property Get MinWidth() As Variant
    MinWidth = FromYogaValue(YGNodeStyleGetMinWidth(NodePtr))
End Property

Property Let MinWidth(vValue As Variant)
    With ToYogaValue(vValue)
        Select Case .Unit
        Case yogaUnitPercent
            Call YGNodeStyleSetMinWidthPercent(NodePtr, .Value)
        Case Else
            Call YGNodeStyleSetMinWidth(NodePtr, .Value)
        End Select
    End With
End Property

Property Get MinHeight() As Variant
    MinHeight = FromYogaValue(YGNodeStyleGetMinHeight(NodePtr))
End Property

Property Let MinHeight(vValue As Variant)
    With ToYogaValue(vValue)
        Select Case .Unit
        Case yogaUnitPercent
            Call YGNodeStyleSetMinHeightPercent(NodePtr, .Value)
        Case Else
            Call YGNodeStyleSetMinHeight(NodePtr, .Value)
        End Select
    End With
End Property

Property Get AspectRatio() As Single
    AspectRatio = YGNodeStyleGetAspectRatio(NodePtr)
End Property

Property Let AspectRatio(ByVal sngValue As Single)
    Call YGNodeStyleSetAspectRatio(NodePtr, sngValue)
End Property

Property Get LayoutX() As Single
    LayoutX = YGNodeLayoutGetLeft(NodePtr)
End Property

Property Get LayoutY() As Single
    LayoutY = YGNodeLayoutGetTop(NodePtr)
End Property

Property Get LayoutWidth() As Single
    LayoutWidth = YGNodeLayoutGetWidth(NodePtr)
End Property

Property Get LayoutHeight() As Single
    LayoutHeight = YGNodeLayoutGetHeight(NodePtr)
End Property

Property Get LayoutDirection() As YogaDirection
    LayoutDirection = YGNodeLayoutGetDirection(NodePtr)
End Property

Property Get Overflow() As YogaOverflow
    Overflow = YGNodeStyleGetOverflow(NodePtr)
End Property

Property Let Overflow(ByVal eValue As YogaOverflow)
    Call YGNodeStyleSetOverflow(NodePtr, eValue)
End Property

Property Get Item(ByVal index As Long) As cYogaNode
    Set Item = m_cChildren.Item(index + 1)
End Property

Property Get Count() As Long
    If Not m_cChildren Is Nothing Then
        Count = m_cChildren.Count
    End If
End Property

Property Get IsReferenceBaseline() As Boolean
    IsReferenceBaseline = (YGNodeIsReferenceBaseline(NodePtr) <> 0)
End Property

Property Let IsReferenceBaseline(ByVal bValue As Boolean)
    Call YGNodeSetIsReferenceBaseline(NodePtr, -bValue)
End Property

'= spacing ===============================================================

Property Get Left() As Variant
    Left = FromYogaValue(YGNodeStyleGetPosition(NodePtr, yogaEdgeLeft))
End Property

Property Let Left(vValue As Variant)
    pvSetStylePosition yogaEdgeLeft, vValue
End Property

Property Get Top() As Variant
    Top = FromYogaValue(YGNodeStyleGetPosition(NodePtr, yogaEdgeTop))
End Property

Property Let Top(vValue As Variant)
    pvSetStylePosition yogaEdgeTop, vValue
End Property

Property Get Right() As Variant
    Right = FromYogaValue(YGNodeStyleGetPosition(NodePtr, yogaEdgeRight))
End Property

Property Let Right(vValue As Variant)
    pvSetStylePosition yogaEdgeRight, vValue
End Property

Property Get Bottom() As Variant
    Bottom = FromYogaValue(YGNodeStyleGetPosition(NodePtr, yogaEdgeBottom))
End Property

Property Let Bottom(vValue As Variant)
    pvSetStylePosition yogaEdgeBottom, vValue
End Property

Property Get Start() As Variant
    Start = FromYogaValue(YGNodeStyleGetPosition(NodePtr, yogaEdgeStart))
End Property

Property Let Start(vValue As Variant)
    pvSetStylePosition yogaEdgeStart, vValue
End Property

Property Get End_() As Variant
    End_ = FromYogaValue(YGNodeStyleGetPosition(NodePtr, yogaEdgeEnd))
End Property

Property Let End_(vValue As Variant)
    pvSetStylePosition yogaEdgeEnd, vValue
End Property

Private Sub pvSetStylePosition(ByVal lEdge As YogaEdge, vValue As Variant)
    With ToYogaValue(vValue)
        If .Unit = yogaUnitPercent Then
            Call YGNodeStyleSetPositionPercent(NodePtr, lEdge, .Value)
        Else
            Call YGNodeStyleSetPosition(NodePtr, lEdge, .Value)
        End If
    End With
End Sub

Property Get MarginLeft() As Variant
    MarginLeft = FromYogaValue(YGNodeStyleGetMargin(NodePtr, yogaEdgeLeft))
End Property

Property Let MarginLeft(vValue As Variant)
    pvSetStyleMargin yogaEdgeLeft, vValue
End Property

Property Get MarginTop() As Variant
    MarginTop = FromYogaValue(YGNodeStyleGetMargin(NodePtr, yogaEdgeTop))
End Property

Property Let MarginTop(vValue As Variant)
    pvSetStyleMargin yogaEdgeTop, vValue
End Property

Property Get MarginRight() As Variant
    MarginRight = FromYogaValue(YGNodeStyleGetMargin(NodePtr, yogaEdgeRight))
End Property

Property Let MarginRight(vValue As Variant)
    pvSetStyleMargin yogaEdgeRight, vValue
End Property

Property Get MarginBottom() As Variant
    MarginBottom = FromYogaValue(YGNodeStyleGetMargin(NodePtr, yogaEdgeBottom))
End Property

Property Let MarginBottom(vValue As Variant)
    pvSetStyleMargin yogaEdgeBottom, vValue
End Property

Property Get MarginStart() As Variant
    MarginStart = FromYogaValue(YGNodeStyleGetMargin(NodePtr, yogaEdgeStart))
End Property

Property Let MarginStart(vValue As Variant)
    pvSetStyleMargin yogaEdgeStart, vValue
End Property

Property Get MarginEnd() As Variant
    MarginEnd = FromYogaValue(YGNodeStyleGetMargin(NodePtr, yogaEdgeEnd))
End Property

Property Let MarginEnd(vValue As Variant)
    pvSetStyleMargin yogaEdgeEnd, vValue
End Property

Property Get MarginHorizontal() As Variant
    MarginHorizontal = FromYogaValue(YGNodeStyleGetMargin(NodePtr, yogaEdgeHorizontal))
End Property

Property Let MarginHorizontal(vValue As Variant)
    pvSetStyleMargin yogaEdgeHorizontal, vValue
End Property

Property Get MarginVertical() As Variant
    MarginVertical = FromYogaValue(YGNodeStyleGetMargin(NodePtr, yogaEdgeVertical))
End Property

Property Let MarginVertical(vValue As Variant)
    pvSetStyleMargin yogaEdgeVertical, vValue
End Property

Property Get Margin() As Variant
    Margin = FromYogaValue(YGNodeStyleGetMargin(NodePtr, yogaEdgeAll))
End Property

Property Let Margin(vValue As Variant)
    pvSetStyleMargin yogaEdgeAll, vValue
End Property

Private Sub pvSetStyleMargin(ByVal lEdge As YogaEdge, vValue As Variant)
    With ToYogaValue(vValue)
        Select Case .Unit
        Case yogaUnitPercent
            Call YGNodeStyleSetMarginPercent(NodePtr, lEdge, .Value)
        Case yogaUnitAuto
            Call YGNodeStyleSetMarginAuto(NodePtr, lEdge)
        Case Else
            Call YGNodeStyleSetMargin(NodePtr, lEdge, .Value)
        End Select
    End With
End Sub

Property Get PaddingLeft() As Variant
    PaddingLeft = FromYogaValue(YGNodeStyleGetPadding(NodePtr, yogaEdgeLeft))
End Property

Property Let PaddingLeft(vValue As Variant)
    pvSetStylePadding yogaEdgeLeft, vValue
End Property

Property Get PaddingTop() As Variant
    PaddingTop = FromYogaValue(YGNodeStyleGetPadding(NodePtr, yogaEdgeTop))
End Property

Property Let PaddingTop(vValue As Variant)
    pvSetStylePadding yogaEdgeTop, vValue
End Property

Property Get PaddingRight() As Variant
    PaddingRight = FromYogaValue(YGNodeStyleGetPadding(NodePtr, yogaEdgeRight))
End Property

Property Let PaddingRight(vValue As Variant)
    pvSetStylePadding yogaEdgeRight, vValue
End Property

Property Get PaddingBottom() As Variant
    PaddingBottom = FromYogaValue(YGNodeStyleGetPadding(NodePtr, yogaEdgeBottom))
End Property

Property Let PaddingBottom(vValue As Variant)
    pvSetStylePadding yogaEdgeBottom, vValue
End Property

Property Get PaddingStart() As Variant
    PaddingStart = FromYogaValue(YGNodeStyleGetPadding(NodePtr, yogaEdgeStart))
End Property

Property Let PaddingStart(vValue As Variant)
    pvSetStylePadding yogaEdgeStart, vValue
End Property

Property Get PaddingEnd() As Variant
    PaddingEnd = FromYogaValue(YGNodeStyleGetPadding(NodePtr, yogaEdgeEnd))
End Property

Property Let PaddingEnd(vValue As Variant)
    pvSetStylePadding yogaEdgeEnd, vValue
End Property

Property Get PaddingHorizontal() As Variant
    PaddingHorizontal = FromYogaValue(YGNodeStyleGetPadding(NodePtr, yogaEdgeHorizontal))
End Property

Property Let PaddingHorizontal(vValue As Variant)
    pvSetStylePadding yogaEdgeHorizontal, vValue
End Property

Property Get PaddingVertical() As Variant
    PaddingVertical = FromYogaValue(YGNodeStyleGetPadding(NodePtr, yogaEdgeVertical))
End Property

Property Let PaddingVertical(vValue As Variant)
    pvSetStylePadding yogaEdgeVertical, vValue
End Property

Property Get Padding() As Variant
    Padding = FromYogaValue(YGNodeStyleGetPadding(NodePtr, yogaEdgeAll))
End Property

Property Let Padding(vValue As Variant)
    pvSetStylePadding yogaEdgeAll, vValue
End Property

Private Sub pvSetStylePadding(ByVal lEdge As YogaEdge, vValue As Variant)
    With ToYogaValue(vValue)
        Select Case .Unit
        Case yogaUnitPercent
            Call YGNodeStyleSetPaddingPercent(NodePtr, lEdge, .Value)
        Case Else
            Call YGNodeStyleSetPadding(NodePtr, lEdge, .Value)
        End Select
    End With
End Sub

Property Get BorderLeftWidth() As Single
    BorderLeftWidth = YGNodeStyleGetBorder(NodePtr, yogaEdgeLeft)
End Property

Property Let BorderLeftWidth(ByVal sngValue As Single)
    Call YGNodeStyleSetBorder(NodePtr, yogaEdgeLeft, sngValue)
End Property

Property Get BorderTopWidth() As Single
    BorderTopWidth = YGNodeStyleGetBorder(NodePtr, yogaEdgeTop)
End Property

Property Let BorderTopWidth(ByVal sngValue As Single)
    Call YGNodeStyleSetBorder(NodePtr, yogaEdgeTop, sngValue)
End Property

Property Get BorderRightWidth() As Single
    BorderRightWidth = YGNodeStyleGetBorder(NodePtr, yogaEdgeRight)
End Property

Property Let BorderRightWidth(ByVal sngValue As Single)
    Call YGNodeStyleSetBorder(NodePtr, yogaEdgeRight, sngValue)
End Property

Property Get BorderBottomWidth() As Single
    BorderBottomWidth = YGNodeStyleGetBorder(NodePtr, yogaEdgeBottom)
End Property

Property Let BorderBottomWidth(ByVal sngValue As Single)
    Call YGNodeStyleSetBorder(NodePtr, yogaEdgeBottom, sngValue)
End Property

Property Get BorderStartWidth() As Single
    BorderStartWidth = YGNodeStyleGetBorder(NodePtr, yogaEdgeStart)
End Property

Property Let BorderStartWidth(ByVal sngValue As Single)
    Call YGNodeStyleSetBorder(NodePtr, yogaEdgeStart, sngValue)
End Property

Property Get BorderEndWidth() As Single
    BorderEndWidth = YGNodeStyleGetBorder(NodePtr, yogaEdgeEnd)
End Property

Property Let BorderEndWidth(ByVal sngValue As Single)
    Call YGNodeStyleSetBorder(NodePtr, yogaEdgeEnd, sngValue)
End Property

Property Get BorderWidth() As Single
    BorderWidth = YGNodeStyleGetBorder(NodePtr, yogaEdgeAll)
End Property

Property Let BorderWidth(ByVal sngValue As Single)
    Call YGNodeStyleSetBorder(NodePtr, yogaEdgeAll, sngValue)
End Property

Property Get LayoutMarginLeft() As Single
    LayoutMarginLeft = YGNodeLayoutGetMargin(NodePtr, yogaEdgeLeft)
End Property

Property Get LayoutMarginTop() As Single
    LayoutMarginTop = YGNodeLayoutGetMargin(NodePtr, yogaEdgeTop)
End Property

Property Get LayoutMarginRight() As Single
    LayoutMarginRight = YGNodeLayoutGetMargin(NodePtr, yogaEdgeRight)
End Property

Property Get LayoutMarginBottom() As Single
    LayoutMarginBottom = YGNodeLayoutGetMargin(NodePtr, yogaEdgeBottom)
End Property

Property Get LayoutMarginStart() As Single
    LayoutMarginStart = YGNodeLayoutGetMargin(NodePtr, yogaEdgeStart)
End Property

Property Get LayoutMarginEnd() As Single
    LayoutMarginEnd = YGNodeLayoutGetMargin(NodePtr, yogaEdgeEnd)
End Property

Property Get LayoutPaddingLeft() As Single
    LayoutPaddingLeft = YGNodeLayoutGetPadding(NodePtr, yogaEdgeLeft)
End Property

Property Get LayoutPaddingTop() As Single
    LayoutPaddingTop = YGNodeLayoutGetPadding(NodePtr, yogaEdgeTop)
End Property

Property Get LayoutPaddingRight() As Single
    LayoutPaddingRight = YGNodeLayoutGetPadding(NodePtr, yogaEdgeRight)
End Property

Property Get LayoutPaddingBottom() As Single
    LayoutPaddingBottom = YGNodeLayoutGetPadding(NodePtr, yogaEdgeBottom)
End Property

Property Get LayoutPaddingStart() As Single
    LayoutPaddingStart = YGNodeLayoutGetPadding(NodePtr, yogaEdgeStart)
End Property

Property Get LayoutPaddingEnd() As Single
    LayoutPaddingEnd = YGNodeLayoutGetPadding(NodePtr, yogaEdgeEnd)
End Property

Friend Property Get frMeasureFn() As Object
    Set frMeasureFn = m_oMeasureFn
End Property

Friend Property Get frBaselineFn() As Object
    Set frBaselineFn = m_oBaselineFn
End Property

'=========================================================================
' Methods
'=========================================================================

Public Sub Init(oConfig As cYogaConfig, Optional ByVal lSrcNodePtr As Long)
    If Not oConfig Is Nothing Then
        Set m_oConfig = oConfig
    Else
        Set m_oConfig = YogaConfigDefault
    End If
    m_lNodePtr = YGNodeNewWithConfig(m_oConfig.ConfigPtr)
    Call YGNodeSetContext(m_lNodePtr, ObjPtr(Me))
    If lSrcNodePtr <> 0 Then
        Call YGNodeCopyStyle(m_lNodePtr, lSrcNodePtr)
    End If
End Sub

Public Sub Reset()
    Set m_oMeasureFn = Nothing
    Set m_oBaselineFn = Nothing
    Tag = Empty
    Call YGNodeReset(NodePtr)
    Call YGNodeSetContext(NodePtr, ObjPtr(Me))
End Sub

Public Sub CopyStyle(oSrcNode As cYogaNode)
    Call YGNodeCopyStyle(NodePtr, oSrcNode.NodePtr)
End Sub

Public Function ValuesEqual(ByVal f1 As Single, ByVal f2 As Single) As Boolean
    ValuesEqual = Abs(f1 - f2) < FLOAT_EPSILON
End Function

Public Sub Insert(ByVal index As Long, oChild As cYogaNode)
    If m_cChildren Is Nothing Then
        Set m_cChildren = New Collection
    End If
    If index < m_cChildren.Count Then
        m_cChildren.Add oChild, Before:=index + 1
    Else
        m_cChildren.Add oChild
    End If
    Set oChild.Parent = Me
    Call YGNodeInsertChild(NodePtr, oChild.NodePtr, index)
End Sub

Public Sub RemoveAt(ByVal index As Long)
    Dim oChild              As cYogaNode
    
    Set oChild = m_cChildren.Item(index + 1)
    m_cChildren.Remove index + 1
    Set oChild.Parent = Nothing
    Call YGNodeRemoveChild(NodePtr, oChild.NodePtr)
End Sub

Public Sub AddChild(oNode As cYogaNode)
    Insert Count, oNode
End Sub

Public Sub RemoveChild(oChild As cYogaNode)
    Dim lIdx            As Long
    
    lIdx = IndexOf(oChild)
    If lIdx >= 0 Then
        RemoveAt lIdx
    End If
End Sub

Public Sub Clear()
    If Not m_cChildren Is Nothing Then
        Do While m_cChildren.Count > 0
            RemoveAt m_cChildren.Count - 1
        Loop
    End If
End Sub

Public Function IndexOf(oNode As cYogaNode) As Long
    Dim lIdx            As Long
    
    For lIdx = 1 To m_cChildren.Count
        If m_cChildren.Item(lIdx) Is oNode Then
            IndexOf = lIdx - 1
            Exit Function
        End If
    Next
    IndexOf = -1
End Function

Public Sub SetMeasureFunction(oCallback As Object)
    Set m_oMeasureFn = oCallback
    If Not oCallback Is Nothing Then
        Call YGNodeSetMeasureFunc(NodePtr, AddressOf YogaNodeMeasureRedirect)
    Else
        Call YGNodeSetMeasureFunc(NodePtr, 0)
    End If
End Sub

Public Sub SetBaselineFunction(oCallback As Object)
    Set m_oBaselineFn = oCallback
    If Not oCallback Is Nothing Then
        Call YGNodeSetBaselineFunc(NodePtr, AddressOf YogaNodeBaselineRedirect)
    Else
        Call YGNodeSetBaselineFunc(NodePtr, 0)
    End If
End Sub

Public Sub CalculateLayout( _
            Optional ByVal sngWidth As Single = FLOAT_UNDEFINED, _
            Optional ByVal sngHeight As Single = FLOAT_UNDEFINED)
    If sngWidth = FLOAT_UNDEFINED Then
        sngWidth = YogaFloatNan
    End If
    If sngHeight = FLOAT_UNDEFINED Then
        sngHeight = YogaFloatNan
    End If
    Call YGNodeCalculateLayout(NodePtr, sngWidth, sngHeight, YGNodeStyleGetDirection(NodePtr))
End Sub

Public Function ToString(Optional ByVal eOptions As YogaPrintOptions = yogaProLayout Or yogaProStyle Or yogaProChildren) As String
    Dim oOrig           As Object
    
    Set oOrig = m_oConfig.Logger
    Set m_oConfig.Logger = Me
    Set m_cLog = New Collection
    Call YGNodePrint(NodePtr, eOptions)
    ToString = ConcatCollection(m_cLog)
    Set m_cLog = Nothing
    Set m_oConfig.Logger = oOrig
End Function

Public Function GetEnumerator() As IUnknown
Attribute GetEnumerator.VB_UserMemId = -4
Attribute GetEnumerator.VB_MemberFlags = "40"
    If m_cChildren Is Nothing Then
        Set m_cChildren = New Collection
    End If
    Set GetEnumerator = m_cChildren.[_NewEnum]
End Function

'--- hidden
Public Sub LogCallback(oNode As cYogaNode, ByVal eLevel As YogaLogLevel, sMessage As String)
Attribute LogCallback.VB_MemberFlags = "40"
    #If oNode And eLevel Then '--- touch args
    #End If
    If Not m_cLog Is Nothing Then
        m_cLog.Add Replace(sMessage, vbLf, vbCrLf)
    End If
End Sub

'= private ===============================================================

Private Function ToYogaValue(vValue As Variant) As YogaValue
    If IsArray(vValue) Then
        ToYogaValue.Unit = vValue(0)
        If ToYogaValue.Unit = yogaUnitAuto Or ToYogaValue.Unit = yogaUnitUndefined Then
            ToYogaValue.Value = YogaFloatNan
        Else
            ToYogaValue.Value = vValue(1)
        End If
    Else
        ToYogaValue.Value = vValue
        If YGFloatIsUndefined(ToYogaValue.Value) <> 0 Then
            ToYogaValue.Unit = yogaUnitUndefined
        Else
            ToYogaValue.Unit = yogaUnitPoint
        End If
    End If
End Function

Private Function FromYogaValue(uValue As YogaValue) As Variant
    FromYogaValue = Array(uValue.Unit, uValue.Value)
End Function

Private Function ConcatCollection(oCol As Collection, Optional Separator As String = vbCrLf) As String
    Dim lSize           As Long
    Dim vElem           As Variant
    
    For Each vElem In oCol
        lSize = lSize + Len(vElem) + Len(Separator)
    Next
    If lSize > 0 Then
        ConcatCollection = String$(lSize - Len(Separator), 0)
        lSize = 1
        For Each vElem In oCol
            If lSize <= Len(ConcatCollection) Then
                Mid$(ConcatCollection, lSize, Len(vElem) + Len(Separator)) = vElem & Separator
            End If
            lSize = lSize + Len(vElem) + Len(Separator)
        Next
    End If
End Function

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Terminate()
    If m_lNodePtr <> 0 Then
        Call YGNodeFree(m_lNodePtr)
        m_lNodePtr = 0
    End If
    Call CopyMemory(m_oParent, 0&, 4)
End Sub


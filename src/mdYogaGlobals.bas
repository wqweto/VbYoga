Attribute VB_Name = "mdYogaGlobals"
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
' Public Types
'=========================================================================

#If False Then
Const Value = 1, Unit = 1, Width = 1, Height = 1
#End If

Public Type YogaValue
    Value           As Single
    Unit            As YogaUnit
End Type

Private Type YogaSize
    Width           As Single
    Height          As Single
End Type

'=========================================================================
' API
'=========================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function YGConfigGetDefault Lib "yoga" Alias "_YGConfigGetDefault@0" () As Long
Private Declare Function YGConfigGetContext Lib "yoga" Alias "_YGConfigGetContext@4" (ByVal lConfigPtr As Long) As Long
Private Declare Function YGConfigGetInstanceCount Lib "yoga" Alias "_YGConfigGetInstanceCount@0" () As Long
Private Declare Function YGInteropSetLogger Lib "yoga" Alias "_YGInteropSetLogger@4" (ByVal pfn As Long) As Long
Private Declare Function YGNodeGetContext Lib "yoga" Alias "_YGNodeGetContext@4" (ByVal lNodePtr As Long) As Long
Private Declare Function YGNodeGetInstanceCount Lib "yoga" Alias "_YGNodeGetInstanceCount@0" () As Long
Private Declare Function YGFloatIsUndefined Lib "yoga" Alias "_YGFloatIsUndefined@4" (ByVal sngValue As Single) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Public Const FLOAT_UNDEFINED        As Single = 3.40282347E+38
Private Const FLOAT_NAN_BYTES       As Long = &HFFC00000

Public YogaFloatNan             As Single
Public YogaDefConfigPtr         As Long
Private m_oDefaultConfig        As Object

'=========================================================================
' Functions
'=========================================================================

Public Function YogaConfigDefault() As cYogaConfig
    If YogaDefConfigPtr = 0 Then
        Call CopyMemory(YogaFloatNan, FLOAT_NAN_BYTES, 4)
        YogaDefConfigPtr = YGConfigGetDefault()
        Set m_oDefaultConfig = YogaConfigNew(YogaDefConfigPtr)
        Call YGInteropSetLogger(AddressOf YogaConfigLoggerRedirect)
    End If
    Set YogaConfigDefault = m_oDefaultConfig
End Function

Public Function YogaConfigNew(Optional ByVal lConfigPtr As Long) As cYogaConfig
    Set YogaConfigNew = New cYogaConfig
    YogaConfigNew.Init lConfigPtr
End Function

Property Get YogaConfigInstanceCount() As Long
    YogaConfigInstanceCount = YGConfigGetInstanceCount()
End Property

Public Function YogaNodeNew(Optional oConfig As cYogaConfig) As cYogaNode
    Set YogaNodeNew = New cYogaNode
    If Not oConfig Is Nothing Then
        YogaNodeNew.Init oConfig
    Else
        YogaNodeNew.Init YogaConfigDefault
    End If
End Function

Public Function YogaNodeClone(oSrcNode As cYogaNode) As cYogaNode
    Set YogaNodeClone = New cYogaNode
    YogaNodeClone.Init oSrcNode.Config, oSrcNode.NodePtr
End Function

Public Function YogaNodeInstanceCount() As Long
    YogaNodeInstanceCount = YGNodeGetInstanceCount()
End Function

Public Function YogaNodeMeasureRedirect( _
            ByVal lNodePtr As Long, _
            ByVal sngWidth As Single, _
            ByVal eWidthMode As YogaMeasureMode, _
            ByVal sngHeight As Single, _
            ByVal eHeightMode As YogaMeasureMode) As YogaSize
    Const FUNC_NAME     As String = "YogaNodeMeasureRedirect"
    Dim oNode           As cYogaNode
    
    On Error GoTo EH
    Set oNode = pvToObject(YGNodeGetContext(lNodePtr))
    Call oNode.frMeasureFn.MeasureCallback(oNode, sngWidth, eWidthMode, sngHeight, eHeightMode, _
        YogaNodeMeasureRedirect.Width, YogaNodeMeasureRedirect.Height)
    Exit Function
EH:
    Debug.Print "Critical error: " & Err.Description & " [" & FUNC_NAME & "]"
End Function

Public Function YogaNodeBaselineRedirect( _
            ByVal lNodePtr As Long, _
            ByVal sngWidth As Single, _
            ByVal sngHeight As Single) As Single
    Const FUNC_NAME     As String = "YogaNodeMeasureRedirect"
    Dim oNode           As cYogaNode
    
    Set oNode = pvToObject(YGNodeGetContext(lNodePtr))
    YogaNodeBaselineRedirect = oNode.frBaselineFn.BaselineCallback(oNode, sngWidth, sngHeight)
    Exit Function
EH:
    Debug.Print "Critical error: " & Err.Description & " [" & FUNC_NAME & "]"
End Function

Public Function ToYogaValue(vValue As Variant) As YogaValue
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

Public Function FromYogaValue(uValue As YogaValue) As Variant
    FromYogaValue = Array(uValue.Unit, uValue.Value)
End Function

Public Function YogaConstantsIsUndefined(vValue As Variant) As Boolean
    If IsArray(vValue) Then
        YogaConstantsIsUndefined = (vValue(0) = yogaUnitUndefined)
    Else
        YogaConstantsIsUndefined = (YGFloatIsUndefined(vValue) <> 0)
    End If
End Function

'= private ===============================================================

Private Function YogaConfigLoggerRedirect( _
            ByVal lConfigPtr As Long, _
            ByVal lNodePtr As Long, _
            ByVal eLevel As YogaLogLevel, _
            ByVal lMsgPtr As Long) As Long
    Const FUNC_NAME     As String = "YogaConfigLoggerRedirect"
    Dim oConfig         As cYogaConfig
    Dim oNode           As cYogaNode
    Dim sMessage        As String
    Dim bLogged         As Boolean
    
    On Error GoTo EH
    Set oConfig = pvToObject(YGConfigGetContext(lConfigPtr))
    Set oNode = pvToObject(YGNodeGetContext(lNodePtr))
    sMessage = pvToString(lMsgPtr)
    If Right$(sMessage, 1) = vbLf Then
        sMessage = Left$(sMessage, Len(sMessage) - 1)
    End If
    If Right$(sMessage, 1) = vbCr Then
        sMessage = Left$(sMessage, Len(sMessage) - 1)
    End If
    If Right$(sMessage, 1) = "." Then
        sMessage = Left$(sMessage, Len(sMessage) - 1)
    End If
    If Not oConfig Is Nothing Then
        If Not oConfig.Logger Is Nothing Then
            Call oConfig.Logger.LogCallback(oNode, eLevel, sMessage)
            bLogged = True
        End If
    End If
    If LenB(sMessage) <> 0 Then
        '--- use default "logging" if not suppressed by logger
        If (eLevel = yogaLogError Or eLevel = yogaLogFatal) Then
            On Error GoTo 0
            Err.Raise vbObjectError, , sMessage
        ElseIf Not bLogged Then
            Debug.Print "VbYoga: eLevel=" & eLevel & ", lMsgPtr=" & sMessage
        End If
    End If
    Exit Function
EH:
    Debug.Print "Critical error: " & Err.Description & " [" & FUNC_NAME & "]"
End Function

Private Function pvToString(ByVal lPtr As Long) As String
    If lPtr <> 0 Then
        pvToString = String$(lstrlen(lPtr), Chr$(0))
        Call CopyMemory(ByVal pvToString, ByVal lPtr, lstrlen(lPtr))
    End If
End Function

Private Function pvToObject(ByVal lPtr As Long) As Object
    Dim pUnk            As IUnknown
    
    If lPtr <> 0 Then
        Call CopyMemory(pUnk, lPtr, 4)
        Set pvToObject = pUnk
        Call CopyMemory(pUnk, 0&, 4)
    End If
End Function

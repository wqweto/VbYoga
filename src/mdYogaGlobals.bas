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
' API
'=========================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function YGConfigGetDefault Lib "yoga" Alias "_YGConfigGetDefault@0" () As Long
Private Declare Function YGConfigGetContext Lib "yoga" Alias "_YGConfigGetContext@4" (ByVal lConfigPtr As Long) As Long
Private Declare Function YGConfigGetInstanceCount Lib "yoga" Alias "_YGConfigGetInstanceCount@0" () As Long
Private Declare Function YGInteropSetLogger Lib "yoga" Alias "_YGInteropSetLogger@4" (ByVal pfn As Long) As Long
Private Declare Function YGNodeGetContext Lib "yoga" Alias "_YGNodeGetContext@4" (ByVal lNodePtr As Long) As Long
Private Declare Function YGNodeGetInstanceCount Lib "yoga" Alias "_YGNodeGetInstanceCount@0" () As Long
Private Declare Function YGFloatIsUndefined Lib "yoga" Alias "_YGFloatIsUndefined@4" (ByVal sngValue As Single) As Byte

#If False Then
Const Width = 1, Height = 1
#End If

Private Type YogaSize
    Width           As Single
    Height          As Single
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const FLOAT_NAN_BYTES       As Long = &HFFC00000

Public YogaFloatNan             As Single
Public YogaDefConfigPtr         As Long
Private m_oDefaultConfig        As Object

'=========================================================================
' Functions
'=========================================================================

Public Function YogaConfigDefault() As cYogaConfig
    If YogaDefConfigPtr = 0 Then
        If GetModuleHandle("yoga.dll") = 0 Then
            Call LoadLibrary(LocateFile(App.Path & "\yoga.dll"))
        End If
        Call CopyMemory(YogaFloatNan, FLOAT_NAN_BYTES, 4)
        YogaDefConfigPtr = YGConfigGetDefault()
        Set m_oDefaultConfig = YogaConfigNew(YogaDefConfigPtr)
        Call YGInteropSetLogger(AddressOf pvYogaConfigLoggerRedirect)
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
            ByVal eWidthMode As Long, _
            ByVal sngHeight As Single, _
            ByVal eHeightMode As Long) As YogaSize
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

Public Function YogaConstantsIsUndefined(vValue As Variant) As Boolean
    If IsArray(vValue) Then
        YogaConstantsIsUndefined = (vValue(0) = yogaUnitUndefined)
    Else
        YogaConstantsIsUndefined = (YGFloatIsUndefined(vValue) <> 0)
    End If
End Function

'= private ===============================================================

Private Function pvYogaConfigLoggerRedirect( _
            ByVal lConfigPtr As Long, _
            ByVal lNodePtr As Long, _
            ByVal eLevel As YogaLogLevel, _
            ByVal lMsgPtr As Long) As Long
    Const FUNC_NAME     As String = "pvYogaConfigLoggerRedirect"
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

Private Function LocateFile(sFile As String) As String
    LocateFile = sFile
End Function

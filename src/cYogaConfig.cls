VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cYogaConfig"
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
Private Const MODULE_NAME As String = "cYogaConfig"

'=========================================================================
' API
'=========================================================================

Private Declare Function YGConfigNew Lib "yoga" Alias "_YGConfigNew@0" () As Long
Private Declare Function YGConfigFree Lib "yoga" Alias "_YGConfigFree@4" (ByVal lConfigPtr As Long) As Long
Private Declare Function YGConfigSetContext Lib "yoga" Alias "_YGConfigSetContext@8" (ByVal lConfigPtr As Long, ByVal lContext As Long) As Long
Private Declare Function YGConfigIsExperimentalFeatureEnabled Lib "yoga" Alias "_YGConfigIsExperimentalFeatureEnabled@8" (ByVal lConfigPtr As Long, ByVal lFeature As Long) As Byte
Private Declare Function YGConfigSetExperimentalFeatureEnabled Lib "yoga" Alias "_YGConfigSetExperimentalFeatureEnabled@12" (ByVal lConfigPtr As Long, ByVal lFeature As Long, ByVal lEnabled As Long) As Long
Private Declare Function YGConfigGetUseWebDefaults Lib "yoga" Alias "_YGConfigGetUseWebDefaults@4" (ByVal lConfigPtr As Long) As Byte
Private Declare Function YGConfigSetUseWebDefaults Lib "yoga" Alias "_YGConfigSetUseWebDefaults@8" (ByVal lConfigPtr As Long, ByVal lEnabled As Long) As Long
Private Declare Function YGConfigSetPointScaleFactor Lib "yoga" Alias "_YGConfigSetPointScaleFactor@8" (ByVal lConfigPtr As Long, ByVal sngFactor As Single) As Long

'=========================================================================
' Constants and member variables
'=========================================================================

Public Logger                   As Object
Private m_lConfigPtr            As Long
'--- debug
#If DebugMode Then
    Private m_sDebugID          As String
#End If

'=========================================================================
' Properties
'=========================================================================

Property Get ConfigPtr() As Long
    If m_lConfigPtr = 0 Then
        m_lConfigPtr = YGConfigNew()
        Call YGConfigSetContext(m_lConfigPtr, ObjPtr(Me))
    End If
    ConfigPtr = m_lConfigPtr
End Property

Property Get ExperimentalFeature(ByVal eValue As YogaExperimentalFeature) As Boolean
    ExperimentalFeature = (YGConfigIsExperimentalFeatureEnabled(ConfigPtr, eValue) <> 0)
End Property

Property Let YogaConfigExperimentalFeature(ByVal eValue As YogaExperimentalFeature, ByVal bValue As Boolean)
    Call YGConfigSetExperimentalFeatureEnabled(ConfigPtr, eValue, -bValue)
End Property

Property Get UseWebDefaults() As Boolean
    UseWebDefaults = (YGConfigGetUseWebDefaults(ConfigPtr) <> 0)
End Property

Property Let UseWebDefaults(ByVal bValue As Boolean)
    Call YGConfigSetUseWebDefaults(ConfigPtr, -bValue)
End Property

Property Get PointScaleFactor() As Single
    Err.Raise vbObjectError, , "Write-only property"
End Property

Property Let PointScaleFactor(ByVal sngValue As Single)
    Call YGConfigSetPointScaleFactor(ConfigPtr, sngValue)
End Property

'=========================================================================
' Methods
'=========================================================================

Public Sub Init(Optional ByVal lConfigPtr As Long)
    If lConfigPtr <> 0 Then
        m_lConfigPtr = lConfigPtr
    Else
        m_lConfigPtr = YGConfigNew()
    End If
    Call YGConfigSetContext(m_lConfigPtr, ObjPtr(Me))
End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub Class_Initialize()
    #If DebugMode Then
        DebugInstanceInit MODULE_NAME, m_sDebugID, Me
    #End If
    Call YogaConfigDefault
End Sub

Private Sub Class_Terminate()
    If m_lConfigPtr <> 0 And m_lConfigPtr <> YogaDefConfigPtr Then
        Call YGConfigFree(m_lConfigPtr)
        m_lConfigPtr = 0
    End If
    #If DebugMode Then
        DebugInstanceTerm MODULE_NAME, m_sDebugID
    #End If
End Sub

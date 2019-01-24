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

'=========================================================================
' API
'=========================================================================

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'=========================================================================
' Constants and member variables
'=========================================================================

Private m_oFlexBox              As ctxFlexContainer     '--- weak-ref
Private m_oParent               As cFlexDomNode         '--- weak-ref
Private m_cChildren             As Collection
Private m_oControl              As Object               '--- weak-ref
Public Name                     As String
Public CssClass                 As String
Public Layout                   As cYogaNode
'--- debug
#If DebugMode Then
    Private m_sDebugID          As String
#End If

'=========================================================================
' Properties
'=========================================================================

Friend Property Get frFlexBox() As ctxFlexContainer
    If Not m_oFlexBox Is Nothing Then
        Set frFlexBox = m_oFlexBox
    ElseIf Not m_oParent Is Nothing Then
        Set frFlexBox = m_oParent.frFlexBox
    End If
End Property

Friend Property Set frFlexBox(oValue As ctxFlexContainer)
    Call CopyMemory(m_oFlexBox, oValue, 4)
End Property

Property Get Parent() As cFlexDomNode
    Set Parent = m_oParent
End Property

Friend Property Set frParent(oValue As cFlexDomNode)
    Call CopyMemory(m_oParent, oValue, 4)
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
    Set Control = m_oControl
End Property

Friend Property Set frControl(oValue As Object)
    Call CopyMemory(m_oControl, oValue, 4)
End Property

Property Get CtlLabel() As VB.Label
    Set CtlLabel = m_oControl
End Property

Property Get CtlButton() As ctxNineButton
    Set CtlButton = m_oControl
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

Public Sub ApplyLayout()
    Dim oItem           As cFlexDomNode
    
    If Not Control Is Nothing Then
        With Layout
            If .Display = yogaDisplayNone Then
                Control.Visible = False
            Else
                Control.Visible = True
                Control.Move .GetRelativeX, .GetRelativeY, .LayoutWidth, .LayoutHeight
            End If
        End With
    End If
    If Not m_cChildren Is Nothing Then
        For Each oItem In m_cChildren
            oItem.ApplyLayout
        Next
    End If
End Sub

Public Sub MeasureCallback( _
            oNode As Object, _
            ByVal sngWidth As Single, _
            ByVal eWidthMode As YogaMeasureMode, _
            ByVal sngHeight As Single, _
            ByVal eHeightMode As YogaMeasureMode, _
            sngLayoutWidth As Single, _
            sngLayoutHeight As Single)
    With CtlLabel
        If .AutoSize Then
            If .WordWrap Then
                .Width = Int(sngWidth / Screen.TwipsPerPixelX) * Screen.TwipsPerPixelX
            Else
                .Width = .Width
            End If
        End If
        sngLayoutWidth = .Width
        sngLayoutHeight = .Height
    End With
End Sub

'=========================================================================
' Base class events
'=========================================================================

#If DebugMode Then
    Private Sub Class_Initialize()
        DebugInstanceInit MODULE_NAME, m_sDebugID, Me
    End Sub
#End If

Private Sub Class_Terminate()
    Call CopyMemory(m_oFlexBox, 0&, 4)
    Call CopyMemory(m_oParent, 0&, 4)
    Set m_cChildren = Nothing
    Call CopyMemory(m_oControl, 0&, 4)
    Set Layout = Nothing
    #If DebugMode Then
        DebugInstanceTerm MODULE_NAME, m_sDebugID
    #End If
End Sub

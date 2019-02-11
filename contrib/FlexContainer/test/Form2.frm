VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000005&
   Caption         =   "Form2"
   ClientHeight    =   5436
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6336
   LinkTopic       =   "Form2"
   ScaleHeight     =   5436
   ScaleWidth      =   6336
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picScroll 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   4800
      Left            =   84
      ScaleHeight     =   4800
      ScaleWidth      =   6060
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   504
      Width           =   6060
      Begin Project1.ctxFlexContainer ctxFlexContainer1 
         Height          =   3876
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   4128
         _ExtentX        =   7281
         _ExtentY        =   6837
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoApplyLayout =   -1  'True
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reset"
      Height          =   348
      Left            =   1428
      TabIndex        =   1
      Top             =   84
      Width           =   1272
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Form"
      Height          =   348
      Left            =   84
      TabIndex        =   0
      Top             =   84
      Width           =   1188
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_oCtlCancelMode        As Object

Private Sub Command1_Click()
    With New Form2
        .Show
    End With
End Sub

Private Sub Command2_Click()
    ctxFlexContainer1.Reset
End Sub

Private Sub ctxFlexContainer1_Click(DomNode As cFlexDomNode)
    Debug.Print DomNode.Control.Name & "(" & DomNode.Control.Index & ") clicked", Timer
End Sub

Private Sub ctxFlexContainer1_StartDrag(Button As Integer, Shift As Integer, X As Single, Y As Single, Handled As Boolean)
    Handled = True
End Sub

Private Sub ctxFlexContainer1_MouseMove(DomNode As cFlexDomNode, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim dblMin          As Double
    
    If ctxFlexContainer1.Dragging Then
        dblMin = Limit(picScroll.Width - ctxFlexContainer1.Width, , 0)
        ctxFlexContainer1.Left = Limit(X - (ctxFlexContainer1.DownX - ctxFlexContainer1.Left), dblMin, 0)
        dblMin = Limit(picScroll.Height - ctxFlexContainer1.Height, , 0)
        ctxFlexContainer1.Top = Limit(Y - (ctxFlexContainer1.DownY - ctxFlexContainer1.Top), dblMin, 0)
    End If
End Sub

Private Sub Form_Load()
    With ctxFlexContainer1.Root
        .Layout.Overflow = yogaOverflowScroll
        With .AddLabel(CssClass:="section-caption")
            .Control.Caption = "Section A"
        End With
        With .AddContainer(CssClass:="top-section")
            .Layout.FlexDirection = yogaFlexRow
            With .AddButton()
                .CtlButton.Style = ucsBtyButtonDefault
                .Layout.Width = Array(yogaUnitPercent, 30)
            End With
        End With
        With .AddLabel(CssClass:="section-caption")
            .Control.Caption = "Section B"
        End With
        With .AddContainer(CssClass:="main-section")
            .Layout.FlexDirection = yogaFlexRow
            .Layout.MinHeight = 1200
            With .AddButton()
                .CtlButton.Style = ucsBtyButtonDefault
                With .Layout
                    .Width = Array(yogaUnitPercent, 30)
                    .MinWidth = 1200
                    .MaxHeight = 1200
                    .Height = Array(yogaUnitPercent, 100)
                End With
            End With
            With .AddButton()
                .CtlButton.Style = ucsBtyButtonTurnGreen
                With .Layout
                    .Width = Array(yogaUnitPercent, 30)
                    .MinWidth = 1200
                    .MaxHeight = 1200
                End With
            End With
            With .AddButton()
                .CtlButton.Style = ucsBtyButtonTurnRed
                With .Layout
                    .Width = Array(yogaUnitPercent, 30)
                    .MinWidth = 1200
                    .MaxHeight = 1200
                End With
            End With
            With .AddLabel()
                .CtlLabel.WordWrap = True
                .CtlLabel.Caption = "In publishing and graphic design, lorem ipsum is a placeholder text commonly used to demonstrate the visual form of a document without relying on meaningful"
                With .Layout
                    .Width = Array(yogaUnitPercent, 9.9)
                End With
            End With
            With .AddButton()
                .CtlButton.Style = ucsBtyButtonDefault
                With .Layout
                    .Width = Array(yogaUnitPercent, 30)
                    .MinWidth = 1200
                    .MaxHeight = 1200
                    .Height = Array(yogaUnitPercent, 100)
                End With
            End With
            With .AddButton()
                .CtlButton.Style = ucsBtyButtonTurnGreen
                With .Layout
                    .Width = Array(yogaUnitPercent, 30)
                    .MinWidth = 1200
                    .MaxHeight = 1200
                End With
            End With
            With .AddButton()
                .CtlButton.Style = ucsBtyButtonTurnRed
                With .Layout
                    .Width = Array(yogaUnitPercent, 30)
                    .MinWidth = 1200
                    .MaxHeight = 1200
                End With
            End With
        End With
        With .AddLabel(CssClass:="section-caption")
            .Control.Caption = "Section C"
        End With
        With .AddContainer(CssClass:="bottom-section")
            .Layout.FlexDirection = yogaFlexRow
            With .AddButton()
                .CtlButton.Style = ucsBtyButtonDefault
                With .Layout
                    .Width = Array(yogaUnitPercent, 30)
                    .FlexShrink = 1
                End With
            End With
        End With
    End With
    Set ctxFlexContainer1.Styles = pvGetStyles()
    Form_Resize
End Sub

Private Function pvGetStyles() As Object
    Dim oStyle          As Object
    
    Set pvGetStyles = CreateObject("Scripting.Dictionary")
    Set oStyle = CreateObject("Scripting.Dictionary")
    Set pvGetStyles.Item(".root") = oStyle
    oStyle.Item("flex-direction") = "column"
    oStyle.Item("padding") = 120
    Set oStyle = CreateObject("Scripting.Dictionary")
    Set pvGetStyles.Item(".top-section") = oStyle
    oStyle.Item("width") = "100%"
    oStyle.Item("height") = 1200
    oStyle.Item("justify-content") = "center"
    Set oStyle = CreateObject("Scripting.Dictionary")
    Set pvGetStyles.Item(".main-section") = oStyle
    oStyle.Item("height") = 1200
    oStyle.Item("flex-grow") = 1
    oStyle.Item("flex-shrink") = 1
    oStyle.Item("flex-wrap") = "wrap"
    oStyle.Item("align-content") = "flex-start"
    Set oStyle = CreateObject("Scripting.Dictionary")
    Set pvGetStyles.Item(".bottom-section") = oStyle
    oStyle.Item("height") = 1200
    Set oStyle = CreateObject("Scripting.Dictionary")
    Set pvGetStyles.Item(".section-caption") = oStyle
    oStyle.Item("min-height") = 240
    oStyle.Item("margin-left") = 60
    oStyle.Item("flex-shrink") = 0
End Function

Private Sub Form_Resize()
    Dim dblMin          As Double
    
    If WindowState <> vbMinimized And ScaleHeight > picScroll.Top Then
        picScroll.Move 0, picScroll.Top, ScaleWidth, ScaleHeight - picScroll.Top
        ctxFlexContainer1.ApplyLayout picScroll.ScaleWidth, picScroll.ScaleHeight
        ctxFlexContainer1.Width = ctxFlexContainer1.Root.Layout.GetActualWidth()
        ctxFlexContainer1.Height = ctxFlexContainer1.Root.Layout.GetActualHeight()
        dblMin = Limit(picScroll.Width - ctxFlexContainer1.Width, , 0)
        ctxFlexContainer1.Left = Limit(ctxFlexContainer1.Left, dblMin, 0)
        dblMin = Limit(picScroll.Height - ctxFlexContainer1.Height, , 0)
        ctxFlexContainer1.Top = Limit(ctxFlexContainer1.Top, dblMin, 0)
    End If
End Sub

Public Sub RegisterCancelMode(oCtl As Object)
    If Not m_oCtlCancelMode Is Nothing And Not m_oCtlCancelMode Is oCtl Then
        m_oCtlCancelMode.CancelMode
    End If
    Set m_oCtlCancelMode = oCtl
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not m_oCtlCancelMode Is Nothing Then
        m_oCtlCancelMode.CancelMode
        Set m_oCtlCancelMode = Nothing
    End If
End Sub

Private Function Limit(ByVal Value As Double, Optional Min As Variant, Optional Max As Variant) As Double
    Limit = Value
    If Not IsMissing(Min) And Not IsEmpty(Min) Then
        If Value < CDbl(Min) Then
            Limit = CDbl(Min)
        End If
    End If
    If Not IsMissing(Max) And Not IsEmpty(Max) Then
        If Value > CDbl(Max) Then
            Limit = CDbl(Max)
        End If
    End If
End Function

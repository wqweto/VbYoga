VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5436
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6336
   LinkTopic       =   "Form2"
   ScaleHeight     =   5436
   ScaleWidth      =   6336
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Reset"
      Height          =   348
      Left            =   1428
      TabIndex        =   2
      Top             =   84
      Width           =   1272
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Form"
      Height          =   348
      Left            =   84
      TabIndex        =   1
      Top             =   84
      Width           =   1188
   End
   Begin Project1.ctxFlexContainer ctxFlexContainer1 
      Height          =   3876
      Left            =   840
      TabIndex        =   0
      Top             =   672
      Width           =   4128
      _ExtentX        =   7281
      _ExtentY        =   6837
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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

Private Sub Form_Load()
    With ctxFlexContainer1.Root
        With .AddLabel(CssClass:="section-caption")
            .Control.Caption = "Section A"
        End With
        With .AddContainer(CssClass:="top-section")
            With .AddButton().Layout
                .Width = Array(yogaUnitPercent, 30)
            End With
        End With
        With .AddLabel(CssClass:="section-caption")
            .Control.Caption = "Section B"
        End With
        With .AddContainer(CssClass:="main-section")
            With .AddButton().Layout
                .Width = Array(yogaUnitPercent, 30)
                .MinWidth = 1200
                .MaxHeight = 1200
                .Height = Array(yogaUnitPercent, 100)
            End With
            With .AddButton()
                .Control.Style = ucsBtyButtonTurnGreen
                With .Layout
                    .Width = Array(yogaUnitPercent, 30)
                    .MinWidth = 1200
                    .MaxHeight = 1200
                End With
            End With
            With .AddButton()
                .Control.Style = ucsBtyButtonTurnRed
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
            With .AddButton().Layout
                .Width = Array(yogaUnitPercent, 30)
                .FlexShrink = 1
            End With
        End With
    End With
    Set ctxFlexContainer1.Styles = pvGetStyles()
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
    If WindowState <> vbMinimized And ScaleHeight > ctxFlexContainer1.Top Then
        ctxFlexContainer1.Move 0, ctxFlexContainer1.Top, ScaleWidth, ScaleHeight - ctxFlexContainer1.Top
    End If
End Sub

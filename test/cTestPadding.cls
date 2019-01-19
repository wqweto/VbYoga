VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTestPadding"
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

Public Sub Test_padding_no_size()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.PaddingLeft = 10
    root.PaddingTop = 10
    root.PaddingRight = 10
    root.PaddingBottom = 10
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 20!, root.LayoutWidth
    Assert.AreEqual 20!, root.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 20!, root.LayoutWidth
    Assert.AreEqual 20!, root.LayoutHeight
End Sub

Public Sub Test_padding_container_match_child()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.PaddingLeft = 10
    root.PaddingTop = 10
    root.PaddingRight = 10
    root.PaddingBottom = 10

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 10
    root_child0.Height = 10
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 30!, root.LayoutWidth
    Assert.AreEqual 30!, root.LayoutHeight

    Assert.AreEqual 10!, root_child0.LayoutX
    Assert.AreEqual 10!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 30!, root.LayoutWidth
    Assert.AreEqual 30!, root.LayoutHeight

    Assert.AreEqual 10!, root_child0.LayoutX
    Assert.AreEqual 10!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight
End Sub

Public Sub Test_padding_flex_child()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.PaddingLeft = 10
    root.PaddingTop = 10
    root.PaddingRight = 10
    root.PaddingBottom = 10
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.Width = 10
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 10!, root_child0.LayoutX
    Assert.AreEqual 10!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 80!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 80!, root_child0.LayoutX
    Assert.AreEqual 10!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 80!, root_child0.LayoutHeight
End Sub

Public Sub Test_padding_stretch_child()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.PaddingLeft = 10
    root.PaddingTop = 10
    root.PaddingRight = 10
    root.PaddingBottom = 10
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Height = 10
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 10!, root_child0.LayoutX
    Assert.AreEqual 10!, root_child0.LayoutY
    Assert.AreEqual 80!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 10!, root_child0.LayoutX
    Assert.AreEqual 10!, root_child0.LayoutY
    Assert.AreEqual 80!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight
End Sub

Public Sub Test_padding_center_child()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.JustifyContent = YogaJustify.yogaJustCenter
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.PaddingStart = 10
    root.PaddingEnd = 20
    root.PaddingBottom = 20
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 10
    root_child0.Height = 10
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 40!, root_child0.LayoutX
    Assert.AreEqual 35!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 50!, root_child0.LayoutX
    Assert.AreEqual 35!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight
End Sub

Public Sub Test_child_with_padding_align_end()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.JustifyContent = YogaJustify.yogaJustFlexEnd
    root.AlignItems = YogaAlign.yogaAlignFlexEnd
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.PaddingLeft = 20
    root_child0.PaddingTop = 20
    root_child0.PaddingRight = 20
    root_child0.PaddingBottom = 20
    root_child0.Width = 100
    root_child0.Height = 100
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 100!, root_child0.LayoutX
    Assert.AreEqual 100!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 100!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight
End Sub


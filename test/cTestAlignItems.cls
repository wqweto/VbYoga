VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTestAlignItems"
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
Implements ITestFixture

Public Sub Test_align_items_stretch()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Height = 10
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight
End Sub

Public Sub Test_align_items_center()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Width = 10
    root_child0.Height = 10
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 45!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 45!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight
End Sub

Public Sub Test_align_items_flex_start()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.AlignItems = YogaAlign.yogaAlignFlexStart
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Width = 10
    root_child0.Height = 10
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 90!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight
End Sub

Public Sub Test_align_items_flex_end()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.AlignItems = YogaAlign.yogaAlignFlexEnd
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Width = 10
    root_child0.Height = 10
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 90!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight
End Sub

Public Sub Test_align_baseline()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.FlexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignItems = YogaAlign.yogaAlignBaseline
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.Width = 50
    root_child1.Height = 20
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 30!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 50!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 30!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight
End Sub

Public Sub Test_align_baseline_child()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.FlexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignItems = YogaAlign.yogaAlignBaseline
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.Width = 50
    root_child1.Height = 20
    root.Insert 1, root_child1

    Dim root_child1_child0 As cYogaNode: Set root_child1_child0 = YogaNodeNew(Config)
    root_child1_child0.Width = 50
    root_child1_child0.Height = 10
    root_child1.Insert 0, root_child1_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 40!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 50!, root_child1_child0.LayoutWidth
    Assert.AreEqual 10!, root_child1_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 50!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 40!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 50!, root_child1_child0.LayoutWidth
    Assert.AreEqual 10!, root_child1_child0.LayoutHeight
End Sub

Public Sub Test_align_baseline_child_multiline()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.FlexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignItems = YogaAlign.yogaAlignBaseline
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Width = 50
    root_child0.Height = 60
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.FlexDirection = YogaFlexDirection.yogaFlexRow
    root_child1.Wrap = YogaWrap.yogaWrapWrap
    root_child1.Width = 50
    root_child1.Height = 25
    root.Insert 1, root_child1

    Dim root_child1_child0 As cYogaNode: Set root_child1_child0 = YogaNodeNew(Config)
    root_child1_child0.Width = 25
    root_child1_child0.Height = 20
    root_child1.Insert 0, root_child1_child0

    Dim root_child1_child1 As cYogaNode: Set root_child1_child1 = YogaNodeNew(Config)
    root_child1_child1.Width = 25
    root_child1_child1.Height = 10
    root_child1.Insert 1, root_child1_child1

    Dim root_child1_child2 As cYogaNode: Set root_child1_child2 = YogaNodeNew(Config)
    root_child1_child2.Width = 25
    root_child1_child2.Height = 20
    root_child1.Insert 2, root_child1_child2

    Dim root_child1_child3 As cYogaNode: Set root_child1_child3 = YogaNodeNew(Config)
    root_child1_child3.Width = 25
    root_child1_child3.Height = 10
    root_child1.Insert 3, root_child1_child3
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 60!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 40!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 25!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 25!, root_child1_child0.LayoutWidth
    Assert.AreEqual 20!, root_child1_child0.LayoutHeight

    Assert.AreEqual 25!, root_child1_child1.LayoutX
    Assert.AreEqual 0!, root_child1_child1.LayoutY
    Assert.AreEqual 25!, root_child1_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child2.LayoutX
    Assert.AreEqual 20!, root_child1_child2.LayoutY
    Assert.AreEqual 25!, root_child1_child2.LayoutWidth
    Assert.AreEqual 20!, root_child1_child2.LayoutHeight

    Assert.AreEqual 25!, root_child1_child3.LayoutX
    Assert.AreEqual 20!, root_child1_child3.LayoutY
    Assert.AreEqual 25!, root_child1_child3.LayoutWidth
    Assert.AreEqual 10!, root_child1_child3.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 50!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 60!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 40!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 25!, root_child1.LayoutHeight

    Assert.AreEqual 25!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 25!, root_child1_child0.LayoutWidth
    Assert.AreEqual 20!, root_child1_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1_child1.LayoutX
    Assert.AreEqual 0!, root_child1_child1.LayoutY
    Assert.AreEqual 25!, root_child1_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1_child1.LayoutHeight

    Assert.AreEqual 25!, root_child1_child2.LayoutX
    Assert.AreEqual 20!, root_child1_child2.LayoutY
    Assert.AreEqual 25!, root_child1_child2.LayoutWidth
    Assert.AreEqual 20!, root_child1_child2.LayoutHeight

    Assert.AreEqual 0!, root_child1_child3.LayoutX
    Assert.AreEqual 20!, root_child1_child3.LayoutY
    Assert.AreEqual 25!, root_child1_child3.LayoutWidth
    Assert.AreEqual 10!, root_child1_child3.LayoutHeight
End Sub

Public Sub Test_align_baseline_child_multiline_override()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.FlexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignItems = YogaAlign.yogaAlignBaseline
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Width = 50
    root_child0.Height = 60
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.FlexDirection = YogaFlexDirection.yogaFlexRow
    root_child1.Wrap = YogaWrap.yogaWrapWrap
    root_child1.Width = 50
    root_child1.Height = 25
    root.Insert 1, root_child1

    Dim root_child1_child0 As cYogaNode: Set root_child1_child0 = YogaNodeNew(Config)
    root_child1_child0.Width = 25
    root_child1_child0.Height = 20
    root_child1.Insert 0, root_child1_child0

    Dim root_child1_child1 As cYogaNode: Set root_child1_child1 = YogaNodeNew(Config)
    root_child1_child1.AlignSelf = YogaAlign.yogaAlignBaseline
    root_child1_child1.Width = 25
    root_child1_child1.Height = 10
    root_child1.Insert 1, root_child1_child1

    Dim root_child1_child2 As cYogaNode: Set root_child1_child2 = YogaNodeNew(Config)
    root_child1_child2.Width = 25
    root_child1_child2.Height = 20
    root_child1.Insert 2, root_child1_child2

    Dim root_child1_child3 As cYogaNode: Set root_child1_child3 = YogaNodeNew(Config)
    root_child1_child3.AlignSelf = YogaAlign.yogaAlignBaseline
    root_child1_child3.Width = 25
    root_child1_child3.Height = 10
    root_child1.Insert 3, root_child1_child3
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 60!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 25!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 25!, root_child1_child0.LayoutWidth
    Assert.AreEqual 20!, root_child1_child0.LayoutHeight

    Assert.AreEqual 25!, root_child1_child1.LayoutX
    Assert.AreEqual 0!, root_child1_child1.LayoutY
    Assert.AreEqual 25!, root_child1_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child2.LayoutX
    Assert.AreEqual 20!, root_child1_child2.LayoutY
    Assert.AreEqual 25!, root_child1_child2.LayoutWidth
    Assert.AreEqual 20!, root_child1_child2.LayoutHeight

    Assert.AreEqual 25!, root_child1_child3.LayoutX
    Assert.AreEqual 20!, root_child1_child3.LayoutY
    Assert.AreEqual 25!, root_child1_child3.LayoutWidth
    Assert.AreEqual 10!, root_child1_child3.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 50!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 60!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 25!, root_child1.LayoutHeight

    Assert.AreEqual 25!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 25!, root_child1_child0.LayoutWidth
    Assert.AreEqual 20!, root_child1_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1_child1.LayoutX
    Assert.AreEqual 0!, root_child1_child1.LayoutY
    Assert.AreEqual 25!, root_child1_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1_child1.LayoutHeight

    Assert.AreEqual 25!, root_child1_child2.LayoutX
    Assert.AreEqual 20!, root_child1_child2.LayoutY
    Assert.AreEqual 25!, root_child1_child2.LayoutWidth
    Assert.AreEqual 20!, root_child1_child2.LayoutHeight

    Assert.AreEqual 0!, root_child1_child3.LayoutX
    Assert.AreEqual 20!, root_child1_child3.LayoutY
    Assert.AreEqual 25!, root_child1_child3.LayoutWidth
    Assert.AreEqual 10!, root_child1_child3.LayoutHeight
End Sub

Public Sub Test_align_baseline_child_multiline_no_override_on_secondline()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.FlexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignItems = YogaAlign.yogaAlignBaseline
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Width = 50
    root_child0.Height = 60
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.FlexDirection = YogaFlexDirection.yogaFlexRow
    root_child1.Wrap = YogaWrap.yogaWrapWrap
    root_child1.Width = 50
    root_child1.Height = 25
    root.Insert 1, root_child1

    Dim root_child1_child0 As cYogaNode: Set root_child1_child0 = YogaNodeNew(Config)
    root_child1_child0.Width = 25
    root_child1_child0.Height = 20
    root_child1.Insert 0, root_child1_child0

    Dim root_child1_child1 As cYogaNode: Set root_child1_child1 = YogaNodeNew(Config)
    root_child1_child1.Width = 25
    root_child1_child1.Height = 10
    root_child1.Insert 1, root_child1_child1

    Dim root_child1_child2 As cYogaNode: Set root_child1_child2 = YogaNodeNew(Config)
    root_child1_child2.Width = 25
    root_child1_child2.Height = 20
    root_child1.Insert 2, root_child1_child2

    Dim root_child1_child3 As cYogaNode: Set root_child1_child3 = YogaNodeNew(Config)
    root_child1_child3.AlignSelf = YogaAlign.yogaAlignBaseline
    root_child1_child3.Width = 25
    root_child1_child3.Height = 10
    root_child1.Insert 3, root_child1_child3
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 60!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 40!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 25!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 25!, root_child1_child0.LayoutWidth
    Assert.AreEqual 20!, root_child1_child0.LayoutHeight

    Assert.AreEqual 25!, root_child1_child1.LayoutX
    Assert.AreEqual 0!, root_child1_child1.LayoutY
    Assert.AreEqual 25!, root_child1_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child2.LayoutX
    Assert.AreEqual 20!, root_child1_child2.LayoutY
    Assert.AreEqual 25!, root_child1_child2.LayoutWidth
    Assert.AreEqual 20!, root_child1_child2.LayoutHeight

    Assert.AreEqual 25!, root_child1_child3.LayoutX
    Assert.AreEqual 20!, root_child1_child3.LayoutY
    Assert.AreEqual 25!, root_child1_child3.LayoutWidth
    Assert.AreEqual 10!, root_child1_child3.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 50!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 60!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 40!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 25!, root_child1.LayoutHeight

    Assert.AreEqual 25!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 25!, root_child1_child0.LayoutWidth
    Assert.AreEqual 20!, root_child1_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1_child1.LayoutX
    Assert.AreEqual 0!, root_child1_child1.LayoutY
    Assert.AreEqual 25!, root_child1_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1_child1.LayoutHeight

    Assert.AreEqual 25!, root_child1_child2.LayoutX
    Assert.AreEqual 20!, root_child1_child2.LayoutY
    Assert.AreEqual 25!, root_child1_child2.LayoutWidth
    Assert.AreEqual 20!, root_child1_child2.LayoutHeight

    Assert.AreEqual 0!, root_child1_child3.LayoutX
    Assert.AreEqual 20!, root_child1_child3.LayoutY
    Assert.AreEqual 25!, root_child1_child3.LayoutWidth
    Assert.AreEqual 10!, root_child1_child3.LayoutHeight
End Sub

Public Sub Test_align_baseline_child_top()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.FlexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignItems = YogaAlign.yogaAlignBaseline
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Top = 10
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.Width = 50
    root_child1.Height = 20
    root.Insert 1, root_child1

    Dim root_child1_child0 As cYogaNode: Set root_child1_child0 = YogaNodeNew(Config)
    root_child1_child0.Width = 50
    root_child1_child0.Height = 10
    root_child1.Insert 0, root_child1_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 10!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 40!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 50!, root_child1_child0.LayoutWidth
    Assert.AreEqual 10!, root_child1_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 50!, root_child0.LayoutX
    Assert.AreEqual 10!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 40!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 50!, root_child1_child0.LayoutWidth
    Assert.AreEqual 10!, root_child1_child0.LayoutHeight
End Sub

Public Sub Test_align_baseline_child_top2()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.FlexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignItems = YogaAlign.yogaAlignBaseline
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.Top = 5
    root_child1.Width = 50
    root_child1.Height = 20
    root.Insert 1, root_child1

    Dim root_child1_child0 As cYogaNode: Set root_child1_child0 = YogaNodeNew(Config)
    root_child1_child0.Width = 50
    root_child1_child0.Height = 10
    root_child1.Insert 0, root_child1_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 45!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 50!, root_child1_child0.LayoutWidth
    Assert.AreEqual 10!, root_child1_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 50!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 45!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 50!, root_child1_child0.LayoutWidth
    Assert.AreEqual 10!, root_child1_child0.LayoutHeight
End Sub

Public Sub Test_align_baseline_double_nested_child()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.FlexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignItems = YogaAlign.yogaAlignBaseline
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(Config)
    root_child0_child0.Width = 50
    root_child0_child0.Height = 20
    root_child0.Insert 0, root_child0_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.Width = 50
    root_child1.Height = 20
    root.Insert 1, root_child1

    Dim root_child1_child0 As cYogaNode: Set root_child1_child0 = YogaNodeNew(Config)
    root_child1_child0.Width = 50
    root_child1_child0.Height = 15
    root_child1.Insert 0, root_child1_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 50!, root_child0_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 5!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 50!, root_child1_child0.LayoutWidth
    Assert.AreEqual 15!, root_child1_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 50!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 50!, root_child0_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 5!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 50!, root_child1_child0.LayoutWidth
    Assert.AreEqual 15!, root_child1_child0.LayoutHeight
End Sub

Public Sub Test_align_baseline_column()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.AlignItems = YogaAlign.yogaAlignBaseline
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.Width = 50
    root_child1.Height = 20
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 50!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight
End Sub

Public Sub Test_align_baseline_child_margin()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.FlexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignItems = YogaAlign.yogaAlignBaseline
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.MarginLeft = 5
    root_child0.MarginTop = 5
    root_child0.MarginRight = 5
    root_child0.MarginBottom = 5
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.Width = 50
    root_child1.Height = 20
    root.Insert 1, root_child1

    Dim root_child1_child0 As cYogaNode: Set root_child1_child0 = YogaNodeNew(Config)
    root_child1_child0.MarginLeft = 1
    root_child1_child0.MarginTop = 1
    root_child1_child0.MarginRight = 1
    root_child1_child0.MarginBottom = 1
    root_child1_child0.Width = 50
    root_child1_child0.Height = 10
    root_child1.Insert 0, root_child1_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 5!, root_child0.LayoutX
    Assert.AreEqual 5!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 60!, root_child1.LayoutX
    Assert.AreEqual 44!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 1!, root_child1_child0.LayoutX
    Assert.AreEqual 1!, root_child1_child0.LayoutY
    Assert.AreEqual 50!, root_child1_child0.LayoutWidth
    Assert.AreEqual 10!, root_child1_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 45!, root_child0.LayoutX
    Assert.AreEqual 5!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual -10!, root_child1.LayoutX
    Assert.AreEqual 44!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual -1!, root_child1_child0.LayoutX
    Assert.AreEqual 1!, root_child1_child0.LayoutY
    Assert.AreEqual 50!, root_child1_child0.LayoutWidth
    Assert.AreEqual 10!, root_child1_child0.LayoutHeight
End Sub

Public Sub Test_align_baseline_child_padding()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.FlexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignItems = YogaAlign.yogaAlignBaseline
    root.PaddingLeft = 5
    root.PaddingTop = 5
    root.PaddingRight = 5
    root.PaddingBottom = 5
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.PaddingLeft = 5
    root_child1.PaddingTop = 5
    root_child1.PaddingRight = 5
    root_child1.PaddingBottom = 5
    root_child1.Width = 50
    root_child1.Height = 20
    root.Insert 1, root_child1

    Dim root_child1_child0 As cYogaNode: Set root_child1_child0 = YogaNodeNew(Config)
    root_child1_child0.Width = 50
    root_child1_child0.Height = 10
    root_child1.Insert 0, root_child1_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 5!, root_child0.LayoutX
    Assert.AreEqual 5!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 55!, root_child1.LayoutX
    Assert.AreEqual 40!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 5!, root_child1_child0.LayoutX
    Assert.AreEqual 5!, root_child1_child0.LayoutY
    Assert.AreEqual 50!, root_child1_child0.LayoutWidth
    Assert.AreEqual 10!, root_child1_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 45!, root_child0.LayoutX
    Assert.AreEqual 5!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual -5!, root_child1.LayoutX
    Assert.AreEqual 40!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual -5!, root_child1_child0.LayoutX
    Assert.AreEqual 5!, root_child1_child0.LayoutY
    Assert.AreEqual 50!, root_child1_child0.LayoutWidth
    Assert.AreEqual 10!, root_child1_child0.LayoutHeight
End Sub

Public Sub Test_align_baseline_multiline()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.FlexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignItems = YogaAlign.yogaAlignBaseline
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.Width = 50
    root_child1.Height = 20
    root.Insert 1, root_child1

    Dim root_child1_child0 As cYogaNode: Set root_child1_child0 = YogaNodeNew(Config)
    root_child1_child0.Width = 50
    root_child1_child0.Height = 10
    root_child1.Insert 0, root_child1_child0

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(Config)
    root_child2.Width = 50
    root_child2.Height = 20
    root.Insert 2, root_child2

    Dim root_child2_child0 As cYogaNode: Set root_child2_child0 = YogaNodeNew(Config)
    root_child2_child0.Width = 50
    root_child2_child0.Height = 10
    root_child2.Insert 0, root_child2_child0

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(Config)
    root_child3.Width = 50
    root_child3.Height = 50
    root.Insert 3, root_child3
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 40!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 50!, root_child1_child0.LayoutWidth
    Assert.AreEqual 10!, root_child1_child0.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 100!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 20!, root_child2.LayoutHeight

    Assert.AreEqual 0!, root_child2_child0.LayoutX
    Assert.AreEqual 0!, root_child2_child0.LayoutY
    Assert.AreEqual 50!, root_child2_child0.LayoutWidth
    Assert.AreEqual 10!, root_child2_child0.LayoutHeight

    Assert.AreEqual 50!, root_child3.LayoutX
    Assert.AreEqual 60!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 50!, root_child3.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 50!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 40!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 50!, root_child1_child0.LayoutWidth
    Assert.AreEqual 10!, root_child1_child0.LayoutHeight

    Assert.AreEqual 50!, root_child2.LayoutX
    Assert.AreEqual 100!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 20!, root_child2.LayoutHeight

    Assert.AreEqual 0!, root_child2_child0.LayoutX
    Assert.AreEqual 0!, root_child2_child0.LayoutY
    Assert.AreEqual 50!, root_child2_child0.LayoutWidth
    Assert.AreEqual 10!, root_child2_child0.LayoutHeight

    Assert.AreEqual 0!, root_child3.LayoutX
    Assert.AreEqual 60!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 50!, root_child3.LayoutHeight
End Sub

Public Sub Test_align_baseline_multiline_column()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.AlignItems = YogaAlign.yogaAlignBaseline
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.Width = 30
    root_child1.Height = 50
    root.Insert 1, root_child1

    Dim root_child1_child0 As cYogaNode: Set root_child1_child0 = YogaNodeNew(Config)
    root_child1_child0.Width = 20
    root_child1_child0.Height = 20
    root_child1.Insert 0, root_child1_child0

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(Config)
    root_child2.Width = 40
    root_child2.Height = 70
    root.Insert 2, root_child2

    Dim root_child2_child0 As cYogaNode: Set root_child2_child0 = YogaNodeNew(Config)
    root_child2_child0.Width = 10
    root_child2_child0.Height = 10
    root_child2.Insert 0, root_child2_child0

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(Config)
    root_child3.Width = 50
    root_child3.Height = 20
    root.Insert 3, root_child3
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 30!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 20!, root_child1_child0.LayoutWidth
    Assert.AreEqual 20!, root_child1_child0.LayoutHeight

    Assert.AreEqual 50!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 40!, root_child2.LayoutWidth
    Assert.AreEqual 70!, root_child2.LayoutHeight

    Assert.AreEqual 0!, root_child2_child0.LayoutX
    Assert.AreEqual 0!, root_child2_child0.LayoutY
    Assert.AreEqual 10!, root_child2_child0.LayoutWidth
    Assert.AreEqual 10!, root_child2_child0.LayoutHeight

    Assert.AreEqual 50!, root_child3.LayoutX
    Assert.AreEqual 70!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 20!, root_child3.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 50!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 70!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 30!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    Assert.AreEqual 10!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 20!, root_child1_child0.LayoutWidth
    Assert.AreEqual 20!, root_child1_child0.LayoutHeight

    Assert.AreEqual 10!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 40!, root_child2.LayoutWidth
    Assert.AreEqual 70!, root_child2.LayoutHeight

    Assert.AreEqual 30!, root_child2_child0.LayoutX
    Assert.AreEqual 0!, root_child2_child0.LayoutY
    Assert.AreEqual 10!, root_child2_child0.LayoutWidth
    Assert.AreEqual 10!, root_child2_child0.LayoutHeight

    Assert.AreEqual 0!, root_child3.LayoutX
    Assert.AreEqual 70!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 20!, root_child3.LayoutHeight
End Sub

Public Sub Test_align_baseline_multiline_column2()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.AlignItems = YogaAlign.yogaAlignBaseline
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.Width = 30
    root_child1.Height = 50
    root.Insert 1, root_child1

    Dim root_child1_child0 As cYogaNode: Set root_child1_child0 = YogaNodeNew(Config)
    root_child1_child0.Width = 20
    root_child1_child0.Height = 20
    root_child1.Insert 0, root_child1_child0

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(Config)
    root_child2.Width = 40
    root_child2.Height = 70
    root.Insert 2, root_child2

    Dim root_child2_child0 As cYogaNode: Set root_child2_child0 = YogaNodeNew(Config)
    root_child2_child0.Width = 10
    root_child2_child0.Height = 10
    root_child2.Insert 0, root_child2_child0

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(Config)
    root_child3.Width = 50
    root_child3.Height = 20
    root.Insert 3, root_child3
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 30!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 20!, root_child1_child0.LayoutWidth
    Assert.AreEqual 20!, root_child1_child0.LayoutHeight

    Assert.AreEqual 50!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 40!, root_child2.LayoutWidth
    Assert.AreEqual 70!, root_child2.LayoutHeight

    Assert.AreEqual 0!, root_child2_child0.LayoutX
    Assert.AreEqual 0!, root_child2_child0.LayoutY
    Assert.AreEqual 10!, root_child2_child0.LayoutWidth
    Assert.AreEqual 10!, root_child2_child0.LayoutHeight

    Assert.AreEqual 50!, root_child3.LayoutX
    Assert.AreEqual 70!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 20!, root_child3.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 50!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 70!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 30!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    Assert.AreEqual 10!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 20!, root_child1_child0.LayoutWidth
    Assert.AreEqual 20!, root_child1_child0.LayoutHeight

    Assert.AreEqual 10!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 40!, root_child2.LayoutWidth
    Assert.AreEqual 70!, root_child2.LayoutHeight

    Assert.AreEqual 30!, root_child2_child0.LayoutX
    Assert.AreEqual 0!, root_child2_child0.LayoutY
    Assert.AreEqual 10!, root_child2_child0.LayoutWidth
    Assert.AreEqual 10!, root_child2_child0.LayoutHeight

    Assert.AreEqual 0!, root_child3.LayoutX
    Assert.AreEqual 70!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 20!, root_child3.LayoutHeight
End Sub

Public Sub Test_align_baseline_multiline_row_and_column()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.FlexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignItems = YogaAlign.yogaAlignBaseline
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.Width = 50
    root_child1.Height = 50
    root.Insert 1, root_child1

    Dim root_child1_child0 As cYogaNode: Set root_child1_child0 = YogaNodeNew(Config)
    root_child1_child0.Width = 50
    root_child1_child0.Height = 10
    root_child1.Insert 0, root_child1_child0

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(Config)
    root_child2.Width = 50
    root_child2.Height = 20
    root.Insert 2, root_child2

    Dim root_child2_child0 As cYogaNode: Set root_child2_child0 = YogaNodeNew(Config)
    root_child2_child0.Width = 50
    root_child2_child0.Height = 10
    root_child2.Insert 0, root_child2_child0

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(Config)
    root_child3.Width = 50
    root_child3.Height = 20
    root.Insert 3, root_child3
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 40!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 50!, root_child1_child0.LayoutWidth
    Assert.AreEqual 10!, root_child1_child0.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 100!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 20!, root_child2.LayoutHeight

    Assert.AreEqual 0!, root_child2_child0.LayoutX
    Assert.AreEqual 0!, root_child2_child0.LayoutY
    Assert.AreEqual 50!, root_child2_child0.LayoutWidth
    Assert.AreEqual 10!, root_child2_child0.LayoutHeight

    Assert.AreEqual 50!, root_child3.LayoutX
    Assert.AreEqual 90!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 20!, root_child3.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 50!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 40!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 50!, root_child1_child0.LayoutWidth
    Assert.AreEqual 10!, root_child1_child0.LayoutHeight

    Assert.AreEqual 50!, root_child2.LayoutX
    Assert.AreEqual 100!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 20!, root_child2.LayoutHeight

    Assert.AreEqual 0!, root_child2_child0.LayoutX
    Assert.AreEqual 0!, root_child2_child0.LayoutY
    Assert.AreEqual 50!, root_child2_child0.LayoutWidth
    Assert.AreEqual 10!, root_child2_child0.LayoutHeight

    Assert.AreEqual 0!, root_child3.LayoutX
    Assert.AreEqual 90!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 20!, root_child3.LayoutHeight
End Sub

Public Sub Test_align_items_center_child_with_margin_bigger_than_parent()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.JustifyContent = YogaJustify.yogaJustCenter
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.Width = 52
    root.Height = 52

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.AlignItems = YogaAlign.yogaAlignCenter
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(Config)
    root_child0_child0.MarginLeft = 10
    root_child0_child0.MarginRight = 10
    root_child0_child0.Width = 52
    root_child0_child0.Height = 52
    root_child0.Insert 0, root_child0_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 52!, root.LayoutWidth
    Assert.AreEqual 52!, root.LayoutHeight

    Assert.AreEqual -10!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 72!, root_child0.LayoutWidth
    Assert.AreEqual 52!, root_child0.LayoutHeight

    Assert.AreEqual 10!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 52!, root_child0_child0.LayoutWidth
    Assert.AreEqual 52!, root_child0_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 52!, root.LayoutWidth
    Assert.AreEqual 52!, root.LayoutHeight

    Assert.AreEqual -10!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 72!, root_child0.LayoutWidth
    Assert.AreEqual 52!, root_child0.LayoutHeight

    Assert.AreEqual 10!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 52!, root_child0_child0.LayoutWidth
    Assert.AreEqual 52!, root_child0_child0.LayoutHeight
End Sub

Public Sub Test_align_items_flex_end_child_with_margin_bigger_than_parent()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.JustifyContent = YogaJustify.yogaJustCenter
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.Width = 52
    root.Height = 52

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.AlignItems = YogaAlign.yogaAlignFlexEnd
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(Config)
    root_child0_child0.MarginLeft = 10
    root_child0_child0.MarginRight = 10
    root_child0_child0.Width = 52
    root_child0_child0.Height = 52
    root_child0.Insert 0, root_child0_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 52!, root.LayoutWidth
    Assert.AreEqual 52!, root.LayoutHeight

    Assert.AreEqual -10!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 72!, root_child0.LayoutWidth
    Assert.AreEqual 52!, root_child0.LayoutHeight

    Assert.AreEqual 10!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 52!, root_child0_child0.LayoutWidth
    Assert.AreEqual 52!, root_child0_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 52!, root.LayoutWidth
    Assert.AreEqual 52!, root.LayoutHeight

    Assert.AreEqual -10!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 72!, root_child0.LayoutWidth
    Assert.AreEqual 52!, root_child0.LayoutHeight

    Assert.AreEqual 10!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 52!, root_child0_child0.LayoutWidth
    Assert.AreEqual 52!, root_child0_child0.LayoutHeight
End Sub

Public Sub Test_align_items_center_child_without_margin_bigger_than_parent()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.JustifyContent = YogaJustify.yogaJustCenter
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.Width = 52
    root.Height = 52

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.AlignItems = YogaAlign.yogaAlignCenter
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(Config)
    root_child0_child0.Width = 72
    root_child0_child0.Height = 72
    root_child0.Insert 0, root_child0_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 52!, root.LayoutWidth
    Assert.AreEqual 52!, root.LayoutHeight

    Assert.AreEqual -10!, root_child0.LayoutX
    Assert.AreEqual -10!, root_child0.LayoutY
    Assert.AreEqual 72!, root_child0.LayoutWidth
    Assert.AreEqual 72!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 72!, root_child0_child0.LayoutWidth
    Assert.AreEqual 72!, root_child0_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 52!, root.LayoutWidth
    Assert.AreEqual 52!, root.LayoutHeight

    Assert.AreEqual -10!, root_child0.LayoutX
    Assert.AreEqual -10!, root_child0.LayoutY
    Assert.AreEqual 72!, root_child0.LayoutWidth
    Assert.AreEqual 72!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 72!, root_child0_child0.LayoutWidth
    Assert.AreEqual 72!, root_child0_child0.LayoutHeight
End Sub

Public Sub Test_align_items_flex_end_child_without_margin_bigger_than_parent()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.JustifyContent = YogaJustify.yogaJustCenter
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.Width = 52
    root.Height = 52

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.AlignItems = YogaAlign.yogaAlignFlexEnd
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(Config)
    root_child0_child0.Width = 72
    root_child0_child0.Height = 72
    root_child0.Insert 0, root_child0_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 52!, root.LayoutWidth
    Assert.AreEqual 52!, root.LayoutHeight

    Assert.AreEqual -10!, root_child0.LayoutX
    Assert.AreEqual -10!, root_child0.LayoutY
    Assert.AreEqual 72!, root_child0.LayoutWidth
    Assert.AreEqual 72!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 72!, root_child0_child0.LayoutWidth
    Assert.AreEqual 72!, root_child0_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 52!, root.LayoutWidth
    Assert.AreEqual 52!, root.LayoutHeight

    Assert.AreEqual -10!, root_child0.LayoutX
    Assert.AreEqual -10!, root_child0.LayoutY
    Assert.AreEqual 72!, root_child0.LayoutWidth
    Assert.AreEqual 72!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 72!, root_child0_child0.LayoutWidth
    Assert.AreEqual 72!, root_child0_child0.LayoutHeight
End Sub

Public Sub Test_align_center_should_size_based_on_content()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.MarginTop = 20
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.JustifyContent = YogaJustify.yogaJustCenter
    root_child0.FlexShrink = 1
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(Config)
    root_child0_child0.FlexGrow = 1
    root_child0_child0.FlexShrink = 1
    root_child0.Insert 0, root_child0_child0

    Dim root_child0_child0_child0 As cYogaNode: Set root_child0_child0_child0 = YogaNodeNew(Config)
    root_child0_child0_child0.Width = 20
    root_child0_child0_child0.Height = 20
    root_child0_child0.Insert 0, root_child0_child0_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 20!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 40!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 20!, root_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 20!, root_child0_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutY
    Assert.AreEqual 20!, root_child0_child0_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0_child0_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 20!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 40!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 20!, root_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 20!, root_child0_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutY
    Assert.AreEqual 20!, root_child0_child0_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0_child0_child0.LayoutHeight
End Sub

Public Sub Test_align_strech_should_size_based_on_parent()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.MarginTop = 20
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.JustifyContent = YogaJustify.yogaJustCenter
    root_child0.FlexShrink = 1
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(Config)
    root_child0_child0.FlexGrow = 1
    root_child0_child0.FlexShrink = 1
    root_child0.Insert 0, root_child0_child0

    Dim root_child0_child0_child0 As cYogaNode: Set root_child0_child0_child0 = YogaNodeNew(Config)
    root_child0_child0_child0.Width = 20
    root_child0_child0_child0.Height = 20
    root_child0_child0.Insert 0, root_child0_child0_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 20!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 100!, root_child0_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutY
    Assert.AreEqual 20!, root_child0_child0_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0_child0_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 20!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 100!, root_child0_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0_child0.LayoutHeight

    Assert.AreEqual 80!, root_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutY
    Assert.AreEqual 20!, root_child0_child0_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0_child0_child0.LayoutHeight
End Sub

Public Sub Test_align_flex_start_with_shrinking_children()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.Width = 500
    root.Height = 500

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.AlignItems = YogaAlign.yogaAlignFlexStart
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(Config)
    root_child0_child0.FlexGrow = 1
    root_child0_child0.FlexShrink = 1
    root_child0.Insert 0, root_child0_child0

    Dim root_child0_child0_child0 As cYogaNode: Set root_child0_child0_child0 = YogaNodeNew(Config)
    root_child0_child0_child0.FlexGrow = 1
    root_child0_child0_child0.FlexShrink = 1
    root_child0_child0.Insert 0, root_child0_child0_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 500!, root.LayoutWidth
    Assert.AreEqual 500!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 500!, root_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 0!, root_child0_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutY
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 500!, root.LayoutWidth
    Assert.AreEqual 500!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 500!, root_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0.LayoutHeight

    Assert.AreEqual 500!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 0!, root_child0_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutY
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutHeight
End Sub

Public Sub Test_align_flex_start_with_stretching_children()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.Width = 500
    root.Height = 500

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(Config)
    root_child0_child0.FlexGrow = 1
    root_child0_child0.FlexShrink = 1
    root_child0.Insert 0, root_child0_child0

    Dim root_child0_child0_child0 As cYogaNode: Set root_child0_child0_child0 = YogaNodeNew(Config)
    root_child0_child0_child0.FlexGrow = 1
    root_child0_child0_child0.FlexShrink = 1
    root_child0_child0.Insert 0, root_child0_child0_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 500!, root.LayoutWidth
    Assert.AreEqual 500!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 500!, root_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 500!, root_child0_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutY
    Assert.AreEqual 500!, root_child0_child0_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 500!, root.LayoutWidth
    Assert.AreEqual 500!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 500!, root_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 500!, root_child0_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutY
    Assert.AreEqual 500!, root_child0_child0_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutHeight
End Sub

Public Sub Test_align_flex_start_with_shrinking_children_with_stretch()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.Width = 500
    root.Height = 500

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.AlignItems = YogaAlign.yogaAlignFlexStart
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(Config)
    root_child0_child0.FlexGrow = 1
    root_child0_child0.FlexShrink = 1
    root_child0.Insert 0, root_child0_child0

    Dim root_child0_child0_child0 As cYogaNode: Set root_child0_child0_child0 = YogaNodeNew(Config)
    root_child0_child0_child0.FlexGrow = 1
    root_child0_child0_child0.FlexShrink = 1
    root_child0_child0.Insert 0, root_child0_child0_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 500!, root.LayoutWidth
    Assert.AreEqual 500!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 500!, root_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 0!, root_child0_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutY
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 500!, root.LayoutWidth
    Assert.AreEqual 500!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 500!, root_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0.LayoutHeight

    Assert.AreEqual 500!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 0!, root_child0_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutY
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutHeight
End Sub

Private Sub ITestFixture_GetTestCases(ByVal Tests As SimplyVBComp.TestFixtureBuilder)
    Tests.Add "Test_align_items_stretch"
    Tests.Add "Test_align_items_center"
    Tests.Add "Test_align_items_flex_start"
    Tests.Add "Test_align_items_flex_end"
    Tests.Add "Test_align_baseline"
    Tests.Add "Test_align_baseline_child"
    Tests.Add "Test_align_baseline_child_multiline"
    Tests.Add "Test_align_baseline_child_multiline_override"
    Tests.Add "Test_align_baseline_child_multiline_no_override_on_secondline"
    Tests.Add "Test_align_baseline_child_top"
    Tests.Add "Test_align_baseline_child_top2"
    Tests.Add "Test_align_baseline_double_nested_child"
    Tests.Add "Test_align_baseline_column"
    Tests.Add "Test_align_baseline_child_margin"
    Tests.Add "Test_align_baseline_child_padding"
    Tests.Add "Test_align_baseline_multiline"
    Tests.Add "Test_align_baseline_multiline_column"
    Tests.Add "Test_align_baseline_multiline_column2"
    Tests.Add "Test_align_baseline_multiline_row_and_column"
    Tests.Add "Test_align_items_center_child_with_margin_bigger_than_parent"
    Tests.Add "Test_align_items_flex_end_child_with_margin_bigger_than_parent"
    Tests.Add "Test_align_items_center_child_without_margin_bigger_than_parent"
    Tests.Add "Test_align_items_flex_end_child_without_margin_bigger_than_parent"
    Tests.Add "Test_align_center_should_size_based_on_content"
    Tests.Add "Test_align_strech_should_size_based_on_parent"
    Tests.Add "Test_align_flex_start_with_shrinking_children"
    Tests.Add "Test_align_flex_start_with_stretching_children"
    Tests.Add "Test_align_flex_start_with_shrinking_children_with_stretch"
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTestFlexWrap"
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

Public Sub Test_wrap_column()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Width = 30
    root_child0.Height = 30
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.Width = 30
    root_child1.Height = 30
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(Config)
    root_child2.Width = 30
    root_child2.Height = 30
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(Config)
    root_child3.Width = 30
    root_child3.Height = 30
    root.Insert 3, root_child3
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 60!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 30!, root_child0.LayoutWidth
    Assert.AreEqual 30!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 30!, root_child1.LayoutY
    Assert.AreEqual 30!, root_child1.LayoutWidth
    Assert.AreEqual 30!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 60!, root_child2.LayoutY
    Assert.AreEqual 30!, root_child2.LayoutWidth
    Assert.AreEqual 30!, root_child2.LayoutHeight

    Assert.AreEqual 30!, root_child3.LayoutX
    Assert.AreEqual 0!, root_child3.LayoutY
    Assert.AreEqual 30!, root_child3.LayoutWidth
    Assert.AreEqual 30!, root_child3.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 60!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 30!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 30!, root_child0.LayoutWidth
    Assert.AreEqual 30!, root_child0.LayoutHeight

    Assert.AreEqual 30!, root_child1.LayoutX
    Assert.AreEqual 30!, root_child1.LayoutY
    Assert.AreEqual 30!, root_child1.LayoutWidth
    Assert.AreEqual 30!, root_child1.LayoutHeight

    Assert.AreEqual 30!, root_child2.LayoutX
    Assert.AreEqual 60!, root_child2.LayoutY
    Assert.AreEqual 30!, root_child2.LayoutWidth
    Assert.AreEqual 30!, root_child2.LayoutHeight

    Assert.AreEqual 0!, root_child3.LayoutX
    Assert.AreEqual 0!, root_child3.LayoutY
    Assert.AreEqual 30!, root_child3.LayoutWidth
    Assert.AreEqual 30!, root_child3.LayoutHeight
End Sub

Public Sub Test_wrap_row()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.FlexDirection = YogaFlexDirection.yogaFlexRow
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Width = 30
    root_child0.Height = 30
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.Width = 30
    root_child1.Height = 30
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(Config)
    root_child2.Width = 30
    root_child2.Height = 30
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(Config)
    root_child3.Width = 30
    root_child3.Height = 30
    root.Insert 3, root_child3
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 60!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 30!, root_child0.LayoutWidth
    Assert.AreEqual 30!, root_child0.LayoutHeight

    Assert.AreEqual 30!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 30!, root_child1.LayoutWidth
    Assert.AreEqual 30!, root_child1.LayoutHeight

    Assert.AreEqual 60!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 30!, root_child2.LayoutWidth
    Assert.AreEqual 30!, root_child2.LayoutHeight

    Assert.AreEqual 0!, root_child3.LayoutX
    Assert.AreEqual 30!, root_child3.LayoutY
    Assert.AreEqual 30!, root_child3.LayoutWidth
    Assert.AreEqual 30!, root_child3.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 60!, root.LayoutHeight

    Assert.AreEqual 70!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 30!, root_child0.LayoutWidth
    Assert.AreEqual 30!, root_child0.LayoutHeight

    Assert.AreEqual 40!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 30!, root_child1.LayoutWidth
    Assert.AreEqual 30!, root_child1.LayoutHeight

    Assert.AreEqual 10!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 30!, root_child2.LayoutWidth
    Assert.AreEqual 30!, root_child2.LayoutHeight

    Assert.AreEqual 70!, root_child3.LayoutX
    Assert.AreEqual 30!, root_child3.LayoutY
    Assert.AreEqual 30!, root_child3.LayoutWidth
    Assert.AreEqual 30!, root_child3.LayoutHeight
End Sub

Public Sub Test_wrap_row_align_items_flex_end()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.FlexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignItems = YogaAlign.yogaAlignFlexEnd
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Width = 30
    root_child0.Height = 10
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.Width = 30
    root_child1.Height = 20
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(Config)
    root_child2.Width = 30
    root_child2.Height = 30
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(Config)
    root_child3.Width = 30
    root_child3.Height = 30
    root.Insert 3, root_child3
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 60!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 20!, root_child0.LayoutY
    Assert.AreEqual 30!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 30!, root_child1.LayoutX
    Assert.AreEqual 10!, root_child1.LayoutY
    Assert.AreEqual 30!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 60!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 30!, root_child2.LayoutWidth
    Assert.AreEqual 30!, root_child2.LayoutHeight

    Assert.AreEqual 0!, root_child3.LayoutX
    Assert.AreEqual 30!, root_child3.LayoutY
    Assert.AreEqual 30!, root_child3.LayoutWidth
    Assert.AreEqual 30!, root_child3.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 60!, root.LayoutHeight

    Assert.AreEqual 70!, root_child0.LayoutX
    Assert.AreEqual 20!, root_child0.LayoutY
    Assert.AreEqual 30!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 40!, root_child1.LayoutX
    Assert.AreEqual 10!, root_child1.LayoutY
    Assert.AreEqual 30!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 10!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 30!, root_child2.LayoutWidth
    Assert.AreEqual 30!, root_child2.LayoutHeight

    Assert.AreEqual 70!, root_child3.LayoutX
    Assert.AreEqual 30!, root_child3.LayoutY
    Assert.AreEqual 30!, root_child3.LayoutWidth
    Assert.AreEqual 30!, root_child3.LayoutHeight
End Sub

Public Sub Test_wrap_row_align_items_center()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.FlexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Width = 30
    root_child0.Height = 10
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.Width = 30
    root_child1.Height = 20
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(Config)
    root_child2.Width = 30
    root_child2.Height = 30
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(Config)
    root_child3.Width = 30
    root_child3.Height = 30
    root.Insert 3, root_child3
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 60!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 10!, root_child0.LayoutY
    Assert.AreEqual 30!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 30!, root_child1.LayoutX
    Assert.AreEqual 5!, root_child1.LayoutY
    Assert.AreEqual 30!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 60!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 30!, root_child2.LayoutWidth
    Assert.AreEqual 30!, root_child2.LayoutHeight

    Assert.AreEqual 0!, root_child3.LayoutX
    Assert.AreEqual 30!, root_child3.LayoutY
    Assert.AreEqual 30!, root_child3.LayoutWidth
    Assert.AreEqual 30!, root_child3.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 60!, root.LayoutHeight

    Assert.AreEqual 70!, root_child0.LayoutX
    Assert.AreEqual 10!, root_child0.LayoutY
    Assert.AreEqual 30!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 40!, root_child1.LayoutX
    Assert.AreEqual 5!, root_child1.LayoutY
    Assert.AreEqual 30!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 10!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 30!, root_child2.LayoutWidth
    Assert.AreEqual 30!, root_child2.LayoutHeight

    Assert.AreEqual 70!, root_child3.LayoutX
    Assert.AreEqual 30!, root_child3.LayoutY
    Assert.AreEqual 30!, root_child3.LayoutWidth
    Assert.AreEqual 30!, root_child3.LayoutHeight
End Sub

Public Sub Test_flex_wrap_children_with_min_main_overriding_flex_basis()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.FlexDirection = YogaFlexDirection.yogaFlexRow
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.FlexBasis = 50
    root_child0.MinWidth = 55
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.FlexBasis = 50
    root_child1.MinWidth = 55
    root_child1.Height = 50
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 55!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 55!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 45!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 55!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 45!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 55!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight
End Sub

Public Sub Test_flex_wrap_wrap_to_child_height()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.FlexDirection = YogaFlexDirection.yogaFlexRow
    root_child0.AlignItems = YogaAlign.yogaAlignFlexStart
    root_child0.Wrap = YogaWrap.yogaWrapWrap
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(Config)
    root_child0_child0.Width = 100
    root_child0.Insert 0, root_child0_child0

    Dim root_child0_child0_child0 As cYogaNode: Set root_child0_child0_child0 = YogaNodeNew(Config)
    root_child0_child0_child0.Width = 100
    root_child0_child0_child0.Height = 100
    root_child0_child0.Insert 0, root_child0_child0_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.Width = 100
    root_child1.Height = 100
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 100!, root_child0_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutY
    Assert.AreEqual 100!, root_child0_child0_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 100!, root_child1.LayoutY
    Assert.AreEqual 100!, root_child1.LayoutWidth
    Assert.AreEqual 100!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 100!, root_child0_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutY
    Assert.AreEqual 100!, root_child0_child0_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 100!, root_child1.LayoutY
    Assert.AreEqual 100!, root_child1.LayoutWidth
    Assert.AreEqual 100!, root_child1.LayoutHeight
End Sub

Public Sub Test_flex_wrap_align_stretch_fits_one_row()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.FlexDirection = YogaFlexDirection.yogaFlexRow
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 150
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Width = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.Width = 50
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 150!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 100!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 150!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 100!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 100!, root_child1.LayoutHeight
End Sub

Public Sub Test_wrap_reverse_row_align_content_flex_start()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.FlexDirection = YogaFlexDirection.yogaFlexRow
    root.Wrap = YogaWrap.yogaWrapWrapReverse
    root.Width = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Width = 30
    root_child0.Height = 10
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.Width = 30
    root_child1.Height = 20
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(Config)
    root_child2.Width = 30
    root_child2.Height = 30
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(Config)
    root_child3.Width = 30
    root_child3.Height = 40
    root.Insert 3, root_child3

    Dim root_child4 As cYogaNode: Set root_child4 = YogaNodeNew(Config)
    root_child4.Width = 30
    root_child4.Height = 50
    root.Insert 4, root_child4
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 80!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 70!, root_child0.LayoutY
    Assert.AreEqual 30!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 30!, root_child1.LayoutX
    Assert.AreEqual 60!, root_child1.LayoutY
    Assert.AreEqual 30!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 60!, root_child2.LayoutX
    Assert.AreEqual 50!, root_child2.LayoutY
    Assert.AreEqual 30!, root_child2.LayoutWidth
    Assert.AreEqual 30!, root_child2.LayoutHeight

    Assert.AreEqual 0!, root_child3.LayoutX
    Assert.AreEqual 10!, root_child3.LayoutY
    Assert.AreEqual 30!, root_child3.LayoutWidth
    Assert.AreEqual 40!, root_child3.LayoutHeight

    Assert.AreEqual 30!, root_child4.LayoutX
    Assert.AreEqual 0!, root_child4.LayoutY
    Assert.AreEqual 30!, root_child4.LayoutWidth
    Assert.AreEqual 50!, root_child4.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 80!, root.LayoutHeight

    Assert.AreEqual 70!, root_child0.LayoutX
    Assert.AreEqual 70!, root_child0.LayoutY
    Assert.AreEqual 30!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 40!, root_child1.LayoutX
    Assert.AreEqual 60!, root_child1.LayoutY
    Assert.AreEqual 30!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 10!, root_child2.LayoutX
    Assert.AreEqual 50!, root_child2.LayoutY
    Assert.AreEqual 30!, root_child2.LayoutWidth
    Assert.AreEqual 30!, root_child2.LayoutHeight

    Assert.AreEqual 70!, root_child3.LayoutX
    Assert.AreEqual 10!, root_child3.LayoutY
    Assert.AreEqual 30!, root_child3.LayoutWidth
    Assert.AreEqual 40!, root_child3.LayoutHeight

    Assert.AreEqual 40!, root_child4.LayoutX
    Assert.AreEqual 0!, root_child4.LayoutY
    Assert.AreEqual 30!, root_child4.LayoutWidth
    Assert.AreEqual 50!, root_child4.LayoutHeight
End Sub

Public Sub Test_wrap_reverse_row_align_content_center()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.FlexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignContent = YogaAlign.yogaAlignCenter
    root.Wrap = YogaWrap.yogaWrapWrapReverse
    root.Width = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Width = 30
    root_child0.Height = 10
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.Width = 30
    root_child1.Height = 20
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(Config)
    root_child2.Width = 30
    root_child2.Height = 30
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(Config)
    root_child3.Width = 30
    root_child3.Height = 40
    root.Insert 3, root_child3

    Dim root_child4 As cYogaNode: Set root_child4 = YogaNodeNew(Config)
    root_child4.Width = 30
    root_child4.Height = 50
    root.Insert 4, root_child4
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 80!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 70!, root_child0.LayoutY
    Assert.AreEqual 30!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 30!, root_child1.LayoutX
    Assert.AreEqual 60!, root_child1.LayoutY
    Assert.AreEqual 30!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 60!, root_child2.LayoutX
    Assert.AreEqual 50!, root_child2.LayoutY
    Assert.AreEqual 30!, root_child2.LayoutWidth
    Assert.AreEqual 30!, root_child2.LayoutHeight

    Assert.AreEqual 0!, root_child3.LayoutX
    Assert.AreEqual 10!, root_child3.LayoutY
    Assert.AreEqual 30!, root_child3.LayoutWidth
    Assert.AreEqual 40!, root_child3.LayoutHeight

    Assert.AreEqual 30!, root_child4.LayoutX
    Assert.AreEqual 0!, root_child4.LayoutY
    Assert.AreEqual 30!, root_child4.LayoutWidth
    Assert.AreEqual 50!, root_child4.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 80!, root.LayoutHeight

    Assert.AreEqual 70!, root_child0.LayoutX
    Assert.AreEqual 70!, root_child0.LayoutY
    Assert.AreEqual 30!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 40!, root_child1.LayoutX
    Assert.AreEqual 60!, root_child1.LayoutY
    Assert.AreEqual 30!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 10!, root_child2.LayoutX
    Assert.AreEqual 50!, root_child2.LayoutY
    Assert.AreEqual 30!, root_child2.LayoutWidth
    Assert.AreEqual 30!, root_child2.LayoutHeight

    Assert.AreEqual 70!, root_child3.LayoutX
    Assert.AreEqual 10!, root_child3.LayoutY
    Assert.AreEqual 30!, root_child3.LayoutWidth
    Assert.AreEqual 40!, root_child3.LayoutHeight

    Assert.AreEqual 40!, root_child4.LayoutX
    Assert.AreEqual 0!, root_child4.LayoutY
    Assert.AreEqual 30!, root_child4.LayoutWidth
    Assert.AreEqual 50!, root_child4.LayoutHeight
End Sub

Public Sub Test_wrap_reverse_row_single_line_different_size()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.FlexDirection = YogaFlexDirection.yogaFlexRow
    root.Wrap = YogaWrap.yogaWrapWrapReverse
    root.Width = 300

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Width = 30
    root_child0.Height = 10
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.Width = 30
    root_child1.Height = 20
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(Config)
    root_child2.Width = 30
    root_child2.Height = 30
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(Config)
    root_child3.Width = 30
    root_child3.Height = 40
    root.Insert 3, root_child3

    Dim root_child4 As cYogaNode: Set root_child4 = YogaNodeNew(Config)
    root_child4.Width = 30
    root_child4.Height = 50
    root.Insert 4, root_child4
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 300!, root.LayoutWidth
    Assert.AreEqual 50!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 40!, root_child0.LayoutY
    Assert.AreEqual 30!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 30!, root_child1.LayoutX
    Assert.AreEqual 30!, root_child1.LayoutY
    Assert.AreEqual 30!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 60!, root_child2.LayoutX
    Assert.AreEqual 20!, root_child2.LayoutY
    Assert.AreEqual 30!, root_child2.LayoutWidth
    Assert.AreEqual 30!, root_child2.LayoutHeight

    Assert.AreEqual 90!, root_child3.LayoutX
    Assert.AreEqual 10!, root_child3.LayoutY
    Assert.AreEqual 30!, root_child3.LayoutWidth
    Assert.AreEqual 40!, root_child3.LayoutHeight

    Assert.AreEqual 120!, root_child4.LayoutX
    Assert.AreEqual 0!, root_child4.LayoutY
    Assert.AreEqual 30!, root_child4.LayoutWidth
    Assert.AreEqual 50!, root_child4.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 300!, root.LayoutWidth
    Assert.AreEqual 50!, root.LayoutHeight

    Assert.AreEqual 270!, root_child0.LayoutX
    Assert.AreEqual 40!, root_child0.LayoutY
    Assert.AreEqual 30!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 240!, root_child1.LayoutX
    Assert.AreEqual 30!, root_child1.LayoutY
    Assert.AreEqual 30!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 210!, root_child2.LayoutX
    Assert.AreEqual 20!, root_child2.LayoutY
    Assert.AreEqual 30!, root_child2.LayoutWidth
    Assert.AreEqual 30!, root_child2.LayoutHeight

    Assert.AreEqual 180!, root_child3.LayoutX
    Assert.AreEqual 10!, root_child3.LayoutY
    Assert.AreEqual 30!, root_child3.LayoutWidth
    Assert.AreEqual 40!, root_child3.LayoutHeight

    Assert.AreEqual 150!, root_child4.LayoutX
    Assert.AreEqual 0!, root_child4.LayoutY
    Assert.AreEqual 30!, root_child4.LayoutWidth
    Assert.AreEqual 50!, root_child4.LayoutHeight
End Sub

Public Sub Test_wrap_reverse_row_align_content_stretch()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.FlexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignContent = YogaAlign.yogaAlignStretch
    root.Wrap = YogaWrap.yogaWrapWrapReverse
    root.Width = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Width = 30
    root_child0.Height = 10
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.Width = 30
    root_child1.Height = 20
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(Config)
    root_child2.Width = 30
    root_child2.Height = 30
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(Config)
    root_child3.Width = 30
    root_child3.Height = 40
    root.Insert 3, root_child3

    Dim root_child4 As cYogaNode: Set root_child4 = YogaNodeNew(Config)
    root_child4.Width = 30
    root_child4.Height = 50
    root.Insert 4, root_child4
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 80!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 70!, root_child0.LayoutY
    Assert.AreEqual 30!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 30!, root_child1.LayoutX
    Assert.AreEqual 60!, root_child1.LayoutY
    Assert.AreEqual 30!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 60!, root_child2.LayoutX
    Assert.AreEqual 50!, root_child2.LayoutY
    Assert.AreEqual 30!, root_child2.LayoutWidth
    Assert.AreEqual 30!, root_child2.LayoutHeight

    Assert.AreEqual 0!, root_child3.LayoutX
    Assert.AreEqual 10!, root_child3.LayoutY
    Assert.AreEqual 30!, root_child3.LayoutWidth
    Assert.AreEqual 40!, root_child3.LayoutHeight

    Assert.AreEqual 30!, root_child4.LayoutX
    Assert.AreEqual 0!, root_child4.LayoutY
    Assert.AreEqual 30!, root_child4.LayoutWidth
    Assert.AreEqual 50!, root_child4.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 80!, root.LayoutHeight

    Assert.AreEqual 70!, root_child0.LayoutX
    Assert.AreEqual 70!, root_child0.LayoutY
    Assert.AreEqual 30!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 40!, root_child1.LayoutX
    Assert.AreEqual 60!, root_child1.LayoutY
    Assert.AreEqual 30!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 10!, root_child2.LayoutX
    Assert.AreEqual 50!, root_child2.LayoutY
    Assert.AreEqual 30!, root_child2.LayoutWidth
    Assert.AreEqual 30!, root_child2.LayoutHeight

    Assert.AreEqual 70!, root_child3.LayoutX
    Assert.AreEqual 10!, root_child3.LayoutY
    Assert.AreEqual 30!, root_child3.LayoutWidth
    Assert.AreEqual 40!, root_child3.LayoutHeight

    Assert.AreEqual 40!, root_child4.LayoutX
    Assert.AreEqual 0!, root_child4.LayoutY
    Assert.AreEqual 30!, root_child4.LayoutWidth
    Assert.AreEqual 50!, root_child4.LayoutHeight
End Sub

Public Sub Test_wrap_reverse_row_align_content_space_around()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.FlexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignContent = YogaAlign.yogaAlignSpaceAround
    root.Wrap = YogaWrap.yogaWrapWrapReverse
    root.Width = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Width = 30
    root_child0.Height = 10
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.Width = 30
    root_child1.Height = 20
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(Config)
    root_child2.Width = 30
    root_child2.Height = 30
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(Config)
    root_child3.Width = 30
    root_child3.Height = 40
    root.Insert 3, root_child3

    Dim root_child4 As cYogaNode: Set root_child4 = YogaNodeNew(Config)
    root_child4.Width = 30
    root_child4.Height = 50
    root.Insert 4, root_child4
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 80!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 70!, root_child0.LayoutY
    Assert.AreEqual 30!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 30!, root_child1.LayoutX
    Assert.AreEqual 60!, root_child1.LayoutY
    Assert.AreEqual 30!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 60!, root_child2.LayoutX
    Assert.AreEqual 50!, root_child2.LayoutY
    Assert.AreEqual 30!, root_child2.LayoutWidth
    Assert.AreEqual 30!, root_child2.LayoutHeight

    Assert.AreEqual 0!, root_child3.LayoutX
    Assert.AreEqual 10!, root_child3.LayoutY
    Assert.AreEqual 30!, root_child3.LayoutWidth
    Assert.AreEqual 40!, root_child3.LayoutHeight

    Assert.AreEqual 30!, root_child4.LayoutX
    Assert.AreEqual 0!, root_child4.LayoutY
    Assert.AreEqual 30!, root_child4.LayoutWidth
    Assert.AreEqual 50!, root_child4.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 80!, root.LayoutHeight

    Assert.AreEqual 70!, root_child0.LayoutX
    Assert.AreEqual 70!, root_child0.LayoutY
    Assert.AreEqual 30!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 40!, root_child1.LayoutX
    Assert.AreEqual 60!, root_child1.LayoutY
    Assert.AreEqual 30!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 10!, root_child2.LayoutX
    Assert.AreEqual 50!, root_child2.LayoutY
    Assert.AreEqual 30!, root_child2.LayoutWidth
    Assert.AreEqual 30!, root_child2.LayoutHeight

    Assert.AreEqual 70!, root_child3.LayoutX
    Assert.AreEqual 10!, root_child3.LayoutY
    Assert.AreEqual 30!, root_child3.LayoutWidth
    Assert.AreEqual 40!, root_child3.LayoutHeight

    Assert.AreEqual 40!, root_child4.LayoutX
    Assert.AreEqual 0!, root_child4.LayoutY
    Assert.AreEqual 30!, root_child4.LayoutWidth
    Assert.AreEqual 50!, root_child4.LayoutHeight
End Sub

Public Sub Test_wrap_reverse_column_fixed_size()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.Wrap = YogaWrap.yogaWrapWrapReverse
    root.Width = 200
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Width = 30
    root_child0.Height = 10
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.Width = 30
    root_child1.Height = 20
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(Config)
    root_child2.Width = 30
    root_child2.Height = 30
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(Config)
    root_child3.Width = 30
    root_child3.Height = 40
    root.Insert 3, root_child3

    Dim root_child4 As cYogaNode: Set root_child4 = YogaNodeNew(Config)
    root_child4.Width = 30
    root_child4.Height = 50
    root.Insert 4, root_child4
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 170!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 30!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 170!, root_child1.LayoutX
    Assert.AreEqual 10!, root_child1.LayoutY
    Assert.AreEqual 30!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 170!, root_child2.LayoutX
    Assert.AreEqual 30!, root_child2.LayoutY
    Assert.AreEqual 30!, root_child2.LayoutWidth
    Assert.AreEqual 30!, root_child2.LayoutHeight

    Assert.AreEqual 170!, root_child3.LayoutX
    Assert.AreEqual 60!, root_child3.LayoutY
    Assert.AreEqual 30!, root_child3.LayoutWidth
    Assert.AreEqual 40!, root_child3.LayoutHeight

    Assert.AreEqual 140!, root_child4.LayoutX
    Assert.AreEqual 0!, root_child4.LayoutY
    Assert.AreEqual 30!, root_child4.LayoutWidth
    Assert.AreEqual 50!, root_child4.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 30!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 10!, root_child1.LayoutY
    Assert.AreEqual 30!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 30!, root_child2.LayoutY
    Assert.AreEqual 30!, root_child2.LayoutWidth
    Assert.AreEqual 30!, root_child2.LayoutHeight

    Assert.AreEqual 0!, root_child3.LayoutX
    Assert.AreEqual 60!, root_child3.LayoutY
    Assert.AreEqual 30!, root_child3.LayoutWidth
    Assert.AreEqual 40!, root_child3.LayoutHeight

    Assert.AreEqual 30!, root_child4.LayoutX
    Assert.AreEqual 0!, root_child4.LayoutY
    Assert.AreEqual 30!, root_child4.LayoutWidth
    Assert.AreEqual 50!, root_child4.LayoutHeight
End Sub

Public Sub Test_wrapped_row_within_align_items_center()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.FlexDirection = YogaFlexDirection.yogaFlexRow
    root_child0.Wrap = YogaWrap.yogaWrapWrap
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(Config)
    root_child0_child0.Width = 150
    root_child0_child0.Height = 80
    root_child0.Insert 0, root_child0_child0

    Dim root_child0_child1 As cYogaNode: Set root_child0_child1 = YogaNodeNew(Config)
    root_child0_child1.Width = 80
    root_child0_child1.Height = 80
    root_child0.Insert 1, root_child0_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 200!, root_child0.LayoutWidth
    Assert.AreEqual 160!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 150!, root_child0_child0.LayoutWidth
    Assert.AreEqual 80!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child1.LayoutX
    Assert.AreEqual 80!, root_child0_child1.LayoutY
    Assert.AreEqual 80!, root_child0_child1.LayoutWidth
    Assert.AreEqual 80!, root_child0_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 200!, root_child0.LayoutWidth
    Assert.AreEqual 160!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 150!, root_child0_child0.LayoutWidth
    Assert.AreEqual 80!, root_child0_child0.LayoutHeight

    Assert.AreEqual 120!, root_child0_child1.LayoutX
    Assert.AreEqual 80!, root_child0_child1.LayoutY
    Assert.AreEqual 80!, root_child0_child1.LayoutWidth
    Assert.AreEqual 80!, root_child0_child1.LayoutHeight
End Sub

Public Sub Test_wrapped_row_within_align_items_flex_start()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.AlignItems = YogaAlign.yogaAlignFlexStart
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.FlexDirection = YogaFlexDirection.yogaFlexRow
    root_child0.Wrap = YogaWrap.yogaWrapWrap
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(Config)
    root_child0_child0.Width = 150
    root_child0_child0.Height = 80
    root_child0.Insert 0, root_child0_child0

    Dim root_child0_child1 As cYogaNode: Set root_child0_child1 = YogaNodeNew(Config)
    root_child0_child1.Width = 80
    root_child0_child1.Height = 80
    root_child0.Insert 1, root_child0_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 200!, root_child0.LayoutWidth
    Assert.AreEqual 160!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 150!, root_child0_child0.LayoutWidth
    Assert.AreEqual 80!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child1.LayoutX
    Assert.AreEqual 80!, root_child0_child1.LayoutY
    Assert.AreEqual 80!, root_child0_child1.LayoutWidth
    Assert.AreEqual 80!, root_child0_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 200!, root_child0.LayoutWidth
    Assert.AreEqual 160!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 150!, root_child0_child0.LayoutWidth
    Assert.AreEqual 80!, root_child0_child0.LayoutHeight

    Assert.AreEqual 120!, root_child0_child1.LayoutX
    Assert.AreEqual 80!, root_child0_child1.LayoutY
    Assert.AreEqual 80!, root_child0_child1.LayoutWidth
    Assert.AreEqual 80!, root_child0_child1.LayoutHeight
End Sub

Public Sub Test_wrapped_row_within_align_items_flex_end()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.AlignItems = YogaAlign.yogaAlignFlexEnd
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.FlexDirection = YogaFlexDirection.yogaFlexRow
    root_child0.Wrap = YogaWrap.yogaWrapWrap
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(Config)
    root_child0_child0.Width = 150
    root_child0_child0.Height = 80
    root_child0.Insert 0, root_child0_child0

    Dim root_child0_child1 As cYogaNode: Set root_child0_child1 = YogaNodeNew(Config)
    root_child0_child1.Width = 80
    root_child0_child1.Height = 80
    root_child0.Insert 1, root_child0_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 200!, root_child0.LayoutWidth
    Assert.AreEqual 160!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 150!, root_child0_child0.LayoutWidth
    Assert.AreEqual 80!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child1.LayoutX
    Assert.AreEqual 80!, root_child0_child1.LayoutY
    Assert.AreEqual 80!, root_child0_child1.LayoutWidth
    Assert.AreEqual 80!, root_child0_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 200!, root_child0.LayoutWidth
    Assert.AreEqual 160!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 150!, root_child0_child0.LayoutWidth
    Assert.AreEqual 80!, root_child0_child0.LayoutHeight

    Assert.AreEqual 120!, root_child0_child1.LayoutX
    Assert.AreEqual 80!, root_child0_child1.LayoutY
    Assert.AreEqual 80!, root_child0_child1.LayoutWidth
    Assert.AreEqual 80!, root_child0_child1.LayoutHeight
End Sub

Public Sub Test_wrapped_column_max_height()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.JustifyContent = YogaJustify.yogaJustCenter
    root.AlignContent = YogaAlign.yogaAlignCenter
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 700
    root.Height = 500

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.Width = 100
    root_child0.Height = 500
    root_child0.MaxHeight = 200
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.MarginLeft = 20
    root_child1.MarginTop = 20
    root_child1.MarginRight = 20
    root_child1.MarginBottom = 20
    root_child1.Width = 200
    root_child1.Height = 200
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(Config)
    root_child2.Width = 100
    root_child2.Height = 100
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 700!, root.LayoutWidth
    Assert.AreEqual 500!, root.LayoutHeight

    Assert.AreEqual 250!, root_child0.LayoutX
    Assert.AreEqual 30!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 200!, root_child0.LayoutHeight

    Assert.AreEqual 200!, root_child1.LayoutX
    Assert.AreEqual 250!, root_child1.LayoutY
    Assert.AreEqual 200!, root_child1.LayoutWidth
    Assert.AreEqual 200!, root_child1.LayoutHeight

    Assert.AreEqual 420!, root_child2.LayoutX
    Assert.AreEqual 200!, root_child2.LayoutY
    Assert.AreEqual 100!, root_child2.LayoutWidth
    Assert.AreEqual 100!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 700!, root.LayoutWidth
    Assert.AreEqual 500!, root.LayoutHeight

    Assert.AreEqual 350!, root_child0.LayoutX
    Assert.AreEqual 30!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 200!, root_child0.LayoutHeight

    Assert.AreEqual 300!, root_child1.LayoutX
    Assert.AreEqual 250!, root_child1.LayoutY
    Assert.AreEqual 200!, root_child1.LayoutWidth
    Assert.AreEqual 200!, root_child1.LayoutHeight

    Assert.AreEqual 180!, root_child2.LayoutX
    Assert.AreEqual 200!, root_child2.LayoutY
    Assert.AreEqual 100!, root_child2.LayoutWidth
    Assert.AreEqual 100!, root_child2.LayoutHeight
End Sub

Public Sub Test_wrapped_column_max_height_flex()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.JustifyContent = YogaJustify.yogaJustCenter
    root.AlignContent = YogaAlign.yogaAlignCenter
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 700
    root.Height = 500

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.FlexGrow = 1
    root_child0.FlexShrink = 1
    root_child0.FlexBasis = Array(yogaUnitPercent, 0)
    root_child0.Width = 100
    root_child0.Height = 500
    root_child0.MaxHeight = 200
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.FlexGrow = 1
    root_child1.FlexShrink = 1
    root_child1.FlexBasis = Array(yogaUnitPercent, 0)
    root_child1.MarginLeft = 20
    root_child1.MarginTop = 20
    root_child1.MarginRight = 20
    root_child1.MarginBottom = 20
    root_child1.Width = 200
    root_child1.Height = 200
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(Config)
    root_child2.Width = 100
    root_child2.Height = 100
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 700!, root.LayoutWidth
    Assert.AreEqual 500!, root.LayoutHeight

    Assert.AreEqual 300!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 180!, root_child0.LayoutHeight

    Assert.AreEqual 250!, root_child1.LayoutX
    Assert.AreEqual 200!, root_child1.LayoutY
    Assert.AreEqual 200!, root_child1.LayoutWidth
    Assert.AreEqual 180!, root_child1.LayoutHeight

    Assert.AreEqual 300!, root_child2.LayoutX
    Assert.AreEqual 400!, root_child2.LayoutY
    Assert.AreEqual 100!, root_child2.LayoutWidth
    Assert.AreEqual 100!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 700!, root.LayoutWidth
    Assert.AreEqual 500!, root.LayoutHeight

    Assert.AreEqual 300!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 180!, root_child0.LayoutHeight

    Assert.AreEqual 250!, root_child1.LayoutX
    Assert.AreEqual 200!, root_child1.LayoutY
    Assert.AreEqual 200!, root_child1.LayoutWidth
    Assert.AreEqual 180!, root_child1.LayoutHeight

    Assert.AreEqual 300!, root_child2.LayoutX
    Assert.AreEqual 400!, root_child2.LayoutY
    Assert.AreEqual 100!, root_child2.LayoutWidth
    Assert.AreEqual 100!, root_child2.LayoutHeight
End Sub

Public Sub Test_wrap_nodes_with_content_sizing_overflowing_margin()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.Width = 500
    root.Height = 500

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.FlexDirection = YogaFlexDirection.yogaFlexRow
    root_child0.Wrap = YogaWrap.yogaWrapWrap
    root_child0.Width = 85
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(Config)
    root_child0.Insert 0, root_child0_child0

    Dim root_child0_child0_child0 As cYogaNode: Set root_child0_child0_child0 = YogaNodeNew(Config)
    root_child0_child0_child0.Width = 40
    root_child0_child0_child0.Height = 40
    root_child0_child0.Insert 0, root_child0_child0_child0

    Dim root_child0_child1 As cYogaNode: Set root_child0_child1 = YogaNodeNew(Config)
    root_child0_child1.MarginRight = 10
    root_child0.Insert 1, root_child0_child1

    Dim root_child0_child1_child0 As cYogaNode: Set root_child0_child1_child0 = YogaNodeNew(Config)
    root_child0_child1_child0.Width = 40
    root_child0_child1_child0.Height = 40
    root_child0_child1.Insert 0, root_child0_child1_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 500!, root.LayoutWidth
    Assert.AreEqual 500!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 85!, root_child0.LayoutWidth
    Assert.AreEqual 80!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 40!, root_child0_child0.LayoutWidth
    Assert.AreEqual 40!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutY
    Assert.AreEqual 40!, root_child0_child0_child0.LayoutWidth
    Assert.AreEqual 40!, root_child0_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child1.LayoutX
    Assert.AreEqual 40!, root_child0_child1.LayoutY
    Assert.AreEqual 40!, root_child0_child1.LayoutWidth
    Assert.AreEqual 40!, root_child0_child1.LayoutHeight

    Assert.AreEqual 0!, root_child0_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child1_child0.LayoutY
    Assert.AreEqual 40!, root_child0_child1_child0.LayoutWidth
    Assert.AreEqual 40!, root_child0_child1_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 500!, root.LayoutWidth
    Assert.AreEqual 500!, root.LayoutHeight

    Assert.AreEqual 415!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 85!, root_child0.LayoutWidth
    Assert.AreEqual 80!, root_child0.LayoutHeight

    Assert.AreEqual 45!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 40!, root_child0_child0.LayoutWidth
    Assert.AreEqual 40!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutY
    Assert.AreEqual 40!, root_child0_child0_child0.LayoutWidth
    Assert.AreEqual 40!, root_child0_child0_child0.LayoutHeight

    Assert.AreEqual 35!, root_child0_child1.LayoutX
    Assert.AreEqual 40!, root_child0_child1.LayoutY
    Assert.AreEqual 40!, root_child0_child1.LayoutWidth
    Assert.AreEqual 40!, root_child0_child1.LayoutHeight

    Assert.AreEqual 0!, root_child0_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child1_child0.LayoutY
    Assert.AreEqual 40!, root_child0_child1_child0.LayoutWidth
    Assert.AreEqual 40!, root_child0_child1_child0.LayoutHeight
End Sub

Public Sub Test_wrap_nodes_with_content_sizing_margin_cross()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.Width = 500
    root.Height = 500

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.FlexDirection = YogaFlexDirection.yogaFlexRow
    root_child0.Wrap = YogaWrap.yogaWrapWrap
    root_child0.Width = 70
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(Config)
    root_child0.Insert 0, root_child0_child0

    Dim root_child0_child0_child0 As cYogaNode: Set root_child0_child0_child0 = YogaNodeNew(Config)
    root_child0_child0_child0.Width = 40
    root_child0_child0_child0.Height = 40
    root_child0_child0.Insert 0, root_child0_child0_child0

    Dim root_child0_child1 As cYogaNode: Set root_child0_child1 = YogaNodeNew(Config)
    root_child0_child1.MarginTop = 10
    root_child0.Insert 1, root_child0_child1

    Dim root_child0_child1_child0 As cYogaNode: Set root_child0_child1_child0 = YogaNodeNew(Config)
    root_child0_child1_child0.Width = 40
    root_child0_child1_child0.Height = 40
    root_child0_child1.Insert 0, root_child0_child1_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 500!, root.LayoutWidth
    Assert.AreEqual 500!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 70!, root_child0.LayoutWidth
    Assert.AreEqual 90!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 40!, root_child0_child0.LayoutWidth
    Assert.AreEqual 40!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutY
    Assert.AreEqual 40!, root_child0_child0_child0.LayoutWidth
    Assert.AreEqual 40!, root_child0_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child1.LayoutX
    Assert.AreEqual 50!, root_child0_child1.LayoutY
    Assert.AreEqual 40!, root_child0_child1.LayoutWidth
    Assert.AreEqual 40!, root_child0_child1.LayoutHeight

    Assert.AreEqual 0!, root_child0_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child1_child0.LayoutY
    Assert.AreEqual 40!, root_child0_child1_child0.LayoutWidth
    Assert.AreEqual 40!, root_child0_child1_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 500!, root.LayoutWidth
    Assert.AreEqual 500!, root.LayoutHeight

    Assert.AreEqual 430!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 70!, root_child0.LayoutWidth
    Assert.AreEqual 90!, root_child0.LayoutHeight

    Assert.AreEqual 30!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 40!, root_child0_child0.LayoutWidth
    Assert.AreEqual 40!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutY
    Assert.AreEqual 40!, root_child0_child0_child0.LayoutWidth
    Assert.AreEqual 40!, root_child0_child0_child0.LayoutHeight

    Assert.AreEqual 30!, root_child0_child1.LayoutX
    Assert.AreEqual 50!, root_child0_child1.LayoutY
    Assert.AreEqual 40!, root_child0_child1.LayoutWidth
    Assert.AreEqual 40!, root_child0_child1.LayoutHeight

    Assert.AreEqual 0!, root_child0_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child1_child0.LayoutY
    Assert.AreEqual 40!, root_child0_child1_child0.LayoutWidth
    Assert.AreEqual 40!, root_child0_child1_child0.LayoutHeight
End Sub

Private Sub ITestFixture_GetTestCases(ByVal Tests As SimplyVBComp.TestFixtureBuilder)
    Tests.Add "Test_wrap_column"
    Tests.Add "Test_wrap_row"
    Tests.Add "Test_wrap_row_align_items_flex_end"
    Tests.Add "Test_wrap_row_align_items_center"
    Tests.Add "Test_flex_wrap_children_with_min_main_overriding_flex_basis"
    Tests.Add "Test_flex_wrap_wrap_to_child_height"
    Tests.Add "Test_flex_wrap_align_stretch_fits_one_row"
    Tests.Add "Test_wrap_reverse_row_align_content_flex_start"
    Tests.Add "Test_wrap_reverse_row_align_content_center"
    Tests.Add "Test_wrap_reverse_row_single_line_different_size"
    Tests.Add "Test_wrap_reverse_row_align_content_stretch"
    Tests.Add "Test_wrap_reverse_row_align_content_space_around"
    Tests.Add "Test_wrap_reverse_column_fixed_size"
    Tests.Add "Test_wrapped_row_within_align_items_center"
    Tests.Add "Test_wrapped_row_within_align_items_flex_start"
    Tests.Add "Test_wrapped_row_within_align_items_flex_end"
    Tests.Add "Test_wrapped_column_max_height"
    Tests.Add "Test_wrapped_column_max_height_flex"
    Tests.Add "Test_wrap_nodes_with_content_sizing_overflowing_margin"
    Tests.Add "Test_wrap_nodes_with_content_sizing_margin_cross"
End Sub


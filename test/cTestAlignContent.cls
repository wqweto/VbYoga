VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTestAlignContent"
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

Public Sub Test_align_content_flex_start()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 130
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 50
    root_child0.Height = 10
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 50
    root_child1.Height = 10
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Width = 50
    root_child2.Height = 10
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(config)
    root_child3.Width = 50
    root_child3.Height = 10
    root.Insert 3, root_child3

    Dim root_child4 As cYogaNode: Set root_child4 = YogaNodeNew(config)
    root_child4.Width = 50
    root_child4.Height = 10
    root.Insert 4, root_child4
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 130!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 10!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 10!, root_child2.LayoutHeight

    Assert.AreEqual 50!, root_child3.LayoutX
    Assert.AreEqual 10!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 10!, root_child3.LayoutHeight

    Assert.AreEqual 0!, root_child4.LayoutX
    Assert.AreEqual 20!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 10!, root_child4.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 130!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 80!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 30!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1.LayoutHeight

    Assert.AreEqual 80!, root_child2.LayoutX
    Assert.AreEqual 10!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 10!, root_child2.LayoutHeight

    Assert.AreEqual 30!, root_child3.LayoutX
    Assert.AreEqual 10!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 10!, root_child3.LayoutHeight

    Assert.AreEqual 80!, root_child4.LayoutX
    Assert.AreEqual 20!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 10!, root_child4.LayoutHeight
End Sub

Public Sub Test_align_content_flex_start_without_height_on_children()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 50
    root_child1.Height = 10
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Width = 50
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(config)
    root_child3.Width = 50
    root_child3.Height = 10
    root.Insert 3, root_child3

    Dim root_child4 As cYogaNode: Set root_child4 = YogaNodeNew(config)
    root_child4.Width = 50
    root.Insert 4, root_child4
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 10!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 0!, root_child2.LayoutHeight

    Assert.AreEqual 0!, root_child3.LayoutX
    Assert.AreEqual 10!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 10!, root_child3.LayoutHeight

    Assert.AreEqual 0!, root_child4.LayoutX
    Assert.AreEqual 20!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 0!, root_child4.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 50!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1.LayoutHeight

    Assert.AreEqual 50!, root_child2.LayoutX
    Assert.AreEqual 10!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 0!, root_child2.LayoutHeight

    Assert.AreEqual 50!, root_child3.LayoutX
    Assert.AreEqual 10!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 10!, root_child3.LayoutHeight

    Assert.AreEqual 50!, root_child4.LayoutX
    Assert.AreEqual 20!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 0!, root_child4.LayoutHeight
End Sub

Public Sub Test_align_content_flex_start_with_flex()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 100
    root.Height = 120

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.FlexBasis = Array(yogaUnitPercent, 0)
    root_child0.Width = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 1
    root_child1.FlexBasis = Array(yogaUnitPercent, 0)
    root_child1.Width = 50
    root_child1.Height = 10
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Width = 50
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(config)
    root_child3.FlexGrow = 1
    root_child3.FlexShrink = 1
    root_child3.FlexBasis = Array(yogaUnitPercent, 0)
    root_child3.Width = 50
    root.Insert 3, root_child3

    Dim root_child4 As cYogaNode: Set root_child4 = YogaNodeNew(config)
    root_child4.Width = 50
    root.Insert 4, root_child4
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 120!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 40!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 40!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 40!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 80!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 0!, root_child2.LayoutHeight

    Assert.AreEqual 0!, root_child3.LayoutX
    Assert.AreEqual 80!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 40!, root_child3.LayoutHeight

    Assert.AreEqual 0!, root_child4.LayoutX
    Assert.AreEqual 120!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 0!, root_child4.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 120!, root.LayoutHeight

    Assert.AreEqual 50!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 40!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 40!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 40!, root_child1.LayoutHeight

    Assert.AreEqual 50!, root_child2.LayoutX
    Assert.AreEqual 80!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 0!, root_child2.LayoutHeight

    Assert.AreEqual 50!, root_child3.LayoutX
    Assert.AreEqual 80!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 40!, root_child3.LayoutHeight

    Assert.AreEqual 50!, root_child4.LayoutX
    Assert.AreEqual 120!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 0!, root_child4.LayoutHeight
End Sub

Public Sub Test_align_content_flex_end()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.AlignContent = YogaAlign.yogaAlignFlexEnd
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 50
    root_child0.Height = 10
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 50
    root_child1.Height = 10
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Width = 50
    root_child2.Height = 10
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(config)
    root_child3.Width = 50
    root_child3.Height = 10
    root.Insert 3, root_child3

    Dim root_child4 As cYogaNode: Set root_child4 = YogaNodeNew(config)
    root_child4.Width = 50
    root_child4.Height = 10
    root.Insert 4, root_child4
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 10!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 20!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 10!, root_child2.LayoutHeight

    Assert.AreEqual 0!, root_child3.LayoutX
    Assert.AreEqual 30!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 10!, root_child3.LayoutHeight

    Assert.AreEqual 0!, root_child4.LayoutX
    Assert.AreEqual 40!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 10!, root_child4.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 50!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 10!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1.LayoutHeight

    Assert.AreEqual 50!, root_child2.LayoutX
    Assert.AreEqual 20!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 10!, root_child2.LayoutHeight

    Assert.AreEqual 50!, root_child3.LayoutX
    Assert.AreEqual 30!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 10!, root_child3.LayoutHeight

    Assert.AreEqual 50!, root_child4.LayoutX
    Assert.AreEqual 40!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 10!, root_child4.LayoutHeight
End Sub

Public Sub Test_align_content_stretch()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.AlignContent = YogaAlign.yogaAlignStretch
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 150
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 50
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Width = 50
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(config)
    root_child3.Width = 50
    root.Insert 3, root_child3

    Dim root_child4 As cYogaNode: Set root_child4 = YogaNodeNew(config)
    root_child4.Width = 50
    root.Insert 4, root_child4
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 150!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 0!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 0!, root_child2.LayoutHeight

    Assert.AreEqual 0!, root_child3.LayoutX
    Assert.AreEqual 0!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 0!, root_child3.LayoutHeight

    Assert.AreEqual 0!, root_child4.LayoutX
    Assert.AreEqual 0!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 0!, root_child4.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 150!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 100!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0.LayoutHeight

    Assert.AreEqual 100!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 0!, root_child1.LayoutHeight

    Assert.AreEqual 100!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 0!, root_child2.LayoutHeight

    Assert.AreEqual 100!, root_child3.LayoutX
    Assert.AreEqual 0!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 0!, root_child3.LayoutHeight

    Assert.AreEqual 100!, root_child4.LayoutX
    Assert.AreEqual 0!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 0!, root_child4.LayoutHeight
End Sub

Public Sub Test_align_content_spacebetween()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignContent = YogaAlign.yogaAlignSpaceBetween
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 130
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 50
    root_child0.Height = 10
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 50
    root_child1.Height = 10
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Width = 50
    root_child2.Height = 10
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(config)
    root_child3.Width = 50
    root_child3.Height = 10
    root.Insert 3, root_child3

    Dim root_child4 As cYogaNode: Set root_child4 = YogaNodeNew(config)
    root_child4.Width = 50
    root_child4.Height = 10
    root.Insert 4, root_child4
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 130!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 45!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 10!, root_child2.LayoutHeight

    Assert.AreEqual 50!, root_child3.LayoutX
    Assert.AreEqual 45!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 10!, root_child3.LayoutHeight

    Assert.AreEqual 0!, root_child4.LayoutX
    Assert.AreEqual 90!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 10!, root_child4.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 130!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 80!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 30!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1.LayoutHeight

    Assert.AreEqual 80!, root_child2.LayoutX
    Assert.AreEqual 45!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 10!, root_child2.LayoutHeight

    Assert.AreEqual 30!, root_child3.LayoutX
    Assert.AreEqual 45!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 10!, root_child3.LayoutHeight

    Assert.AreEqual 80!, root_child4.LayoutX
    Assert.AreEqual 90!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 10!, root_child4.LayoutHeight
End Sub

Public Sub Test_align_content_spacearound()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignContent = YogaAlign.yogaAlignSpaceAround
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 140
    root.Height = 120

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 50
    root_child0.Height = 10
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 50
    root_child1.Height = 10
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Width = 50
    root_child2.Height = 10
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(config)
    root_child3.Width = 50
    root_child3.Height = 10
    root.Insert 3, root_child3

    Dim root_child4 As cYogaNode: Set root_child4 = YogaNodeNew(config)
    root_child4.Width = 50
    root_child4.Height = 10
    root.Insert 4, root_child4
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 140!, root.LayoutWidth
    Assert.AreEqual 120!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 15!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 15!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 55!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 10!, root_child2.LayoutHeight

    Assert.AreEqual 50!, root_child3.LayoutX
    Assert.AreEqual 55!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 10!, root_child3.LayoutHeight

    Assert.AreEqual 0!, root_child4.LayoutX
    Assert.AreEqual 95!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 10!, root_child4.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 140!, root.LayoutWidth
    Assert.AreEqual 120!, root.LayoutHeight

    Assert.AreEqual 90!, root_child0.LayoutX
    Assert.AreEqual 15!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 40!, root_child1.LayoutX
    Assert.AreEqual 15!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1.LayoutHeight

    Assert.AreEqual 90!, root_child2.LayoutX
    Assert.AreEqual 55!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 10!, root_child2.LayoutHeight

    Assert.AreEqual 40!, root_child3.LayoutX
    Assert.AreEqual 55!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 10!, root_child3.LayoutHeight

    Assert.AreEqual 90!, root_child4.LayoutX
    Assert.AreEqual 95!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 10!, root_child4.LayoutHeight
End Sub

Public Sub Test_align_content_stretch_row()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignContent = YogaAlign.yogaAlignStretch
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 150
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 50
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Width = 50
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(config)
    root_child3.Width = 50
    root.Insert 3, root_child3

    Dim root_child4 As cYogaNode: Set root_child4 = YogaNodeNew(config)
    root_child4.Width = 50
    root.Insert 4, root_child4
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 150!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    Assert.AreEqual 100!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 50!, root_child2.LayoutHeight

    Assert.AreEqual 0!, root_child3.LayoutX
    Assert.AreEqual 50!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 50!, root_child3.LayoutHeight

    Assert.AreEqual 50!, root_child4.LayoutX
    Assert.AreEqual 50!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 50!, root_child4.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 150!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 100!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 50!, root_child2.LayoutHeight

    Assert.AreEqual 100!, root_child3.LayoutX
    Assert.AreEqual 50!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 50!, root_child3.LayoutHeight

    Assert.AreEqual 50!, root_child4.LayoutX
    Assert.AreEqual 50!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 50!, root_child4.LayoutHeight
End Sub

Public Sub Test_align_content_stretch_row_with_children()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignContent = YogaAlign.yogaAlignStretch
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 150
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 50
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(config)
    root_child0_child0.FlexGrow = 1
    root_child0_child0.FlexShrink = 1
    root_child0_child0.FlexBasis = Array(yogaUnitPercent, 0)
    root_child0.Insert 0, root_child0_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 50
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Width = 50
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(config)
    root_child3.Width = 50
    root.Insert 3, root_child3

    Dim root_child4 As cYogaNode: Set root_child4 = YogaNodeNew(config)
    root_child4.Width = 50
    root.Insert 4, root_child4
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 150!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 50!, root_child0_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    Assert.AreEqual 100!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 50!, root_child2.LayoutHeight

    Assert.AreEqual 0!, root_child3.LayoutX
    Assert.AreEqual 50!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 50!, root_child3.LayoutHeight

    Assert.AreEqual 50!, root_child4.LayoutX
    Assert.AreEqual 50!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 50!, root_child4.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 150!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 100!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 50!, root_child0_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 50!, root_child2.LayoutHeight

    Assert.AreEqual 100!, root_child3.LayoutX
    Assert.AreEqual 50!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 50!, root_child3.LayoutHeight

    Assert.AreEqual 50!, root_child4.LayoutX
    Assert.AreEqual 50!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 50!, root_child4.LayoutHeight
End Sub

Public Sub Test_align_content_stretch_row_with_flex()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignContent = YogaAlign.yogaAlignStretch
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 150
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 1
    root_child1.FlexShrink = 1
    root_child1.FlexBasis = Array(yogaUnitPercent, 0)
    root_child1.Width = 50
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Width = 50
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(config)
    root_child3.FlexGrow = 1
    root_child3.FlexShrink = 1
    root_child3.FlexBasis = Array(yogaUnitPercent, 0)
    root_child3.Width = 50
    root.Insert 3, root_child3

    Dim root_child4 As cYogaNode: Set root_child4 = YogaNodeNew(config)
    root_child4.Width = 50
    root.Insert 4, root_child4
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
    Assert.AreEqual 0!, root_child1.LayoutWidth
    Assert.AreEqual 100!, root_child1.LayoutHeight

    Assert.AreEqual 50!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 100!, root_child2.LayoutHeight

    Assert.AreEqual 100!, root_child3.LayoutX
    Assert.AreEqual 0!, root_child3.LayoutY
    Assert.AreEqual 0!, root_child3.LayoutWidth
    Assert.AreEqual 100!, root_child3.LayoutHeight

    Assert.AreEqual 100!, root_child4.LayoutX
    Assert.AreEqual 0!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 100!, root_child4.LayoutHeight

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

    Assert.AreEqual 100!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 0!, root_child1.LayoutWidth
    Assert.AreEqual 100!, root_child1.LayoutHeight

    Assert.AreEqual 50!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 100!, root_child2.LayoutHeight

    Assert.AreEqual 50!, root_child3.LayoutX
    Assert.AreEqual 0!, root_child3.LayoutY
    Assert.AreEqual 0!, root_child3.LayoutWidth
    Assert.AreEqual 100!, root_child3.LayoutHeight

    Assert.AreEqual 0!, root_child4.LayoutX
    Assert.AreEqual 0!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 100!, root_child4.LayoutHeight
End Sub

Public Sub Test_align_content_stretch_row_with_flex_no_shrink()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignContent = YogaAlign.yogaAlignStretch
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 150
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 1
    root_child1.FlexShrink = 1
    root_child1.FlexBasis = Array(yogaUnitPercent, 0)
    root_child1.Width = 50
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Width = 50
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(config)
    root_child3.FlexGrow = 1
    root_child3.FlexBasis = Array(yogaUnitPercent, 0)
    root_child3.Width = 50
    root.Insert 3, root_child3

    Dim root_child4 As cYogaNode: Set root_child4 = YogaNodeNew(config)
    root_child4.Width = 50
    root.Insert 4, root_child4
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
    Assert.AreEqual 0!, root_child1.LayoutWidth
    Assert.AreEqual 100!, root_child1.LayoutHeight

    Assert.AreEqual 50!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 100!, root_child2.LayoutHeight

    Assert.AreEqual 100!, root_child3.LayoutX
    Assert.AreEqual 0!, root_child3.LayoutY
    Assert.AreEqual 0!, root_child3.LayoutWidth
    Assert.AreEqual 100!, root_child3.LayoutHeight

    Assert.AreEqual 100!, root_child4.LayoutX
    Assert.AreEqual 0!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 100!, root_child4.LayoutHeight

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

    Assert.AreEqual 100!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 0!, root_child1.LayoutWidth
    Assert.AreEqual 100!, root_child1.LayoutHeight

    Assert.AreEqual 50!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 100!, root_child2.LayoutHeight

    Assert.AreEqual 50!, root_child3.LayoutX
    Assert.AreEqual 0!, root_child3.LayoutY
    Assert.AreEqual 0!, root_child3.LayoutWidth
    Assert.AreEqual 100!, root_child3.LayoutHeight

    Assert.AreEqual 0!, root_child4.LayoutX
    Assert.AreEqual 0!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 100!, root_child4.LayoutHeight
End Sub

Public Sub Test_align_content_stretch_row_with_margin()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignContent = YogaAlign.yogaAlignStretch
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 150
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.MarginLeft = 10
    root_child1.MarginTop = 10
    root_child1.MarginRight = 10
    root_child1.MarginBottom = 10
    root_child1.Width = 50
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Width = 50
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(config)
    root_child3.MarginLeft = 10
    root_child3.MarginTop = 10
    root_child3.MarginRight = 10
    root_child3.MarginBottom = 10
    root_child3.Width = 50
    root.Insert 3, root_child3

    Dim root_child4 As cYogaNode: Set root_child4 = YogaNodeNew(config)
    root_child4.Width = 50
    root.Insert 4, root_child4
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 150!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 40!, root_child0.LayoutHeight

    Assert.AreEqual 60!, root_child1.LayoutX
    Assert.AreEqual 10!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 40!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 40!, root_child2.LayoutHeight

    Assert.AreEqual 60!, root_child3.LayoutX
    Assert.AreEqual 50!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 20!, root_child3.LayoutHeight

    Assert.AreEqual 0!, root_child4.LayoutX
    Assert.AreEqual 80!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 20!, root_child4.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 150!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 100!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 40!, root_child0.LayoutHeight

    Assert.AreEqual 40!, root_child1.LayoutX
    Assert.AreEqual 10!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 100!, root_child2.LayoutX
    Assert.AreEqual 40!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 40!, root_child2.LayoutHeight

    Assert.AreEqual 40!, root_child3.LayoutX
    Assert.AreEqual 50!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 20!, root_child3.LayoutHeight

    Assert.AreEqual 100!, root_child4.LayoutX
    Assert.AreEqual 80!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 20!, root_child4.LayoutHeight
End Sub

Public Sub Test_align_content_stretch_row_with_padding()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignContent = YogaAlign.yogaAlignStretch
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 150
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.PaddingLeft = 10
    root_child1.PaddingTop = 10
    root_child1.PaddingRight = 10
    root_child1.PaddingBottom = 10
    root_child1.Width = 50
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Width = 50
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(config)
    root_child3.PaddingLeft = 10
    root_child3.PaddingTop = 10
    root_child3.PaddingRight = 10
    root_child3.PaddingBottom = 10
    root_child3.Width = 50
    root.Insert 3, root_child3

    Dim root_child4 As cYogaNode: Set root_child4 = YogaNodeNew(config)
    root_child4.Width = 50
    root.Insert 4, root_child4
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 150!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    Assert.AreEqual 100!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 50!, root_child2.LayoutHeight

    Assert.AreEqual 0!, root_child3.LayoutX
    Assert.AreEqual 50!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 50!, root_child3.LayoutHeight

    Assert.AreEqual 50!, root_child4.LayoutX
    Assert.AreEqual 50!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 50!, root_child4.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 150!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 100!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 50!, root_child2.LayoutHeight

    Assert.AreEqual 100!, root_child3.LayoutX
    Assert.AreEqual 50!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 50!, root_child3.LayoutHeight

    Assert.AreEqual 50!, root_child4.LayoutX
    Assert.AreEqual 50!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 50!, root_child4.LayoutHeight
End Sub

Public Sub Test_align_content_stretch_row_with_single_row()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignContent = YogaAlign.yogaAlignStretch
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 150
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
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

Public Sub Test_align_content_stretch_row_with_fixed_height()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignContent = YogaAlign.yogaAlignStretch
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 150
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 50
    root_child1.Height = 60
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Width = 50
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(config)
    root_child3.Width = 50
    root.Insert 3, root_child3

    Dim root_child4 As cYogaNode: Set root_child4 = YogaNodeNew(config)
    root_child4.Width = 50
    root.Insert 4, root_child4
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 150!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 80!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 60!, root_child1.LayoutHeight

    Assert.AreEqual 100!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 80!, root_child2.LayoutHeight

    Assert.AreEqual 0!, root_child3.LayoutX
    Assert.AreEqual 80!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 20!, root_child3.LayoutHeight

    Assert.AreEqual 50!, root_child4.LayoutX
    Assert.AreEqual 80!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 20!, root_child4.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 150!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 100!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 80!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 60!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 80!, root_child2.LayoutHeight

    Assert.AreEqual 100!, root_child3.LayoutX
    Assert.AreEqual 80!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 20!, root_child3.LayoutHeight

    Assert.AreEqual 50!, root_child4.LayoutX
    Assert.AreEqual 80!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 20!, root_child4.LayoutHeight
End Sub

Public Sub Test_align_content_stretch_row_with_max_height()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignContent = YogaAlign.yogaAlignStretch
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 150
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 50
    root_child1.MaxHeight = 20
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Width = 50
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(config)
    root_child3.Width = 50
    root.Insert 3, root_child3

    Dim root_child4 As cYogaNode: Set root_child4 = YogaNodeNew(config)
    root_child4.Width = 50
    root.Insert 4, root_child4
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 150!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 100!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 50!, root_child2.LayoutHeight

    Assert.AreEqual 0!, root_child3.LayoutX
    Assert.AreEqual 50!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 50!, root_child3.LayoutHeight

    Assert.AreEqual 50!, root_child4.LayoutX
    Assert.AreEqual 50!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 50!, root_child4.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 150!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 100!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 50!, root_child2.LayoutHeight

    Assert.AreEqual 100!, root_child3.LayoutX
    Assert.AreEqual 50!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 50!, root_child3.LayoutHeight

    Assert.AreEqual 50!, root_child4.LayoutX
    Assert.AreEqual 50!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 50!, root_child4.LayoutHeight
End Sub

Public Sub Test_align_content_stretch_row_with_min_height()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignContent = YogaAlign.yogaAlignStretch
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 150
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 50
    root_child1.MinHeight = 80
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Width = 50
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(config)
    root_child3.Width = 50
    root.Insert 3, root_child3

    Dim root_child4 As cYogaNode: Set root_child4 = YogaNodeNew(config)
    root_child4.Width = 50
    root.Insert 4, root_child4
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 150!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 90!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 90!, root_child1.LayoutHeight

    Assert.AreEqual 100!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 90!, root_child2.LayoutHeight

    Assert.AreEqual 0!, root_child3.LayoutX
    Assert.AreEqual 90!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 10!, root_child3.LayoutHeight

    Assert.AreEqual 50!, root_child4.LayoutX
    Assert.AreEqual 90!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 10!, root_child4.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 150!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 100!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 90!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 90!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 90!, root_child2.LayoutHeight

    Assert.AreEqual 100!, root_child3.LayoutX
    Assert.AreEqual 90!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 10!, root_child3.LayoutHeight

    Assert.AreEqual 50!, root_child4.LayoutX
    Assert.AreEqual 90!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 10!, root_child4.LayoutHeight
End Sub

Public Sub Test_align_content_stretch_column()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.AlignContent = YogaAlign.yogaAlignStretch
    root.Wrap = YogaWrap.yogaWrapWrap
    root.Width = 100
    root.Height = 150

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(config)
    root_child0_child0.FlexGrow = 1
    root_child0_child0.FlexShrink = 1
    root_child0_child0.FlexBasis = Array(yogaUnitPercent, 0)
    root_child0.Insert 0, root_child0_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 1
    root_child1.FlexShrink = 1
    root_child1.FlexBasis = Array(yogaUnitPercent, 0)
    root_child1.Height = 50
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Height = 50
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(config)
    root_child3.Height = 50
    root.Insert 3, root_child3

    Dim root_child4 As cYogaNode: Set root_child4 = YogaNodeNew(config)
    root_child4.Height = 50
    root.Insert 4, root_child4
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 150!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 50!, root_child0_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 0!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 50!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 50!, root_child2.LayoutHeight

    Assert.AreEqual 0!, root_child3.LayoutX
    Assert.AreEqual 100!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 50!, root_child3.LayoutHeight

    Assert.AreEqual 50!, root_child4.LayoutX
    Assert.AreEqual 0!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 50!, root_child4.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 150!, root.LayoutHeight

    Assert.AreEqual 50!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 50!, root_child0_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 0!, root_child1.LayoutHeight

    Assert.AreEqual 50!, root_child2.LayoutX
    Assert.AreEqual 50!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 50!, root_child2.LayoutHeight

    Assert.AreEqual 50!, root_child3.LayoutX
    Assert.AreEqual 100!, root_child3.LayoutY
    Assert.AreEqual 50!, root_child3.LayoutWidth
    Assert.AreEqual 50!, root_child3.LayoutHeight

    Assert.AreEqual 0!, root_child4.LayoutX
    Assert.AreEqual 0!, root_child4.LayoutY
    Assert.AreEqual 50!, root_child4.LayoutWidth
    Assert.AreEqual 50!, root_child4.LayoutHeight
End Sub

Public Sub Test_align_content_stretch_is_not_overriding_align_items()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.AlignContent = YogaAlign.yogaAlignStretch

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.flexDirection = YogaFlexDirection.yogaFlexRow
    root_child0.AlignContent = YogaAlign.yogaAlignStretch
    root_child0.AlignItems = YogaAlign.yogaAlignCenter
    root_child0.Width = 100
    root_child0.Height = 100
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(config)
    root_child0_child0.AlignContent = YogaAlign.yogaAlignStretch
    root_child0_child0.Width = 10
    root_child0_child0.Height = 10
    root_child0.Insert 0, root_child0_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 45!, root_child0_child0.LayoutY
    Assert.AreEqual 10!, root_child0_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    Assert.AreEqual 90!, root_child0_child0.LayoutX
    Assert.AreEqual 45!, root_child0_child0.LayoutY
    Assert.AreEqual 10!, root_child0_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0_child0.LayoutHeight
End Sub


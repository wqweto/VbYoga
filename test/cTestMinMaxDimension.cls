VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTestMinMaxDimension"
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

Public Sub Test_max_width()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.MaxWidth = 50
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
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

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
End Sub

Public Sub Test_max_height()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 10
    root_child0.MaxHeight = 50
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
    Assert.AreEqual 50!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 90!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight
End Sub

Public Sub Test_min_height()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.MinHeight = 60
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 1
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 80!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 80!, root_child1.LayoutY
    Assert.AreEqual 100!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 80!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 80!, root_child1.LayoutY
    Assert.AreEqual 100!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight
End Sub

Public Sub Test_min_width()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.MinWidth = 60
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 1
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 80!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    Assert.AreEqual 80!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 20!, root_child1.LayoutWidth
    Assert.AreEqual 100!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 20!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 80!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 20!, root_child1.LayoutWidth
    Assert.AreEqual 100!, root_child1.LayoutHeight
End Sub

Public Sub Test_justify_content_min_max()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.JustifyContent = YogaJustify.yogaJustCenter
    root.Width = 100
    root.MinHeight = 100
    root.MaxHeight = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 60
    root_child0.Height = 60
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 20!, root_child0.LayoutY
    Assert.AreEqual 60!, root_child0.LayoutWidth
    Assert.AreEqual 60!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 40!, root_child0.LayoutX
    Assert.AreEqual 20!, root_child0.LayoutY
    Assert.AreEqual 60!, root_child0.LayoutWidth
    Assert.AreEqual 60!, root_child0.LayoutHeight
End Sub

Public Sub Test_align_items_min_max()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.MinWidth = 100
    root.MaxWidth = 200
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 60
    root_child0.Height = 60
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 20!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 60!, root_child0.LayoutWidth
    Assert.AreEqual 60!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 20!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 60!, root_child0.LayoutWidth
    Assert.AreEqual 60!, root_child0.LayoutHeight
End Sub

Public Sub Test_justify_content_overflow_min_max()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.JustifyContent = YogaJustify.yogaJustCenter
    root.MinHeight = 100
    root.MaxHeight = 110

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 50
    root_child1.Height = 50
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Width = 50
    root_child2.Height = 50
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 50!, root.LayoutWidth
    Assert.AreEqual 110!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual -20!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 30!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 80!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 50!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 50!, root.LayoutWidth
    Assert.AreEqual 110!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual -20!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 30!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 80!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 50!, root_child2.LayoutHeight
End Sub

Public Sub Test_flex_grow_to_min()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 100
    root.MinHeight = 100
    root.MaxHeight = 500

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.FlexShrink = 1
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
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
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 100!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 100!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight
End Sub

Public Sub Test_flex_grow_in_at_most_container()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignItems = YogaAlign.yogaAlignFlexStart
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.flexDirection = YogaFlexDirection.yogaFlexRow
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(config)
    root_child0_child0.FlexGrow = 1
    root_child0_child0.FlexBasis = 0
    root_child0.Insert 0, root_child0_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 0!, root_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 0!, root_child0_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 100!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 0!, root_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 0!, root_child0_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0_child0.LayoutHeight
End Sub

Public Sub Test_flex_grow_child()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.FlexBasis = 0
    root_child0.Height = 100
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 0!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 0!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 0!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 0!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight
End Sub

Public Sub Test_flex_grow_within_constrained_min_max_column()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.MinHeight = 100
    root.MaxHeight = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Height = 50
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 0!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 0!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 0!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 0!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 0!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 0!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight
End Sub

Public Sub Test_flex_grow_within_max_width()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 200
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.flexDirection = YogaFlexDirection.yogaFlexRow
    root_child0.MaxWidth = 100
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(config)
    root_child0_child0.FlexGrow = 1
    root_child0_child0.Height = 20
    root_child0.Insert 0, root_child0_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 100!, root_child0_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 100!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 100!, root_child0_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0_child0.LayoutHeight
End Sub

Public Sub Test_flex_grow_within_constrained_max_width()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 200
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.flexDirection = YogaFlexDirection.yogaFlexRow
    root_child0.MaxWidth = 300
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(config)
    root_child0_child0.FlexGrow = 1
    root_child0_child0.Height = 20
    root_child0.Insert 0, root_child0_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 200!, root_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 200!, root_child0_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 200!, root_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 200!, root_child0_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0_child0.LayoutHeight
End Sub

Public Sub Test_flex_root_ignored()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.FlexGrow = 1
    root.Width = 100
    root.MinHeight = 100
    root.MaxHeight = 500

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.FlexBasis = 200
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Height = 100
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 300!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 200!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 200!, root_child1.LayoutY
    Assert.AreEqual 100!, root_child1.LayoutWidth
    Assert.AreEqual 100!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 300!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 200!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 200!, root_child1.LayoutY
    Assert.AreEqual 100!, root_child1.LayoutWidth
    Assert.AreEqual 100!, root_child1.LayoutHeight
End Sub

Public Sub Test_flex_grow_root_minimized()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 100
    root.MinHeight = 100
    root.MaxHeight = 500

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.MinHeight = 100
    root_child0.MaxHeight = 500
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(config)
    root_child0_child0.FlexGrow = 1
    root_child0_child0.FlexBasis = 200
    root_child0.Insert 0, root_child0_child0

    Dim root_child0_child1 As cYogaNode: Set root_child0_child1 = YogaNodeNew(config)
    root_child0_child1.Height = 100
    root_child0.Insert 1, root_child0_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 300!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 300!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 100!, root_child0_child0.LayoutWidth
    Assert.AreEqual 200!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child1.LayoutX
    Assert.AreEqual 200!, root_child0_child1.LayoutY
    Assert.AreEqual 100!, root_child0_child1.LayoutWidth
    Assert.AreEqual 100!, root_child0_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 300!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 300!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 100!, root_child0_child0.LayoutWidth
    Assert.AreEqual 200!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child1.LayoutX
    Assert.AreEqual 200!, root_child0_child1.LayoutY
    Assert.AreEqual 100!, root_child0_child1.LayoutWidth
    Assert.AreEqual 100!, root_child0_child1.LayoutHeight
End Sub

Public Sub Test_flex_grow_height_maximized()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 100
    root.Height = 500

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.MinHeight = 100
    root_child0.MaxHeight = 500
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(config)
    root_child0_child0.FlexGrow = 1
    root_child0_child0.FlexBasis = 200
    root_child0.Insert 0, root_child0_child0

    Dim root_child0_child1 As cYogaNode: Set root_child0_child1 = YogaNodeNew(config)
    root_child0_child1.Height = 100
    root_child0.Insert 1, root_child0_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 500!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 500!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 100!, root_child0_child0.LayoutWidth
    Assert.AreEqual 400!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child1.LayoutX
    Assert.AreEqual 400!, root_child0_child1.LayoutY
    Assert.AreEqual 100!, root_child0_child1.LayoutWidth
    Assert.AreEqual 100!, root_child0_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 500!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 500!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 100!, root_child0_child0.LayoutWidth
    Assert.AreEqual 400!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child1.LayoutX
    Assert.AreEqual 400!, root_child0_child1.LayoutY
    Assert.AreEqual 100!, root_child0_child1.LayoutWidth
    Assert.AreEqual 100!, root_child0_child1.LayoutHeight
End Sub

Public Sub Test_flex_grow_within_constrained_min_row()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.MinWidth = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 50
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
    Assert.AreEqual 100!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 100!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 50!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 100!, root_child1.LayoutHeight
End Sub

Public Sub Test_flex_grow_within_constrained_min_column()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.MinHeight = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Height = 50
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 0!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 0!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 0!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 0!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 0!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 0!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight
End Sub

Public Sub Test_flex_grow_within_constrained_max_row()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.flexDirection = YogaFlexDirection.yogaFlexRow
    root_child0.MaxWidth = 100
    root_child0.Height = 100
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(config)
    root_child0_child0.FlexShrink = 1
    root_child0_child0.FlexBasis = 100
    root_child0.Insert 0, root_child0_child0

    Dim root_child0_child1 As cYogaNode: Set root_child0_child1 = YogaNodeNew(config)
    root_child0_child1.Width = 50
    root_child0.Insert 1, root_child0_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 50!, root_child0_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0_child0.LayoutHeight

    Assert.AreEqual 50!, root_child0_child1.LayoutX
    Assert.AreEqual 0!, root_child0_child1.LayoutY
    Assert.AreEqual 50!, root_child0_child1.LayoutWidth
    Assert.AreEqual 100!, root_child0_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 100!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 50!, root_child0_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child1.LayoutX
    Assert.AreEqual 0!, root_child0_child1.LayoutY
    Assert.AreEqual 50!, root_child0_child1.LayoutWidth
    Assert.AreEqual 100!, root_child0_child1.LayoutHeight
End Sub

Public Sub Test_flex_grow_within_constrained_max_column()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 100
    root.MaxHeight = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexShrink = 1
    root_child0.FlexBasis = 100
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
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
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 100!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 100!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight
End Sub

Public Sub Test_child_min_max_width_flexing()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.Width = 120
    root.Height = 50

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.FlexBasis = 0
    root_child0.MinWidth = 60
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 1
    root_child1.FlexBasis = Array(yogaUnitPercent, 50)
    root_child1.MaxWidth = 20
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 120!, root.LayoutWidth
    Assert.AreEqual 50!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 100!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 20!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 120!, root.LayoutWidth
    Assert.AreEqual 50!, root.LayoutHeight

    Assert.AreEqual 20!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 20!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight
End Sub

Public Sub Test_min_width_overrides_width()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 50
    root.MinWidth = 100
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 0!, root.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 0!, root.LayoutHeight
End Sub

Public Sub Test_max_width_overrides_width()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 200
    root.MaxWidth = 100
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 0!, root.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 0!, root.LayoutHeight
End Sub

Public Sub Test_min_height_overrides_height()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Height = 50
    root.MinHeight = 100
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 0!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 0!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight
End Sub

Public Sub Test_max_height_overrides_height()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Height = 200
    root.MaxHeight = 100
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 0!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 0!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight
End Sub

Public Sub Test_min_max_percent_no_width_height()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.AlignItems = YogaAlign.yogaAlignFlexStart
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.MinWidth = Array(yogaUnitPercent, 10)
    root_child0.MaxWidth = Array(yogaUnitPercent, 10)
    root_child0.MinHeight = Array(yogaUnitPercent, 10)
    root_child0.MaxHeight = Array(yogaUnitPercent, 10)
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

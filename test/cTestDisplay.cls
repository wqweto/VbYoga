VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTestDisplay"
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

Public Sub Test_display_none()
            Dim config As cYogaConfig: Set config = YogaConfigNew()

            Dim root As cYogaNode: Set root = YogaNodeNew(config)
            root.flexDirection = YogaFlexDirection.yogaFlexRow
            root.Width = 100
            root.Height = 100

            Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
            root_child0.FlexGrow = 1
            root.Insert 0, root_child0

            Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
            root_child1.FlexGrow = 1
            root_child1.Display = YogaDisplay.yogaDisplayNone
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
            Assert.AreEqual 100!, root_child0.LayoutHeight

            Assert.AreEqual 0!, root_child1.LayoutX
            Assert.AreEqual 0!, root_child1.LayoutY
            Assert.AreEqual 0!, root_child1.LayoutWidth
            Assert.AreEqual 0!, root_child1.LayoutHeight

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

            Assert.AreEqual 0!, root_child1.LayoutX
            Assert.AreEqual 0!, root_child1.LayoutY
            Assert.AreEqual 0!, root_child1.LayoutWidth
            Assert.AreEqual 0!, root_child1.LayoutHeight
End Sub

Public Sub Test_display_none_fixed_size()
            Dim config As cYogaConfig: Set config = YogaConfigNew()

            Dim root As cYogaNode: Set root = YogaNodeNew(config)
            root.flexDirection = YogaFlexDirection.yogaFlexRow
            root.Width = 100
            root.Height = 100

            Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
            root_child0.FlexGrow = 1
            root.Insert 0, root_child0

            Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
            root_child1.Width = 20
            root_child1.Height = 20
            root_child1.Display = YogaDisplay.yogaDisplayNone
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
            Assert.AreEqual 100!, root_child0.LayoutHeight

            Assert.AreEqual 0!, root_child1.LayoutX
            Assert.AreEqual 0!, root_child1.LayoutY
            Assert.AreEqual 0!, root_child1.LayoutWidth
            Assert.AreEqual 0!, root_child1.LayoutHeight

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

            Assert.AreEqual 0!, root_child1.LayoutX
            Assert.AreEqual 0!, root_child1.LayoutY
            Assert.AreEqual 0!, root_child1.LayoutWidth
            Assert.AreEqual 0!, root_child1.LayoutHeight
End Sub

Public Sub Test_display_none_with_margin()
            Dim config As cYogaConfig: Set config = YogaConfigNew()

            Dim root As cYogaNode: Set root = YogaNodeNew(config)
            root.flexDirection = YogaFlexDirection.yogaFlexRow
            root.Width = 100
            root.Height = 100

            Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
            root_child0.MarginLeft = 10
            root_child0.MarginTop = 10
            root_child0.MarginRight = 10
            root_child0.MarginBottom = 10
            root_child0.Width = 20
            root_child0.Height = 20
            root_child0.Display = YogaDisplay.yogaDisplayNone
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
            Assert.AreEqual 0!, root_child0.LayoutWidth
            Assert.AreEqual 0!, root_child0.LayoutHeight

            Assert.AreEqual 0!, root_child1.LayoutX
            Assert.AreEqual 0!, root_child1.LayoutY
            Assert.AreEqual 100!, root_child1.LayoutWidth
            Assert.AreEqual 100!, root_child1.LayoutHeight

            root.StyleDirection = YogaDirection.yogaDirRTL
            root.CalculateLayout

            Assert.AreEqual 0!, root.LayoutX
            Assert.AreEqual 0!, root.LayoutY
            Assert.AreEqual 100!, root.LayoutWidth
            Assert.AreEqual 100!, root.LayoutHeight

            Assert.AreEqual 0!, root_child0.LayoutX
            Assert.AreEqual 0!, root_child0.LayoutY
            Assert.AreEqual 0!, root_child0.LayoutWidth
            Assert.AreEqual 0!, root_child0.LayoutHeight

            Assert.AreEqual 0!, root_child1.LayoutX
            Assert.AreEqual 0!, root_child1.LayoutY
            Assert.AreEqual 100!, root_child1.LayoutWidth
            Assert.AreEqual 100!, root_child1.LayoutHeight
End Sub

Public Sub Test_display_none_with_child()
            Dim config As cYogaConfig: Set config = YogaConfigNew()

            Dim root As cYogaNode: Set root = YogaNodeNew(config)
            root.flexDirection = YogaFlexDirection.yogaFlexRow
            root.Width = 100
            root.Height = 100

            Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
            root_child0.FlexGrow = 1
            root_child0.FlexShrink = 1
            root_child0.FlexBasis = Array(yogaUnitPercent, 0)
            root.Insert 0, root_child0

            Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
            root_child1.FlexGrow = 1
            root_child1.FlexShrink = 1
            root_child1.FlexBasis = Array(yogaUnitPercent, 0)
            root_child1.Display = YogaDisplay.yogaDisplayNone
            root.Insert 1, root_child1

            Dim root_child1_child0 As cYogaNode: Set root_child1_child0 = YogaNodeNew(config)
            root_child1_child0.FlexGrow = 1
            root_child1_child0.FlexShrink = 1
            root_child1_child0.FlexBasis = Array(yogaUnitPercent, 0)
            root_child1_child0.Width = 20
            root_child1_child0.MinWidth = 0
            root_child1_child0.MinHeight = 0
            root_child1.Insert 0, root_child1_child0

            Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
            root_child2.FlexGrow = 1
            root_child2.FlexShrink = 1
            root_child2.FlexBasis = Array(yogaUnitPercent, 0)
            root.Insert 2, root_child2
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

            Assert.AreEqual 0!, root_child1.LayoutX
            Assert.AreEqual 0!, root_child1.LayoutY
            Assert.AreEqual 0!, root_child1.LayoutWidth
            Assert.AreEqual 0!, root_child1.LayoutHeight

            Assert.AreEqual 0!, root_child1_child0.LayoutX
            Assert.AreEqual 0!, root_child1_child0.LayoutY
            Assert.AreEqual 0!, root_child1_child0.LayoutWidth
            Assert.AreEqual 0!, root_child1_child0.LayoutHeight

            Assert.AreEqual 50!, root_child2.LayoutX
            Assert.AreEqual 0!, root_child2.LayoutY
            Assert.AreEqual 50!, root_child2.LayoutWidth
            Assert.AreEqual 100!, root_child2.LayoutHeight

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
            Assert.AreEqual 0!, root_child1.LayoutWidth
            Assert.AreEqual 0!, root_child1.LayoutHeight

            Assert.AreEqual 0!, root_child1_child0.LayoutX
            Assert.AreEqual 0!, root_child1_child0.LayoutY
            Assert.AreEqual 0!, root_child1_child0.LayoutWidth
            Assert.AreEqual 0!, root_child1_child0.LayoutHeight

            Assert.AreEqual 0!, root_child2.LayoutX
            Assert.AreEqual 0!, root_child2.LayoutY
            Assert.AreEqual 50!, root_child2.LayoutWidth
            Assert.AreEqual 100!, root_child2.LayoutHeight
End Sub

Public Sub Test_display_none_with_position()
            Dim config As cYogaConfig: Set config = YogaConfigNew()

            Dim root As cYogaNode: Set root = YogaNodeNew(config)
            root.flexDirection = YogaFlexDirection.yogaFlexRow
            root.Width = 100
            root.Height = 100

            Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
            root_child0.FlexGrow = 1
            root.Insert 0, root_child0

            Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
            root_child1.FlexGrow = 1
            root_child1.Top = 10
            root_child1.Display = YogaDisplay.yogaDisplayNone
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
            Assert.AreEqual 100!, root_child0.LayoutHeight

            Assert.AreEqual 0!, root_child1.LayoutX
            Assert.AreEqual 0!, root_child1.LayoutY
            Assert.AreEqual 0!, root_child1.LayoutWidth
            Assert.AreEqual 0!, root_child1.LayoutHeight

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

            Assert.AreEqual 0!, root_child1.LayoutX
            Assert.AreEqual 0!, root_child1.LayoutY
            Assert.AreEqual 0!, root_child1.LayoutWidth
            Assert.AreEqual 0!, root_child1.LayoutHeight
End Sub


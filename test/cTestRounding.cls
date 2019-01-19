VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTestRounding"
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

Public Sub Test_rounding_flex_basis_flex_grow_row_width_of_100()
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
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.FlexGrow = 1
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 33!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    Assert.AreEqual 33!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 34!, root_child1.LayoutWidth
    Assert.AreEqual 100!, root_child1.LayoutHeight

    Assert.AreEqual 67!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 33!, root_child2.LayoutWidth
    Assert.AreEqual 100!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 67!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 33!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    Assert.AreEqual 33!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 34!, root_child1.LayoutWidth
    Assert.AreEqual 100!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 33!, root_child2.LayoutWidth
    Assert.AreEqual 100!, root_child2.LayoutHeight
End Sub

Public Sub Test_rounding_flex_basis_flex_grow_row_prime_number_width()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.Width = 113
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 1
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.FlexGrow = 1
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(config)
    root_child3.FlexGrow = 1
    root.Insert 3, root_child3

    Dim root_child4 As cYogaNode: Set root_child4 = YogaNodeNew(config)
    root_child4.FlexGrow = 1
    root.Insert 4, root_child4
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 113!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 23!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    Assert.AreEqual 23!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 22!, root_child1.LayoutWidth
    Assert.AreEqual 100!, root_child1.LayoutHeight

    Assert.AreEqual 45!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 23!, root_child2.LayoutWidth
    Assert.AreEqual 100!, root_child2.LayoutHeight

    Assert.AreEqual 68!, root_child3.LayoutX
    Assert.AreEqual 0!, root_child3.LayoutY
    Assert.AreEqual 22!, root_child3.LayoutWidth
    Assert.AreEqual 100!, root_child3.LayoutHeight

    Assert.AreEqual 90!, root_child4.LayoutX
    Assert.AreEqual 0!, root_child4.LayoutY
    Assert.AreEqual 23!, root_child4.LayoutWidth
    Assert.AreEqual 100!, root_child4.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 113!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 90!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 23!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    Assert.AreEqual 68!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 22!, root_child1.LayoutWidth
    Assert.AreEqual 100!, root_child1.LayoutHeight

    Assert.AreEqual 45!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 23!, root_child2.LayoutWidth
    Assert.AreEqual 100!, root_child2.LayoutHeight

    Assert.AreEqual 23!, root_child3.LayoutX
    Assert.AreEqual 0!, root_child3.LayoutY
    Assert.AreEqual 22!, root_child3.LayoutWidth
    Assert.AreEqual 100!, root_child3.LayoutHeight

    Assert.AreEqual 0!, root_child4.LayoutX
    Assert.AreEqual 0!, root_child4.LayoutY
    Assert.AreEqual 23!, root_child4.LayoutWidth
    Assert.AreEqual 100!, root_child4.LayoutHeight
End Sub

Public Sub Test_rounding_flex_basis_flex_shrink_row()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.Width = 101
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexShrink = 1
    root_child0.FlexBasis = 100
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexBasis = 25
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.FlexBasis = 25
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 101!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 51!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    Assert.AreEqual 51!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 25!, root_child1.LayoutWidth
    Assert.AreEqual 100!, root_child1.LayoutHeight

    Assert.AreEqual 76!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 25!, root_child2.LayoutWidth
    Assert.AreEqual 100!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 101!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 50!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 51!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    Assert.AreEqual 25!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 25!, root_child1.LayoutWidth
    Assert.AreEqual 100!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 25!, root_child2.LayoutWidth
    Assert.AreEqual 100!, root_child2.LayoutHeight
End Sub

Public Sub Test_rounding_flex_basis_overrides_main_size()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 100
    root.Height = 113

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.FlexBasis = 50
    root_child0.Height = 20
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 1
    root_child1.Height = 10
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.FlexGrow = 1
    root_child2.Height = 10
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 113!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 64!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 64!, root_child1.LayoutY
    Assert.AreEqual 100!, root_child1.LayoutWidth
    Assert.AreEqual 25!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 89!, root_child2.LayoutY
    Assert.AreEqual 100!, root_child2.LayoutWidth
    Assert.AreEqual 24!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 113!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 64!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 64!, root_child1.LayoutY
    Assert.AreEqual 100!, root_child1.LayoutWidth
    Assert.AreEqual 25!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 89!, root_child2.LayoutY
    Assert.AreEqual 100!, root_child2.LayoutWidth
    Assert.AreEqual 24!, root_child2.LayoutHeight
End Sub

Public Sub Test_rounding_total_fractial()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 87.4!
    root.Height = 113.4!

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 0.7!
    root_child0.FlexBasis = 50.3!
    root_child0.Height = 20.3!
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 1.6!
    root_child1.Height = 10
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.FlexGrow = 1.1!
    root_child2.Height = 10.7!
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 87!, root.LayoutWidth
    Assert.AreEqual 113!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 87!, root_child0.LayoutWidth
    Assert.AreEqual 59!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 59!, root_child1.LayoutY
    Assert.AreEqual 87!, root_child1.LayoutWidth
    Assert.AreEqual 30!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 89!, root_child2.LayoutY
    Assert.AreEqual 87!, root_child2.LayoutWidth
    Assert.AreEqual 24!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 87!, root.LayoutWidth
    Assert.AreEqual 113!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 87!, root_child0.LayoutWidth
    Assert.AreEqual 59!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 59!, root_child1.LayoutY
    Assert.AreEqual 87!, root_child1.LayoutWidth
    Assert.AreEqual 30!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 89!, root_child2.LayoutY
    Assert.AreEqual 87!, root_child2.LayoutWidth
    Assert.AreEqual 24!, root_child2.LayoutHeight
End Sub

Public Sub Test_rounding_total_fractial_nested()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 87.4!
    root.Height = 113.4!

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 0.7!
    root_child0.FlexBasis = 50.3!
    root_child0.Height = 20.3!
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(config)
    root_child0_child0.FlexGrow = 1
    root_child0_child0.FlexBasis = 0.3!
    root_child0_child0.Bottom = 13.3!
    root_child0_child0.Height = 9.9!
    root_child0.Insert 0, root_child0_child0

    Dim root_child0_child1 As cYogaNode: Set root_child0_child1 = YogaNodeNew(config)
    root_child0_child1.FlexGrow = 4
    root_child0_child1.FlexBasis = 0.3!
    root_child0_child1.Top = 13.3!
    root_child0_child1.Height = 1.1!
    root_child0.Insert 1, root_child0_child1

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 1.6!
    root_child1.Height = 10
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.FlexGrow = 1.1!
    root_child2.Height = 10.7!
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 87!, root.LayoutWidth
    Assert.AreEqual 113!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 87!, root_child0.LayoutWidth
    Assert.AreEqual 59!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual -13!, root_child0_child0.LayoutY
    Assert.AreEqual 87!, root_child0_child0.LayoutWidth
    Assert.AreEqual 12!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child1.LayoutX
    Assert.AreEqual 25!, root_child0_child1.LayoutY
    Assert.AreEqual 87!, root_child0_child1.LayoutWidth
    Assert.AreEqual 47!, root_child0_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 59!, root_child1.LayoutY
    Assert.AreEqual 87!, root_child1.LayoutWidth
    Assert.AreEqual 30!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 89!, root_child2.LayoutY
    Assert.AreEqual 87!, root_child2.LayoutWidth
    Assert.AreEqual 24!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 87!, root.LayoutWidth
    Assert.AreEqual 113!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 87!, root_child0.LayoutWidth
    Assert.AreEqual 59!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual -13!, root_child0_child0.LayoutY
    Assert.AreEqual 87!, root_child0_child0.LayoutWidth
    Assert.AreEqual 12!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child1.LayoutX
    Assert.AreEqual 25!, root_child0_child1.LayoutY
    Assert.AreEqual 87!, root_child0_child1.LayoutWidth
    Assert.AreEqual 47!, root_child0_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 59!, root_child1.LayoutY
    Assert.AreEqual 87!, root_child1.LayoutWidth
    Assert.AreEqual 30!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 89!, root_child2.LayoutY
    Assert.AreEqual 87!, root_child2.LayoutWidth
    Assert.AreEqual 24!, root_child2.LayoutHeight
End Sub

Public Sub Test_rounding_fractial_input_1()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 100
    root.Height = 113.4!

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.FlexBasis = 50
    root_child0.Height = 20
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 1
    root_child1.Height = 10
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.FlexGrow = 1
    root_child2.Height = 10
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 113!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 64!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 64!, root_child1.LayoutY
    Assert.AreEqual 100!, root_child1.LayoutWidth
    Assert.AreEqual 25!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 89!, root_child2.LayoutY
    Assert.AreEqual 100!, root_child2.LayoutWidth
    Assert.AreEqual 24!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 113!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 64!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 64!, root_child1.LayoutY
    Assert.AreEqual 100!, root_child1.LayoutWidth
    Assert.AreEqual 25!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 89!, root_child2.LayoutY
    Assert.AreEqual 100!, root_child2.LayoutWidth
    Assert.AreEqual 24!, root_child2.LayoutHeight
End Sub

Public Sub Test_rounding_fractial_input_2()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 100
    root.Height = 113.6!

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.FlexBasis = 50
    root_child0.Height = 20
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 1
    root_child1.Height = 10
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.FlexGrow = 1
    root_child2.Height = 10
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 114!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 65!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 65!, root_child1.LayoutY
    Assert.AreEqual 100!, root_child1.LayoutWidth
    Assert.AreEqual 24!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 89!, root_child2.LayoutY
    Assert.AreEqual 100!, root_child2.LayoutWidth
    Assert.AreEqual 25!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 114!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 65!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 65!, root_child1.LayoutY
    Assert.AreEqual 100!, root_child1.LayoutWidth
    Assert.AreEqual 24!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 89!, root_child2.LayoutY
    Assert.AreEqual 100!, root_child2.LayoutWidth
    Assert.AreEqual 25!, root_child2.LayoutHeight
End Sub

Public Sub Test_rounding_fractial_input_3()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Top = 0.3!
    root.Width = 100
    root.Height = 113.4!

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.FlexBasis = 50
    root_child0.Height = 20
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 1
    root_child1.Height = 10
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.FlexGrow = 1
    root_child2.Height = 10
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 114!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 65!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 64!, root_child1.LayoutY
    Assert.AreEqual 100!, root_child1.LayoutWidth
    Assert.AreEqual 24!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 89!, root_child2.LayoutY
    Assert.AreEqual 100!, root_child2.LayoutWidth
    Assert.AreEqual 25!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 114!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 65!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 64!, root_child1.LayoutY
    Assert.AreEqual 100!, root_child1.LayoutWidth
    Assert.AreEqual 24!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 89!, root_child2.LayoutY
    Assert.AreEqual 100!, root_child2.LayoutWidth
    Assert.AreEqual 25!, root_child2.LayoutHeight
End Sub

Public Sub Test_rounding_fractial_input_4()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Top = 0.7!
    root.Width = 100
    root.Height = 113.4!

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.FlexBasis = 50
    root_child0.Height = 20
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 1
    root_child1.Height = 10
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.FlexGrow = 1
    root_child2.Height = 10
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 1!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 113!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 64!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 64!, root_child1.LayoutY
    Assert.AreEqual 100!, root_child1.LayoutWidth
    Assert.AreEqual 25!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 89!, root_child2.LayoutY
    Assert.AreEqual 100!, root_child2.LayoutWidth
    Assert.AreEqual 24!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 1!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 113!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 64!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 64!, root_child1.LayoutY
    Assert.AreEqual 100!, root_child1.LayoutWidth
    Assert.AreEqual 25!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 89!, root_child2.LayoutY
    Assert.AreEqual 100!, root_child2.LayoutWidth
    Assert.AreEqual 24!, root_child2.LayoutHeight
End Sub

Public Sub Test_rounding_inner_node_controversy_horizontal()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.Width = 320

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.Height = 10
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 1
    root_child1.Height = 10
    root.Insert 1, root_child1

    Dim root_child1_child0 As cYogaNode: Set root_child1_child0 = YogaNodeNew(config)
    root_child1_child0.FlexGrow = 1
    root_child1_child0.Height = 10
    root_child1.Insert 0, root_child1_child0

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.FlexGrow = 1
    root_child2.Height = 10
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 320!, root.LayoutWidth
    Assert.AreEqual 10!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 107!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 107!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 106!, root_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 106!, root_child1_child0.LayoutWidth
    Assert.AreEqual 10!, root_child1_child0.LayoutHeight

    Assert.AreEqual 213!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 107!, root_child2.LayoutWidth
    Assert.AreEqual 10!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 320!, root.LayoutWidth
    Assert.AreEqual 10!, root.LayoutHeight

    Assert.AreEqual 213!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 107!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 107!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 106!, root_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 106!, root_child1_child0.LayoutWidth
    Assert.AreEqual 10!, root_child1_child0.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 107!, root_child2.LayoutWidth
    Assert.AreEqual 10!, root_child2.LayoutHeight
End Sub

Public Sub Test_rounding_inner_node_controversy_vertical()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Height = 320

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.Width = 10
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 1
    root_child1.Width = 10
    root.Insert 1, root_child1

    Dim root_child1_child0 As cYogaNode: Set root_child1_child0 = YogaNodeNew(config)
    root_child1_child0.FlexGrow = 1
    root_child1_child0.Width = 10
    root_child1.Insert 0, root_child1_child0

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.FlexGrow = 1
    root_child2.Width = 10
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 10!, root.LayoutWidth
    Assert.AreEqual 320!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 107!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 107!, root_child1.LayoutY
    Assert.AreEqual 10!, root_child1.LayoutWidth
    Assert.AreEqual 106!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 10!, root_child1_child0.LayoutWidth
    Assert.AreEqual 106!, root_child1_child0.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 213!, root_child2.LayoutY
    Assert.AreEqual 10!, root_child2.LayoutWidth
    Assert.AreEqual 107!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 10!, root.LayoutWidth
    Assert.AreEqual 320!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 107!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 107!, root_child1.LayoutY
    Assert.AreEqual 10!, root_child1.LayoutWidth
    Assert.AreEqual 106!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 10!, root_child1_child0.LayoutWidth
    Assert.AreEqual 106!, root_child1_child0.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 213!, root_child2.LayoutY
    Assert.AreEqual 10!, root_child2.LayoutWidth
    Assert.AreEqual 107!, root_child2.LayoutHeight
End Sub

Public Sub Test_rounding_inner_node_controversy_combined()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.Width = 640
    root.Height = 320

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.Height = Array(yogaUnitPercent, 100)
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 1
    root_child1.Height = Array(yogaUnitPercent, 100)
    root.Insert 1, root_child1

    Dim root_child1_child0 As cYogaNode: Set root_child1_child0 = YogaNodeNew(config)
    root_child1_child0.FlexGrow = 1
    root_child1_child0.Width = Array(yogaUnitPercent, 100)
    root_child1.Insert 0, root_child1_child0

    Dim root_child1_child1 As cYogaNode: Set root_child1_child1 = YogaNodeNew(config)
    root_child1_child1.FlexGrow = 1
    root_child1_child1.Width = Array(yogaUnitPercent, 100)
    root_child1.Insert 1, root_child1_child1

    Dim root_child1_child1_child0 As cYogaNode: Set root_child1_child1_child0 = YogaNodeNew(config)
    root_child1_child1_child0.FlexGrow = 1
    root_child1_child1_child0.Width = Array(yogaUnitPercent, 100)
    root_child1_child1.Insert 0, root_child1_child1_child0

    Dim root_child1_child2 As cYogaNode: Set root_child1_child2 = YogaNodeNew(config)
    root_child1_child2.FlexGrow = 1
    root_child1_child2.Width = Array(yogaUnitPercent, 100)
    root_child1.Insert 2, root_child1_child2

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.FlexGrow = 1
    root_child2.Height = Array(yogaUnitPercent, 100)
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 640!, root.LayoutWidth
    Assert.AreEqual 320!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 213!, root_child0.LayoutWidth
    Assert.AreEqual 320!, root_child0.LayoutHeight

    Assert.AreEqual 213!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 214!, root_child1.LayoutWidth
    Assert.AreEqual 320!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 214!, root_child1_child0.LayoutWidth
    Assert.AreEqual 107!, root_child1_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1_child1.LayoutX
    Assert.AreEqual 107!, root_child1_child1.LayoutY
    Assert.AreEqual 214!, root_child1_child1.LayoutWidth
    Assert.AreEqual 106!, root_child1_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child1_child0.LayoutY
    Assert.AreEqual 214!, root_child1_child1_child0.LayoutWidth
    Assert.AreEqual 106!, root_child1_child1_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1_child2.LayoutX
    Assert.AreEqual 213!, root_child1_child2.LayoutY
    Assert.AreEqual 214!, root_child1_child2.LayoutWidth
    Assert.AreEqual 107!, root_child1_child2.LayoutHeight

    Assert.AreEqual 427!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 213!, root_child2.LayoutWidth
    Assert.AreEqual 320!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 640!, root.LayoutWidth
    Assert.AreEqual 320!, root.LayoutHeight

    Assert.AreEqual 427!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 213!, root_child0.LayoutWidth
    Assert.AreEqual 320!, root_child0.LayoutHeight

    Assert.AreEqual 213!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 214!, root_child1.LayoutWidth
    Assert.AreEqual 320!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 214!, root_child1_child0.LayoutWidth
    Assert.AreEqual 107!, root_child1_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1_child1.LayoutX
    Assert.AreEqual 107!, root_child1_child1.LayoutY
    Assert.AreEqual 214!, root_child1_child1.LayoutWidth
    Assert.AreEqual 106!, root_child1_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child1_child0.LayoutY
    Assert.AreEqual 214!, root_child1_child1_child0.LayoutWidth
    Assert.AreEqual 106!, root_child1_child1_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1_child2.LayoutX
    Assert.AreEqual 213!, root_child1_child2.LayoutY
    Assert.AreEqual 214!, root_child1_child2.LayoutWidth
    Assert.AreEqual 107!, root_child1_child2.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 213!, root_child2.LayoutWidth
    Assert.AreEqual 320!, root_child2.LayoutHeight
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTestPercentage"
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

Public Sub Test_percentage_width_height()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = Array(yogaUnitPercent, 30)
    root_child0.Height = Array(yogaUnitPercent, 30)
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 60!, root_child0.LayoutWidth
    Assert.AreEqual 60!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 140!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 60!, root_child0.LayoutWidth
    Assert.AreEqual 60!, root_child0.LayoutHeight
End Sub

Public Sub Test_percentage_position_left_top()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.Width = 400
    root.Height = 400

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Left = Array(yogaUnitPercent, 10)
    root_child0.Top = Array(yogaUnitPercent, 20)
    root_child0.Width = Array(yogaUnitPercent, 45)
    root_child0.Height = Array(yogaUnitPercent, 55)
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 400!, root.LayoutWidth
    Assert.AreEqual 400!, root.LayoutHeight

    Assert.AreEqual 40!, root_child0.LayoutX
    Assert.AreEqual 80!, root_child0.LayoutY
    Assert.AreEqual 180!, root_child0.LayoutWidth
    Assert.AreEqual 220!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 400!, root.LayoutWidth
    Assert.AreEqual 400!, root.LayoutHeight

    Assert.AreEqual 260!, root_child0.LayoutX
    Assert.AreEqual 80!, root_child0.LayoutY
    Assert.AreEqual 180!, root_child0.LayoutWidth
    Assert.AreEqual 220!, root_child0.LayoutHeight
End Sub

Public Sub Test_percentage_position_bottom_right()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.Width = 500
    root.Height = 500

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Right = Array(yogaUnitPercent, 20)
    root_child0.Bottom = Array(yogaUnitPercent, 10)
    root_child0.Width = Array(yogaUnitPercent, 55)
    root_child0.Height = Array(yogaUnitPercent, 15)
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 500!, root.LayoutWidth
    Assert.AreEqual 500!, root.LayoutHeight

    Assert.AreEqual -100!, root_child0.LayoutX
    Assert.AreEqual -50!, root_child0.LayoutY
    Assert.AreEqual 275!, root_child0.LayoutWidth
    Assert.AreEqual 75!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 500!, root.LayoutWidth
    Assert.AreEqual 500!, root.LayoutHeight

    Assert.AreEqual 125!, root_child0.LayoutX
    Assert.AreEqual -50!, root_child0.LayoutY
    Assert.AreEqual 275!, root_child0.LayoutWidth
    Assert.AreEqual 75!, root_child0.LayoutHeight
End Sub

Public Sub Test_percentage_flex_basis()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.FlexBasis = Array(yogaUnitPercent, 50)
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 1
    root_child1.FlexBasis = Array(yogaUnitPercent, 25)
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 125!, root_child0.LayoutWidth
    Assert.AreEqual 200!, root_child0.LayoutHeight

    Assert.AreEqual 125!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 75!, root_child1.LayoutWidth
    Assert.AreEqual 200!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 75!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 125!, root_child0.LayoutWidth
    Assert.AreEqual 200!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 75!, root_child1.LayoutWidth
    Assert.AreEqual 200!, root_child1.LayoutHeight
End Sub

Public Sub Test_percentage_flex_basis_cross()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.FlexBasis = Array(yogaUnitPercent, 50)
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 1
    root_child1.FlexBasis = Array(yogaUnitPercent, 25)
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 200!, root_child0.LayoutWidth
    Assert.AreEqual 125!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 125!, root_child1.LayoutY
    Assert.AreEqual 200!, root_child1.LayoutWidth
    Assert.AreEqual 75!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 200!, root_child0.LayoutWidth
    Assert.AreEqual 125!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 125!, root_child1.LayoutY
    Assert.AreEqual 200!, root_child1.LayoutWidth
    Assert.AreEqual 75!, root_child1.LayoutHeight
End Sub

Public Sub Test_percentage_flex_basis_cross_min_height()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.MinHeight = Array(yogaUnitPercent, 60)
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 2
    root_child1.MinHeight = Array(yogaUnitPercent, 10)
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 200!, root_child0.LayoutWidth
    Assert.AreEqual 140!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 140!, root_child1.LayoutY
    Assert.AreEqual 200!, root_child1.LayoutWidth
    Assert.AreEqual 60!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 200!, root_child0.LayoutWidth
    Assert.AreEqual 140!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 140!, root_child1.LayoutY
    Assert.AreEqual 200!, root_child1.LayoutWidth
    Assert.AreEqual 60!, root_child1.LayoutHeight
End Sub

Public Sub Test_percentage_flex_basis_main_max_height()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.FlexBasis = Array(yogaUnitPercent, 10)
    root_child0.MaxHeight = Array(yogaUnitPercent, 60)
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 4
    root_child1.FlexBasis = Array(yogaUnitPercent, 10)
    root_child1.MaxHeight = Array(yogaUnitPercent, 20)
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 52!, root_child0.LayoutWidth
    Assert.AreEqual 120!, root_child0.LayoutHeight

    Assert.AreEqual 52!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 148!, root_child1.LayoutWidth
    Assert.AreEqual 40!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 148!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 52!, root_child0.LayoutWidth
    Assert.AreEqual 120!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 148!, root_child1.LayoutWidth
    Assert.AreEqual 40!, root_child1.LayoutHeight
End Sub

Public Sub Test_percentage_flex_basis_cross_max_height()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.FlexBasis = Array(yogaUnitPercent, 10)
    root_child0.MaxHeight = Array(yogaUnitPercent, 60)
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 4
    root_child1.FlexBasis = Array(yogaUnitPercent, 10)
    root_child1.MaxHeight = Array(yogaUnitPercent, 20)
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 200!, root_child0.LayoutWidth
    Assert.AreEqual 120!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 120!, root_child1.LayoutY
    Assert.AreEqual 200!, root_child1.LayoutWidth
    Assert.AreEqual 40!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 200!, root_child0.LayoutWidth
    Assert.AreEqual 120!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 120!, root_child1.LayoutY
    Assert.AreEqual 200!, root_child1.LayoutWidth
    Assert.AreEqual 40!, root_child1.LayoutHeight
End Sub

Public Sub Test_percentage_flex_basis_main_max_width()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.FlexBasis = Array(yogaUnitPercent, 15)
    root_child0.MaxWidth = Array(yogaUnitPercent, 60)
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 4
    root_child1.FlexBasis = Array(yogaUnitPercent, 10)
    root_child1.MaxWidth = Array(yogaUnitPercent, 20)
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 120!, root_child0.LayoutWidth
    Assert.AreEqual 200!, root_child0.LayoutHeight

    Assert.AreEqual 120!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 40!, root_child1.LayoutWidth
    Assert.AreEqual 200!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 80!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 120!, root_child0.LayoutWidth
    Assert.AreEqual 200!, root_child0.LayoutHeight

    Assert.AreEqual 40!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 40!, root_child1.LayoutWidth
    Assert.AreEqual 200!, root_child1.LayoutHeight
End Sub

Public Sub Test_percentage_flex_basis_cross_max_width()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.FlexBasis = Array(yogaUnitPercent, 10)
    root_child0.MaxWidth = Array(yogaUnitPercent, 60)
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 4
    root_child1.FlexBasis = Array(yogaUnitPercent, 15)
    root_child1.MaxWidth = Array(yogaUnitPercent, 20)
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 120!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 40!, root_child1.LayoutWidth
    Assert.AreEqual 150!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 80!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 120!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 160!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 40!, root_child1.LayoutWidth
    Assert.AreEqual 150!, root_child1.LayoutHeight
End Sub

Public Sub Test_percentage_flex_basis_main_min_width()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.FlexBasis = Array(yogaUnitPercent, 15)
    root_child0.MinWidth = Array(yogaUnitPercent, 60)
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 4
    root_child1.FlexBasis = Array(yogaUnitPercent, 10)
    root_child1.MinWidth = Array(yogaUnitPercent, 20)
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 120!, root_child0.LayoutWidth
    Assert.AreEqual 200!, root_child0.LayoutHeight

    Assert.AreEqual 120!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 80!, root_child1.LayoutWidth
    Assert.AreEqual 200!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 80!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 120!, root_child0.LayoutWidth
    Assert.AreEqual 200!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 80!, root_child1.LayoutWidth
    Assert.AreEqual 200!, root_child1.LayoutHeight
End Sub

Public Sub Test_percentage_flex_basis_cross_min_width()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.FlexBasis = Array(yogaUnitPercent, 10)
    root_child0.MinWidth = Array(yogaUnitPercent, 60)
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 4
    root_child1.FlexBasis = Array(yogaUnitPercent, 15)
    root_child1.MinWidth = Array(yogaUnitPercent, 20)
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 200!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 200!, root_child1.LayoutWidth
    Assert.AreEqual 150!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 200!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 200!, root_child1.LayoutWidth
    Assert.AreEqual 150!, root_child1.LayoutHeight
End Sub

Public Sub Test_percentage_multiple_nested_with_padding_margin_and_percentage_values()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.FlexBasis = Array(yogaUnitPercent, 10)
    root_child0.MarginLeft = 5
    root_child0.MarginTop = 5
    root_child0.MarginRight = 5
    root_child0.MarginBottom = 5
    root_child0.PaddingLeft = 3
    root_child0.PaddingTop = 3
    root_child0.PaddingRight = 3
    root_child0.PaddingBottom = 3
    root_child0.MinWidth = Array(yogaUnitPercent, 60)
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(config)
    root_child0_child0.MarginLeft = 5
    root_child0_child0.MarginTop = 5
    root_child0_child0.MarginRight = 5
    root_child0_child0.MarginBottom = 5
    root_child0_child0.PaddingLeft = Array(yogaUnitPercent, 3)
    root_child0_child0.PaddingTop = Array(yogaUnitPercent, 3)
    root_child0_child0.PaddingRight = Array(yogaUnitPercent, 3)
    root_child0_child0.PaddingBottom = Array(yogaUnitPercent, 3)
    root_child0_child0.Width = Array(yogaUnitPercent, 50)
    root_child0.Insert 0, root_child0_child0

    Dim root_child0_child0_child0 As cYogaNode: Set root_child0_child0_child0 = YogaNodeNew(config)
    root_child0_child0_child0.MarginLeft = Array(yogaUnitPercent, 5)
    root_child0_child0_child0.MarginTop = Array(yogaUnitPercent, 5)
    root_child0_child0_child0.MarginRight = Array(yogaUnitPercent, 5)
    root_child0_child0_child0.MarginBottom = Array(yogaUnitPercent, 5)
    root_child0_child0_child0.PaddingLeft = 3
    root_child0_child0_child0.PaddingTop = 3
    root_child0_child0_child0.PaddingRight = 3
    root_child0_child0_child0.PaddingBottom = 3
    root_child0_child0_child0.Width = Array(yogaUnitPercent, 45)
    root_child0_child0.Insert 0, root_child0_child0_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 4
    root_child1.FlexBasis = Array(yogaUnitPercent, 15)
    root_child1.MinWidth = Array(yogaUnitPercent, 20)
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 5!, root_child0.LayoutX
    Assert.AreEqual 5!, root_child0.LayoutY
    Assert.AreEqual 190!, root_child0.LayoutWidth
    Assert.AreEqual 48!, root_child0.LayoutHeight

    Assert.AreEqual 8!, root_child0_child0.LayoutX
    Assert.AreEqual 8!, root_child0_child0.LayoutY
    Assert.AreEqual 92!, root_child0_child0.LayoutWidth
    Assert.AreEqual 25!, root_child0_child0.LayoutHeight

    Assert.AreEqual 10!, root_child0_child0_child0.LayoutX
    Assert.AreEqual 10!, root_child0_child0_child0.LayoutY
    Assert.AreEqual 36!, root_child0_child0_child0.LayoutWidth
    Assert.AreEqual 6!, root_child0_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 58!, root_child1.LayoutY
    Assert.AreEqual 200!, root_child1.LayoutWidth
    Assert.AreEqual 142!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 5!, root_child0.LayoutX
    Assert.AreEqual 5!, root_child0.LayoutY
    Assert.AreEqual 190!, root_child0.LayoutWidth
    Assert.AreEqual 48!, root_child0.LayoutHeight

    Assert.AreEqual 90!, root_child0_child0.LayoutX
    Assert.AreEqual 8!, root_child0_child0.LayoutY
    Assert.AreEqual 92!, root_child0_child0.LayoutWidth
    Assert.AreEqual 25!, root_child0_child0.LayoutHeight

    Assert.AreEqual 46!, root_child0_child0_child0.LayoutX
    Assert.AreEqual 10!, root_child0_child0_child0.LayoutY
    Assert.AreEqual 36!, root_child0_child0_child0.LayoutWidth
    Assert.AreEqual 6!, root_child0_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 58!, root_child1.LayoutY
    Assert.AreEqual 200!, root_child1.LayoutWidth
    Assert.AreEqual 142!, root_child1.LayoutHeight
End Sub

Public Sub Test_percentage_margin_should_calculate_based_only_on_width()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 200
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.MarginLeft = Array(yogaUnitPercent, 10)
    root_child0.MarginTop = Array(yogaUnitPercent, 10)
    root_child0.MarginRight = Array(yogaUnitPercent, 10)
    root_child0.MarginBottom = Array(yogaUnitPercent, 10)
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(config)
    root_child0_child0.Width = 10
    root_child0_child0.Height = 10
    root_child0.Insert 0, root_child0_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 20!, root_child0.LayoutX
    Assert.AreEqual 20!, root_child0.LayoutY
    Assert.AreEqual 160!, root_child0.LayoutWidth
    Assert.AreEqual 60!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 10!, root_child0_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 20!, root_child0.LayoutX
    Assert.AreEqual 20!, root_child0.LayoutY
    Assert.AreEqual 160!, root_child0.LayoutWidth
    Assert.AreEqual 60!, root_child0.LayoutHeight

    Assert.AreEqual 150!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 10!, root_child0_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0_child0.LayoutHeight
End Sub

Public Sub Test_percentage_padding_should_calculate_based_only_on_width()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 200
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.PaddingLeft = Array(yogaUnitPercent, 10)
    root_child0.PaddingTop = Array(yogaUnitPercent, 10)
    root_child0.PaddingRight = Array(yogaUnitPercent, 10)
    root_child0.PaddingBottom = Array(yogaUnitPercent, 10)
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(config)
    root_child0_child0.Width = 10
    root_child0_child0.Height = 10
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
    Assert.AreEqual 100!, root_child0.LayoutHeight

    Assert.AreEqual 20!, root_child0_child0.LayoutX
    Assert.AreEqual 20!, root_child0_child0.LayoutY
    Assert.AreEqual 10!, root_child0_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 200!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    Assert.AreEqual 170!, root_child0_child0.LayoutX
    Assert.AreEqual 20!, root_child0_child0.LayoutY
    Assert.AreEqual 10!, root_child0_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0_child0.LayoutHeight
End Sub

Public Sub Test_percentage_absolute_position()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 200
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.PositionType = YogaPositionType.yogaPosAbsolute
    root_child0.Left = Array(yogaUnitPercent, 30)
    root_child0.Top = Array(yogaUnitPercent, 10)
    root_child0.Width = 10
    root_child0.Height = 10
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 60!, root_child0.LayoutX
    Assert.AreEqual 10!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 60!, root_child0.LayoutX
    Assert.AreEqual 10!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight
End Sub

Public Sub Test_percentage_width_height_undefined_parent_size()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = Array(yogaUnitPercent, 50)
    root_child0.Height = Array(yogaUnitPercent, 50)
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 0!, root.LayoutWidth
    Assert.AreEqual 0!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 0!, root_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 0!, root.LayoutWidth
    Assert.AreEqual 0!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 0!, root_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0.LayoutHeight
End Sub

Public Sub Test_percent_within_flex_grow()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.Width = 350
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 100
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.FlexGrow = 1
    root.Insert 1, root_child1

    Dim root_child1_child0 As cYogaNode: Set root_child1_child0 = YogaNodeNew(config)
    root_child1_child0.Width = Array(yogaUnitPercent, 100)
    root_child1.Insert 0, root_child1_child0

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Width = 100
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 350!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    Assert.AreEqual 100!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 150!, root_child1.LayoutWidth
    Assert.AreEqual 100!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 150!, root_child1_child0.LayoutWidth
    Assert.AreEqual 0!, root_child1_child0.LayoutHeight

    Assert.AreEqual 250!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 100!, root_child2.LayoutWidth
    Assert.AreEqual 100!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 350!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 250!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    Assert.AreEqual 100!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 150!, root_child1.LayoutWidth
    Assert.AreEqual 100!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 150!, root_child1_child0.LayoutWidth
    Assert.AreEqual 0!, root_child1_child0.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 100!, root_child2.LayoutWidth
    Assert.AreEqual 100!, root_child2.LayoutHeight
End Sub

Public Sub Test_percentage_container_in_wrapping_container()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.JustifyContent = YogaJustify.yogaJustCenter
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(config)
    root_child0_child0.flexDirection = YogaFlexDirection.yogaFlexRow
    root_child0_child0.JustifyContent = YogaJustify.yogaJustCenter
    root_child0_child0.Width = Array(yogaUnitPercent, 100)
    root_child0.Insert 0, root_child0_child0

    Dim root_child0_child0_child0 As cYogaNode: Set root_child0_child0_child0 = YogaNodeNew(config)
    root_child0_child0_child0.Width = 50
    root_child0_child0_child0.Height = 50
    root_child0_child0.Insert 0, root_child0_child0_child0

    Dim root_child0_child0_child1 As cYogaNode: Set root_child0_child0_child1 = YogaNodeNew(config)
    root_child0_child0_child1.Width = 50
    root_child0_child0_child1.Height = 50
    root_child0_child0.Insert 1, root_child0_child0_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 50!, root_child0.LayoutX
    Assert.AreEqual 75!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 100!, root_child0_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutY
    Assert.AreEqual 50!, root_child0_child0_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0_child0_child0.LayoutHeight

    Assert.AreEqual 50!, root_child0_child0_child1.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child1.LayoutY
    Assert.AreEqual 50!, root_child0_child0_child1.LayoutWidth
    Assert.AreEqual 50!, root_child0_child0_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 50!, root_child0.LayoutX
    Assert.AreEqual 75!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 100!, root_child0_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0_child0.LayoutHeight

    Assert.AreEqual 50!, root_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutY
    Assert.AreEqual 50!, root_child0_child0_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child1.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child1.LayoutY
    Assert.AreEqual 50!, root_child0_child0_child1.LayoutWidth
    Assert.AreEqual 50!, root_child0_child0_child1.LayoutHeight
End Sub

Public Sub Test_percent_absolute_position()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 60
    root.Height = 50

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.flexDirection = YogaFlexDirection.yogaFlexRow
    root_child0.PositionType = YogaPositionType.yogaPosAbsolute
    root_child0.Left = Array(yogaUnitPercent, 50)
    root_child0.Width = Array(yogaUnitPercent, 100)
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(config)
    root_child0_child0.Width = Array(yogaUnitPercent, 100)
    root_child0.Insert 0, root_child0_child0

    Dim root_child0_child1 As cYogaNode: Set root_child0_child1 = YogaNodeNew(config)
    root_child0_child1.Width = Array(yogaUnitPercent, 100)
    root_child0.Insert 1, root_child0_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 60!, root.LayoutWidth
    Assert.AreEqual 50!, root.LayoutHeight

    Assert.AreEqual 30!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 60!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 60!, root_child0_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0_child0.LayoutHeight

    Assert.AreEqual 60!, root_child0_child1.LayoutX
    Assert.AreEqual 0!, root_child0_child1.LayoutY
    Assert.AreEqual 60!, root_child0_child1.LayoutWidth
    Assert.AreEqual 50!, root_child0_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 60!, root.LayoutWidth
    Assert.AreEqual 50!, root.LayoutHeight

    Assert.AreEqual 30!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 60!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 60!, root_child0_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0_child0.LayoutHeight

    Assert.AreEqual -60!, root_child0_child1.LayoutX
    Assert.AreEqual 0!, root_child0_child1.LayoutY
    Assert.AreEqual 60!, root_child0_child1.LayoutWidth
    Assert.AreEqual 50!, root_child0_child1.LayoutHeight
End Sub

Private Sub ITestFixture_GetTestCases(ByVal Tests As SimplyVBComp.TestFixtureBuilder)
    Tests.Add "Test_percentage_width_height"
    Tests.Add "Test_percentage_position_left_top"
    Tests.Add "Test_percentage_position_bottom_right"
    Tests.Add "Test_percentage_flex_basis"
    Tests.Add "Test_percentage_flex_basis_cross"
    Tests.Add "Test_percentage_flex_basis_cross_min_height"
    Tests.Add "Test_percentage_flex_basis_main_max_height"
    Tests.Add "Test_percentage_flex_basis_cross_max_height"
    Tests.Add "Test_percentage_flex_basis_main_max_width"
    Tests.Add "Test_percentage_flex_basis_cross_max_width"
    Tests.Add "Test_percentage_flex_basis_main_min_width"
    Tests.Add "Test_percentage_flex_basis_cross_min_width"
    Tests.Add "Test_percentage_multiple_nested_with_padding_margin_and_percentage_values"
    Tests.Add "Test_percentage_margin_should_calculate_based_only_on_width"
    Tests.Add "Test_percentage_padding_should_calculate_based_only_on_width"
    Tests.Add "Test_percentage_absolute_position"
    Tests.Add "Test_percentage_width_height_undefined_parent_size"
    Tests.Add "Test_percent_within_flex_grow"
    Tests.Add "Test_percentage_container_in_wrapping_container"
    Tests.Add "Test_percent_absolute_position"
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTestFlex"
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

Public Sub Test_flex_basis_flex_grow_column()
    Dim oConfig         As cYogaConfig
    Dim oRoot           As cYogaNode
    Dim oRoot_child0    As cYogaNode
    Dim oRoot_child1    As cYogaNode
    
    Set oConfig = YogaConfigNew()

    Set oRoot = YogaNodeNew(oConfig)
    oRoot.Width = 100
    oRoot.Height = 100

    Set oRoot_child0 = YogaNodeNew(oConfig)
    oRoot_child0.FlexGrow = 1
    oRoot_child0.FlexBasis = 50
    oRoot.Insert 0, oRoot_child0

    Set oRoot_child1 = YogaNodeNew(oConfig)
    oRoot_child1.FlexGrow = 1
    oRoot.Insert 1, oRoot_child1
    oRoot.StyleDirection = yogaDirLTR
    oRoot.CalculateLayout

    Assert.AreEqual 0!, oRoot.LayoutX
    Assert.AreEqual 0!, oRoot.LayoutY
    Assert.AreEqual 100!, oRoot.LayoutWidth
    Assert.AreEqual 100!, oRoot.LayoutHeight

    Assert.AreEqual 0!, oRoot_child0.LayoutX
    Assert.AreEqual 0!, oRoot_child0.LayoutY
    Assert.AreEqual 100!, oRoot_child0.LayoutWidth
    Assert.AreEqual 75!, oRoot_child0.LayoutHeight

    Assert.AreEqual 0!, oRoot_child1.LayoutX
    Assert.AreEqual 75!, oRoot_child1.LayoutY
    Assert.AreEqual 100!, oRoot_child1.LayoutWidth
    Assert.AreEqual 25!, oRoot_child1.LayoutHeight

    oRoot.StyleDirection = yogaDirRTL
    oRoot.CalculateLayout

    Assert.AreEqual 0!, oRoot.LayoutX
    Assert.AreEqual 0!, oRoot.LayoutY
    Assert.AreEqual 100!, oRoot.LayoutWidth
    Assert.AreEqual 100!, oRoot.LayoutHeight

    Assert.AreEqual 0!, oRoot_child0.LayoutX
    Assert.AreEqual 0!, oRoot_child0.LayoutY
    Assert.AreEqual 100!, oRoot_child0.LayoutWidth
    Assert.AreEqual 75!, oRoot_child0.LayoutHeight

    Assert.AreEqual 0!, oRoot_child1.LayoutX
    Assert.AreEqual 75!, oRoot_child1.LayoutY
    Assert.AreEqual 100!, oRoot_child1.LayoutWidth
    Assert.AreEqual 25!, oRoot_child1.LayoutHeight
End Sub

Public Sub Test_flex_shrink_flex_grow_row()
    Dim oConfig         As cYogaConfig
    Dim oRoot           As cYogaNode
    Dim oRoot_child0    As cYogaNode
    Dim oRoot_child1    As cYogaNode
    
    Set oConfig = YogaConfigNew()

    Set oRoot = YogaNodeNew(oConfig)
    oRoot.flexDirection = yogaFlexRow
    oRoot.Width = 500
    oRoot.Height = 500

    Set oRoot_child0 = YogaNodeNew(oConfig)
    oRoot_child0.FlexShrink = 1
    oRoot_child0.Width = 500
    oRoot_child0.Height = 100
    oRoot.Insert 0, oRoot_child0

    Set oRoot_child1 = YogaNodeNew(oConfig)
    oRoot_child1.FlexShrink = 1
    oRoot_child1.Width = 500
    oRoot_child1.Height = 100
    oRoot.Insert 1, oRoot_child1
    oRoot.StyleDirection = yogaDirLTR
    oRoot.CalculateLayout

    Assert.AreEqual 0!, oRoot.LayoutX
    Assert.AreEqual 0!, oRoot.LayoutY
    Assert.AreEqual 500!, oRoot.LayoutWidth
    Assert.AreEqual 500!, oRoot.LayoutHeight

    Assert.AreEqual 0!, oRoot_child0.LayoutX
    Assert.AreEqual 0!, oRoot_child0.LayoutY
    Assert.AreEqual 250!, oRoot_child0.LayoutWidth
    Assert.AreEqual 100!, oRoot_child0.LayoutHeight

    Assert.AreEqual 250!, oRoot_child1.LayoutX
    Assert.AreEqual 0!, oRoot_child1.LayoutY
    Assert.AreEqual 250!, oRoot_child1.LayoutWidth
    Assert.AreEqual 100!, oRoot_child1.LayoutHeight

    oRoot.StyleDirection = yogaDirRTL
    oRoot.CalculateLayout

    Assert.AreEqual 0!, oRoot.LayoutX
    Assert.AreEqual 0!, oRoot.LayoutY
    Assert.AreEqual 500!, oRoot.LayoutWidth
    Assert.AreEqual 500!, oRoot.LayoutHeight

    Assert.AreEqual 250!, oRoot_child0.LayoutX
    Assert.AreEqual 0!, oRoot_child0.LayoutY
    Assert.AreEqual 250!, oRoot_child0.LayoutWidth
    Assert.AreEqual 100!, oRoot_child0.LayoutHeight

    Assert.AreEqual 0!, oRoot_child1.LayoutX
    Assert.AreEqual 0!, oRoot_child1.LayoutY
    Assert.AreEqual 250!, oRoot_child1.LayoutWidth
    Assert.AreEqual 100!, oRoot_child1.LayoutHeight
End Sub

Public Sub Test_flex_shrink_flex_grow_child_flex_shrink_other_child()
    Dim oConfig         As cYogaConfig
    Dim oRoot           As cYogaNode
    Dim oRoot_child0    As cYogaNode
    Dim oRoot_child1    As cYogaNode
    
    Set oConfig = YogaConfigNew()

    Set oRoot = YogaNodeNew(oConfig)
    oRoot.flexDirection = yogaFlexRow
    oRoot.Width = 500
    oRoot.Height = 500

    Set oRoot_child0 = YogaNodeNew(oConfig)
    oRoot_child0.FlexShrink = 1
    oRoot_child0.Width = 500
    oRoot_child0.Height = 100
    oRoot.Insert 0, oRoot_child0

    Set oRoot_child1 = YogaNodeNew(oConfig)
    oRoot_child1.FlexGrow = 1
    oRoot_child1.FlexShrink = 1
    oRoot_child1.Width = 500
    oRoot_child1.Height = 100
    oRoot.Insert 1, oRoot_child1
    oRoot.StyleDirection = yogaDirLTR
    oRoot.CalculateLayout

    Assert.AreEqual 0!, oRoot.LayoutX
    Assert.AreEqual 0!, oRoot.LayoutY
    Assert.AreEqual 500!, oRoot.LayoutWidth
    Assert.AreEqual 500!, oRoot.LayoutHeight

    Assert.AreEqual 0!, oRoot_child0.LayoutX
    Assert.AreEqual 0!, oRoot_child0.LayoutY
    Assert.AreEqual 250!, oRoot_child0.LayoutWidth
    Assert.AreEqual 100!, oRoot_child0.LayoutHeight

    Assert.AreEqual 250!, oRoot_child1.LayoutX
    Assert.AreEqual 0!, oRoot_child1.LayoutY
    Assert.AreEqual 250!, oRoot_child1.LayoutWidth
    Assert.AreEqual 100!, oRoot_child1.LayoutHeight

    oRoot.StyleDirection = yogaDirRTL
    oRoot.CalculateLayout

    Assert.AreEqual 0!, oRoot.LayoutX
    Assert.AreEqual 0!, oRoot.LayoutY
    Assert.AreEqual 500!, oRoot.LayoutWidth
    Assert.AreEqual 500!, oRoot.LayoutHeight

    Assert.AreEqual 250!, oRoot_child0.LayoutX
    Assert.AreEqual 0!, oRoot_child0.LayoutY
    Assert.AreEqual 250!, oRoot_child0.LayoutWidth
    Assert.AreEqual 100!, oRoot_child0.LayoutHeight

    Assert.AreEqual 0!, oRoot_child1.LayoutX
    Assert.AreEqual 0!, oRoot_child1.LayoutY
    Assert.AreEqual 250!, oRoot_child1.LayoutWidth
    Assert.AreEqual 100!, oRoot_child1.LayoutHeight
End Sub

Public Sub Test_flex_basis_flex_grow_row()
    Dim oConfig         As cYogaConfig
    Dim oRoot           As cYogaNode
    Dim oRoot_child0    As cYogaNode
    Dim oRoot_child1    As cYogaNode
    
    Set oConfig = YogaConfigNew()

    Set oRoot = YogaNodeNew(oConfig)
    oRoot.flexDirection = yogaFlexRow
    oRoot.Width = 100
    oRoot.Height = 100

    Set oRoot_child0 = YogaNodeNew(oConfig)
    oRoot_child0.FlexGrow = 1
    oRoot_child0.FlexBasis = 50
    oRoot.Insert 0, oRoot_child0

    Set oRoot_child1 = YogaNodeNew(oConfig)
    oRoot_child1.FlexGrow = 1
    oRoot.Insert 1, oRoot_child1
    oRoot.StyleDirection = yogaDirLTR
    oRoot.CalculateLayout

    Assert.AreEqual 0!, oRoot.LayoutX
    Assert.AreEqual 0!, oRoot.LayoutY
    Assert.AreEqual 100!, oRoot.LayoutWidth
    Assert.AreEqual 100!, oRoot.LayoutHeight

    Assert.AreEqual 0!, oRoot_child0.LayoutX
    Assert.AreEqual 0!, oRoot_child0.LayoutY
    Assert.AreEqual 75!, oRoot_child0.LayoutWidth
    Assert.AreEqual 100!, oRoot_child0.LayoutHeight

    Assert.AreEqual 75!, oRoot_child1.LayoutX
    Assert.AreEqual 0!, oRoot_child1.LayoutY
    Assert.AreEqual 25!, oRoot_child1.LayoutWidth
    Assert.AreEqual 100!, oRoot_child1.LayoutHeight

    oRoot.StyleDirection = yogaDirRTL
    oRoot.CalculateLayout

    Assert.AreEqual 0!, oRoot.LayoutX
    Assert.AreEqual 0!, oRoot.LayoutY
    Assert.AreEqual 100!, oRoot.LayoutWidth
    Assert.AreEqual 100!, oRoot.LayoutHeight

    Assert.AreEqual 25!, oRoot_child0.LayoutX
    Assert.AreEqual 0!, oRoot_child0.LayoutY
    Assert.AreEqual 75!, oRoot_child0.LayoutWidth
    Assert.AreEqual 100!, oRoot_child0.LayoutHeight

    Assert.AreEqual 0!, oRoot_child1.LayoutX
    Assert.AreEqual 0!, oRoot_child1.LayoutY
    Assert.AreEqual 25!, oRoot_child1.LayoutWidth
    Assert.AreEqual 100!, oRoot_child1.LayoutHeight
End Sub

Public Sub Test_flex_basis_flex_shrink_column()
    Dim oConfig         As cYogaConfig
    Dim oRoot           As cYogaNode
    Dim oRoot_child0    As cYogaNode
    Dim oRoot_child1    As cYogaNode
    
    Set oConfig = YogaConfigNew()

    Set oRoot = YogaNodeNew(oConfig)
    oRoot.Width = 100
    oRoot.Height = 100

    Set oRoot_child0 = YogaNodeNew(oConfig)
    oRoot_child0.FlexShrink = 1
    oRoot_child0.FlexBasis = 100
    oRoot.Insert 0, oRoot_child0

    Set oRoot_child1 = YogaNodeNew(oConfig)
    oRoot_child1.FlexBasis = 50
    oRoot.Insert 1, oRoot_child1
    oRoot.StyleDirection = yogaDirLTR
    oRoot.CalculateLayout

    Assert.AreEqual 0!, oRoot.LayoutX
    Assert.AreEqual 0!, oRoot.LayoutY
    Assert.AreEqual 100!, oRoot.LayoutWidth
    Assert.AreEqual 100!, oRoot.LayoutHeight

    Assert.AreEqual 0!, oRoot_child0.LayoutX
    Assert.AreEqual 0!, oRoot_child0.LayoutY
    Assert.AreEqual 100!, oRoot_child0.LayoutWidth
    Assert.AreEqual 50!, oRoot_child0.LayoutHeight

    Assert.AreEqual 0!, oRoot_child1.LayoutX
    Assert.AreEqual 50!, oRoot_child1.LayoutY
    Assert.AreEqual 100!, oRoot_child1.LayoutWidth
    Assert.AreEqual 50!, oRoot_child1.LayoutHeight

    oRoot.StyleDirection = yogaDirRTL
    oRoot.CalculateLayout

    Assert.AreEqual 0!, oRoot.LayoutX
    Assert.AreEqual 0!, oRoot.LayoutY
    Assert.AreEqual 100!, oRoot.LayoutWidth
    Assert.AreEqual 100!, oRoot.LayoutHeight

    Assert.AreEqual 0!, oRoot_child0.LayoutX
    Assert.AreEqual 0!, oRoot_child0.LayoutY
    Assert.AreEqual 100!, oRoot_child0.LayoutWidth
    Assert.AreEqual 50!, oRoot_child0.LayoutHeight

    Assert.AreEqual 0!, oRoot_child1.LayoutX
    Assert.AreEqual 50!, oRoot_child1.LayoutY
    Assert.AreEqual 100!, oRoot_child1.LayoutWidth
    Assert.AreEqual 50!, oRoot_child1.LayoutHeight
End Sub

Public Sub Test_flex_basis_flex_shrink_row()
    Dim oConfig         As cYogaConfig
    Dim oRoot           As cYogaNode
    Dim oRoot_child0    As cYogaNode
    Dim oRoot_child1    As cYogaNode
    
    Set oConfig = YogaConfigNew()

    Set oRoot = YogaNodeNew(oConfig)
    oRoot.flexDirection = yogaFlexRow
    oRoot.Width = 100
    oRoot.Height = 100

    Set oRoot_child0 = YogaNodeNew(oConfig)
    oRoot_child0.FlexShrink = 1
    oRoot_child0.FlexBasis = 100
    oRoot.Insert 0, oRoot_child0

    Set oRoot_child1 = YogaNodeNew(oConfig)
    oRoot_child1.FlexBasis = 50
    oRoot.Insert 1, oRoot_child1
    oRoot.StyleDirection = yogaDirLTR
    oRoot.CalculateLayout

    Assert.AreEqual 0!, oRoot.LayoutX
    Assert.AreEqual 0!, oRoot.LayoutY
    Assert.AreEqual 100!, oRoot.LayoutWidth
    Assert.AreEqual 100!, oRoot.LayoutHeight

    Assert.AreEqual 0!, oRoot_child0.LayoutX
    Assert.AreEqual 0!, oRoot_child0.LayoutY
    Assert.AreEqual 50!, oRoot_child0.LayoutWidth
    Assert.AreEqual 100!, oRoot_child0.LayoutHeight

    Assert.AreEqual 50!, oRoot_child1.LayoutX
    Assert.AreEqual 0!, oRoot_child1.LayoutY
    Assert.AreEqual 50!, oRoot_child1.LayoutWidth
    Assert.AreEqual 100!, oRoot_child1.LayoutHeight

    oRoot.StyleDirection = yogaDirRTL
    oRoot.CalculateLayout

    Assert.AreEqual 0!, oRoot.LayoutX
    Assert.AreEqual 0!, oRoot.LayoutY
    Assert.AreEqual 100!, oRoot.LayoutWidth
    Assert.AreEqual 100!, oRoot.LayoutHeight

    Assert.AreEqual 50!, oRoot_child0.LayoutX
    Assert.AreEqual 0!, oRoot_child0.LayoutY
    Assert.AreEqual 50!, oRoot_child0.LayoutWidth
    Assert.AreEqual 100!, oRoot_child0.LayoutHeight

    Assert.AreEqual 0!, oRoot_child1.LayoutX
    Assert.AreEqual 0!, oRoot_child1.LayoutY
    Assert.AreEqual 50!, oRoot_child1.LayoutWidth
    Assert.AreEqual 100!, oRoot_child1.LayoutHeight
End Sub

Public Sub Test_flex_shrink_to_zero()
    Dim oConfig         As cYogaConfig
    Dim oRoot           As cYogaNode
    Dim oRoot_child0    As cYogaNode
    Dim oRoot_child1    As cYogaNode
    Dim oRoot_child2    As cYogaNode
    
    Set oConfig = YogaConfigNew()

    Set oRoot = YogaNodeNew(oConfig)
    oRoot.Height = 75

    Set oRoot_child0 = YogaNodeNew(oConfig)
    oRoot_child0.Width = 50
    oRoot_child0.Height = 50
    oRoot.Insert 0, oRoot_child0

    Set oRoot_child1 = YogaNodeNew(oConfig)
    oRoot_child1.FlexShrink = 1
    oRoot_child1.Width = 50
    oRoot_child1.Height = 50
    oRoot.Insert 1, oRoot_child1

    Set oRoot_child2 = YogaNodeNew(oConfig)
    oRoot_child2.Width = 50
    oRoot_child2.Height = 50
    oRoot.Insert 2, oRoot_child2
    oRoot.StyleDirection = yogaDirLTR
    oRoot.CalculateLayout

    Assert.AreEqual 0!, oRoot.LayoutX
    Assert.AreEqual 0!, oRoot.LayoutY
    Assert.AreEqual 50!, oRoot.LayoutWidth
    Assert.AreEqual 75!, oRoot.LayoutHeight

    Assert.AreEqual 0!, oRoot_child0.LayoutX
    Assert.AreEqual 0!, oRoot_child0.LayoutY
    Assert.AreEqual 50!, oRoot_child0.LayoutWidth
    Assert.AreEqual 50!, oRoot_child0.LayoutHeight

    Assert.AreEqual 0!, oRoot_child1.LayoutX
    Assert.AreEqual 50!, oRoot_child1.LayoutY
    Assert.AreEqual 50!, oRoot_child1.LayoutWidth
    Assert.AreEqual 0!, oRoot_child1.LayoutHeight

    Assert.AreEqual 0!, oRoot_child2.LayoutX
    Assert.AreEqual 50!, oRoot_child2.LayoutY
    Assert.AreEqual 50!, oRoot_child2.LayoutWidth
    Assert.AreEqual 50!, oRoot_child2.LayoutHeight

    oRoot.StyleDirection = yogaDirRTL
    oRoot.CalculateLayout

    Assert.AreEqual 0!, oRoot.LayoutX
    Assert.AreEqual 0!, oRoot.LayoutY
    Assert.AreEqual 50!, oRoot.LayoutWidth
    Assert.AreEqual 75!, oRoot.LayoutHeight

    Assert.AreEqual 0!, oRoot_child0.LayoutX
    Assert.AreEqual 0!, oRoot_child0.LayoutY
    Assert.AreEqual 50!, oRoot_child0.LayoutWidth
    Assert.AreEqual 50!, oRoot_child0.LayoutHeight

    Assert.AreEqual 0!, oRoot_child1.LayoutX
    Assert.AreEqual 50!, oRoot_child1.LayoutY
    Assert.AreEqual 50!, oRoot_child1.LayoutWidth
    Assert.AreEqual 0!, oRoot_child1.LayoutHeight

    Assert.AreEqual 0!, oRoot_child2.LayoutX
    Assert.AreEqual 50!, oRoot_child2.LayoutY
    Assert.AreEqual 50!, oRoot_child2.LayoutWidth
    Assert.AreEqual 50!, oRoot_child2.LayoutHeight
End Sub

Public Sub Test_flex_basis_overrides_main_size()
    Dim oConfig         As cYogaConfig
    Dim oRoot           As cYogaNode
    Dim oRoot_child0    As cYogaNode
    Dim oRoot_child1    As cYogaNode
    Dim oRoot_child2    As cYogaNode
    
    Set oConfig = YogaConfigNew()

    Set oRoot = YogaNodeNew(oConfig)
    oRoot.Width = 100
    oRoot.Height = 100

    Set oRoot_child0 = YogaNodeNew(oConfig)
    oRoot_child0.FlexGrow = 1
    oRoot_child0.FlexBasis = 50
    oRoot_child0.Height = 20
    oRoot.Insert 0, oRoot_child0

    Set oRoot_child1 = YogaNodeNew(oConfig)
    oRoot_child1.FlexGrow = 1
    oRoot_child1.Height = 10
    oRoot.Insert 1, oRoot_child1

    Set oRoot_child2 = YogaNodeNew(oConfig)
    oRoot_child2.FlexGrow = 1
    oRoot_child2.Height = 10
    oRoot.Insert 2, oRoot_child2
    oRoot.StyleDirection = yogaDirLTR
    oRoot.CalculateLayout

    Assert.AreEqual 0!, oRoot.LayoutX
    Assert.AreEqual 0!, oRoot.LayoutY
    Assert.AreEqual 100!, oRoot.LayoutWidth
    Assert.AreEqual 100!, oRoot.LayoutHeight

    Assert.AreEqual 0!, oRoot_child0.LayoutX
    Assert.AreEqual 0!, oRoot_child0.LayoutY
    Assert.AreEqual 100!, oRoot_child0.LayoutWidth
    Assert.AreEqual 60!, oRoot_child0.LayoutHeight

    Assert.AreEqual 0!, oRoot_child1.LayoutX
    Assert.AreEqual 60!, oRoot_child1.LayoutY
    Assert.AreEqual 100!, oRoot_child1.LayoutWidth
    Assert.AreEqual 20!, oRoot_child1.LayoutHeight

    Assert.AreEqual 0!, oRoot_child2.LayoutX
    Assert.AreEqual 80!, oRoot_child2.LayoutY
    Assert.AreEqual 100!, oRoot_child2.LayoutWidth
    Assert.AreEqual 20!, oRoot_child2.LayoutHeight

    oRoot.StyleDirection = yogaDirRTL
    oRoot.CalculateLayout

    Assert.AreEqual 0!, oRoot.LayoutX
    Assert.AreEqual 0!, oRoot.LayoutY
    Assert.AreEqual 100!, oRoot.LayoutWidth
    Assert.AreEqual 100!, oRoot.LayoutHeight

    Assert.AreEqual 0!, oRoot_child0.LayoutX
    Assert.AreEqual 0!, oRoot_child0.LayoutY
    Assert.AreEqual 100!, oRoot_child0.LayoutWidth
    Assert.AreEqual 60!, oRoot_child0.LayoutHeight

    Assert.AreEqual 0!, oRoot_child1.LayoutX
    Assert.AreEqual 60!, oRoot_child1.LayoutY
    Assert.AreEqual 100!, oRoot_child1.LayoutWidth
    Assert.AreEqual 20!, oRoot_child1.LayoutHeight

    Assert.AreEqual 0!, oRoot_child2.LayoutX
    Assert.AreEqual 80!, oRoot_child2.LayoutY
    Assert.AreEqual 100!, oRoot_child2.LayoutWidth
    Assert.AreEqual 20!, oRoot_child2.LayoutHeight
End Sub

Public Sub Test_flex_grow_shrink_at_most()
    Dim oConfig         As cYogaConfig
    Dim oRoot           As cYogaNode
    Dim oRoot_child0    As cYogaNode
    Dim oRoot_child0_child0 As cYogaNode
    
    Set oConfig = YogaConfigNew()

    Set oRoot = YogaNodeNew(oConfig)
    oRoot.Width = 100
    oRoot.Height = 100

    Set oRoot_child0 = YogaNodeNew(oConfig)
    oRoot.Insert 0, oRoot_child0

    Set oRoot_child0_child0 = YogaNodeNew(oConfig)
    oRoot_child0_child0.FlexGrow = 1
    oRoot_child0_child0.FlexShrink = 1
    oRoot_child0.Insert 0, oRoot_child0_child0
    oRoot.StyleDirection = yogaDirLTR
    oRoot.CalculateLayout

    Assert.AreEqual 0!, oRoot.LayoutX
    Assert.AreEqual 0!, oRoot.LayoutY
    Assert.AreEqual 100!, oRoot.LayoutWidth
    Assert.AreEqual 100!, oRoot.LayoutHeight

    Assert.AreEqual 0!, oRoot_child0.LayoutX
    Assert.AreEqual 0!, oRoot_child0.LayoutY
    Assert.AreEqual 100!, oRoot_child0.LayoutWidth
    Assert.AreEqual 0!, oRoot_child0.LayoutHeight

    Assert.AreEqual 0!, oRoot_child0_child0.LayoutX
    Assert.AreEqual 0!, oRoot_child0_child0.LayoutY
    Assert.AreEqual 100!, oRoot_child0_child0.LayoutWidth
    Assert.AreEqual 0!, oRoot_child0_child0.LayoutHeight

    oRoot.StyleDirection = yogaDirRTL
    oRoot.CalculateLayout

    Assert.AreEqual 0!, oRoot.LayoutX
    Assert.AreEqual 0!, oRoot.LayoutY
    Assert.AreEqual 100!, oRoot.LayoutWidth
    Assert.AreEqual 100!, oRoot.LayoutHeight

    Assert.AreEqual 0!, oRoot_child0.LayoutX
    Assert.AreEqual 0!, oRoot_child0.LayoutY
    Assert.AreEqual 100!, oRoot_child0.LayoutWidth
    Assert.AreEqual 0!, oRoot_child0.LayoutHeight

    Assert.AreEqual 0!, oRoot_child0_child0.LayoutX
    Assert.AreEqual 0!, oRoot_child0_child0.LayoutY
    Assert.AreEqual 100!, oRoot_child0_child0.LayoutWidth
    Assert.AreEqual 0!, oRoot_child0_child0.LayoutHeight
End Sub

Public Sub Test_flex_grow_less_than_factor_one()
    Dim oConfig         As cYogaConfig
    Dim oRoot           As cYogaNode
    Dim oRoot_child0    As cYogaNode
    Dim oRoot_child1    As cYogaNode
    Dim oRoot_child2    As cYogaNode
    
    Set oConfig = YogaConfigNew()

    Set oRoot = YogaNodeNew(oConfig)
    oRoot.Width = 200
    oRoot.Height = 500

    Set oRoot_child0 = YogaNodeNew(oConfig)
    oRoot_child0.FlexGrow = 0.2!
    oRoot_child0.FlexBasis = 40
    oRoot.Insert 0, oRoot_child0

    Set oRoot_child1 = YogaNodeNew(oConfig)
    oRoot_child1.FlexGrow = 0.2!
    oRoot.Insert 1, oRoot_child1

    Set oRoot_child2 = YogaNodeNew(oConfig)
    oRoot_child2.FlexGrow = 0.4!
    oRoot.Insert 2, oRoot_child2
    oRoot.StyleDirection = yogaDirLTR
    oRoot.CalculateLayout

    Assert.AreEqual 0!, oRoot.LayoutX
    Assert.AreEqual 0!, oRoot.LayoutY
    Assert.AreEqual 200!, oRoot.LayoutWidth
    Assert.AreEqual 500!, oRoot.LayoutHeight

    Assert.AreEqual 0!, oRoot_child0.LayoutX
    Assert.AreEqual 0!, oRoot_child0.LayoutY
    Assert.AreEqual 200!, oRoot_child0.LayoutWidth
    Assert.AreEqual 132!, oRoot_child0.LayoutHeight

    Assert.AreEqual 0!, oRoot_child1.LayoutX
    Assert.AreEqual 132!, oRoot_child1.LayoutY
    Assert.AreEqual 200!, oRoot_child1.LayoutWidth
    Assert.AreEqual 92!, oRoot_child1.LayoutHeight

    Assert.AreEqual 0!, oRoot_child2.LayoutX
    Assert.AreEqual 224!, oRoot_child2.LayoutY
    Assert.AreEqual 200!, oRoot_child2.LayoutWidth
    Assert.AreEqual 184!, oRoot_child2.LayoutHeight

    oRoot.StyleDirection = yogaDirRTL
    oRoot.CalculateLayout

    Assert.AreEqual 0!, oRoot.LayoutX
    Assert.AreEqual 0!, oRoot.LayoutY
    Assert.AreEqual 200!, oRoot.LayoutWidth
    Assert.AreEqual 500!, oRoot.LayoutHeight

    Assert.AreEqual 0!, oRoot_child0.LayoutX
    Assert.AreEqual 0!, oRoot_child0.LayoutY
    Assert.AreEqual 200!, oRoot_child0.LayoutWidth
    Assert.AreEqual 132!, oRoot_child0.LayoutHeight

    Assert.AreEqual 0!, oRoot_child1.LayoutX
    Assert.AreEqual 132!, oRoot_child1.LayoutY
    Assert.AreEqual 200!, oRoot_child1.LayoutWidth
    Assert.AreEqual 92!, oRoot_child1.LayoutHeight

    Assert.AreEqual 0!, oRoot_child2.LayoutX
    Assert.AreEqual 224!, oRoot_child2.LayoutY
    Assert.AreEqual 200!, oRoot_child2.LayoutWidth
    Assert.AreEqual 184!, oRoot_child2.LayoutHeight
End Sub


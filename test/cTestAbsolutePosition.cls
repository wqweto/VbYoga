VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTestAbsolutePosition"
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

Public Sub Test_absolute_layout_width_height_start_top()
    Dim oConfig         As cYogaConfig
    Dim oRoot           As cYogaNode
    Dim oRoot_child0    As cYogaNode
    
    Set oConfig = YogaConfigNew
    Set oRoot = YogaNodeNew(oConfig)
    oRoot.Width = 100
    oRoot.Height = 100
    
    Set oRoot_child0 = YogaNodeNew(oConfig)
    oRoot_child0.PositionType = yogaPosAbsolute
    oRoot_child0.Start = 10
    oRoot_child0.Top = 10
    oRoot_child0.Width = 10
    oRoot_child0.Height = 10
    oRoot.Insert 0, oRoot_child0
    oRoot.StyleDirection = yogaDirLTR
    oRoot.CalculateLayout
            
    Assert.AreEqual 0!, oRoot.LayoutX
    Assert.AreEqual 0!, oRoot.LayoutY
    Assert.AreEqual 100!, oRoot.LayoutWidth
    Assert.AreEqual 100!, oRoot.LayoutHeight
    
    Assert.AreEqual 10!, oRoot_child0.LayoutX
    Assert.AreEqual 10!, oRoot_child0.LayoutY
    Assert.AreEqual 10!, oRoot_child0.LayoutWidth
    Assert.AreEqual 10!, oRoot_child0.LayoutHeight
    
    oRoot.StyleDirection = yogaDirRTL
    oRoot.CalculateLayout
    
    Assert.AreEqual 0!, oRoot.LayoutX
    Assert.AreEqual 0!, oRoot.LayoutY
    Assert.AreEqual 100!, oRoot.LayoutWidth
    Assert.AreEqual 100!, oRoot.LayoutHeight
    
    Assert.AreEqual 80!, oRoot_child0.LayoutX
    Assert.AreEqual 10!, oRoot_child0.LayoutY
    Assert.AreEqual 10!, oRoot_child0.LayoutWidth
    Assert.AreEqual 10!, oRoot_child0.LayoutHeight
End Sub





Public Sub Test_absolute_layout_width_height_end_bottom()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.PositionType = YogaPositionType.yogaPosAbsolute
    root_child0.End_ = 10
    root_child0.Bottom = 10
    root_child0.Width = 10
    root_child0.Height = 10
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 100!, root.LayoutWidth, "100! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 80!, root_child0.LayoutX, "80! = root_child0.LayoutX"
    Assert.AreEqual 80!, root_child0.LayoutY, "80! = root_child0.LayoutY"
    Assert.AreEqual 10!, root_child0.LayoutWidth, "10! = root_child0.LayoutWidth"
    Assert.AreEqual 10!, root_child0.LayoutHeight, "10! = root_child0.LayoutHeight"

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 100!, root.LayoutWidth, "100! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 10!, root_child0.LayoutX, "10! = root_child0.LayoutX"
    Assert.AreEqual 80!, root_child0.LayoutY, "80! = root_child0.LayoutY"
    Assert.AreEqual 10!, root_child0.LayoutWidth, "10! = root_child0.LayoutWidth"
    Assert.AreEqual 10!, root_child0.LayoutHeight, "10! = root_child0.LayoutHeight"
End Sub

Public Sub Test_absolute_layout_start_top_end_bottom()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.PositionType = YogaPositionType.yogaPosAbsolute
    root_child0.Start = 10
    root_child0.Top = 10
    root_child0.End_ = 10
    root_child0.Bottom = 10
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 100!, root.LayoutWidth, "100! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 10!, root_child0.LayoutX, "10! = root_child0.LayoutX"
    Assert.AreEqual 10!, root_child0.LayoutY, "10! = root_child0.LayoutY"
    Assert.AreEqual 80!, root_child0.LayoutWidth, "80! = root_child0.LayoutWidth"
    Assert.AreEqual 80!, root_child0.LayoutHeight, "80! = root_child0.LayoutHeight"

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 100!, root.LayoutWidth, "100! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 10!, root_child0.LayoutX, "10! = root_child0.LayoutX"
    Assert.AreEqual 10!, root_child0.LayoutY, "10! = root_child0.LayoutY"
    Assert.AreEqual 80!, root_child0.LayoutWidth, "80! = root_child0.LayoutWidth"
    Assert.AreEqual 80!, root_child0.LayoutHeight, "80! = root_child0.LayoutHeight"
End Sub

Public Sub Test_absolute_layout_width_height_start_top_end_bottom()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.PositionType = YogaPositionType.yogaPosAbsolute
    root_child0.Start = 10
    root_child0.Top = 10
    root_child0.End_ = 10
    root_child0.Bottom = 10
    root_child0.Width = 10
    root_child0.Height = 10
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 100!, root.LayoutWidth, "100! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 10!, root_child0.LayoutX, "10! = root_child0.LayoutX"
    Assert.AreEqual 10!, root_child0.LayoutY, "10! = root_child0.LayoutY"
    Assert.AreEqual 10!, root_child0.LayoutWidth, "10! = root_child0.LayoutWidth"
    Assert.AreEqual 10!, root_child0.LayoutHeight, "10! = root_child0.LayoutHeight"

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 100!, root.LayoutWidth, "100! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 80!, root_child0.LayoutX, "80! = root_child0.LayoutX"
    Assert.AreEqual 10!, root_child0.LayoutY, "10! = root_child0.LayoutY"
    Assert.AreEqual 10!, root_child0.LayoutWidth, "10! = root_child0.LayoutWidth"
    Assert.AreEqual 10!, root_child0.LayoutHeight, "10! = root_child0.LayoutHeight"
End Sub

Public Sub Test_do_not_clamp_height_of_absolute_node_to_height_of_its_overflow_hidden_parent()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.FlexDirection = YogaFlexDirection.yogaFlexRow
    root.Overflow = YogaOverflow.yogaOverflowHidden
    root.Width = 50
    root.Height = 50

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.PositionType = YogaPositionType.yogaPosAbsolute
    root_child0.Start = 0
    root_child0.Top = 0
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(Config)
    root_child0_child0.Width = 100
    root_child0_child0.Height = 100
    root_child0.Insert 0, root_child0_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 50!, root.LayoutWidth, "50! = root.LayoutWidth"
    Assert.AreEqual 50!, root.LayoutHeight, "50! = root.LayoutHeight"

    Assert.AreEqual 0!, root_child0.LayoutX, "0! = root_child0.LayoutX"
    Assert.AreEqual 0!, root_child0.LayoutY, "0! = root_child0.LayoutY"
    Assert.AreEqual 100!, root_child0.LayoutWidth, "100! = root_child0.LayoutWidth"
    Assert.AreEqual 100!, root_child0.LayoutHeight, "100! = root_child0.LayoutHeight"

    Assert.AreEqual 0!, root_child0_child0.LayoutX, "0! = root_child0_child0.LayoutX"
    Assert.AreEqual 0!, root_child0_child0.LayoutY, "0! = root_child0_child0.LayoutY"
    Assert.AreEqual 100!, root_child0_child0.LayoutWidth, "100! = root_child0_child0.LayoutWidth"
    Assert.AreEqual 100!, root_child0_child0.LayoutHeight, "100! = root_child0_child0.LayoutHeight"

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 50!, root.LayoutWidth, "50! = root.LayoutWidth"
    Assert.AreEqual 50!, root.LayoutHeight, "50! = root.LayoutHeight"

    Assert.AreEqual -50!, root_child0.LayoutX, "-50! = root_child0.LayoutX"
    Assert.AreEqual 0!, root_child0.LayoutY, "0! = root_child0.LayoutY"
    Assert.AreEqual 100!, root_child0.LayoutWidth, "100! = root_child0.LayoutWidth"
    Assert.AreEqual 100!, root_child0.LayoutHeight, "100! = root_child0.LayoutHeight"

    Assert.AreEqual 0!, root_child0_child0.LayoutX, "0! = root_child0_child0.LayoutX"
    Assert.AreEqual 0!, root_child0_child0.LayoutY, "0! = root_child0_child0.LayoutY"
    Assert.AreEqual 100!, root_child0_child0.LayoutWidth, "100! = root_child0_child0.LayoutWidth"
    Assert.AreEqual 100!, root_child0_child0.LayoutHeight, "100! = root_child0_child0.LayoutHeight"
End Sub

Public Sub Test_absolute_layout_within_border()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.MarginLeft = 10
    root.MarginTop = 10
    root.MarginRight = 10
    root.MarginBottom = 10
    root.PaddingLeft = 10
    root.PaddingTop = 10
    root.PaddingRight = 10
    root.PaddingBottom = 10
    root.BorderLeftWidth = 10
    root.BorderTopWidth = 10
    root.BorderRightWidth = 10
    root.BorderBottomWidth = 10
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.PositionType = YogaPositionType.yogaPosAbsolute
    root_child0.Left = 0
    root_child0.Top = 0
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.PositionType = YogaPositionType.yogaPosAbsolute
    root_child1.Right = 0
    root_child1.Bottom = 0
    root_child1.Width = 50
    root_child1.Height = 50
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(Config)
    root_child2.PositionType = YogaPositionType.yogaPosAbsolute
    root_child2.Left = 0
    root_child2.Top = 0
    root_child2.MarginLeft = 10
    root_child2.MarginTop = 10
    root_child2.MarginRight = 10
    root_child2.MarginBottom = 10
    root_child2.Width = 50
    root_child2.Height = 50
    root.Insert 2, root_child2

    Dim root_child3 As cYogaNode: Set root_child3 = YogaNodeNew(Config)
    root_child3.PositionType = YogaPositionType.yogaPosAbsolute
    root_child3.Right = 0
    root_child3.Bottom = 0
    root_child3.MarginLeft = 10
    root_child3.MarginTop = 10
    root_child3.MarginRight = 10
    root_child3.MarginBottom = 10
    root_child3.Width = 50
    root_child3.Height = 50
    root.Insert 3, root_child3
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 10!, root.LayoutX, "10! = root.LayoutX"
    Assert.AreEqual 10!, root.LayoutY, "10! = root.LayoutY"
    Assert.AreEqual 100!, root.LayoutWidth, "100! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 10!, root_child0.LayoutX, "10! = root_child0.LayoutX"
    Assert.AreEqual 10!, root_child0.LayoutY, "10! = root_child0.LayoutY"
    Assert.AreEqual 50!, root_child0.LayoutWidth, "50! = root_child0.LayoutWidth"
    Assert.AreEqual 50!, root_child0.LayoutHeight, "50! = root_child0.LayoutHeight"

    Assert.AreEqual 40!, root_child1.LayoutX, "40! = root_child1.LayoutX"
    Assert.AreEqual 40!, root_child1.LayoutY, "40! = root_child1.LayoutY"
    Assert.AreEqual 50!, root_child1.LayoutWidth, "50! = root_child1.LayoutWidth"
    Assert.AreEqual 50!, root_child1.LayoutHeight, "50! = root_child1.LayoutHeight"

    Assert.AreEqual 20!, root_child2.LayoutX, "20! = root_child2.LayoutX"
    Assert.AreEqual 20!, root_child2.LayoutY, "20! = root_child2.LayoutY"
    Assert.AreEqual 50!, root_child2.LayoutWidth, "50! = root_child2.LayoutWidth"
    Assert.AreEqual 50!, root_child2.LayoutHeight, "50! = root_child2.LayoutHeight"

    Assert.AreEqual 30!, root_child3.LayoutX, "30! = root_child3.LayoutX"
    Assert.AreEqual 30!, root_child3.LayoutY, "30! = root_child3.LayoutY"
    Assert.AreEqual 50!, root_child3.LayoutWidth, "50! = root_child3.LayoutWidth"
    Assert.AreEqual 50!, root_child3.LayoutHeight, "50! = root_child3.LayoutHeight"

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 10!, root.LayoutX, "10! = root.LayoutX"
    Assert.AreEqual 10!, root.LayoutY, "10! = root.LayoutY"
    Assert.AreEqual 100!, root.LayoutWidth, "100! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 10!, root_child0.LayoutX, "10! = root_child0.LayoutX"
    Assert.AreEqual 10!, root_child0.LayoutY, "10! = root_child0.LayoutY"
    Assert.AreEqual 50!, root_child0.LayoutWidth, "50! = root_child0.LayoutWidth"
    Assert.AreEqual 50!, root_child0.LayoutHeight, "50! = root_child0.LayoutHeight"

    Assert.AreEqual 40!, root_child1.LayoutX, "40! = root_child1.LayoutX"
    Assert.AreEqual 40!, root_child1.LayoutY, "40! = root_child1.LayoutY"
    Assert.AreEqual 50!, root_child1.LayoutWidth, "50! = root_child1.LayoutWidth"
    Assert.AreEqual 50!, root_child1.LayoutHeight, "50! = root_child1.LayoutHeight"

    Assert.AreEqual 20!, root_child2.LayoutX, "20! = root_child2.LayoutX"
    Assert.AreEqual 20!, root_child2.LayoutY, "20! = root_child2.LayoutY"
    Assert.AreEqual 50!, root_child2.LayoutWidth, "50! = root_child2.LayoutWidth"
    Assert.AreEqual 50!, root_child2.LayoutHeight, "50! = root_child2.LayoutHeight"

    Assert.AreEqual 30!, root_child3.LayoutX, "30! = root_child3.LayoutX"
    Assert.AreEqual 30!, root_child3.LayoutY, "30! = root_child3.LayoutY"
    Assert.AreEqual 50!, root_child3.LayoutWidth, "50! = root_child3.LayoutWidth"
    Assert.AreEqual 50!, root_child3.LayoutHeight, "50! = root_child3.LayoutHeight"
End Sub

Public Sub Test_absolute_layout_align_items_and_justify_content_center()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.JustifyContent = YogaJustify.yogaJustCenter
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.FlexGrow = 1
    root.Width = 110
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.PositionType = YogaPositionType.yogaPosAbsolute
    root_child0.Width = 60
    root_child0.Height = 40
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 110!, root.LayoutWidth, "110! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 25!, root_child0.LayoutX, "25! = root_child0.LayoutX"
    Assert.AreEqual 30!, root_child0.LayoutY, "30! = root_child0.LayoutY"
    Assert.AreEqual 60!, root_child0.LayoutWidth, "60! = root_child0.LayoutWidth"
    Assert.AreEqual 40!, root_child0.LayoutHeight, "40! = root_child0.LayoutHeight"

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 110!, root.LayoutWidth, "110! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 25!, root_child0.LayoutX, "25! = root_child0.LayoutX"
    Assert.AreEqual 30!, root_child0.LayoutY, "30! = root_child0.LayoutY"
    Assert.AreEqual 60!, root_child0.LayoutWidth, "60! = root_child0.LayoutWidth"
    Assert.AreEqual 40!, root_child0.LayoutHeight, "40! = root_child0.LayoutHeight"
End Sub

Public Sub Test_absolute_layout_align_items_and_justify_content_flex_end()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.JustifyContent = YogaJustify.yogaJustFlexEnd
    root.AlignItems = YogaAlign.yogaAlignFlexEnd
    root.FlexGrow = 1
    root.Width = 110
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.PositionType = YogaPositionType.yogaPosAbsolute
    root_child0.Width = 60
    root_child0.Height = 40
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 110!, root.LayoutWidth, "110! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 50!, root_child0.LayoutX, "50! = root_child0.LayoutX"
    Assert.AreEqual 60!, root_child0.LayoutY, "60! = root_child0.LayoutY"
    Assert.AreEqual 60!, root_child0.LayoutWidth, "60! = root_child0.LayoutWidth"
    Assert.AreEqual 40!, root_child0.LayoutHeight, "40! = root_child0.LayoutHeight"

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 110!, root.LayoutWidth, "110! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 0!, root_child0.LayoutX, "0! = root_child0.LayoutX"
    Assert.AreEqual 60!, root_child0.LayoutY, "60! = root_child0.LayoutY"
    Assert.AreEqual 60!, root_child0.LayoutWidth, "60! = root_child0.LayoutWidth"
    Assert.AreEqual 40!, root_child0.LayoutHeight, "40! = root_child0.LayoutHeight"
End Sub

Public Sub Test_absolute_layout_justify_content_center()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.JustifyContent = YogaJustify.yogaJustCenter
    root.FlexGrow = 1
    root.Width = 110
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.PositionType = YogaPositionType.yogaPosAbsolute
    root_child0.Width = 60
    root_child0.Height = 40
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 110!, root.LayoutWidth, "110! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 0!, root_child0.LayoutX, "0! = root_child0.LayoutX"
    Assert.AreEqual 30!, root_child0.LayoutY, "30! = root_child0.LayoutY"
    Assert.AreEqual 60!, root_child0.LayoutWidth, "60! = root_child0.LayoutWidth"
    Assert.AreEqual 40!, root_child0.LayoutHeight, "40! = root_child0.LayoutHeight"

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 110!, root.LayoutWidth, "110! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 50!, root_child0.LayoutX, "50! = root_child0.LayoutX"
    Assert.AreEqual 30!, root_child0.LayoutY, "30! = root_child0.LayoutY"
    Assert.AreEqual 60!, root_child0.LayoutWidth, "60! = root_child0.LayoutWidth"
    Assert.AreEqual 40!, root_child0.LayoutHeight, "40! = root_child0.LayoutHeight"
End Sub

Public Sub Test_absolute_layout_align_items_center()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.FlexGrow = 1
    root.Width = 110
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.PositionType = YogaPositionType.yogaPosAbsolute
    root_child0.Width = 60
    root_child0.Height = 40
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 110!, root.LayoutWidth, "110! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 25!, root_child0.LayoutX, "25! = root_child0.LayoutX"
    Assert.AreEqual 0!, root_child0.LayoutY, "0! = root_child0.LayoutY"
    Assert.AreEqual 60!, root_child0.LayoutWidth, "60! = root_child0.LayoutWidth"
    Assert.AreEqual 40!, root_child0.LayoutHeight, "40! = root_child0.LayoutHeight"

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 110!, root.LayoutWidth, "110! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 25!, root_child0.LayoutX, "25! = root_child0.LayoutX"
    Assert.AreEqual 0!, root_child0.LayoutY, "0! = root_child0.LayoutY"
    Assert.AreEqual 60!, root_child0.LayoutWidth, "60! = root_child0.LayoutWidth"
    Assert.AreEqual 40!, root_child0.LayoutHeight, "40! = root_child0.LayoutHeight"
End Sub

Public Sub Test_absolute_layout_align_items_center_on_child_only()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.FlexGrow = 1
    root.Width = 110
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.AlignSelf = YogaAlign.yogaAlignCenter
    root_child0.PositionType = YogaPositionType.yogaPosAbsolute
    root_child0.Width = 60
    root_child0.Height = 40
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 110!, root.LayoutWidth, "110! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 25!, root_child0.LayoutX, "25! = root_child0.LayoutX"
    Assert.AreEqual 0!, root_child0.LayoutY, "0! = root_child0.LayoutY"
    Assert.AreEqual 60!, root_child0.LayoutWidth, "60! = root_child0.LayoutWidth"
    Assert.AreEqual 40!, root_child0.LayoutHeight, "40! = root_child0.LayoutHeight"

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 110!, root.LayoutWidth, "110! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 25!, root_child0.LayoutX, "25! = root_child0.LayoutX"
    Assert.AreEqual 0!, root_child0.LayoutY, "0! = root_child0.LayoutY"
    Assert.AreEqual 60!, root_child0.LayoutWidth, "60! = root_child0.LayoutWidth"
    Assert.AreEqual 40!, root_child0.LayoutHeight, "40! = root_child0.LayoutHeight"
End Sub

Public Sub Test_absolute_layout_align_items_and_justify_content_center_and_top_position()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.JustifyContent = YogaJustify.yogaJustCenter
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.FlexGrow = 1
    root.Width = 110
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.PositionType = YogaPositionType.yogaPosAbsolute
    root_child0.Top = 10
    root_child0.Width = 60
    root_child0.Height = 40
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 110!, root.LayoutWidth, "110! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 25!, root_child0.LayoutX, "25! = root_child0.LayoutX"
    Assert.AreEqual 10!, root_child0.LayoutY, "10! = root_child0.LayoutY"
    Assert.AreEqual 60!, root_child0.LayoutWidth, "60! = root_child0.LayoutWidth"
    Assert.AreEqual 40!, root_child0.LayoutHeight, "40! = root_child0.LayoutHeight"

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 110!, root.LayoutWidth, "110! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 25!, root_child0.LayoutX, "25! = root_child0.LayoutX"
    Assert.AreEqual 10!, root_child0.LayoutY, "10! = root_child0.LayoutY"
    Assert.AreEqual 60!, root_child0.LayoutWidth, "60! = root_child0.LayoutWidth"
    Assert.AreEqual 40!, root_child0.LayoutHeight, "40! = root_child0.LayoutHeight"
End Sub

Public Sub Test_absolute_layout_align_items_and_justify_content_center_and_bottom_position()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.JustifyContent = YogaJustify.yogaJustCenter
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.FlexGrow = 1
    root.Width = 110
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.PositionType = YogaPositionType.yogaPosAbsolute
    root_child0.Bottom = 10
    root_child0.Width = 60
    root_child0.Height = 40
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 110!, root.LayoutWidth, "110! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 25!, root_child0.LayoutX, "25! = root_child0.LayoutX"
    Assert.AreEqual 50!, root_child0.LayoutY, "50! = root_child0.LayoutY"
    Assert.AreEqual 60!, root_child0.LayoutWidth, "60! = root_child0.LayoutWidth"
    Assert.AreEqual 40!, root_child0.LayoutHeight, "40! = root_child0.LayoutHeight"

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 110!, root.LayoutWidth, "110! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 25!, root_child0.LayoutX, "25! = root_child0.LayoutX"
    Assert.AreEqual 50!, root_child0.LayoutY, "50! = root_child0.LayoutY"
    Assert.AreEqual 60!, root_child0.LayoutWidth, "60! = root_child0.LayoutWidth"
    Assert.AreEqual 40!, root_child0.LayoutHeight, "40! = root_child0.LayoutHeight"
End Sub

Public Sub Test_absolute_layout_align_items_and_justify_content_center_and_left_position()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.JustifyContent = YogaJustify.yogaJustCenter
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.FlexGrow = 1
    root.Width = 110
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.PositionType = YogaPositionType.yogaPosAbsolute
    root_child0.Left = 5
    root_child0.Width = 60
    root_child0.Height = 40
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 110!, root.LayoutWidth, "110! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 5!, root_child0.LayoutX, "5! = root_child0.LayoutX"
    Assert.AreEqual 30!, root_child0.LayoutY, "30! = root_child0.LayoutY"
    Assert.AreEqual 60!, root_child0.LayoutWidth, "60! = root_child0.LayoutWidth"
    Assert.AreEqual 40!, root_child0.LayoutHeight, "40! = root_child0.LayoutHeight"

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 110!, root.LayoutWidth, "110! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 5!, root_child0.LayoutX, "5! = root_child0.LayoutX"
    Assert.AreEqual 30!, root_child0.LayoutY, "30! = root_child0.LayoutY"
    Assert.AreEqual 60!, root_child0.LayoutWidth, "60! = root_child0.LayoutWidth"
    Assert.AreEqual 40!, root_child0.LayoutHeight, "40! = root_child0.LayoutHeight"
End Sub

Public Sub Test_absolute_layout_align_items_and_justify_content_center_and_right_position()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.JustifyContent = YogaJustify.yogaJustCenter
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.FlexGrow = 1
    root.Width = 110
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.PositionType = YogaPositionType.yogaPosAbsolute
    root_child0.Right = 5
    root_child0.Width = 60
    root_child0.Height = 40
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 110!, root.LayoutWidth, "110! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 45!, root_child0.LayoutX, "45! = root_child0.LayoutX"
    Assert.AreEqual 30!, root_child0.LayoutY, "30! = root_child0.LayoutY"
    Assert.AreEqual 60!, root_child0.LayoutWidth, "60! = root_child0.LayoutWidth"
    Assert.AreEqual 40!, root_child0.LayoutHeight, "40! = root_child0.LayoutHeight"

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 110!, root.LayoutWidth, "110! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 45!, root_child0.LayoutX, "45! = root_child0.LayoutX"
    Assert.AreEqual 30!, root_child0.LayoutY, "30! = root_child0.LayoutY"
    Assert.AreEqual 60!, root_child0.LayoutWidth, "60! = root_child0.LayoutWidth"
    Assert.AreEqual 40!, root_child0.LayoutHeight, "40! = root_child0.LayoutHeight"
End Sub

Public Sub Test_position_root_with_rtl_should_position_withoutdirection()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.Left = 72
    root.Width = 52
    root.Height = 52
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 72!, root.LayoutX, "72! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 52!, root.LayoutWidth, "52! = root.LayoutWidth"
    Assert.AreEqual 52!, root.LayoutHeight, "52! = root.LayoutHeight"

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 72!, root.LayoutX, "72! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 52!, root.LayoutWidth, "52! = root.LayoutWidth"
    Assert.AreEqual 52!, root.LayoutHeight, "52! = root.LayoutHeight"
End Sub

Public Sub Test_absolute_layout_percentage_bottom_based_on_parent_height()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.Width = 100
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.PositionType = YogaPositionType.yogaPosAbsolute
    root_child0.Top = Array(yogaUnitPercent, 50)
    root_child0.Width = 10
    root_child0.Height = 10
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.PositionType = YogaPositionType.yogaPosAbsolute
    root_child1.Bottom = Array(yogaUnitPercent, 50)
    root_child1.Width = 10
    root_child1.Height = 10
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(Config)
    root_child2.PositionType = YogaPositionType.yogaPosAbsolute
    root_child2.Top = Array(yogaUnitPercent, 10)
    root_child2.Bottom = Array(yogaUnitPercent, 10)
    root_child2.Width = 10
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 100!, root.LayoutWidth, "100! = root.LayoutWidth"
    Assert.AreEqual 200!, root.LayoutHeight, "200! = root.LayoutHeight"

    Assert.AreEqual 0!, root_child0.LayoutX, "0! = root_child0.LayoutX"
    Assert.AreEqual 100!, root_child0.LayoutY, "100! = root_child0.LayoutY"
    Assert.AreEqual 10!, root_child0.LayoutWidth, "10! = root_child0.LayoutWidth"
    Assert.AreEqual 10!, root_child0.LayoutHeight, "10! = root_child0.LayoutHeight"

    Assert.AreEqual 0!, root_child1.LayoutX, "0! = root_child1.LayoutX"
    Assert.AreEqual 90!, root_child1.LayoutY, "90! = root_child1.LayoutY"
    Assert.AreEqual 10!, root_child1.LayoutWidth, "10! = root_child1.LayoutWidth"
    Assert.AreEqual 10!, root_child1.LayoutHeight, "10! = root_child1.LayoutHeight"

    Assert.AreEqual 0!, root_child2.LayoutX, "0! = root_child2.LayoutX"
    Assert.AreEqual 20!, root_child2.LayoutY, "20! = root_child2.LayoutY"
    Assert.AreEqual 10!, root_child2.LayoutWidth, "10! = root_child2.LayoutWidth"
    Assert.AreEqual 160!, root_child2.LayoutHeight, "160! = root_child2.LayoutHeight"

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 100!, root.LayoutWidth, "100! = root.LayoutWidth"
    Assert.AreEqual 200!, root.LayoutHeight, "200! = root.LayoutHeight"

    Assert.AreEqual 90!, root_child0.LayoutX, "90! = root_child0.LayoutX"
    Assert.AreEqual 100!, root_child0.LayoutY, "100! = root_child0.LayoutY"
    Assert.AreEqual 10!, root_child0.LayoutWidth, "10! = root_child0.LayoutWidth"
    Assert.AreEqual 10!, root_child0.LayoutHeight, "10! = root_child0.LayoutHeight"

    Assert.AreEqual 90!, root_child1.LayoutX, "90! = root_child1.LayoutX"
    Assert.AreEqual 90!, root_child1.LayoutY, "90! = root_child1.LayoutY"
    Assert.AreEqual 10!, root_child1.LayoutWidth, "10! = root_child1.LayoutWidth"
    Assert.AreEqual 10!, root_child1.LayoutHeight, "10! = root_child1.LayoutHeight"

    Assert.AreEqual 90!, root_child2.LayoutX, "90! = root_child2.LayoutX"
    Assert.AreEqual 20!, root_child2.LayoutY, "20! = root_child2.LayoutY"
    Assert.AreEqual 10!, root_child2.LayoutWidth, "10! = root_child2.LayoutWidth"
    Assert.AreEqual 160!, root_child2.LayoutHeight, "160! = root_child2.LayoutHeight"
End Sub

Public Sub Test_absolute_layout_in_wrap_reverse_column_container()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.Wrap = YogaWrap.yogaWrapWrapReverse
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.PositionType = YogaPositionType.yogaPosAbsolute
    root_child0.Width = 20
    root_child0.Height = 20
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 100!, root.LayoutWidth, "100! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 80!, root_child0.LayoutX, "80! = root_child0.LayoutX"
    Assert.AreEqual 0!, root_child0.LayoutY, "0! = root_child0.LayoutY"
    Assert.AreEqual 20!, root_child0.LayoutWidth, "20! = root_child0.LayoutWidth"
    Assert.AreEqual 20!, root_child0.LayoutHeight, "20! = root_child0.LayoutHeight"

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 100!, root.LayoutWidth, "100! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 0!, root_child0.LayoutX, "0! = root_child0.LayoutX"
    Assert.AreEqual 0!, root_child0.LayoutY, "0! = root_child0.LayoutY"
    Assert.AreEqual 20!, root_child0.LayoutWidth, "20! = root_child0.LayoutWidth"
    Assert.AreEqual 20!, root_child0.LayoutHeight, "20! = root_child0.LayoutHeight"
End Sub

Public Sub Test_absolute_layout_in_wrap_reverse_row_container()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.FlexDirection = YogaFlexDirection.yogaFlexRow
    root.Wrap = YogaWrap.yogaWrapWrapReverse
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.PositionType = YogaPositionType.yogaPosAbsolute
    root_child0.Width = 20
    root_child0.Height = 20
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 100!, root.LayoutWidth, "100! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 0!, root_child0.LayoutX, "0! = root_child0.LayoutX"
    Assert.AreEqual 80!, root_child0.LayoutY, "80! = root_child0.LayoutY"
    Assert.AreEqual 20!, root_child0.LayoutWidth, "20! = root_child0.LayoutWidth"
    Assert.AreEqual 20!, root_child0.LayoutHeight, "20! = root_child0.LayoutHeight"

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 100!, root.LayoutWidth, "100! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 80!, root_child0.LayoutX, "80! = root_child0.LayoutX"
    Assert.AreEqual 80!, root_child0.LayoutY, "80! = root_child0.LayoutY"
    Assert.AreEqual 20!, root_child0.LayoutWidth, "20! = root_child0.LayoutWidth"
    Assert.AreEqual 20!, root_child0.LayoutHeight, "20! = root_child0.LayoutHeight"
End Sub

Public Sub Test_absolute_layout_in_wrap_reverse_column_container_flex_end()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.Wrap = YogaWrap.yogaWrapWrapReverse
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.AlignSelf = YogaAlign.yogaAlignFlexEnd
    root_child0.PositionType = YogaPositionType.yogaPosAbsolute
    root_child0.Width = 20
    root_child0.Height = 20
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 100!, root.LayoutWidth, "100! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 0!, root_child0.LayoutX, "0! = root_child0.LayoutX"
    Assert.AreEqual 0!, root_child0.LayoutY, "0! = root_child0.LayoutY"
    Assert.AreEqual 20!, root_child0.LayoutWidth, "20! = root_child0.LayoutWidth"
    Assert.AreEqual 20!, root_child0.LayoutHeight, "20! = root_child0.LayoutHeight"

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 100!, root.LayoutWidth, "100! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 80!, root_child0.LayoutX, "80! = root_child0.LayoutX"
    Assert.AreEqual 0!, root_child0.LayoutY, "0! = root_child0.LayoutY"
    Assert.AreEqual 20!, root_child0.LayoutWidth, "20! = root_child0.LayoutWidth"
    Assert.AreEqual 20!, root_child0.LayoutHeight, "20! = root_child0.LayoutHeight"
End Sub

Public Sub Test_absolute_layout_in_wrap_reverse_row_container_flex_end()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.FlexDirection = YogaFlexDirection.yogaFlexRow
    root.Wrap = YogaWrap.yogaWrapWrapReverse
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.AlignSelf = YogaAlign.yogaAlignFlexEnd
    root_child0.PositionType = YogaPositionType.yogaPosAbsolute
    root_child0.Width = 20
    root_child0.Height = 20
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 100!, root.LayoutWidth, "100! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 0!, root_child0.LayoutX, "0! = root_child0.LayoutX"
    Assert.AreEqual 0!, root_child0.LayoutY, "0! = root_child0.LayoutY"
    Assert.AreEqual 20!, root_child0.LayoutWidth, "20! = root_child0.LayoutWidth"
    Assert.AreEqual 20!, root_child0.LayoutHeight, "20! = root_child0.LayoutHeight"

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX, "0! = root.LayoutX"
    Assert.AreEqual 0!, root.LayoutY, "0! = root.LayoutY"
    Assert.AreEqual 100!, root.LayoutWidth, "100! = root.LayoutWidth"
    Assert.AreEqual 100!, root.LayoutHeight, "100! = root.LayoutHeight"

    Assert.AreEqual 80!, root_child0.LayoutX, "80! = root_child0.LayoutX"
    Assert.AreEqual 0!, root_child0.LayoutY, "0! = root_child0.LayoutY"
    Assert.AreEqual 20!, root_child0.LayoutWidth, "20! = root_child0.LayoutWidth"
    Assert.AreEqual 20!, root_child0.LayoutHeight, "20! = root_child0.LayoutHeight"
End Sub

Private Sub ITestFixture_GetTestCases(ByVal Tests As SimplyVBComp.TestFixtureBuilder)
    Tests.Add "Test_absolute_layout_width_height_start_top"
    Tests.Add "Test_absolute_layout_width_height_end_bottom"
    Tests.Add "Test_absolute_layout_start_top_end_bottom"
    Tests.Add "Test_absolute_layout_width_height_start_top_end_bottom"
    Tests.Add "Test_do_not_clamp_height_of_absolute_node_to_height_of_its_overflow_hidden_parent"
    Tests.Add "Test_absolute_layout_within_border"
    Tests.Add "Test_absolute_layout_align_items_and_justify_content_center"
    Tests.Add "Test_absolute_layout_align_items_and_justify_content_flex_end"
    Tests.Add "Test_absolute_layout_justify_content_center"
    Tests.Add "Test_absolute_layout_align_items_center"
    Tests.Add "Test_absolute_layout_align_items_center_on_child_only"
    Tests.Add "Test_absolute_layout_align_items_and_justify_content_center_and_top_position"
    Tests.Add "Test_absolute_layout_align_items_and_justify_content_center_and_bottom_position"
    Tests.Add "Test_absolute_layout_align_items_and_justify_content_center_and_left_position"
    Tests.Add "Test_absolute_layout_align_items_and_justify_content_center_and_right_position"
    Tests.Add "Test_position_root_with_rtl_should_position_withoutdirection"
    Tests.Add "Test_absolute_layout_percentage_bottom_based_on_parent_height"
    Tests.Add "Test_absolute_layout_in_wrap_reverse_column_container"
    Tests.Add "Test_absolute_layout_in_wrap_reverse_row_container"
    Tests.Add "Test_absolute_layout_in_wrap_reverse_column_container_flex_end"
    Tests.Add "Test_absolute_layout_in_wrap_reverse_row_container_flex_end"
End Sub

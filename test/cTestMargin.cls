VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTestMargin"
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

Public Sub Test_margin_start()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.MarginStart = 10
    root_child0.Width = 10
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 10!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 80!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight
End Sub

Public Sub Test_margin_top()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.MarginTop = 10
    root_child0.Height = 10
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 10!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 10!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight
End Sub

Public Sub Test_margin_end()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.JustifyContent = YogaJustify.yogaJustFlexEnd
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.MarginEnd = 10
    root_child0.Width = 10
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 80!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 10!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight
End Sub

Public Sub Test_margin_bottom()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.JustifyContent = YogaJustify.yogaJustFlexEnd
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.MarginBottom = 10
    root_child0.Height = 10
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 80!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 80!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight
End Sub

Public Sub Test_margin_and_flex_row()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.MarginStart = 10
    root_child0.MarginEnd = 10
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 10!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 80!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 10!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 80!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight
End Sub

Public Sub Test_margin_and_flex_column()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.MarginTop = 10
    root_child0.MarginBottom = 10
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 10!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 80!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 10!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 80!, root_child0.LayoutHeight
End Sub

Public Sub Test_margin_and_stretch_row()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.MarginTop = 10
    root_child0.MarginBottom = 10
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 10!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 80!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 10!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 80!, root_child0.LayoutHeight
End Sub

Public Sub Test_margin_and_stretch_column()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.MarginStart = 10
    root_child0.MarginEnd = 10
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 10!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 80!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 10!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 80!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight
End Sub

Public Sub Test_margin_with_sibling_row()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.MarginEnd = 10
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
    Assert.AreEqual 45!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    Assert.AreEqual 55!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 45!, root_child1.LayoutWidth
    Assert.AreEqual 100!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 55!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 45!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 45!, root_child1.LayoutWidth
    Assert.AreEqual 100!, root_child1.LayoutHeight
End Sub

Public Sub Test_margin_with_sibling_column()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.MarginBottom = 10
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
    Assert.AreEqual 45!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 55!, root_child1.LayoutY
    Assert.AreEqual 100!, root_child1.LayoutWidth
    Assert.AreEqual 45!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 45!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 55!, root_child1.LayoutY
    Assert.AreEqual 100!, root_child1.LayoutWidth
    Assert.AreEqual 45!, root_child1.LayoutHeight
End Sub

Public Sub Test_margin_auto_bottom()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.MarginBottom = Array(yogaUnitAuto)
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 50
    root_child1.Height = 50
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 75!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 75!, root_child1.LayoutX
    Assert.AreEqual 150!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 75!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 75!, root_child1.LayoutX
    Assert.AreEqual 150!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight
End Sub

Public Sub Test_margin_auto_top()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.MarginTop = Array(yogaUnitAuto)
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 50
    root_child1.Height = 50
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 75!, root_child0.LayoutX
    Assert.AreEqual 100!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 75!, root_child1.LayoutX
    Assert.AreEqual 150!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 75!, root_child0.LayoutX
    Assert.AreEqual 100!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 75!, root_child1.LayoutX
    Assert.AreEqual 150!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight
End Sub

Public Sub Test_margin_auto_bottom_and_top()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.MarginTop = Array(yogaUnitAuto)
    root_child0.MarginBottom = Array(yogaUnitAuto)
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 50
    root_child1.Height = 50
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 75!, root_child0.LayoutX
    Assert.AreEqual 50!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 75!, root_child1.LayoutX
    Assert.AreEqual 150!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 75!, root_child0.LayoutX
    Assert.AreEqual 50!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 75!, root_child1.LayoutX
    Assert.AreEqual 150!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight
End Sub

Public Sub Test_margin_auto_bottom_and_top_justify_center()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.JustifyContent = YogaJustify.yogaJustCenter
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.MarginTop = Array(yogaUnitAuto)
    root_child0.MarginBottom = Array(yogaUnitAuto)
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 50
    root_child1.Height = 50
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 50!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 150!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 150!, root_child0.LayoutX
    Assert.AreEqual 50!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 150!, root_child1.LayoutX
    Assert.AreEqual 150!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight
End Sub

Public Sub Test_margin_auto_mutiple_children_column()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.MarginTop = Array(yogaUnitAuto)
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.MarginTop = Array(yogaUnitAuto)
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
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 75!, root_child0.LayoutX
    Assert.AreEqual 25!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 75!, root_child1.LayoutX
    Assert.AreEqual 100!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    Assert.AreEqual 75!, root_child2.LayoutX
    Assert.AreEqual 150!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 50!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 75!, root_child0.LayoutX
    Assert.AreEqual 25!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 75!, root_child1.LayoutX
    Assert.AreEqual 100!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    Assert.AreEqual 75!, root_child2.LayoutX
    Assert.AreEqual 150!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 50!, root_child2.LayoutHeight
End Sub

Public Sub Test_margin_auto_mutiple_children_row()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.MarginRight = Array(yogaUnitAuto)
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.MarginRight = Array(yogaUnitAuto)
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
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 75!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 75!, root_child1.LayoutX
    Assert.AreEqual 75!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    Assert.AreEqual 150!, root_child2.LayoutX
    Assert.AreEqual 75!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 50!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 125!, root_child0.LayoutX
    Assert.AreEqual 75!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 75!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 75!, root_child2.LayoutY
    Assert.AreEqual 50!, root_child2.LayoutWidth
    Assert.AreEqual 50!, root_child2.LayoutHeight
End Sub

Public Sub Test_margin_auto_left_and_right_column()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.MarginLeft = Array(yogaUnitAuto)
    root_child0.MarginRight = Array(yogaUnitAuto)
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 50
    root_child1.Height = 50
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 50!, root_child0.LayoutX
    Assert.AreEqual 75!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 150!, root_child1.LayoutX
    Assert.AreEqual 75!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 100!, root_child0.LayoutX
    Assert.AreEqual 75!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 75!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight
End Sub

Public Sub Test_margin_auto_left_and_right()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.MarginLeft = Array(yogaUnitAuto)
    root_child0.MarginRight = Array(yogaUnitAuto)
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 50
    root_child1.Height = 50
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 75!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 75!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 150!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight
End Sub

Public Sub Test_margin_auto_start_and_end_column()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.MarginStart = Array(yogaUnitAuto)
    root_child0.MarginEnd = Array(yogaUnitAuto)
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 50
    root_child1.Height = 50
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 50!, root_child0.LayoutX
    Assert.AreEqual 75!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 150!, root_child1.LayoutX
    Assert.AreEqual 75!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 100!, root_child0.LayoutX
    Assert.AreEqual 75!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 75!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight
End Sub

Public Sub Test_margin_auto_start_and_end()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.MarginStart = Array(yogaUnitAuto)
    root_child0.MarginEnd = Array(yogaUnitAuto)
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 50
    root_child1.Height = 50
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 75!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 75!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 150!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight
End Sub

Public Sub Test_margin_auto_left_and_right_column_and_center()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.MarginLeft = Array(yogaUnitAuto)
    root_child0.MarginRight = Array(yogaUnitAuto)
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 50
    root_child1.Height = 50
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 75!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 75!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 75!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 75!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight
End Sub

Public Sub Test_margin_auto_left()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.MarginLeft = Array(yogaUnitAuto)
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 50
    root_child1.Height = 50
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 150!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 75!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 150!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 75!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight
End Sub

Public Sub Test_margin_auto_right()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.MarginRight = Array(yogaUnitAuto)
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 50
    root_child1.Height = 50
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 75!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 75!, root_child1.LayoutX
    Assert.AreEqual 50!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight
End Sub

Public Sub Test_margin_auto_left_and_right_strech()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.MarginLeft = Array(yogaUnitAuto)
    root_child0.MarginRight = Array(yogaUnitAuto)
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 50
    root_child1.Height = 50
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 50!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 150!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 100!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight
End Sub

Public Sub Test_margin_auto_top_and_bottom_strech()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.MarginTop = Array(yogaUnitAuto)
    root_child0.MarginBottom = Array(yogaUnitAuto)
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 50
    root_child1.Height = 50
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 50!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 150!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 150!, root_child0.LayoutX
    Assert.AreEqual 50!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 150!, root_child1.LayoutX
    Assert.AreEqual 150!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight
End Sub

Public Sub Test_margin_should_not_be_part_of_max_height()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 250
    root.Height = 250

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.MarginTop = 20
    root_child0.Width = 100
    root_child0.Height = 100
    root_child0.MaxHeight = 100
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 250!, root.LayoutWidth
    Assert.AreEqual 250!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 20!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 250!, root.LayoutWidth
    Assert.AreEqual 250!, root.LayoutHeight

    Assert.AreEqual 150!, root_child0.LayoutX
    Assert.AreEqual 20!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight
End Sub

Public Sub Test_margin_should_not_be_part_of_max_width()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 250
    root.Height = 250

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.MarginLeft = 20
    root_child0.Width = 100
    root_child0.MaxWidth = 100
    root_child0.Height = 100
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 250!, root.LayoutWidth
    Assert.AreEqual 250!, root.LayoutHeight

    Assert.AreEqual 20!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 250!, root.LayoutWidth
    Assert.AreEqual 250!, root.LayoutHeight

    Assert.AreEqual 150!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 100!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight
End Sub

Public Sub Test_margin_auto_left_right_child_bigger_than_parent()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.JustifyContent = YogaJustify.yogaJustCenter
    root.Width = 52
    root.Height = 52

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.MarginLeft = Array(yogaUnitAuto)
    root_child0.MarginRight = Array(yogaUnitAuto)
    root_child0.Width = 72
    root_child0.Height = 72
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 52!, root.LayoutWidth
    Assert.AreEqual 52!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual -10!, root_child0.LayoutY
    Assert.AreEqual 72!, root_child0.LayoutWidth
    Assert.AreEqual 72!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 52!, root.LayoutWidth
    Assert.AreEqual 52!, root.LayoutHeight

    Assert.AreEqual -20!, root_child0.LayoutX
    Assert.AreEqual -10!, root_child0.LayoutY
    Assert.AreEqual 72!, root_child0.LayoutWidth
    Assert.AreEqual 72!, root_child0.LayoutHeight
End Sub

Public Sub Test_margin_auto_left_child_bigger_than_parent()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.JustifyContent = YogaJustify.yogaJustCenter
    root.Width = 52
    root.Height = 52

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.MarginLeft = Array(yogaUnitAuto)
    root_child0.Width = 72
    root_child0.Height = 72
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 52!, root.LayoutWidth
    Assert.AreEqual 52!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual -10!, root_child0.LayoutY
    Assert.AreEqual 72!, root_child0.LayoutWidth
    Assert.AreEqual 72!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 52!, root.LayoutWidth
    Assert.AreEqual 52!, root.LayoutHeight

    Assert.AreEqual -20!, root_child0.LayoutX
    Assert.AreEqual -10!, root_child0.LayoutY
    Assert.AreEqual 72!, root_child0.LayoutWidth
    Assert.AreEqual 72!, root_child0.LayoutHeight
End Sub

Public Sub Test_margin_fix_left_auto_right_child_bigger_than_parent()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.JustifyContent = YogaJustify.yogaJustCenter
    root.Width = 52
    root.Height = 52

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.MarginLeft = 10
    root_child0.MarginRight = Array(yogaUnitAuto)
    root_child0.Width = 72
    root_child0.Height = 72
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 52!, root.LayoutWidth
    Assert.AreEqual 52!, root.LayoutHeight

    Assert.AreEqual 10!, root_child0.LayoutX
    Assert.AreEqual -10!, root_child0.LayoutY
    Assert.AreEqual 72!, root_child0.LayoutWidth
    Assert.AreEqual 72!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 52!, root.LayoutWidth
    Assert.AreEqual 52!, root.LayoutHeight

    Assert.AreEqual -20!, root_child0.LayoutX
    Assert.AreEqual -10!, root_child0.LayoutY
    Assert.AreEqual 72!, root_child0.LayoutWidth
    Assert.AreEqual 72!, root_child0.LayoutHeight
End Sub

Public Sub Test_margin_auto_left_fix_right_child_bigger_than_parent()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.JustifyContent = YogaJustify.yogaJustCenter
    root.Width = 52
    root.Height = 52

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.MarginLeft = Array(yogaUnitAuto)
    root_child0.MarginRight = 10
    root_child0.Width = 72
    root_child0.Height = 72
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 52!, root.LayoutWidth
    Assert.AreEqual 52!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual -10!, root_child0.LayoutY
    Assert.AreEqual 72!, root_child0.LayoutWidth
    Assert.AreEqual 72!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 52!, root.LayoutWidth
    Assert.AreEqual 52!, root.LayoutHeight

    Assert.AreEqual -30!, root_child0.LayoutX
    Assert.AreEqual -10!, root_child0.LayoutY
    Assert.AreEqual 72!, root_child0.LayoutWidth
    Assert.AreEqual 72!, root_child0.LayoutHeight
End Sub

Public Sub Test_margin_auto_top_stretching_child()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.FlexShrink = 1
    root_child0.FlexBasis = Array(yogaUnitPercent, 0)
    root_child0.MarginTop = Array(yogaUnitAuto)
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 50
    root_child1.Height = 50
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 100!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 0!, root_child0.LayoutWidth
    Assert.AreEqual 150!, root_child0.LayoutHeight

    Assert.AreEqual 75!, root_child1.LayoutX
    Assert.AreEqual 150!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 100!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 0!, root_child0.LayoutWidth
    Assert.AreEqual 150!, root_child0.LayoutHeight

    Assert.AreEqual 75!, root_child1.LayoutX
    Assert.AreEqual 150!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight
End Sub

Public Sub Test_margin_auto_left_stretching_child()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.AlignItems = YogaAlign.yogaAlignCenter
    root.Width = 200
    root.Height = 200

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.FlexGrow = 1
    root_child0.FlexShrink = 1
    root_child0.FlexBasis = Array(yogaUnitPercent, 0)
    root_child0.MarginLeft = Array(yogaUnitAuto)
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 50
    root_child1.Height = 50
    root.Insert 1, root_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 200!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 0!, root_child0.LayoutWidth
    Assert.AreEqual 150!, root_child0.LayoutHeight

    Assert.AreEqual 75!, root_child1.LayoutX
    Assert.AreEqual 150!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 200!, root.LayoutWidth
    Assert.AreEqual 200!, root.LayoutHeight

    Assert.AreEqual 200!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 0!, root_child0.LayoutWidth
    Assert.AreEqual 150!, root_child0.LayoutHeight

    Assert.AreEqual 75!, root_child1.LayoutX
    Assert.AreEqual 150!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 50!, root_child1.LayoutHeight
End Sub


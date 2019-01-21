VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTestJustifyContent"
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

Public Sub Test_justify_content_row_flex_start()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.Width = 102
    root.Height = 102

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 10
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 10
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Width = 10
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 102!, root.LayoutWidth
    Assert.AreEqual 102!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 102!, root_child0.LayoutHeight

    Assert.AreEqual 10!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 10!, root_child1.LayoutWidth
    Assert.AreEqual 102!, root_child1.LayoutHeight

    Assert.AreEqual 20!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 10!, root_child2.LayoutWidth
    Assert.AreEqual 102!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 102!, root.LayoutWidth
    Assert.AreEqual 102!, root.LayoutHeight

    Assert.AreEqual 92!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 102!, root_child0.LayoutHeight

    Assert.AreEqual 82!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 10!, root_child1.LayoutWidth
    Assert.AreEqual 102!, root_child1.LayoutHeight

    Assert.AreEqual 72!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 10!, root_child2.LayoutWidth
    Assert.AreEqual 102!, root_child2.LayoutHeight
End Sub

Public Sub Test_justify_content_row_flex_end()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.JustifyContent = YogaJustify.yogaJustFlexEnd
    root.Width = 102
    root.Height = 102

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 10
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 10
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Width = 10
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 102!, root.LayoutWidth
    Assert.AreEqual 102!, root.LayoutHeight

    Assert.AreEqual 72!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 102!, root_child0.LayoutHeight

    Assert.AreEqual 82!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 10!, root_child1.LayoutWidth
    Assert.AreEqual 102!, root_child1.LayoutHeight

    Assert.AreEqual 92!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 10!, root_child2.LayoutWidth
    Assert.AreEqual 102!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 102!, root.LayoutWidth
    Assert.AreEqual 102!, root.LayoutHeight

    Assert.AreEqual 20!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 102!, root_child0.LayoutHeight

    Assert.AreEqual 10!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 10!, root_child1.LayoutWidth
    Assert.AreEqual 102!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 10!, root_child2.LayoutWidth
    Assert.AreEqual 102!, root_child2.LayoutHeight
End Sub

Public Sub Test_justify_content_row_center()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.JustifyContent = YogaJustify.yogaJustCenter
    root.Width = 102
    root.Height = 102

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 10
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 10
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Width = 10
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 102!, root.LayoutWidth
    Assert.AreEqual 102!, root.LayoutHeight

    Assert.AreEqual 36!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 102!, root_child0.LayoutHeight

    Assert.AreEqual 46!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 10!, root_child1.LayoutWidth
    Assert.AreEqual 102!, root_child1.LayoutHeight

    Assert.AreEqual 56!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 10!, root_child2.LayoutWidth
    Assert.AreEqual 102!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 102!, root.LayoutWidth
    Assert.AreEqual 102!, root.LayoutHeight

    Assert.AreEqual 56!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 102!, root_child0.LayoutHeight

    Assert.AreEqual 46!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 10!, root_child1.LayoutWidth
    Assert.AreEqual 102!, root_child1.LayoutHeight

    Assert.AreEqual 36!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 10!, root_child2.LayoutWidth
    Assert.AreEqual 102!, root_child2.LayoutHeight
End Sub

Public Sub Test_justify_content_row_space_between()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.JustifyContent = YogaJustify.yogaJustSpaceBetween
    root.Width = 102
    root.Height = 102

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 10
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 10
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Width = 10
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 102!, root.LayoutWidth
    Assert.AreEqual 102!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 102!, root_child0.LayoutHeight

    Assert.AreEqual 46!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 10!, root_child1.LayoutWidth
    Assert.AreEqual 102!, root_child1.LayoutHeight

    Assert.AreEqual 92!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 10!, root_child2.LayoutWidth
    Assert.AreEqual 102!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 102!, root.LayoutWidth
    Assert.AreEqual 102!, root.LayoutHeight

    Assert.AreEqual 92!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 102!, root_child0.LayoutHeight

    Assert.AreEqual 46!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 10!, root_child1.LayoutWidth
    Assert.AreEqual 102!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 10!, root_child2.LayoutWidth
    Assert.AreEqual 102!, root_child2.LayoutHeight
End Sub

Public Sub Test_justify_content_row_space_around()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.JustifyContent = YogaJustify.yogaJustSpaceAround
    root.Width = 102
    root.Height = 102

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 10
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Width = 10
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Width = 10
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 102!, root.LayoutWidth
    Assert.AreEqual 102!, root.LayoutHeight

    Assert.AreEqual 12!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 102!, root_child0.LayoutHeight

    Assert.AreEqual 46!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 10!, root_child1.LayoutWidth
    Assert.AreEqual 102!, root_child1.LayoutHeight

    Assert.AreEqual 80!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 10!, root_child2.LayoutWidth
    Assert.AreEqual 102!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 102!, root.LayoutWidth
    Assert.AreEqual 102!, root.LayoutHeight

    Assert.AreEqual 80!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 102!, root_child0.LayoutHeight

    Assert.AreEqual 46!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 10!, root_child1.LayoutWidth
    Assert.AreEqual 102!, root_child1.LayoutHeight

    Assert.AreEqual 12!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 10!, root_child2.LayoutWidth
    Assert.AreEqual 102!, root_child2.LayoutHeight
End Sub

Public Sub Test_justify_content_column_flex_start()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.Width = 102
    root.Height = 102

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Height = 10
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Height = 10
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Height = 10
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 102!, root.LayoutWidth
    Assert.AreEqual 102!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 102!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 10!, root_child1.LayoutY
    Assert.AreEqual 102!, root_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 20!, root_child2.LayoutY
    Assert.AreEqual 102!, root_child2.LayoutWidth
    Assert.AreEqual 10!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 102!, root.LayoutWidth
    Assert.AreEqual 102!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 102!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 10!, root_child1.LayoutY
    Assert.AreEqual 102!, root_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 20!, root_child2.LayoutY
    Assert.AreEqual 102!, root_child2.LayoutWidth
    Assert.AreEqual 10!, root_child2.LayoutHeight
End Sub

Public Sub Test_justify_content_column_flex_end()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.JustifyContent = YogaJustify.yogaJustFlexEnd
    root.Width = 102
    root.Height = 102

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Height = 10
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Height = 10
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Height = 10
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 102!, root.LayoutWidth
    Assert.AreEqual 102!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 72!, root_child0.LayoutY
    Assert.AreEqual 102!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 82!, root_child1.LayoutY
    Assert.AreEqual 102!, root_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 92!, root_child2.LayoutY
    Assert.AreEqual 102!, root_child2.LayoutWidth
    Assert.AreEqual 10!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 102!, root.LayoutWidth
    Assert.AreEqual 102!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 72!, root_child0.LayoutY
    Assert.AreEqual 102!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 82!, root_child1.LayoutY
    Assert.AreEqual 102!, root_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 92!, root_child2.LayoutY
    Assert.AreEqual 102!, root_child2.LayoutWidth
    Assert.AreEqual 10!, root_child2.LayoutHeight
End Sub

Public Sub Test_justify_content_column_center()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.JustifyContent = YogaJustify.yogaJustCenter
    root.Width = 102
    root.Height = 102

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Height = 10
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Height = 10
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Height = 10
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 102!, root.LayoutWidth
    Assert.AreEqual 102!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 36!, root_child0.LayoutY
    Assert.AreEqual 102!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 46!, root_child1.LayoutY
    Assert.AreEqual 102!, root_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 56!, root_child2.LayoutY
    Assert.AreEqual 102!, root_child2.LayoutWidth
    Assert.AreEqual 10!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 102!, root.LayoutWidth
    Assert.AreEqual 102!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 36!, root_child0.LayoutY
    Assert.AreEqual 102!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 46!, root_child1.LayoutY
    Assert.AreEqual 102!, root_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 56!, root_child2.LayoutY
    Assert.AreEqual 102!, root_child2.LayoutWidth
    Assert.AreEqual 10!, root_child2.LayoutHeight
End Sub

Public Sub Test_justify_content_column_space_between()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.JustifyContent = YogaJustify.yogaJustSpaceBetween
    root.Width = 102
    root.Height = 102

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Height = 10
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Height = 10
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Height = 10
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 102!, root.LayoutWidth
    Assert.AreEqual 102!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 102!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 46!, root_child1.LayoutY
    Assert.AreEqual 102!, root_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 92!, root_child2.LayoutY
    Assert.AreEqual 102!, root_child2.LayoutWidth
    Assert.AreEqual 10!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 102!, root.LayoutWidth
    Assert.AreEqual 102!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 102!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 46!, root_child1.LayoutY
    Assert.AreEqual 102!, root_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 92!, root_child2.LayoutY
    Assert.AreEqual 102!, root_child2.LayoutWidth
    Assert.AreEqual 10!, root_child2.LayoutHeight
End Sub

Public Sub Test_justify_content_column_space_around()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.JustifyContent = YogaJustify.yogaJustSpaceAround
    root.Width = 102
    root.Height = 102

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Height = 10
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Height = 10
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Height = 10
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 102!, root.LayoutWidth
    Assert.AreEqual 102!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 12!, root_child0.LayoutY
    Assert.AreEqual 102!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 46!, root_child1.LayoutY
    Assert.AreEqual 102!, root_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 80!, root_child2.LayoutY
    Assert.AreEqual 102!, root_child2.LayoutWidth
    Assert.AreEqual 10!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 102!, root.LayoutWidth
    Assert.AreEqual 102!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 12!, root_child0.LayoutY
    Assert.AreEqual 102!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 46!, root_child1.LayoutY
    Assert.AreEqual 102!, root_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 80!, root_child2.LayoutY
    Assert.AreEqual 102!, root_child2.LayoutWidth
    Assert.AreEqual 10!, root_child2.LayoutHeight
End Sub

Public Sub Test_justify_content_row_min_width_and_margin()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.JustifyContent = YogaJustify.yogaJustCenter
    root.MarginLeft = 100
    root.MinWidth = 50

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 20
    root_child0.Height = 20
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 100!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 50!, root.LayoutWidth
    Assert.AreEqual 20!, root.LayoutHeight

    Assert.AreEqual 15!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 20!, root_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 100!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 50!, root.LayoutWidth
    Assert.AreEqual 20!, root.LayoutHeight

    Assert.AreEqual 15!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 20!, root_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0.LayoutHeight
End Sub

Public Sub Test_justify_content_row_max_width_and_margin()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.JustifyContent = YogaJustify.yogaJustCenter
    root.MarginLeft = 100
    root.Width = 100
    root.MaxWidth = 80

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 20
    root_child0.Height = 20
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 100!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 80!, root.LayoutWidth
    Assert.AreEqual 20!, root.LayoutHeight

    Assert.AreEqual 30!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 20!, root_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 100!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 80!, root.LayoutWidth
    Assert.AreEqual 20!, root.LayoutHeight

    Assert.AreEqual 30!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 20!, root_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0.LayoutHeight
End Sub

Public Sub Test_justify_content_column_min_height_and_margin()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.JustifyContent = YogaJustify.yogaJustCenter
    root.MarginTop = 100
    root.MinHeight = 50

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 20
    root_child0.Height = 20
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 100!, root.LayoutY
    Assert.AreEqual 20!, root.LayoutWidth
    Assert.AreEqual 50!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 15!, root_child0.LayoutY
    Assert.AreEqual 20!, root_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 100!, root.LayoutY
    Assert.AreEqual 20!, root.LayoutWidth
    Assert.AreEqual 50!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 15!, root_child0.LayoutY
    Assert.AreEqual 20!, root_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0.LayoutHeight
End Sub

Public Sub Test_justify_content_colunn_max_height_and_margin()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.JustifyContent = YogaJustify.yogaJustCenter
    root.MarginTop = 100
    root.Height = 100
    root.MaxHeight = 80

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Width = 20
    root_child0.Height = 20
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 100!, root.LayoutY
    Assert.AreEqual 20!, root.LayoutWidth
    Assert.AreEqual 80!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 30!, root_child0.LayoutY
    Assert.AreEqual 20!, root_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 100!, root.LayoutY
    Assert.AreEqual 20!, root.LayoutWidth
    Assert.AreEqual 80!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 30!, root_child0.LayoutY
    Assert.AreEqual 20!, root_child0.LayoutWidth
    Assert.AreEqual 20!, root_child0.LayoutHeight
End Sub

Public Sub Test_justify_content_column_space_evenly()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.JustifyContent = YogaJustify.yogaJustSpaceEvenly
    root.Width = 102
    root.Height = 102

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Height = 10
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Height = 10
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Height = 10
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 102!, root.LayoutWidth
    Assert.AreEqual 102!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 18!, root_child0.LayoutY
    Assert.AreEqual 102!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 46!, root_child1.LayoutY
    Assert.AreEqual 102!, root_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 74!, root_child2.LayoutY
    Assert.AreEqual 102!, root_child2.LayoutWidth
    Assert.AreEqual 10!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 102!, root.LayoutWidth
    Assert.AreEqual 102!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 18!, root_child0.LayoutY
    Assert.AreEqual 102!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 46!, root_child1.LayoutY
    Assert.AreEqual 102!, root_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child2.LayoutX
    Assert.AreEqual 74!, root_child2.LayoutY
    Assert.AreEqual 102!, root_child2.LayoutWidth
    Assert.AreEqual 10!, root_child2.LayoutHeight
End Sub

Public Sub Test_justify_content_row_space_evenly()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.flexDirection = YogaFlexDirection.yogaFlexRow
    root.JustifyContent = YogaJustify.yogaJustSpaceEvenly
    root.Width = 102
    root.Height = 102

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.Height = 10
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(config)
    root_child1.Height = 10
    root.Insert 1, root_child1

    Dim root_child2 As cYogaNode: Set root_child2 = YogaNodeNew(config)
    root_child2.Height = 10
    root.Insert 2, root_child2
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 102!, root.LayoutWidth
    Assert.AreEqual 102!, root.LayoutHeight

    Assert.AreEqual 26!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 0!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 51!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 0!, root_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1.LayoutHeight

    Assert.AreEqual 77!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 0!, root_child2.LayoutWidth
    Assert.AreEqual 10!, root_child2.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 102!, root.LayoutWidth
    Assert.AreEqual 102!, root.LayoutHeight

    Assert.AreEqual 77!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 0!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    Assert.AreEqual 51!, root_child1.LayoutX
    Assert.AreEqual 0!, root_child1.LayoutY
    Assert.AreEqual 0!, root_child1.LayoutWidth
    Assert.AreEqual 10!, root_child1.LayoutHeight

    Assert.AreEqual 26!, root_child2.LayoutX
    Assert.AreEqual 0!, root_child2.LayoutY
    Assert.AreEqual 0!, root_child2.LayoutWidth
    Assert.AreEqual 10!, root_child2.LayoutHeight
End Sub

Public Sub Test_justify_content_min_width_with_padding_child_width_greater_than_parent()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.AlignContent = YogaAlign.yogaAlignStretch
    root.Width = 1000
    root.Height = 1584

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.flexDirection = YogaFlexDirection.yogaFlexRow
    root_child0.AlignContent = YogaAlign.yogaAlignStretch
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(config)
    root_child0_child0.flexDirection = YogaFlexDirection.yogaFlexRow
    root_child0_child0.JustifyContent = YogaJustify.yogaJustCenter
    root_child0_child0.AlignContent = YogaAlign.yogaAlignStretch
    root_child0_child0.PaddingLeft = 100
    root_child0_child0.PaddingRight = 100
    root_child0_child0.MinWidth = 400
    root_child0.Insert 0, root_child0_child0

    Dim root_child0_child0_child0 As cYogaNode: Set root_child0_child0_child0 = YogaNodeNew(config)
    root_child0_child0_child0.flexDirection = YogaFlexDirection.yogaFlexRow
    root_child0_child0_child0.AlignContent = YogaAlign.yogaAlignStretch
    root_child0_child0_child0.Width = 300
    root_child0_child0_child0.Height = 100
    root_child0_child0.Insert 0, root_child0_child0_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 1000!, root.LayoutWidth
    Assert.AreEqual 1584!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 1000!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 500!, root_child0_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0_child0.LayoutHeight

    Assert.AreEqual 100!, root_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutY
    Assert.AreEqual 300!, root_child0_child0_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0_child0_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 1000!, root.LayoutWidth
    Assert.AreEqual 1584!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 1000!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    Assert.AreEqual 500!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 500!, root_child0_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0_child0.LayoutHeight

    Assert.AreEqual 100!, root_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutY
    Assert.AreEqual 300!, root_child0_child0_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0_child0_child0.LayoutHeight
End Sub

Public Sub Test_justify_content_min_width_with_padding_child_width_lower_than_parent()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.AlignContent = YogaAlign.yogaAlignStretch
    root.Width = 1080
    root.Height = 1584

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root_child0.flexDirection = YogaFlexDirection.yogaFlexRow
    root_child0.AlignContent = YogaAlign.yogaAlignStretch
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(config)
    root_child0_child0.flexDirection = YogaFlexDirection.yogaFlexRow
    root_child0_child0.JustifyContent = YogaJustify.yogaJustCenter
    root_child0_child0.AlignContent = YogaAlign.yogaAlignStretch
    root_child0_child0.PaddingLeft = 100
    root_child0_child0.PaddingRight = 100
    root_child0_child0.MinWidth = 400
    root_child0.Insert 0, root_child0_child0

    Dim root_child0_child0_child0 As cYogaNode: Set root_child0_child0_child0 = YogaNodeNew(config)
    root_child0_child0_child0.flexDirection = YogaFlexDirection.yogaFlexRow
    root_child0_child0_child0.AlignContent = YogaAlign.yogaAlignStretch
    root_child0_child0_child0.Width = 199
    root_child0_child0_child0.Height = 100
    root_child0_child0.Insert 0, root_child0_child0_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 1080!, root.LayoutWidth
    Assert.AreEqual 1584!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 1080!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 400!, root_child0_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0_child0.LayoutHeight

    Assert.AreEqual 101!, root_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutY
    Assert.AreEqual 199!, root_child0_child0_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0_child0_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 1080!, root.LayoutWidth
    Assert.AreEqual 1584!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 1080!, root_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0.LayoutHeight

    Assert.AreEqual 680!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 400!, root_child0_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0_child0.LayoutHeight

    Assert.AreEqual 101!, root_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutY
    Assert.AreEqual 199!, root_child0_child0_child0.LayoutWidth
    Assert.AreEqual 100!, root_child0_child0_child0.LayoutHeight
End Sub

Private Sub ITestFixture_GetTestCases(ByVal Tests As SimplyVBComp.TestFixtureBuilder)
    Tests.Add "Test_justify_content_row_flex_start"
    Tests.Add "Test_justify_content_row_flex_end"
    Tests.Add "Test_justify_content_row_center"
    Tests.Add "Test_justify_content_row_space_between"
    Tests.Add "Test_justify_content_row_space_around"
    Tests.Add "Test_justify_content_column_flex_start"
    Tests.Add "Test_justify_content_column_flex_end"
    Tests.Add "Test_justify_content_column_center"
    Tests.Add "Test_justify_content_column_space_between"
    Tests.Add "Test_justify_content_column_space_around"
    Tests.Add "Test_justify_content_row_min_width_and_margin"
    Tests.Add "Test_justify_content_row_max_width_and_margin"
    Tests.Add "Test_justify_content_column_min_height_and_margin"
    Tests.Add "Test_justify_content_colunn_max_height_and_margin"
    Tests.Add "Test_justify_content_column_space_evenly"
    Tests.Add "Test_justify_content_row_space_evenly"
    Tests.Add "Test_justify_content_min_width_with_padding_child_width_greater_than_parent"
    Tests.Add "Test_justify_content_min_width_with_padding_child_width_lower_than_parent"
End Sub

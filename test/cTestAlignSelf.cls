VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTestAlignSelf"
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

Public Sub Test_align_self_center()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.AlignSelf = YogaAlign.yogaAlignCenter
    root_child0.Width = 10
    root_child0.Height = 10
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 45!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 45!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight
End Sub

Public Sub Test_align_self_flex_end()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.AlignSelf = YogaAlign.yogaAlignFlexEnd
    root_child0.Width = 10
    root_child0.Height = 10
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 90!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight
End Sub

Public Sub Test_align_self_flex_start()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.AlignSelf = YogaAlign.yogaAlignFlexStart
    root_child0.Width = 10
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

Public Sub Test_align_self_flex_end_override_flex_start()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.AlignItems = YogaAlign.yogaAlignFlexStart
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.AlignSelf = YogaAlign.yogaAlignFlexEnd
    root_child0.Width = 10
    root_child0.Height = 10
    root.Insert 0, root_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 90!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 10!, root_child0.LayoutWidth
    Assert.AreEqual 10!, root_child0.LayoutHeight
End Sub

Public Sub Test_align_self_baseline()
    Dim Config As cYogaConfig: Set Config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(Config)
    root.FlexDirection = YogaFlexDirection.yogaFlexRow
    root.Width = 100
    root.Height = 100

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(Config)
    root_child0.AlignSelf = YogaAlign.yogaAlignBaseline
    root_child0.Width = 50
    root_child0.Height = 50
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = YogaNodeNew(Config)
    root_child1.AlignSelf = YogaAlign.yogaAlignBaseline
    root_child1.Width = 50
    root_child1.Height = 20
    root.Insert 1, root_child1

    Dim root_child1_child0 As cYogaNode: Set root_child1_child0 = YogaNodeNew(Config)
    root_child1_child0.Width = 50
    root_child1_child0.Height = 10
    root_child1.Insert 0, root_child1_child0
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 50!, root_child1.LayoutX
    Assert.AreEqual 40!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 50!, root_child1_child0.LayoutWidth
    Assert.AreEqual 10!, root_child1_child0.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 100!, root.LayoutWidth
    Assert.AreEqual 100!, root.LayoutHeight

    Assert.AreEqual 50!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 50!, root_child0.LayoutWidth
    Assert.AreEqual 50!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child1.LayoutX
    Assert.AreEqual 40!, root_child1.LayoutY
    Assert.AreEqual 50!, root_child1.LayoutWidth
    Assert.AreEqual 20!, root_child1.LayoutHeight

    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    Assert.AreEqual 50!, root_child1_child0.LayoutWidth
    Assert.AreEqual 10!, root_child1_child0.LayoutHeight
End Sub

Private Sub ITestFixture_GetTestCases(ByVal Tests As SimplyVBComp.TestFixtureBuilder)
    Tests.Add "Test_align_self_center"
    Tests.Add "Test_align_self_flex_end"
    Tests.Add "Test_align_self_flex_start"
    Tests.Add "Test_align_self_flex_end_override_flex_start"
    Tests.Add "Test_align_self_baseline"
End Sub

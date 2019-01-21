VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTestAlignBaseline"
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

Public Sub Test_align_baseline_parent_using_child_in_column_as_reference()
    Dim config As cYogaConfig: Set config = YogaConfigNew()
    Dim root As cYogaNode: Set root = createYGNode(config, YogaFlexDirection.yogaFlexRow, 1000, 1000, True)
    
    Dim root_child0 As cYogaNode: Set root_child0 = createYGNode(config, YogaFlexDirection.yogaFlexColumn, 500, 600, False)
    root.Insert 0, root_child0
    
    Dim root_child1 As cYogaNode: Set root_child1 = createYGNode(config, YogaFlexDirection.yogaFlexColumn, 500, 800, False)
    root.Insert 1, root_child1
    
    Dim root_child1_child0 As cYogaNode: Set root_child1_child0 = createYGNode(config, YogaFlexDirection.yogaFlexColumn, 500, 300, False)
    root_child1.Insert 0, root_child1_child0
    
    Dim root_child1_child1 As cYogaNode: Set root_child1_child1 = createYGNode(config, YogaFlexDirection.yogaFlexColumn, 500, 400, False)
    root_child1_child1.SetBaselineFunction Me
    root_child1_child1.IsReferenceBaseline = True
    root_child1.Insert 1, root_child1_child1
    
    root.CalculateLayout
    
    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    
    Assert.AreEqual 500!, root_child1.LayoutX
    Assert.AreEqual 100!, root_child1.LayoutY
    
    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY
    
    Assert.AreEqual 0!, root_child1_child1.LayoutX
    Assert.AreEqual 300!, root_child1_child1.LayoutY
End Sub

Public Function BaselineCallback(oNode As cYogaNode, ByVal sngWidth As Single, ByVal sngHeight As Single) As Single
    #If oNode And sngWidth Then
    #End If
    BaselineCallback = sngHeight / 2
End Function

Public Sub Test_align_baseline_parent_using_child_in_row_as_reference()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = createYGNode(config, YogaFlexDirection.yogaFlexRow, 1000, 1000, True)

    Dim root_child0 As cYogaNode: Set root_child0 = createYGNode(config, YogaFlexDirection.yogaFlexColumn, 500, 600, False)
    root.Insert 0, root_child0

    Dim root_child1 As cYogaNode: Set root_child1 = createYGNode(config, YogaFlexDirection.yogaFlexRow, 500, 800, True)
    root.Insert 1, root_child1

    Dim root_child1_child0 As cYogaNode: Set root_child1_child0 = createYGNode(config, YogaFlexDirection.yogaFlexRow, 500, 500, False)
    root_child1.Insert 0, root_child1_child0

    Dim root_child1_child1 As cYogaNode: Set root_child1_child1 = createYGNode(config, YogaFlexDirection.yogaFlexRow, 500, 400, False)
    root_child1_child1.SetBaselineFunction Me
    root_child1_child1.IsReferenceBaseline = True
    root_child1.Insert 1, root_child1_child1

    root.StyleDirection = YogaDirection.yogaDirLTR

    root.CalculateLayout

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY

    Assert.AreEqual 500!, root_child1.LayoutX
    Assert.AreEqual 100!, root_child1.LayoutY

    Assert.AreEqual 0!, root_child1_child0.LayoutX
    Assert.AreEqual 0!, root_child1_child0.LayoutY

    Assert.AreEqual 500!, root_child1_child1.LayoutX
    Assert.AreEqual 300!, root_child1_child1.LayoutY
End Sub

Private Function createYGNode(config As cYogaConfig, flexDirection As YogaFlexDirection, Width As Long, Height As Long, alignBaseline As Boolean) As cYogaNode
    Dim node As cYogaNode: Set node = YogaNodeNew(config)
    node.flexDirection = flexDirection
    node.Width = Width
    node.Height = Height
    If alignBaseline Then
        node.AlignItems = YogaAlign.yogaAlignBaseline
    End If
    Set createYGNode = node
End Function

Private Sub ITestFixture_GetTestCases(ByVal Tests As SimplyVBComp.TestFixtureBuilder)
    Tests.Add "Test_align_baseline_parent_using_child_in_column_as_reference"
    Tests.Add "Test_align_baseline_parent_using_child_in_row_as_reference"
End Sub

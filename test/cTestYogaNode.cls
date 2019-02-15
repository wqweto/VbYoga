VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTestYogaNode"
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

Private Const ObjectVariableNotSetException     As Long = 91
Private Const SubscriptOutOfRangeException      As Long = 9
Private Const InvalidOperationException         As Long = vbObjectError

Private Function C_Obj(Value As Variant) As Object
    Set C_Obj = Value
End Function

Public Sub TestAddChildGetParent()
    Dim Parent As cYogaNode: Set Parent = YogaNodeNew()
    Dim child As cYogaNode: Set child = YogaNodeNew()

    Assert.IsNothing child.Parent
    Assert.AreEqual 0, Parent.Count

    Parent.Insert 0, child

    Assert.AreEqual 1, Parent.Count
    Assert.AreSame child, Parent.Item(0)
    Assert.AreSame Parent, child.Parent

    Parent.RemoveAt 0

    Assert.IsNothing child.Parent
    Assert.AreEqual 0, Parent.Count
End Sub

Public Sub TestChildren()
    Dim Parent As cYogaNode: Set Parent = YogaNodeNew()
    Dim node As cYogaNode
    For Each node In Parent
        Assert.Fail node.ToString()
    Next

    Dim child0 As cYogaNode: Set child0 = YogaNodeNew()
    Assert.AreEqual -1, Parent.IndexOf(child0)
    Parent.Insert 0, child0
    For Each node In Parent
        Assert.AreEqual 0, Parent.IndexOf(node)
    Next

    Dim child1 As cYogaNode: Set child1 = YogaNodeNew()
    Parent.Insert 1, child1
    Dim Index As Long
    For Each node In Parent
        Assert.AreEqual Index, Parent.IndexOf(node)
        Index = Index + 1
    Next

    Parent.RemoveAt 0
    Assert.AreEqual -1, Parent.IndexOf(child0)
    Assert.AreEqual 0, Parent.IndexOf(child1)

    Parent.Clear
    Assert.AreEqual 0, Parent.Count

    Parent.Clear
    Assert.AreEqual 0, Parent.Count
End Sub

Public Sub TestRemoveAtFromEmpty()
    Dim Parent As cYogaNode: Set Parent = YogaNodeNew()
    Assert.Throws ObjectVariableNotSetException, , "Object variable or With block variable not set"
    Parent.RemoveAt 0
End Sub

Public Sub TestRemoveAtOutOfRange()
    Dim Parent As cYogaNode: Set Parent = YogaNodeNew()
    Dim child As cYogaNode: Set child = YogaNodeNew()
    Parent.Insert 0, child
    Assert.Throws SubscriptOutOfRangeException, , "Subscript out of range"
    Parent.RemoveAt 1
End Sub

Public Sub TestCannotAddChildToMultipleParents()
    Dim parent1 As cYogaNode: Set parent1 = YogaNodeNew()
    Dim parent2 As cYogaNode: Set parent2 = YogaNodeNew()
    Dim child As cYogaNode: Set child = YogaNodeNew()

    parent1.Insert 0, child
    Assert.Throws InvalidOperationException, , "Child already has a parent, it must be removed first"
    parent2.Insert 0, child
End Sub

Public Sub TestReset()
    Dim instanceCount As Long: instanceCount = YogaNodeInstanceCount
    Dim node As cYogaNode: Set node = YogaNodeNew()
    Assert.AreEqual instanceCount + 1, YogaNodeInstanceCount()
    node.Reset
    Assert.AreEqual instanceCount + 1, YogaNodeInstanceCount()
End Sub

Public Sub TestResetParent()
    Dim Parent As cYogaNode: Set Parent = YogaNodeNew()
    Dim child As cYogaNode: Set child = YogaNodeNew()
    Parent.Insert 0, child
    Assert.Throws InvalidOperationException, , "Cannot reset a node which still has children attached"
    Parent.Reset
End Sub

Public Sub TestResetChild()
    Dim Parent As cYogaNode: Set Parent = YogaNodeNew()
    Dim child As cYogaNode: Set child = YogaNodeNew()
    Parent.Insert 0, child
    Assert.Throws InvalidOperationException, , "Cannot reset a node still attached to a parent"
    child.Reset
End Sub

Public Sub TestClear()
    On Error GoTo EH
    Dim instanceCount As Long: instanceCount = YogaNodeInstanceCount
    Dim Parent As cYogaNode: Set Parent = YogaNodeNew()
    Assert.AreEqual instanceCount + 1, YogaNodeInstanceCount()
    Dim child As cYogaNode: Set child = YogaNodeNew()
    Assert.AreEqual instanceCount + 2, YogaNodeInstanceCount()
    Parent.Insert 0, child
    Assert.AreEqual 1, Parent.Count
    Assert.AreSame Parent, child.Parent
    Parent.Clear
    Assert.AreEqual 0, Parent.Count
    Assert.IsNothing child.Parent
    Assert.AreEqual instanceCount + 2, YogaNodeInstanceCount()
    Exit Sub
EH:
    Err.Raise Err.Number, Err.Source, Err.Description & "(" & Erl & ")"
End Sub

Public Function MeasureCallback( _
            oNode As cYogaNode, _
            ByVal sngWidth As Single, _
            ByVal eWidthMode As YogaMeasureMode, _
            ByVal sngHeight As Single, _
            ByVal eHeightMode As YogaMeasureMode, _
            sngOutputWidth As Single, _
            sngOutputHeight As Single)
    #If sngWidth And eWidthMode And sngHeight And eHeightMode Then '--- touch args
    #End If
    Select Case oNode.Tag
    Case 0
        sngOutputWidth = 100
        sngOutputHeight = 150
    Case 1
        sngOutputWidth = 123.4!
        sngOutputHeight = 81.7!
    Case 2
        sngOutputWidth = 120
        sngOutputHeight = 130
    End Select
End Function

Public Sub TestMeasureFunc()
    Dim node As cYogaNode: Set node = YogaNodeNew()
    node.Tag = 0
    node.SetMeasureFunction Array(C_Obj(Me), "MeasureCallback")
    node.CalculateLayout
    Assert.AreEqual 100, node.LayoutWidth
    Assert.AreEqual 150, node.LayoutHeight
End Sub

Public Sub TestMeasureFuncWithFloat()
    Dim node As cYogaNode: Set node = YogaNodeNew()
    node.Tag = 1
    node.SetMeasureFunction Array(C_Obj(Me), "MeasureCallback")
    node.CalculateLayout
    Assert.AreEqual 124!, node.LayoutWidth
    Assert.AreEqual 82!, node.LayoutHeight

    Dim config0 As cYogaConfig: Set config0 = YogaConfigNew
    config0.PointScaleFactor = 0
    Set node = YogaNodeNew(config0)
    node.Tag = 1
    node.SetMeasureFunction Array(C_Obj(Me), "MeasureCallback")
    node.CalculateLayout
    Assert.AreEqual 123.4!, node.LayoutWidth
    Assert.AreEqual 81.7!, node.LayoutHeight
End Sub

Public Sub TestChildWithMeasureFunc()
    Dim node As cYogaNode: Set node = YogaNodeNew()
    node.SetMeasureFunction Array(C_Obj(Me), "MeasureCallback")
    Dim child As cYogaNode: Set child = YogaNodeNew()
    Assert.Throws InvalidOperationException, , "Cannot add child: Nodes with measure functions cannot have children"
    node.Insert 0, child
End Sub

Public Sub TestMeasureFuncWithChild()
    Dim node As cYogaNode: Set node = YogaNodeNew()
    Dim child As cYogaNode: Set child = YogaNodeNew()
    node.Insert 0, child
    Assert.Throws InvalidOperationException, , "Cannot set measure function: Nodes with measure functions cannot have children"
    node.SetMeasureFunction Array(C_Obj(Me), "MeasureCallback")
End Sub

Public Function BaselineCallback( _
            oNode As cYogaNode, _
            ByVal sngWidth As Single, _
            ByVal sngHeight As Single) As Single
    Select Case oNode.Tag
    Case 0
        Assert.AreEqual 100, sngWidth
        Assert.AreEqual 110, sngHeight
        BaselineCallback = 65
    Case 1
        Assert.AreEqual 100, sngWidth
        Assert.AreEqual 110, sngHeight
        BaselineCallback = 80
    Case 2
        Assert.AreEqual 100, sngWidth
        Assert.AreEqual 110, sngHeight
        BaselineCallback = 88
    End Select
End Function

Public Sub TestBaselineFunc()
    Dim node As cYogaNode: Set node = YogaNodeNew()
    node.Height = 200
    node.FlexDirection = YogaFlexDirection.yogaFlexRow
    node.AlignItems = YogaAlign.yogaAlignBaseline

    Dim child0 As cYogaNode: Set child0 = YogaNodeNew()
    child0.Width = 100
    child0.Height = 110
    child0.Tag = 0
    child0.SetBaselineFunction Array(C_Obj(Me), "BaselineCallback")
    node.Insert 0, child0

    Dim child1 As cYogaNode: Set child1 = YogaNodeNew()
    child1.Width = 100
    child1.Height = 110
    child1.Tag = 1
    child1.SetBaselineFunction Array(C_Obj(Me), "BaselineCallback")
    node.Insert 1, child1

    Dim child2 As cYogaNode: Set child2 = YogaNodeNew()
    child2.Width = 100
    child2.Height = 110
    child2.Tag = 2
    child2.SetBaselineFunction Array(C_Obj(Me), "BaselineCallback")
    node.Insert 2, child2

    node.CalculateLayout

    Assert.AreEqual 0, child0.LayoutX
    Assert.AreEqual 23, child0.LayoutY
    Assert.AreEqual 100, child1.LayoutX
    Assert.AreEqual 8, child1.LayoutY
    Assert.AreEqual 200, child2.LayoutX
    Assert.AreEqual 0, child2.LayoutY
End Sub

Public Sub TestPrintOneNode()
    Dim node As cYogaNode: Set node = YogaNodeNew()
    node.Width = 100
    node.Height = 120
    node.CalculateLayout
    Assert.AreEqual "<div layout=""width: 100; height: 120; top: 0; left: 0; "" style=""width: 100px; height: 120px; "" ></div>", node.ToString()
End Sub

Public Sub TestPrintWithLogger()
    Dim node As cYogaNode: Set node = YogaNodeNew()
    node.Width = 110
    node.Height = 105
    node.CalculateLayout
    Assert.AreEqual "<div layout=""width: 110; height: 105; top: 0; left: 0; "" style=""width: 110px; height: 105px; "" ></div>", node.ToString()
End Sub

Public Sub TestPrint()
    Dim Parent As cYogaNode: Set Parent = YogaNodeNew()
    Parent.Width = 100
    Parent.Height = 120
    Dim child0 As cYogaNode: Set child0 = YogaNodeNew()
    child0.Width = 30
    child0.Height = 40
    Dim child1 As cYogaNode: Set child1 = YogaNodeNew()
    child1.Width = 35
    child1.Height = 45
    Parent.Insert 0, child0
    Parent.Insert 0, child1
    Parent.CalculateLayout
    Assert.AreEqual "<div layout=""width: 100; height: 120; top: 0; left: 0; "" style=""width: 100px; height: 120px; "" >" & vbCrLf & _
        "  <div layout=""width: 35; height: 45; top: 0; left: 0; "" style=""width: 35px; height: 45px; "" ></div>" & vbCrLf & _
        "  <div layout=""width: 30; height: 40; top: 45; left: 0; "" style=""width: 30px; height: 40px; "" ></div>" & vbCrLf & _
        "</div>", Parent.ToString()
End Sub

Public Sub TestCopyStyle()
    Dim node0 As cYogaNode: Set node0 = YogaNodeNew()
    Assert.IsTrue YogaConstantsIsUndefined(node0.MaxHeight)

    Dim node1 As cYogaNode: Set node1 = YogaNodeNew()
    node1.MaxHeight = 100

    node0.CopyStyle node1
    Assert.AreEqual Array(yogaUnitPoint, 100), node0.MaxHeight
End Sub

Public Sub TestCopyConstructor()
    Dim node0 As cYogaNode: Set node0 = YogaNodeNew()
    node0.MaxWidth = 80

    Dim node1 As cYogaNode: Set node1 = YogaNodeClone(node0)
    Assert.AreEqual Array(yogaUnitPoint, 80), node1.MaxWidth

    Dim node2 As cYogaNode: Set node2 = YogaNodeClone(node1)
    node2.MaxHeight = 90

    Assert.AreEqual Array(yogaUnitPoint, 80), node2.MaxWidth
    Assert.AreEqual Array(yogaUnitPoint, 90), node2.MaxHeight

    Dim node3 As cYogaNode: Set node3 = YogaNodeClone(node0)
    node3.MaxWidth = 100

    Assert.AreEqual Array(yogaUnitPoint, 100), node3.MaxWidth

    Dim node4 As cYogaNode: Set node4 = YogaNodeClone(node2)
    node4.MaxWidth = 100

    Assert.AreEqual Array(yogaUnitPoint, 100), node4.MaxWidth
    Assert.AreEqual Array(yogaUnitPoint, 90), node4.MaxHeight
End Sub

Private Sub ForceGC()
    '--- do nothing
End Sub

Public Sub TestDestructor()
    On Error GoTo EH
    ForceGC
    Dim instanceCount As Long: instanceCount = YogaNodeInstanceCount
    TestDestructorForGC instanceCount
    ForceGC
    Assert.AreEqual instanceCount, YogaNodeInstanceCount()
    Exit Sub
EH:
    Err.Raise Err.Number, Err.Source, Err.Description & ", TestDestructor(" & Erl & ")"
End Sub

Private Sub TestDestructorForGC(ByVal instanceCount As Long)
    On Error GoTo EH
    Dim node As cYogaNode: Set node = YogaNodeNew()
    Assert.AreNotSame Nothing, node
    Assert.AreEqual instanceCount + 1, YogaNodeInstanceCount()
    Set node = Nothing
    Exit Sub
EH:
    Err.Raise Err.Number, Err.Source, Err.Description & ", TestDestructorForGC(" & Erl & ")"
End Sub

Public Sub TestDestructorWithChildren()
    ForceGC
    Dim instanceCount As Long: instanceCount = YogaNodeInstanceCount
    TestDestructorWithChildrenForGC1 instanceCount
    ForceGC
    Assert.AreEqual instanceCount, YogaNodeInstanceCount()
End Sub

Private Sub TestDestructorWithChildrenForGC1(ByVal instanceCount As Long)
    Dim node As cYogaNode: Set node = YogaNodeNew()
    Assert.AreEqual instanceCount + 1, YogaNodeInstanceCount()

    TestDestructorWithChildrenForGC2 node, instanceCount + 1
    ForceGC
    Assert.AreEqual instanceCount + 2, YogaNodeInstanceCount()

    TestDestructorWithChildrenForGC2 node, instanceCount + 2
    ForceGC
    Assert.AreEqual instanceCount + 3, YogaNodeInstanceCount()

    Set node = Nothing
End Sub

Private Sub TestDestructorWithChildrenForGC2(Parent As cYogaNode, ByVal instanceCount As Long)
    Dim child As cYogaNode: Set child = YogaNodeNew()
    Assert.AreEqual instanceCount + 1, YogaNodeInstanceCount()

    Parent.Insert 0, child
    Set child = Nothing
End Sub

Public Sub TestClearWithChildDestructor()
    ForceGC
    Dim instanceCount As Long: instanceCount = YogaNodeInstanceCount
    Dim node As cYogaNode: Set node = YogaNodeNew()
    Assert.AreEqual instanceCount + 1, YogaNodeInstanceCount()
    TestClearWithChildDestructorForGC node, instanceCount + 1
    ForceGC
    Assert.AreEqual instanceCount + 2, YogaNodeInstanceCount()
    node.Clear
    Assert.AreEqual 0, node.Count
    ForceGC
    Assert.AreEqual instanceCount + 1, YogaNodeInstanceCount()
End Sub

Private Sub TestClearWithChildDestructorForGC(Parent As cYogaNode, ByVal instanceCount As Long)
    Dim child As cYogaNode: Set child = YogaNodeNew()
    Assert.AreEqual instanceCount + 1, YogaNodeInstanceCount()
    Parent.Insert 0, child
End Sub

Public Sub TestMeasureFuncWithDestructor()
    ForceGC
    Dim instanceCount As Long: instanceCount = YogaNodeInstanceCount
    Dim Parent As cYogaNode: Set Parent = YogaNodeNew()
    Assert.AreEqual instanceCount + 1, YogaNodeInstanceCount()
    TestMeasureFuncWithDestructorForGC Parent
    ForceGC
    Assert.AreEqual instanceCount + 2, YogaNodeInstanceCount()
    Parent.CalculateLayout
    Assert.AreEqual 120, CLng(Parent.LayoutWidth)
    Assert.AreEqual 130, CLng(Parent.LayoutHeight)
End Sub

Private Sub TestMeasureFuncWithDestructorForGC(Parent As cYogaNode)
    Dim child As cYogaNode: Set child = YogaNodeNew()
    Parent.Insert 0, child
    child.Tag = 2
    child.SetMeasureFunction Array(C_Obj(Me), "MeasureCallback")
End Sub

Public Sub TestLayoutMargin()
    Dim node As cYogaNode: Set node = YogaNodeNew()
    node.Width = 100
    node.Height = 100
    node.MarginStart = 1
    node.MarginEnd = 2
    node.MarginTop = 3
    node.MarginBottom = 4
    node.CalculateLayout
    
    Assert.AreEqual 1, node.LayoutMarginLeft
    Assert.AreEqual 2, node.LayoutMarginRight
    Assert.AreEqual 3, node.LayoutMarginTop
    Assert.AreEqual 4, node.LayoutMarginBottom
End Sub

Public Sub TestLayoutPadding()
    Dim node As cYogaNode: Set node = YogaNodeNew()
    node.Width = 100
    node.Height = 100
    node.PaddingStart = 1
    node.PaddingEnd = 2
    node.PaddingTop = 3
    node.PaddingBottom = 4
    node.CalculateLayout
    
    Assert.AreEqual 1, node.LayoutPaddingLeft
    Assert.AreEqual 2, node.LayoutPaddingRight
    Assert.AreEqual 3, node.LayoutPaddingTop
    Assert.AreEqual 4, node.LayoutPaddingBottom
End Sub

Private Sub ITestFixture_GetTestCases(ByVal Tests As SimplyVBComp.TestFixtureBuilder)
    Tests.Add "TestAddChildGetParent"
    Tests.Add "TestChildren"
    Tests.Add "TestRemoveAtFromEmpty"
    Tests.Add "TestRemoveAtOutOfRange"
    Tests.Add "TestCannotAddChildToMultipleParents"
    Tests.Add "TestReset"
    Tests.Add "TestResetParent"
    Tests.Add "TestResetChild"
    Tests.Add "TestClear"
    Tests.Add "TestMeasureFunc"
    Tests.Add "TestMeasureFuncWithFloat"
    Tests.Add "TestChildWithMeasureFunc"
    Tests.Add "TestMeasureFuncWithChild"
    Tests.Add "TestBaselineFunc"
    Tests.Add "TestPrintOneNode"
    Tests.Add "TestPrintWithLogger"
    Tests.Add "TestPrint"
    Tests.Add "TestCopyStyle"
    Tests.Add "TestCopyConstructor"
    Tests.Add "TestDestructor"
    Tests.Add "TestDestructorWithChildren"
    Tests.Add "TestClearWithChildDestructor"
    Tests.Add "TestMeasureFuncWithDestructor"
    Tests.Add "TestLayoutMargin"
    Tests.Add "TestLayoutPadding"
End Sub

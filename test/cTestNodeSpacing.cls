VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTestNodeSpacing"
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

Public Sub TestObjectInitializer()
    Dim node As cYogaNode: Set node = YogaNodeNew()
    With node
        .Top = 1
        .Bottom = 2
        .Left = 3
        .Right = 4

        .MarginTop = 5
        .MarginBottom = 6
        .MarginLeft = 7
        .MarginRight = 8

        .PaddingTop = 9
        .PaddingBottom = 10
        .PaddingLeft = 11
        .PaddingRight = 12

        .BorderTopWidth = 13
        .BorderBottomWidth = 14
        .BorderLeftWidth = 15
        .BorderRightWidth = 16
    End With

    Assert.AreEqual Array(yogaUnitPoint, 1), node.Top
    Assert.AreEqual Array(yogaUnitPoint, 2), node.Bottom
    Assert.AreEqual Array(yogaUnitPoint, 3), node.Left
    Assert.AreEqual Array(yogaUnitPoint, 4), node.Right

    Assert.AreEqual Array(yogaUnitPoint, 5), node.MarginTop
    Assert.AreEqual Array(yogaUnitPoint, 6), node.MarginBottom
    Assert.AreEqual Array(yogaUnitPoint, 7), node.MarginLeft
    Assert.AreEqual Array(yogaUnitPoint, 8), node.MarginRight

    Assert.AreEqual Array(yogaUnitPoint, 9), node.PaddingTop
    Assert.AreEqual Array(yogaUnitPoint, 10), node.PaddingBottom
    Assert.AreEqual Array(yogaUnitPoint, 11), node.PaddingLeft
    Assert.AreEqual Array(yogaUnitPoint, 12), node.PaddingRight

    Assert.AreEqual 13, node.BorderTopWidth
    Assert.AreEqual 14, node.BorderBottomWidth
    Assert.AreEqual 15, node.BorderLeftWidth
    Assert.AreEqual 16, node.BorderRightWidth
End Sub

Public Sub TestWriteRead()
    Dim node As cYogaNode: Set node = YogaNodeNew()

    node.Top = 1
    node.Bottom = 2
    node.Left = 3
    node.Right = 4

    node.MarginTop = 5
    node.MarginBottom = 6
    node.MarginLeft = 7
    node.MarginRight = 8

    node.PaddingTop = 9
    node.PaddingBottom = 10
    node.PaddingLeft = 11
    node.PaddingRight = 12

    node.BorderTopWidth = 13
    node.BorderBottomWidth = 14
    node.BorderLeftWidth = 15
    node.BorderRightWidth = 16

    Assert.AreEqual Array(yogaUnitPoint, 1), node.Top
    Assert.AreEqual Array(yogaUnitPoint, 2), node.Bottom
    Assert.AreEqual Array(yogaUnitPoint, 3), node.Left
    Assert.AreEqual Array(yogaUnitPoint, 4), node.Right

    Assert.AreEqual Array(yogaUnitPoint, 5), node.MarginTop
    Assert.AreEqual Array(yogaUnitPoint, 6), node.MarginBottom
    Assert.AreEqual Array(yogaUnitPoint, 7), node.MarginLeft
    Assert.AreEqual Array(yogaUnitPoint, 8), node.MarginRight

    Assert.AreEqual Array(yogaUnitPoint, 9), node.PaddingTop
    Assert.AreEqual Array(yogaUnitPoint, 10), node.PaddingBottom
    Assert.AreEqual Array(yogaUnitPoint, 11), node.PaddingLeft
    Assert.AreEqual Array(yogaUnitPoint, 12), node.PaddingRight

    Assert.AreEqual 13, node.BorderTopWidth
    Assert.AreEqual 14, node.BorderBottomWidth
    Assert.AreEqual 15, node.BorderLeftWidth
    Assert.AreEqual 16, node.BorderRightWidth
End Sub

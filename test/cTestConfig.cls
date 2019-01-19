VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTestConfig"
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

Public Sub TestUseWebDefaults()
    Dim config0 As cYogaConfig: Set config0 = YogaConfigNew()
    config0.UseWebDefaults = True
    Dim node0 As cYogaNode: Set node0 = YogaNodeNew(config0)
    Assert.AreEqual YogaFlexDirection.yogaFlexRow, node0.flexDirection

    node0.Reset
    Assert.AreEqual YogaFlexDirection.yogaFlexRow, node0.flexDirection

    Dim config As cYogaConfig: Set config = YogaConfigNew()
    config.UseWebDefaults = True
    Dim node1 As cYogaNode: Set node1 = YogaNodeNew(config)
    Assert.AreEqual YogaFlexDirection.yogaFlexRow, node1.flexDirection

    node1.Reset
    Assert.AreEqual YogaFlexDirection.yogaFlexRow, node1.flexDirection
End Sub

Public Sub TestDefaultConfig()
    Dim node0 As cYogaNode: Set node0 = YogaNodeNew()
    Assert.AreEqual YogaFlexDirection.yogaFlexColumn, node0.flexDirection

    Dim node1 As cYogaNode: Set node1 = YogaNodeNew(YogaConfigNew())
    Assert.AreEqual YogaFlexDirection.yogaFlexColumn, node1.flexDirection
End Sub

Public Sub TestCopyConstructor()
    Dim config0 As cYogaConfig: Set config0 = YogaConfigNew()
    config0.UseWebDefaults = True
    Dim srcNode As cYogaNode: Set srcNode = YogaNodeNew(config0)
    Dim node0 As cYogaNode: Set node0 = YogaNodeClone(srcNode)
    Assert.AreEqual YogaFlexDirection.yogaFlexRow, node0.flexDirection

    node0.flexDirection = YogaFlexDirection.yogaFlexColumn
    Assert.AreEqual YogaFlexDirection.yogaFlexColumn, node0.flexDirection

    node0.Reset
    Assert.AreEqual YogaFlexDirection.yogaFlexRow, node0.flexDirection

    Dim node1 As cYogaNode: Set node1 = YogaNodeClone(srcNode)
    node1.flexDirection = YogaFlexDirection.yogaFlexColumn

    Assert.AreEqual YogaFlexDirection.yogaFlexColumn, node1.flexDirection

    node1.Reset
    Assert.AreEqual YogaFlexDirection.yogaFlexRow, node1.flexDirection
End Sub


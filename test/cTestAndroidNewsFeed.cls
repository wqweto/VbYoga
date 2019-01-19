VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTestAndroidNewsFeed"
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

Public Sub Test_android_news_feed()
    Dim config As cYogaConfig: Set config = YogaConfigNew()

    Dim root As cYogaNode: Set root = YogaNodeNew(config)
    root.AlignContent = YogaAlign.yogaAlignStretch
    root.Width = 1080

    Dim root_child0 As cYogaNode: Set root_child0 = YogaNodeNew(config)
    root.Insert 0, root_child0

    Dim root_child0_child0 As cYogaNode: Set root_child0_child0 = YogaNodeNew(config)
    root_child0_child0.AlignContent = YogaAlign.yogaAlignStretch
    root_child0.Insert 0, root_child0_child0

    Dim root_child0_child0_child0 As cYogaNode: Set root_child0_child0_child0 = YogaNodeNew(config)
    root_child0_child0_child0.AlignContent = YogaAlign.yogaAlignStretch
    root_child0_child0.Insert 0, root_child0_child0_child0

    Dim root_child0_child0_child0_child0 As cYogaNode: Set root_child0_child0_child0_child0 = YogaNodeNew(config)
    root_child0_child0_child0_child0.flexDirection = YogaFlexDirection.yogaFlexRow
    root_child0_child0_child0_child0.AlignContent = YogaAlign.yogaAlignStretch
    root_child0_child0_child0_child0.AlignItems = YogaAlign.yogaAlignFlexStart
    root_child0_child0_child0_child0.MarginStart = 36
    root_child0_child0_child0_child0.MarginTop = 24
    root_child0_child0_child0.Insert 0, root_child0_child0_child0_child0

    Dim root_child0_child0_child0_child0_child0 As cYogaNode: Set root_child0_child0_child0_child0_child0 = YogaNodeNew(config)
    root_child0_child0_child0_child0_child0.flexDirection = YogaFlexDirection.yogaFlexRow
    root_child0_child0_child0_child0_child0.AlignContent = YogaAlign.yogaAlignStretch
    root_child0_child0_child0_child0.Insert 0, root_child0_child0_child0_child0_child0

    Dim root_child0_child0_child0_child0_child0_child0 As cYogaNode: Set root_child0_child0_child0_child0_child0_child0 = YogaNodeNew(config)
    root_child0_child0_child0_child0_child0_child0.AlignContent = YogaAlign.yogaAlignStretch
    root_child0_child0_child0_child0_child0_child0.Width = 120
    root_child0_child0_child0_child0_child0_child0.Height = 120
    root_child0_child0_child0_child0_child0.Insert 0, root_child0_child0_child0_child0_child0_child0

    Dim root_child0_child0_child0_child0_child1 As cYogaNode: Set root_child0_child0_child0_child0_child1 = YogaNodeNew(config)
    root_child0_child0_child0_child0_child1.AlignContent = YogaAlign.yogaAlignStretch
    root_child0_child0_child0_child0_child1.FlexShrink = 1
    root_child0_child0_child0_child0_child1.MarginRight = 36
    root_child0_child0_child0_child0_child1.PaddingLeft = 36
    root_child0_child0_child0_child0_child1.PaddingTop = 21
    root_child0_child0_child0_child0_child1.PaddingRight = 36
    root_child0_child0_child0_child0_child1.PaddingBottom = 18
    root_child0_child0_child0_child0.Insert 1, root_child0_child0_child0_child0_child1

    Dim root_child0_child0_child0_child0_child1_child0 As cYogaNode: Set root_child0_child0_child0_child0_child1_child0 = YogaNodeNew(config)
    root_child0_child0_child0_child0_child1_child0.flexDirection = YogaFlexDirection.yogaFlexRow
    root_child0_child0_child0_child0_child1_child0.AlignContent = YogaAlign.yogaAlignStretch
    root_child0_child0_child0_child0_child1_child0.FlexShrink = 1
    root_child0_child0_child0_child0_child1.Insert 0, root_child0_child0_child0_child0_child1_child0

    Dim root_child0_child0_child0_child0_child1_child1 As cYogaNode: Set root_child0_child0_child0_child0_child1_child1 = YogaNodeNew(config)
    root_child0_child0_child0_child0_child1_child1.AlignContent = YogaAlign.yogaAlignStretch
    root_child0_child0_child0_child0_child1_child1.FlexShrink = 1
    root_child0_child0_child0_child0_child1.Insert 1, root_child0_child0_child0_child0_child1_child1

    Dim root_child0_child0_child1 As cYogaNode: Set root_child0_child0_child1 = YogaNodeNew(config)
    root_child0_child0_child1.AlignContent = YogaAlign.yogaAlignStretch
    root_child0_child0.Insert 1, root_child0_child0_child1

    Dim root_child0_child0_child1_child0 As cYogaNode: Set root_child0_child0_child1_child0 = YogaNodeNew(config)
    root_child0_child0_child1_child0.flexDirection = YogaFlexDirection.yogaFlexRow
    root_child0_child0_child1_child0.AlignContent = YogaAlign.yogaAlignStretch
    root_child0_child0_child1_child0.AlignItems = YogaAlign.yogaAlignFlexStart
    root_child0_child0_child1_child0.MarginStart = 174
    root_child0_child0_child1_child0.MarginTop = 24
    root_child0_child0_child1.Insert 0, root_child0_child0_child1_child0

    Dim root_child0_child0_child1_child0_child0 As cYogaNode: Set root_child0_child0_child1_child0_child0 = YogaNodeNew(config)
    root_child0_child0_child1_child0_child0.flexDirection = YogaFlexDirection.yogaFlexRow
    root_child0_child0_child1_child0_child0.AlignContent = YogaAlign.yogaAlignStretch
    root_child0_child0_child1_child0.Insert 0, root_child0_child0_child1_child0_child0

    Dim root_child0_child0_child1_child0_child0_child0 As cYogaNode: Set root_child0_child0_child1_child0_child0_child0 = YogaNodeNew(config)
    root_child0_child0_child1_child0_child0_child0.AlignContent = YogaAlign.yogaAlignStretch
    root_child0_child0_child1_child0_child0_child0.Width = 72
    root_child0_child0_child1_child0_child0_child0.Height = 72
    root_child0_child0_child1_child0_child0.Insert 0, root_child0_child0_child1_child0_child0_child0

    Dim root_child0_child0_child1_child0_child1 As cYogaNode: Set root_child0_child0_child1_child0_child1 = YogaNodeNew(config)
    root_child0_child0_child1_child0_child1.AlignContent = YogaAlign.yogaAlignStretch
    root_child0_child0_child1_child0_child1.FlexShrink = 1
    root_child0_child0_child1_child0_child1.MarginRight = 36
    root_child0_child0_child1_child0_child1.PaddingLeft = 36
    root_child0_child0_child1_child0_child1.PaddingTop = 21
    root_child0_child0_child1_child0_child1.PaddingRight = 36
    root_child0_child0_child1_child0_child1.PaddingBottom = 18
    root_child0_child0_child1_child0.Insert 1, root_child0_child0_child1_child0_child1

    Dim root_child0_child0_child1_child0_child1_child0 As cYogaNode: Set root_child0_child0_child1_child0_child1_child0 = YogaNodeNew(config)
    root_child0_child0_child1_child0_child1_child0.flexDirection = YogaFlexDirection.yogaFlexRow
    root_child0_child0_child1_child0_child1_child0.AlignContent = YogaAlign.yogaAlignStretch
    root_child0_child0_child1_child0_child1_child0.FlexShrink = 1
    root_child0_child0_child1_child0_child1.Insert 0, root_child0_child0_child1_child0_child1_child0

    Dim root_child0_child0_child1_child0_child1_child1 As cYogaNode: Set root_child0_child0_child1_child0_child1_child1 = YogaNodeNew(config)
    root_child0_child0_child1_child0_child1_child1.AlignContent = YogaAlign.yogaAlignStretch
    root_child0_child0_child1_child0_child1_child1.FlexShrink = 1
    root_child0_child0_child1_child0_child1.Insert 1, root_child0_child0_child1_child0_child1_child1
    root.StyleDirection = YogaDirection.yogaDirLTR
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 1080!, root.LayoutWidth
    Assert.AreEqual 240!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 1080!, root_child0.LayoutWidth
    Assert.AreEqual 240!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 1080!, root_child0_child0.LayoutWidth
    Assert.AreEqual 240!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutY
    Assert.AreEqual 1080!, root_child0_child0_child0.LayoutWidth
    Assert.AreEqual 144!, root_child0_child0_child0.LayoutHeight

    Assert.AreEqual 36!, root_child0_child0_child0_child0.LayoutX
    Assert.AreEqual 24!, root_child0_child0_child0_child0.LayoutY
    Assert.AreEqual 1044!, root_child0_child0_child0_child0.LayoutWidth
    Assert.AreEqual 120!, root_child0_child0_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0_child0_child0.LayoutY
    Assert.AreEqual 120!, root_child0_child0_child0_child0_child0.LayoutWidth
    Assert.AreEqual 120!, root_child0_child0_child0_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child0_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0_child0_child0_child0.LayoutY
    Assert.AreEqual 120!, root_child0_child0_child0_child0_child0_child0.LayoutWidth
    Assert.AreEqual 120!, root_child0_child0_child0_child0_child0_child0.LayoutHeight

    Assert.AreEqual 120!, root_child0_child0_child0_child0_child1.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0_child0_child1.LayoutY
    Assert.AreEqual 72!, root_child0_child0_child0_child0_child1.LayoutWidth
    Assert.AreEqual 39!, root_child0_child0_child0_child0_child1.LayoutHeight

    Assert.AreEqual 36!, root_child0_child0_child0_child0_child1_child0.LayoutX
    Assert.AreEqual 21!, root_child0_child0_child0_child0_child1_child0.LayoutY
    Assert.AreEqual 0!, root_child0_child0_child0_child0_child1_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0_child0_child0_child0_child1_child0.LayoutHeight

    Assert.AreEqual 36!, root_child0_child0_child0_child0_child1_child1.LayoutX
    Assert.AreEqual 21!, root_child0_child0_child0_child0_child1_child1.LayoutY
    Assert.AreEqual 0!, root_child0_child0_child0_child0_child1_child1.LayoutWidth
    Assert.AreEqual 0!, root_child0_child0_child0_child0_child1_child1.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child1.LayoutX
    Assert.AreEqual 144!, root_child0_child0_child1.LayoutY
    Assert.AreEqual 1080!, root_child0_child0_child1.LayoutWidth
    Assert.AreEqual 96!, root_child0_child0_child1.LayoutHeight

    Assert.AreEqual 174!, root_child0_child0_child1_child0.LayoutX
    Assert.AreEqual 24!, root_child0_child0_child1_child0.LayoutY
    Assert.AreEqual 906!, root_child0_child0_child1_child0.LayoutWidth
    Assert.AreEqual 72!, root_child0_child0_child1_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child1_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child1_child0_child0.LayoutY
    Assert.AreEqual 72!, root_child0_child0_child1_child0_child0.LayoutWidth
    Assert.AreEqual 72!, root_child0_child0_child1_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child1_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child1_child0_child0_child0.LayoutY
    Assert.AreEqual 72!, root_child0_child0_child1_child0_child0_child0.LayoutWidth
    Assert.AreEqual 72!, root_child0_child0_child1_child0_child0_child0.LayoutHeight

    Assert.AreEqual 72!, root_child0_child0_child1_child0_child1.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child1_child0_child1.LayoutY
    Assert.AreEqual 72!, root_child0_child0_child1_child0_child1.LayoutWidth
    Assert.AreEqual 39!, root_child0_child0_child1_child0_child1.LayoutHeight

    Assert.AreEqual 36!, root_child0_child0_child1_child0_child1_child0.LayoutX
    Assert.AreEqual 21!, root_child0_child0_child1_child0_child1_child0.LayoutY
    Assert.AreEqual 0!, root_child0_child0_child1_child0_child1_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0_child0_child1_child0_child1_child0.LayoutHeight

    Assert.AreEqual 36!, root_child0_child0_child1_child0_child1_child1.LayoutX
    Assert.AreEqual 21!, root_child0_child0_child1_child0_child1_child1.LayoutY
    Assert.AreEqual 0!, root_child0_child0_child1_child0_child1_child1.LayoutWidth
    Assert.AreEqual 0!, root_child0_child0_child1_child0_child1_child1.LayoutHeight

    root.StyleDirection = YogaDirection.yogaDirRTL
    root.CalculateLayout

    Assert.AreEqual 0!, root.LayoutX
    Assert.AreEqual 0!, root.LayoutY
    Assert.AreEqual 1080!, root.LayoutWidth
    Assert.AreEqual 240!, root.LayoutHeight

    Assert.AreEqual 0!, root_child0.LayoutX
    Assert.AreEqual 0!, root_child0.LayoutY
    Assert.AreEqual 1080!, root_child0.LayoutWidth
    Assert.AreEqual 240!, root_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0.LayoutY
    Assert.AreEqual 1080!, root_child0_child0.LayoutWidth
    Assert.AreEqual 240!, root_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0.LayoutY
    Assert.AreEqual 1080!, root_child0_child0_child0.LayoutWidth
    Assert.AreEqual 144!, root_child0_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child0_child0.LayoutX
    Assert.AreEqual 24!, root_child0_child0_child0_child0.LayoutY
    Assert.AreEqual 1044!, root_child0_child0_child0_child0.LayoutWidth
    Assert.AreEqual 120!, root_child0_child0_child0_child0.LayoutHeight

    Assert.AreEqual 924!, root_child0_child0_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0_child0_child0.LayoutY
    Assert.AreEqual 120!, root_child0_child0_child0_child0_child0.LayoutWidth
    Assert.AreEqual 120!, root_child0_child0_child0_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child0_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0_child0_child0_child0.LayoutY
    Assert.AreEqual 120!, root_child0_child0_child0_child0_child0_child0.LayoutWidth
    Assert.AreEqual 120!, root_child0_child0_child0_child0_child0_child0.LayoutHeight

    Assert.AreEqual 816!, root_child0_child0_child0_child0_child1.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child0_child0_child1.LayoutY
    Assert.AreEqual 72!, root_child0_child0_child0_child0_child1.LayoutWidth
    Assert.AreEqual 39!, root_child0_child0_child0_child0_child1.LayoutHeight

    Assert.AreEqual 36!, root_child0_child0_child0_child0_child1_child0.LayoutX
    Assert.AreEqual 21!, root_child0_child0_child0_child0_child1_child0.LayoutY
    Assert.AreEqual 0!, root_child0_child0_child0_child0_child1_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0_child0_child0_child0_child1_child0.LayoutHeight

    Assert.AreEqual 36!, root_child0_child0_child0_child0_child1_child1.LayoutX
    Assert.AreEqual 21!, root_child0_child0_child0_child0_child1_child1.LayoutY
    Assert.AreEqual 0!, root_child0_child0_child0_child0_child1_child1.LayoutWidth
    Assert.AreEqual 0!, root_child0_child0_child0_child0_child1_child1.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child1.LayoutX
    Assert.AreEqual 144!, root_child0_child0_child1.LayoutY
    Assert.AreEqual 1080!, root_child0_child0_child1.LayoutWidth
    Assert.AreEqual 96!, root_child0_child0_child1.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child1_child0.LayoutX
    Assert.AreEqual 24!, root_child0_child0_child1_child0.LayoutY
    Assert.AreEqual 906!, root_child0_child0_child1_child0.LayoutWidth
    Assert.AreEqual 72!, root_child0_child0_child1_child0.LayoutHeight

    Assert.AreEqual 834!, root_child0_child0_child1_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child1_child0_child0.LayoutY
    Assert.AreEqual 72!, root_child0_child0_child1_child0_child0.LayoutWidth
    Assert.AreEqual 72!, root_child0_child0_child1_child0_child0.LayoutHeight

    Assert.AreEqual 0!, root_child0_child0_child1_child0_child0_child0.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child1_child0_child0_child0.LayoutY
    Assert.AreEqual 72!, root_child0_child0_child1_child0_child0_child0.LayoutWidth
    Assert.AreEqual 72!, root_child0_child0_child1_child0_child0_child0.LayoutHeight

    Assert.AreEqual 726!, root_child0_child0_child1_child0_child1.LayoutX
    Assert.AreEqual 0!, root_child0_child0_child1_child0_child1.LayoutY
    Assert.AreEqual 72!, root_child0_child0_child1_child0_child1.LayoutWidth
    Assert.AreEqual 39!, root_child0_child0_child1_child0_child1.LayoutHeight

    Assert.AreEqual 36!, root_child0_child0_child1_child0_child1_child0.LayoutX
    Assert.AreEqual 21!, root_child0_child0_child1_child0_child1_child0.LayoutY
    Assert.AreEqual 0!, root_child0_child0_child1_child0_child1_child0.LayoutWidth
    Assert.AreEqual 0!, root_child0_child0_child1_child0_child1_child0.LayoutHeight

    Assert.AreEqual 36!, root_child0_child0_child1_child0_child1_child1.LayoutX
    Assert.AreEqual 21!, root_child0_child0_child1_child0_child1_child1.LayoutY
    Assert.AreEqual 0!, root_child0_child0_child1_child0_child1_child1.LayoutWidth
    Assert.AreEqual 0!, root_child0_child0_child1_child0_child1_child1.LayoutHeight
End Sub


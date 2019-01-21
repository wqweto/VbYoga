VERSION 5.00
Object = "{7983BD3B-752A-43EA-9BFF-444BBA1FC293}#5.0#0"; "SimplyVBUnit.Component.ocx"
Begin VB.Form frmTestRunner 
   ClientHeight    =   5532
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   9456
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5532
   ScaleWidth      =   9456
   StartUpPosition =   3  'Windows Default
   Begin SimplyVBComp.UIRunner UIRunner1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      _ExtentX        =   16108
      _ExtentY        =   9123
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmTestRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
'
' frmTestRunner
'
' ** NOTE **
' Please set Tools->Options->General->Error_Trapping to 'Break on Unhandled Errors'
'
' Namespaces Available:
'       Assert.*            ie. Assert.That Value, Iz.EqualTo(5)
'
' Public Functions Availabe:
'       AddTest <TestObject>
'       WriteText "Message"
'       WriteLine "Message"
'
' Adding a test fixture:
'   Use AddTest <object>
'
' Steps to create a TestCase:
'
'   1. Add a new class
'   2. Name it as desired
'   3. (Optionally) Add a Setup/Teardown method to be run before and after every test.
'   4. (Optionally) Add a FixtureSetup/FixtureTeardown method to be run at the
'      before the first test and after the last test.
'   5. Add public Subs of the tests you want run.
'
'      Public Sub MyTest()
'          Assert.That a, Iz.EqualTo(b)
'      End Sub
'

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private m_cCodegen          As Collection

Private Sub Form_Load()
    Call LoadLibrary(App.Path & "\yoga.dll")
    Call LoadLibrary(App.Path & "\..\lib\yoga.dll")
    
'    Set m_cCodegen = New Collection
    AddTest CodegenTestCases(New cTestConfig)
    AddTest CodegenTestCases(New cTestYogaNode)
    AddTest CodegenTestCases(New cTestAbsolutePosition)
    AddTest CodegenTestCases(New cTestFlex)
    AddTest CodegenTestCases(New cTestAlignBaseline)
    AddTest CodegenTestCases(New cTestAlignContent)
    AddTest CodegenTestCases(New cTestAlignItems)
    AddTest CodegenTestCases(New cTestAlignSelf)
    AddTest CodegenTestCases(New cTestAndroidNewsFeed)
    AddTest CodegenTestCases(New cTestBorder)
    AddTest CodegenTestCases(New cTestDimension)
    AddTest CodegenTestCases(New cTestDisplay)
    AddTest CodegenTestCases(New cTestFlexDirection)
    AddTest CodegenTestCases(New cTestFlexWrap)
    AddTest CodegenTestCases(New cTestJustifyContent)
    AddTest CodegenTestCases(New cTestMargin)
    AddTest CodegenTestCases(New cTestMinMaxDimension)
    AddTest CodegenTestCases(New cTestNodeSpacing)
    AddTest CodegenTestCases(New cTestPadding)
    AddTest CodegenTestCases(New cTestPercentage)
    AddTest CodegenTestCases(New cTestRounding)
    AddTest CodegenTestCases(New cTestSizeOverflow)
    If Not m_cCodegen Is Nothing Then
        Clipboard.Clear
        On Error Resume Next
        Clipboard.SetText ConcatCollection(m_cCodegen, vbCrLf)
    End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Form Initialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Initialize()
    Call Me.UIRunner1.Init(App)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Call Unload(Me)
    If KeyCode = vbKeyF5 Then Call UIRunner1.Run
End Sub


Private Function CodegenTestCases(oObj As Object) As Object
    Dim oInfo           As TLI.MemberInfo
    
    Set CodegenTestCases = oObj
    If Not m_cCodegen Is Nothing Then
        m_cCodegen.Add "'--- " & TypeName(oObj)
        m_cCodegen.Add "Private Sub ITestFixture_GetTestCases(ByVal Tests As SimplyVBComp.TestFixtureBuilder)"
        With New TLI.TLIApplication
            For Each oInfo In .InterfaceInfoFromObject(oObj).Members
                If oInfo.InvokeKind = INVOKE_FUNC Then
                    If oInfo.Parameters.Count = 0 Then
                        m_cCodegen.Add Replace("    Tests.Add ""%1""", "%1", oInfo.Name)
                    End If
                End If
            Next
        End With
        m_cCodegen.Add "End Sub"
        m_cCodegen.Add vbNullString
    End If
End Function

Private Function ConcatCollection(oCol As Collection, Optional Separator As String = vbCrLf) As String
    Dim lSize           As Long
    Dim vElem           As Variant
    
    For Each vElem In oCol
        lSize = lSize + Len(vElem) + Len(Separator)
    Next
    If lSize > 0 Then
        ConcatCollection = String$(lSize - Len(Separator), 0)
        lSize = 1
        For Each vElem In oCol
            If lSize <= Len(ConcatCollection) Then
                Mid$(ConcatCollection, lSize, Len(vElem) + Len(Separator)) = vElem & Separator
            End If
            lSize = lSize + Len(vElem) + Len(Separator)
        Next
    End If
End Function

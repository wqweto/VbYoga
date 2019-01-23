Attribute VB_Name = "Module1"
Option Explicit

Private m_lDebugID          As Long
Private m_lDebugCount       As Long

Public Sub DebugInstanceInit(sModuleName As String, sDebugID As String, oObj As Object)
    m_lDebugCount = m_lDebugCount + 1
    m_lDebugID = m_lDebugID + 1
    sDebugID = m_lDebugID
End Sub

Public Sub DebugInstanceTerm(sModuleName As String, sDebugID As String)
    m_lDebugCount = m_lDebugCount - 1
    Debug.Print sModuleName & ".DebugInstanceTerm: " & sDebugID & "/" & m_lDebugCount
End Sub

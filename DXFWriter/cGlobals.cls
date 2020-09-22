VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGlobals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------
'     © 2006, Athanasios Gardos
'e-mail: gardos@hol.gr
'You may freely use, modify and distribute this source code
'
'Last update: November 16, 2006
'Please visit:
'     http://business.hol.gr/gardos/
' or
'     http://avax.invisionzone.com/
'for development tools and more source code
'-----------------------------------------------------------

Option Explicit

Private m_Blocks() As cDrawing

Public Sub SetBlocks(Blocks() As cDrawing)
    m_Blocks() = Blocks()
End Sub

Public Function GetBlockBorder(sName As String, xMin As Double, yMin As Double, zMin As Double, xMax As Double, yMax As Double, zMax As Double) As Boolean
    Dim lMax As Long, lCnt As Long
    On Local Error Resume Next
    lMax = UBound(m_Blocks)
    For lCnt = 1 To lMax
        If Trim$(UCase$(sName)) = Trim$(UCase$(m_Blocks(lCnt).Name)) Then
           Call m_Blocks(lCnt).GetBorder(xMin, yMin, zMin, xMax, yMax, zMax)
           GetBlockBorder = True
           Exit Function
        End If
    Next lCnt
End Function



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPolyline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Private m_count As Long

Public LayerName As String
Public LineTypeName As String
Public ColorIndex As Integer

Private m_Polyline As String
Private m_Vert As cAddString

Private m_xMin As Double
Private m_yMin As Double
Private m_zMin As Double
Private m_xMax As Double
Private m_yMax As Double
Private m_zMax As Double

Friend Property Get xMin() As Double
    xMin = m_xMin
End Property

Friend Property Let xMin(ByVal v As Double)
    m_xMin = v
End Property

Friend Property Get yMin() As Double
    yMin = m_yMin
End Property

Friend Property Let yMin(ByVal v As Double)
    m_yMin = v
End Property

Friend Property Get zMin() As Double
    zMin = m_zMin
End Property

Friend Property Let zMin(ByVal v As Double)
    m_zMin = v
End Property

Friend Property Get xMax() As Double
    xMax = m_xMax
End Property

Friend Property Let xMax(ByVal v As Double)
    m_xMax = v
End Property

Friend Property Get yMax() As Double
    yMax = m_yMax
End Property

Friend Property Let yMax(ByVal v As Double)
    m_yMax = v
End Property

Friend Property Get zMax() As Double
    zMax = m_zMax
End Property

Friend Property Let zMax(ByVal v As Double)
    m_zMax = v
End Property

Private Sub Class_Initialize()
    Set m_Vert = New cAddString
    m_Vert.BeginAdd
    m_count = 0
    LayerName = "0"
    LineTypeName = "CONTINUOUS"
    ColorIndex = 255
End Sub

Public Function InsertVertex(x As Double, y As Double, z As Double) As Boolean
    m_count = m_count + 1
    If m_count = 1 Then
       xMin = x
       yMin = y
       zMin = z
       xMax = x
       yMax = y
       zMax = z
    End If
    If xMin > x Then xMin = x
    If yMin > y Then yMin = y
    If zMin > z Then zMin = z
    If xMax < x Then xMax = x
    If yMax < y Then yMax = y
    If zMax < z Then zMax = z
    m_Vert.Add2Strings "  0", vbCrLf
    m_Vert.Add2Strings "VERTEX", vbCrLf
    m_Vert.Add2Strings "  8", vbCrLf
    If LayerName = "" Then LayerName = "0"
    m_Vert.Add2Strings LayerName, vbCrLf
    m_Vert.Add2Strings " 62", vbCrLf
    m_Vert.Add2Strings Format$(ColorIndex), vbCrLf
    m_Vert.Add2Strings "  6", vbCrLf
    If LineTypeName = "" Then LineTypeName = "CONTINUOUS"
    m_Vert.Add2Strings LineTypeName, vbCrLf
    m_Vert.AddString DxfNb(10, x)
    m_Vert.AddString DxfNb(20, y)
    m_Vert.AddString DxfNb(30, z)
    InsertVertex = True
End Function

Friend Property Get DxfPolyline() As String
    Dim oStr As cAddString
    If m_count = 0 Then Exit Function
    Set oStr = New cAddString
    oStr.BeginAdd
    oStr.Add2Strings "  0", vbCrLf
    oStr.Add2Strings "POLYLINE", vbCrLf
    oStr.Add2Strings "  8", vbCrLf
    If LayerName = "" Then LayerName = "0"
    oStr.Add2Strings LayerName, vbCrLf
    oStr.Add2Strings " 62", vbCrLf
    oStr.Add2Strings Format$(ColorIndex), vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "    1", vbCrLf
    oStr.Add2Strings " 66", vbCrLf
    oStr.Add2Strings "    1", vbCrLf
    oStr.AddString m_Vert.CurString
    oStr.Add2Strings "  0", vbCrLf
    oStr.Add2Strings "SEQEND", vbCrLf
    DxfPolyline = oStr.CurString
    Set oStr = Nothing
End Property

Private Sub Class_Terminate()
    Set m_Vert = Nothing
End Sub

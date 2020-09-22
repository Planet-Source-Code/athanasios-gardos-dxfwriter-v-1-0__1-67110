VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLinearHatch"
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
Private m_HatchObjects As Long

Private xV() As Double
Private yV() As Double
Private zV() As Double

Public LayerName As String
Public LineTypeName As String
Public ColorIndex As Integer
Public Distance As Double
Public Mask As String
Public Angle As Double
Public Outside As Boolean
Public Border As Boolean
Public Áround As Boolean

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
    ReDim xV(100) As Double
    ReDim yV(100) As Double
    ReDim zV(100) As Double
    m_count = 0
    m_HatchObjects = 0
    LayerName = "0"
    LineTypeName = "CONTINUOUS"
    ColorIndex = 255
    Distance = 1
    Angle = 0
    Outside = False
    Border = False
    Áround = False
    Mask = ""
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
    If UBound(xV) < m_count Then
       ReDim Preserve xV(m_count) As Double
       ReDim Preserve yV(m_count) As Double
       ReDim Preserve zV(m_count) As Double
    End If
    xV(m_count) = x
    yV(m_count) = y
    zV(m_count) = z
    InsertVertex = True
End Function

Friend Property Get DxfHatch() As String
    Dim oStr As cAddString
    Dim lCnt As Long
    Dim oLine As cLine
    Dim oPolyLine As cPolyline
    Dim oGeometry As cGeometry
    Dim xV1() As Double, yV1() As Double, zV1() As Double
    Dim xV2() As Double, yV2() As Double, zV2() As Double
    Dim Hin As Integer
    Dim Hbrd As Integer
    If Outside = False Then Hin = 1
    If Border = True Then Hbrd = 1
    If m_count <= 2 Then Exit Property
    ReDim Preserve xV(m_count) As Double
    ReDim Preserve yV(m_count) As Double
    ReDim Preserve zV(m_count) As Double
    Set oGeometry = New cGeometry
    m_HatchObjects = oGeometry.GetHatchBlockLines(xV(), yV(), Distance, Angle, Hin, Hbrd, Mask$, xV1(), yV1(), zV1(), xV2(), yV2(), zV2())
    If m_HatchObjects > 0 Then
       Set oLine = New cLine
       Set oStr = New cAddString
       oStr.BeginAdd
       oLine.LineTypeName = LineTypeName
       oLine.ColorIndex = ColorIndex
       oLine.LayerName = LayerName
       For lCnt = 1 To m_HatchObjects
           oLine.x1 = xV1(lCnt)
           oLine.y1 = yV1(lCnt)
           oLine.x2 = xV2(lCnt)
           oLine.y2 = yV2(lCnt)
           If xMin > xV1(lCnt) Then xMin = xV1(lCnt)
           If yMin > yV1(lCnt) Then yMin = yV1(lCnt)
           If xMax < xV1(lCnt) Then xMax = xV1(lCnt)
           If yMax < yV1(lCnt) Then yMax = yV1(lCnt)
           If xMin > xV2(lCnt) Then xMin = xV2(lCnt)
           If yMin > yV2(lCnt) Then yMin = yV2(lCnt)
           If xMax < xV2(lCnt) Then xMax = xV2(lCnt)
           If yMax < yV2(lCnt) Then yMax = yV2(lCnt)
           oStr.AddString oLine.DxfLine
       Next lCnt
       Set oLine = Nothing
       If Áround = True Then
          m_HatchObjects = m_HatchObjects + 1
          Set oPolyLine = New cPolyline
          oPolyLine.LineTypeName = LineTypeName
          oPolyLine.ColorIndex = ColorIndex
          oPolyLine.LayerName = LayerName
          For lCnt = 1 To m_count
              oPolyLine.InsertVertex xV(lCnt), yV(lCnt), zV(lCnt)
          Next lCnt
          oStr.AddString oPolyLine.DxfPolyline
          Set oPolyLine = Nothing
       End If
       DxfHatch = oStr.CurString
       Set oStr = Nothing
    End If
    Set oGeometry = Nothing
End Property

Friend Property Get HatchObjects() As Long
    HatchObjects = m_HatchObjects
End Property

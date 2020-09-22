VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAngularDimension"
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
Private m_DimensionObjects As Long

Public LayerName As String
Public LineTypeName As String
Public ColorIndex As Integer

Public ax1 As Double
Public ay1 As Double
Public az1 As Double
Public ax2 As Double
Public ay2 As Double
Public az2 As Double

Public bx1 As Double
Public by1 As Double
Public bz1 As Double
Public bx2 As Double
Public by2 As Double
Public bz2 As Double

Public DirX As Double
Public DirY As Double
Public DirZ As Double

Public AngFontIndex As Integer
Public AngTextHeight As Double
Public AngDecimalNum As Integer
Public AngArrowType As Integer
Public AngFactor As Double
Public AngTickSize As Double
Public AngDefaultText As String

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
    m_count = 0
    m_DimensionObjects = 0
    LayerName = "0"
    LineTypeName = "CONTINUOUS"
    ColorIndex = 255
    AngFontIndex = 0
    AngTextHeight = 0.2
    AngDecimalNum = 2
    AngArrowType = 1
    AngTickSize = 0.1
    AngFactor = 1
    AngDefaultText = ""
End Sub

Friend Property Get DxfAngularDimension() As String
    Dim DTStyle%, DTxtScl As Double, DimAkr As Double
    Dim DimSynt As Double, DimArrow%
    Dim oGeometry As cGeometry
    Set oGeometry = New cGeometry
    oGeometry.LayerName = LayerName
    oGeometry.LineTypeName = LineTypeName
    oGeometry.ColorIndex = ColorIndex
    DTStyle% = AngFontIndex
    DTxtScl = AngTextHeight
    DimAkr = AngDecimalNum
    DimSynt = AngFactor
    DimArrow% = AngArrowType
    DxfAngularDimension = oGeometry.PutAngle(ax1, ay1, ax2, ay2, bx1, by1, bx2, by2, DirX, DirY, AngDecimalNum, DimSynt, AngDefaultText, DTStyle%, DTxtScl, DimArrow%, AngTickSize)
    m_DimensionObjects = oGeometry.m_count
    m_xMin = oGeometry.m_xMin
    m_yMin = oGeometry.m_yMin
    m_zMin = oGeometry.m_zMin
    m_xMax = oGeometry.m_xMax
    m_yMax = oGeometry.m_yMax
    m_zMax = oGeometry.m_zMax
    Set oGeometry = Nothing
End Property

Friend Property Get DimensionObjects() As Long
    DimensionObjects = m_DimensionObjects
End Property



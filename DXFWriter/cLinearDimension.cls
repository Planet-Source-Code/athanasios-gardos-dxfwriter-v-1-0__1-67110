VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLinearDimension"
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

Public x1 As Double
Public y1 As Double
Public z1 As Double
Public x2 As Double
Public y2 As Double
Public z2 As Double
Public DirX As Double
Public DirY As Double
Public DirZ As Double

Public Distance As Double
Public DimDist As Double
Public DimTextAbove As Boolean
Public DimTextHorizontal As Boolean
Public DimTextVerticalToLine As Boolean
Public DimArrowSize As Double
Public DimTickSize As Double
Public DimOutExtLines As Double
Public DimInExtLines As Double
Public DimFontIndex As Integer
Public DimTextHeight As Double
Public DimDecimalNum As Integer
Public DimDefText As String
Public DimFactor As Double
Public DimArrowType As Integer

Private DimScale As Double
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
    DimTextAbove = False
    DimTextHorizontal = False
    DimTextVerticalToLine = False
    DimArrowSize = 0.2
    DimTickSize = 0.1
    DimOutExtLines = 0.1
    DimInExtLines = 0.2
    DimFontIndex = 0
    DimTextHeight = 0.2
    DimDecimalNum = 2
    DimDefText = ""
    DimFactor = 1
    DimScale = 1
    DimArrowType = 0
End Sub

Friend Property Get DxfLinearDimension() As String
    Dim mthd%
    Dim x1dist As Double, y1dist As Double, DTStyle%
    Dim xd1 As Double, yd1 As Double, xd2 As Double, yd2 As Double
    Dim DimTadOn%, DimTihOn%, DimAsz As Double
    Dim DimTsz As Double, DimExe As Double, DimExo As Double
    Dim DTxtScl As Double, DimAkr As Double, DimT$
    Dim DimSynt As Double, DimArrow%, DimVertToLine%
    Dim oGeometry As cGeometry
    Set oGeometry = New cGeometry
    oGeometry.LayerName = LayerName
    oGeometry.LineTypeName = LineTypeName
    oGeometry.ColorIndex = ColorIndex
    If DimTextAbove = True Then
       DimTadOn% = 1
    Else
       DimTadOn% = 0
    End If
    If DimTextHorizontal = True Then
       DimTihOn% = 1
    Else
       DimTihOn% = 0
    End If
    If DimTextVerticalToLine = True Then
       DimVertToLine% = 1
    Else
       DimVertToLine% = 0
    End If
    DimAsz = DimArrowSize
    DimTsz = DimTickSize
    DimExe = DimOutExtLines
    DimExo = DimInExtLines
    DTStyle% = DimFontIndex
    DTxtScl = DimTextHeight
    DimAkr = DimDecimalNum
    DimT$ = DimDefText
    DimSynt = DimFactor
    DimArrow% = DimArrowType
    DxfLinearDimension = oGeometry.PutDimension(mthd%, x1, y1, x2, y2, Distance, DirX, DirY, xd1, yd1, xd2, yd2, DimScale, DimTadOn%, DimTihOn%, DimAsz, DimTsz, DimExe, DimExo, DTxtScl, DimAkr, DimT$, DimSynt, DimArrow%, DimVertToLine%, DTStyle%)
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


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegularPolygon"
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

Public x As Double
Public y As Double
Public z As Double
Public R As Double
Public Apexes As Long
Public LayerName As String
Public LineTypeName As String
Public ColorIndex As Integer

Friend Property Get DxfRegularPolygon() As String
    If Apexes <= 0 Then Exit Property
    Dim lCnt As Long
    Dim oPolyLine As cPolyline
    Dim oGeometry As cGeometry
    Dim xV() As Double, yV() As Double, zV() As Double
    Set oPolyLine = New cPolyline
    Set oGeometry = New cGeometry
    Call oGeometry.GetRegularPolygonVertices(x, y, R, Apexes, xV(), yV(), zV())
    oPolyLine.LineTypeName = LineTypeName
    oPolyLine.ColorIndex = ColorIndex
    oPolyLine.LayerName = LayerName
    For lCnt = 1 To Apexes
        oPolyLine.InsertVertex xV(lCnt), yV(lCnt), zV(lCnt)
    Next lCnt
    DxfRegularPolygon = oPolyLine.DxfPolyline
    Set oPolyLine = Nothing
    Set oGeometry = Nothing
End Property

Private Sub Class_Initialize()
    LayerName = "0"
    LineTypeName = "CONTINUOUS"
    ColorIndex = 255
    Apexes = 3
End Sub



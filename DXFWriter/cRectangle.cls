VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRectangle"
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
Public Width As Double
Public Height As Double
Public LayerName As String
Public LineTypeName As String
Public ColorIndex As Integer

Friend Property Get DxfRectangle() As String
    Dim lCnt As Long
    Dim oPolyLine As cPolyline
    Dim oGeometry As cGeometry
    Dim xV() As Double, yV() As Double, zV() As Double
    Set oPolyLine = New cPolyline
    Set oGeometry = New cGeometry
    Call oGeometry.GetBoxVertices(x, y, Width, Height, xV(), yV(), zV())
    oPolyLine.LineTypeName = LineTypeName
    oPolyLine.ColorIndex = ColorIndex
    oPolyLine.LayerName = LayerName
    For lCnt = 1 To 4
        oPolyLine.InsertVertex xV(lCnt), yV(lCnt), zV(lCnt)
    Next lCnt
    DxfRectangle = oPolyLine.DxfPolyline
    Set oPolyLine = Nothing
    Set oGeometry = Nothing
End Property

Private Sub Class_Initialize()
    LayerName = "0"
    LineTypeName = "CONTINUOUS"
    ColorIndex = 255
End Sub

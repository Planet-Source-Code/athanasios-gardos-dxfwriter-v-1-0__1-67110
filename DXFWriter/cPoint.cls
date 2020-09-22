VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPoint"
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
Public LayerName As String
Public LineTypeName As String
Public ColorIndex As Integer

Friend Property Get dxfPoint() As String
    Dim oStr As cAddString
    Set oStr = New cAddString
    oStr.BeginAdd
    oStr.Add2Strings "  0", vbCrLf
    oStr.Add2Strings "POINT", vbCrLf
    oStr.Add2Strings "  8", vbCrLf
    If LayerName = "" Then LayerName = "0"
    oStr.Add2Strings LayerName, vbCrLf
    oStr.Add2Strings " 62", vbCrLf
    oStr.Add2Strings Format$(ColorIndex), vbCrLf
    oStr.Add2Strings "  6", vbCrLf
    If LineTypeName = "" Then LineTypeName = "CONTINUOUS"
    oStr.Add2Strings LineTypeName, vbCrLf
    oStr.AddString DxfNb(10, x)
    oStr.AddString DxfNb(20, y)
    oStr.AddString DxfNb(30, z)
    dxfPoint = oStr.CurString
    Set oStr = Nothing
End Property

Private Sub Class_Initialize()
    LayerName = "0"
    LineTypeName = "CONTINUOUS"
    ColorIndex = 255
End Sub


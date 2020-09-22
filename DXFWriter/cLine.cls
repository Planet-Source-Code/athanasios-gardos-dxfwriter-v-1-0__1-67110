VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLine"
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

Public x1 As Double
Public y1 As Double
Public z1 As Double
Public x2 As Double
Public y2 As Double
Public z2 As Double
Public LayerName As String
Public LineTypeName As String
Public ColorIndex As Integer

Friend Property Get DxfLine() As String
    Dim oStr As cAddString
    Set oStr = New cAddString
    oStr.BeginAdd
    oStr.Add2Strings "  0", vbCrLf
    oStr.Add2Strings "LINE", vbCrLf
    oStr.Add2Strings "  8", vbCrLf
    If LayerName = "" Then LayerName = "0"
    oStr.Add2Strings LayerName, vbCrLf
    oStr.Add2Strings " 62", vbCrLf
    oStr.Add2Strings Format$(ColorIndex), vbCrLf
    oStr.Add2Strings "  6", vbCrLf
    If LineTypeName = "" Then LineTypeName = "CONTINUOUS"
    oStr.Add2Strings LineTypeName, vbCrLf
    oStr.AddString DxfNb(10, x1)
    oStr.AddString DxfNb(20, y1)
    oStr.AddString DxfNb(30, z1)
    oStr.AddString DxfNb(11, x2)
    oStr.AddString DxfNb(21, y2)
    oStr.AddString DxfNb(31, z2)
    DxfLine = oStr.CurString
    Set oStr = Nothing
End Property

Private Sub Class_Initialize()
    LayerName = "0"
    LineTypeName = "CONTINUOUS"
    ColorIndex = 255
End Sub

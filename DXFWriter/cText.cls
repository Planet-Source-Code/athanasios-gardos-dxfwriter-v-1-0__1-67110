VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cText"
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

Public Text As String
Public x As Double
Public y As Double
Public z As Double
Public Height As Double
Public Angle As Double
Public LayerName As String
Public FontIndex As Integer
Private FontName As String
Public ColorIndex As Integer

Private Bold As Boolean
Private Italic As Boolean
Private Strike As Boolean
Private m_Underline As Boolean
Private m_UnderlineDot As Boolean
Private m_UnderlineDouble As Boolean
Private m_UnderlineWords As Boolean

Private Property Get Underline() As Boolean
    Underline = m_Underline
End Property

Private Property Let Underline(ByVal v As Boolean)
    m_Underline = v
    If v = True Then
       UnderlineWords = False
       UnderlineDot = False
       UnderlineDouble = False
    End If
End Property

Private Property Get UnderlineDot() As Boolean
    UnderlineDot = m_UnderlineDot
End Property

Private Property Let UnderlineDot(ByVal v As Boolean)
    m_UnderlineDot = v
    If v = True Then
       Underline = False
       UnderlineDouble = False
    End If
End Property

Private Property Get UnderlineDouble() As Boolean
    UnderlineDouble = m_UnderlineDouble
End Property

Private Property Let UnderlineDouble(ByVal v As Boolean)
    m_UnderlineDouble = v
    If v = True Then
       Underline = False
       UnderlineDot = False
    End If
End Property

Private Property Get UnderlineWords() As Boolean
    UnderlineWords = m_UnderlineWords
End Property

Private Property Let UnderlineWords(ByVal v As Boolean)
    m_UnderlineWords = v
    If v = True Then Underline = False
End Property

Private Sub Class_Initialize()
    Text = ""
    Height = 0.2
    Bold = False
    Italic = False
    Underline = False
    UnderlineDot = False
    UnderlineDouble = False
    UnderlineWords = False
    Strike = False
    LayerName = "0"
    ColorIndex = 255
    FontIndex = 0
    FontName = "STANDARD"
End Sub

Friend Property Get DxfText() As String
    Dim oStr As cAddString
    Set oStr = New cAddString
    oStr.BeginAdd
    oStr.Add2Strings "  0", vbCrLf
    oStr.Add2Strings "TEXT", vbCrLf
    oStr.Add2Strings "  8", vbCrLf
    If LayerName = "" Then LayerName = "0"
    oStr.Add2Strings LayerName, vbCrLf
    oStr.Add2Strings " 62", vbCrLf
    oStr.Add2Strings Format$(ColorIndex), vbCrLf
    oStr.AddString DxfNb(10, x)
    oStr.AddString DxfNb(20, y)
    oStr.AddString DxfNb(30, z)
    oStr.AddString DxfNb(40, Height)
    oStr.Add2Strings "  1", vbCrLf
    oStr.Add2Strings Text, vbCrLf
    oStr.AddString DxfNb(50, Angle)
    oStr.Add2Strings "  41", vbCrLf
    oStr.Add2Strings "1", vbCrLf
    oStr.Add2Strings "  7", vbCrLf
    If FontIndex = 0 Then
       oStr.Add2Strings "STANDARD", vbCrLf
    Else
       oStr.Add2Strings "FONT" & Format$(FontIndex), vbCrLf
    End If
    DxfText = oStr.CurString
    Set oStr = Nothing
End Property


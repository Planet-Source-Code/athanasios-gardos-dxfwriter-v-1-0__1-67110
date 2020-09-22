VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDXF"
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

Private oHeader As cHeader
Private m_Entities As cDrawing
Private m_Blocks() As cDrawing
Private m_FontCharSet As Long

Public Function SetLayers(LayerNames() As String, LayerFlags() As Integer, LayerColorIndex() As Integer, LayerLineTypeName() As String) As Boolean
    SetLayers = oHeader.SetLayers(LayerNames(), LayerFlags(), LayerColorIndex(), LayerLineTypeName())
End Function

Public Function GetLayers(LayerNames() As String, LayerFlags() As Integer, LayerColorIndex() As Integer, LayerLineTypeName() As String) As Long
    GetLayers = oHeader.GetLayers(LayerNames(), LayerFlags(), LayerColorIndex(), LayerLineTypeName())
End Function

Public Function SetFonts(Fonts() As String, FontCharSet As Long) As Boolean
    SetFonts = oHeader.SetFonts(Fonts())
    m_FontCharSet = FontCharSet
End Function

Public Function GetFonts(Fonts() As String, FontCharSet As Long) As Long
    GetFonts = oHeader.GetFonts(Fonts())
    FontCharSet = m_FontCharSet
End Function

Public Function SetLineTypes(LineTypes() As String) As Boolean
    SetLineTypes = oHeader.SetLineTypes(LineTypes())
End Function

Public Function GetLineTypes(LineTypes() As String) As Long
    GetLineTypes = oHeader.GetLineTypes(LineTypes())
End Function

Public Property Let LineTypeScale(ByVal v As Double)
    If v <= 0 Then v = 1
    oHeader.LineTypeScale = v
End Property

Public Property Get LineTypeScale() As Double
    LineTypeScale = oHeader.LineTypeScale
End Property

Public Function InsertEntities(Entities As cDrawing) As Boolean
    If (Entities Is Nothing) Then Exit Function
    Set m_Entities = Entities
    InsertEntities = True
End Function
    
Public Function InsertBlocks(Blocks() As cDrawing) As Boolean
    Dim lMax As Long, lCnt As Long
    Dim xMin As Double, yMin As Double, zMin As Double
    Dim xMax As Double, yMax As Double, zMax As Double
    On Local Error Resume Next
    lMax = UBound(Blocks)
    If lMax = 0 Then Exit Function
    ReDim m_Blocks(lMax) As cDrawing
    For lCnt = 1 To lMax
        If Not (Blocks(lCnt) Is Nothing) Then
           Set m_Blocks(lCnt) = Blocks(lCnt)
           InsertBlocks = True
        End If
    Next lCnt
    Call oGlobals.SetBlocks(Blocks())
End Function
    
Private Sub InitHeader()
    Dim xMin As Double, yMin As Double, zMin As Double
    Dim xMax As Double, yMax As Double, zMax As Double
    Call m_Entities.GetBorder(xMin, yMin, zMin, xMax, yMax, zMax)
    oHeader.EXTMIN_x = xMin
    oHeader.EXTMIN_y = yMin
    oHeader.EXTMIN_z = zMin
    oHeader.EXTMAX_x = xMax
    oHeader.EXTMAX_y = yMax
    oHeader.EXTMAX_z = zMax
End Sub

Private Sub InsertDemo()
    Dim dx As Double
    Dim oLine As cLine
    Dim oText As cText
    If fDemoMode = True Then
       Set oLine = New cLine
       Set oText = New cText
       oLine.x1 = oHeader.EXTMIN_x
       oLine.y1 = oHeader.EXTMIN_y
       oLine.z1 = oHeader.EXTMIN_z
       oLine.x2 = oHeader.EXTMAX_x
       oLine.y2 = oHeader.EXTMIN_y
       oLine.z2 = oHeader.EXTMAX_z
       oLine.ColorIndex = 1
       Call m_Entities.InsertLine(oLine)
       dx = oHeader.EXTMAX_x - oHeader.EXTMIN_x
       oText.Text = D_DemoMSG
       oText.Height = dx / Len(D_DemoMSG)
       oText.x = oHeader.EXTMIN_x
       oText.y = oHeader.EXTMIN_y - 1.1 * oText.Height
       oText.z = oHeader.EXTMIN_z
       oText.ColorIndex = 1
       Call m_Entities.InsertText(oText)
       Set oLine = Nothing
       Set oText = Nothing
    End If
End Sub

Public Function Save(sDXFFile As String) As Boolean
    Dim iFr As Integer
    Dim oStr As cAddString
    Dim lMax As Long, lCnt As Long
    If sDXFFile = "" Then Exit Function
    If (m_Entities Is Nothing) Then Exit Function
    Set oStr = New cAddString
    oStr.BeginAdd
    '------------ Header ---------------
    Call InitHeader
    oStr.AddString oHeader.DxfHeader
    '------------ Tables ---------------
    oStr.AddString oHeader.DxfTables
    '------------ Blocks ---------------
    On Local Error Resume Next
    lMax = UBound(m_Blocks)
    If lMax <> 0 Then
       oStr.Add2Strings "  0", vbCrLf
       oStr.Add2Strings "SECTION", vbCrLf
       oStr.Add2Strings "  2", vbCrLf
       oStr.Add2Strings "BLOCKS", vbCrLf
       For lCnt = 1 To lMax
           If Not m_Blocks(lCnt) Is Nothing Then
              oStr.Add2Strings "  0", vbCrLf
              oStr.Add2Strings "BLOCK", vbCrLf
              oStr.Add2Strings "  2", vbCrLf
              oStr.Add2Strings m_Blocks(lCnt).Name, vbCrLf
              oStr.Add2Strings "  3", vbCrLf
              oStr.Add2Strings m_Blocks(lCnt).Name, vbCrLf
              oStr.Add2Strings "  70", vbCrLf
              oStr.Add2Strings "    0", vbCrLf
              oStr.AddString DxfNb(10, m_Blocks(lCnt).BaseX)
              oStr.AddString DxfNb(20, m_Blocks(lCnt).BaseY)
              oStr.AddString DxfNb(30, m_Blocks(lCnt).BaseZ)
              oStr.AddString m_Blocks(lCnt).DxfDrawing
              oStr.Add2Strings "  0", vbCrLf
              oStr.Add2Strings "ENDBLK", vbCrLf
           End If
       Next lCnt
       oStr.Add2Strings "  0", vbCrLf
       oStr.Add2Strings "ENDSEC", vbCrLf
    End If
    '------------ Entities -------------
    Call InsertDemo
    oStr.Add2Strings "  0", vbCrLf
    oStr.Add2Strings "SECTION", vbCrLf
    oStr.Add2Strings "  2", vbCrLf
    oStr.Add2Strings "ENTITIES", vbCrLf
    oStr.AddString m_Entities.DxfDrawing
    oStr.Add2Strings "  0", vbCrLf
    oStr.Add2Strings "ENDSEC", vbCrLf
    '-----------------------------------
    oStr.Add2Strings "  0", vbCrLf
    oStr.AddString "EOF"
    Call DeleteFile(sDXFFile)
    iFr = FreeFile
    Open sDXFFile For Output As #iFr
    Print #iFr, oStr.CurString
    Close #iFr
    Save = IsFile(sDXFFile)
    Set oStr = Nothing
End Function

Private Sub Class_Initialize()
    Set oGlobals = New cGlobals
    Set oHeader = New cHeader
    LineTypeScale = 1
End Sub

Private Sub Class_Terminate()
    Dim lMax As Long, lCnt As Long
    On Local Error Resume Next
    Set oHeader = Nothing
    Set m_Entities = Nothing
    lMax = UBound(m_Blocks)
    For lCnt = 1 To lMax
        Set m_Blocks(lCnt) = Nothing
    Next lCnt
    Set oGlobals = Nothing
End Sub

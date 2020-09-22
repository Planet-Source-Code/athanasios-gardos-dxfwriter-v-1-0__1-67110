VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Write DXF files"
   ClientHeight    =   1020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1020
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CMD_CreateDXF 
      Caption         =   "Create DXF"
      Height          =   615
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub CMD_CreateDXF_Click()
    Dim sDXFFile As String
    Dim FontCharSet As Long
    Dim oDXF As DXFWriter.cDXF
    Dim Blocks() As DXFWriter.cDrawing
    Dim oDrawing As DXFWriter.cDrawing
    Dim oLine As DXFWriter.cLine
    Dim oPoint As DXFWriter.cPoint
    Dim oCircle As DXFWriter.cCircle
    Dim oText As DXFWriter.cText
    Dim oPolyline As DXFWriter.cPolyline
    Dim oArc As DXFWriter.cArc
    Dim oRegularPolygon As DXFWriter.cRegularPolygon
    Dim oEllipse As DXFWriter.cEllipse
    Dim oSolid As DXFWriter.cSolid
    Dim oRectangle As DXFWriter.cRectangle
    Dim oHatch As DXFWriter.cLinearHatch
    Dim oLinearDimension As DXFWriter.cLinearDimension
    Dim oAngularDimension As DXFWriter.cAngularDimension
    Set oDXF = New DXFWriter.cDXF
    Set oDrawing = New DXFWriter.cDrawing
    Set oLine = New DXFWriter.cLine
    
    ReDim Fonts(2) As String
    Fonts(1) = "Arial"
    Fonts(2) = "Times New Roman"
    Call oDXF.SetFonts(Fonts(), FontCharSet)
    
    sDXFFile = App.Path & "\test.dxf"
  
    oLine.X1 = 10
    oLine.Y1 = 20
    oLine.z1 = 0
    oLine.X2 = 11
    oLine.Y2 = 2.4
    oLine.z2 = 0
    oLine.ColorIndex = 1
    oLine.LineTypeName = "DASHDOT"
    Call oDrawing.InsertLine(oLine)
    oLine.X1 = 12
    oLine.Y1 = 10
    oLine.z1 = 0
    oLine.X2 = 2.3
    oLine.Y2 = 12
    oLine.z2 = 0
    oLine.ColorIndex = 1
    Call oDrawing.InsertLine(oLine)
    ReDim Blocks(1) As cDrawing
    Set Blocks(1) = oDrawing
    Blocks(1).Name = "TEST"
    Call oDXF.InsertBlocks(Blocks())
    
    Set oDrawing = Nothing
    Set oDrawing = New DXFWriter.cDrawing
    Dim x As Double
    oLine.LineTypeName = "CONTINUOUS"
    For x = 0 To 50 Step 10
        oLine.X1 = x
        oLine.Y1 = 20
        oLine.z1 = 0
        oLine.X2 = 11
        oLine.Y2 = 22.4
        oLine.z2 = 0
        oLine.ColorIndex = 2
        Call oDrawing.InsertLine(oLine)
    Next x
    oLine.X1 = 15
    oLine.Y1 = 10
    oLine.z1 = 0
    oLine.X2 = 22.3
    oLine.Y2 = 12
    oLine.z2 = 0
    oLine.ColorIndex = 2
    Call oDrawing.InsertLine(oLine)
    
    Set oCircle = New DXFWriter.cCircle
    oCircle.R = 10
    
    Call oDrawing.InsertCircle(oCircle)
    
    Set oText = New DXFWriter.cText
    oText.Text = "asdfasdfadsf"
    oText.Height = 2
    oText.Angle = 45
    oText.FontIndex = 2
    Call oDrawing.InsertText(oText)
    
    Set oPoint = New DXFWriter.cPoint
    oPoint.ColorIndex = 3
    oPoint.x = 10
    oPoint.y = 10
    Call oDrawing.InsertPoint(oPoint)
    
    
    Set oPolyline = New DXFWriter.cPolyline
    oPolyline.InsertVertex 10, 10, 0
    oPolyline.InsertVertex 0, 10, 0
    oPolyline.InsertVertex 10, 0, 0
    oPolyline.InsertVertex 10, 10, 0
    Call oDrawing.InsertPolyLine(oPolyline)
    
    
    Set oRegularPolygon = New DXFWriter.cRegularPolygon
    oRegularPolygon.Apexes = 5
    oRegularPolygon.R = 10
    Call oDrawing.InsertRegularPolygon(oRegularPolygon)
    
    
    Set oEllipse = New DXFWriter.cEllipse
    oEllipse.aR = 10
    oEllipse.bR = 5
    Call oDrawing.InsertEllipse(oEllipse)
    
    Set oSolid = New DXFWriter.cSolid
    oSolid.ColorIndex = 1
    oSolid.X1 = 0
    oSolid.Y1 = 0
    oSolid.X2 = 10
    oSolid.Y2 = 0
    oSolid.x3 = 10
    oSolid.y3 = 10
    oSolid.x4 = 10
    oSolid.y4 = 0
    Call oDrawing.InsertSolid(oSolid)

    Set oRectangle = New DXFWriter.cRectangle
    oRectangle.Width = 15
    oRectangle.Height = 5
    Call oDrawing.InsertRectangle(oRectangle)
    
    Set oHatch = New DXFWriter.cLinearHatch
    oHatch.InsertVertex 0, 0, 0
    oHatch.InsertVertex 13, 0, 0
    oHatch.InsertVertex 5, 10, 0
    oHatch.InsertVertex 0, 7, 0
    oHatch.Distance = 0.5
    oHatch.ColorIndex = 3
    oHatch.Angle = 245
    oHatch.Áround = True
    oHatch.Outside = True
    oHatch.Border = True
    Call oDrawing.InsertLinearHatch(oHatch)
    
    Set oLinearDimension = New DXFWriter.cLinearDimension
    oLinearDimension.X1 = 10
    oLinearDimension.Y1 = 10
    oLinearDimension.X2 = 20
    oLinearDimension.Y2 = 20
    oLinearDimension.DirX = 5
    oLinearDimension.DirY = 15
    oLinearDimension.Distance = 1
    oLinearDimension.ColorIndex = 3
    Call oDrawing.InsertLinearDimension(oLinearDimension)
    
    Set oAngularDimension = New DXFWriter.cAngularDimension
    oAngularDimension.ax1 = 10
    oAngularDimension.ay1 = 10
    oAngularDimension.ax2 = 20
    oAngularDimension.ay2 = 20
    oAngularDimension.bx1 = 10
    oAngularDimension.by1 = 10
    oAngularDimension.bx2 = 20
    oAngularDimension.by2 = 10
    oAngularDimension.DirX = 15
    oAngularDimension.DirY = 11
    oAngularDimension.AngArrowType = 2
    oAngularDimension.ColorIndex = 3
    Call oDrawing.InsertAngularDimension(oAngularDimension)
    
    Set oAngularDimension = New DXFWriter.cAngularDimension
    oAngularDimension.ax1 = 10
    oAngularDimension.ay1 = 10
    oAngularDimension.ax2 = 20
    oAngularDimension.ay2 = 20
    oAngularDimension.bx1 = 10
    oAngularDimension.by1 = 10
    oAngularDimension.bx2 = 20
    oAngularDimension.by2 = 10
    oAngularDimension.DirX = 9
    oAngularDimension.DirY = 8
    oAngularDimension.AngArrowType = 1
    oAngularDimension.ColorIndex = 3
    Call oDrawing.InsertAngularDimension(oAngularDimension)
    
    Set oArc = New DXFWriter.cArc
    oArc.x = 10
    oArc.y = 10
    oArc.R = 5
    oArc.ColorIndex = 5
    oArc.StartAngle = 45
    oArc.EndAngle = 180
    Call oDrawing.InsertArc(oArc)
    
    Call oDXF.InsertEntities(oDrawing)
    If oDXF.Save(sDXFFile) = True Then
       MsgBox sDXFFile, vbInformation, "DXF is done successfully"
       ShellExecute 0, vbNullString, sDXFFile, vbNullString, vbNullString, 1
    End If
    Set oCircle = Nothing
    Set oLine = Nothing
    Set oPoint = Nothing
    Set oDrawing = Nothing
    Set oText = Nothing
    Set oPolyline = Nothing
    Set oRegularPolygon = Nothing
    Set oEllipse = Nothing
    Set oSolid = Nothing
    Set oArc = Nothing
    Set oRectangle = Nothing
    Set oHatch = Nothing
    Set oLinearDimension = Nothing
    Set oAngularDimension = Nothing
    Set oDXF = Nothing
End Sub

Private Sub Form_Load()
    On Local Error Resume Next
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
End Sub

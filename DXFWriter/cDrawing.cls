VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDrawing"
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

Private m_xMin As Double
Private m_yMin As Double
Private m_zMin As Double
Private m_xMax As Double
Private m_yMax As Double
Private m_zMax As Double

Private m_count As Long

Private m_Drawing As cAddString

Public Name As String
Public BaseX As Double
Public BaseY As Double
Public BaseZ As Double

Private Sub Class_Initialize()
    Name = ""
    m_count = 0
    Set m_Drawing = New cAddString
    m_Drawing.BeginAdd
End Sub

Private Sub Class_Terminate()
    Set m_Drawing = Nothing
End Sub

Friend Sub GetBorder(xMin As Double, yMin As Double, zMin As Double, xMax As Double, yMax As Double, zMax As Double)
    xMin = m_xMin
    yMin = m_yMin
    zMin = m_zMin
    xMax = m_xMax
    yMax = m_yMax
    zMax = m_zMax
End Sub

Public Function InsertCircle(oCircle As cCircle) As Boolean
    If (oCircle Is Nothing) Then Exit Function
    m_Drawing.AddString oCircle.DxfCircle
    m_count = m_count + 1
    If m_count = 1 Then
       m_xMin = oCircle.x
       m_yMin = oCircle.y
       m_zMin = oCircle.z
    End If
    If m_xMin > oCircle.x - oCircle.R Then m_xMin = oCircle.x - oCircle.R
    If m_yMin > oCircle.y - oCircle.R Then m_yMin = oCircle.y - oCircle.R
    If m_zMin > oCircle.z - oCircle.R Then m_zMin = oCircle.z - oCircle.R
    If m_xMax < oCircle.x + oCircle.R Then m_xMax = oCircle.x + oCircle.R
    If m_yMax < oCircle.y + oCircle.R Then m_yMax = oCircle.y + oCircle.R
    If m_zMax < oCircle.z + oCircle.R Then m_zMax = oCircle.z + oCircle.R
    InsertCircle = True
End Function

Public Function InsertText(oText As cText) As Boolean
    Dim dTextLength As Double
    If (oText Is Nothing) Then Exit Function
    m_Drawing.AddString oText.DxfText
    m_count = m_count + 1
    If m_count = 1 Then
       m_xMin = oText.x
       m_yMin = oText.y
       m_zMin = oText.z
    End If
    dTextLength = oText.Height * Len(oText.Text)
    If m_xMin > oText.x - dTextLength Then m_xMin = oText.x - dTextLength
    If m_yMin > oText.y - dTextLength Then m_yMin = oText.y - dTextLength
    If m_zMin > oText.z - dTextLength Then m_zMin = oText.z - dTextLength
    If m_xMax < oText.x + dTextLength Then m_xMax = oText.x + dTextLength
    If m_yMax < oText.y + dTextLength Then m_yMax = oText.y + dTextLength
    If m_zMax < oText.z + dTextLength Then m_zMax = oText.z + dTextLength
    InsertText = True
End Function

Public Function InsertPoint(oPoint As cPoint) As Boolean
    If (oPoint Is Nothing) Then Exit Function
    m_Drawing.AddString oPoint.dxfPoint
    m_count = m_count + 1
    If m_count = 1 Then
       m_xMin = oPoint.x
       m_yMin = oPoint.y
       m_zMin = oPoint.z
    End If
    If m_xMin > oPoint.x Then m_xMin = oPoint.x
    If m_yMin > oPoint.y Then m_yMin = oPoint.y
    If m_zMin > oPoint.z Then m_zMin = oPoint.z
    If m_xMax < oPoint.x Then m_xMax = oPoint.x
    If m_yMax < oPoint.y Then m_yMax = oPoint.y
    If m_zMax < oPoint.z Then m_zMax = oPoint.z
    InsertPoint = True
End Function

Public Function InsertLine(oLine As cLine) As Boolean
    If (oLine Is Nothing) Then Exit Function
    m_Drawing.AddString oLine.DxfLine
    m_count = m_count + 1
    If m_count = 1 Then
       m_xMin = oLine.x1
       m_yMin = oLine.y1
       m_zMin = oLine.z1
       m_xMax = oLine.x2
       m_yMax = oLine.y2
       m_zMax = oLine.z2
    End If
    If m_xMin > oLine.x1 Then m_xMin = oLine.x1
    If m_xMin > oLine.x2 Then m_xMin = oLine.x2
    If m_yMin > oLine.y1 Then m_yMin = oLine.y1
    If m_yMin > oLine.y2 Then m_yMin = oLine.y2
    If m_zMin > oLine.z1 Then m_zMin = oLine.z1
    If m_zMin > oLine.z2 Then m_zMin = oLine.z2
    If m_xMax < oLine.x1 Then m_xMax = oLine.x1
    If m_xMax < oLine.x2 Then m_xMax = oLine.x2
    If m_yMax < oLine.y1 Then m_yMax = oLine.y1
    If m_yMax < oLine.y2 Then m_yMax = oLine.y2
    If m_zMax < oLine.z1 Then m_zMax = oLine.z1
    If m_zMax < oLine.z2 Then m_zMax = oLine.z2
    InsertLine = True
End Function

Public Function InsertPolyLine(oPolyLine As cPolyline) As Boolean
    If (oPolyLine Is Nothing) Then Exit Function
    m_Drawing.AddString oPolyLine.DxfPolyline
    m_count = m_count + 1
    If m_count = 1 Then
       m_xMin = oPolyLine.xMin
       m_yMin = oPolyLine.yMin
       m_zMin = oPolyLine.zMin
       m_xMax = oPolyLine.xMax
       m_yMax = oPolyLine.yMax
       m_zMax = oPolyLine.zMax
    End If
    If m_xMin > oPolyLine.xMin Then m_xMin = oPolyLine.xMin
    If m_xMin > oPolyLine.xMax Then m_xMin = oPolyLine.xMax
    If m_yMin > oPolyLine.yMin Then m_yMin = oPolyLine.yMin
    If m_yMin > oPolyLine.yMax Then m_yMin = oPolyLine.yMax
    If m_zMin > oPolyLine.zMin Then m_zMin = oPolyLine.zMin
    If m_zMin > oPolyLine.zMax Then m_zMin = oPolyLine.zMax
    If m_xMax < oPolyLine.xMin Then m_xMax = oPolyLine.xMin
    If m_xMax < oPolyLine.xMax Then m_xMax = oPolyLine.xMax
    If m_yMax < oPolyLine.yMin Then m_yMax = oPolyLine.yMin
    If m_yMax < oPolyLine.yMax Then m_yMax = oPolyLine.yMax
    If m_zMax < oPolyLine.zMin Then m_zMax = oPolyLine.zMin
    If m_zMax < oPolyLine.zMax Then m_zMax = oPolyLine.zMax
    InsertPolyLine = True
End Function

Public Function InsertArc(oArc As cArc) As Boolean
    If (oArc Is Nothing) Then Exit Function
    m_Drawing.AddString oArc.DxfArc
    m_count = m_count + 1
    If m_count = 1 Then
       m_xMin = oArc.x
       m_yMin = oArc.y
       m_zMin = oArc.z
    End If
    If m_xMin > oArc.x - oArc.R Then m_xMin = oArc.x - oArc.R
    If m_yMin > oArc.y - oArc.R Then m_yMin = oArc.y - oArc.R
    If m_zMin > oArc.z - oArc.R Then m_zMin = oArc.z - oArc.R
    If m_xMax < oArc.x + oArc.R Then m_xMax = oArc.x + oArc.R
    If m_yMax < oArc.y + oArc.R Then m_yMax = oArc.y + oArc.R
    If m_zMax < oArc.z + oArc.R Then m_zMax = oArc.z + oArc.R
    InsertArc = True
End Function

Public Function InsertRegularPolygon(oRegularPolygon As cRegularPolygon) As Boolean
    If (oRegularPolygon Is Nothing) Then Exit Function
    m_Drawing.AddString oRegularPolygon.DxfRegularPolygon
    m_count = m_count + 1
    If m_count = 1 Then
       m_xMin = oRegularPolygon.x
       m_yMin = oRegularPolygon.y
       m_zMin = oRegularPolygon.z
    End If
    If m_xMin > oRegularPolygon.x - oRegularPolygon.R Then m_xMin = oRegularPolygon.x - oRegularPolygon.R
    If m_yMin > oRegularPolygon.y - oRegularPolygon.R Then m_yMin = oRegularPolygon.y - oRegularPolygon.R
    If m_xMax < oRegularPolygon.x + oRegularPolygon.R Then m_xMax = oRegularPolygon.x + oRegularPolygon.R
    If m_yMax < oRegularPolygon.y + oRegularPolygon.R Then m_yMax = oRegularPolygon.y + oRegularPolygon.R
    InsertRegularPolygon = True
End Function

Public Function InsertEllipse(oEllipse As cEllipse) As Boolean
    If (oEllipse Is Nothing) Then Exit Function
    m_Drawing.AddString oEllipse.DxfEllipse
    m_count = m_count + 1
    If m_count = 1 Then
       m_xMin = oEllipse.x
       m_yMin = oEllipse.y
       m_zMin = oEllipse.z
    End If
    If m_xMin > oEllipse.x - oEllipse.aR Then m_xMin = oEllipse.x - oEllipse.aR
    If m_yMin > oEllipse.y - oEllipse.bR Then m_yMin = oEllipse.y - oEllipse.bR
    If m_xMax < oEllipse.x + oEllipse.aR Then m_xMax = oEllipse.x + oEllipse.aR
    If m_yMax < oEllipse.y + oEllipse.bR Then m_yMax = oEllipse.y + oEllipse.bR
    InsertEllipse = True
End Function

Public Function InsertSolid(oSolid As cSolid) As Boolean
    If (oSolid Is Nothing) Then Exit Function
    m_Drawing.AddString oSolid.DxfSolid
    m_count = m_count + 1
    If m_count = 1 Then
       m_xMin = oSolid.x1
       m_yMin = oSolid.y1
       m_zMin = oSolid.z1
       m_xMax = oSolid.x2
       m_yMax = oSolid.y2
       m_zMax = oSolid.z2
    End If
    If m_xMin > oSolid.x1 Then m_xMin = oSolid.x1
    If m_xMin > oSolid.x2 Then m_xMin = oSolid.x2
    If m_yMin > oSolid.y1 Then m_yMin = oSolid.y1
    If m_yMin > oSolid.y2 Then m_yMin = oSolid.y2
    If m_zMin > oSolid.z1 Then m_zMin = oSolid.z1
    If m_zMin > oSolid.z2 Then m_zMin = oSolid.z2
    If m_xMax < oSolid.x1 Then m_xMax = oSolid.x1
    If m_xMax < oSolid.x2 Then m_xMax = oSolid.x2
    If m_yMax < oSolid.y1 Then m_yMax = oSolid.y1
    If m_yMax < oSolid.y2 Then m_yMax = oSolid.y2
    If m_zMax < oSolid.z1 Then m_zMax = oSolid.z1
    If m_zMax < oSolid.z2 Then m_zMax = oSolid.z2
    If m_xMin > oSolid.x3 Then m_xMin = oSolid.x3
    If m_xMin > oSolid.x4 Then m_xMin = oSolid.x4
    If m_yMin > oSolid.y3 Then m_yMin = oSolid.y3
    If m_yMin > oSolid.y4 Then m_yMin = oSolid.y4
    If m_zMin > oSolid.z3 Then m_zMin = oSolid.z3
    If m_zMin > oSolid.z4 Then m_zMin = oSolid.z4
    If m_xMax < oSolid.x3 Then m_xMax = oSolid.x3
    If m_xMax < oSolid.x4 Then m_xMax = oSolid.x4
    If m_yMax < oSolid.y3 Then m_yMax = oSolid.y3
    If m_yMax < oSolid.y4 Then m_yMax = oSolid.y4
    If m_zMax < oSolid.z3 Then m_zMax = oSolid.z3
    If m_zMax < oSolid.z4 Then m_zMax = oSolid.z4
    InsertSolid = True
End Function

Public Function InsertRectangle(oRectangle As cRectangle) As Boolean
    If (oRectangle Is Nothing) Then Exit Function
    m_Drawing.AddString oRectangle.DxfRectangle
    m_count = m_count + 1
    If m_count = 1 Then
       m_xMin = oRectangle.x
       m_yMin = oRectangle.y
       m_zMin = oRectangle.z
    End If
    If m_xMin > oRectangle.x - oRectangle.Width Then m_xMin = oRectangle.x - oRectangle.Width
    If m_yMin > oRectangle.y - oRectangle.Height Then m_yMin = oRectangle.y - oRectangle.Height
    If m_xMax < oRectangle.x + oRectangle.Width Then m_xMax = oRectangle.x + oRectangle.Width
    If m_yMax < oRectangle.y + oRectangle.Height Then m_yMax = oRectangle.y + oRectangle.Height
    InsertRectangle = True
End Function

Public Function InsertLinearHatch(oLinearHatch As cLinearHatch) As Boolean
    If (oLinearHatch Is Nothing) Then Exit Function
    m_Drawing.AddString oLinearHatch.DxfHatch
    m_count = m_count + oLinearHatch.HatchObjects
    If m_count = 1 Then
       m_xMin = oLinearHatch.xMin
       m_yMin = oLinearHatch.yMin
       m_zMin = oLinearHatch.zMin
       m_xMax = oLinearHatch.xMax
       m_yMax = oLinearHatch.yMax
       m_zMax = oLinearHatch.zMax
    End If
    If m_xMin > oLinearHatch.xMin Then m_xMin = oLinearHatch.xMin
    If m_xMin > oLinearHatch.xMax Then m_xMin = oLinearHatch.xMax
    If m_yMin > oLinearHatch.yMin Then m_yMin = oLinearHatch.yMin
    If m_yMin > oLinearHatch.yMax Then m_yMin = oLinearHatch.yMax
    If m_zMin > oLinearHatch.zMin Then m_zMin = oLinearHatch.zMin
    If m_zMin > oLinearHatch.zMax Then m_zMin = oLinearHatch.zMax
    If m_xMax < oLinearHatch.xMin Then m_xMax = oLinearHatch.xMin
    If m_xMax < oLinearHatch.xMax Then m_xMax = oLinearHatch.xMax
    If m_yMax < oLinearHatch.yMin Then m_yMax = oLinearHatch.yMin
    If m_yMax < oLinearHatch.yMax Then m_yMax = oLinearHatch.yMax
    If m_zMax < oLinearHatch.zMin Then m_zMax = oLinearHatch.zMin
    If m_zMax < oLinearHatch.zMax Then m_zMax = oLinearHatch.zMax
    InsertLinearHatch = True
End Function

Public Function InsertLinearDimension(oLinearDimension As cLinearDimension) As Boolean
    If (oLinearDimension Is Nothing) Then Exit Function
    m_Drawing.AddString oLinearDimension.DxfLinearDimension
    m_count = m_count + oLinearDimension.DimensionObjects
    If m_count = 1 Then
       m_xMin = oLinearDimension.xMin
       m_yMin = oLinearDimension.yMin
       m_zMin = oLinearDimension.zMin
       m_xMax = oLinearDimension.xMax
       m_yMax = oLinearDimension.yMax
       m_zMax = oLinearDimension.zMax
    End If
    If m_xMin > oLinearDimension.xMin Then m_xMin = oLinearDimension.xMin
    If m_xMin > oLinearDimension.xMax Then m_xMin = oLinearDimension.xMax
    If m_yMin > oLinearDimension.yMin Then m_yMin = oLinearDimension.yMin
    If m_yMin > oLinearDimension.yMax Then m_yMin = oLinearDimension.yMax
    If m_zMin > oLinearDimension.zMin Then m_zMin = oLinearDimension.zMin
    If m_zMin > oLinearDimension.zMax Then m_zMin = oLinearDimension.zMax
    If m_xMax < oLinearDimension.xMin Then m_xMax = oLinearDimension.xMin
    If m_xMax < oLinearDimension.xMax Then m_xMax = oLinearDimension.xMax
    If m_yMax < oLinearDimension.yMin Then m_yMax = oLinearDimension.yMin
    If m_yMax < oLinearDimension.yMax Then m_yMax = oLinearDimension.yMax
    If m_zMax < oLinearDimension.zMin Then m_zMax = oLinearDimension.zMin
    If m_zMax < oLinearDimension.zMax Then m_zMax = oLinearDimension.zMax
    InsertLinearDimension = True
End Function

Public Function InsertAngularDimension(oAngularDimension As cAngularDimension) As Boolean
    If (oAngularDimension Is Nothing) Then Exit Function
    m_Drawing.AddString oAngularDimension.DxfAngularDimension
    m_count = m_count + oAngularDimension.DimensionObjects
    If m_count = 1 Then
       m_xMin = oAngularDimension.xMin
       m_yMin = oAngularDimension.yMin
       m_zMin = oAngularDimension.zMin
       m_xMax = oAngularDimension.xMax
       m_yMax = oAngularDimension.yMax
       m_zMax = oAngularDimension.zMax
    End If
    If m_xMin > oAngularDimension.xMin Then m_xMin = oAngularDimension.xMin
    If m_xMin > oAngularDimension.xMax Then m_xMin = oAngularDimension.xMax
    If m_yMin > oAngularDimension.yMin Then m_yMin = oAngularDimension.yMin
    If m_yMin > oAngularDimension.yMax Then m_yMin = oAngularDimension.yMax
    If m_zMin > oAngularDimension.zMin Then m_zMin = oAngularDimension.zMin
    If m_zMin > oAngularDimension.zMax Then m_zMin = oAngularDimension.zMax
    If m_xMax < oAngularDimension.xMin Then m_xMax = oAngularDimension.xMin
    If m_xMax < oAngularDimension.xMax Then m_xMax = oAngularDimension.xMax
    If m_yMax < oAngularDimension.yMin Then m_yMax = oAngularDimension.yMin
    If m_yMax < oAngularDimension.yMax Then m_yMax = oAngularDimension.yMax
    If m_zMax < oAngularDimension.zMin Then m_zMax = oAngularDimension.zMin
    If m_zMax < oAngularDimension.zMax Then m_zMax = oAngularDimension.zMax
    InsertAngularDimension = True
End Function

Friend Property Get DxfDrawing() As String
    DxfDrawing = m_Drawing.CurString
End Property


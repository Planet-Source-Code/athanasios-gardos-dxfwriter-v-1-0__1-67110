VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGeometry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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

Public Enum PolygonsComp
    In_p = 1
    Out_p = 2
    Cross_p = 3
    ExtContact_p = 4
End Enum

Public m_xMin As Double
Public m_yMin As Double
Public m_zMin As Double
Public m_xMax As Double
Public m_yMax As Double
Public m_zMax As Double
Public m_count As Long
Public LayerName As String
Public LineTypeName As String
Public ColorIndex As Integer

Private Reslt() As Double
Private Const VertRec% = 8
Const pi# = 3.14159265358979

Private Sub strogilS(bs As Double, akr1 As Double)
    Dim m As Double, k As Double
    If bs > 0 Then
       If akr1 <> 0 Then
          m = 10 ^ akr1
          k = bs * m
          Call strogS(k)
          bs = k / m
        Else
          Call strogS(bs)
       End If
    End If
End Sub

Private Sub strogS(k As Double)
    Dim m%
    If k > 0 Then
       m% = k - Int(k)
       k = Int(k) + m%
    End If
End Sub

Private Sub Strog(k#)
    Dim s1#
    Dim m%
    s1# = Sgn(k#)
    k# = k# * s1#
    m% = k# - Int(k#)
    k# = Int(k#) + m%
    k# = k# * s1#
End Sub

Private Sub Strogil(bs#, akrv1 As Double)
    Dim s1#
    Dim m#
    Dim k#
    s1# = Sgn(bs#)
    bs# = bs# * s1#
    If akrv1 <> 0 Then
       m# = 10 ^ akrv1
       k# = bs# * m#
       Call Strog(k#)
       bs# = k# / m#
    Else
       Call Strog(bs#)
    End If
    bs# = bs# * s1#
End Sub

Private Sub StrogilUp(bs#, akrv1 As Double)
    Dim m#
    Dim k#
    If bs# > 0 Then
        If akrv1 <> 0 Then
            m# = 10 ^ akrv1
            k# = bs# * m#
            Call StrogUp(k#)
            bs# = k# / m#
        Else
            Call StrogUp(bs#)
        End If
    End If
End Sub

Private Sub StrogUp(k#)
    If k# > 0 Then
        If k# > Int(k#) Then
            k# = Int(k#) + 1
        End If
    End If
End Sub

Private Sub LineABC(x1 As Double, y1 As Double, x2 As Double, y2 As Double, a As Double, b As Double, c As Double)
    Dim ax1 As Double, ay1 As Double, ax2 As Double, ay2 As Double
    ax1 = x1
    ay1 = y1
    ax2 = x2
    ay2 = y2
    a = ay1 - ay2
    b = ax2 - ax1
    c = -ax1 * a - ay1 * b
End Sub

Public Sub InterSectLinesX(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double, x4 As Double, y4 As Double, xin As Double, yin As Double, ans%)
    Dim ina%, inb%, denom#
    Dim xt1 As Double, xt2 As Double
    Dim yt1 As Double, yt2 As Double
    Dim A1#, B1#, c1#, A2#, B2#, c2#
    Call LineABC(x1, y1, x2, y2, A1#, B1#, c1#)
    Call LineABC(x3, y3, x4, y4, A2#, B2#, c2#)
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Âñßóêåé óçìåßá ôïìÞò (áí õðÜñ÷ïõí) óå äýï åõè.ôìÞìáôá  '
    '      -1 ôáýôéóç åõèåéþí                               '
    ' ans%=0 ÐáñÜëëçëá                                      '
    '      1 Óçìåßï ôïìÞò ðÜíù óôï ôìÞìá á                  '
    '      2 Óçìåßï ôïìÞò ðÜíù óôï ôìÞìá â                  '
    '      3 Ôá åõè. ôìÞìáôá ôÝìíïíôáé                      '
    '      4 Óçìåßï ôïìÞò åêôüò ôùí åõè.ôìçìÜôùí            '
    '      5 Óçìåßï ôïìÞò óôï Üêñï êáðïéïõ ôìçìáôïò         '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Abs(A1# - A2#) < 0.01 And Abs(B1# - B2#) < 0.01 And Abs(c1# - c2#) < 0.01 Then
        ans% = -1
        Exit Sub
    End If
    If A2# = 0 And B2# = 0 Then
        ans% = -1
        Exit Sub
    End If
    If A2# = 0 Then
        If Abs(B1# / B2# - c1# / c2#) < 0.0001 Then
            ans% = -1
            Exit Sub
        End If
    ElseIf B2# = 0 Then
        If Abs(A1# / A2# - c1# / c2#) < 0.0001 Then
            ans% = -1
            Exit Sub
        End If
    ElseIf c2# = 0 Then
        If Abs(A1# / A2# - B1# / B2#) < 0.0001 Then
            ans% = -1
            Exit Sub
        End If
    Else
        If Abs(A1# / A2# - B1# / B2#) < 0.0001 And Abs(A1# / A2# - c1# / c2#) < 0.0001 Then
            ans% = -1
            Exit Sub
        End If
    End If
    ina% = False
    inb% = False
    ans% = 0
    xin = 0
    yin = 0
    denom = A1# * B2# - A2# * B1#
    If Abs(denom) < 0.01 Then ans% = 0: Exit Sub
    xin = -(c1# * B2# - c2# * B1#) / denom
    yin = -(A1# * c2# - A2# * c1#) / denom
    xt1 = x1: xt2 = x2: yt1 = y1: yt2 = y2
    If yt1 >= yt2 Then
        yt2 = y1
        yt1 = y2
    End If
    If xt1 >= xt2 Then
        xt2 = x1
        xt1 = x2
    End If
    If (xin >= xt1 And xin <= xt2) Or Abs(xin - xt1) <= 0.01 Or Abs(xin - xt2) <= 0.01 Then
       If (yin >= yt1 And yin <= yt2) Or Abs(yin - yt1) <= 0.01 Or Abs(yin - yt2) < 0.01 Then
          ina% = True
       End If
    End If
    xt1 = x3: xt2 = x4: yt1 = y3: yt2 = y4
    If yt1 >= yt2 Then
        yt2 = y3
        yt1 = y4
    End If
    If xt1 >= xt2 Then
        xt2 = x3
        xt1 = x4
    End If
    If (xin >= xt1 And xin <= xt2) Or Abs(xin - xt1) <= 0.01 Or Abs(xin - xt2) < 0.01 Then
        If (yin >= yt1 And yin <= yt2) Or _
            Abs(yin - yt1) <= 0.01 Or _
            Abs(yin - yt2) < 0.01 Then
              inb% = True
        End If
    End If
    If ina% * inb% Then
        ans% = 3
        If (xin = x1 And yin = y1) Or _
           (xin = x2 And yin = y2) Or _
           (xin = x3 And yin = y3) Or _
           (xin = x4 And yin = y4) Then
              ans% = 5
        End If
        Exit Sub
    End If
    If ina% Then
        ans% = 1
        Exit Sub
    End If
    If inb% Then
        ans% = 2
        Exit Sub
    End If
    ans% = 4
End Sub

Public Sub TangetFromPointToCircle(x As Double, y As Double, x0 As Double, y0 As Double, aR As Double, bR As Double, newx As Double, newy As Double, tangDn%)
    Dim R As Double, d As Double, swp%
    Dim x1 As Double, y1 As Double
    Dim x2 As Double, y2 As Double
    Dim D1 As Double, D2 As Double
    Dim vx0#, vy0#, vR#, vX#, vY#, gx#, gy#
    Dim xx0#, yy0#, a#, gR#, b#, c#, vay1#, vay2#
    Dim vax1#, vax2#
    tangDn% = 0
    If FnZero(aR - bR) = 0 Then Exit Sub
    R = aR
    Call Distance(x, y, x0, y0, d)
    If d < R Then Exit Sub
    If FnZero(d - R) = 1 Then
       newx = x
       newy = y
       tangDn% = 1
       Exit Sub
    End If
    Rem .................................
    swp% = 0
    If FnZero(x - x0) = 1 Then swp% = 1
    If swp% = 1 Then Swap x, y: Swap x0, y0
    vx0# = x0: vy0# = y0: vR# = R: vX# = x: vY# = y
    gy# = vY# - vy0#: gx# = vX# - vx0#: xx0# = vX# * vx0#: yy0# = vY# * vy0#
    a# = (gy# / gx#) ^ 2 + 1: gR# = vR# ^ 2 + xx0# - vx0# ^ 2 + yy0# - vy0# ^ 2
    b# = -2 * gy# * gR# / gx# ^ 2: b# = b# + 2 * vx0# * gy# / gx# - 2 * vy0#
    c# = (gR# / gx#) ^ 2 + vx0# ^ 2 + vy0# ^ 2 - vR# ^ 2
    c# = c# - 2 * vx0# * (vR# ^ 2 + xx0# - vx0# ^ 2 + yy0# - vy0# ^ 2) / gx#
    Call SolveSystem2(a#, b#, c#, vay1#, vay2#, tangDn%)
    If tangDn% = 1 Then
       vax1# = (vR# ^ 2 + xx0# - vx0# ^ 2 - vay1# * gy# + yy0# - vy0# ^ 2) / gx#
       vax2# = (vR# ^ 2 + xx0# - vx0# ^ 2 - vay2# * gy# + yy0# - vy0# ^ 2) / gx#
       If swp% = 1 Then Swap vax1#, vay1#: Swap vax2#, vay2#
       x1 = vax1#: y1 = vay1#: x2 = vax2#: y2 = vay2#
       Call Distance(newx, newy, x1, y1, D1)
       Call Distance(newx, newy, x2, y2, D2)
       newx = x1: newy = y1
       If D2 < D1 Then newx = x2: newy = y2
    End If
    If swp% = 1 Then Swap x, y: Swap x0, y0
End Sub

Public Sub QuadrantFromPointToEllipse(x As Double, y As Double, x0 As Double, y0 As Double, aR As Double, bR As Double, newx As Double, newy As Double, QuadDn%)
    Dim R As Double, i%, mini%, mind As Double
    QuadDn% = 0
    If FnZero(aR - bR) = 0 Or FnSamePoint(newx, newy, x0, y0) = 1 Then Exit Sub
    R = aR
    ReDim xr(4) As Double, yr(4) As Double, dxyr(4) As Double
    xr(1) = x0 - R: yr(1) = y0: xr(2) = x0 + R: yr(2) = y0
    xr(3) = x0: yr(3) = y0 - R: xr(4) = x0: yr(4) = y0 + R
    For i% = 1 To 4
        Call Distance(newx, newy, xr(i%), yr(i%), dxyr(i%))
    Next i%
    mini% = 0: mind = dxyr(1)
    For i% = 1 To 4
        If dxyr(i%) <= mind Then mini% = i%: mind = dxyr(i%)
    Next i%
    If mini% <> 0 Then
       QuadDn% = 1
       newx = xr(mini%)
       newy = yr(mini%)
    End If
    Erase xr, yr, dxyr
End Sub

Public Function CircleFromThreePoints(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double, xC As Double, yC As Double, R As Double) As Boolean
    Dim swp%, vx1#, vy1#, vx2#, vy2#
    Dim vx3#, vy3#, vyc#, ad#, vxc#
    CircleFromThreePoints = False
    If FnSamePoint(x1, y1, x2, y2) = 1 Or FnSamePoint(x1, y1, x3, y3) = 1 Or FnSamePoint(x3, y3, x2, y2) = 1 Then Exit Function
    If (FnZero(x1 - x2) = 1 And FnZero(x1 - x3) = 1) Or (FnZero(y1 - y2) = 1 And FnZero(y1 - y3) = 1) Then Exit Function
    If FnZero(x1 - x2) = 1 Then swp% = 1 Else swp% = 0
    If swp% = 1 Then Swap x1, x3: Swap y1, y3
    vx1# = x1: vy1# = y1: vx2# = x2
    vy2# = y2: vx3# = x3: vy3# = y3
    vyc# = (vx1# ^ 2 - vx2# ^ 2 + vy1# ^ 2 - vy2# ^ 2) * (vx3# - vx2#)
    vyc# = vyc# - (vx3# ^ 2 - vx2# ^ 2 + vy3# ^ 2 - vy2# ^ 2) * (vx1# - vx2#)
    ad# = 2 * ((vy1# - vy2#) * (vx3# - vx2#) - (vy3# - vy2#) * (vx1# - vx2#))
    If ad# <> 0 Then vyc# = vyc# / ad# Else GoTo endCircl
    vxc# = vx1# ^ 2 - vx2# ^ 2 + vy1# ^ 2 - vy2# ^ 2 - 2 * vyc# * (vy1# - vy2#)
    ad# = 2 * (vx1# - vx2#)
    If ad# <> 0 Then vxc# = vxc# / ad# Else GoTo endCircl
    xC = vxc#
    yC = vyc#
    Call Distance(x1, y1, xC, yC, R)
    If R <> 0 Then CircleFromThreePoints = True
endCircl:
    If swp% = 1 Then Swap x1, x3: Swap y1, y3
End Function

Public Sub VertFromPointToLine(x0 As Double, y0 As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double, xp As Double, yp As Double)
    Dim fx0#, fy0#, fx1#, fy1#, fx2#, fy2#, dd1#, dd2#, dd3#
    Dim A1#, a#, b#, B1#, k#, fxp#, fyp#
    If x1 = x2 And y1 = y2 Then xp = x1: yp = y1: Exit Sub
    If FnZero(x1 - x2) = 1 Then
       xp = x1: yp = y0
    Else
       If FnZero(y1 - y2) = 1 Then
          xp = x0: yp = y1
       Else
          fx0# = x0: fy0# = y0: fx1# = x1: fy1# = y1: fx2# = x2: fy2# = y2
          dd1# = fy2# * fx1# - fy1# * fx2#: dd2# = fy1# - fy2#: dd3# = fx1# - fx2#
          a# = dd2# / dd3#: b# = dd1# / dd3#
          B1# = dd1# / dd2#: A1# = dd3# / dd2#: k# = B1# + a# * fy0# + fx0#
          fxp# = -k# / (A1# ^ 2 + 1) + a# * fy0# + fx0#: fyp# = k# / (A1# + a#)
          xp = fxp#: yp = fyp#
       End If
    End If
End Sub

Public Sub LineCenter(x1m As Double, y1m As Double, x2m As Double, y2m As Double, xb As Double, yb As Double)
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    If x1m = x2m And y1m = y2m Then xb = x1m: yb = y1m: Exit Sub
    x1 = x1m: y1 = y1m: x2 = x2m: y2 = y2m
    Call SwapSingle(x1, y1, x2, y2)
    xb = x1 + (x2 - x1) / 2
    yb = y1 + (y2 - y1) / 2
End Sub

Public Sub LineEquation(x1 As Double, y1 As Double, x2 As Double, y2 As Double, a As Double, b As Double, c As Double)
    Dim A1#, B1#, c1#
    Call LineABC(x1, y1, x2, y2, A1#, B1#, c1#)
    a = A1#
    b = B1#
    c = c1#
End Sub

Public Function InterSectLines(ByVal xa1 As Double, ByVal ya1 As Double, ByVal xa2 As Double, ByVal ya2 As Double, ByVal xb1 As Double, ByVal yb1 As Double, ByVal xb2 As Double, ByVal yb2 As Double, xn As Double, yn As Double, dne%) As Boolean
    Dim ok As Double
    Call InterSection(xa1, ya1, xa2, ya2, xb1, yb1, xb2, yb2, xn, yn, dne%)
    If dne% = 0 Then Exit Function
    ok = 0
    Call CheckPointOnLine(xa1, ya1, xa2, ya2, xn, yn, ok)
    Call CheckPointOnLine(xb1, yb1, xb2, yb2, xn, yn, ok)
    If ok = 2 Then dne% = 2
    InterSectLines = True
End Function

Private Sub InterSection(xa1 As Double, ya1 As Double, xa2 As Double, ya2 As Double, xb1 As Double, yb1 As Double, xb2 As Double, yb2 As Double, xn As Double, yn As Double, dne%)
    Dim A1 As Double, B1 As Double, g1 As Double
    Dim A2 As Double, B2 As Double, g2 As Double
    Dim da1#, da2#, db1#, db2#, dg1#, dg2#, dd#, dxn#, dyn#
    dne% = 0
    Call LineEquation(xa1, ya1, xa2, ya2, A1, B1, g1)
    If Abs(A1) + Abs(B1) = 0 Then Exit Sub
    Call LineEquation(xb1, yb1, xb2, yb2, A2, B2, g2)
    If Abs(A2) + Abs(B2) = 0 Then Exit Sub
    If FnZeroS(A1 * B2 - B1 * A2) = 1 Then Exit Sub
    If (FnZeroS(xa1 - xb1) = 1 And FnZeroS(ya1 - yb1) = 1) Or (FnZeroS(xa1 - xb2) = 1 And FnZeroS(ya1 - yb2) = 1) Then
       xn = xa1: yn = ya1: dne% = 2: Exit Sub
    ElseIf (FnZeroS(xa2 - xb1) = 1 And FnZeroS(ya2 - yb1) = 1) Or (FnZeroS(xa2 - xb2) = 1 And FnZeroS(ya2 - yb2) = 1) Then
       xn = xa2: yn = ya2: dne% = 2: Exit Sub
    ElseIf (FnZeroS(xa1 - xa2) = 1 And FnZeroS(yb1 - yb2) = 1) Then
       xn = xa1: yn = yb1: dne% = 1: Exit Sub
    ElseIf (FnZeroS(xb1 - xb2) = 1 And FnZeroS(ya1 - ya2) = 1) Then
       xn = xb1: yn = ya1: dne% = 1: Exit Sub
    Else
       da1# = A1: da2# = A2: db1# = B1
       db2# = B2: dg1# = -g1: dg2# = -g2
       dd# = da1# * db2# - db1# * da2#
       If dd# = 0 Then Exit Sub
       dxn# = (db2# * dg1# - db1# * dg2#) / dd#
       dyn# = (da1# * dg2# - da2# * dg1#) / dd#
       xn = dxn#: yn = dyn#: dne% = 1
    End If
End Sub

Public Function ExtendLines(x1 As Double, y1 As Double, x2 As Double, y2 As Double, ix1 As Double, iy1 As Double, ix2 As Double, iy2 As Double) As Boolean
    Dim ok1 As Double, ok2 As Double
    Dim xn As Double, yn As Double, dne%
    Dim x0 As Double, y0 As Double
    ok1 = 0: ok2 = 0
    Call InterSection(x1, y1, x2, y2, ix1, iy1, ix2, iy2, xn, yn, dne%)
    If dne% <> 1 Then Exit Function
    x0 = xn: y0 = yn
    Call CheckPointOnLine(x1, y1, x2, y2, x0, y0, ok1)
    If ok1 = 1 Then Exit Function
    x0 = xn: y0 = yn
    Call CheckPointOnLine(ix1, iy1, ix2, iy2, x0, y0, ok2)
    If ok2 <> 1 Then Exit Function
    Call ExtendLineToPoint(x1, y1, x2, y2, xn, yn)
    ExtendLines = True
End Function

Public Function FilletLines(x11 As Double, y11 As Double, x12 As Double, y12 As Double, etx1 As Double, ety1 As Double, x21 As Double, y21 As Double, x22 As Double, y22 As Double, etx2 As Double, ety2 As Double) As Boolean
    Dim xn As Double, yn As Double, dne%
    Dim ok1 As Double, ok2 As Double, mth As Double
    Dim x0 As Double, y0 As Double
    Call InterSection(x11, y11, x12, y12, x21, y21, x22, y22, xn, yn, dne%)
    If dne% = 0 Then
       Exit Function
    ElseIf dne% = 1 Then
       ok1 = 0: ok2 = 0: mth = 1
       x0 = xn
       y0 = yn
       Call CheckPointOnLine(x11, y11, x12, y12, x0, y0, ok1)
       x0 = xn
       y0 = yn
       Call CheckPointOnLine(x21, y21, x22, y22, x0, y0, ok2)
       If ok1 = 0 Then
          Call ExtendLineToPoint(x11, y11, x12, y12, xn, yn)
       Else
          mth = 0
          Call TrimLineToPoint(mth, etx1, ety1, x11, y11, x12, y12, xn, yn)
       End If
       If ok2 = 0 Then
          Call ExtendLineToPoint(x21, y21, x22, y22, xn, yn)
       Else
          mth = 0
          Call TrimLineToPoint(mth, etx2, ety2, x21, y21, x22, y22, xn, yn)
       End If
       FilletLines = True
    End If
End Function

Public Sub OffsetLinePoint(x1 As Double, y1 As Double, x2 As Double, y2 As Double, OffsetD As Double, x1dist As Double, y1dist As Double, x1d As Double, y1d As Double)
    Dim PdistX As Double, PdistY As Double, xp1 As Double, yp1 As Double
    Call ParallelFromDistance(x1, y1, x2, y2, OffsetD, PdistX, PdistY)
    Call VertFromPointToLine(x1dist, y1dist, x1, y1, x2, y2, xp1, yp1)
    x1d = x1dist - xp1
    y1d = y1dist - yp1
    x1d = Sgn(x1d) * PdistX
    y1d = Sgn(y1d) * PdistY
End Sub

Public Function TrimLines(wx1 As Double, wy1 As Double, wx2 As Double, wy2 As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double, px1 As Double, py1 As Double) As Boolean
    Dim xn As Double, yn As Double, dne%, mth As Double
    Call InterSectLines(wx1, wy1, wx2, wy2, x1, y1, x2, y2, xn, yn, dne%)
    If dne% <> 2 Then Exit Function
    If (FnZero(x1 - xn) = 1 And FnZero(y1 - yn) = 1) Or (FnZero(x2 - xn) = 1 And FnZero(y2 - yn) = 1) Then Exit Function
    mth = 1
    Call TrimLineToPoint(mth, px1, py1, x1, y1, x2, y2, xn, yn)
    TrimLines = True
End Function

Private Sub ExtendLineToPoint(x1 As Double, y1 As Double, x2 As Double, y2 As Double, xn As Double, yn As Double)
    Dim D1 As Double, D2 As Double
    Call Distance(x1, y1, xn, yn, D1)
    Call Distance(x2, y2, xn, yn, D2)
    If D2 > D1 Then
       x1 = xn: y1 = yn
    Else
       x2 = xn: y2 = yn
    End If
End Sub

Private Sub TrimLineToPoint(mth As Double, etx1 As Double, ety1 As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double, xn As Double, yn As Double)
    Dim D1 As Double, D2 As Double
    Call Distance(x1, y1, xn, yn, D1)
    Call Distance(x1, y1, etx1, ety1, D2)
    If mth = 0 Then
       If D2 > D1 Then
          x1 = xn: y1 = yn
       Else
          x2 = xn: y2 = yn
       End If
    Else
       If D2 > D1 Then
          x2 = xn: y2 = yn
       Else
          x1 = xn: y1 = yn
       End If
    End If
End Sub

Public Sub CheckPointOnLine(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x0 As Double, y0 As Double, ok As Double)
    Dim xk1 As Double, yk1 As Double, xk2 As Double, yk2 As Double
    Dim dm As Double, D1 As Double, D2 As Double
    xk1 = x1: yk1 = y1: xk2 = x2: yk2 = y2
    Call SwapSingle(xk1, yk1, xk2, yk2)
    If (xk1 <= x0 And x0 <= xk2) And (yk1 <= y0 And y0 <= yk2) Then
       ok = ok + 1
    ElseIf FnZeroS(x1 - x0) = 1 And FnZeroS(y1 - y0) = 1 Then
       ok = ok + 1: x0 = x1: y0 = y1
    ElseIf FnZeroS(x2 - x0) = 1 And FnZeroS(y2 - y0) = 1 Then
       ok = ok + 1: x0 = x2: y0 = y2
    Else
      Call Distance(x1, y1, x2, y2, dm)
      Call Distance(x1, y1, x0, y0, D1)
      Call Distance(x2, y2, x0, y0, D2)
      If FnZeroS(dm - D1 - D2) = 1 Then ok = ok + 1
    End If
End Sub

Public Function Scale2dMat(sclx As Double, scly As Double, px As Double, py As Double, done%, mat2d() As Double) As Boolean
    Call MATidn(mat2d())
    mat2d(1, 1) = sclx
    mat2d(2, 2) = scly
    mat2d(3, 1) = px * (1 - sclx)
    mat2d(3, 2) = py * (1 - scly)
    mat2d(3, 3) = 1
    done% = 1
    Scale2dMat = True
End Function

Public Function Move2dMat(mvx As Double, mvy As Double, done%, mat2d() As Double) As Boolean
    Call MATidn(mat2d())
    mat2d(3, 1) = mvx
    mat2d(3, 2) = mvy
    done% = 1
    Move2dMat = True
End Function

Public Function Mirror2dMat(mirx1 As Double, miry1 As Double, mirx2 As Double, miry2 As Double, done%, mat2d() As Double) As Boolean
    Dim a As Double, b As Double, g As Double
    Dim theta As Double, xb1 As Double, yb1 As Double
    Dim xb2 As Double, yb2 As Double, m As Double, n As Double, dne%
    done% = 0
    ReDim rfl(3, 3) As Double, rot1(3, 3) As Double
    ReDim rot2(3, 3) As Double, tht1(3, 3) As Double, tht2(3, 3) As Double
    Call MATidn(rfl()): Call MATidn(rot1()): Call MATidn(rot2())
    Call LineEquation(mirx1, miry1, mirx2, miry2, a, b, g)
    If Abs(a) + Abs(b) = 0 Then GoTo endmirror
    If b <> 0 Then theta = -FnArcTan(-a / b) Else theta = -90
    xb1 = 1: yb1 = 0: xb2 = 2: yb2 = 0
    Call InterSectLines(mirx1, miry1, mirx2, miry2, xb1, yb1, xb2, yb2, m, n, dne%)
    If dne% = 0 Then
       xb1 = 0: yb1 = 1: xb2 = 0: yb2 = 2
       Call InterSectLines(mirx1, miry1, mirx2, miry2, xb1, yb1, xb2, yb2, m, n, dne%)
       If dne% = 0 Then GoTo endmirror
       If theta = 0 Then rfl(2, 2) = -1 Else rfl(1, 1) = -1
    Else
       If theta = 0 Then rfl(1, 1) = -1 Else rfl(2, 2) = -1
    End If
    rot1(3, 1) = -m: rot1(3, 2) = -n: rot2(3, 1) = m: rot2(3, 2) = n
    tht1(1, 1) = FnCos(theta): tht1(1, 2) = FnSin(theta)
    tht1(2, 1) = -FnSin(theta): tht1(2, 2) = FnCos(theta): tht1(3, 3) = 1
    tht2(1, 1) = FnCos(theta): tht2(1, 2) = -FnSin(theta)
    tht2(2, 1) = FnSin(theta): tht2(2, 2) = FnCos(theta): tht2(3, 3) = 1
    Call MATmul(rot1(), tht1(), done%): Call MATputRes(rot1(), done%)
    Call MATmul(rot1(), rfl(), done%): Call MATputRes(rot1(), done%)
    Call MATmul(rot1(), tht2(), done%): Call MATputRes(rot1(), done%)
    Call MATmul(rot1(), rot2(), done%): Call MATputRes(mat2d(), done%)
    If done% <> 0 Then Mirror2dMat = True
endmirror:
    Erase rfl, rot1, tht1, rot2, tht2
End Function

Public Function Rot2dMat(theta As Double, rotx As Double, roty As Double, done%, mat2d() As Double) As Boolean
    Call MATidn(mat2d())
    mat2d(1, 1) = FnCos(theta): mat2d(1, 2) = FnSin(theta)
    mat2d(2, 1) = -FnSin(theta): mat2d(2, 2) = FnCos(theta)
    mat2d(3, 1) = -rotx * (FnCos(theta) - 1) + roty * FnSin(theta)
    mat2d(3, 2) = -roty * (FnCos(theta) - 1) - rotx * FnSin(theta)
    mat2d(3, 3) = 1: done% = 1
    Rot2dMat = True
End Function

Public Function Oblique3dMat(alpha As Double, fz As Double, done%, mat3d() As Double) As Boolean
    Call MATidn(mat3d())
    mat3d(3, 3) = 1
    mat3d(3, 1) = -fz * FnCos(alpha)
    mat3d(3, 2) = -fz * FnSin(alpha)
    done% = 1
    Oblique3dMat = True
End Function

Public Function Perspective3dMat(l As Double, m As Double, n As Double, alphax As Double, alphay As Double, zC As Double, done%, mat3d() As Double) As Boolean
    done% = 0
    If zC = 0 Then Exit Function
    ReDim matry(4, 4) As Double
    ReDim matrx(4, 4) As Double
    ReDim prz(4, 4) As Double
    ReDim trxyz(4, 4) As Double
    done% = 0
    matry(1, 1) = FnCos(alphay): matry(1, 3) = -FnSin(alphay): matry(2, 2) = 1
    matry(3, 1) = FnSin(alphay): matry(3, 3) = FnCos(alphay): matry(4, 4) = 1
    matrx(1, 1) = 1: matrx(2, 2) = FnCos(alphax): matrx(2, 3) = FnSin(alphax)
    matrx(3, 2) = -FnSin(alphax): matrx(3, 3) = FnCos(alphax): matrx(4, 4) = 1
    prz(1, 1) = 1: prz(2, 2) = 1: prz(3, 3) = 1: prz(3, 4) = -1 / zC: prz(4, 4) = 1
    Call MATidn(trxyz()): trxyz(4, 1) = l: trxyz(4, 2) = m: trxyz(4, 3) = n
    Call MATmul(matry(), matrx(), done%): Call MATputRes(mat3d(), done%)
    Call MATmul(mat3d(), trxyz(), done%): Call MATputRes(mat3d(), done%)
    Call MATmul(mat3d(), prz(), done%): Call MATputRes(mat3d(), done%)
    If done% <> 0 Then Perspective3dMat = True
    Erase matry, matrx, prz, trxyz
End Function

Public Function Axonometric3dMat(alphax As Double, alphay As Double, done%, mat3d() As Double) As Boolean
    ReDim matry(4, 4) As Double
    ReDim matrx(4, 4) As Double
    done% = 0
    matry(1, 1) = FnCos(alphay): matry(1, 3) = -FnSin(alphay): matry(2, 2) = 1
    matry(3, 1) = FnSin(alphay): matry(3, 3) = FnCos(alphay): matry(4, 4) = 1
    matrx(1, 1) = 1: matrx(2, 2) = FnCos(alphax): matrx(2, 3) = FnSin(alphax)
    matrx(3, 2) = -FnSin(alphax): matrx(3, 3) = FnCos(alphax): matrx(4, 4) = 1
    Call MATmul(matry(), matrx(), done%): Call MATputRes(mat3d(), done%)
    If done% <> 0 Then Axonometric3dMat = True
    Erase matry, matrx
End Function

Public Sub Distance(x1 As Double, y1 As Double, x2 As Double, y2 As Double, d As Double)
    Dim xx1#, xx2#, yy1#, yy2#, dd#
    xx1# = x1
    yy1# = y1
    xx2# = x2
    yy2# = y2
    dd# = Sqr((xx1# - xx2#) ^ 2# + (yy1# - yy2#) ^ 2#)
    d = dd#
End Sub

Private Function DistanceX(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
    Dim xx1#, xx2#, yy1#, yy2#, dd#
    xx1# = x1
    yy1# = y1
    xx2# = x2
    yy2# = y2
    DistanceX = Sqr((xx1# - xx2#) ^ 2# + (yy1# - yy2#) ^ 2#)
End Function

Public Sub ParallelFromDistance(x1 As Double, y1 As Double, x2 As Double, y2 As Double, OffsetD As Double, PdistX As Double, PdistY As Double)
    Dim a As Double, b As Double, g As Double
    Dim theta As Double
    PdistX = 0: PdistY = 0
    If FnZero(x1 - x2) = 1 Then
       PdistX = OffsetD
    ElseIf FnZero(y1 - y2) = 1 Then
       PdistY = OffsetD
    Else
      Call LineEquation(x1, y1, x2, y2, a, b, g)
      If Abs(a) + Abs(b) = 0 Then Exit Sub
      If b <> 0 Then theta = FnArcTan(-a / b) Else theta = 90
      PdistX = FnSin(Abs(theta)) * OffsetD
      PdistY = FnCos(Abs(theta)) * OffsetD
    End If
End Sub

Private Sub FindRealAngle(x As Double, y As Double, xC As Double, yC As Double, theta As Double)
    If xC = x And yC = y Then
       theta = 0
       Exit Sub
    ElseIf FnZero(xC - x) = 1 Then
       If yC < y Then
          theta = 90
       Else
          theta = 270
       End If
       Exit Sub
    ElseIf FnZero(yC - y) = 1 Then
       If xC < x Then
          theta = 0
       Else
          theta = 180
       End If
       Exit Sub
    Else
       theta = FnArcTan((yC - y) / (xC - x))
    End If
    theta = Abs(theta)
    If x >= xC And y < yC Then
       theta = 360 - theta
    ElseIf x < xC And y < yC Then
       theta = 180 + theta
    ElseIf x < xC And y >= yC Then
       theta = 180 - theta
    End If
End Sub

Public Sub Transform2xy(x As Double, y As Double, mat2() As Double)
    Dim done%
    ReDim tr(1, 3) As Double
    tr(1, 1) = x
    tr(1, 2) = y
    tr(1, 3) = 1
    Call MATmul(tr(), mat2(), done%)
    Call MATputRes(tr(), done%)
    If done% = 1 Then x = tr(1, 1): y = tr(1, 2)
    Erase tr
End Sub

Public Sub MATmul(a() As Double, b() As Double, done%)
    Dim g1 As Double
    Dim g2 As Double
    Dim s1 As Double
    Dim s2 As Double
    Dim i%, j%, k%
    If done% = -1 Then Exit Sub
    done% = 0
    g1 = UBound(a, 1)
    s1 = UBound(a, 2)
    g2 = UBound(b, 1)
    s2 = UBound(b, 2)
    If s1 <> g2 Then done% = -1: Exit Sub
    ReDim Reslt(g1, s2)
    For i% = 1 To g1
        For j% = 1 To s2
            For k% = 1 To s1
                Reslt(i%, j%) = Reslt(i%, j%) + a(i%, k%) * b(k%, j%)
            Next k%
        Next j%
    Next i%
    done% = 1
End Sub

Public Sub MATidn(a() As Double)
    Dim g1 As Double
    Dim s1 As Double
    Dim i%, j%
    g1 = UBound(a, 1)
    s1 = UBound(a, 2)
    For i% = 1 To g1
        For j% = 1 To s1
            If i% = j% Then a(i%, j%) = 1 Else a(i%, j%) = 0
        Next j%
    Next i%
End Sub

Public Sub MATputRes(a() As Double, done%)
    Dim g1 As Double
    Dim s1 As Double
    Dim i%, j%
    If done% = 1 Then
       g1 = UBound(Reslt, 1)
       s1 = UBound(Reslt, 2)
       For i% = 1 To g1
           For j% = 1 To s1
               a(i%, j%) = Reslt(i%, j%)
           Next j%
       Next i%
    End If
    Erase Reslt
End Sub

Public Function FnCos(a As Double) As Double
    FnCos = Cos((CDbl(a) / 180#) * pi#)
End Function

Public Function FnSin(a As Double) As Double
    FnSin = Sin((CDbl(a) / 180#) * pi#)
End Function

Private Function FnSinD(a As Double) As Double
    FnSinD = Sin((a / 180#) * pi#)
End Function

Private Function FnCosD(a As Double) As Double
    FnCosD = Cos((a / 180#) * pi#)
End Function

Public Function FnArcTan(a As Double) As Double
    FnArcTan = Atn(a) * 180# / pi#
End Function

Public Function FnArcSin(a As Double) As Double
    Dim b As Double
    If a = 1 Or a = -1 Then
       FnArcSin = a * 90!
    ElseIf Abs(a) < 1 Then
       b = a / Sqr(1 - a ^ 2!)
       FnArcSin = FnArcTan(b)
    Else
       FnArcSin = 0
    End If
End Function

Public Function FnTan(a As Double) As Double
    FnTan = Tan((CDbl(a) / 180#) * pi#)
End Function

Public Sub LineInclination(x As Double, y As Double, xC As Double, yC As Double, Angle As Double)
    Call FindRealAngle(x, y, xC, yC, Angle)
End Sub

Public Sub Angle2Lines(ByVal px As Double, ByVal py As Double, ByVal xa1 As Double, ByVal ya1 As Double, ByVal xa2 As Double, ByVal ya2 As Double, ByVal xb1 As Double, ByVal yb1 As Double, ByVal xb2 As Double, ByVal yb2 As Double, Angle As Double, ok As Double)
    Dim theta As Double, thetb As Double, dne%, xn As Double, yn As Double
    Dim ax2 As Double, ay2 As Double, bx2 As Double, by2 As Double
    Dim theta1 As Double, theta2 As Double, theta3 As Double, R As Double
    Call FindRealAngle(xa1, ya1, xa2, ya2, theta)
    Call FindRealAngle(xb1, yb1, xb2, yb2, thetb)
    ok = 0
    If FnZero(theta - thetb) = 1 Then Exit Sub
    dne% = 0
    Call InterSectLines(xa1, ya1, xa2, ya2, xb1, yb1, xb2, yb2, xn, yn, dne%)
    If dne% = 0 Then Exit Sub
    ax2 = xa1: ay2 = ya1
    If FnSamePoint(xa1, ya1, xn, yn) = 1 Then ax2 = xa2: ay2 = ya2
    bx2 = xb1: by2 = yb1
    If FnSamePoint(xb1, yb1, xn, yn) = 1 Then bx2 = xb2: by2 = yb2
    Call FindRealAngle(ax2, ay2, xn, yn, theta1)
    Call FindRealAngle(bx2, by2, xn, yn, theta2)
    Call FindRealAngle(px, py, xn, yn, theta3)
    If theta1 > theta2 Then Swap theta1, theta2
    ok = 1
    If theta3 >= theta1 And theta3 <= theta2 Then Angle = theta2 - theta1 Else Angle = 360 - (theta2 - theta1)
    Call Distance(px, py, xn, yn, R)
    Rem -----
    xa1 = xn + FnCos(theta1) * R
    ya1 = yn + FnSin(theta1) * R
    xa2 = xn + FnCos(theta2) * R
    ya2 = yn + FnSin(theta2) * R
    xb1 = xn
    yb1 = yn
    xb2 = R
    Rem -----
End Sub

Public Sub PlaneEquation(x1 As Double, y1 As Double, z1 As Double, x2 As Double, y2 As Double, z2 As Double, x3 As Double, y3 As Double, z3 As Double, a As Double, b As Double, c As Double, d As Double)
    Dim aMin As Double
    a = (y1 - y2) * (z1 + z2) + (y2 - y3) * (z2 + z3) + (y3 - y1) * (z3 + z1)
    b = (z1 - z2) * (x1 + x2) + (z2 - z3) * (x2 + x3) + (z3 - z1) * (x3 + x1)
    c = (x1 - x2) * (y1 + y2) + (x2 - x3) * (y2 + y3) + (x3 - x1) * (y3 + y1)
    If a <> 0 And Abs(a) <= Abs(b) And Abs(a) <= Abs(c) Then
       aMin = a
    ElseIf b <> 0 And Abs(b) <= Abs(a) And Abs(b) <= Abs(c) Then
       aMin = b
    Else
       aMin = c
    End If
    If aMin <> 0 Then a = a / aMin: b = b / aMin: c = c / aMin
    d = -(a * x1 + b * y1 + c * z1)
End Sub

Public Sub Solve2x2(a() As Double, x() As Double, nosolve%)
    Dim d As Double, dx As Double, dy As Double
    nosolve% = 0
    d = a(1, 1) * a(2, 2) - a(2, 1) * a(1, 2)
    If d = 0 Then nosolve% = 1: Exit Sub
    dx = a(1, 3) * a(2, 2) - a(2, 3) * a(1, 2)
    dy = a(1, 1) * a(2, 3) - a(2, 1) * a(1, 3)
    x(1) = dx / d: x(2) = dy / d
End Sub

Public Sub IntersecLineAnd2Planes(A1 As Double, B1 As Double, c1 As Double, D1 As Double, A2 As Double, B2 As Double, c2 As Double, D2 As Double, x1 As Double, y1 As Double, z1 As Double, x2 As Double, y2 As Double, z2 As Double, xC As Double, yC As Double, zC As Double, okd%)
    Dim dx As Double, dy As Double, dz As Double, nosolve%, dd As Double
    Dim zc1 As Double, zc2 As Double
    okd% = 0
    ReDim aSt(2, 3) As Double
    ReDim xSt(2) As Double
    dx = x2 - x1: dy = y2 - y1: dz = z2 - z1
    If FnZeroS(dz) = 1 And FnZeroS(dy) = 1 And FnZeroS(A2) = 0 Then
       zC = z1: yC = y1: xC = -(B2 * yC + c2 * zC + D2) / A2
    ElseIf FnZeroS(dx) = 1 And FnZeroS(dz) = 1 And FnZeroS(B2) = 0 Then
       xC = x1: zC = z1: yC = -(c2 * zC + A2 * xC + D2) / B2
    ElseIf FnZeroS(dx) = 1 And FnZeroS(dy) = 1 And FnZeroS(c2) = 0 Then
       xC = x1: yC = y1: zC = -(B2 * yC + A2 * xC + D2) / c2
    ElseIf FnZeroS(dz) = 1 Then
       zC = z1
       aSt(1, 1) = A1: aSt(1, 2) = B1: aSt(1, 3) = -(c1 * zC + D1)
       aSt(2, 1) = A2: aSt(2, 2) = B2: aSt(2, 3) = -(c2 * zC + D2)
       Call Solve2x2(aSt(), xSt(), nosolve%)
       If nosolve% = 1 Then GoTo telostomi
       xC = xSt(1): yC = xSt(2)
    ElseIf FnZeroS(dx) = 1 Then
       xC = x1
       aSt(1, 1) = B1: aSt(1, 2) = c1: aSt(1, 3) = -(A1 * xC + D1)
       aSt(2, 1) = B2: aSt(2, 2) = c2: aSt(2, 3) = -(A2 * xC + D2)
       Call Solve2x2(aSt(), xSt(), nosolve%)
       If nosolve% = 1 Then GoTo telostomi
       yC = xSt(1): zC = xSt(2)
    ElseIf FnZeroS(dy) = 1 Then
       yC = y1
       aSt(1, 1) = A1: aSt(1, 2) = c1: aSt(1, 3) = -(B1 * yC + D1)
       aSt(2, 1) = A2: aSt(2, 2) = c2: aSt(2, 3) = -(B2 * yC + D2)
       Call Solve2x2(aSt(), xSt(), nosolve%)
       If nosolve% = 1 Then GoTo telostomi
       xC = xSt(1): zC = xSt(2)
    ElseIf FnZeroS(dx) = 0 Then
       dd = dz / dx
       aSt(1, 1) = A1 + c1 * dd: aSt(1, 2) = B1: aSt(1, 3) = -(c1 * (z1 - x1 * dd) + D1)
       aSt(2, 1) = A2 + c2 * dd: aSt(2, 2) = B2: aSt(2, 3) = -(c2 * (z2 - x2 * dd) + D2)
       Call Solve2x2(aSt(), xSt(), nosolve%)
       If nosolve% = 1 Then GoTo telostomi
       xC = xSt(1): yC = xSt(2): zc1 = (xC - x1) * dd + z1: zc2 = (xC - x2) * dd + z2
       If FnZeroM1(zc1 - zc2) = 0 Then GoTo telostomi
       zC = zc1: If FnZeroM1((xC - x1) / dx - (zC - z1) / dz) = 0 Then GoTo telostomi
    ElseIf FnZeroS(dy) = 0 Then
       dd = dz / dy
       aSt(1, 1) = A1: aSt(1, 2) = B1 + c1 * dd: aSt(1, 3) = -(c1 * (z1 - y1 * dd) + D1)
       aSt(2, 1) = A2: aSt(2, 2) = B2 + c2 * dd: aSt(2, 3) = -(c2 * (z2 - y2 * dd) + D2)
       Call Solve2x2(aSt(), xSt(), nosolve%)
       If nosolve% = 1 Then GoTo telostomi
       xC = xSt(1): yC = xSt(2): zc1 = (xC - y1) * dd + z1: zc2 = (xC - y2) * dd + z2
       If FnZeroM1(zc1 - zc2) = 0 Then GoTo telostomi
       zC = zc1: If FnZeroM1((yC - y1) / dy - (zC - z1) / dz) = 0 Then GoTo telostomi
    Else
       GoTo telostomi
    End If
    okd% = 1
telostomi:
    Erase aSt, xSt
End Sub

Private Sub Swap(a As Variant, b As Variant)
    Dim c As Variant
    c = b
    b = a
    a = c
End Sub

Private Function FnZero(a As Double) As Double
    If Abs(a) < 0.000001 Then FnZero = 1 Else FnZero = 0
End Function

Private Function FnZeroM(a As Double) As Double
    If Abs(a) < 0.00001 Then FnZeroM = 1 Else FnZeroM = 0
End Function

Private Function FnZeroS(a As Double) As Double
    If Abs(a) < 0.0000001 Then
       FnZeroS = 1
    Else
       FnZeroS = 0
    End If
End Function

Private Function FnZeroM1(a)
    If Abs(a) < 0.00005 Then FnZeroM1 = 1 Else FnZeroM1 = 0
End Function

Private Function FnSamePoint(xa1 As Double, ya1 As Double, xb1 As Double, yb1 As Double) As Double
    If FnZero(xa1 - xb1) = 1 And FnZero(ya1 - yb1) = 1 Then
       FnSamePoint = 1
    Else
       FnSamePoint = 0
    End If
End Function

Private Sub Equal4Double(pmnx1 As Double, pmny1 As Double, pmxx1 As Double, pmxy1 As Double, pmnx As Double, pmny As Double, pmxx As Double, pmxy As Double)
    pmnx1 = pmnx
    pmny1 = pmny
    pmxx1 = pmxx
    pmxy1 = pmxy
End Sub

Private Sub SwapSingle(xx1 As Double, yy1 As Double, xx2 As Double, yy2 As Double)
    If xx1 > xx2 Then Swap xx1, xx2
    If yy1 > yy2 Then Swap yy1, yy2
End Sub

Private Sub SwapInteger(xx1%, yy1%, xx2%, yy2%)
    If xx1% > xx2% Then Swap xx1%, xx2%
    If yy1% > yy2% Then Swap yy1%, yy2%
End Sub

Private Sub SolveSystem2(a#, b#, c#, y1#, y2#, SolveDn%)
    Dim Diak#
    SolveDn% = 1
    Diak# = b# ^ 2 - 4 * a# * c#
    If a# = 0# Or Diak# < 0# Then SolveDn% = 0: Exit Sub
    y1# = (-b# + Sqr(Diak#)) / (2 * a#)
    y2# = (-b# - Sqr(Diak#)) / (2 * a#)
End Sub

Private Function Cvd(s As String) As Double
    Dim A2 As Double
    Dim MyByte() As Byte
    If Len(s) = 0 Then Exit Function
    MyByte() = StrConv(s, vbFromUnicode, 0)
    Call CopyMemory(ByVal VarPtr(A2), ByVal VarPtr(MyByte(0)), CLng(VertRec%))
    Cvd = A2
End Function

Private Function Mkd(A2 As Double) As String
    Dim MyByte(VertRec% - 1) As Byte
    Call CopyMemory(ByVal VarPtr(MyByte(0)), ByVal VarPtr(A2), CLng(VertRec%))
    Mkd = StrConv(MyByte, vbUnicode, 0)
End Function

Private Sub FindMaxMin(ax1 As Double, amaxx As Double, aminx As Double)
    If ax1 > amaxx Then amaxx = ax1
    If ax1 < aminx Then aminx = ax1
End Sub

Private Sub FindMaxMinXY(sx1 As Double, sy1 As Double, MaxPx As Double, MinPx As Double, MaxPy As Double, MinPy As Double)
    Call FindMaxMin(sx1, MaxPx, MinPx)
    Call FindMaxMin(sy1, MaxPy, MinPy)
End Sub

Private Sub GetFourCvs(Plane$, va1 As Double, vb1 As Double, vc1 As Double, vd1 As Double)
    va1 = Cvd(Mid$(Plane$, 1, VertRec%))
    vb1 = Cvd(Mid$(Plane$, VertRec% + 1, VertRec%))
    vc1 = Cvd(Mid$(Plane$, 2 * VertRec% + 1, VertRec%))
    vd1 = Cvd(Mid$(Plane$, 3 * VertRec% + 1, VertRec%))
End Sub

Private Sub CheckIntersec(xa1 As Double, ya1 As Double, xa2 As Double, ya2 As Double, xb1 As Double, yb1 As Double, xb2 As Double, yb2 As Double, okint%)
    Dim ax1 As Double, ay1 As Double
    Dim ax2 As Double, ay2 As Double
    Dim bx1 As Double, by1 As Double
    Dim bx2 As Double, by2 As Double
    okint% = 0
    ax1 = xa1: ay1 = ya1: ax2 = xa2: ay2 = ya2
    bx1 = xb1: by1 = yb1: bx2 = xb2: by2 = yb2
    Call SwapSingle(ax1, ay1, ax2, ay2)
    Call SwapSingle(bx1, by1, bx2, by2)
    If ax1 > bx2 Or ax2 < bx1 Then Exit Sub
    If ay1 > by2 Or ay2 < by1 Then Exit Sub
    If bx1 > ax2 Or bx2 < ax1 Then Exit Sub
    If by1 > ay2 Or by2 < ay1 Then Exit Sub
    okint% = 1
End Sub

Private Function CheckOutMinMax(x As Double, y As Double, xMin As Double, yMin As Double, xMax As Double, yMax As Double) As Boolean
    If (x < xMin And FnZero(x - xMin) = 0) Or _
       (x > xMax And FnZero(x - xMax) = 0) Or _
       (y < yMin And FnZero(y - yMin) = 0) Or _
       (y > yMax And FnZero(y - yMax) = 0) Then
          CheckOutMinMax = True
    End If
End Function

Private Sub PointInVertex(x As Double, y As Double, VertexSel$, xMin As Double, yMin As Double, xMax As Double, yMax As Double, okin%)
    Dim recs%, h1$, counLf%, counRt%, Lf%, Rt%, ii%, okint%
    Dim xb1 As Double, yb1 As Double, xb2 As Double, yb2 As Double
    Dim xa1 As Double, ya1 As Double, xa2 As Double, ya2 As Double
    Dim xn As Double, yn As Double
    Dim dne%, coun%, now1%, Rtok%, Lfok%, i%, j%, first%
    Dim xMinTomi As Double, yMinTomi As Double
    Dim xMaxTomi As Double, yMaxTomi As Double
    okin% = 0: first% = 0
    If VertRec% = 0 Then Exit Sub
    recs% = Len(VertexSel$) / (2 * VertRec%)
    If recs% = 0 Then Exit Sub
    If xMin = xMax Or yMin = yMax Then Exit Sub
    If CheckOutMinMax(x, y, xMin, yMin, xMax, yMax) = True Then Exit Sub
    xb1 = xMin - 10
    yb1 = y
    xb2 = xMax + 10
    yb2 = y
    GoSub Getcoun
    h1$ = ""
    Exit Sub
Getcoun:
    counLf% = 0: counRt% = 0: Lf% = 1: Rt% = 2
    ReDim tomes(2, recs%, 2) As Double
    For ii% = 1 To recs%
        h1$ = Mid$(VertexSel$, (ii% - 1) * 2 * VertRec% + 1, 2 * VertRec%)
        Call GiveMeVline(VertexSel$, recs%, ii%, xa1, ya1, xa2, ya2)
        Call CheckIntersec(xa1, ya1, xa2, ya2, xb1, yb1, xb2, yb2, okint%)
        If okint% = 1 Then
           Call InterSectLines(xa1, ya1, xa2, ya2, xb1, yb1, xb2, yb2, xn, yn, dne%)
           If dne% = 2 Then
              If xn > x Then
                 coun% = counRt%
                 now1% = Rt%
                 GoSub CheckIfSame
                 counRt% = coun%
              Else
                 coun% = counLf%
                 now1% = Lf%
                 GoSub CheckIfSame
                 counLf% = coun%
              End If
           End If
        End If
    Next ii%
    coun% = counRt%: now1% = Rt%
    GoSub CheckAkra: Rtok% = okin%
    If okin% = 0 Then GoTo extsub
    coun% = counLf%: now1% = Lf%
    GoSub CheckAkra: Lfok% = okin%
    okin% = 0
    If Rtok% = 1 And Lfok% = 1 Then okin% = 1
    If okin% = 1 Then
       If first% > 1 Then
          If CheckOutMinMax(x, y, xMinTomi, yMinTomi, xMaxTomi, yMaxTomi) = True Then
             okin% = 0
          End If
       End If
    End If
extsub:
    Erase tomes
    Return
CheckIfSame:
    If coun% <> 0 Then
       If FnSamePoint(xn, yn, xa1, ya1) = 1 Or FnSamePoint(xn, yn, xa2, ya2) = 1 Then
          For j% = coun% To 1 Step -1
              If FnSamePoint(xn, yn, tomes(now1%, j%, 1), tomes(now1%, j%, 2)) = 1 Then
                 Return
              End If
          Next j%
       End If
    End If
    coun% = coun% + 1
    tomes(now1%, coun%, 1) = xn
    tomes(now1%, coun%, 2) = yn
    first% = first% + 1
    If first% = 1 Then
       xMinTomi = xn
       xMaxTomi = xn
       yMinTomi = yn
       yMaxTomi = yn
    Else
       Call FindMaxMinXY(xn, yn, xMaxTomi, xMinTomi, yMaxTomi, yMinTomi)
    End If
    Return
CheckAkra:
    okin% = 0
    If coun% <> 0 Then
       If coun% Mod 2 = 1 Then
          If tomes(now1%, 1, 2) = yMax Or tomes(now1%, 1, 2) = yMin Then
             If coun% = 1 Then
                GoTo extsub
             Else
                For i% = 2 To coun%
                    If tomes(now1%, i%, 2) <> yMax And tomes(now1%, i%, 2) <> yMin Then
                       GoTo Lab_2
                    End If
                Next i%
                Return
Lab_2:
             End If
          End If
          okin% = 1
       End If
    End If
    Return
End Sub

Private Function pÂetween(p As Double, p1 As Double, p2 As Double) As Boolean
    Dim mp1 As Double, mp2 As Double
    If p1 < p2 Then
       mp1 = p1
       mp2 = p2
    Else
       mp1 = p2
       mp2 = p1
    End If
    If FnZero(mp1 - p) = 1 Or FnZero(mp2 - p) = 1 Or (p >= mp1 And p <= mp2) Then
       pÂetween = True
    End If
End Function

Private Function PointInVertexArr(x As Double, y As Double, xV() As Double, yV() As Double, xMin As Double, yMin As Double, xMax As Double, yMax As Double) As Boolean
    Dim lRecs As Long, counLf%, counRt%, Lf%, Rt%, lCnt1 As Long, okint%
    Dim xb1 As Double, yb1 As Double, xb2 As Double, yb2 As Double
    Dim xa1 As Double, ya1 As Double, xa2 As Double, ya2 As Double
    Dim xn As Double, yn As Double, okin%
    Dim dne%, coun%, now1%, Rtok%, Lfok%, i%, j%, first%
    Dim xMinTomi As Double, yMinTomi As Double
    Dim xMaxTomi As Double, yMaxTomi As Double
    okin% = 0: first% = 0
    If VertRec% = 0 Then Exit Function
    lRecs = UboundSngX(xV())
    If lRecs <= 1 Then Exit Function
    If xV(1) = xV(lRecs) And yV(1) = yV(lRecs) Then lRecs = lRecs - 1
    If lRecs <= 1 Then Exit Function
    If xMin = xMax Or yMin = yMax Then Exit Function
    If CheckOutMinMax(x, y, xMin, yMin, xMax, yMax) = True Then Exit Function
    xb1 = xMin - 10
    yb1 = y
    xb2 = xMax + 10
    yb2 = y
    GoSub Getcoun
    If okin% = 1 Then PointInVertexArr = True
    Exit Function
Getcoun:
    counLf% = 0: counRt% = 0: Lf% = 1: Rt% = 2
    ReDim tomes(2, lRecs, 2) As Double
    For lCnt1 = 1 To lRecs
        xa1 = xV(lCnt1)
        ya1 = yV(lCnt1)
        If lCnt1 <> lRecs Then
           xa2 = xV(lCnt1 + 1)
           ya2 = yV(lCnt1 + 1)
        Else
           xa2 = xV(1)
           ya2 = yV(1)
        End If
        '-----
        If FnZero(xa1 - xa2) = 1 Then 'ÊÜèåôç
           If FnZero(xa1 - x) = 1 And FnZero(xa2 - x) = 1 And pÂetween(y, ya1, ya2) = True Then
              PointInVertexArr = True
              Exit Function
           End If
        End If
        If FnZero(ya1 - ya2) = 1 Then 'Ïñéæüíôéá
           If FnZero(ya1 - y) = 1 And FnZero(ya2 - y) = 1 And pÂetween(x, xa1, xa2) = True Then
              PointInVertexArr = True
              Exit Function
           End If
        End If
        If FnSamePoint(x, y, xa1, ya1) = 1 Or FnSamePoint(x, y, xa2, ya2) = 1 Then
           PointInVertexArr = True
           Exit Function
        End If
        '-----
        Call CheckIntersec(xa1, ya1, xa2, ya2, xb1, yb1, xb2, yb2, okint%)
        If okint% = 1 Then
           Call InterSectLines(xa1, ya1, xa2, ya2, xb1, yb1, xb2, yb2, xn, yn, dne%)
           If dne% = 2 Then
              If xn > x Then
                 coun% = counRt%
                 now1% = Rt%
                 GoSub CheckIfSame
                 counRt% = coun%
              Else
                 coun% = counLf%
                 now1% = Lf%
                 GoSub CheckIfSame
                 counLf% = coun%
              End If
           End If
        End If
    Next lCnt1
    coun% = counRt%: now1% = Rt%
    GoSub CheckAkra: Rtok% = okin%
    If okin% = 0 Then GoTo extsub
    coun% = counLf%: now1% = Lf%
    GoSub CheckAkra: Lfok% = okin%
    okin% = 0
    If Rtok% = 1 And Lfok% = 1 Then okin% = 1
    If okin% = 1 Then
       If first% > 1 Then
          If CheckOutMinMax(x, y, xMinTomi, yMinTomi, xMaxTomi, yMaxTomi) = True Then
             okin% = 0
          End If
       End If
    End If
extsub:
    Erase tomes
    Return
CheckIfSame:
    If coun% <> 0 Then
       If FnSamePoint(xn, yn, xa1, ya1) = 1 Or FnSamePoint(xn, yn, xa2, ya2) = 1 Then
          For j% = coun% To 1 Step -1
              If FnSamePoint(xn, yn, tomes(now1%, j%, 1), tomes(now1%, j%, 2)) = 1 Then
                 Return
              End If
          Next j%
       End If
    End If
    coun% = coun% + 1
    tomes(now1%, coun%, 1) = xn
    tomes(now1%, coun%, 2) = yn
    first% = first% + 1
    If first% = 1 Then
       xMinTomi = xn
       xMaxTomi = xn
       yMinTomi = yn
       yMaxTomi = yn
    Else
       Call FindMaxMinXY(xn, yn, xMaxTomi, xMinTomi, yMaxTomi, yMinTomi)
    End If
    Return
CheckAkra:
    okin% = 0
    If coun% <> 0 Then
       If coun% Mod 2 = 1 Then
          If tomes(now1%, 1, 2) = yMax Or tomes(now1%, 1, 2) = yMin Then
             If coun% = 1 Then
                GoTo extsub
             Else
                For i% = 2 To coun%
                    If tomes(now1%, i%, 2) <> yMax And tomes(now1%, i%, 2) <> yMin Then
                       GoTo Lab_2
                    End If
                Next i%
                Return
Lab_2:
             End If
          End If
          okin% = 1
       End If
    End If
    Return
End Function

Private Sub CheckIfPointInPolygon(x As Double, y As Double, VertexSel$, okin%)
    Dim Minx As Double, Miny As Double, Maxx As Double, Maxy As Double
    Call FindVertexMinMax(VertexSel$, Minx, Miny, Maxx, Maxy)
    Call PointInVertex(x, y, VertexSel$, Minx, Miny, Maxx, Maxy, okin%)
End Sub

Private Sub PsaCodes(mnx As Double, mny As Double, mxx As Double, mxy As Double, c%, x As Double, y As Double)
    c% = 0
    If x < mnx Then
       c% = c% Or 1
    Else
       If x > mxx Then c% = c% Or 4
    End If
    If y < mny Then
       c% = c% Or 2
    Else
       If y > mxy Then c% = c% Or 8
    End If
End Sub

Private Sub PsaMake(mnx As Double, mny As Double, mxx As Double, mxy As Double, c%, rx1 As Double, ry1 As Double, rx2 As Double, ry2 As Double)
    Dim zdx As Double, zdy As Double
    If c% = 0 Then Exit Sub
    zdx = rx2 - rx1: zdy = ry2 - ry1
    If (c% And 1) Then
       If zdx <> 0 Then ry1 = ry1 + zdy * (mnx - rx1) / zdx: rx1 = mnx
    ElseIf (c% And 2) Then
       If zdy <> 0 Then rx1 = rx1 + zdx * (mny - ry1) / zdy: ry1 = mny
    ElseIf (c% And 4) Then
       If zdx <> 0 Then ry1 = ry1 + zdy * (mxx - rx1) / zdx: rx1 = mxx
    ElseIf (c% And 8) Then
       If zdy <> 0 Then rx1 = rx1 + zdx * (mxy - ry1) / zdy: ry1 = mxy
    End If
End Sub

Private Sub PsalidisiLine(nowsct%, inWin%, mnx As Double, mny As Double, mxx As Double, mxy As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double, ok%)
    Dim c1%, c2%
    ok% = 0
    While ok% = 0
       Call PsaCodes(mnx, mny, mxx, mxy, c1%, x1, y1)
       Call PsaCodes(mnx, mny, mxx, mxy, c2%, x2, y2)
       If inWin% = 1 And (c1% <> 0 Or c2% <> 0) Then
          ok% = 2
       ElseIf (c1% = 0 Or c2% = 0) And nowsct% <> 0 Then
          ok% = 3
       ElseIf (c1% Or c2%) = 0 Then
          ok% = 1
       ElseIf (c1% And c2%) <> 0 Then
          ok% = 2
       Else
          Call PsaMake(mnx, mny, mxx, mxy, c1%, x1, y1, x2, y2)
          If c1% = 0 Then Call PsaMake(mnx, mny, mxx, mxy, c2%, x2, y2, x1, y1)
       End If
    Wend
End Sub

Function LineInBox(ByVal Minx As Double, ByVal Miny As Double, ByVal Maxx As Double, ByVal Maxy As Double, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, Optional finWin As Boolean = False) As Boolean
    Dim c1%, c2%
    Call SwapSingle(Minx, Miny, Maxx, Maxy)
    Call PsaCodes(Minx, Miny, Maxx, Maxy, c1%, x1, y1)
    Call PsaCodes(Minx, Miny, Maxx, Maxy, c2%, x2, y2)
    If finWin = True And (c1% <> 0 Or c2% <> 0) Then
       LineInBox = False
    ElseIf (c1% And c2%) <> 0 Then
       LineInBox = False
    Else
       LineInBox = True
    End If
End Function

Public Function ClipLine(ByVal Minx As Double, ByVal Miny As Double, ByVal Maxx As Double, ByVal Maxy As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double, Optional finWin As Boolean = False) As Boolean
    Dim nowsct%, ok%, inWin1%
    Call SwapSingle(Minx, Miny, Maxx, Maxy)
    nowsct% = 0: ok% = 0
    If finWin = True Then inWin1% = 1 Else inWin1% = 0
    Call PsalidisiLine(nowsct%, inWin1%, Minx, Miny, Maxx, Maxy, x1, y1, x2, y2, ok%)
    If ok% <> 2 Then ClipLine = True
End Function
 
Private Function UboundLngX(a() As Long) As Long
    On Local Error GoTo Lab_Error
    UboundLngX = UBound(a)
    Exit Function
Lab_Error:
    UboundLngX = 0
End Function

Private Function UboundSngX(a() As Double) As Long
    On Local Error GoTo Lab_Err
    UboundSngX = UBound(a)
    Exit Function
Lab_Err:
    UboundSngX = 0
End Function

Private Function UboundStrX(sArray() As String) As Integer
    On Local Error GoTo Lab_Err
    UboundStrX = UBound(sArray)
    Exit Function
Lab_Err:
    UboundStrX = 0
End Function

Private Sub FindVertexMinMax(VertexSel$, mnx As Double, mny As Double, mxx As Double, mxy As Double)
    Dim e As Double, i%, h$, recs%
    Dim x1 As Double, y1 As Double
    mnx = 0: mny = 0: mxx = 0: mxy = 0
    If VertRec% = 0 Then Exit Sub
    recs% = Len(VertexSel$) / (2 * VertRec%)
    For i% = 1 To recs%
        h$ = Mid$(VertexSel$, (i% - 1) * 2 * VertRec% + 1, 2 * VertRec%)
        x1 = Cvd(Mid$(h$, 1, VertRec%))
        y1 = Cvd(Mid$(h$, VertRec% + 1, VertRec%))
        If i% = 1 Then
           mxx = x1
           mnx = x1
           mxy = y1
           mny = y1
        End If
        Call FindMaxMinXY(x1, y1, mxx, mnx, mxy, mny)
    Next i%
    Call SwapSingle(mnx, mny, mxx, mxy)
End Sub

Private Function FindVertexMinMaxArr(xV() As Double, yV() As Double, mnx As Double, mny As Double, mxx As Double, mxy As Double) As Boolean
    Dim lCnt As Long, lMax As Long
    lMax = UboundSngX(xV())
    If lMax = 0 Then Exit Function
    mxx = xV(1)
    mnx = xV(1)
    mxy = yV(1)
    mny = yV(1)
    For lCnt = 1 To lMax
        Call FindMaxMinXY(xV(lCnt), yV(lCnt), mxx, mnx, mxy, mny)
    Next lCnt
    Call SwapSingle(mnx, mny, mxx, mxy)
    FindVertexMinMaxArr = True
End Function

Private Sub VetrexRadiate(VertexSel$, ok As Double)
    Dim x1 As Double, y1 As Double, Embado As Double, z As Double
    Dim NewVertex$, a$, recs%, i%
    ok = 0
    If VertRec% = 0 Then Exit Sub
    Call CentroidAndArea(VertexSel$, x1, y1, Embado, ok)
    If ok = 1 Then
       NewVertex$ = ""
       z = 0: a$ = Mkd$(x1) + Mkd$(y1) + Mkd$(z)
       recs% = Len(VertexSel$) / (2 * VertRec%)
       For i% = 1 To recs%
           NewVertex$ = NewVertex$ + a$ + Mid$(VertexSel$, (i% - 1) * 2 * VertRec% + 1, 2 * VertRec%) + Mkd$(z)
       Next i%
       VertexSel$ = NewVertex$: NewVertex$ = "": a$ = ""
    End If
End Sub

Private Sub GiveMeVertex(VertexSel$, recs%, i%, xb1 As Double, yb1 As Double)
    Dim h$
    If i% > recs% Or VertRec% = 0 Then Exit Sub
    h$ = Mid$(VertexSel$, (i% - 1) * 2 * VertRec% + 1, 2 * VertRec%)
    xb1 = Cvd(Mid$(h$, 1, VertRec%))
    yb1 = Cvd(Mid$(h$, VertRec% + 1, VertRec%))
    h$ = ""
End Sub

Private Sub GiveMeVline(VertexSel$, recs%, i%, xb1 As Double, yb1 As Double, xb2 As Double, yb2 As Double)
    Dim h$, ni%
    h$ = Mid$(VertexSel$, (i% - 1) * 2 * VertRec% + 1, 2 * VertRec%)
    xb1 = Cvd(Mid$(h$, 1, VertRec%))
    yb1 = Cvd(Mid$(h$, VertRec% + 1, VertRec%))
    If i% = recs% Then ni% = 1 Else ni% = i% + 1
    If recs% > 1 Then
       h$ = Mid$(VertexSel$, (ni% - 1) * 2 * VertRec% + 1, 2 * VertRec%)
       xb2 = Cvd(Mid$(h$, 1, VertRec%))
       yb2 = Cvd(Mid$(h$, VertRec% + 1, VertRec%))
    Else
       xb2 = xb1
       yb2 = yb1
    End If
    h$ = ""
End Sub

Private Sub MhkosVertex(VertexSel$, Mhkos As Double, ok As Double)
    Dim recs%, trec%, i%, d As Double
    Dim xb1 As Double, yb1 As Double, xb2 As Double, yb2 As Double
    ok = 0
    If VertRec% = 0 Then Exit Sub
    Mhkos = 0
    If VertexSel$ = "" Then Exit Sub
    recs% = Len(VertexSel$) / (2 * VertRec%)
    Mhkos = 0
    If recs% = 1 Then Exit Sub
    trec% = recs%
    If trec% <= 2 Then trec% = trec% - 1
    For i% = 1 To trec%
        Call GiveMeVline(VertexSel$, recs%, i%, xb1, yb1, xb2, yb2)
        Call Distance(xb1, yb1, xb2, yb2, d)
        Mhkos = Mhkos + d
    Next i%
    ok = 1
End Sub

Private Sub CheckCrossVertex(recs%, VertexSel$, CrossOk%)
    Dim i%, j%, dne%
    Dim xa1 As Double, ya1 As Double, xa2 As Double, ya2 As Double
    Dim xb1 As Double, yb1 As Double, xb2 As Double, yb2 As Double
    Dim xn As Double, yn As Double
    CrossOk% = 0
    For i% = 1 To recs% - 1
        Call GiveMeVline(VertexSel$, recs%, i%, xa1, ya1, xa2, ya2)
        For j% = i% + 1 To recs%
            Call GiveMeVline(VertexSel$, recs%, j%, xb1, yb1, xb2, yb2)
            dne% = 0
            Call InterSectLines(xa1, ya1, xa2, ya2, xb1, yb1, xb2, yb2, xn, yn, dne%)
            If dne% = 2 Then
               If ((FnZero(xn - xb1) = 1 And FnZero(yn - yb1) = 1) Or (FnZero(xn - xb2) = 1 And FnZero(yn - yb2) = 1)) _
                  And ((FnZero(xn - xa1) = 1 And FnZero(yn - ya1) = 1) Or (FnZero(xn - xa2) = 1 And FnZero(yn - ya2) = 1)) Then
                  GoTo MetaCross
               Else
                  CrossOk% = 1
                  Exit Sub
               End If
            End If
MetaCross:
        Next j%
    Next i%
End Sub

Private Function CheckCrossVertexArr(xV() As Double, yV() As Double) As Boolean
    Dim i As Long, j As Long, lRecs As Long, dne%
    Dim xa1 As Double, ya1 As Double, xa2 As Double, ya2 As Double
    Dim xb1 As Double, yb1 As Double, xb2 As Double, yb2 As Double
    Dim xn As Double, yn As Double
    lRecs = UboundSngX(xV)
    If lRecs < 3 Then Exit Function
    If UboundSngX(yV) <> lRecs Then CheckCrossVertexArr = True: Exit Function
    For i = 1 To lRecs - 2
        xa1 = xV(i)
        ya1 = yV(i)
        xa2 = xV(i + 1)
        ya2 = yV(i + 1)
        For j = (i + 1) To lRecs - 1
            xb1 = xV(j)
            yb1 = yV(j)
            xb2 = xV(j + 1)
            yb2 = yV(j + 1)
            If ClearInterSectLines(xa1, ya1, xa2, ya2, xb1, yb1, xb2, yb2) = True Then
               CheckCrossVertexArr = True
               Exit Function
            End If
        Next j
    Next i
End Function

Private Function ClearInterSectLines(xa1 As Double, ya1 As Double, xa2 As Double, ya2 As Double, xb1 As Double, yb1 As Double, xb2 As Double, yb2 As Double) As Boolean
    Dim xn As Double, yn   As Double, dne%
    dne% = 0
    Call InterSectLines(xa1, ya1, xa2, ya2, xb1, yb1, xb2, yb2, xn, yn, dne%)
    If dne% = 2 Then
       If ((FnZero(xn - xb1) = 1 And _
          FnZero(yn - yb1) = 1) Or _
          (FnZero(xn - xb2) = 1 And FnZero(yn - yb2) = 1)) _
          And ((FnZero(xn - xa1) = 1 And FnZero(yn - ya1) = 1) Or (FnZero(xn - xa2) = 1 And FnZero(yn - ya2) = 1)) Then
            Exit Function
       Else
            ClearInterSectLines = True
            Exit Function
       End If
    End If
End Function

Private Sub CentroidAndAreaArr(xV() As Double, yV() As Double, vxc As Double, vyc As Double, Embado As Double, ok As Double)
    Dim lRecs As Long, emb#, xtmp#, ytmp#, l As Long, xxa1#, yya1#
    Dim xa1 As Double, ya1 As Double, xa2 As Double, ya2 As Double
    Dim xxa2#, yya2#, ai#
    ok = 0
    Embado = 0
    lRecs = UboundSngX(xV())
    If lRecs < 3 Then Exit Sub
    emb# = 0#
    If CrossVertices(xV(), yV()) = True Then Exit Sub
    xtmp# = 0#: ytmp# = 0#
    For l = 1 To lRecs - 1
        xxa1# = xV(l)
        yya1# = yV(l)
        xxa2# = xV(l + 1)
        yya2# = yV(l + 1)
        ai# = xxa1# * yya2# - xxa2# * yya1#
        emb# = emb# + ai#
        xtmp# = xtmp# + (xxa1# + xxa2#) * ai#
        ytmp# = ytmp# + (yya1# + yya2#) * ai#
    Next l
    If emb# = 0# Then Exit Sub
    ok = 1
    xtmp# = xtmp# / (3# * emb#)
    ytmp# = ytmp# / (3# * emb#)
    vxc = xtmp#
    vyc = ytmp#
    emb# = emb# / 2#
    Embado = Abs(emb#)
End Sub

Private Sub CentroidAndArea(VertexSel$, vxc As Double, vyc As Double, Embado As Double, ok As Double)
    Dim recs%, emb#, CrossOk%, xtmp#, ytmp#, i%, xxa1#, yya1#
    Dim xa1 As Double, ya1 As Double, xa2 As Double, ya2 As Double
    Dim xxa2#, yya2#, ai#
    ok = 0
    If VertRec% = 0 Then Exit Sub
    Embado = 0
    If VertexSel$ = "" Then Exit Sub
    recs% = Len(VertexSel$) / (2 * VertRec%)
    If recs% < 3 Then Exit Sub
    emb# = 0#
    If recs% < 3 Then Exit Sub
    Call CheckCrossVertex(recs%, VertexSel$, CrossOk%)
    If CrossOk% = 1 Then Exit Sub
    xtmp# = 0#: ytmp# = 0#
    For i% = 1 To recs%
        Call GiveMeVline(VertexSel$, recs%, i%, xa1, ya1, xa2, ya2)
        xxa1# = xa1
        yya1# = ya1
        xxa2# = xa2
        yya2# = ya2
        ai# = xxa1# * yya2# - xxa2# * yya1#
        emb# = emb# + ai#
        xtmp# = xtmp# + (xxa1# + xxa2#) * ai#
        ytmp# = ytmp# + (yya1# + yya2#) * ai#
    Next i%
    If emb# = 0# Then Exit Sub
    ok = 1
    xtmp# = xtmp# / (3# * emb#)
    ytmp# = ytmp# / (3# * emb#)
    vxc = xtmp#
    vyc = ytmp#
    emb# = emb# / 2#
    Embado = Abs(emb#)
End Sub

Private Sub Rcvs(s As Double, aa$, nn As Double)
    If nn + VertRec% - 1 <= Len(aa$) Then
       s = Cvd(Mid$(aa$, nn, VertRec%))
    Else
       s = 0
    End If
    nn = nn + VertRec%
End Sub

Private Function FnMks2$(A1 As Double, A2 As Double)
    FnMks2$ = Mkd$(A1) & Mkd$(A2)
End Function

Private Function FnMks4$(A1 As Double, A2 As Double, a3 As Double, a4 As Double)
    FnMks4$ = Mkd$(A1) & Mkd$(A2) & Mkd$(a3) & Mkd$(a4)
End Function

Private Sub FastTransForm2xy(x As Double, y As Double, mat2() As Double)
    Dim x1 As Double
    Dim y1 As Double
    x1 = mat2(1, 1) * x
    x1 = x1 + mat2(2, 1) * y
    x1 = x1 + mat2(3, 1)
    y1 = mat2(1, 2) * x
    y1 = y1 + mat2(2, 2) * y
    y1 = y1 + mat2(3, 2)
    x = x1
    y = y1
End Sub

Public Function EllipsePerimeter(aR As Double, bR As Double) As Double
    Dim aR1#, bR1#, per#
    aR1# = aR
    bR1# = bR
    per# = pi# * (aR1# + bR1#)
    EllipsePerimeter = per#
End Function

Public Function EllipseArea(aR As Double, bR As Double) As Double
    Dim aR1#, bR1#, Embado#
    aR1# = aR
    bR1# = bR
    Embado# = pi# * aR1# * bR1#
    EllipseArea = Embado#
End Function

Public Function RegularPolygonSideLength(R As Double, nSides As Long) As Double
    Dim a#, Rd#
    If nSides < 3 Then Exit Function
    Rd# = R
    a# = (2# * pi#) / CDbl(nSides)
    RegularPolygonSideLength = 2# * Rd# * Sin(a# / 2#)
End Function

Public Function RegularPolygonPerimeter(R As Double, nSides As Long) As Double
    Dim a#, Rd#
    If nSides < 3 Then Exit Function
    Rd# = R
    a# = (2# * pi#) / CDbl(nSides)
    RegularPolygonPerimeter = 2# * CDbl(nSides) * Rd# * Sin(a# / 2#)
End Function

Public Function RegularPolygonArea(R As Double, nSides As Long) As Double
    Dim a#, Rd#
    If nSides < 3 Then Exit Function
    Rd# = R
    a# = (2# * pi#) / CDbl(nSides)
    RegularPolygonArea = CDbl(nSides) * ((Rd# ^ 2#) / 2#) * Sin(a#)
End Function

Public Function CirclePerimeter(R As Double) As Double
    Dim R1#, per#
    R1# = R
    per# = 2 * pi# * R1#
    CirclePerimeter = per#
End Function

Public Function CircleArea(R As Double) As Double
    Dim R1#, Embado#
    R1# = R
    Embado# = pi# * (R1# ^ 2)
    CircleArea = Embado#
End Function

Public Function TransformVertices(mat2d() As Double, xV() As Double, yV() As Double) As Boolean
    Dim nber As Long, l As Long
    On Local Error GoTo Lab_Error
    nber = UboundSngX(xV())
    If UboundSngX(yV()) <> nber Then Exit Function
    For l = 1 To nber
        Call Transform2xy(xV(l), yV(l), mat2d())
    Next l
    TransformVertices = True
    Exit Function
Lab_Error:
    TransformVertices = False
End Function

Private Sub Dejikena(ss$)
    ss$ = RTrim$(ss$)
End Sub

Private Sub CheckClosedVertex(VertexSel$)
    Dim aa$, bb$, nn As Double
    Dim xpr1 As Double, ypr1 As Double
    Dim xpr2 As Double, ypr2 As Double
    If VertexSel$ = "" Then Exit Sub
    aa$ = Left$(VertexSel$, 2 * VertRec%)
    bb$ = Right$(VertexSel$, 2 * VertRec%)
    If aa$ = bb$ Then
       VertexSel$ = Mid$(VertexSel$, 1, Len(VertexSel$) - 2 * VertRec%)
       aa$ = ""
       bb$ = ""
       Exit Sub
    End If
    nn = 1
    Call Rcvs(xpr1, aa$, nn)
    Call Rcvs(ypr1, aa$, nn)
    aa$ = ""
    nn = 1
    Call Rcvs(xpr2, bb$, nn)
    Call Rcvs(ypr2, bb$, nn)
    bb$ = ""
    If FnSamePoint(xpr1, ypr1, xpr2, ypr2) = 1 Then
       VertexSel$ = Mid$(VertexSel$, 1, Len(VertexSel$) - 2 * VertRec%)
    End If
End Sub

Private Sub MakeVertexStr(xV() As Double, yV() As Double, VertexSel$)
    On Local Error GoTo Lab_Error1
    Dim l As Long
    Dim nber As Long
    VertexSel$ = ""
    nber = UboundSngX(xV())
    For l = 1 To nber
        VertexSel$ = VertexSel$ & Mkd$(xV(l)) & Mkd$(yV(l))
    Next l
    Exit Sub
Lab_Error1:
    VertexSel$ = ""
End Sub

Public Function PointInVertices(x As Double, y As Double, xV() As Double, yV() As Double) As Boolean
    Dim Minx As Double, Miny As Double, Maxx As Double, Maxy As Double
    Call FindVertexMinMaxArr(xV(), yV(), Minx, Miny, Maxx, Maxy)
    PointInVertices = PointInVertexArr(x, y, xV(), yV(), Minx, Miny, Maxx, Maxy)
End Function

Public Function VerticesArea(xV() As Double, yV() As Double, Area As Double) As Boolean
    Dim nber As Long, ok As Double
    Dim xC As Double, yC As Double
    Dim VertexSel$, akr1 As Double
    On Local Error GoTo Lab_Error
    nber = UboundSngX(xV())
    Area = 0
    If UboundSngX(yV()) <> nber Then Exit Function
    ok = 0: xC = 0: yC = 0
    Call CentroidAndAreaArr(xV(), yV(), xC, yC, Area, ok)
    akr1 = 5
    Call strogilS(Area, akr1)
    If ok = 1 Then VerticesArea = True
    Exit Function
Lab_Error:
    VerticesArea = False
End Function

Public Function VerticesPerimeter(xV() As Double, yV() As Double, Perimeter As Double) As Boolean
    Dim nber As Long, l As Long, d As Double, per As Double
    On Local Error GoTo Lab_Error
    Perimeter = 0
    nber = UboundSngX(xV())
    If UboundSngX(yV()) <> nber Then Exit Function
    For l = 1 To nber - 1
        d = 0
        If l <> nber - 1 Then
           d = DistanceX(xV(l), yV(l), xV(l + 1), yV(l + 1))
        Else
           If FnSamePoint(xV(l), yV(l), xV(1), yV(1)) = 0 Then
              d = DistanceX(xV(l), yV(l), xV(1), yV(1))
           End If
        End If
        per = per + d
    Next l
    Perimeter = per
    VerticesPerimeter = True
    Exit Function
Lab_Error:
    VerticesPerimeter = False
End Function

Public Function VerticesBorder(xV() As Double, yV() As Double, Minx As Double, Miny As Double, Maxx As Double, Maxy As Double) As Boolean
    Dim nber As Long
    nber = UboundSngX(xV)
    If UboundSngX(yV) <> nber Then Exit Function
    Call FindVertexMinMaxArr(xV(), yV(), Minx, Miny, Maxx, Maxy)
End Function

Public Function VerticesCenter(xV() As Double, yV() As Double, xC As Double, yC As Double) As Boolean
    Dim nber As Long, Embado As Double, ok As Double
    Dim akr1 As Double
    Dim VertexSel$
    On Local Error GoTo Lab_Error
    nber = UboundSngX(xV())
    If UboundSngX(yV()) <> nber Then Exit Function
    ok = 0: xC = 0: yC = 0
    Call CentroidAndAreaArr(xV(), yV(), xC, yC, Embado, ok)
    akr1 = 5
    Call strogilS(xC, akr1)
    Call strogilS(yC, akr1)
    If ok = 1 Then VerticesCenter = True
    Exit Function
Lab_Error:
    VerticesCenter = False
End Function

Public Function CrossVertices(xV() As Double, yV() As Double) As Boolean
    CrossVertices = CheckCrossVertexArr(xV(), yV())
End Function

Public Function GetEllipseVertices(fx As Double, fy As Double, aR As Double, bR As Double, nVert As Long, xV() As Double, yV() As Double, zV() As Double) As Boolean
    Dim dtheta As Double, cdtheta As Double
    Dim sdtheta As Double
    Dim c1 As Double, c2 As Double
    Dim x0 As Double, y0 As Double
    Dim x1 As Double, y1 As Double
    Dim xx1 As Double, yy1 As Double
    Dim xx2 As Double, yy2 As Double
    Dim inber As Long, l As Long, z As Double
    Dim xn2 As Double, yn2 As Double, zn2 As Double
    Dim nber As Long
    On Local Error GoTo Lab_Error
    If nVert < 4 Or aR = 0 Or bR = 0 Then Exit Function
    nVert = nVert + 1
    ReDim xV(nVert) As Double
    ReDim yV(nVert) As Double
    ReDim zV(nVert) As Double
    nber = nVert - 1
    dtheta = 360 / (nber - 1)
    cdtheta = FnCos(dtheta)
    sdtheta = FnSin(dtheta)
    c1 = (aR / bR) * sdtheta
    c2 = (bR / aR) * sdtheta
    x0 = aR: y0 = 0
    inber = nber + 1
    For l = 2 To inber
        x1 = x0 * cdtheta - c1 * y0
        y1 = c2 * x0 + y0 * cdtheta
        z = 0
        xx1 = x0 + fx
        yy1 = y0 + fy
        xx2 = x1 + fx
        yy2 = y1 + fy
        If l > 1 Then
           xV(l - 1) = xx1
           yV(l - 1) = yy1
        End If
        x0 = x1
        y0 = y1
    Next l
    xV(nVert) = xV(1)
    yV(nVert) = yV(1)
    GetEllipseVertices = True
    Exit Function
Lab_Error:
    GetEllipseVertices = False
End Function

Public Function GetCircleVertices(fx As Double, fy As Double, R As Double, nVert As Long, xV() As Double, yV() As Double, zV() As Double) As Boolean
    Dim aR As Double, bR As Double
    If R = 0 Then Exit Function
    aR = R
    bR = R
    GetCircleVertices = GetEllipseVertices(fx, fy, aR, bR, nVert, xV(), yV(), zV())
End Function

Public Function GetBoxVertices(x0 As Double, y0 As Double, Width As Double, Height As Double, xV() As Double, yV() As Double, zV() As Double) As Boolean
    If Width = 0 Or Height = 0 Then Exit Function
    ReDim xV(5) As Double, yV(5) As Double, zV(5) As Double
    xV(1) = x0
    yV(1) = y0
    xV(2) = x0
    yV(2) = y0 + Height
    xV(3) = x0 + Width
    yV(3) = y0 + Height
    xV(4) = x0 + Width
    yV(4) = y0
    xV(5) = x0
    yV(5) = y0
    GetBoxVertices = True
End Function

Public Function GetRegularPolygonVertices(fx As Double, fy As Double, R As Double, nVert As Long, xV() As Double, yV() As Double, zV() As Double) As Boolean
    Dim inberl As Long, l As Long
    Dim ctheta As Double, stheta As Double, th As Double
    Dim x0 As Double, y0 As Double
    Dim x1 As Double, y1 As Double
    Dim xx1 As Double, yy1 As Double
    Dim xx2 As Double, yy2 As Double
    If nVert < 3 Or R = 0 Then Exit Function
    ReDim xV(nVert + 1) As Double
    ReDim yV(nVert + 1) As Double
    ReDim zV(nVert + 1) As Double
    th = 360# / CDbl(nVert)
    ctheta = FnCosD(th)
    stheta = FnSinD(th)
    x0 = R
    y0 = 0
    inberl = nVert + 1
    For l = 2 To inberl
        x1 = ctheta * x0 - stheta * y0
        y1 = stheta * x0 + ctheta * y0
        xx1 = x0 + CDbl(fx)
        yy1 = y0 + CDbl(fy)
        xx2 = x1 + CDbl(fx)
        yy2 = y1 + CDbl(fy)
        xV(l - 1) = xx1
        yV(l - 1) = yy1
        x0 = x1
        y0 = y1
    Next l
    xV(nVert + 1) = xV(1)
    yV(nVert + 1) = yV(1)
    GetRegularPolygonVertices = True
    Exit Function
Lab_Error:
    GetRegularPolygonVertices = False
End Function

Public Function GetArcVertices(x1 As Double, y1 As Double, x2 As Double, y2 As Double, xc1 As Double, yc1 As Double, nVert As Long, xV() As Double, yV() As Double, zV() As Double, Optional fClosed As Boolean = True) As Boolean
    On Local Error GoTo Lab_Error
    Dim z1 As Double
    Dim z2 As Double
    Dim zc1 As Double
    Dim xr1 As Double
    Dim yr1 As Double
    Dim zr1 As Double
    Dim R As Double
    Dim aR1 As Double
    Dim bR1 As Double
    Dim theta1 As Double
    Dim theta2 As Double
    Dim theta3 As Double
    If nVert < 3 Then Exit Function
    Call Distance(x1, y1, xc1, yc1, R)
    If R = 0 Then Exit Function
    Call FindRealAngle(x1, y1, xc1, yc1, theta1)
    Call FindRealAngle(x2, y2, xc1, yc1, theta2)
    If theta1 = theta2 Then
       GetArcVertices = GetCircleVertices(xc1, yc1, R, nVert, xV(), yV(), zV())
       Exit Function
    End If
    If fClosed = True Then nVert = nVert + 1
    ReDim xV(nVert) As Double
    ReDim yV(nVert) As Double
    ReDim zV(nVert) As Double
    If theta1 < theta2 Then
       theta3 = theta1 + (theta2 - theta1) / 2
    Else
       theta3 = theta2 + (theta1 - theta2) / 2 + 180
    End If
    xr1 = xc1 + FnCos(theta3) * R
    yr1 = yc1 + FnSin(theta3) * R
    zr1 = 0
    GetArcVertices = GetArc(x1, y1, z1, x2, y2, z2, xc1, yc1, zc1, xr1, yr1, zr1, nVert, xV(), yV(), zV(), fClosed)
    Exit Function
Lab_Error:
    GetArcVertices = False
End Function

Private Function GetArc(x1 As Double, y1 As Double, z1 As Double, x2 As Double, y2 As Double, z2 As Double, xc1 As Double, yc1 As Double, zc1 As Double, xr1 As Double, yr1 As Double, zr1 As Double, nVert As Long, xV() As Double, yV() As Double, zV() As Double, fClosed As Boolean) As Boolean
    Dim ArcPoints As Double, R As Double
    GetArc = False
    ArcPoints = UboundSngX(xV())
    If ArcPoints < 3 Then Exit Function
    Call Distance(x1, y1, xc1, yc1, R)
    GetArc = Arc3Points(xc1, yc1, zc1, R, x1, y1, x2, y2, xr1, yr1, ArcPoints, xV(), yV(), zV(), fClosed)
End Function

Private Function Arc3Points(xC As Double, yC As Double, zC As Double, R As Double, Startx As Double, Starty As Double, Endx As Double, Endy As Double, newx As Double, newy As Double, ArcPoints As Double, xV() As Double, yV() As Double, zV() As Double, fClosed As Boolean) As Boolean
    Dim oldStartx As Double, oldStarty As Double, swp%
    Dim thetaS As Double, thetaE As Double, thetaN As Double
    Dim thetaStp As Double, th As Double, theta As Double
    Dim zc0 As Double, zc1 As Double, cth%, iCnt As Integer
    Dim x1 As Double, y1 As Double, ArcPoints1 As Double
    Dim x0 As Double, y0 As Double, CirclePoints As Double
    oldStartx = Startx
    oldStarty = Starty
    ArcPoints1 = ArcPoints
    ReDim xV(ArcPoints1), yV(ArcPoints1), zV(ArcPoints1)
    If ArcPoints1 <= 1 Then Exit Function
    Call FindRealAngle(Startx, Starty, xC, yC, thetaS)
    Call FindRealAngle(Endx, Endy, xC, yC, thetaE)
    Call FindRealAngle(newx, newy, xC, yC, thetaN)
    If thetaS > thetaE Then swp% = 1 Else swp% = 0
    If thetaS = thetaE Then Exit Function
    If swp% = 1 Then
       Swap Startx, Endx
       Swap Starty, Endy
       Swap thetaS, thetaE
    End If
    If fClosed = True Then
       CirclePoints = (360 / (thetaE - thetaS)) * (ArcPoints1 - 1)
    Else
       CirclePoints = (360 / (thetaE - thetaS)) * ArcPoints1
    End If
    thetaStp = 360 / CirclePoints
    If thetaN > thetaE Or thetaN < thetaS Then
       Swap Startx, Endx
       Swap Starty, Endy
       Swap thetaS, thetaE
       thetaE = 360 + thetaE
       If swp% = 1 Then swp% = 0 Else swp% = 1
    End If
    zc0 = zC
    zc1 = zC
    iCnt = 0
    cth% = 0
    For th = thetaS To thetaE Step thetaStp
        cth% = cth% + 1
        theta = th
        If theta > 360 Then theta = theta - 360
        x1 = xC + R * FnCos(theta)
        y1 = yC + R * FnSin(theta)
        If cth% = 2 Then x0 = Startx: y0 = Starty
        If cth% > 1 Then
           If x0 <> x1 Or y0 <> y1 Or iCnt = 0 Then
              GoSub Inc_iCnt
              xV(iCnt) = x0
              yV(iCnt) = y0
           End If
        End If
        x0 = x1
        y0 = y1
    Next th
    If ArcPoints1 >= 1 Then
       xV(1) = Startx
       yV(1) = Starty
    End If
    If xV(iCnt) <> Endx Or yV(iCnt) = Endy Then
       GoSub Inc_iCnt
       xV(iCnt) = Endx
       yV(iCnt) = Endy
    End If
    If fClosed = True Then
       GoSub Inc_iCnt
       xV(iCnt) = Startx
       yV(iCnt) = Starty
    End If
    ReDim Preserve xV(iCnt), yV(iCnt), zV(iCnt)
    If swp% = 1 Then
       Swap Startx, Endx
       Swap Starty, Endy
    End If
    If iCnt <> 0 Then Arc3Points = True
    Exit Function
Inc_iCnt:
    iCnt = iCnt + 1
    If iCnt > ArcPoints1 Then
       ReDim Preserve xV(iCnt), yV(iCnt), zV(iCnt)
       ArcPoints1 = iCnt
    End If
    Return
End Function

Private Sub FindA1(a As Double, b As Double, c As Double, th1 As Double, problem%)
    Dim p As Double, dd As Double, k As Double, tnth1 As Double
    If a * b * c = 0 Then problem% = 1: Exit Sub
    p = (a + b + c) / 2
    dd = p * (p - a)
    If dd = 0 Then problem% = 1: Exit Sub
    k = (p - b) * (p - c) / dd
    If k < 0 Then problem% = 1: Exit Sub
    tnth1 = Sqr(k)
    th1 = FnArcTan(tnth1) * 2
End Sub

Private Sub SolveAnyTriangle(Trig As Double, D1 As Double, D2 As Double, d3 As Double, xV() As Double, yV() As Double, zV() As Double, problem%)
    ReDim xV(4) As Double, yV(4) As Double, zV(4) As Double
    Dim th1 As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim newx As Double, newy As Double, k As Double, a As Double, c1 As Double
    Dim B1 As Double, A1 As Double, b As Double, c As Double
    problem% = 0
    If Trig = 1 Then
       'ABC
       If D1 < 0 Or D2 < 0 Or d3 < 0 Then problem% = 1: Exit Sub
Trig1: Call FindA1(D2, d3, D1, th1, problem%)
       If problem% = 1 Then Exit Sub
       x1 = 0: y1 = 0: x2 = D1: y2 = 0
       newx = x1 + FnCos(th1) * d3
       newy = y1 + FnSin(th1) * d3
       xV(1) = x1
       yV(1) = y1
       xV(2) = x2
       yV(2) = y2
       xV(3) = newx
       yV(3) = newy
       xV(4) = x1
       yV(4) = y1
    ElseIf Trig = 2 Then
       'AcB
       k = D1 ^ 2 + d3 ^ 2 - 2 * D1 * d3 * FnCos(D2)
       If k < 0 Then problem% = 1: Exit Sub
       D2 = Sqr(k): Swap D2, d3: GoTo Trig1
    ElseIf Trig = 3 Then
       'bAc
       a = D2: c1 = d3: B1 = D1
       A1 = 180 - B1 - c1
       If FnSin(B1) = 0 Or FnSin(A1) = 0 Then problem% = 1: Exit Sub
       b = a / (FnSin(A1) / FnSin(B1))
       c = b * FnSin(c1) / FnSin(B1)
       D1 = a: D2 = b: d3 = c: GoTo Trig1
    End If
End Sub

Public Function GetTriangleVerticesABC(A1 As Double, B1 As Double, c1 As Double, xV() As Double, yV() As Double, zV() As Double, Optional fClosed As Boolean = True) As Boolean
    On Local Error GoTo Lab_Error
    Dim Trig As Double, problem%
    Trig = 1: problem% = 0
    Call SolveAnyTriangle(Trig, A1, B1, c1, xV(), yV(), zV(), problem%)
    If problem% = 0 Then GetTriangleVerticesABC = True
    Exit Function
Lab_Error:
    GetTriangleVerticesABC = False
End Function

Public Function GetTriangleVerticesAcB(A1 As Double, c As Double, B1 As Double, xV() As Double, yV() As Double, zV() As Double, Optional fClosed As Boolean = True) As Boolean
    On Local Error GoTo Lab_Error
    Dim Trig As Double, problem%
    Trig = 2: problem% = 0
    Call SolveAnyTriangle(Trig, A1, c, B1, xV(), yV(), zV(), problem%)
    If problem% = 0 Then GetTriangleVerticesAcB = True
    Exit Function
Lab_Error:
    GetTriangleVerticesAcB = False
End Function

Public Function GetTriangleVerticesbAc(b As Double, A1 As Double, c As Double, xV() As Double, yV() As Double, zV() As Double, Optional fClosed As Boolean = True) As Boolean
    On Local Error GoTo Lab_Error
    Dim Trig As Double, problem%
    Trig = 3: problem% = 0
    Call SolveAnyTriangle(Trig, b, A1, c, xV(), yV(), zV(), problem%)
    If problem% = 0 Then GetTriangleVerticesbAc = True
    Exit Function
Lab_Error:
    GetTriangleVerticesbAc = False
End Function

Public Function IsSamePoint(xa1 As Double, ya1 As Double, xb1 As Double, yb1 As Double) As Boolean
    If FnSamePoint(xa1, ya1, xb1, yb1) = 1 Then IsSamePoint = True
End Function

Private Sub DoHatch(brdpoi%, OsX%, Hin%, Hbrd%, sx1 As Double, sy1 As Double, sx2 As Double, sy2 As Double, pmnx As Double, pmny As Double, pmxx As Double, pmxy As Double, VertexSel$, VertRec%, recs%, lHatchCnt As Long, xV1() As Double, yV1() As Double, xV2() As Double, yV2() As Double, ok As Double)
    Dim z0 As Double, xx1 As Double, yy1 As Double, xx2 As Double, yy2 As Double
    Dim nwst%, inwnl%, vok1%, i%
    Dim hx1 As Double, hy1 As Double, hx2 As Double, hy2 As Double
    Dim xh1 As Double, yh1 As Double, xh2 As Double, yh2 As Double
    Dim cnx As Double, zh1 As Double, zh2 As Double
    Dim xb1 As Double, yb1 As Double, xb2 As Double, yb2 As Double
    Dim xn As Double, yn As Double, dne%, apk As Double, k As Double
    Dim j As Double, ubXY As Long
    ubXY = UboundSngX(xV1())
    z0 = 0
    Call Equal4Double(xx1, yy1, xx2, yy2, sx1, sy1, sx2, sy2): nwst% = 0: inwnl% = 0
    Call PsalidisiLine(nwst%, inwnl%, pmnx, pmny, pmxx, pmxy, xx1, yy1, xx2, yy2, vok1%)
    If vok1% = 2 Then Exit Sub
    If (Hbrd% = 1 Or Hbrd% = 3) And brdpoi% = 0 Then
       Call Equal4Double(hx1, hy1, hx2, hy2, pmnx, pmny, pmxx, pmxy)
       Call Equal4Double(xh1, yh1, xh2, yh2, hx1, hy1, hx1, hy2): GoSub PutHatch
       Call Equal4Double(xh1, yh1, xh2, yh2, hx1, hy2, hx2, hy2): GoSub PutHatch
       Call Equal4Double(xh1, yh1, xh2, yh2, hx2, hy2, hx2, hy1): GoSub PutHatch
       Call Equal4Double(xh1, yh1, xh2, yh2, hx2, hy1, hx1, hy1): GoSub PutHatch
    End If
    cnx = 0: zh1 = 0: zh2 = 0
    ReDim intXsec(recs% + 1) As Double, intYsec(recs% + 1) As Double
    For i% = 1 To recs%
        Call GiveMeVline(VertexSel$, recs%, i%, xb1, yb1, xb2, yb2)
        Call InterSectLines(xx1, yy1, xx2, yy2, xb1, yb1, xb2, yb2, xn, yn, dne%)
        If Hbrd% > 1 And brdpoi% = 0 Then Call Equal4Double(xh1, yh1, xh2, yh2, xb1, yb1, xb2, yb2): GoSub PutHatch
        If dne% = 2 Then cnx = cnx + 1: intXsec(cnx) = xn: intYsec(cnx) = yn
    Next i%
    brdpoi% = brdpoi% + 1: If Hin% = 1 Then apk = 1 Else apk = 0
janahatch:
    If cnx > 1 Then
       For k = 1 To cnx - 1
           For j = k + 1 To cnx
               If (OsX% = 1 And intXsec(k) > intXsec(j)) Or (OsX% = 0 And intYsec(k) > intYsec(j)) Then
                  Swap intXsec(k), intXsec(j): Swap intYsec(k), intYsec(j)
               End If
           Next j
       Next k
       For k = 1 To cnx - 1
           If FnSamePoint(intXsec(k), intYsec(k), intXsec(k + 1), intYsec(k + 1)) = 1 Then
              For j = k + 1 To cnx - 1: Swap intXsec(j), intXsec(j + 1): Swap intYsec(j), intYsec(j + 1): Next j
              cnx = cnx - 1
           End If
       Next k
       If cnx <= 1 Then GoTo janahatch
       For k = apk To cnx Step 2
           If k = 0 Then
              xh1 = xx1: yh1 = yy1: xh2 = intXsec(k + 1): yh2 = intYsec(k + 1)
           ElseIf k = cnx Then
              xh1 = intXsec(k): yh1 = intYsec(k): xh2 = xx2: yh2 = yy2
           Else
              xh1 = intXsec(k): yh1 = intYsec(k): xh2 = intXsec(k + 1): yh2 = intYsec(k + 1)
           End If
           GoSub PutHatch
       Next k
    ElseIf Hin% = 0 Then
       xh1 = xx1
       yh1 = yy1
       xh2 = xx2
       yh2 = yy2
       GoSub PutHatch
    End If
    Erase intXsec, intYsec
    Exit Sub
Rem ................................................
PutHatch:
    If xh1 = xh2 And yh1 = yh2 Then Return
    lHatchCnt = lHatchCnt + 1
    If lHatchCnt > ubXY Then
       ReDim Preserve xV1(lHatchCnt) As Double
       ReDim Preserve yV1(lHatchCnt) As Double
       ReDim Preserve xV2(lHatchCnt) As Double
       ReDim Preserve yV2(lHatchCnt) As Double
       ubXY = lHatchCnt
    End If
    xV1(lHatchCnt) = xh1
    yV1(lHatchCnt) = yh1
    xV2(lHatchCnt) = xh2
    yV2(lHatchCnt) = yh2
    Return
End Sub

Private Function MakeHatch(VertexSel$, HatchDistance As Double, theta As Double, Hin%, Hbrd%, HatchMask$, xV1() As Double, yV1() As Double, zV1() As Double, xV2() As Double, yV2() As Double, zV2() As Double, ok As Double) As Long
    Dim recs%, m%, LeMask%, MaskCnt%, brdpoi%, OsX%, OkMask%
    Dim mnx As Double, mny As Double, mxx As Double, mxy As Double
    Dim pmnx As Double, pmny As Double, pmxx As Double, pmxy As Double
    Dim xx1 As Double, yy1 As Double, xx2 As Double, yy2 As Double
    Dim HTheta As Double, stpx As Double, stpy As Double, Jy As Double
    Dim lHatchCnt As Long, StartDim As Double
    ReDim xV1(0) As Double, yV1(0) As Double, zV1(0) As Double
    ReDim xV2(0) As Double, yV2(0) As Double, zV2(0) As Double
    lHatchCnt = 0
    ok = 0
    If VertexSel$ = "" Then GoTo Function_Exit
    If VertRec% = 0 Then GoTo Function_Exit
    recs% = Len(VertexSel$) / (2 * VertRec%)
    If recs% < 3 Then GoTo Function_Exit
    If Left$(VertexSel$, 2 * VertRec%) = Right$(VertexSel$, 2 * VertRec%) Then recs% = recs% - 1
    If recs% < 3 Then GoTo Function_Exit
    Call Dejikena(HatchMask$)
    If HatchMask$ <> "" Then
       For m% = 1 To Len(HatchMask$)
           If Mid$(HatchMask$, m%, 1) <> "1" And Mid$(HatchMask$, m%, 1) <> "0" Then
              HatchMask$ = ""
              GoTo NoMask
           End If
       Next m%
    End If
NoMask:
    LeMask% = Len(HatchMask$)
    MaskCnt% = 0
    Call FindVertexMinMax(VertexSel$, mnx, mny, mxx, mxy)
    Call Equal4Double(pmnx, pmny, pmxx, pmxy, mnx, mny, mxx, mxy)
    Call Equal4Double(xx1, yy1, xx2, yy2, mnx, mny, mxx, mxy)
    HTheta = theta
    brdpoi% = 0
    If HTheta >= 180 Then HTheta = HTheta - 180
    If HTheta > 90 Then HTheta = HTheta - 180
    If HTheta = 90 Or FnZero(HTheta - 90) = 1 Then
       stpx = HatchDistance
       StartDim = Int((mxx - mnx) / stpx)
       If StartDim > 1200 Then GoTo Function_Exit
       ReDim xV(StartDim) As Double, yV(StartDim) As Double
       For xx1 = mnx To mxx Step stpx
           xx2 = xx1
           OsX% = 0
           GoSub ChecKOkMask
           If OkMask% = 1 Then
              Call DoHatch(brdpoi%, OsX%, Hin%, Hbrd%, xx1, yy1, xx2, yy2, pmnx, pmny, pmxx, pmxy, VertexSel$, VertRec%, recs%, lHatchCnt, xV1(), yV1(), xV2(), yV2(), ok)
           End If
       Next xx1
    Else
       If FnCos(HTheta) <> 0 Then
          Jy = (mxx - mnx) * FnSin(HTheta) / FnCos(HTheta)
          Jy = Abs(Jy)
          stpy = HatchDistance / FnCos(HTheta)
          stpy = Abs(stpy)
          mny = mny - Jy
          StartDim = Int((mxy - mny) / stpy)
          If StartDim > 1200 Then GoTo Function_Exit
          ReDim xV(StartDim) As Double, yV(StartDim) As Double
          For yy1 = mny To mxy Step stpy
              yy2 = yy1 + Jy
              OsX% = 1
              If HTheta < 0 Then Swap yy1, yy2
              GoSub ChecKOkMask
              If OkMask% = 1 Then
                 Call DoHatch(brdpoi%, OsX%, Hin%, Hbrd%, xx1, yy1, xx2, yy2, pmnx, pmny, pmxx, pmxy, VertexSel$, VertRec%, recs%, lHatchCnt, xV1(), yV1(), xV2(), yV2(), ok)
              End If
              If HTheta < 0 Then Swap yy1, yy2
          Next yy1
       End If
    End If
Function_Exit:
    MakeHatch = lHatchCnt
    ReDim Preserve xV1(lHatchCnt) As Double
    ReDim Preserve yV1(lHatchCnt) As Double
    ReDim zV1(lHatchCnt) As Double
    ReDim Preserve xV2(lHatchCnt) As Double
    ReDim Preserve yV2(lHatchCnt) As Double
    ReDim zV2(lHatchCnt) As Double
    Exit Function
Rem ................................................
ChecKOkMask:
    OkMask% = 0: If LeMask% = 0 Then OkMask% = 1: Return
    MaskCnt% = MaskCnt% + 1
    If MaskCnt% > LeMask% Then MaskCnt% = 1
    If Mid$(HatchMask$, MaskCnt%, 1) = "1" Then OkMask% = 1
    Return
End Function

Public Function GetHatchBlockLines(xV() As Double, yV() As Double, HatchDistance As Double, Angle As Double, Hin%, Hbrd%, HatchMask$, x1() As Double, y1() As Double, z1() As Double, x2() As Double, y2() As Double, z2() As Double) As Long
    Dim VertexSel$, ok As Double, lMaxLines As Long
    Call MakeVertexStr(xV(), yV(), VertexSel$)
    If VertexSel$ <> "" Then
       lMaxLines = MakeHatch(VertexSel$, HatchDistance, Angle, Hin%, Hbrd%, HatchMask$, x1(), y1(), z1(), x2(), y2(), z2(), ok)
       If ok <> 0 Then GetHatchBlockLines = 0 Else GetHatchBlockLines = lMaxLines
    End If
End Function

Private Function mComparePolygons(xMainV() As Double, yMainV() As Double, xV1() As Double, yV1() As Double, iAns As PolygonsComp) As Boolean
    On Local Error GoTo Lab_Error1
    Dim i As Long, cntIn As Long, cntOut As Long
    Dim lRecs1 As Long, lRecs2 As Long
    Dim Minx As Double, Miny As Double, Maxx As Double, Maxy As Double
    iAns = 0
    If CrossVertices(xMainV(), yMainV()) = True Then Exit Function
    lRecs1 = UboundSngX(xMainV())
    If lRecs1 < 4 Then Exit Function
    lRecs2 = UboundSngX(xV1())
    If UboundSngX(yV1()) <> lRecs2 Then Exit Function
    cntIn = 0
    Call FindVertexMinMaxArr(xMainV(), yMainV(), Minx, Miny, Maxx, Maxy)
    For i = 1 To lRecs2
        If PointInVertexArr(xV1(i), yV1(i), xMainV(), yMainV(), Minx, Miny, Maxx, Maxy) = True Then
           cntIn = cntIn + 1
        Else
           cntOut = cntOut + 1
        End If
    Next i
    If cntIn = lRecs2 Then
       iAns = In_p
    ElseIf cntOut = lRecs2 Then
       iAns = Out_p
    Else
       iAns = Cross_p
    End If
    mComparePolygons = True
    Exit Function
Lab_Error1:
    mComparePolygons = False
End Function

Private Function mCompareCrossPolygons(xMainV() As Double, yMainV() As Double, xV1() As Double, yV1() As Double, ans As Integer) As Boolean
    Dim lRecs1 As Long, lRecs2 As Long
    Dim iL As Long, jL As Long, dne%, okint%
    Dim xa1 As Double, ya1 As Double
    Dim xa2 As Double, ya2 As Double
    Dim xb1 As Double, yb1 As Double
    Dim xb2 As Double, yb2 As Double
    Dim xn As Double, yn As Double
    ans = 0
    lRecs1 = UboundSngX(xMainV())
    If UboundSngX(yMainV) <> lRecs1 Or lRecs1 < 3 Then Exit Function
    lRecs2 = UboundSngX(xV1())
    If UboundSngX(yV1()) <> lRecs2 Or lRecs2 < 3 Then Exit Function
    mCompareCrossPolygons = True
    For iL = 1 To lRecs1 - 1
        xa1 = xMainV(iL)
        ya1 = yMainV(iL)
        xa2 = xMainV(iL + 1)
        ya2 = yMainV(iL + 1)
        For jL = 1 To lRecs2 - 1
            xb1 = xV1(jL)
            yb1 = yV1(jL)
            xb2 = xV1(jL + 1)
            yb2 = yV1(jL + 1)
            Call CheckIntersec(xa1, ya1, xa2, ya2, xb1, yb1, xb2, yb2, okint%)
            If okint% = 1 Then
               Call InterSectLines(xa1, ya1, xa2, ya2, xb1, yb1, xb2, yb2, xn, yn, dne%)
               If dne% = 2 Then
                  If FnSamePoint(xn, yn, xa1, ya1) = 0 And _
                     FnSamePoint(xn, yn, xa2, ya2) = 0 And _
                     FnSamePoint(xn, yn, xb1, yb1) = 0 And _
                     FnSamePoint(xn, yn, xb2, yb2) = 0 Then
                       ans = 1
                       Exit Function
                  End If
               End If
            End If
        Next jL
    Next iL
End Function

Function ComparePolygons(xMainV() As Double, yMainV() As Double, xV1() As Double, yV1() As Double, iAns As PolygonsComp) As Boolean
    Dim fOk1 As Boolean, fOk2 As Boolean
    Dim iAns1 As PolygonsComp
    Dim ans As Integer
    Dim xC As Double, yC As Double
    Dim MinxV As Double, MinyV As Double
    Dim MaxxV As Double, MaxyV As Double
    Dim Minx As Double, Miny As Double
    Dim Maxx As Double, Maxy As Double
    ComparePolygons = False
    If mCompareCrossPolygons(xMainV(), yMainV(), xV1(), yV1(), ans) = True Then
       If ans = 1 Then
          iAns = Cross_p
          ComparePolygons = True
          Exit Function
       End If
    End If
    fOk1 = mComparePolygons(xMainV(), yMainV(), xV1(), yV1(), iAns)
    If fOk1 = False Then Exit Function
    If iAns = Cross_p Then
       If VerticesCenter(xV1(), yV1(), xC, yC) = True Then
          If PointInVertices(xC, yC, xMainV(), yMainV()) = True Then
             iAns = In_p
          Else
             iAns = Out_p
          End If
       Else
          Exit Function
       End If
    End If
    If iAns = Out_p Or iAns = Cross_p Then
       If FindVertexMinMaxArr(xMainV(), yMainV(), Minx, Miny, Maxx, Maxy) = True Then
          If FindVertexMinMaxArr(xV1(), yV1(), MinxV, MinyV, MaxxV, MaxyV) = True Then
             If Maxx = MinxV Or _
                Maxy = MinyV Or _
                Minx = MaxxV Or _
                Miny = MaxyV Then
                  iAns = ExtContact_p
             End If
          End If
       End If
    End If
    ComparePolygons = True
End Function

Function GetAvaxShapeVertices(xyz() As Double, xV() As Double, yV() As Double, zV() As Double) As Boolean
    Dim lCnt As Integer, lCount As Long, lMax As Long
    lMax = UboundSngX(xyz)
    If lMax = 0 Or (lMax Mod 3) <> 0 Then Exit Function
    lMax = lMax / 3
    If lMax <= 2 Then Exit Function
    ReDim xV(lMax) As Double, yV(lMax) As Double, zV(lMax) As Double
    lCount = 0
    For lCnt = 1 To lMax
        lCount = lCount + 1
        xV(lCnt) = xyz(lCount)
        lCount = lCount + 1
        yV(lCnt) = xyz(lCount)
        lCount = lCount + 1
        zV(lCnt) = xyz(lCount)
    Next lCnt
    If FnSamePoint(xV(1), yV(1), xV(lMax), yV(lMax)) = 0 Then
       lCnt = lMax + 1
       ReDim Preserve xV(lCnt) As Double
       ReDim Preserve yV(lCnt) As Double
       ReDim Preserve zV(lCnt) As Double
       xV(lCnt) = xV(1)
       yV(lCnt) = yV(1)
       zV(lCnt) = zV(1)
    End If
    GetAvaxShapeVertices = True
End Function

Function SetAvaxShapeVertices(xV() As Double, yV() As Double, zV() As Double, xyz() As Double) As Boolean
    Dim lMax As Long
    Dim lCnt As Long
    Dim lDiast As Long
    Dim lCount As Long
    lMax = UboundSngX(xV())
    If lMax <> UboundSngX(yV()) Or lMax <> UboundSngX(zV()) Then Exit Function
    lDiast = lMax * 3
    If FnSamePoint(xV(1), yV(1), xV(lMax), yV(lMax)) = 1 Then
       lDiast = lDiast - 3
       lMax = lMax - 1
    End If
    ReDim xyz(lDiast) As Double
    lCount = 0
    For lCnt = 1 To lMax
        xyz(lCount + 1) = xV(lCnt)
        xyz(lCount + 2) = yV(lCnt)
        xyz(lCount + 3) = zV(lCnt)
        lCount = lCount + 3
    Next lCnt
    SetAvaxShapeVertices = True
End Function

Public Sub PolarPoint(xC As Double, yC As Double, Dist As Double, Angle As Double, xNew As Double, yNew As Double)
    xNew = xC + FnCos(Angle) * Dist
    yNew = yC + FnSin(Angle) * Dist
End Sub

Public Function OrthoPoint(xC As Double, yC As Double, x As Double, y As Double, newx As Double, newy As Double) As Boolean
    Dim dxx  As Double, dyy  As Double
    newx = x
    newy = y
    dxx = Abs(x - xC)
    dyy = Abs(y - yC)
    If dxx > dyy Then newy = yC Else newx = xC
    If newy <> y Or newx <> x Then OrthoPoint = True
End Function

Public Function SnapPoint(SnapMdx As Double, SnapMdy As Double, x As Double, y As Double, newx As Double, newy As Double) As Boolean
    newx = x
    newy = y
    If SnapMdx * SnapMdy = 0 Then Exit Function
    SnapPoint = DoStrogXY(newx, newy, SnapMdx, SnapMdy)
End Function

Private Function DoStrogXY(xNew As Double, yNew As Double, GridMdx As Double, GridMdy As Double) As Boolean
   Dim akerx1 As Double, akery1  As Double, maa1%
   Dim akerx2 As Double, akery2 As Double
   maa1% = 0
   akerx1 = Int(xNew / GridMdx) * GridMdx
   akerx2 = Int(xNew / GridMdx) * GridMdx + GridMdx
   If xNew <> akerx1 And xNew <> akerx2 Then
      If Abs(xNew - akerx1) < Abs(xNew - akerx2) Then
         xNew = akerx1
      Else
         xNew = akerx2
      End If
      maa1% = maa1% + 1
   End If
   akery1 = Int(yNew / GridMdy) * GridMdy
   akery2 = Int(yNew / GridMdy) * GridMdy + GridMdy
   If yNew <> akery1 And yNew <> akery2 Then
      If Abs(yNew - akery1) < Abs(yNew - akery2) Then
         yNew = akery1
      Else
         yNew = akery2
      End If
      maa1% = maa1% + 1
   End If
   If maa1% <> 0 Then DoStrogXY = True
End Function

Function ClipArrayLines(ByVal eselx1m As Double, ByVal esely1m As Double, ByVal eselx2m As Double, ByVal esely2m As Double, ux1() As Double, uy1() As Double, ux2() As Double, uy2() As Double, Optional lFoundClipLine As Long = 0, Optional finWin As Boolean = False) As Boolean
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim lMax As Long, lCnt As Long
    lMax = UboundSngX(ux1())
    If lFoundClipLine < 0 Then lFoundClipLine = 0
    If lFoundClipLine >= lMax Then Exit Function
    For lCnt = lFoundClipLine + 1 To lMax
        x1 = ux1(lCnt)
        y1 = uy1(lCnt)
        x2 = ux2(lCnt)
        y2 = uy2(lCnt)
        If ClipLine(eselx1m, esely1m, eselx2m, esely2m, x1, y1, x2, y2, finWin) = True Then
           If finWin = False Then
              lFoundClipLine = lCnt
              ClipArrayLines = True
              Exit Function
           End If
        Else
           If finWin = True Then Exit Function
        End If
    Next lCnt
    If finWin = True Then
       lFoundClipLine = lMax
       ClipArrayLines = True
    End If
End Function

Private Function HasSamelPoint(xa1 As Double, ya1 As Double, xa2 As Double, ya2 As Double, xb1 As Double, yb1 As Double, xb2 As Double, yb2 As Double) As Long
    If IsSamePoint(xa1, ya1, xb1, yb1) = True Then HasSamelPoint = 1: Exit Function
    If IsSamePoint(xa2, ya2, xb1, yb1) = True Then HasSamelPoint = 2: Exit Function
    If IsSamePoint(xa1, ya1, xb2, yb2) = True Then HasSamelPoint = 1: Exit Function
    If IsSamePoint(xa2, ya2, xb2, yb2) = True Then HasSamelPoint = 2: Exit Function
End Function

Public Function SortArrayLines(x1() As Double, y1() As Double, z1() As Double, x2() As Double, y2() As Double, z2() As Double) As Boolean
    Dim l As Long, j As Long
    Dim lMax As Long, lSame As Long
    lMax = UboundSngX(x1())
    If lMax = 0 Then Exit Function
    If lMax = 1 Then SortArrayLines = True: Exit Function
    l = 1
Lab_Again:
    j = l + 1
    While j <= lMax
        lSame = HasSamelPoint(x1(j), y1(j), x2(j), y2(j), x1(l), y1(l), x2(l), y2(l))
        If lSame <> 0 Then
           If j <> l + 1 Then
              If lSame = 1 Then
                 Swap x1(j), x1(l + 1)
                 Swap y1(j), y1(l + 1)
                 Swap z1(j), z1(l + 1)
                 Swap x2(j), x2(l + 1)
                 Swap y2(j), y2(l + 1)
                 Swap z2(j), z2(l + 1)
              Else
                 Swap x1(j), x2(l + 1)
                 Swap y1(j), y2(l + 1)
                 Swap z1(j), z2(l + 1)
                 Swap x2(j), x1(l + 1)
                 Swap y2(j), y1(l + 1)
                 Swap z2(j), z1(l + 1)
              End If
           End If
           l = l + 1
           GoTo Lab_Again
        End If
        j = j + 1
    Wend
    SortArrayLines = True
End Function

Public Function GetArrayLinesBorder(x1() As Double, y1() As Double, z1() As Double, x2() As Double, y2() As Double, z2() As Double, xMin As Double, yMin As Double, zMin As Double, xMax As Double, yMax As Double, zMax As Double) As Boolean
    Dim lMax As Long
    Dim lCnt As Long
    lMax = UboundSngX(x1)
    If lMax = 0 Then Exit Function
    GetArrayLinesBorder = True
    If lMax <= 1 Then Exit Function
    For lCnt = 1 To lMax
        If lCnt = 1 Then
           xMin = x1(lCnt)
           yMin = y1(lCnt)
           xMax = x1(lCnt)
           yMax = y1(lCnt)
        End If
        If x1(lCnt) < xMin Then xMin = x1(lCnt)
        If y1(lCnt) < yMin Then yMin = y1(lCnt)
        If x1(lCnt) > xMax Then xMax = x1(lCnt)
        If y1(lCnt) > yMax Then yMax = y1(lCnt)
        If x2(lCnt) < xMin Then xMin = x2(lCnt)
        If y2(lCnt) < yMin Then yMin = y2(lCnt)
        If x2(lCnt) > xMax Then xMax = x2(lCnt)
        If y2(lCnt) > yMax Then yMax = y2(lCnt)
    Next lCnt
End Function

Public Function GetNearestArrayLineFromPoint(ax1() As Double, ay1() As Double, ax2() As Double, ay2() As Double, x As Double, y As Double) As Long
    Dim fFirst As Boolean
    Dim DistMinIn As Double
    Dim lMax As Long, lCnt As Long
    Dim x1 As Double, y1 As Double
    Dim x2 As Double, y2 As Double
    Dim xp As Double, yp As Double
    Dim ok As Double, sDist As Double
    lMax = UboundSngX(ax1())
    If lMax = 0 Then Exit Function
    For lCnt = 1 To lMax
        x1 = ax1(lCnt)
        y1 = ay1(lCnt)
        x2 = ax2(lCnt)
        y2 = ay2(lCnt)
        Call VertFromPointToLine(x, y, x1, y1, x2, y2, xp, yp)
        ok = 0
        Call CheckPointOnLine(x1, y1, x2, y2, xp, yp, ok)
        If ok <> 0 Then
           Call Distance(x, y, xp, yp, sDist)
           If fFirst = False Then
              DistMinIn = sDist
              GetNearestArrayLineFromPoint = lCnt
              fFirst = True
           End If
           If DistMinIn > sDist Then
              DistMinIn = sDist
              GetNearestArrayLineFromPoint = lCnt
           End If
        End If
    Next lCnt
End Function

Public Function InvertMatrix(Mat() As Double) As Boolean
    Dim iBound%, l%, i%, k%
    On Local Error GoTo Lab_Error
    iBound% = UBound(Mat, 1)
    If UBound(Mat, 2) <> iBound% Then Exit Function
    ReDim OldMat(iBound%, iBound%) As Double
    OldMat = Mat
    For l% = 1 To iBound%
        For i% = 1 To iBound%
           If i% <> l% Then Mat(l%, i%) = Mat(l%, i%) / Mat(l%, l%)
        Next i%
        For k% = 1 To iBound%
            If k% <> l% Then
               For i% = 1 To iBound%
                  If i% <> l% Then
                      Mat(k%, i%) = Mat(k%, i%) - Mat(l%, i%) * Mat(k%, l%)
                  End If
               Next i%
            End If
        Next k%
        For k% = 1 To iBound%
            If k% <> l% Then Mat(k%, l%) = -Mat(k%, l%) / Mat(l%, l%)
        Next k%
        Mat(l%, l%) = 1! / Mat(l%, l%)
    Next l%
    InvertMatrix = True
    Exit Function
Lab_Error:
    Mat = OldMat
    InvertMatrix = False
End Function

Function PutAngle(ByVal xx1 As Double, ByVal yy1 As Double, ByVal xx2 As Double, ByVal yy2 As Double, ByVal xd1 As Double, ByVal yd1 As Double, ByVal xd2 As Double, ByVal yd2 As Double, ByVal x1dist As Double, ByVal y1dist As Double, ByVal akr As Double, ByVal DimSynt As Double, ByVal DefaultText As String, ByVal DTStyle%, ByVal DTxtScl As Double, ByVal DimArrow%, ByVal AngTickSize As Double) As String
    Dim Angle As Double, ok As Double
    Dim xC As Double, yC As Double, zC As Double, dne1%
    Dim R As Double
    Dim xV() As Double, yV() As Double, zV() As Double
    Dim Startx As Double, Starty As Double
    Dim Endx As Double, Endy As Double
    Dim Text$, TxtScl As Double
    Dim lMax As Long, lCnt As Long
    Dim theta1 As Double, iTrns%, TStyle%
    Dim Angle1 As Double, Angle2 As Double, Angle3 As Double
    Dim StartAngle As Double, EndAngle As Double
    Dim aaR As Double, x As Double, y As Double, z As Double
    Dim nx As Double, ny As Double, nz As Double
    Dim oStr As cAddString
    On Local Error Resume Next
    Call InterSectLines(xx1, yy1, xx2, yy2, xd1, yd1, xd2, yd2, xC, yC, dne1%)
    If dne1% = 0 Then Exit Function
    Set oStr = New cAddString
    oStr.BeginAdd
    Call Angle2Lines(x1dist, y1dist, xx1, yy1, xx2, yy2, xd1, yd1, xd2, yd2, Angle, ok)
    Call Distance(xC, yC, x1dist, y1dist, R)
    If FnSamePoint(xx1, yy1, xC, yC) = 1 Then
       Startx = xx2
       Starty = yy2
    Else
       Startx = xx1
       Starty = yy1
    End If
    If FnSamePoint(xd1, yd1, xC, yC) = 1 Then
       Endx = xd2
       Endy = yd2
    Else
       Endx = xd1
       Endy = yd1
    End If
    Call FindRealAngle(Startx, Starty, xC, yC, theta1)
    Angle2 = theta1
    Call PolarPoint(xC, yC, R, theta1, Startx, Starty)
    Call FindRealAngle(Endx, Endy, xC, yC, theta1)
    Angle1 = theta1
    Call PolarPoint(xC, yC, R, theta1, Endx, Endy)
    Call FindRealAngle(x1dist, y1dist, xC, yC, Angle3)
    If ok <> 0 Then
       If Angle1 > Angle2 Then
          Swap Angle1, Angle2
          Swap xx1, xd1
          Swap yy1, yd1
          Swap xx2, xd2
          Swap yy2, yd2
       End If
       If Angle3 > Angle2 Or Angle3 < Angle1 Then
          Swap Angle1, Angle2
          Angle1 = Angle1 + 360
       End If
       Call NewArc(xC, yC, zC, R, Angle1, Angle2, oStr)
       If DimArrow% = 1 Then
          Call PolarPoint(xC, yC, R, Angle1, x, y)
          Call PolarPoint(x, y, AngTickSize, Angle1 + 45, nx, ny)
          Call NewDrawLine(nx, ny, nz, x, y, z, oStr)
          Call PolarPoint(x, y, AngTickSize, Angle1 + 135, nx, ny)
          Call NewDrawLine(nx, ny, nz, x, y, z, oStr)
          Call PolarPoint(xC, yC, R, Angle2, x, y)
          Call PolarPoint(x, y, AngTickSize, Angle2 - 45, nx, ny)
          Call NewDrawLine(nx, ny, nz, x, y, z, oStr)
          Call PolarPoint(x, y, AngTickSize, Angle2 - 135, nx, ny)
          Call NewDrawLine(nx, ny, nz, x, y, z, oStr)
       ElseIf DimArrow% = 2 Then
          aaR = AngTickSize / 2
          Call PolarPoint(xC, yC, R, Angle1, x, y)
          Call NewDrawCircle(x, y, z, aaR, oStr)
          Call PolarPoint(xC, yC, R, Angle2, x, y)
          Call NewDrawCircle(x, y, z, aaR, oStr)
       End If
       Call UsingText(akr, Angle * DimSynt, Text$)
       Text$ = Text$ & " " & DefaultText
       Text$ = RTrim$(Text$)
       TStyle% = DTStyle%
       TxtScl = DTxtScl
       theta1 = 0
       iTrns% = 0
       Call NewDrawText(TStyle%, Text$, x1dist, y1dist, zC, TxtScl, iTrns%, theta1, oStr)
    End If
    PutAngle = oStr.CurString
    Set oStr = Nothing
End Function

Function PutDimension(mthd%, xx1 As Double, yy1 As Double, xx2 As Double, yy2 As Double, OffsetD As Double, x1dist As Double, y1dist As Double, xd1 As Double, yd1 As Double, xd2 As Double, yd2 As Double, DimScale As Double, DimTadOn%, DimTihOn%, DimAsz As Double, DimTsz As Double, DimExe As Double, DimExo As Double, TxtScl As Double, akr As Double, DimT$, DimSynt As Double, DimArrow%, DimVertToLine%, TStyle%) As String
    Dim oStr As cAddString
    Set oStr = New cAddString
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    Dim TextWX As Double, TextHY As Double
    Dim TxtWX As Double, TxtHY As Double
    Dim TextLen As Double, cclr%
    Dim oriz%, pano%, d As Double, Dist$, Text$, pointKp%
    Dim x1d As Double, y1d As Double, xyd As Double
    Dim x11 As Double, y11 As Double, xC As Double, yC As Double
    Dim x22 As Double, y22 As Double, snth As Double, prd1 As Double
    Dim theta As Double, z As Double, StartTheta As Double, th As Double
    Dim prd As Double, D1 As Double, D2 As Double, iTrns%
    Dim nxt As Double, nyt As Double, nx1 As Double, nx2 As Double
    Dim ny1 As Double, ny2 As Double, xde As Double, yde As Double
    Dim dl As Double, x As Double, y As Double, kp As Double, dne1%
    Dim xt As Double, yt As Double, done%, aaR As Double, bbR As Double
    Dim ecz1 As Double, ecz2 As Double, far1 As Double, far2 As Double
    Dim ecx1 As Double, ecy1 As Double, ecx2 As Double, ecy2 As Double
    Dim hxn As Double, hyn As Double, ok As Double
    m_count = 0
    Call Equal4Double(x1, y1, x2, y2, xx1, yy1, xx2, yy2)
    Rem *----------------------------*
    If TextWX = 0 Then TextWX = 0.2
    If TextHY = 0 Then TextHY = 0.2
    If TxtScl = 0 Then TxtScl = 1
    If DimScale = 0 Then DimScale = 1
    Rem *----------------------------*
    TxtScl = TxtScl * DimScale
    TxtWX = TextWX * TxtScl
    TxtHY = TextHY * TxtScl
    If TxtHY = 0 Then Exit Function
    DimAsz = DimAsz * DimScale
    DimTsz = DimTsz * DimScale
    DimExe = DimExe * DimScale
    DimExo = DimExo * DimScale
    Rem *----------------------------*
    oriz% = DimTihOn%
    pano% = DimTadOn%
    Call Distance(x1, y1, x2, y2, d)
    d = d * DimSynt
    Call UsingText(akr, d, Dist$)
    Dist$ = Dist$ + DimT$
    Rem *----------------------------*
    TextLen = 0.5 * Len(Dist$) * TxtWX
    Rem *----------------------------*
    If mthd% = 0 Then
       Call OffsetLinePoint(x1, y1, x2, y2, OffsetD, x1dist, y1dist, x1d, y1d)
       If x1d = 0 And y1d = 0 Then Exit Function
       x11 = x1 + x1d
       y11 = y1 + y1d
       x22 = x2 + x1d
       y22 = y2 + y1d
    Else
       Call Equal4Double(x11, y11, x22, y22, xd1, yd1, xd2, yd2)
    End If
    oStr.BeginAdd
    Call FindRealAngle(x22, y22, x11, y11, theta): z = 0
    If theta > 90 And theta <= 270 Then
       Swap x11, x22
       Swap y11, y22
       Swap x1, x2
       Swap y1, y2
       Call FindRealAngle(x22, y22, x11, y11, theta)
    End If
    StartTheta = theta
    If DimAsz >= d / 2 Then DimAsz = d / 3
    If pano% = 0 Then
       snth = Abs(FnSin(theta)): prd1 = 2 * DimAsz
       If oriz% = 1 And FnZero(snth) = 0 Then prd = TxtHY / snth Else prd = TextLen
       If prd > TextLen Then prd = TextLen
       D1 = d - prd - prd1: D1 = D1 / 2: If D1 < 0 Then D1 = 0: D2 = 0: pano% = 1: GoTo Lab1
       D2 = D1 + prd + prd1: If D2 > d Then D1 = 0: D2 = 0: pano% = 1: GoTo Lab1
    End If
Lab1:
    iTrns% = 0
    ReDim tk2(3, 3) As Double
    Call MATidn(tk2())
    Text$ = Dist$
    Call LineCenter(x11, y11, x22, y22, xC, yC)
    If pano% = 0 Then
       nxt = x11 + FnCos(theta) * (D1 + DimAsz): nyt = y11 + FnSin(theta) * (D1 + DimAsz)
       nx1 = x11 + FnCos(theta) * D1: ny1 = y11 + FnSin(theta) * D1
       Call NewDrawLine(x11, y11, z, nx1, ny1, z, oStr)
       nx2 = x11 + FnCos(theta) * D2: ny2 = y11 + FnSin(theta) * D2
       Call NewDrawLine(nx2, ny2, z, x22, y22, z, oStr)
       If oriz% = 1 And theta <> 0 Then
          xde = TextLen / 2
          nxt = xC - xde
          nyt = yC - TxtHY / 2
       End If
    Else
       Call NewDrawLine(x11, y11, z, x22, y22, z, oStr)
       xde = TextLen / 2
       yde = TxtHY / 3
       xyd = Sqr(xde ^ 2 + yde ^ 2)
       th = FnArcTan(xde / yde) + 90 + theta
       nxt = xC + FnCos(th) * xyd
       nyt = yC + FnSin(th) * xyd
       If oriz% = 1 And theta <> 0 Then
          xde = TextLen / 2
          nxt = xC - xde * 2 - yde
          nyt = yC - TxtHY / 2
       End If
    End If
    pointKp% = 0
    If DimTsz <> 0 Then dl = DimTsz: x = x11: y = y11: GoSub SetLines1: x = x22: y = y22: GoSub SetLines1
    If DimExe <> 0 Then dl = DimExe: kp = 0: x = x11: y = y11: xt = x1: yt = y1: GoSub SetLines2: x = x22: y = y22: xt = x2: yt = y2: GoSub SetLines2
    If DimExo <> 0 Then
       dl = DimExo: kp = 1: x = x11
       y = y11: xt = x1: yt = y1
       GoSub SetLines2: x = x22
       y = y22: xt = x2: yt = y2
       GoSub SetLines2
       pointKp% = 1
       dl = DimExo + DimTsz: kp = 1
       x = x11: y = y11: xt = x1: yt = y1
       GoSub SetLines2: x = x22
       y = y22: xt = x2: yt = y2
       GoSub SetLines2
    End If
    If oriz% = 0 And theta <> 0 Then
       Call Rot2dMat(theta, nxt, nyt, done%, tk2())
       iTrns% = 1
    End If
    If DimVertToLine% = 1 Then
       If pano% = 0 Then
          theta = StartTheta + 90
          If theta >= 180 Then theta = theta - 180
          xde = -TextLen / 2
          yde = TxtHY / 3
          xyd = Sqr(xde ^ 2 + yde ^ 2)
          th = 90 + theta
          nxt = xC + FnCos(th) * xyd
          nyt = yC + FnSin(th) * xyd
          If StartTheta = 90 Then
             nyt = nyt - TxtHY
             nxt = nxt - TxtHY / 2
          ElseIf StartTheta = 0 Then
             nxt = nxt + TxtHY
             nyt = nyt - TxtHY / 2
          End If
          Call Rot2dMat(theta, nxt, nyt, done%, tk2())
          iTrns% = 1
       Else
          theta = StartTheta + 90
          If theta >= 180 Then theta = theta - 180
          xde = -TextLen / 2
          yde = TxtHY / 3
          xyd = Sqr(xde ^ 2 + yde ^ 2)
          th = 90 + theta
          nxt = xC + FnCos(th) * xyd
          nyt = yC + FnSin(th) * xyd
          If StartTheta = 90 Then
             nyt = nyt - TxtHY
             Text$ = "  " & Text$
          ElseIf StartTheta = 0 Then
             nxt = nxt + TxtHY
             Text$ = " " & Text$
          Else
             Text$ = " " & Text$
          End If
          Call Rot2dMat(theta, nxt, nyt, done%, tk2())
          iTrns% = 1
       End If
    End If
    Call NewDrawText(TStyle%, Text$, nxt, nyt, z, TxtScl, iTrns%, theta, oStr)
    Erase tk2
    Dist$ = ""
    PutDimension = oStr.CurString
    Set oStr = Nothing
    Exit Function
Rem ==================================
SetLines1:
    If DimArrow% = 2 Then
       aaR = dl / 2
       Call NewDrawCircle(x, y, z, aaR, oStr)
    Else
       far1 = 45
       If DimArrow% = 0 Then far2 = 225 Else far2 = -45
Lab2:
       nx1 = x + FnCos(theta + far1) * dl: ny1 = y + FnSin(theta + far1) * dl
       nx2 = x + FnCos(theta + far2) * dl: ny2 = y + FnSin(theta + far2) * dl
       If DimArrow% = 0 Then
          Call NewDrawLine(nx1, ny1, z, nx2, ny2, z, oStr)
       Else
          dne1% = 0
          Call InterSectLines(x11, y11, x22, y22, nx1, ny1, nx2, ny2, hxn, hyn, dne1%)
          If dne1% <> 2 And far1 = 45 Then far1 = 135: far2 = 225: GoTo Lab2
          Call NewDrawLine(nx1, ny1, z, x, y, z, oStr)
          Call NewDrawLine(x, y, z, nx2, ny2, z, oStr)
       End If
    End If
    Return
Rem ==================================
SetLines2:
    nx1 = x + FnCos(theta + 90) * dl
    ny1 = y + FnSin(theta + 90) * dl
    ok = 0
    Call CheckPointOnLine(xt, yt, x, y, nx1, ny1, ok)
    If ok <> kp Then
       nx1 = x + FnCos(theta + 270) * dl
       ny1 = y + FnSin(theta + 270) * dl
    End If
    If pointKp% = 0 Then
       Call NewDrawLine(nx1, ny1, z, x, y, z, oStr)
    Else
       Call NewDrawLine(nx1, ny1, z, nx1, ny1, z, oStr)
    End If
    Return
Rem ==================================
End Function

Private Sub UsingText(vAkr As Double, ByVal mNumber As Double, sText As String)
    Call strogilS(mNumber, vAkr)
    sText = sFormatDbl(mNumber, vAkr)
End Sub

Private Sub MakeRealEllipseRange(fx As Double, fy As Double, aaR As Double, bbR As Double, ecx1 As Double, ecy1 As Double, ecx2 As Double, ecy2 As Double)
    ecx1 = fx - aaR
    ecy1 = fy - bbR
    ecx2 = fx + aaR
    ecy2 = fy + bbR
End Sub

Private Function sFormatDbl(n As Double, akrib As Double) As String
    Dim sTxt As String
    Call FormatDbl(n, akrib, sTxt)
    sFormatDbl = sTxt
End Function

Private Sub FormatDbl(n As Double, akrib As Double, sTxt As String)
    Dim sformat As String
    If n = 0! Then
        sTxt = " "
    Else
        sformat = "##,##0"
        If (akrib > 0) Then sformat = sformat & "." & String$(akrib, "0")
        sTxt = Format$(n, sformat & ";-" & sformat & "; ")
    End If
End Sub

Private Sub NewDrawLine(x1 As Double, y1 As Double, z1 As Double, x2 As Double, y2 As Double, z2 As Double, oStr As cAddString)
    Dim oLine As cLine
    Set oLine = New cLine
    m_count = m_count + 1
    If m_count = 1 Then
       m_xMin = x1
       m_yMin = y1
       m_zMin = z1
       m_xMax = x2
       m_yMax = y2
       m_zMax = z2
    End If
    If m_xMin > x1 Then m_xMin = x1
    If m_xMin > x2 Then m_xMin = x2
    If m_yMin > y1 Then m_yMin = y1
    If m_yMin > y2 Then m_yMin = y2
    If m_zMin > z1 Then m_zMin = z1
    If m_zMin > z2 Then m_zMin = z2
    If m_xMax < x1 Then m_xMax = x1
    If m_xMax < x2 Then m_xMax = x2
    If m_yMax < y1 Then m_yMax = y1
    If m_yMax < y2 Then m_yMax = y2
    If m_zMax < z1 Then m_zMax = z1
    If m_zMax < z2 Then m_zMax = z2
    oLine.LineTypeName = LineTypeName
    oLine.ColorIndex = ColorIndex
    oLine.LayerName = LayerName
    oLine.x1 = x1
    oLine.y1 = y1
    oLine.z1 = z1
    oLine.x2 = x2
    oLine.y2 = y2
    oLine.z2 = z2
    oStr.AddString oLine.DxfLine
    Set oLine = Nothing
End Sub

Private Sub NewArc(x1 As Double, y1 As Double, z1 As Double, R As Double, StartAngle As Double, EndAngle As Double, oStr As cAddString)
    Dim oArc As cArc
    Set oArc = New cArc
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
    oArc.ColorIndex = ColorIndex
    oArc.LayerName = LayerName
    oArc.x = x1
    oArc.y = y1
    oArc.z = z1
    oArc.R = R
    oArc.StartAngle = StartAngle
    oArc.EndAngle = EndAngle
    oStr.AddString oArc.DxfArc
    Set oArc = Nothing
End Sub

Private Sub NewDrawCircle(x As Double, y As Double, z As Double, aaR As Double, oStr As cAddString)
    Dim oCircle As cCircle
    Dim dCircleLength As Double
    Set oCircle = New cCircle
    oCircle.ColorIndex = ColorIndex
    oCircle.LayerName = LayerName
    oCircle.x = x
    oCircle.y = y
    oCircle.z = z
    oCircle.R = aaR
    m_count = m_count + 1
    If m_count = 1 Then
       m_xMin = x
       m_yMin = y
       m_zMin = z
    End If
    dCircleLength = aaR
    If m_xMin > x - dCircleLength Then m_xMin = x - dCircleLength
    If m_yMin > y - dCircleLength Then m_yMin = y - dCircleLength
    If m_zMin > z - dCircleLength Then m_zMin = z - dCircleLength
    If m_xMax < x + dCircleLength Then m_xMax = x + dCircleLength
    If m_yMax < y + dCircleLength Then m_yMax = y + dCircleLength
    If m_zMax < z + dCircleLength Then m_zMax = z + dCircleLength
    oStr.AddString oCircle.DxfCircle
    Set oCircle = Nothing
End Sub

Private Sub NewDrawText(TStyle%, Text$, x As Double, y As Double, z As Double, Height As Double, iTrns%, theta As Double, oStr As cAddString)
    Dim oText As cText
    Dim dTextLength As Double
    Set oText = New cText
    oText.ColorIndex = ColorIndex
    oText.LayerName = LayerName
    oText.Text = Text$
    oText.x = x
    oText.y = y
    oText.z = z
    oText.Height = Height
    oText.Angle = theta
    m_count = m_count + 1
    If m_count = 1 Then
       m_xMin = x
       m_yMin = y
       m_zMin = z
    End If
    dTextLength = Height * Len(Text)
    If m_xMin > x - dTextLength Then m_xMin = x - dTextLength
    If m_yMin > y - dTextLength Then m_yMin = y - dTextLength
    If m_zMin > z - dTextLength Then m_zMin = z - dTextLength
    If m_xMax < x + dTextLength Then m_xMax = x + dTextLength
    If m_yMax < y + dTextLength Then m_yMax = y + dTextLength
    If m_zMax < z + dTextLength Then m_zMax = z + dTextLength
    oStr.AddString oText.DxfText
    Set oText = Nothing
End Sub

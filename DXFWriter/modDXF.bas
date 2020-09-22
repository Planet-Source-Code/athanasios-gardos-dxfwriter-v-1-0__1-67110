Attribute VB_Name = "modDXF"
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

Public Function DxfNb(nb%, x As Double) As String
    DxfNb = " " & Str$(nb%) & vbCrLf & dFormat(x) & vbCrLf
End Function

Public Function dFormat(x As Double) As String
    Dim sTmp As String, i As Long
    If x = 0 Then
       sTmp = "0.0"
    ElseIf Abs(x) < 1 Then
       sTmp = Format$(x, "###0.0000000000000000")
       Call ChangeChr(sTmp, ",", ".")
       For i = Len(sTmp) To 1 Step -1
           If Asc(Mid$(sTmp, i, 1)) = 46 Then
              Exit For
           ElseIf Asc(Mid$(sTmp, i, 1)) <> 48 Then
              sTmp = Mid$(sTmp, 1, i)
              Exit For
           End If
       Next i
    Else
       sTmp = Format$(x)
       Call ChangeChr(sTmp, ",", ".")
    End If
    dFormat = sTmp
End Function


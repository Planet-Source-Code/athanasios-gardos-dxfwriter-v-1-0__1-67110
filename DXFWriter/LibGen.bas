Attribute VB_Name = "LibGen"
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

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Declare Function GetSystemDirectory Lib "kernel32" _
                 Alias "GetSystemDirectoryA" ( _
                 ByVal lpBuffer As String, _
                 ByVal nSize As Long) As Long
                 
Public Declare Function GetTempPath Lib "kernel32" _
                 Alias "GetTempPathA" ( _
                 ByVal nBufferLength As Long, _
                 ByVal lpBuffer As String) As Long
                 
Public Declare Function GetTempFileName Lib "kernel32" _
                 Alias "GetTempFileNameA" ( _
                 ByVal lpszPath As String, _
                 ByVal lpPrefixString As String, _
                 ByVal wUnique As Long, _
                 ByVal lpTempFileName As String) As Long
                 
Public Const cMaxPath = 260
Public Const TwipsPerPoint% = 20
Public Const TwipsPerInch% = 1440
Public Const TwipsPerCm% = 567
Public Const paperA4w% = 11906
Public Const paperA4h% = 16837

Public D_DemoMSG As String
Public fDemoMode As Boolean
Public m_UserSerialNumber As String
Public m_UserName As String

Public BlackColorIndex As Long
Public WhiteColorIndex As Long

Public Const sEmpty As String = ""

Type POINTAPI
     x As Long
     y As Long
End Type

Sub FileNotFound(fl$)
    MsgBox "file not found =" & fl$, vbCritical, "Error"
End Sub

Function GetTempFile(Optional Prefix As String, Optional PathName As String) As String
    If Prefix = sEmpty Then Prefix = "~gv"
    If PathName = sEmpty Then PathName = GetTempDir
    Dim sRet As String
    sRet = String(cMaxPath, 0)
    Call GetTempFileName(PathName, Prefix, 0, sRet)
    Call ChangeChr(sRet, Chr$(0), Chr$(32))
    sRet = RTrim$(sRet)
    GetTempFile = sRet
End Function

Function GetTempDir() As String
    Dim sRet As String
    Dim c As Long
    Dim sDir As String
    sRet = String(cMaxPath, 0)
    c = GetTempPath(cMaxPath, sRet)
    If c <> 0 Then
       sDir = Left$(sRet, c)
       GetTempDir = NormalizePath(sDir)
    End If
End Function

Function UboundVarX(a() As Variant) As Long
    On Local Error GoTo Lab_Error
    UboundVarX = UBound(a)
    Exit Function
Lab_Error:
    UboundVarX = 0
End Function

Function UboundLngX(a() As Long) As Long
    On Local Error GoTo Lab_Error
    UboundLngX = UBound(a)
    Exit Function
Lab_Error:
    UboundLngX = 0
End Function

Function UboundSngX(a() As Single) As Long
    On Local Error GoTo Lab_Error
    UboundSngX = UBound(a)
    Exit Function
Lab_Error:
    UboundSngX = 0
End Function

Function UboundByteX(a() As Byte) As Long
    On Local Error GoTo Lab_Error
    UboundByteX = UBound(a)
    Exit Function
Lab_Error:
    UboundByteX = 0
End Function

Function UboundStrX(sArray() As String) As Integer
    On Local Error GoTo Lab_Err
    UboundStrX = UBound(sArray)
    Exit Function
Lab_Err:
    UboundStrX = 0
End Function

Function QBColorX(clr As Integer) As Long
    Dim b As Integer
    b = clr Mod 16
    QBColorX = QBColor(b)
End Function

Function SystemDirectory() As String
    Dim buffer As String * 512, Length As Long
    Length = GetSystemDirectory(buffer, Len(buffer))
    SystemDirectory = Left$(buffer, Length)
End Function

Sub Main()
    Dim sSN As String, sUserName As String
    D_DemoMSG = App.ProductName & " v." & Format$(App.Major) & "." & Format$(App.Minor)
    m_UserSerialNumber = D_DemoMSG
    m_UserName = ""
    fDemoMode = True
    Call MsgBoxAbout
End Sub

Private Sub MsgBoxAbout()
    If fDemoMode = True Then
       frmAbout.Show 1
    End If
End Sub

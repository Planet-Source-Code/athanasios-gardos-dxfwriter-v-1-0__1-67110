VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHeader"
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

Public EXTMIN_x As Double
Public EXTMIN_y As Double
Public EXTMIN_z As Double
Public EXTMAX_x As Double
Public EXTMAX_y As Double
Public EXTMAX_z As Double
Public LineTypeScale As Double

Dim sLineTypes() As String
Dim sFonts() As String
Dim sLayers() As String
Dim iLayerFlags() As Integer
Dim iLayerColorIndex() As Integer
Dim sLayerLineTypeName() As String

Public Function DxfHeader() As String
    Dim oStr As cAddString
    Set oStr = New cAddString
    oStr.BeginAdd
    oStr.Add2Strings "  0", vbCrLf
    oStr.Add2Strings "SECTION", vbCrLf
    oStr.Add2Strings "  2", vbCrLf
    oStr.Add2Strings "HEADER", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$ACADVER", vbCrLf
    oStr.Add2Strings "  1", vbCrLf
    oStr.Add2Strings "AC1009", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$ACADMAINTVER", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     6", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DWGCODEPAGE", vbCrLf
    oStr.Add2Strings "  3", vbCrLf
    oStr.Add2Strings "ANSI_1252", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$INSBASE", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 30", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$EXTMIN", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings dFormat(EXTMIN_x), vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings dFormat(EXTMIN_y), vbCrLf
    oStr.Add2Strings " 30", vbCrLf
    oStr.Add2Strings dFormat(EXTMIN_z), vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$EXTMAX", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings dFormat(EXTMAX_x), vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings dFormat(EXTMAX_y), vbCrLf
    oStr.Add2Strings " 30", vbCrLf
    oStr.Add2Strings dFormat(EXTMAX_z), vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$LIMMIN", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$LIMMAX", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings "12.0", vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings "9.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$ORTHOMODE", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$REGENMODE", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     1", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$FILLMODE", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     1", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$QTEXTMODE", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$MIRRTEXT", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     1", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$LTSCALE", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    If LineTypeScale <= 0 Then LineTypeScale = 1
    oStr.Add2Strings dFormat(LineTypeScale), vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$ATTMODE", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     1", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$TEXTSIZE", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.2", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$TRACEWID", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.05", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$TEXTSTYLE", vbCrLf
    oStr.Add2Strings "  7", vbCrLf
    oStr.Add2Strings "STANDARD", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$CLAYER", vbCrLf
    oStr.Add2Strings "  8", vbCrLf
    oStr.Add2Strings "0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$CELTYPE", vbCrLf
    oStr.Add2Strings "  6", vbCrLf
    oStr.Add2Strings "CONTINUOUS", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$CECOLOR", vbCrLf
    oStr.Add2Strings " 62", vbCrLf
    oStr.Add2Strings "   256", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$CELTSCALE", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "1.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DISPSILH", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMSCALE", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "3.229813999999999", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMASZ", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.09", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMEXO", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.03125", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMDLI", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.38", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMRND", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMDLE", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMEXE", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.05", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMTP", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMTM", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMTXT", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.1", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMCEN", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.09", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMTSZ", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMTOL", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMLIM", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMTIH", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     1", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMTOH", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     1", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMSE1", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMSE2", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMTAD", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMZIN", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     8", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMBLK", vbCrLf
    oStr.Add2Strings "  1", vbCrLf
    oStr.Add2Strings "", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMASO", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     1", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMSHO", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     1", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMPOST", vbCrLf
    oStr.Add2Strings "  1", vbCrLf
    oStr.Add2Strings "", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMAPOST", vbCrLf
    oStr.Add2Strings "  1", vbCrLf
    oStr.Add2Strings "", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMALT", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMALTD", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     2", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMALTF", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "25.4", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMLFAC", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "1.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMTOFL", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMTVP", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMTIX", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMSOXD", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMSAH", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMBLK1", vbCrLf
    oStr.Add2Strings "  1", vbCrLf
    oStr.Add2Strings "", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMBLK2", vbCrLf
    oStr.Add2Strings "  1", vbCrLf
    oStr.Add2Strings "", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMSTYLE", vbCrLf
    oStr.Add2Strings "  2", vbCrLf
    oStr.Add2Strings "STANDARD", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMCLRD", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     5", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMCLRE", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     5", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMCLRT", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     7", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMTFAC", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "1.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMGAP", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.09", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMJUST", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMSD1", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMSD2", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMTOLJ", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     1", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMTZIN", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMALTZ", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMALTTZ", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMUPT", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMDEC", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     3", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMTDEC", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     4", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMALTU", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     2", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMALTTD", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     2", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMTXSTY", vbCrLf
    oStr.Add2Strings "  7", vbCrLf
    oStr.Add2Strings "STANDARD", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMAUNIT", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMADEC", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMALTRND", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMAZIN", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMDSEP", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "    46", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMATFIT", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     3", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMFRAC", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMLDRBLK", vbCrLf
    oStr.Add2Strings "  1", vbCrLf
    oStr.Add2Strings "", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMLUNIT", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     2", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMLWD", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "    -2", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMLWE", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "    -2", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$DIMTMOVE", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$LUNITS", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     2", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$LUPREC", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     4", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$SKETCHINC", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.1", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$FILLETRAD", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$AUNITS", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$AUPREC", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$MENU", vbCrLf
    oStr.Add2Strings "  1", vbCrLf
    oStr.Add2Strings ".", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$ELEVATION", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$PELEVATION", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$THICKNESS", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$LIMCHECK", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$CHAMFERA", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$CHAMFERB", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$CHAMFERC", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$CHAMFERD", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$SKPOLY", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$TDCREATE", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "2452514.408270833", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$TDUCREATE", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "2452514.574937500", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$TDUPDATE", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "2452514.408283680", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$TDUUPDATE", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "2452514.574950347", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$TDINDWG", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.0000130208", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$TDUSRTIMER", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.0000130208", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$USRTIMER", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     1", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$ANGBASE", vbCrLf
    oStr.Add2Strings " 50", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$ANGDIR", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$PDMODE", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$PDSIZE", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$PLINEWID", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$SPLFRAME", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$SPLINETYPE", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     6", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$SPLINESEGS", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     8", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$SURFTAB1", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     6", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$SURFTAB2", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     6", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$SURFTYPE", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     6", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$SURFU", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     6", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$SURFV", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     6", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$UCSBASE", vbCrLf
    oStr.Add2Strings "  2", vbCrLf
    oStr.Add2Strings "", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$UCSNAME", vbCrLf
    oStr.Add2Strings "  2", vbCrLf
    oStr.Add2Strings "", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$UCSORG", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 30", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$UCSXDIR", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings "1.0", vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 30", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$UCSYDIR", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings "1.0", vbCrLf
    oStr.Add2Strings " 30", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$UCSORTHOREF", vbCrLf
    oStr.Add2Strings "  2", vbCrLf
    oStr.Add2Strings "", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$UCSORTHOVIEW", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$UCSORGTOP", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 30", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$UCSORGBOTTOM", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 30", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$UCSORGLEFT", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 30", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$UCSORGRIGHT", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 30", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$UCSORGFRONT", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 30", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$UCSORGBACK", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 30", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$PUCSBASE", vbCrLf
    oStr.Add2Strings "  2", vbCrLf
    oStr.Add2Strings "", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$PUCSNAME", vbCrLf
    oStr.Add2Strings "  2", vbCrLf
    oStr.Add2Strings "", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$PUCSORG", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 30", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$PUCSXDIR", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings "1.0", vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 30", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$PUCSYDIR", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings "1.0", vbCrLf
    oStr.Add2Strings " 30", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$PUCSORTHOREF", vbCrLf
    oStr.Add2Strings "  2", vbCrLf
    oStr.Add2Strings "", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$PUCSORTHOVIEW", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$PUCSORGTOP", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 30", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$PUCSORGBOTTOM", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 30", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$PUCSORGLEFT", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 30", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$PUCSORGRIGHT", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 30", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$PUCSORGFRONT", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 30", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$PUCSORGBACK", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 30", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$USERI1", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$USERI2", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$USERI3", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$USERI4", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$USERI5", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$USERR1", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$USERR2", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$USERR3", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$USERR4", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$USERR5", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$WORLDVIEW", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     1", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$SHADEDGE", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     3", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$SHADEDIF", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "    70", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$TILEMODE", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     1", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$MAXACTVP", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "    64", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$PINSBASE", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 30", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$PLIMCHECK", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$PEXTMIN", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings "1.000000000000000E+20", vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings "1.000000000000000E+20", vbCrLf
    oStr.Add2Strings " 30", vbCrLf
    oStr.Add2Strings "1.000000000000000E+20", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$PEXTMAX", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings "-1.000000000000000E+20", vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings "-1.000000000000000E+20", vbCrLf
    oStr.Add2Strings " 30", vbCrLf
    oStr.Add2Strings "-1.000000000000000E+20", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$PLIMMIN", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$PLIMMAX", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings "12.0", vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings "9.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$UNITMODE", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$VISRETAIN", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     1", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$PLINEGEN", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$PSLTSCALE", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     1", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$TREEDEPTH", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "  3020", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$CMLSTYLE", vbCrLf
    oStr.Add2Strings "  2", vbCrLf
    oStr.Add2Strings "STANDARD", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$CMLJUST", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$CMLSCALE", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "1.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$PROXYGRAPHICS", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     1", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$MEASUREMENT", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$CELWEIGHT", vbCrLf
    oStr.Add2Strings "370", vbCrLf
    oStr.Add2Strings "    -1", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$ENDCAPS", vbCrLf
    oStr.Add2Strings "280", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$JOINSTYLE", vbCrLf
    oStr.Add2Strings "280", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$LWDISPLAY", vbCrLf
    oStr.Add2Strings "290", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$INSUNITS", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$HYPERLINKBASE", vbCrLf
    oStr.Add2Strings "  1", vbCrLf
    oStr.Add2Strings "", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$STYLESHEET", vbCrLf
    oStr.Add2Strings "  1", vbCrLf
    oStr.Add2Strings "", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$XEDIT", vbCrLf
    oStr.Add2Strings "290", vbCrLf
    oStr.Add2Strings "     1", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$CEPSNTYPE", vbCrLf
    oStr.Add2Strings "380", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$PSTYLEMODE", vbCrLf
    oStr.Add2Strings "290", vbCrLf
    oStr.Add2Strings "     1", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$EXTNAMES", vbCrLf
    oStr.Add2Strings "290", vbCrLf
    oStr.Add2Strings "     1", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$PSVPSCALE", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings "  9", vbCrLf
    oStr.Add2Strings "$OLESTARTUP", vbCrLf
    oStr.Add2Strings "290", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  0", vbCrLf
    oStr.Add2Strings "ENDSEC", vbCrLf
    DxfHeader = oStr.EndAdd
    Set oStr = Nothing
End Function

Function DxfTables() As String
    Dim oStr As cAddString
    Set oStr = New cAddString
    oStr.BeginAdd
    oStr.Add2Strings "  0", vbCrLf
    oStr.Add2Strings "SECTION", vbCrLf
    oStr.Add2Strings "2", vbCrLf
    oStr.Add2Strings "TABLES", vbCrLf
    oStr.AddString DxfTableVPORT()
    oStr.AddString DxfLineTypes()
    oStr.AddString DxfFonts()
    oStr.AddString DxfLayers()
    oStr.Add2Strings "  0", vbCrLf
    oStr.Add2Strings "ENDSEC", vbCrLf
    DxfTables = oStr.EndAdd
    Set oStr = Nothing
End Function

Function DxfTableVPORT() As String
    Dim dNum As Double
    Dim oStr As cAddString
    Set oStr = New cAddString
    oStr.BeginAdd
    oStr.Add2Strings "  0", vbCrLf
    oStr.Add2Strings "TABLE", vbCrLf
    oStr.Add2Strings "  2", vbCrLf
    oStr.Add2Strings "VPORT", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     2", vbCrLf
    oStr.Add2Strings "  0", vbCrLf
    oStr.Add2Strings "VPORT", vbCrLf
    oStr.Add2Strings "  2", vbCrLf
    oStr.Add2Strings "*Active", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings " 10", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 20", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 11", vbCrLf
    oStr.Add2Strings "1.0", vbCrLf
    oStr.Add2Strings " 21", vbCrLf
    oStr.Add2Strings "1.0", vbCrLf
    oStr.Add2Strings " 12", vbCrLf
    oStr.Add2Strings dFormat(EXTMIN_x + (EXTMAX_x - EXTMIN_x) / 2), vbCrLf
    oStr.Add2Strings " 22", vbCrLf
    oStr.Add2Strings dFormat(EXTMIN_y + (EXTMAX_y - EXTMIN_y) / 2), vbCrLf
    oStr.Add2Strings " 13", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 23", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 14", vbCrLf
    oStr.Add2Strings "1.0", vbCrLf
    oStr.Add2Strings " 24", vbCrLf
    oStr.Add2Strings "1.0", vbCrLf
    oStr.Add2Strings " 15", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 25", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 16", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 26", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 36", vbCrLf
    oStr.Add2Strings "1.0", vbCrLf
    oStr.Add2Strings " 17", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 27", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 37", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 40", vbCrLf
    If Abs(EXTMAX_x) > dNum Then dNum = Abs(EXTMAX_x)
    If Abs(EXTMAX_y) > dNum Then dNum = Abs(EXTMAX_y)
    If Abs(EXTMIN_x) > dNum Then dNum = Abs(EXTMIN_x)
    If Abs(EXTMIN_y) > dNum Then dNum = Abs(EXTMIN_y)
    oStr.Add2Strings dFormat(dNum), vbCrLf
    oStr.Add2Strings " 41", vbCrLf
    oStr.Add2Strings "2.0", vbCrLf
    oStr.Add2Strings " 42", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 43", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 44", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 50", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 51", vbCrLf
    oStr.Add2Strings "0.0", vbCrLf
    oStr.Add2Strings " 71", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings " 72", vbCrLf
    oStr.Add2Strings "  1000", vbCrLf
    oStr.Add2Strings " 73", vbCrLf
    oStr.Add2Strings "     1", vbCrLf
    oStr.Add2Strings " 74", vbCrLf
    oStr.Add2Strings "     2", vbCrLf '3=ÅìöÜíéóç UCS áîüíùí/2=áðüêñõøç
    oStr.Add2Strings " 75", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings " 76", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings " 77", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings " 78", vbCrLf
    oStr.Add2Strings "     0", vbCrLf
    oStr.Add2Strings "  0", vbCrLf
    oStr.Add2Strings "ENDTAB", vbCrLf
    DxfTableVPORT = oStr.EndAdd
    Set oStr = Nothing
End Function

Sub InitLineType()
    Dim MaxLineTypes As Long
    MaxLineTypes = 24
    ReDim sLineTypes(MaxLineTypes) As String
    sLineTypes(1) = "BORDER, 0.5, -0.25, 0.5, -0.25, 0, -0.25"
    sLineTypes(2) = "BORDER2, 0.25, -0.125, 0.25, -0.125, 0, -0.125"
    sLineTypes(3) = "BORDERX2, 1, -0.5, 1, -0.5, 0, -0.5"
    sLineTypes(4) = "CENTER, 1.25, -0.25, 0.25, -0.25"
    sLineTypes(5) = "CENTER2, 0.75, -0.125, 0.125, -0.125"
    sLineTypes(6) = "CENTERX2, 2.5, -0.5, 0.5, -0.5"
    sLineTypes(7) = "DASHDOT, 0.5, -0.25, 0, -0.25"
    sLineTypes(8) = "DASHDOT2, 0.25, -0.125, 0, -0.125"
    sLineTypes(9) = "DASHDOTX2, 1, -0.5, 0, -0.5"
    sLineTypes(10) = "DASHED, 0.5, -0.25"
    sLineTypes(11) = "DASHED2, 0.25, -0.125"
    sLineTypes(12) = "DASHEDX2, 1, -0.5"
    sLineTypes(13) = "DIVIDE, 0.5, -0.25, 0, -0.25, 0, -0.25"
    sLineTypes(14) = "DIVIDE2, 0.25, -0.125, 0, -0.125, 0, -0.125"
    sLineTypes(15) = "DIVIDEX2, 1, -0.5, 0, -0.5, 0, -0.5"
    sLineTypes(16) = "DOT, 0, -0.25"
    sLineTypes(17) = "DOT2, 0, -0.125"
    sLineTypes(18) = "DOTX2, 0, -0.5"
    sLineTypes(19) = "HIDDEN, 0.25, -0.125"
    sLineTypes(20) = "HIDDEN2, 0.125, -0.0625"
    sLineTypes(21) = "HIDDENX2, 0.5, -0.25"
    sLineTypes(22) = "PHANTOM, 1.25, -0.25, 0.25, -0.25, 0.25, -0.25"
    sLineTypes(23) = "PHANTOM2, 0.625, -0.125, 0.125, -0.125, 0.125, -0.125"
    sLineTypes(24) = "PHANTOMX2, 2.5, -0.5, 0.5, -0.5, 0.5, -0.5"
End Sub

Function DxfLineTypes() As String
    Dim MaxLineTypes As Long
    Dim i As Long
    Dim sString As String
    Dim sArr() As String
    Dim sDel As String
    Dim lTokMax As Long
    Dim lTokCnt As Long
    Dim dd As Double
    On Local Error Resume Next
    Dim oStr As cAddString
    Set oStr = New cAddString
    MaxLineTypes = UBound(sLineTypes)
    oStr.BeginAdd
    sDel = ","
    oStr.Add2Strings "  0", vbCrLf
    oStr.Add2Strings "TABLE", vbCrLf
    oStr.Add2Strings "  2", vbCrLf
    oStr.Add2Strings "LTYPE", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "    25", vbCrLf
    oStr.Add2Strings "  0", vbCrLf
    oStr.Add2Strings "LTYPE", vbCrLf
    oStr.Add2Strings "  2", vbCrLf
    oStr.Add2Strings "CONTINUOUS", vbCrLf
    oStr.Add2Strings "  70", vbCrLf
    oStr.Add2Strings "    0", vbCrLf
    oStr.Add2Strings "  3", vbCrLf
    oStr.Add2Strings "Solid line", vbCrLf
    oStr.Add2Strings "  72", vbCrLf
    oStr.Add2Strings "    65", vbCrLf
    oStr.Add2Strings "  73", vbCrLf
    oStr.Add2Strings "    0", vbCrLf
    oStr.Add2Strings "  40", vbCrLf
    oStr.Add2Strings " 0.0", vbCrLf
    For i = 1 To MaxLineTypes
        lTokMax = GetToken(sLineTypes(i), sArr(), sDel)
        If lTokMax > 1 Then
           oStr.Add2Strings "  0", vbCrLf
           oStr.Add2Strings "LTYPE", vbCrLf
           oStr.Add2Strings "  2", vbCrLf
           oStr.Add2Strings Trim$(sArr(1)), vbCrLf
           oStr.Add2Strings "  70", vbCrLf
           oStr.Add2Strings "     0", vbCrLf
           oStr.Add2Strings "  3", vbCrLf
           oStr.Add2Strings "", vbCrLf
           oStr.Add2Strings "  72", vbCrLf
           oStr.Add2Strings "     65", vbCrLf
           oStr.Add2Strings "  73", vbCrLf
           oStr.Add2Strings "     " & Format$(lTokMax - 1), vbCrLf
           dd = 0
           For lTokCnt = 1 To lTokMax - 1
               dd = dd + Abs(Val(Trim$(sArr(lTokCnt + 1))))
           Next lTokCnt
           oStr.Add2Strings "  40", vbCrLf
           oStr.Add2Strings dFormat(dd), vbCrLf
           For lTokCnt = 1 To lTokMax - 1
               oStr.Add2Strings "  49", vbCrLf
               oStr.Add2Strings dFormat(Val(Trim$(sArr(lTokCnt + 1)))), vbCrLf
           Next lTokCnt
        End If
    Next i
    oStr.Add2Strings "  0", vbCrLf
    oStr.Add2Strings "ENDTAB", vbCrLf
    DxfLineTypes = oStr.EndAdd
    Set oStr = Nothing
End Function

Public Function SetLineTypes(LineTypes() As String) As Boolean
    sLineTypes() = LineTypes()
    SetLineTypes = True
End Function

Public Function GetLineTypes(LineTypes() As String) As Long
    On Local Error Resume Next
    LineTypes() = sLineTypes()
    GetLineTypes = UBound(sLineTypes())
End Function

Function DxfFonts() As String
    Dim lCntFont As Long
    Dim lMaxFonts As Long
    Dim oStr As cAddString
    Set oStr = New cAddString
    On Local Error Resume Next
    lMaxFonts = UBound(sFonts)
    oStr.BeginAdd
    oStr.Add2Strings "  0", vbCrLf
    oStr.Add2Strings "TABLE", vbCrLf
    oStr.Add2Strings "  2", vbCrLf
    oStr.Add2Strings "STYLE", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "    2", vbCrLf
    For lCntFont = 1 To lMaxFonts
        If sFonts(lCntFont) <> "" Then
           oStr.Add2Strings "  0", vbCrLf
           oStr.Add2Strings "STYLE", vbCrLf
           oStr.Add2Strings "  2", vbCrLf
           oStr.Add2Strings "FONT" & Format$(lCntFont), vbCrLf
           oStr.Add2Strings "  70", vbCrLf
           oStr.Add2Strings "    64", vbCrLf
           oStr.Add2Strings "  40", vbCrLf
           oStr.Add2Strings " 0.0", vbCrLf
           oStr.Add2Strings "  41", vbCrLf
           oStr.Add2Strings " 1.0", vbCrLf
           oStr.Add2Strings "  50", vbCrLf
           oStr.Add2Strings " 0.0", vbCrLf
           oStr.Add2Strings "  71", vbCrLf
           oStr.Add2Strings "    0", vbCrLf
           oStr.Add2Strings "  42", vbCrLf
           oStr.Add2Strings " 0.2", vbCrLf
           oStr.Add2Strings "  3", vbCrLf
           oStr.Add2Strings sFonts(lCntFont), vbCrLf
           oStr.Add2Strings "  4", vbCrLf
           oStr.AddString vbCrLf
        End If
    Next lCntFont
    oStr.Add2Strings "  0", vbCrLf
    oStr.Add2Strings "ENDTAB", vbCrLf
    DxfFonts = oStr.EndAdd
    Set oStr = Nothing
End Function

Public Function SetFonts(Fonts() As String) As Boolean
    sFonts() = Fonts()
    SetFonts = True
End Function

Public Function GetFonts(Fonts() As String) As Long
    On Local Error Resume Next
    Fonts() = sFonts()
    GetFonts = UBound(sFonts())
End Function

Function DxfLayers() As String
    Dim lCntLayer As Long
    Dim lMaxLayers As Long
    Dim oStr As cAddString
    Set oStr = New cAddString
    On Local Error Resume Next
    lMaxLayers = UBound(sLayers)
    oStr.BeginAdd
    oStr.Add2Strings "  0", vbCrLf
    oStr.Add2Strings "TABLE", vbCrLf
    oStr.Add2Strings "  2", vbCrLf
    oStr.Add2Strings "LAYER", vbCrLf
    oStr.Add2Strings " 70", vbCrLf
    oStr.Add2Strings "    0", vbCrLf
    For lCntLayer = 1 To lMaxLayers
        If sLayers(lCntLayer) <> "" Then
           oStr.Add2Strings "  0", vbCrLf
           oStr.Add2Strings "LAYER", vbCrLf
           oStr.Add2Strings "  2", vbCrLf
           oStr.Add2Strings sLayers(lCntLayer), vbCrLf
           oStr.Add2Strings "  70", vbCrLf
           oStr.Add2Strings "   " & Format$(iLayerFlags(lCntLayer)), vbCrLf
           oStr.Add2Strings "  62", vbCrLf
           oStr.Add2Strings Format$(iLayerColorIndex(lCntLayer)), vbCrLf
           oStr.Add2Strings "  6", vbCrLf
           If sLayerLineTypeName(lCntLayer) = "" Then sLayerLineTypeName(lCntLayer) = "CONTINUOUS"
           oStr.Add2Strings sLayerLineTypeName(lCntLayer), vbCrLf
        End If
    Next lCntLayer
    oStr.Add2Strings "  0", vbCrLf
    oStr.Add2Strings "ENDTAB", vbCrLf
    DxfLayers = oStr.EndAdd
    Set oStr = Nothing
End Function

Public Function SetLayers(LayerNames() As String, LayerFlags() As Integer, LayerColorIndex() As Integer, LayerLineTypeName() As String) As Boolean
    sLayers() = LayerNames()
    iLayerFlags() = LayerFlags()
    iLayerColorIndex() = LayerColorIndex()
    sLayerLineTypeName() = LayerLineTypeName()
    SetLayers = True
End Function

Public Function GetLayers(LayerNames() As String, LayerFlags() As Integer, LayerColorIndex() As Integer, LayerLineTypeName() As String) As Long
    On Local Error Resume Next
    LayerNames() = sLayers()
    LayerFlags() = iLayerFlags()
    LayerColorIndex() = iLayerColorIndex()
    LayerLineTypeName() = sLayerLineTypeName()
    GetLayers = UBound(sLayers())
End Function

Private Sub Class_Initialize()
    LineTypeScale = 1
    Call InitLineType
    ReDim sLayers(1) As String
    ReDim iLayerFlags(1) As Integer
    ReDim iLayerColorIndex(1) As Integer
    ReDim sLayerLineTypeName(1) As String
    sLayers(1) = "0"
    iLayerFlags(1) = 0
    iLayerColorIndex(1) = 255
    sLayerLineTypeName(1) = "CONTINUOUS"
    ReDim sFonts(1) As String
    sFonts(1) = "Arial"
End Sub

Private Sub Class_Terminate()
    Erase sLineTypes
    Erase sFonts
End Sub

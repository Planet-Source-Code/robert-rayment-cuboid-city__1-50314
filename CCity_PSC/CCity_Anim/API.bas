Attribute VB_Name = "Module1"
'Module1: (API.bas)

Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


' Structures for StretchDIBits
Public Type BITMAPINFOHEADER ' 40 bytes
   biSize As Long
   biwidth As Long
   biheight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Public Type BITMAPINFO
   bmiH As BITMAPINFOHEADER
   Colors(0 To 255) As RGBQUAD
End Type
Public bm As BITMAPINFO

' For transferring drawing in an integer array to Form or PicBox
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal DesW As Long, ByVal DesH As Long, _
ByVal SrcXOffset As Long, ByVal SrcYOffset As Long, _
ByVal PICWW As Long, ByVal PICHH As Long, _
lpBits As Any, lpBitsInfo As BITMAPINFO, _
ByVal wUsage As Long, ByVal dwRop As Long) As Long

'Public Const DIB_PAL_COLORS = 1 '  color table in palette indices
Public Const DIB_RGB_COLORS = 0 '  color table in RGBs

'------------------------------------------------------------------------------
'Copy one array to another of same number of bytes

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
 (Destination As Any, Source As Any, ByVal Length As Long)

'------------------------------------------------------------------------------
'
'' For calling machine code
'Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
'(ByVal lpMCode As Long, _
'ByVal Long1 As Long, ByVal Long2 As Long, _
'ByVal Long3 As Long, ByVal Long4 As Long) As Long
''-----------------------------------------------------------------

Public Sub FillBMPStruc(ByVal bwidth As Long, ByVal bheight As Long)
   With bm.bmiH
      .biSize = 40
      .biwidth = bwidth
      .biheight = bheight
      .biPlanes = 1
      .biBitCount = 8            ' Sets up 8-bit colors
      .biCompression = 0
      .biSizeImage = Abs(bwidth) * Abs(bheight)
      .biXPelsPerMeter = 0
      .biYPelsPerMeter = 0
      .biClrUsed = 0
      .biClrImportant = 0
   End With
End Sub

Public Sub FillPalette()
ReDim Colors(0 To 255)
   With bm.Colors(1)
      .rgbRed = 200
      .rgbGreen = 255
      .rgbBlue = 200
      .rgbReserved = 0
   End With
   With bm.Colors(2)
      .rgbRed = 0
      .rgbGreen = 200
      .rgbBlue = 200
      .rgbReserved = 0
   End With
   With bm.Colors(3)
      .rgbRed = 0
      .rgbGreen = 120
      .rgbBlue = 200
      .rgbReserved = 0
   End With
   For i = 5 To 255
      With bm.Colors(i)
         .rgbRed = i
         .rgbGreen = i
         .rgbBlue = i
         .rgbReserved = 0
      End With
   Next i
End Sub

'Public Sub Loadmcode(InFile$, MCCode() As Byte)
''Load machine code into InCode() byte array
'On Error GoTo InFileErr
'If Dir$(InFile$) = "" Then
'   MsgBox (InFile$ & " missing")
'   DoEvents
'   Unload frmPlasTunnel
'   End
'End If
'Open InFile$ For Binary As #1
'MCSize& = LOF(1)
'If MCSize& = 0 Then
'InFileErr:
'   MsgBox (InFile$ & " missing")
'   DoEvents
'   Unload frmPlasTunnel
'   End
'End If
'ReDim MCCode(MCSize&)
'Get #1, , MCCode
'Close #1
'On Error GoTo 0
'End Sub

'Public Function zATan2(ByVal zy As Single, ByVal zx As Single)
'' Find angle Atan from -pi#/2 to +pi#/2
'' Public pi#
'If zx <> 0 Then
'   zATan2 = Atn(zy / zx)
'   If (zx < 0) Then
'      If (zy < 0) Then zATan2 = zATan2 - pi# Else zATan2 = zATan2 + pi#
'   End If
'Else  ' zx=0
'   If Abs(zy) > Abs(zx) Then   'Must be an overflow
'      If zy > 0 Then zATan2 = pi# / 2 Else zATan2 = -pi# / 2
'   Else
'      zATan2 = 0   'Must be an underflow
'   End If
'End If
'End Function
'

Attribute VB_Name = "RCPerspec"
' RCPerspec.bas

Option Explicit
Option Base 1

' Box coords generated from HLXYZ()
Private BoxX() As Long
Private BoxY() As Long
Private BoxZ() As Long

' Transformed Box to 2D
Private transx() As Long
Private transy() As Long
Private zDiv As Single
Private zNum As Single

Public eyeX As Long
Public eyeY As Long
Public eyeZ As Long

' For Bresline
Private ix As Long, iy As Long
Private idx As Long, idy As Long
Private jkstep As Long
Private incx As Long
Private id As Long
Private ainc As Long, binc As Long
Private JJ As Long, kk As Long

Private BN As Long
Private N As Long
Private NB As Long

'Public NumBlocksInRC() As Integer
'Public BlockNumsInRC() As Integer
'Public  BlockCounter As Long

Public Sub DRAW_RC_PERSPEC() ' NB=NumBlocksInRC()
Dim iRB As Long

   NB = NumBlocksInRC(R, C)
   If NB < 1 Then Exit Sub
   
   ReDim BoxX(8, NB)
   ReDim BoxY(8, NB)
   ReDim BoxZ(8, NB)
   ReDim transx(8, NB)
   ReDim transy(8, NB)
   
      'Find BlockNumbers in RC-square ( not nec in order!)
   ReDim BlockNumsInRC(NumBlocksInRC(R, C)) '''''''''''''
   BlockCounter = 0
   iRB = 0 '''''''''''''''
   For i = 1 To NumBlocks
      If NR(i) = R And NC(i) = C Then
         BlockCounter = BlockCounter + 1
         BlockNumsInRC(BlockCounter) = i
         If i = RedBlockNumber Then iRB = BlockCounter
         If BlockCounter = NumBlocksInRC(R, C) Then Exit For
      End If
   Next i
   
   ' FillBoxes
'  Y
'  |
'  | 8-------P7H
'  |/|       /|  Z
'  4--------3 | /
'  | |      | |/
'  | 5------|-6
'  |/       |/
' P1L-------2---- X
   For N = 1 To NB
      BN = BlockNumsInRC(N)
      ' Front face
      ' Pt 1
      BoxX(1, N) = LX(BN) '+ (NC(BN) - 1) * 256
      BoxY(1, N) = LY(BN)
      BoxZ(1, N) = LZ(BN) '+ (NR(BN) - 1) * 256
      ' Pt 2
      BoxX(2, N) = HX(BN) '+ (NC(BN) - 1) * 256
      BoxY(2, N) = BoxY(1, N)
      BoxZ(2, N) = BoxZ(1, N)
      ' Pt 3
      BoxX(3, N) = BoxX(2, N)
      BoxY(3, N) = HY(BN)
      BoxZ(3, N) = BoxZ(1, N)
      ' Pt 4
      BoxX(4, N) = BoxX(1, N)
      BoxY(4, N) = BoxY(3, N)
      BoxZ(4, N) = BoxZ(1, N)
'  Y
'  |
'  | 8-------P7H
'  |/|       /|  Z
'  4--------3 | /
'  | |      | |/
'  | 5------|-6
'  |/       |/
' P1L-------2---- X
      ' Back face
      ' Pt 5
      BoxX(5, N) = BoxX(1, N)
      BoxY(5, N) = LY(BN)
      BoxZ(5, N) = HZ(BN) '+ (NR(BN) - 1) * 256
      ' Pt 6
      BoxX(6, N) = BoxX(2, N)
      BoxY(6, N) = BoxY(5, N)
      BoxZ(6, N) = BoxZ(5, N)
      ' Pt 7
      BoxX(7, N) = BoxX(6, N)
      BoxY(7, N) = HY(BN)
      BoxZ(7, N) = BoxZ(5, N)
      ' Pt 8
      BoxX(8, N) = BoxX(5, N)
      BoxY(8, N) = BoxY(7, N)
      BoxZ(8, N) = BoxZ(5, N)
   Next N
   
   Transform
   
   DrawOnpicFace
   
 End Sub
' Redim Block(8,NB)
' Form 8 points from LX() etc for all blocks
' TransForm  ex,ey,ez
' Bresline to picFace

Private Sub Transform()
Dim PlaneZ As Long
   PlaneZ = 0
   'eyeX = 128 '256 + 128
   eyeY = 512 '650
   eyeZ = -512
   
   For N = 1 To NB
   For i = 1 To 8
      zDiv = CSng(BoxZ(i, N) - eyeZ)
      If zDiv > 0 Then
         zNum = CSng((BoxZ(i, N) - PlaneZ)) / zDiv
         transx(i, N) = CLng((eyeX - BoxX(i, N)) * zNum + 0.5) + BoxX(i, N)
         transy(i, N) = CLng((eyeY - BoxY(i, N)) * zNum + 0.5) + BoxY(i, N)
      End If
   Next i
   Next N
End Sub

Public Sub DrawOnpicFace()
Dim Cul As Long

   
   For N = 1 To NB
      
      ' Back plane
      ' Pt2 5-6
      
      'If transx(5, n) <> transx(6, n) Then
      'If transy(5, n) <> transy(6, n)  Then
      ' Quicker then drawing same pixel twice
      ' when added to every line drawn? -
      ' No unless all boxes of zero thicknes
      If BlockNumsInRC(N) = RedBlockNumber Then
         Cul = RGB(255, 0, 0)
      Else
         Cul = RGB(200, 200, 200)
      End If
      
      BresLine transx(5, N), transy(5, N), transx(6, N), transy(6, N), Cul
      ' Pt2 6-7
      BresLine transx(6, N), transy(6, N), transx(7, N), transy(7, N), Cul
      ' Pt2 7-8
      BresLine transx(7, N), transy(7, N), transx(8, N), transy(8, N), Cul
      ' Pt2 8-5
      BresLine transx(8, N), transy(8, N), transx(5, N), transy(5, N), Cul
      
'  Y
'  | 8-------P7H
'  |/|       /|  Z
'  4--------3 | /
'  | 5------|-6
'  |/       |/
' P1L-------2---- X
      ' Side lines
      ' Pt2 1-5
      If BlockNumsInRC(N) = RedBlockNumber Then
         Cul = RGB(255, 0, 0)
      Else
         Cul = RGB(180, 180, 180)
      End If
      
      BresLine transx(1, N), transy(1, N), transx(5, N), transy(5, N), Cul
      ' Pt2 2-6
      BresLine transx(2, N), transy(2, N), transx(6, N), transy(6, N), Cul
      ' Pt2 3-7
      BresLine transx(3, N), transy(3, N), transx(7, N), transy(7, N), Cul
      ' Pt2 4-8
      BresLine transx(4, N), transy(4, N), transx(8, N), transy(8, N), Cul

      ' Front plane
      ' Pt2 1-2
      'If BoxY(7, N) - BoxY(1, N) < 2 Then Cul = 3
      
      If BlockNumsInRC(N) = RedBlockNumber Then
         Cul = RGB(255, 0, 0)
      Else
         Cul = RGB(0, 0, 0)
      End If
      
      BresLine transx(1, N), transy(1, N), transx(2, N), transy(2, N), Cul
      ' Pt2 2-3
      BresLine transx(2, N), transy(2, N), transx(3, N), transy(3, N), Cul
      ' Pt2 3-4
      BresLine transx(3, N), transy(3, N), transx(4, N), transy(4, N), Cul
      ' Pt2 4-1
      BresLine transx(4, N), transy(4, N), transx(1, N), transy(1, N), Cul

   Next N
   Form1.picFace.Refresh
End Sub

Public Sub BresLine(ByVal ix1 As Long, ByVal iy1 As Long, ByVal ix2 As Long, ByVal iy2 As Long, ByVal Cul As Long)
'** Public BArray()
'** BASIC Bresenham Line for drawing into a 2D Public
'** Byte Array (BArray()) with a color index Cul (256 palette)

'** Plus clipping on 1->BArrayWidth, 1->BArrayHeight

'Dim ix As Long, iy As Long
'Dim idx As Long, idy As Long
'Dim jkstep As Long
'Dim incx As Long
'Dim id As Long
'Dim ainc As Long, binc As Long
'Dim jj As Long, kk As Long

   ' Reject lines outside BArray
   If ix1 > 0 Or ix2 > 0 Then
   If ix1 <= 256 Or ix2 <= 256 Then
   If iy1 > 0 Or iy2 > 0 Then
   If iy1 <= 512 Or iy2 <= 512 Then
      
      idx = Abs(ix2 - ix1)
      idy = Abs(iy2 - iy1)
      jkstep = 1
      incx = 1
      If idx < idy Then   '-- Steep slope
         
         If iy1 > iy2 Then jkstep = -1
         If ix2 < ix1 Then incx = -1
         id = 2 * idx - idy
         ainc = 2 * (idx - idy)   '-ve
         binc = 2 * idx
         JJ = iy1: kk = iy2: ix = ix1
      
         For iy = JJ To kk Step jkstep
            ' Reject any point outside BArray
            If ix > 0 Then
            If ix <= 256 Then
            If iy > 0 Then
            If iy <= 512 Then
               Form1.picFace.PSet (ix, iy), Cul
               'BArray(ix, iy) = Cul
            End If
            End If
            End If
            End If
            If id > 0 Then
               id = id + ainc
               ix = ix + incx
            Else
               id = id + binc
            End If
         Next iy
      
      Else                '-- Shallow slope
         
         If ix1 > ix2 Then jkstep = -1
         If iy2 < iy1 Then incx = -1
         id = 2 * idy - idx
         ainc = 2 * (idy - idx)   '-ve
         binc = 2 * idy
         JJ = ix1: kk = ix2: ix = iy1
      
         For iy = JJ To kk Step jkstep
            ' Reject any point outside BArray
            If iy > 0 Then
            If iy <= 256 Then
            If ix > 0 Then
            If ix <= 512 Then
               Form1.picFace.PSet (iy, ix), Cul
               'BArray(iy, ix) = Cul
            End If
            End If
            End If
            End If
            If id > 0 Then
               id = id + ainc
               ix = ix + incx
            Else
               id = id + binc
            End If
         Next iy
      
      End If
   
   End If
   End If
   End If
   End If

End Sub



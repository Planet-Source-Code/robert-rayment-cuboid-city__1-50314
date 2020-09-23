Attribute VB_Name = "Maze"
' Maze.bas

Option Explicit
Option Base 1

Dim WallRC() As Long
Dim VIS() As Long
Dim PathR() As Long
Dim PathC() As Long
Dim NumPaths As Long
Dim NumSidesLooked As Long
Dim Row As Long
Dim Col As Long
Dim RR As Long
Dim CC As Long
Dim pLX() As Integer, pLZ() As Integer, pLY() As Integer
Dim pHX() As Integer, pHZ() As Integer, pHY() As Integer

Public Sub MakeMaze()
Dim N As Long

' 14 x 14 maze
NumBlocks = 134  ' ENCLOSURE

ReDim pLX(NumBlocks), pLZ(NumBlocks), pLY(NumBlocks)
ReDim pHX(NumBlocks), pHZ(NumBlocks), pHY(NumBlocks)
ReDim LX(NumBlocks)
ReDim LY(NumBlocks)
ReDim LZ(NumBlocks)
ReDim HX(NumBlocks)
ReDim HY(NumBlocks)
ReDim HZ(NumBlocks)
ReDim NR(NumBlocks)
ReDim NC(NumBlocks)
' Num of sub-blocks in RC-square
ReDim NumBlocksInRC(-2 To 21, -5 To 12)

LoadEnclosure

' Transfer to LX(),,,,
For i = 1 To NumBlocks
   LX(i) = pLX(i)
   LY(i) = pLY(i)
   LZ(i) = pLZ(i)
   HX(i) = pHX(i)
   HY(i) = pHY(i)
   HZ(i) = pHZ(i)
   'NR(NumBlocks) = R ' Already transferred by LoadEnclosure
   'NC(NumBlocks) = C
   NumBlocksInRC(NR(i), NC(i)) = NumBlocksInRC(NR(i), NC(i)) + 1
Next i


ReDim VIS(0 To 15, 0 To 15)
ReDim PathR(1)
ReDim PathC(1)
ReDim WallRC(14, 14)
   ' Make outside boundary visited
   For Col = 0 To 15
      VIS(0, Col) = 1
      VIS(15, Col) = 1
   Next Col
   For Row = 0 To 15
      VIS(Row, 0) = 1
      VIS(Row, 15) = 1
   Next Row
   ' Marks in all walls
   For i = 1 To 14
   For j = 1 To 14
      WallRC(i, j) = 15
   Next j
   Next i
   
   Randomize
   NumPaths = 0
   ' Start at random Row, Col in 14x14
   Row = Int((14 - 1 + 1) * Rnd + 1)
   Col = Int((14 - 1 + 1) * Rnd + 1)
   ' Start of first path
   MarkPath
   
   N = 0

   Do
      If VIS(Row, Col) <> 1 Then
         VIS(Row, Col) = 1
         N = N + 1
         NumSidesLooked = 1
         ' Look at 4 sides randomly
         i = Int((4 - 1 + 1) * Rnd + 1)
         j = i + 1: If i = 4 Then j = 1

L1:      ' Check adjacent square according to i value
         If i = 1 And VIS(Row - 1, Col) <> 1 Then
            WallRC(Row, Col) = WallRC(Row, Col) - 1 ' Lower row
            Row = Row - 1
            WallRC(Row, Col) = WallRC(Row, Col) - 4 ' Top row
            MarkPath
         ElseIf i = 2 And VIS(Row, Col + 1) <> 1 Then
            WallRC(Row, Col) = WallRC(Row, Col) - 2 ' Right col
            Col = Col + 1
            WallRC(Row, Col) = WallRC(Row, Col) - 8 ' Left col
            MarkPath
         ElseIf i = 3 And VIS(Row + 1, Col) <> 1 Then
            WallRC(Row, Col) = WallRC(Row, Col) - 4 ' Top row
            Row = Row + 1
            WallRC(Row, Col) = WallRC(Row, Col) - 1 ' Low row
            MarkPath
         
         ElseIf i = 4 And VIS(Row, Col - 1) <> 1 Then
            WallRC(Row, Col) = WallRC(Row, Col) - 8 ' Left col
            Col = Col - 1
            WallRC(Row, Col) = WallRC(Row, Col) - 2 ' Right col
            MarkPath
                     
         Else
            ' Look at another side
            NumSidesLooked = NumSidesLooked + 1
            
            If NumSidesLooked < 5 Then
               i = i + 1: If i = 5 Then i = 1
               j = i + 1: If i = 4 Then j = 1
               GoTo L1  ' Check adjacent square according to i value
            ElseIf NumPaths > 1 Then ' Backtrack on path
                  NumPaths = NumPaths - 1
                  ReDim Preserve PathR(NumPaths), PathC(NumPaths)
                  Row = PathR(NumPaths)
                  Col = PathC(NumPaths)
                  NumSidesLooked = 1
                  ' Look at 4 sides randomly
                  i = Int((4 - 1 + 1) * Rnd + 1)
                  j = i + 1: If i = 4 Then j = 1
                  GoTo L1  ' Check adjacent square according to i value
            Else
               RR = RR
               ' Run thru to Loop
            End If
         
         End If
      
      End If   'If VIS(Row, Col) <> 1 Then
   
   Loop Until N = 14 * 14
   
   ' From WallRC() create walls
   For RR = 1 To 14
   For CC = 1 To 14
      Select Case WallRC(RR, CC)
      Case 9, 11, 13 ' Left wall & Bottom wall
         If (RR <> 8 Or CC <> 1) Then  ' Skip maze entry point
            NumBlocks = NumBlocks + 1
            ReDimPreserveLH
            LX(NumBlocks) = 1    ' Left wall
            LY(NumBlocks) = 1
            LZ(NumBlocks) = 1
            HX(NumBlocks) = 16
            HY(NumBlocks) = 64
            HZ(NumBlocks) = 256
            NR(NumBlocks) = RR + 5
            NC(NumBlocks) = CC - 4
            NumBlocksInRC(NR(NumBlocks), NC(NumBlocks)) = NumBlocksInRC(NR(NumBlocks), NC(NumBlocks)) + 1
         End If
         NumBlocks = NumBlocks + 1
         ReDimPreserveLH
         LX(NumBlocks) = 1    ' Bottom wall
         LY(NumBlocks) = 1
         LZ(NumBlocks) = 1
         HX(NumBlocks) = 256
         HY(NumBlocks) = 64
         HZ(NumBlocks) = 16
         NR(NumBlocks) = RR + 5
         NC(NumBlocks) = CC - 4
         NumBlocksInRC(NR(NumBlocks), NC(NumBlocks)) = NumBlocksInRC(NR(NumBlocks), NC(NumBlocks)) + 1
      Case 1, 3, 7, 5   ' Bottom wall
         NumBlocks = NumBlocks + 1
         ReDimPreserveLH
         LX(NumBlocks) = 1
         LY(NumBlocks) = 1
         LZ(NumBlocks) = 1
         HX(NumBlocks) = 256
         HY(NumBlocks) = 64
         HZ(NumBlocks) = 16
         NR(NumBlocks) = RR + 5
         NC(NumBlocks) = CC - 4
         NumBlocksInRC(NR(NumBlocks), NC(NumBlocks)) = NumBlocksInRC(NR(NumBlocks), NC(NumBlocks)) + 1
      Case 8, 10, 12, 14   ' Left wall
         If (RR <> 8 Or CC <> 1) Then  ' Skip maze entry point
            NumBlocks = NumBlocks + 1
            ReDimPreserveLH
            LX(NumBlocks) = 1
            LY(NumBlocks) = 1
            LZ(NumBlocks) = 1
            HX(NumBlocks) = 16
            HY(NumBlocks) = 64
            HZ(NumBlocks) = 256
            NR(NumBlocks) = RR + 5
            NC(NumBlocks) = CC - 4
            NumBlocksInRC(NR(NumBlocks), NC(NumBlocks)) = NumBlocksInRC(NR(NumBlocks), NC(NumBlocks)) + 1
         End If
      Case Else
         ' 2,4,6
      End Select
   Next CC
   Next RR
   
   
   R = RR - 1 + 5
   C = CC - 1 - 4
   Form1.LabNumBlocks = "NumBlocks =" & Str$(NumBlocks)
   RedBlockNumber = NumBlocks
   Form1.LabBlockNum = "BlockNum =" & Str$(RedBlockNumber)
   Form1.LabBlocksInRC = "Blocks in RC =" & Str$(NumBlocksInRC(R, C))
   REDRAW_PLAN Form1.picPLAN
   DRAW_Faces Form1.picFace

   ' Show on RC grid
   For j = -2 To 21  ' rows
   For i = -5 To 12  ' cols
      iy = GridStep * (21 - j)
      ix = GridStep * (i + 5)
      If NumBlocksInRC(j, i) > 0 Then
         Form1.picRC.Line (ix + 1, iy + 1)-(ix + GridStep - 1, iy + GridStep - 1), vbWhite, BF
      Else
         Form1.picRC.Line (ix + 1, iy + 1)-(ix + GridStep - 1, iy + GridStep - 1), GridBackColor, BF
      End If
   Next i
   Next j
   
   ' Highlight RC-square with last block
   j = RR - 1 + 5
   i = CC - 1 - 4
   iy = GridStep * (21 - j)
   ix = GridStep * (i + 5)
   If NumBlocksInRC(j, i) > 0 Then
      Form1.picRC.Line (ix + 1, iy + 1)-(ix + GridStep - 1, iy + GridStep - 1), SelectCul, BF
   End If

Erase WallRC, VIS, PathR, PathC
Erase pLX, pLY, pLZ, pHX, pHY, pHZ
End Sub

Private Sub MarkPath()
   NumPaths = NumPaths + 1
   ReDim Preserve PathR(NumPaths), PathC(NumPaths)
   PathR(NumPaths) = Row
   PathC(NumPaths) = Col
End Sub

Private Sub ReDimPreserveLH()
   ReDim Preserve LX(NumBlocks)
   ReDim Preserve LZ(NumBlocks)
   ReDim Preserve LY(NumBlocks)
   ReDim Preserve HX(NumBlocks)
   ReDim Preserve HZ(NumBlocks)
   ReDim Preserve HY(NumBlocks)
   ReDim Preserve NR(NumBlocks)
   ReDim Preserve NC(NumBlocks)
End Sub

Private Sub LoadEnclosure()

'ENCLarge.ccs  from Special save
'NoBoxes = 134 '128
pLX(1) = 1: pLY(1) = 1: pLZ(1) = 240: pHX(1) = 256: pHY(1) = 64: pHZ(1) = 256
NR(1) = 21: NC(1) = -5
pLX(2) = 1: pLY(2) = 1: pLZ(2) = 1: pHX(2) = 16: pHY(2) = 64: pHZ(2) = 240
NR(2) = 21: NC(2) = -5
pLX(3) = 1: pLY(3) = 1: pLZ(3) = 240: pHX(3) = 256: pHY(3) = 64: pHZ(3) = 256
NR(3) = 21: NC(3) = -4
pLX(4) = 1: pLY(4) = 1: pLZ(4) = 240: pHX(4) = 256: pHY(4) = 64: pHZ(4) = 256
NR(4) = 21: NC(4) = -3
pLX(5) = 1: pLY(5) = 1: pLZ(5) = 240: pHX(5) = 256: pHY(5) = 64: pHZ(5) = 256
NR(5) = 21: NC(5) = -2
pLX(6) = 1: pLY(6) = 1: pLZ(6) = 240: pHX(6) = 256: pHY(6) = 64: pHZ(6) = 256
NR(6) = 21: NC(6) = -1
pLX(7) = 1: pLY(7) = 1: pLZ(7) = 240: pHX(7) = 256: pHY(7) = 64: pHZ(7) = 256
NR(7) = 21: NC(7) = 0
pLX(8) = 1: pLY(8) = 1: pLZ(8) = 240: pHX(8) = 256: pHY(8) = 64: pHZ(8) = 256
NR(8) = 21: NC(8) = 1
pLX(9) = 1: pLY(9) = 1: pLZ(9) = 240: pHX(9) = 256: pHY(9) = 64: pHZ(9) = 256
NR(9) = 21: NC(9) = 2
pLX(10) = 1: pLY(10) = 1: pLZ(10) = 240: pHX(10) = 256: pHY(10) = 64: pHZ(10) = 256
NR(10) = 21: NC(10) = 3
pLX(11) = 1: pLY(11) = 1: pLZ(11) = 240: pHX(11) = 256: pHY(11) = 64: pHZ(11) = 256
NR(11) = 21: NC(11) = 4
pLX(12) = 1: pLY(12) = 1: pLZ(12) = 240: pHX(12) = 256: pHY(12) = 64: pHZ(12) = 256
NR(12) = 21: NC(12) = 5
pLX(13) = 1: pLY(13) = 1: pLZ(13) = 240: pHX(13) = 256: pHY(13) = 64: pHZ(13) = 256
NR(13) = 21: NC(13) = 6
pLX(14) = 1: pLY(14) = 1: pLZ(14) = 240: pHX(14) = 256: pHY(14) = 64: pHZ(14) = 256
NR(14) = 21: NC(14) = 7
pLX(15) = 1: pLY(15) = 1: pLZ(15) = 240: pHX(15) = 256: pHY(15) = 64: pHZ(15) = 256
NR(15) = 21: NC(15) = 8
pLX(16) = 1: pLY(16) = 1: pLZ(16) = 240: pHX(16) = 256: pHY(16) = 64: pHZ(16) = 256
NR(16) = 21: NC(16) = 9
pLX(17) = 1: pLY(17) = 1: pLZ(17) = 240: pHX(17) = 256: pHY(17) = 64: pHZ(17) = 256
NR(17) = 21: NC(17) = 10
pLX(18) = 1: pLY(18) = 1: pLZ(18) = 240: pHX(18) = 256: pHY(18) = 64: pHZ(18) = 256
NR(18) = 21: NC(18) = 11
pLX(19) = 1: pLY(19) = 1: pLZ(19) = 240: pHX(19) = 256: pHY(19) = 64: pHZ(19) = 256
NR(19) = 21: NC(19) = 12
pLX(20) = 240: pLY(20) = 1: pLZ(20) = 1: pHX(20) = 256: pHY(20) = 64: pHZ(20) = 240
NR(20) = 21: NC(20) = 12
pLX(21) = 1: pLY(21) = 1: pLZ(21) = 1: pHX(21) = 16: pHY(21) = 64: pHZ(21) = 256
NR(21) = 20: NC(21) = -5
pLX(22) = 240: pLY(22) = 1: pLZ(22) = 1: pHX(22) = 256: pHY(22) = 64: pHZ(22) = 256
NR(22) = 20: NC(22) = 12
pLX(23) = 1: pLY(23) = 1: pLZ(23) = 240: pHX(23) = 256: pHY(23) = 64: pHZ(23) = 256
NR(23) = 19: NC(23) = -1
pLX(24) = 1: pLY(24) = 1: pLZ(24) = 240: pHX(24) = 256: pHY(24) = 64: pHZ(24) = 256
NR(24) = 19: NC(24) = -2
pLX(25) = 1: pLY(25) = 1: pLZ(25) = 240: pHX(25) = 256: pHY(25) = 64: pHZ(25) = 256
NR(25) = 19: NC(25) = 1
pLX(26) = 1: pLY(26) = 1: pLZ(26) = 240: pHX(26) = 256: pHY(26) = 64: pHZ(26) = 256
NR(26) = 19: NC(26) = 0
pLX(27) = 1: pLY(27) = 1: pLZ(27) = 240: pHX(27) = 256: pHY(27) = 64: pHZ(27) = 256
NR(27) = 19: NC(27) = 10
pLX(28) = 240: pLY(28) = 1: pLZ(28) = 1: pHX(28) = 256: pHY(28) = 64: pHZ(28) = 240
NR(28) = 19: NC(28) = 10
pLX(29) = 1: pLY(29) = 1: pLZ(29) = 1: pHX(29) = 16: pHY(29) = 64: pHZ(29) = 240
NR(29) = 19: NC(29) = -3
pLX(30) = 1: pLY(30) = 1: pLZ(30) = 240: pHX(30) = 256: pHY(30) = 64: pHZ(30) = 256
NR(30) = 19: NC(30) = -3
pLX(31) = 1: pLY(31) = 1: pLZ(31) = 240: pHX(31) = 256: pHY(31) = 64: pHZ(31) = 256
NR(31) = 19: NC(31) = 7
pLX(32) = 1: pLY(32) = 1: pLZ(32) = 240: pHX(32) = 256: pHY(32) = 64: pHZ(32) = 256
NR(32) = 19: NC(32) = 6
pLX(33) = 1: pLY(33) = 1: pLZ(33) = 240: pHX(33) = 256: pHY(33) = 64: pHZ(33) = 256
NR(33) = 19: NC(33) = 9
pLX(34) = 1: pLY(34) = 1: pLZ(34) = 240: pHX(34) = 256: pHY(34) = 64: pHZ(34) = 256
NR(34) = 19: NC(34) = 8
pLX(35) = 1: pLY(35) = 1: pLZ(35) = 240: pHX(35) = 256: pHY(35) = 64: pHZ(35) = 256
NR(35) = 19: NC(35) = 3
pLX(36) = 1: pLY(36) = 1: pLZ(36) = 240: pHX(36) = 256: pHY(36) = 64: pHZ(36) = 256
NR(36) = 19: NC(36) = 2
pLX(37) = 1: pLY(37) = 1: pLZ(37) = 240: pHX(37) = 256: pHY(37) = 64: pHZ(37) = 256
NR(37) = 19: NC(37) = 5
pLX(38) = 1: pLY(38) = 1: pLZ(38) = 240: pHX(38) = 256: pHY(38) = 64: pHZ(38) = 256
NR(38) = 19: NC(38) = 4
pLX(39) = 240: pLY(39) = 1: pLZ(39) = 1: pHX(39) = 256: pHY(39) = 64: pHZ(39) = 256
NR(39) = 19: NC(39) = 12
pLX(40) = 1: pLY(40) = 1: pLZ(40) = 1: pHX(40) = 16: pHY(40) = 64: pHZ(40) = 256
NR(40) = 19: NC(40) = -5
pLX(41) = 1: pLY(41) = 1: pLZ(41) = 1: pHX(41) = 16: pHY(41) = 64: pHZ(41) = 256
NR(41) = 18: NC(41) = -3
pLX(42) = 240: pLY(42) = 1: pLZ(42) = 1: pHX(42) = 256: pHY(42) = 64: pHZ(42) = 256
NR(42) = 18: NC(42) = 10
pLX(43) = 1: pLY(43) = 1: pLZ(43) = 1: pHX(43) = 16: pHY(43) = 64: pHZ(43) = 256
NR(43) = 18: NC(43) = -5
pLX(44) = 240: pLY(44) = 1: pLZ(44) = 1: pHX(44) = 256: pHY(44) = 64: pHZ(44) = 256
NR(44) = 18: NC(44) = 12
pLX(45) = 1: pLY(45) = 1: pLZ(45) = 1: pHX(45) = 16: pHY(45) = 64: pHZ(45) = 256
NR(45) = 17: NC(45) = -5
pLX(46) = 240: pLY(46) = 1: pLZ(46) = 1: pHX(46) = 256: pHY(46) = 64: pHZ(46) = 256
NR(46) = 17: NC(46) = 12
pLX(47) = 240: pLY(47) = 1: pLZ(47) = 1: pHX(47) = 256: pHY(47) = 64: pHZ(47) = 256
NR(47) = 17: NC(47) = 10
pLX(48) = 1: pLY(48) = 1: pLZ(48) = 1: pHX(48) = 16: pHY(48) = 64: pHZ(48) = 256
NR(48) = 17: NC(48) = -3
pLX(49) = 240: pLY(49) = 1: pLZ(49) = 1: pHX(49) = 256: pHY(49) = 64: pHZ(49) = 256
NR(49) = 16: NC(49) = 10
pLX(50) = 1: pLY(50) = 1: pLZ(50) = 1: pHX(50) = 16: pHY(50) = 64: pHZ(50) = 256
NR(50) = 16: NC(50) = -5
pLX(51) = 240: pLY(51) = 1: pLZ(51) = 1: pHX(51) = 256: pHY(51) = 64: pHZ(51) = 256
NR(51) = 16: NC(51) = 12
pLX(52) = 1: pLY(52) = 1: pLZ(52) = 1: pHX(52) = 16: pHY(52) = 64: pHZ(52) = 256
NR(52) = 16: NC(52) = -3
pLX(53) = 1: pLY(53) = 1: pLZ(53) = 1: pHX(53) = 16: pHY(53) = 64: pHZ(53) = 256
NR(53) = 15: NC(53) = -3
pLX(54) = 240: pLY(54) = 1: pLZ(54) = 1: pHX(54) = 256: pHY(54) = 64: pHZ(54) = 256
NR(54) = 15: NC(54) = 12
pLX(55) = 1: pLY(55) = 1: pLZ(55) = 1: pHX(55) = 16: pHY(55) = 64: pHZ(55) = 256
NR(55) = 15: NC(55) = -5
pLX(56) = 240: pLY(56) = 1: pLZ(56) = 1: pHX(56) = 256: pHY(56) = 64: pHZ(56) = 256
NR(56) = 15: NC(56) = 10
pLX(57) = 1: pLY(57) = 1: pLZ(57) = 1: pHX(57) = 16: pHY(57) = 64: pHZ(57) = 256
NR(57) = 14: NC(57) = -3
pLX(58) = 240: pLY(58) = 1: pLZ(58) = 1: pHX(58) = 256: pHY(58) = 64: pHZ(58) = 256
NR(58) = 14: NC(58) = 10
pLX(59) = 1: pLY(59) = 1: pLZ(59) = 1: pHX(59) = 16: pHY(59) = 64: pHZ(59) = 256
NR(59) = 14: NC(59) = -5
pLX(60) = 240: pLY(60) = 1: pLZ(60) = 1: pHX(60) = 256: pHY(60) = 64: pHZ(60) = 256
NR(60) = 14: NC(60) = 12
pLX(61) = 1: pLY(61) = 1: pLZ(61) = 1: pHX(61) = 16: pHY(61) = 64: pHZ(61) = 256
NR(61) = 13: NC(61) = -5
pLX(62) = 240: pLY(62) = 1: pLZ(62) = 1: pHX(62) = 256: pHY(62) = 64: pHZ(62) = 256
NR(62) = 13: NC(62) = 10
pLX(63) = 240: pLY(63) = 1: pLZ(63) = 1: pHX(63) = 256: pHY(63) = 64: pHZ(63) = 256
NR(63) = 13: NC(63) = 12
pLX(64) = 1: pLY(64) = 1: pLZ(64) = 1: pHX(64) = 16: pHY(64) = 64: pHZ(64) = 256
NR(64) = 12: NC(64) = -5
pLX(65) = 1: pLY(65) = 1: pLZ(65) = 1: pHX(65) = 16: pHY(65) = 64: pHZ(65) = 256
NR(65) = 12: NC(65) = -3
pLX(66) = 240: pLY(66) = 1: pLZ(66) = 1: pHX(66) = 256: pHY(66) = 64: pHZ(66) = 256
NR(66) = 12: NC(66) = 10
pLX(67) = 240: pLY(67) = 1: pLZ(67) = 1: pHX(67) = 256: pHY(67) = 64: pHZ(67) = 256
NR(67) = 12: NC(67) = 12
pLX(68) = 1: pLY(68) = 1: pLZ(68) = 1: pHX(68) = 16: pHY(68) = 64: pHZ(68) = 256
NR(68) = 11: NC(68) = -5
pLX(69) = 1: pLY(69) = 1: pLZ(69) = 1: pHX(69) = 16: pHY(69) = 64: pHZ(69) = 256
NR(69) = 11: NC(69) = -3
pLX(70) = 240: pLY(70) = 1: pLZ(70) = 1: pHX(70) = 256: pHY(70) = 64: pHZ(70) = 256
NR(70) = 11: NC(70) = 12
pLX(71) = 240: pLY(71) = 1: pLZ(71) = 1: pHX(71) = 256: pHY(71) = 64: pHZ(71) = 256
NR(71) = 11: NC(71) = 10
pLX(72) = 1: pLY(72) = 1: pLZ(72) = 1: pHX(72) = 16: pHY(72) = 64: pHZ(72) = 256
NR(72) = 10: NC(72) = -3
pLX(73) = 240: pLY(73) = 1: pLZ(73) = 1: pHX(73) = 256: pHY(73) = 64: pHZ(73) = 256
NR(73) = 10: NC(73) = 10
pLX(74) = 240: pLY(74) = 1: pLZ(74) = 1: pHX(74) = 256: pHY(74) = 64: pHZ(74) = 256
NR(74) = 10: NC(74) = 12
pLX(75) = 1: pLY(75) = 1: pLZ(75) = 1: pHX(75) = 16: pHY(75) = 64: pHZ(75) = 256
NR(75) = 10: NC(75) = -5
pLX(76) = 1: pLY(76) = 1: pLZ(76) = 1: pHX(76) = 16: pHY(76) = 64: pHZ(76) = 256
NR(76) = 9: NC(76) = -5
pLX(77) = 1: pLY(77) = 1: pLZ(77) = 1: pHX(77) = 16: pHY(77) = 64: pHZ(77) = 256
NR(77) = 9: NC(77) = -3
pLX(78) = 240: pLY(78) = 1: pLZ(78) = 1: pHX(78) = 256: pHY(78) = 64: pHZ(78) = 256
NR(78) = 9: NC(78) = 12
pLX(79) = 240: pLY(79) = 1: pLZ(79) = 1: pHX(79) = 256: pHY(79) = 64: pHZ(79) = 256
NR(79) = 9: NC(79) = 10
pLX(80) = 240: pLY(80) = 1: pLZ(80) = 1: pHX(80) = 256: pHY(80) = 64: pHZ(80) = 256
NR(80) = 8: NC(80) = 10
pLX(81) = 1: pLY(81) = 1: pLZ(81) = 1: pHX(81) = 16: pHY(81) = 64: pHZ(81) = 256
NR(81) = 8: NC(81) = -3
pLX(82) = 1: pLY(82) = 1: pLZ(82) = 1: pHX(82) = 16: pHY(82) = 64: pHZ(82) = 256
NR(82) = 8: NC(82) = -5
pLX(83) = 240: pLY(83) = 1: pLZ(83) = 1: pHX(83) = 256: pHY(83) = 64: pHZ(83) = 256
NR(83) = 8: NC(83) = 12
pLX(84) = 1: pLY(84) = 1: pLZ(84) = 1: pHX(84) = 16: pHY(84) = 64: pHZ(84) = 256
NR(84) = 7: NC(84) = -5
pLX(85) = 240: pLY(85) = 1: pLZ(85) = 1: pHX(85) = 256: pHY(85) = 64: pHZ(85) = 256
NR(85) = 7: NC(85) = 12
pLX(86) = 1: pLY(86) = 1: pLZ(86) = 1: pHX(86) = 16: pHY(86) = 64: pHZ(86) = 256
NR(86) = 7: NC(86) = -3
pLX(87) = 240: pLY(87) = 1: pLZ(87) = 1: pHX(87) = 256: pHY(87) = 64: pHZ(87) = 256
NR(87) = 7: NC(87) = 10
pLX(88) = 1: pLY(88) = 1: pLZ(88) = 1: pHX(88) = 256: pHY(88) = 64: pHZ(88) = 16
NR(88) = 6: NC(88) = 9
pLX(89) = 1: pLY(89) = 1: pLZ(89) = 1: pHX(89) = 16: pHY(89) = 64: pHZ(89) = 256
NR(89) = 6: NC(89) = -5
pLX(90) = 1: pLY(90) = 1: pLZ(90) = 1: pHX(90) = 256: pHY(90) = 64: pHZ(90) = 16
NR(90) = 6: NC(90) = 10
pLX(91) = 240: pLY(91) = 1: pLZ(91) = 16: pHX(91) = 256: pHY(91) = 64: pHZ(91) = 256
NR(91) = 6: NC(91) = 10
pLX(92) = 240: pLY(92) = 1: pLZ(92) = 1: pHX(92) = 256: pHY(92) = 64: pHZ(92) = 256
NR(92) = 6: NC(92) = 12
pLX(93) = 1: pLY(93) = 1: pLZ(93) = 1: pHX(93) = 256: pHY(93) = 64: pHZ(93) = 16
NR(93) = 6: NC(93) = 6
pLX(94) = 1: pLY(94) = 1: pLZ(94) = 1: pHX(94) = 256: pHY(94) = 64: pHZ(94) = 16
NR(94) = 6: NC(94) = -2
pLX(95) = 1: pLY(95) = 1: pLZ(95) = 16: pHX(95) = 16: pHY(95) = 64: pHZ(95) = 256
NR(95) = 6: NC(95) = -3
pLX(96) = 1: pLY(96) = 1: pLZ(96) = 1: pHX(96) = 256: pHY(96) = 64: pHZ(96) = 16
NR(96) = 6: NC(96) = -3
pLX(97) = 1: pLY(97) = 1: pLZ(97) = 1: pHX(97) = 256: pHY(97) = 64: pHZ(97) = 16
NR(97) = 6: NC(97) = 1
pLX(98) = 1: pLY(98) = 1: pLZ(98) = 1: pHX(98) = 256: pHY(98) = 64: pHZ(98) = 16
NR(98) = 6: NC(98) = 7
pLX(99) = 1: pLY(99) = 1: pLZ(99) = 1: pHX(99) = 256: pHY(99) = 64: pHZ(99) = 16
NR(99) = 6: NC(99) = -1
pLX(100) = 1: pLY(100) = 1: pLZ(100) = 1: pHX(100) = 256: pHY(100) = 64: pHZ(100) = 16
NR(100) = 6: NC(100) = 5
pLX(101) = 1: pLY(101) = 1: pLZ(101) = 1: pHX(101) = 256: pHY(101) = 64: pHZ(101) = 16
NR(101) = 6: NC(101) = 0
pLX(102) = 1: pLY(102) = 1: pLZ(102) = 1: pHX(102) = 256: pHY(102) = 64: pHZ(102) = 16
NR(102) = 6: NC(102) = 3
pLX(103) = 1: pLY(103) = 1: pLZ(103) = 1: pHX(103) = 256: pHY(103) = 64: pHZ(103) = 16
NR(103) = 6: NC(103) = 4
pLX(104) = 1: pLY(104) = 1: pLZ(104) = 1: pHX(104) = 256: pHY(104) = 64: pHZ(104) = 16
NR(104) = 6: NC(104) = 2
pLX(105) = 1: pLY(105) = 1: pLZ(105) = 1: pHX(105) = 256: pHY(105) = 64: pHZ(105) = 16
NR(105) = 6: NC(105) = 8
pLX(106) = 1: pLY(106) = 1: pLZ(106) = 1: pHX(106) = 16: pHY(106) = 64: pHZ(106) = 256
NR(106) = 5: NC(106) = -5
pLX(107) = 240: pLY(107) = 1: pLZ(107) = 1: pHX(107) = 256: pHY(107) = 64: pHZ(107) = 256
NR(107) = 5: NC(107) = 12
pLX(108) = 240: pLY(108) = 1: pLZ(108) = 1: pHX(108) = 256: pHY(108) = 64: pHZ(108) = 256
NR(108) = 5: NC(108) = 7
pLX(109) = 1: pLY(109) = 1: pLZ(109) = 1: pHX(109) = 256: pHY(109) = 64: pHZ(109) = 16
NR(109) = 4: NC(109) = 0
pLX(110) = 1: pLY(110) = 1: pLZ(110) = 1: pHX(110) = 256: pHY(110) = 64: pHZ(110) = 16
NR(110) = 4: NC(110) = 7
pLX(111) = 240: pLY(111) = 1: pLZ(111) = 16: pHX(111) = 256: pHY(111) = 64: pHZ(111) = 256
NR(111) = 4: NC(111) = 7
pLX(112) = 1: pLY(112) = 1: pLZ(112) = 1: pHX(112) = 256: pHY(112) = 64: pHZ(112) = 16
NR(112) = 4: NC(112) = 1
pLX(113) = 1: pLY(113) = 1: pLZ(113) = 1: pHX(113) = 256: pHY(113) = 64: pHZ(113) = 16
NR(113) = 4: NC(113) = -2
pLX(114) = 1: pLY(114) = 1: pLZ(114) = 1: pHX(114) = 256: pHY(114) = 64: pHZ(114) = 16
NR(114) = 4: NC(114) = -3
pLX(115) = 1: pLY(115) = 1: pLZ(115) = 1: pHX(115) = 256: pHY(115) = 64: pHZ(115) = 16
NR(115) = 4: NC(115) = -1
pLX(116) = 1: pLY(116) = 1: pLZ(116) = 1: pHX(116) = 256: pHY(116) = 64: pHZ(116) = 16
NR(116) = 4: NC(116) = 10
pLX(117) = 1: pLY(117) = 1: pLZ(117) = 1: pHX(117) = 256: pHY(117) = 64: pHZ(117) = 16
NR(117) = 4: NC(117) = 9
pLX(118) = 1: pLY(118) = 1: pLZ(118) = 1: pHX(118) = 256: pHY(118) = 64: pHZ(118) = 16
NR(118) = 4: NC(118) = 11
pLX(119) = 1: pLY(119) = 1: pLZ(119) = 1: pHX(119) = 256: pHY(119) = 64: pHZ(119) = 16
NR(119) = 4: NC(119) = 6
pLX(120) = 1: pLY(120) = 1: pLZ(120) = 1: pHX(120) = 256: pHY(120) = 64: pHZ(120) = 16
NR(120) = 4: NC(120) = 4
pLX(121) = 1: pLY(121) = 1: pLZ(121) = 1: pHX(121) = 256: pHY(121) = 64: pHZ(121) = 16
NR(121) = 4: NC(121) = 5
pLX(122) = 240: pLY(122) = 1: pLZ(122) = 16: pHX(122) = 256: pHY(122) = 64: pHZ(122) = 256
NR(122) = 4: NC(122) = 12
pLX(123) = 1: pLY(123) = 1: pLZ(123) = 1: pHX(123) = 256: pHY(123) = 64: pHZ(123) = 16
NR(123) = 4: NC(123) = -5
pLX(124) = 1: pLY(124) = 1: pLZ(124) = 1: pHX(124) = 256: pHY(124) = 64: pHZ(124) = 16
NR(124) = 4: NC(124) = -4
pLX(125) = 1: pLY(125) = 1: pLZ(125) = 16: pHX(125) = 16: pHY(125) = 64: pHZ(125) = 256
NR(125) = 4: NC(125) = -5
pLX(126) = 1: pLY(126) = 1: pLZ(126) = 1: pHX(126) = 256: pHY(126) = 64: pHZ(126) = 16
NR(126) = 4: NC(126) = 12
pLX(127) = 1: pLY(127) = 1: pLZ(127) = 1: pHX(127) = 256: pHY(127) = 64: pHZ(127) = 16
NR(127) = 4: NC(127) = 3
pLX(128) = 1: pLY(128) = 1: pLZ(128) = 1: pHX(128) = 256: pHY(128) = 64: pHZ(128) = 16
NR(128) = 4: NC(128) = 2


' Pyramid
pLX(129) = 79: pLY(129) = 130: pLZ(129) = 80: pHX(129) = 184: pHY(129) = 194: pHZ(129) = 180
NR(129) = 16: NC(129) = 7
pLX(130) = 99: pLY(130) = 259: pLZ(130) = 96: pHX(130) = 169: pHY(130) = 323: pHZ(130) = 161
NR(130) = 16: NC(130) = 7
pLX(131) = 88: pLY(131) = 194: pLZ(131) = 88: pHX(131) = 178: pHY(131) = 258: pHZ(131) = 172
NR(131) = 16: NC(131) = 7
pLX(132) = 121: pLY(132) = 324: pLZ(132) = 114: pHX(132) = 150: pHY(132) = 388: pHZ(132) = 139
NR(132) = 16: NC(132) = 7
pLX(133) = 71: pLY(133) = 66: pLZ(133) = 73: pHX(133) = 189: pHY(133) = 130: pHZ(133) = 187
NR(133) = 16: NC(133) = 7
pLX(134) = 64: pLY(134) = 1: pLZ(134) = 66: pHX(134) = 195: pHY(134) = 65: pHZ(134) = 193
NR(134) = 16: NC(134) = 7


End Sub


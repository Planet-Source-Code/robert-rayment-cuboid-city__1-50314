Attribute VB_Name = "Module2"
' QSort.bas
Option Explicit
Option Base 1

Public Sub QSORT_BLOCKS()
' Public

'   For i = 1 To NumBlocks
'      Input #1, LX(i), LY(i), LZ(i), HX(i), HY(i), HZ(i), NR(i), NC(i)
'   Next i


' Sorts long array in de/ascending order
Dim Max As Long
Dim k As Long
Dim m As Long
Dim S As Long
Dim sortL() As Long
Dim sortR() As Long
Dim LL As Long
Dim MM As Long
Dim II As Long
Dim JJ As Long
Dim PP As Long
Dim XX As Long
Dim YY As Long
Dim MAZ As Long
MAZ = 7

   Max = NumBlocks
   k = 1
   If Max = 1 Then Exit Sub
   If k = Max Then Exit Sub
   
   
   m = Max \ 2: ReDim sortL(m), sortR(m)
   S = 1: sortL(1) = k: sortR(1) = Max
   Do While S <> 0
      LL = sortL(S): MM = sortR(S): S = S - 1
      
      Do While LL < MM
         II = LL: JJ = MM
         PP = (LL + MM) \ 2
         'XX = NR(PP)
         XX = HY(PP)
         
         Do While II <= JJ
'            ASCENDING
'            Do While NR(II) < XX: II = II + 1: Loop
'            Do While XX < NR(JJ): JJ = JJ - 1: Loop
            
            Do While HY(II) < XX: II = II + 1: Loop
            Do While XX < HY(JJ): JJ = JJ - 1: Loop
            

'            DESCENDING
'           Do While NR(II) > XX: II = II + 1: Loop
'           Do While XX > NR(JJ): JJ = JJ - 1: Loop
           
'           Do While HY(II) > XX: II = II + 1: Loop
'           Do While XX > HY(JJ): JJ = JJ - 1: Loop
            
            
            If II <= JJ Then
               If II = MAZ Then MAZ = JJ
               'SWAP LX(i), LY(i), LZ(i), HX(i), HY(i), HZ(i), NR(i), NC(i)
               YY = LX(II): LX(II) = LX(JJ): LX(JJ) = YY
               YY = LY(II): LY(II) = LY(JJ): LY(JJ) = YY
               YY = LZ(II): LZ(II) = LZ(JJ): LZ(JJ) = YY
               YY = HX(II): HX(II) = HX(JJ): HX(JJ) = YY
               YY = HY(II): HY(II) = HY(JJ): HY(JJ) = YY
               YY = HZ(II): HZ(II) = HZ(JJ): HZ(JJ) = YY
               YY = NR(II): NR(II) = NR(JJ): NR(JJ) = YY
               YY = NC(II): NC(II) = NC(JJ): NC(JJ) = YY
               
               II = II + 1: JJ = JJ - 1
            End If
         Loop
         
         If II < MM Then
            S = S + 1: sortL(S) = II: sortR(S) = MM
         End If
         MM = JJ
      Loop
   
   Loop

Erase sortL, sortR
End Sub



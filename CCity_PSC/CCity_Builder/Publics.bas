Attribute VB_Name = "Publics"
' Publics (Publics.bas)

'#### LOADING & SAVING CCC FILES ################################################
'#### DRAWING PLANS & FACES #####################################################
'#### FRAME MOVER ###############################################################

Option Explicit
Option Base 1

Public NumBlocks As Long
Public R As Long, C As Long
Public NumBlocksInRC() As Integer
Public BlockNumsInRC() As Integer
Public LX() As Integer
Public LZ() As Integer
Public LY() As Integer
Public HX() As Integer
Public HZ() As Integer
Public HY() As Integer
Public NR() As Integer
Public NC() As Integer

Public DummyHeight As Long
Public DummyDepth As Long
Public ARandHts As Boolean

Public RedBlockNumber As Long
Public FaceIndex As Long
Public FaceAction As Boolean
Public PlanAction As Boolean


Public BlockCounter As Long
' Blocks temp coords
Public X1 As Long
Public X2 As Long
Public Z1 As Long
Public Z2 As Long
Public Y1 As Long
Public Y2 As Long
Public SelectCul As Long

Public Cul As Long
Public picPLANCul As Long
Public picFaceBacCul As Long

' Grid on RC pic
Public GridWidth As Integer
Public GridHeight As Integer
Public GridStep As Integer
Public GridBackColor As Long

Public PathSpec$, CCC_Path$, FileSpec$

' For locating Preset forms
Public frmPLANLeft As Long
Public frmPLANTop As Long
Public frmFACELeft As Long
Public frmFACETop As Long
Public FaceNum As Long
Public AfrmFACEVis As Boolean

Public STX As Long
Public STY As Long

Public FName$
Public res As Long
Public i As Long, j As Long
Public ix As Long, iy As Long


'#### LOADING & SAVING CCC FILES #############################################

Public Sub READ_CCC_FILE()
'FileSpec$
On Error GoTo FERR

   Open FileSpec$ For Input As #1
   Line Input #1, FName$
   Input #1, NumBlocks
   
   ReDim LX(NumBlocks)
   ReDim LZ(NumBlocks)
   ReDim LY(NumBlocks)
   ReDim HX(NumBlocks)
   ReDim HZ(NumBlocks)
   ReDim HY(NumBlocks)
   ReDim NR(NumBlocks)
   ReDim NC(NumBlocks)
   ReDim NumBlocksInRC(-2 To 21, -5 To 12)

   For i = 1 To NumBlocks
      Input #1, LX(i), LY(i), LZ(i), HX(i), HY(i), HZ(i), NR(i), NC(i)
   Next i
   Close
   
   ' Fill out NumBlocksInRC()
   For i = 1 To NumBlocks
         NumBlocksInRC(NR(i), NC(i)) = NumBlocksInRC(NR(i), NC(i)) + 1
   Next i
   
   ' Set R & C to first block
   R = NR(1)
   C = NC(1)
Exit Sub
'========
FERR:
   Close
   MsgBox "File Error", vbCritical, "Loading ccc"
   ' Shouldn't happen unless code/file error somewhere
   ' Continue anyway with reduced NumBlocks
   NumBlocks = i
   ReDim LX(NumBlocks)
   ReDim LZ(NumBlocks)
   ReDim LY(NumBlocks)
   ReDim HX(NumBlocks)
   ReDim HZ(NumBlocks)
   ReDim HY(NumBlocks)
   ReDim NR(NumBlocks)
   ReDim NC(NumBlocks)
   ReDim NumBlocksInRC(-2 To 21, -5 To 12)
   NumBlocks = 0
   RedBlockNumber = 0
   Form1.LabBlockNum = "BlockNum =" & Str$(RedBlockNumber)
   Form1.LabBlocksInRC = "Blocks in RC =" & Str$(NumBlocksInRC(R, C))

End Sub

Public Sub ADD_CCC_FILE()
Dim NB As Long
Dim OldNumBlocks As Long
Dim TheR As Long
Dim TheC As Long
Dim AddR As Long
Dim AddC As Long

On Error GoTo ADDERR
   OldNumBlocks = NumBlocks
   Open FileSpec$ For Input As #1
   Line Input #1, FName$
   Input #1, NB
   If NumBlocks = 0 And NB > 0 Then
      NumBlocks = NB
      ReDim LX(NumBlocks)
      ReDim LZ(NumBlocks)
      ReDim LY(NumBlocks)
      ReDim HX(NumBlocks)
      ReDim HZ(NumBlocks)
      ReDim HY(NumBlocks)
      ReDim NR(NumBlocks)
      ReDim NC(NumBlocks)
      ReDim NumBlocksInRC(-2 To 21, -5 To 12)
      
      For i = 1 To NumBlocks
         Input #1, LX(i), LY(i), LZ(i), HX(i), HY(i), HZ(i), NR(i), NC(i)
      Next i
      Close
      
      ' Fill out NumBlocksInRC()
      For i = 1 To NumBlocks
            NumBlocksInRC(NR(i), NC(i)) = NumBlocksInRC(NR(i), NC(i)) + 1
      Next i
      
      ' Set R & C to first block
      R = NR(1)
      C = NC(1)
      
   ElseIf NumBlocks > 0 And NB > 0 Then
      TheR = R
      TheC = C
      NumBlocks = NumBlocks + NB
      ReDim Preserve LX(NumBlocks)
      ReDim Preserve LZ(NumBlocks)
      ReDim Preserve LY(NumBlocks)
      ReDim Preserve HX(NumBlocks)
      ReDim Preserve HZ(NumBlocks)
      ReDim Preserve HY(NumBlocks)
      ReDim Preserve NR(NumBlocks)
      ReDim Preserve NC(NumBlocks)
      ReDim NumBlocksInRC(-2 To 21, -5 To 12)
   
      i = OldNumBlocks + 1
      Input #1, LX(i), LY(i), LZ(i), HX(i), HY(i), HZ(i), AddR, AddC
      NR(i) = TheR: NC(i) = TheC
      
      If OldNumBlocks + 1 < NumBlocks Then
         For i = OldNumBlocks + 2 To NumBlocks
            Input #1, LX(i), LY(i), LZ(i), HX(i), HY(i), HZ(i), NR(i), NC(i)
            If NR(i) = AddR And NC(i) = AddC Then
               NR(i) = TheR: NC(i) = TheC
            End If
         Next i
         Close
      End If
      
      ' Fill out NumBlocksInRC()
      For i = 1 To NumBlocks
            NumBlocksInRC(NR(i), NC(i)) = NumBlocksInRC(NR(i), NC(i)) + 1
      Next i
      
      ' Set R & C to first block
      R = NR(1)
      C = NC(1)
   Exit Sub
Else
'========
ADDERR:
   Close
   MsgBox "File Error", vbCritical, "Loading ccc"
   ' Shouldn't happen unless code/file error somewhere
   ' Continue anyway with reduced NumBlocks
End If
End Sub

Public Sub SAVE_CCC_FILE()
Dim NB As Long

On Error GoTo SERR

QSORT_BLOCKS   ' In Descending order based on Row NR()
' CCityAnim will read these from the start of the file but
' number them 1,2,3,,NumBlocks so the last shall be first
' and get overdrawn by later blocks in later RC-squares

'FileSpec$
   j = InStrRev(FileSpec$, "\")
   FName$ = Mid$(FileSpec$, j + 1)
   
   
   Open FileSpec$ For Output As #1
   Print #1, FName$
   Print #1, NumBlocks
   NB = 0
   For i = 1 To NumBlocks
      If NumBlocksInRC(NR(i), NC(i)) <> 0 Then
         Print #1, Format$(LX(i), "@@@"); ",";
         Print #1, Format$(LY(i), "@@@"); ",";
         Print #1, Format$(LZ(i), "@@@"); ",";
         Print #1, Format$(HX(i), "@@@"); ",";
         Print #1, Format$(HY(i), "@@@"); ",";
         Print #1, Format$(HZ(i), "@@@"); ",";
         Print #1, Format$(NR(i), "@@@"); ",";
         Print #1, Format$(NC(i), "@@@")
      NB = NB + 1
      End If
   Next i
   Close
   If NB <> NumBlocks Then
      ' Shouldn't happen unless code error somewhere
      MsgBox " Block count error", vbCritical, "Saving ccc"
   End If
   Exit Sub
'============
SERR:
   Close
   MsgBox " Saving error @ Block" & Str$(i), vbCritical, "Saving ccc"
End Sub

Public Sub FixExtension(FSpec$, Ext$)
Dim p As Long

If Len(FSpec$) = 0 Then Exit Sub
   Ext$ = LCase$(Ext$)
   
   p = InStr(1, FSpec$, ".")
   
   If p = 0 Then
      FSpec$ = FSpec$ & Ext$
   Else
      'Ext$ = LCase$(Mid$(FSpec$, p))
      If LCase$(Mid$(FSpec$, p)) <> Ext$ Then FSpec$ = Mid$(FSpec$, 1, p) & Ext$
   End If
End Sub
'#### END LOADING & SAVING CCC FILES ############################################


'#### DRAWING PLANS & FACES #####################################################

Public Sub REDRAW_PLAN(PIC As PictureBox)
' @ R,C
' In: RedBlockNumber for red
   PIC.Cls
   BlockCounter = 0
   For i = 1 To NumBlocks
      If NR(i) = R Then
      If NC(i) = C Then
         X1 = LX(i)
         Z1 = LZ(i)
         X2 = HX(i)
         Z2 = HZ(i)
         If i = RedBlockNumber Then
            Cul = picPLANCul Xor SelectCul
         Else
            Cul = picPLANCul Xor 0
         End If
         PIC.Line (X1, Z1)-(X2, Z2), Cul, B
         PIC.Circle (X1, Z1), 2, Cul
         BlockCounter = BlockCounter + 1
         If BlockCounter = NumBlocksInRC(R, C) Then Exit For
      End If
      End If
   Next i
End Sub

Public Sub DRAW_Faces(PIC As PictureBox)
Dim HtExtra As Long
' @ R,C
' In: RedBlockNumber for red
   
   PIC.Enabled = True
   PIC.Cls
   PIC.DrawMode = 7
   HtExtra = 0
   If ARandHts Then
      HtExtra = CLng(256 * Rnd + 1)
   End If
   
   BlockCounter = 0
   
   Select Case FaceIndex
   Case 0   ' 1 Front face X,Y
      For i = 1 To NumBlocks
         If NR(i) = R Then
         If NC(i) = C Then
            X1 = LX(i)
            X2 = HX(i)
            If LY(i) <= 1 Then LY(i) = 1
            If HY(i) <= 1 Then HY(i) = LY(i) + DummyHeight
            
            If ARandHts Then
               If HY(i) + HtExtra < 384 Then
                  HY(i) = HY(i) + HtExtra
               End If
            End If
            
            Y1 = LY(i)
            Y2 = HY(i)
            If i = RedBlockNumber Then
               Cul = picFaceBacCul Xor SelectCul
            Else
               Cul = picFaceBacCul Xor 0
            End If
            PIC.Line (X1, Y1)-(X2, Y2), Cul, B
            PIC.Circle (X1, Y1), 2, Cul
            BlockCounter = BlockCounter + 1
            If BlockCounter = NumBlocksInRC(R, C) Then Exit For
         End If
         End If
      Next i
   
   Case 1   ' 2 Right face  Z,Y
      For i = 1 To NumBlocks
         If NR(i) = R Then
         If NC(i) = C Then
            Z1 = LZ(i)
            Z2 = HZ(i)
            If LY(i) <= 1 Then LY(i) = 1
            If HY(i) <= 1 Then HY(i) = LY(i) + DummyHeight
            
            If ARandHts Then
               If HY(i) + HtExtra < 384 Then
                  HY(i) = HY(i) + HtExtra
               End If
            End If
            
            Y1 = LY(i)
            Y2 = HY(i)
            If i = RedBlockNumber Then
               Cul = picFaceBacCul Xor SelectCul
            Else
               Cul = picFaceBacCul Xor 0
            End If
            PIC.Line (Z1, Y1)-(Z2, Y2), Cul, B
            PIC.Circle (Z1, Y1), 2, Cul
            BlockCounter = BlockCounter + 1
            If BlockCounter = NumBlocksInRC(R, C) Then Exit For
         End If
         End If
      Next i
   
   Case 2   ' 3 Back face -X,Y
      For i = 1 To NumBlocks
         If NR(i) = R Then
         If NC(i) = C Then
            X1 = LX(i)
            X2 = HX(i)
            If LY(i) <= 1 Then LY(i) = 1
            If HY(i) <= 1 Then HY(i) = LY(i) + DummyHeight
            
            If ARandHts Then
               If HY(i) + HtExtra < 384 Then
                  HY(i) = HY(i) + HtExtra
               End If
            End If
            
            Y1 = LY(i)
            Y2 = HY(i)
            If i = RedBlockNumber Then
               Cul = picFaceBacCul Xor SelectCul
            Else
               Cul = picFaceBacCul Xor 0
            End If
            PIC.Line (256 - X1, Y1)-(256 - X2, Y2), Cul, B
            PIC.Circle (256 - X1, Y1), 2, Cul
            BlockCounter = BlockCounter + 1
            If BlockCounter = NumBlocksInRC(R, C) Then Exit For
         End If
         End If
      Next i
   
   Case 3   ' 4 Left face -Z,Y
      For i = 1 To NumBlocks
         If NR(i) = R Then
         If NC(i) = C Then
            Z1 = LZ(i)
            Z2 = HZ(i)
            If LY(i) < 1 Then LY(i) = 1
            If HY(i) < 1 Then HY(i) = LY(i) + DummyHeight
            
            If ARandHts Then
               If HY(i) + HtExtra < 384 Then
                  HY(i) = HY(i) + HtExtra
               End If
            End If
            
            Y1 = LY(i)
            Y2 = HY(i)
            If i = RedBlockNumber Then
               Cul = picFaceBacCul Xor SelectCul
            Else
               Cul = picFaceBacCul Xor 0
            End If
            PIC.Line (256 - Z1, Y1)-(256 - Z2, Y2), Cul, B
            PIC.Circle (256 - Z1, Y1), 2, Cul
            BlockCounter = BlockCounter + 1
            If BlockCounter = NumBlocksInRC(R, C) Then Exit For
         End If
         End If
      Next i
   Case 4   ' 6 Whole plan
      If NumBlocks > 0 Then
         PIC.DrawMode = 13
         ' Top boundary
         PIC.Line (1, 344)-(256, 344), vbBlack
         
         For i = 1 To NumBlocks
            X1 = (1& * LX(i) + 1& * (NC(i) + 5) * 256) \ 18
            Z1 = (1& * LZ(i) + 1& * (NR(i) + 2) * 256) \ 18
            X2 = (1& * HX(i) + 1& * (NC(i) + 5) * 256) \ 18
            Z2 = (1& * HZ(i) + 1& * (NR(i) + 2) * 256) \ 18
            If X2 <= X1 Then X2 = X1 + 2
            If Z2 <= Z1 Then Z2 = Z1 + 1
            If i = RedBlockNumber Then
               PIC.Line (X1, Z1)-(X2, Z2), vbRed, BF
            Else
               PIC.Line (X1, Z1)-(X2, Z2), vbBlack, B
            End If
         Next i
      End If
   Case 5      ' RC Perspec
      If NumBlocks > 0 Then
         PIC.DrawMode = 13
         'eyeX = 128 '256 + 128
         DRAW_RC_PERSPEC
      End If
   End Select
End Sub
'#### END DRAWING PLANS & FACES #####################################################


'#### FRAME MOVER ###################################################################

Public Sub fraMOVER(frm As Form, fra As Frame, Xfra As Single, Yfra As Single, Button As Integer, X As Single, Y As Single)
Dim fraLeft As Long
Dim fraTop As Long

   If Button = vbLeftButton Then
      
      fraLeft = fra.Left + (X - Xfra) \ STX
      If fraLeft < 0 Then fraLeft = 0
      If fraLeft + fra.Width > frm.Width \ STX + fra.Width \ 2 Then
         fraLeft = frm.Width \ STX - fra.Width \ 2
      End If
      fra.Left = fraLeft
      
      fraTop = fra.Top + (Y - Yfra) \ STY
      If fraTop < 8 Then fraTop = 8
      If fraTop + fra.Height > frm.Height \ STY + fra.Height \ 2 Then
         fraTop = frm.Height \ STY - fra.Height \ 2
      End If
      fra.Top = fraTop
      
   End If
End Sub

' Trebor Tnemyar

Public Sub SAVE_CCS_FILE()
Dim NB As Long

On Error GoTo SSERR

QSORT_BLOCKS   ' In Descending order based on Row (R()
' CCityAnim will read these from the start of the file but
' number them 1,2,3,,NumBlocks so the last shall be first
' and get overdrawn by later blocks in later RC-squares

'FileSpec$
   j = InStrRev(FileSpec$, "\")
   FName$ = Mid$(FileSpec$, j + 1)
   
   
   Open FileSpec$ For Output As #1
   Print #1, FName$
   Print #1, "NoBoxes =" & Str$(NumBlocks)
   NB = 0
   For i = 1 To NumBlocks
      If NumBlocksInRC(NR(i), NC(i)) <> 0 Then
         Print #1, "pLX(" & Trim$(Str$(i)) & ") =" & Str$(LX(i)); ": ";
         Print #1, "pLY(" & Trim$(Str$(i)) & ") =" & Str$(LY(i)); ": ";
         Print #1, "pLZ(" & Trim$(Str$(i)) & ") =" & Str$(LZ(i)); ": ";
         Print #1, "pHX(" & Trim$(Str$(i)) & ") =" & Str$(HX(i)); ": ";
         Print #1, "pHY(" & Trim$(Str$(i)) & ") =" & Str$(HY(i)); ": ";
         Print #1, "pHZ(" & Trim$(Str$(i)) & ") =" & Str$(HZ(i))
         Print #1, "R(" & Trim$(Str$(i)) & ") =" & Str$(NR(i)); ": ";
         Print #1, "C(" & Trim$(Str$(i)) & ") =" & Str$(NC(i))
         
      NB = NB + 1
      End If
   Next i
   Close
   If NB <> NumBlocks Then
      ' Shouldn't happen unless code error somewhere
      MsgBox " Block count error", vbCritical, "Saving ccs"
   End If
   Exit Sub
'============
SSERR:
   Close
   MsgBox " Saving error @ Block" & Str$(i), vbCritical, "Saving ccs"
End Sub


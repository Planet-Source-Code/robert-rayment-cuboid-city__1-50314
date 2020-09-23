VERSION 5.00
Begin VB.Form frmPLANPreSets 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " PLAN Shapes"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2445
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   263
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   163
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMaze 
      Caption         =   "Make Maze"
      Height          =   300
      Left            =   585
      TabIndex        =   22
      Top             =   3540
      Width           =   1350
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Index           =   19
      Left            =   1830
      ScaleHeight     =   -32
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   32
      ScaleWidth      =   32
      TabIndex        =   20
      Top             =   2370
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Index           =   18
      Left            =   1245
      ScaleHeight     =   -32
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   32
      ScaleWidth      =   32
      TabIndex        =   19
      Top             =   2370
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Index           =   17
      Left            =   645
      ScaleHeight     =   -32
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   32
      ScaleWidth      =   32
      TabIndex        =   18
      Top             =   2370
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Index           =   16
      Left            =   30
      ScaleHeight     =   -32
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   32
      ScaleWidth      =   32
      TabIndex        =   17
      Top             =   2370
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Index           =   15
      Left            =   1830
      ScaleHeight     =   -32
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   32
      ScaleWidth      =   32
      TabIndex        =   16
      Top             =   1785
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Index           =   14
      Left            =   1245
      ScaleHeight     =   -32
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   32
      ScaleWidth      =   32
      TabIndex        =   15
      Top             =   1785
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Index           =   13
      Left            =   645
      ScaleHeight     =   -32
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   32
      ScaleWidth      =   32
      TabIndex        =   14
      Top             =   1785
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Index           =   12
      Left            =   30
      ScaleHeight     =   -32
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   32
      ScaleWidth      =   32
      TabIndex        =   13
      Top             =   1785
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Index           =   11
      Left            =   1830
      ScaleHeight     =   -32
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   32
      ScaleWidth      =   32
      TabIndex        =   12
      Top             =   1200
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Index           =   10
      Left            =   1245
      ScaleHeight     =   -32
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   32
      ScaleWidth      =   32
      TabIndex        =   11
      Top             =   1200
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Index           =   9
      Left            =   645
      ScaleHeight     =   -32
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   32
      ScaleWidth      =   32
      TabIndex        =   10
      Top             =   1215
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Index           =   8
      Left            =   45
      ScaleHeight     =   -32
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   32
      ScaleWidth      =   32
      TabIndex        =   9
      Top             =   1215
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Index           =   7
      Left            =   1830
      ScaleHeight     =   -32
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   32
      ScaleWidth      =   32
      TabIndex        =   8
      Top             =   630
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Index           =   6
      Left            =   1245
      ScaleHeight     =   -32
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   32
      ScaleWidth      =   32
      TabIndex        =   7
      Top             =   630
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Index           =   5
      Left            =   645
      ScaleHeight     =   -32
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   32
      ScaleWidth      =   32
      TabIndex        =   6
      Top             =   630
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Index           =   4
      Left            =   45
      ScaleHeight     =   -32
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   32
      ScaleWidth      =   32
      TabIndex        =   5
      Top             =   630
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Index           =   3
      Left            =   1830
      ScaleHeight     =   -32
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   32
      ScaleWidth      =   32
      TabIndex        =   4
      Top             =   45
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Index           =   2
      Left            =   1230
      ScaleHeight     =   -32
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   32
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   45
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Index           =   1
      Left            =   630
      ScaleHeight     =   -32
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   32
      ScaleWidth      =   32
      TabIndex        =   2
      Top             =   45
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Index           =   0
      Left            =   45
      ScaleHeight     =   -32
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   32
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   45
      Width           =   540
   End
   Begin VB.Label LabIndex 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   210
      Left            =   2130
      TabIndex        =   21
      Top             =   3255
      Width           =   210
   End
   Begin VB.Label Label1 
      Caption         =   "Click on shape to add to whole RC-PLAN-square."
      Height          =   480
      Left            =   60
      TabIndex        =   1
      Top             =   2985
      Width           =   1875
   End
End
Attribute VB_Name = "frmPLANPreSets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmPLANPreSets (frmPLANPreSets.bas)

Option Explicit
Option Base 1

'--------------------------------------------------------------------------
'  API to make application stay on top
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
ByVal wi As Long, ByVal ht As Long, ByVal wFlags As Long) As Long

Const hWndInsertAfter = -1
Const wFlags = &H40 Or &H20
'---------------------------------------------------------------------------
Dim pLX() As Integer, pLZ() As Integer, pLY() As Integer
Dim pHX() As Integer, pHZ() As Integer, pHY() As Integer
Dim NB As Long
Dim NoBoxes As Long


Private Sub cmdMaze_Click()
   MakeMaze
End Sub

Private Sub Form_Load()
' frmPLANLeft = 430
' frmPLANTop = 200

Dim FW As Long, FH As Long
FW = frmPLANPreSets.Width
FH = frmPLANPreSets.Height
' Size & Make frmPLANPreSets stay on top
res = SetWindowPos(frmPLANPreSets.hWnd, hWndInsertAfter, _
   frmPLANLeft, frmPLANTop, FW \ STX, FH \ STY, wFlags)

frmPLANPreSets.Show

   ReDim pLX(1 To 4)
   ReDim pLZ(1 To 4)
   ReDim pLY(1 To 4)
   ReDim pHX(1 To 4)
   ReDim pHZ(1 To 4)
   ReDim pHY(1 To 4)

   ' DrawShapes
   For i = 0 To 19
      picPDraw
   Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmPLANLeft = frmPLANPreSets.Left \ STX
   frmPLANTop = frmPLANPreSets.Top \ STY
   Erase pLX, pLZ, pLY, pHX, pHY, pHZ
   Unload frmPLANPreSets
End Sub

Private Sub picP_Click(Index As Integer)
   LabIndex = Trim$(Str$(Index))
   
   LCoords Index
   
   NB = NumBlocks + 1
   NumBlocks = NumBlocks + NoBoxes 'NOB(Index + 1)
   ReDimPreserveLH
   For i = 1 To NoBoxes
      LX(NB) = pLX(i)
      LZ(NB) = pLZ(i)
      LY(NB) = pLY(i)
      HX(NB) = pHX(i)
      HZ(NB) = pHZ(i)
      HY(NB) = pHY(i)
      NR(NB) = R
      NC(NB) = C
      NumBlocksInRC(NR(NB), NC(NB)) = NumBlocksInRC(NR(NB), NC(NB)) + 1
      NB = NB + 1
   Next i
   
   Form1.LabNumBlocks = "NumBlocks =" & Str$(NumBlocks)
   RedBlockNumber = NumBlocks
   Form1.LabBlockNum = "BlockNum =" & Str$(RedBlockNumber)
   Form1.LabBlocksInRC = "Blocks in RC =" & Str$(NumBlocksInRC(R, C))
   REDRAW_PLAN Form1.picPLAN
   DRAW_Faces Form1.picFace
   'Unload Me
End Sub

Private Sub ReDimPreserveLH()
   If NB = 1 Then
      ReDim LX(NumBlocks)
      ReDim LZ(NumBlocks)
      ReDim LY(NumBlocks)
      ReDim HX(NumBlocks)
      ReDim HZ(NumBlocks)
      ReDim HY(NumBlocks)
      ReDim NR(NumBlocks)
      ReDim NC(NumBlocks)
   Else
      ReDim Preserve LX(NumBlocks)
      ReDim Preserve LZ(NumBlocks)
      ReDim Preserve LY(NumBlocks)
      ReDim Preserve HX(NumBlocks)
      ReDim Preserve HZ(NumBlocks)
      ReDim Preserve HY(NumBlocks)
      ReDim Preserve NR(NumBlocks)
      ReDim Preserve NC(NumBlocks)
   End If
End Sub

Private Sub LCoords(Index As Integer)
   Select Case Index
   Case 0
      NoBoxes = 1
      pLX(1) = 1: pLY(1) = 1: pLZ(1) = 240: pHX(1) = 256: pHY(1) = 16: pHZ(1) = 256
   Case 1
      NoBoxes = 1
      pLX(1) = 1: pLY(1) = 1: pLZ(1) = 1: pHX(1) = 16: pHY(1) = 16: pHZ(1) = 256
   Case 2
      NoBoxes = 1
      pLX(1) = 1: pLY(1) = 1: pLZ(1) = 1: pHX(1) = 256: pHY(1) = 16: pHZ(1) = 16
   Case 3
      NoBoxes = 1
      pLX(1) = 240: pLY(1) = 1: pLZ(1) = 1: pHX(1) = 256: pHY(1) = 16: pHZ(1) = 256
   Case 4
      NoBoxes = 2
      pLX(1) = 1: pLY(1) = 1: pLZ(1) = 240: pHX(1) = 256: pHY(1) = 16: pHZ(1) = 256
      pLX(2) = 240: pLY(2) = 1: pLZ(2) = 1:  pHX(2) = 256: pHY(2) = 16: pHZ(2) = 240
   Case 5
      NoBoxes = 2
      pLX(1) = 1: pLY(1) = 1: pLZ(1) = 240: pHX(1) = 256: pHY(1) = 16: pHZ(1) = 256
      pLX(2) = 1: pLY(2) = 1: pLZ(2) = 1: pHX(2) = 16: pHY(2) = 16: pHZ(2) = 240
   Case 6
      NoBoxes = 2
      pLX(1) = 1: pLY(1) = 1: pLZ(1) = 16: pHX(1) = 16: pHY(1) = 16: pHZ(1) = 256
      pLX(2) = 1: pLY(2) = 1: pLZ(2) = 1: pHX(2) = 256: pHY(2) = 16: pHZ(2) = 16
   Case 7
      NoBoxes = 2
      pLX(1) = 1: pLY(1) = 1: pLZ(1) = 1: pHX(1) = 256: pHY(1) = 16: pHZ(1) = 16
      pLX(2) = 240: pLY(2) = 1: pLZ(2) = 16: pHX(2) = 256: pHY(2) = 16: pHZ(2) = 256
   Case 8
      NoBoxes = 2
      pLX(1) = 1: pLY(1) = 1: pLZ(1) = 96: pHX(1) = 160: pHY(1) = 16: pHZ(1) = 160
      pLX(2) = 96: pLY(2) = 1: pLZ(2) = 160: pHX(2) = 160: pHY(2) = 16: pHZ(2) = 256
   Case 9
      NoBoxes = 2
      pLX(1) = 1: pLY(1) = 1: pLZ(1) = 96: pHX(1) = 160: pHY(1) = 16: pHZ(1) = 160
      pLX(2) = 96: pLY(2) = 1: pLZ(2) = 1: pHX(2) = 160: pHY(2) = 16: pHZ(2) = 96
   Case 10
      NoBoxes = 2
      pLX(1) = 96: pLY(1) = 1: pLZ(1) = 1: pHX(1) = 160: pHY(1) = 16: pHZ(1) = 96
      pLX(2) = 96: pLY(2) = 1: pLZ(2) = 96: pHX(2) = 256: pHY(2) = 16: pHZ(2) = 160
   Case 11
      NoBoxes = 2
      pLX(1) = 96: pLY(1) = 1: pLZ(1) = 96: pHX(1) = 256: pHY(1) = 16: pHZ(1) = 160
      pLX(2) = 96: pLY(2) = 1: pLZ(2) = 160: pHX(2) = 160: pHY(2) = 16: pHZ(2) = 256
   Case 12
      NoBoxes = 2
      pLX(1) = 1: pLY(1) = 1: pLZ(1) = 96: pHX(1) = 256: pHY(1) = 16: pHZ(1) = 160
      pLX(2) = 96: pLY(2) = 1: pLZ(2) = 160: pHX(2) = 160: pHY(2) = 16: pHZ(2) = 256
   Case 13
      NoBoxes = 2
      pLX(1) = 1: pLY(1) = 1: pLZ(1) = 96: pHX(1) = 96: pHY(1) = 16: pHZ(1) = 160
      pLX(2) = 96: pLY(2) = 1: pLZ(2) = 1: pHX(2) = 160: pHY(2) = 16: pHZ(2) = 256
   Case 14
      NoBoxes = 2
      pLX(1) = 1: pLY(1) = 1: pLZ(1) = 96: pHX(1) = 256: pHY(1) = 16: pHZ(1) = 160
      pLX(2) = 96: pLY(2) = 1: pLZ(2) = 1: pHX(2) = 160: pHY(2) = 16: pHZ(2) = 96
   Case 15
      NoBoxes = 2
      pLX(1) = 96: pLY(1) = 1: pLZ(1) = 1: pHX(1) = 160: pHY(1) = 16: pHZ(1) = 256
      pLX(2) = 160: pLY(2) = 1: pLZ(2) = 96: pHX(2) = 256: pHY(2) = 16: pHZ(2) = 160
   Case 16
      NoBoxes = 1
      pLX(1) = 1: pLY(1) = 1: pLZ(1) = 96: pHX(1) = 256: pHY(1) = 16: pHZ(1) = 160
   Case 17
      NoBoxes = 1
      pLX(1) = 96: pLY(1) = 1: pLZ(1) = 1: pHX(1) = 160: pHY(1) = 16: pHZ(1) = 256
   Case 18
      NoBoxes = 3
      pLX(1) = 1: pLY(1) = 1: pLZ(1) = 96: pHX(1) = 256: pHY(1) = 16: pHZ(1) = 160
      pLX(2) = 96: pLY(2) = 1: pLZ(2) = 1: pHX(2) = 160: pHY(2) = 16: pHZ(2) = 96
      pLX(3) = 96: pLY(3) = 1: pLZ(3) = 160: pHX(3) = 160: pHY(3) = 16: pHZ(3) = 256
   Case 19
      NoBoxes = 4
      pLX(1) = 1: pLY(1) = 1: pLZ(1) = 1: pHX(1) = 256: pHY(1) = 16: pHZ(1) = 32
      pLX(2) = 1: pLY(2) = 1: pLZ(2) = 32: pHX(2) = 32: pHY(2) = 16: pHZ(2) = 224
      pLX(3) = 224: pLY(3) = 1: pLZ(3) = 32: pHX(3) = 256: pHY(3) = 16: pHZ(3) = 224
      pLX(4) = 1: pLY(4) = 1: pLZ(4) = 224: pHX(4) = 256: pHY(4) = 16: pHZ(4) = 256
   End Select
End Sub

Private Sub picPDraw()

'   ' DrawShapes
'   For i = 0 To 19
'      picPDraw
'   Next i
' Public i

Dim ixL As Long
Dim izL As Long
Dim ixH As Long
Dim izH As Long
   LCoords CInt(i)  ' Sets pLX() etc & NoBoxes
   For j = 1 To NoBoxes
      ixL = pLX(j): If ixL > 1 Then ixL = ixL \ 8
      izL = pLZ(j): If izL > 1 Then izL = izL \ 8
      ixH = pHX(j): If ixH > 1 Then ixH = ixH \ 8
      izH = pHZ(j): If izH > 1 Then izH = izH \ 8
      Select Case i
      Case 1
         picP(i).Line (ixL, izL)-(ixH + 2, izH), 0, B
      Case 2
         picP(i).Line (ixL, izL)-(ixH, izH + 2), 0, B
      Case 3
         picP(i).Line (ixL - 1, izL)-(ixH, izH), 0, B
      Case 4
         If j = 2 Then
            picP(i).Line (ixL - 1, izL)-(ixH, izH), 0, B
         Else
            picP(i).Line (ixL, izL)-(ixH, izH), 0, B
         End If
      Case 5
         If j = 2 Then
            picP(i).Line (ixL, izL)-(ixH + 2, izH), 0, B
         Else
            picP(i).Line (ixL, izL)-(ixH, izH), 0, B
         End If
      Case 6
         If j = 1 Then
            picP(i).Line (ixL, izL)-(ixH + 2, izH), 0, B
         Else
            picP(i).Line (ixL, izL)-(ixH, izH + 2), 0, B
         End If
      Case 7
         If j = 1 Then
            picP(i).Line (ixL, izL)-(ixH, izH + 2), 0, B
         Else
            picP(i).Line (ixL - 1, izL)-(ixH, izH), 0, B
         End If
      Case Else
         picP(i).Line (ixL, izL)-(ixH, izH), 0, B
      End Select
   Next j
End Sub

' Trebor Tnemyar


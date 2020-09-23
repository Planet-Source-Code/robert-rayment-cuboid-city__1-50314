VERSION 5.00
Begin VB.Form frmFACEPresets 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FACE shapes"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2460
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   279
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   164
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   1020
      Index           =   11
      Left            =   1815
      ScaleHeight     =   -64
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   64
      ScaleWidth      =   32
      TabIndex        =   14
      Top             =   2565
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   1020
      Index           =   10
      Left            =   1230
      ScaleHeight     =   -64
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   64
      ScaleWidth      =   32
      TabIndex        =   13
      Top             =   2565
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   1020
      Index           =   9
      Left            =   645
      ScaleHeight     =   -64
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   64
      ScaleWidth      =   32
      TabIndex        =   12
      Top             =   2565
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   1020
      Index           =   8
      Left            =   45
      ScaleHeight     =   -64
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   64
      ScaleWidth      =   32
      TabIndex        =   11
      Top             =   2565
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   1020
      Index           =   7
      Left            =   1830
      ScaleHeight     =   -64
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   64
      ScaleWidth      =   32
      TabIndex        =   10
      Top             =   1470
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   1020
      Index           =   6
      Left            =   1230
      ScaleHeight     =   -64
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   64
      ScaleWidth      =   32
      TabIndex        =   9
      Top             =   1470
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   1020
      Index           =   5
      Left            =   645
      ScaleHeight     =   -64
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   64
      ScaleWidth      =   32
      TabIndex        =   8
      Top             =   1485
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   1020
      Index           =   4
      Left            =   45
      ScaleHeight     =   -64
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   64
      ScaleWidth      =   32
      TabIndex        =   7
      Top             =   1485
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   1020
      Index           =   3
      Left            =   1830
      ScaleHeight     =   -64
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   64
      ScaleWidth      =   32
      TabIndex        =   6
      Top             =   420
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   1020
      Index           =   2
      Left            =   1230
      ScaleHeight     =   -64
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   64
      ScaleWidth      =   32
      TabIndex        =   5
      Top             =   420
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   1020
      Index           =   1
      Left            =   645
      ScaleHeight     =   -64
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   64
      ScaleWidth      =   32
      TabIndex        =   4
      Top             =   420
      Width           =   540
   End
   Begin VB.PictureBox picP 
      BackColor       =   &H00FFFFFF&
      Height          =   1020
      Index           =   0
      Left            =   60
      ScaleHeight     =   -64
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleTop        =   64
      ScaleWidth      =   32
      TabIndex        =   3
      Top             =   420
      Width           =   540
   End
   Begin VB.CommandButton cmdFaceShift 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1860
      TabIndex        =   2
      Top             =   75
      Width           =   285
   End
   Begin VB.CommandButton cmdFaceShift 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1515
      TabIndex        =   1
      Top             =   75
      Width           =   285
   End
   Begin VB.Label LabIndex 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   285
      Left            =   2145
      TabIndex        =   16
      Top             =   3870
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Click on shape to add to whole FACE RC-square."
      Height          =   465
      Left            =   60
      TabIndex        =   15
      Top             =   3660
      Width           =   1920
   End
   Begin VB.Label LabFaces 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "Front face"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   195
      TabIndex        =   0
      Top             =   75
      Width           =   1200
   End
End
Attribute VB_Name = "frmFACEPresets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmFACEPreSets (frmFACEPreSets.bas)

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

'Dim FaceNum As Long

Private Sub Form_Load()
' frmFACELeft = 700
' frmFACETop = 140

Dim FW As Long, FH As Long
   FW = frmFACEPresets.Width
   FH = frmFACEPresets.Height
   ' Size & Make frmPLANPreSets stay on top
   res = SetWindowPos(frmFACEPresets.hWnd, hWndInsertAfter, _
      frmFACELeft, frmFACETop, FW \ STX, FH \ STY, wFlags)
   
   AfrmFACEVis = True
   frmFACEPresets.Show
   
   ReDim pLX(1 To 16)
   ReDim pLZ(1 To 16)
   ReDim pLY(1 To 16)
   ReDim pHX(1 To 16)
   ReDim pHZ(1 To 16)
   ReDim pHY(1 To 16)

   ' DrawShapes
   For i = 0 To 11
      picPDraw CInt(i)
   Next i
   
   If FaceIndex < 4 Then
      FaceNum = FaceIndex
   Else  ' Avoid WholePlan & RC Perspec
      FaceIndex = 0
      FaceNum = 0
   End If
   Select Case FaceNum
   Case 0: LabFaces = "  Front face "
   Case 1: LabFaces = "  Right face "
   Case 2: LabFaces = "  Back face "
   Case 3: LabFaces = "  Left face "
   End Select
   SetCorrectFace

End Sub

'#### FACES ##################################################################################

Private Sub cmdFaceShift_Click(Index As Integer)
' Cycle thru' faces
   Select Case Index
   Case 0   '<
      If FaceNum = 0 Then FaceNum = 4
      FaceNum = FaceNum - 1
   Case 1   '>
      If FaceNum = 3 Then FaceNum = -1
      FaceNum = FaceNum + 1
   End Select
   Select Case FaceNum
   Case 0: LabFaces = "  Front face "
   Case 1: LabFaces = "  Right face "
   Case 2: LabFaces = "  Back face "
   Case 3: LabFaces = "  Left face "
   End Select
   
   SetCorrectFace
End Sub

Private Sub picP_Click(Index As Integer)
   LabIndex = Trim$(Str$(Index))
   
   LCoords Index
   
   NB = NumBlocks + 1
   NumBlocks = NumBlocks + NoBoxes 'NOB(Index + 1)
   ReDimPreserveLH
   For i = 1 To NoBoxes
      Select Case FaceNum
      Case 0   ' FrontFace
         LX(NB) = pLX(i)
         LY(NB) = pLY(i)
         LZ(NB) = pLZ(i)
         HX(NB) = pHX(i)
         HY(NB) = pHY(i)
         HZ(NB) = pHZ(i)
      Case 1   ' Right face
         LX(NB) = 256 - pLZ(i)
         HX(NB) = 256 - pHZ(i)
         LY(NB) = pLY(i)
         LZ(NB) = pLX(i)
         HY(NB) = pHY(i)
         HZ(NB) = pHX(i)
      Case 2   ' Back face
'Depends on where front face wanted
'         LX(NB) = 256 - pLX(i)
'         HX(NB) = 256 - pHX(i)
'         LZ(NB) = 256 - pLZ(i)
'         HZ(NB) = 256 - pHZ(i)
         LX(NB) = 256 - pHX(i)
         HX(NB) = 256 - pLX(i)
         LZ(NB) = 256 - pHZ(i)
         HZ(NB) = 256 - pLZ(i)
         
         LY(NB) = pLY(i)
         HY(NB) = pHY(i)
      Case 3   ' Left face
         LY(NB) = pLY(i)
         HY(NB) = pHY(i)
' Ditto
'         LX(NB) = pLZ(i)
'         HX(NB) = pHZ(i)
'         LZ(NB) = 256 - pLX(i)
'         HZ(NB) = 256 - pHX(i)
         LX(NB) = pHZ(i)
         HX(NB) = pLZ(i)
         LZ(NB) = 256 - pHX(i)
         HZ(NB) = 256 - pLX(i)
      End Select
      
      
      NR(NB) = R
      NC(NB) = C
      NumBlocksInRC(NR(NB), NC(NB)) = NumBlocksInRC(NR(NB), NC(NB)) + 1
      NB = NB + 1
      
   Next i
   
   SetCorrectFace
   
   Form1.LabNumBlocks = "NumBlocks =" & Str$(NumBlocks)
   RedBlockNumber = NumBlocks
   Form1.LabBlockNum = "BlockNum =" & Str$(RedBlockNumber)
   Form1.LabBlocksInRC = "Blocks in RC =" & Str$(NumBlocksInRC(R, C))
   REDRAW_PLAN Form1.picPLAN
   DRAW_Faces Form1.picFace
   'Unload Me
End Sub

Private Sub SetCorrectFace()
   FaceIndex = FaceNum
   With Form1
      .LabFaceCap = "Top"
      Select Case FaceIndex
      Case 0: .LabFace = "  Front face "
              .LabFaceAxis(0) = "X"
              .LabFaceAxis(1) = "1,1"
              .LabFaceAxis(2) = "Y"
              .LabFaceAxis(3) = ""
              DRAW_Faces .picFace
      Case 1: .LabFace = "  Right face "
              .LabFaceAxis(0) = "Z"
              .LabFaceAxis(1) = "1.1"
              .LabFaceAxis(2) = "Y"
              .LabFaceAxis(3) = ""
              DRAW_Faces .picFace
      Case 2: .LabFace = "  Back face "
              .LabFaceAxis(0) = "1,1"
              .LabFaceAxis(1) = "X"
              .LabFaceAxis(2) = ""
              .LabFaceAxis(3) = "Y"
              DRAW_Faces .picFace
      Case 3: .LabFace = "  Left face "
              .LabFaceAxis(0) = "1,1"
              .LabFaceAxis(1) = "Z"
              .LabFaceAxis(2) = ""
              .LabFaceAxis(3) = "Y"
              DRAW_Faces .picFace
      End Select
   End With
End Sub


Private Sub Form_Unload(Cancel As Integer)
   frmFACELeft = frmFACEPresets.Left \ STX
   frmFACETop = frmFACEPresets.Top \ STY
   AfrmFACEVis = False
   Unload frmFACEPresets
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
' Made with special save
   Select Case Index
   Case 0   ' circ
      NoBoxes = 12
      pLX(1) = 192: pLY(1) = 160: pLZ(1) = 1: pHX(1) = 224: pHY(1) = 192: pHZ(1) = 16
      pLX(2) = 160: pLY(2) = 192: pLZ(2) = 1: pHX(2) = 192: pHY(2) = 224: pHZ(2) = 16
      pLX(3) = 64: pLY(3) = 33: pLZ(3) = 1: pHX(3) = 96: pHY(3) = 64: pHZ(3) = 16
      pLX(4) = 160: pLY(4) = 32: pLZ(4) = 1: pHX(4) = 192: pHY(4) = 64: pHZ(4) = 16
      pLX(5) = 192: pLY(5) = 64: pLZ(5) = 1: pHX(5) = 224: pHY(5) = 96: pHZ(5) = 16
      pLX(6) = 224: pLY(6) = 96: pLZ(6) = 1: pHX(6) = 256: pHY(6) = 160: pHZ(6) = 16
      pLX(7) = 64: pLY(7) = 192: pLZ(7) = 1: pHX(7) = 96: pHY(7) = 224: pHZ(7) = 16
      pLX(8) = 96: pLY(8) = 224: pLZ(8) = 1: pHX(8) = 160: pHY(8) = 256: pHZ(8) = 16
      pLX(9) = 96: pLY(9) = 1: pLZ(9) = 1: pHX(9) = 160: pHY(9) = 32: pHZ(9) = 16
      pLX(10) = 32: pLY(10) = 64: pLZ(10) = 1: pHX(10) = 64: pHY(10) = 96: pHZ(10) = 16
      pLX(11) = 1: pLY(11) = 96: pLZ(11) = 1: pHX(11) = 32: pHY(11) = 161: pHZ(11) = 16
      pLX(12) = 32: pLY(12) = 160: pLZ(12) = 1: pHX(12) = 64: pHY(12) = 192: pHZ(12) = 16
   Case 1   ' bridge
      NoBoxes = 7
      pLX(1) = 96: pLY(1) = 192: pLZ(1) = 1: pHX(1) = 160: pHY(1) = 226: pHZ(1) = 16
      pLX(2) = 1: pLY(2) = 2: pLZ(2) = 1: pHX(2) = 33: pHY(2) = 66: pHZ(2) = 16
      pLX(3) = 223: pLY(3) = 1: pLZ(3) = 1: pHX(3) = 256: pHY(3) = 65: pHZ(3) = 16
      pLX(4) = 30: pLY(4) = 66: pLZ(4) = 1: pHX(4) = 64: pHY(4) = 129: pHZ(4) = 16
      pLX(5) = 191: pLY(5) = 64: pLZ(5) = 1: pHX(5) = 223: pHY(5) = 129: pHZ(5) = 16
      pLX(6) = 63: pLY(6) = 129: pLZ(6) = 1: pHX(6) = 95: pHY(6) = 193: pHZ(6) = 16
      pLX(7) = 160: pLY(7) = 129: pLZ(7) = 1: pHX(7) = 193: pHY(7) = 191: pHZ(7) = 16
   Case 2   ' ramp
      NoBoxes = 7
      pLX(1) = 160: pLY(1) = 1: pLZ(1) = 1: pHX(1) = 192: pHY(1) = 192: pHZ(1) = 16
      pLX(2) = 64: pLY(2) = 1: pLZ(2) = 1: pHX(2) = 96: pHY(2) = 192: pHZ(2) = 16
      pLX(3) = 96: pLY(3) = 1: pLZ(3) = 1: pHX(3) = 160: pHY(3) = 224: pHZ(3) = 16
      pLX(4) = 192: pLY(4) = 1: pLZ(4) = 1: pHX(4) = 224: pHY(4) = 128: pHZ(4) = 16
      pLX(5) = 1: pLY(5) = 1: pLZ(5) = 1: pHX(5) = 32: pHY(5) = 64: pHZ(5) = 16
      pLX(6) = 224: pLY(6) = 1: pLZ(6) = 1: pHX(6) = 255: pHY(6) = 64: pHZ(6) = 16
      pLX(7) = 32: pLY(7) = 1: pLZ(7) = 1: pHX(7) = 64: pHY(7) = 128: pHZ(7) = 16
   Case 3   ' cross
      NoBoxes = 16
      pLX(1) = 1: pLY(1) = 224: pLZ(1) = 1: pHX(1) = 32: pHY(1) = 256: pHZ(1) = 16
      pLX(2) = 32: pLY(2) = 192: pLZ(2) = 1: pHX(2) = 64: pHY(2) = 224: pHZ(2) = 16
      pLX(3) = 64: pLY(3) = 160: pLZ(3) = 1: pHX(3) = 96: pHY(3) = 192: pHZ(3) = 16
      pLX(4) = 96: pLY(4) = 128: pLZ(4) = 1: pHX(4) = 128: pHY(4) = 160: pHZ(4) = 16
      pLX(5) = 128: pLY(5) = 96: pLZ(5) = 1: pHX(5) = 160: pHY(5) = 128: pHZ(5) = 16
      pLX(6) = 160: pLY(6) = 64: pLZ(6) = 1: pHX(6) = 192: pHY(6) = 96: pHZ(6) = 16
      pLX(7) = 192: pLY(7) = 32: pLZ(7) = 1: pHX(7) = 224: pHY(7) = 64: pHZ(7) = 16
      pLX(8) = 224: pLY(8) = 1: pLZ(8) = 1: pHX(8) = 256: pHY(8) = 33: pHZ(8) = 16
      pLX(9) = 224: pLY(9) = 224: pLZ(9) = 1: pHX(9) = 256: pHY(9) = 256: pHZ(9) = 16
      pLX(10) = 192: pLY(10) = 192: pLZ(10) = 1: pHX(10) = 224: pHY(10) = 224: pHZ(10) = 16
      pLX(11) = 160: pLY(11) = 160: pLZ(11) = 1: pHX(11) = 192: pHY(11) = 192: pHZ(11) = 16
      pLX(12) = 128: pLY(12) = 128: pLZ(12) = 1: pHX(12) = 160: pHY(12) = 160: pHZ(12) = 16
      pLX(13) = 96: pLY(13) = 96: pLZ(13) = 1: pHX(13) = 128: pHY(13) = 128: pHZ(13) = 16
      pLX(14) = 64: pLY(14) = 64: pLZ(14) = 1: pHX(14) = 96: pHY(14) = 96: pHZ(14) = 16
      pLX(15) = 32: pLY(15) = 32: pLZ(15) = 1: pHX(15) = 64: pHY(15) = 64: pHZ(15) = 16
      pLX(16) = 1: pLY(16) = 1: pLZ(16) = 1: pHX(16) = 33: pHY(16) = 33: pHZ(16) = 16
   Case 4   ' circ left
      NoBoxes = 7
      pLX(1) = 192: pLY(1) = 192: pLZ(1) = 1: pHX(1) = 255: pHY(1) = 224: pHZ(1) = 16
      pLX(2) = 127: pLY(2) = 160: pLZ(2) = 1: pHX(2) = 192: pHY(2) = 192: pHZ(2) = 16
      pLX(3) = 64: pLY(3) = 128: pLZ(3) = 1: pHX(3) = 127: pHY(3) = 160: pHZ(3) = 16
      pLX(4) = 31: pLY(4) = 96: pLZ(4) = 1: pHX(4) = 63: pHY(4) = 129: pHZ(4) = 16
      pLX(5) = 63: pLY(5) = 64: pLZ(5) = 1: pHX(5) = 127: pHY(5) = 96: pHZ(5) = 16
      pLX(6) = 127: pLY(6) = 32: pLZ(6) = 1: pHX(6) = 192: pHY(6) = 64: pHZ(6) = 16
      pLX(7) = 192: pLY(7) = 1: pLZ(7) = 1: pHX(7) = 256: pHY(7) = 32: pHZ(7) = 16
   Case 5   ' circ right
      NoBoxes = 7
      pLX(1) = 129: pLY(1) = 64: pLZ(1) = 1: pHX(1) = 193: pHY(1) = 96: pHZ(1) = 16
      pLX(2) = 64: pLY(2) = 32: pLZ(2) = 1: pHX(2) = 129: pHY(2) = 64: pHZ(2) = 16
      pLX(3) = 1: pLY(3) = 1: pLZ(3) = 1: pHX(3) = 65: pHY(3) = 32: pHZ(3) = 16
      pLX(4) = 192: pLY(4) = 95: pLZ(4) = 1: pHX(4) = 224: pHY(4) = 128: pHZ(4) = 16
      pLX(5) = 1: pLY(5) = 192: pLZ(5) = 1: pHX(5) = 64: pHY(5) = 224: pHZ(5) = 16
      pLX(6) = 63: pLY(6) = 160: pLZ(6) = 1: pHX(6) = 128: pHY(6) = 192: pHZ(6) = 16
      pLX(7) = 128: pLY(7) = 128: pLZ(7) = 1: pHX(7) = 191: pHY(7) = 160: pHZ(7) = 16
   Case 6   ' bridge left
      NoBoxes = 5
      pLX(1) = 193: pLY(1) = 194: pLZ(1) = 1: pHX(1) = 256: pHY(1) = 226: pHZ(1) = 16
      pLX(2) = 1: pLY(2) = 1: pLZ(2) = 1: pHX(2) = 31: pHY(2) = 66: pHZ(2) = 16
      pLX(3) = 31: pLY(3) = 64: pLZ(3) = 1: pHX(3) = 64: pHY(3) = 130: pHZ(3) = 16
      pLX(4) = 64: pLY(4) = 128: pLZ(4) = 1: pHX(4) = 128: pHY(4) = 161: pHZ(4) = 16
      pLX(5) = 127: pLY(5) = 160: pLZ(5) = 1: pHX(5) = 192: pHY(5) = 193: pHZ(5) = 16
   Case 7   ' bridge right
      NoBoxes = 5
      pLX(1) = 128: pLY(1) = 127: pLZ(1) = 1: pHX(1) = 192: pHY(1) = 160: pHZ(1) = 16
      pLX(2) = 63: pLY(2) = 159: pLZ(2) = 1: pHX(2) = 128: pHY(2) = 192: pHZ(2) = 16
      pLX(3) = 191: pLY(3) = 63: pLZ(3) = 1: pHX(3) = 224: pHY(3) = 129: pHZ(3) = 16
      pLX(4) = 1: pLY(4) = 192: pLZ(4) = 1: pHX(4) = 64: pHY(4) = 224: pHZ(4) = 16
      pLX(5) = 226: pLY(5) = 1: pLZ(5) = 1: pHX(5) = 256: pHY(5) = 66: pHZ(5) = 16
   Case 8   ' ramp left -------
      NoBoxes = 4
      pLX(1) = 224: pLY(1) = 1: pLZ(1) = 1: pHX(1) = 256: pHY(1) = 224: pHZ(1) = 16
      pLX(2) = 161: pLY(2) = 1: pLZ(2) = 1: pHX(2) = 223: pHY(2) = 193: pHZ(2) = 16
      pLX(3) = 96: pLY(3) = 1: pLZ(3) = 1: pHX(3) = 161: pHY(3) = 129: pHZ(3) = 16
      pLX(4) = 32: pLY(4) = 1: pLZ(4) = 1: pHX(4) = 96: pHY(4) = 65: pHZ(4) = 16
   Case 9   ' ramp right ------
      NoBoxes = 4
      pLX(1) = 1: pLY(1) = 1: pLZ(1) = 1: pHX(1) = 33: pHY(1) = 224: pHZ(1) = 16
      pLX(2) = 33: pLY(2) = 1: pLZ(2) = 1: pHX(2) = 95: pHY(2) = 193: pHZ(2) = 16
      pLX(3) = 96: pLY(3) = 1: pLZ(3) = 1: pHX(3) = 161: pHY(3) = 129: pHZ(3) = 16
      pLX(4) = 160: pLY(4) = 1: pLZ(4) = 1: pHX(4) = 224: pHY(4) = 65: pHZ(4) = 16
   Case 10  ' cross left
      NoBoxes = 7
      pLX(1) = 224: pLY(1) = 96: pLZ(1) = 1: pHX(1) = 256: pHY(1) = 129: pHZ(1) = 16
      pLX(2) = 2: pLY(2) = 192: pLZ(2) = 1: pHX(2) = 64: pHY(2) = 224: pHZ(2) = 16
      pLX(3) = 1: pLY(3) = 1: pLZ(3) = 1: pHX(3) = 65: pHY(3) = 33: pHZ(3) = 16
      pLX(4) = 65: pLY(4) = 161: pLZ(4) = 1: pHX(4) = 159: pHY(4) = 192: pHZ(4) = 16
      pLX(5) = 159: pLY(5) = 129: pLZ(5) = 1: pHX(5) = 224: pHY(5) = 161: pHZ(5) = 16
      pLX(6) = 159: pLY(6) = 64: pLZ(6) = 1: pHX(6) = 224: pHY(6) = 97: pHZ(6) = 16
      pLX(7) = 64: pLY(7) = 32: pLZ(7) = 1: pHX(7) = 159: pHY(7) = 64: pHZ(7) = 16
   Case 11  ' cross right
      NoBoxes = 7
      pLX(1) = 32: pLY(1) = 128: pLZ(1) = 1: pHX(1) = 97: pHY(1) = 160: pHZ(1) = 16
      pLX(2) = 32: pLY(2) = 63: pLZ(2) = 1: pHX(2) = 97: pHY(2) = 96: pHZ(2) = 16
      pLX(3) = 96: pLY(3) = 32: pLZ(3) = 1: pHX(3) = 191: pHY(3) = 64: pHZ(3) = 16
      pLX(4) = 96: pLY(4) = 161: pLZ(4) = 1: pHX(4) = 190: pHY(4) = 192: pHZ(4) = 16
      pLX(5) = 1: pLY(5) = 96: pLZ(5) = 1: pHX(5) = 33: pHY(5) = 129: pHZ(5) = 16
      pLX(6) = 194: pLY(6) = 192: pLZ(6) = 1: pHX(6) = 256: pHY(6) = 224: pHZ(6) = 16
      pLX(7) = 192: pLY(7) = 1: pLZ(7) = 1: pHX(7) = 256: pHY(7) = 33: pHZ(7) = 16
   End Select
End Sub


Private Sub picPDraw(Index As Integer)
Dim ixL As Long
Dim iyL As Long
Dim ixH As Long
Dim iyH As Long
   LCoords Index
   For j = 1 To NoBoxes
      ixL = pLX(j): If ixL > 1 Then ixL = ixL \ 8
      iyL = pLY(j): If iyL > 1 Then iyL = iyL \ 8
      ixH = pHX(j): If ixH > 1 Then ixH = ixH \ 8
      iyH = pHY(j): If iyH > 1 Then iyH = iyH \ 8
      picP(i).Line (ixL, iyL)-(ixH, iyH), 0, B
   Next j
End Sub

' Trebor Tnemyar


VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7680
   ClientLeft      =   105
   ClientTop       =   -180
   ClientWidth     =   11520
   DrawWidth       =   2
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   512
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   768
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picHITS 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   330
      Left            =   15
      ScaleHeight     =   330
      ScaleWidth      =   1365
      TabIndex        =   0
      Top             =   -15
      Width           =   1365
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   90
      Picture         =   "Main.frx":0442
      Stretch         =   -1  'True
      Top             =   675
      Width           =   1140
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Cuboid City by Robert Rayment Dec 2003

' COMPILE FOR SPEED !!

Option Explicit
Option Base 1

Private Type POINTAPI
        X As Long
        Y As Long
End Type
Dim PT As POINTAPI

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Const SM_CXSCREEN = 0 'X Size of screen
Const SM_CYSCREEN = 1 'Y Size of Screen
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function GetAsyncKeyState Lib "user32" _
   (ByVal vKey As KeyCodeConstants) As Long

Dim aDone As Boolean
Dim ScrWidth As Long
Dim ScrHeight As Long

Dim xfrm As Single
Dim yfrm As Single
Dim FrmLeft As Long
Dim FrmTop As Long


Private Sub Form_Initialize()
   
   If App.PrevInstance Then End

   STX = Screen.TwipsPerPixelX
   STY = Screen.TwipsPerPixelY
   
   FillPalette
   
   'Drawing array
   BArrayWidth = 3 * 256
   BArrayHeight = 9 * 256
   
   FillBMPStruc BArrayWidth, BArrayHeight
    
   BackArrayWidth = BArrayWidth
   BackArrayHeight = 512
   
   WindowWidth = 768
   WindowHeight = 512
   
   MakeBackArray
   
   PlaneZ = 0   ' Z View plane
   
   eyeX = 256 + 128
   eyeY = 512 '650
   eyeZ = -512
   
   SumStepX = 0
   SumStepZ = 0
   StepX = 0
   StepY = 0
   StepZ = 0
   
   XIncr = 2
   ZIncr = 2
   EyeIncr = 4
   
   MaxStepX = 128 '32
   Max_eyeY = 2048
   Min_eyeY = 128
   MaxStepZ = 128 '64
   
   AMAZE = False
   HITS = 0
End Sub

Private Sub Form_Load()
Dim fso As New Scripting.FileSystemObject

   PathSpec$ = App.Path
   If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"
   
   ' Check if CCC_Files exists else App.Path
   i = InStrRev(PathSpec$, "\", Len(PathSpec$) - 1)
   CCC_Path$ = Mid$(PathSpec$, 1, i) + "CCC_Files\"
   If Not fso.FolderExists(CCC_Path$) Then
      CCC_Path$ = PathSpec$
   End If
   
   With Me
      .Width = WindowWidth * Screen.TwipsPerPixelX
      .Height = WindowHeight * Screen.TwipsPerPixelY
      .Show
      .Cls
   End With
   ' Collision image
   Image1.Visible = False
   
   frmStart.Show 1
   
   If Len(FileSpec$) = 0 Then Make_Locate_Boxes
   
   FillBoxes

   ReDim transx(8, NumBoxes)
   ReDim transy(8, NumBoxes)
   Transform

   'ReDim BArray(BArrayWidth, BArrayHeight)
   DrawOnBArray
   
   ' OK in IDE but Flickers in EXE ??
   Screen.MousePointer = vbCustom
   Screen.MouseIcon = LoadResPicture(101, vbResCursor)
   
   ACTION
End Sub


'#### ACTION ###################################################

Private Sub ACTION()
Dim mx As Long
Dim mz As Long
Dim divx As Long
Dim divz As Long
   
Dim mpx As Long
Dim mpy As Long

'Dim vKey As KeyCodeConstants

   ' For calculation cursor offset
   divx = ScrWidth \ (2 * MaxStepX)
   divz = ScrHeight \ (2 * MaxStepZ)
   picHITS.Cls
   picHITS.Print "0"
   
   aDone = False
   Do
      
      ' Check for Quit
      If GetAsyncKeyState(vbKeyEscape) And &H8000 Then
         aDone = True
         Exit Do
      End If
      
      ' Check for screen res change
      If GetSystemMetrics(SM_CXSCREEN) <> ScrWidth Then Form_Resize
      
      ' Check for key F1 or Right click - Show Start Window
      If GetAsyncKeyState(vbKeyF1) And &H8000 _
         Or GetAsyncKeyState(vbRightButton) And &H8000 Then
         
         frmStart.Show 1
         RESET_ALL
      
      ' Eye height change
      ElseIf GetAsyncKeyState(vbKeyControl) And &H8000 Then
         ' Shift eye height
         If GetAsyncKeyState(vbKeyUp) And &H8000 Then
            eyeY = eyeY - EyeIncr   ' Decrease eye height
            If eyeY < Min_eyeY Then eyeY = Min_eyeY
         ElseIf GetAsyncKeyState(vbKeyDown) And &H8000 Then
            eyeY = eyeY + EyeIncr   ' Increase eye height
            If eyeY > Max_eyeY Then eyeY = Max_eyeY
         End If
      ' Block height change
      ElseIf GetAsyncKeyState(vbKeyShift) And &H8000 Then
         If GetAsyncKeyState(vbKeyUp) And &H8000 Then
            StepY = StepY + 4 ' Increase height
            If StepY <= 252 Then IncY 4 Else StepY = 256
         ElseIf GetAsyncKeyState(vbKeyDown) And &H8000 Then
            StepY = StepY - 4 ' Decrease heights
            If StepY >= 0 Then IncY -4 Else StepY = 0
         End If
      End If
      
      ' Get cursor offset from screen center
      ' to set step size   StepX or StepZ
      If GetAsyncKeyState(vbKeyControl) And &H8000 Then  ' Change X only
         GetCursorPos PT
         mx = PT.X - ScrWidth \ 2
         StepX = -mx \ divx
         StepZ = 0
      ElseIf GetAsyncKeyState(vbKeyShift) And &H8000 Then   ' Change Z only
         GetCursorPos PT
         mz = PT.Y - ScrHeight \ 2
         StepZ = mz \ divz
         StepX = 0
      Else
         GetCursorPos PT
         mx = PT.X - ScrWidth \ 2
         StepX = -mx \ divx
         mz = PT.Y - ScrHeight \ 2
         StepZ = mz \ divz
      End If
      
      IncrZ ' by StepZ
      IncrX ' by StepX
      Transform
      DrawOnBArray
      
      '----------------------------------------------
      ' Detecting collisions
      ' Get Form coords
      GetCursorPos PT
      mpx = PT.X - FrmLeft
      mpy = 512 - (PT.Y - FrmTop)  ' 512 is forms pixel height

      For i = 1 To NumBoxes
          
          ' Front face detection only
'         If mpx > transx(1, i) Then ' >= etc will detect flat
'         If mpx < transx(2, i) Then '    surfaces
'         If mpy > transy(1, i) Then ' Detection depends on
'         If mpy < transy(3, i) Then ' ss 1 <= ss 2 <= ss 3
'OR
         ' Detect on diagonal, but can pick lower blocks before
         ' upper blocks
         If transx(7, i) > transx(1, i) Then
            If mpx > transx(1, i) Then
            If mpx < transx(7, i) Then
            If mpy > transy(1, i) Then
            If mpy < transy(7, i) Then
               ' Avoid collisions on ground plates of thickness 1
               If BoxY(4, i) - BoxY(1, i) > 1 Then
                  PROCHITS
                  Exit For
               End If
            End If
            End If
            End If
            End If
         Else
            If mpx < transx(1, i) Then
            If mpx > transx(7, i) Then
            If mpy > transy(1, i) Then
            If mpy < transy(7, i) Then
               ' Avoid collisions on ground plates of thickness 1
               If BoxY(4, i) - BoxY(1, i) > 1 Then
                  PROCHITS
                  Exit For
               End If
            End If
            End If
            End If
            End If
         End If
      Next i
      '----------------------------------------------
      
      'Display
      'Blit Stretch byte-array to Form
      StretchDIBits Me.hdc, 0, 0, WindowWidth, WindowHeight, _
      0, 0, WindowWidth, WindowHeight, _
      BArray(1, 1), bm, DIB_RGB_COLORS, vbSrcCopy
   
      Me.Refresh
      
      Image1.Visible = False
      
      'DoEvents ' To allow mouse to move form, arbitrary
   Loop Until aDone

   Screen.MousePointer = vbDefault
   Erase BArray, BackArray
   Unload Me
   End

End Sub

Private Sub PROCHITS()
   With Image1
      If transx(7, i) > transx(1, i) Then
         .Left = transx(1, i)
      Else
         .Left = transx(7, i)
      End If
      .Top = 512 - transy(3, i)
      .Width = Abs(transx(2, i) - transx(1, i))
      .Height = transy(3, i) - transy(1, i)
      .Visible = True
      HITS = HITS + 1
      picHITS.Cls
      picHITS.Print HITS;
   End With
   'Label1 = Str$(i)   ' Test
   If AMAZE And i = PyramidBlockNumber And HITS = 1 Then
      CLEARALL
      frmStart.Show 1
      RESET_ALL
   End If
End Sub

Private Sub RESET_ALL()
   If Len(FileSpec$) = 0 Then Make_Locate_Boxes ' Default
   FillBoxes
   SumStepX = 0
   SumStepZ = 0
   StepX = 0
   StepY = 0
   StepZ = 0
   ReDim transx(8, NumBoxes)
   ReDim transy(8, NumBoxes)
   Transform
   DrawOnBArray
   StepY = 0
   ' Not colored but OK in IDE but,
   ' though colored, can Flicker in EXE ??
   Screen.MousePointer = vbCustom
   Screen.MouseIcon = LoadResPicture(101, vbResCursor)
End Sub

Private Sub CLEARALL()
Dim iL As Long
Dim iT As Long
Dim iW As Long
Dim iH As Long
Dim icx As Long
Dim icy As Long

   icx = Me.Width / (2 * STX)
   icy = Me.Height / (2 * STX)
   iL = icx - 2
   iT = icy - 2
   iW = 4
   iH = 4
   Image1.Visible = True
   
   'Cls
   
   For i = 1 To icy \ 2 Step 8
      With Image1
         .Left = iL
         .Top = iT
         .Width = iW
         .Height = iH
      End With
      iL = iL - icy \ 8
      iT = iT - icx \ 8
      iW = iW + 2 * icy \ 8
      iH = iH + 2 * icx \ 8
      Refresh
   Next i
End Sub
'#### END ACTION ################################################


'#### MOVE FORM #################################################
' Not used

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   xfrm = X
   yfrm = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim FTop As Long
Dim FLeft As Long

   If Button = vbLeftButton Then
      ScrWidth = GetSystemMetrics(SM_CXSCREEN)
      ScrHeight = GetSystemMetrics(SM_CYSCREEN)
   
      FTop = Me.Top + Y - yfrm
      FLeft = Me.Left + X - xfrm
      
      ' Ensure form stays on screen
      If FTop < 0 Then
         FTop = 1
         Y = yfrm
      End If
      If FLeft < 0 Then
         FLeft = 1
         X = xfrm
      End If
      If FTop > ScrHeight * Screen.TwipsPerPixelY - Me.Height Then
         FTop = ScrHeight * Screen.TwipsPerPixelY - Me.Height - 1
         Y = yfrm
      End If
      If FLeft > ScrWidth * Screen.TwipsPerPixelX - Me.Width Then
         FLeft = ScrWidth * Screen.TwipsPerPixelX - Me.Width - 1
         X = xfrm
      End If
      
      Me.Top = FTop + Y - yfrm
      Me.Left = FLeft + X - xfrm
   End If
End Sub

Private Sub Form_Resize()
   ScrWidth = GetSystemMetrics(SM_CXSCREEN)
   ScrHeight = GetSystemMetrics(SM_CYSCREEN)
   Me.Left = (ScrWidth * STX - Me.Width) \ 2
   Me.Top = (ScrHeight * STY - Me.Height) \ 2

   FrmLeft = Me.Left \ STX
   FrmTop = Me.Top \ STY
   
   Show
End Sub

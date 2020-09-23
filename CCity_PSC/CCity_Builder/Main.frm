VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   " Cuboid City Coords"
   ClientHeight    =   8295
   ClientLeft      =   165
   ClientTop       =   -180
   ClientWidth     =   11280
   DrawWidth       =   2
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   553
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   752
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDefaults 
      Caption         =   "Default face dimensions"
      Height          =   1800
      Left            =   6255
      TabIndex        =   53
      Top             =   75
      Width           =   2625
      Begin VB.CheckBox chkRand 
         Caption         =   "Random heights"
         Height          =   270
         Left            =   300
         TabIndex        =   62
         Top             =   975
         Width           =   1530
      End
      Begin VB.CommandButton cmdCloseDefault 
         Caption         =   "Close"
         Height          =   285
         Left            =   855
         TabIndex        =   60
         Top             =   1365
         Width           =   1020
      End
      Begin VB.HScrollBar HSDefault 
         Height          =   240
         Index           =   1
         Left            =   1740
         Max             =   32
         TabIndex        =   59
         Top             =   615
         Width           =   510
      End
      Begin VB.HScrollBar HSDefault 
         Height          =   240
         Index           =   0
         LargeChange     =   16
         Left            =   1545
         Max             =   256
         TabIndex        =   58
         Top             =   285
         Width           =   855
      End
      Begin VB.TextBox txtDefault 
         Height          =   285
         Index           =   1
         Left            =   675
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   "Text1"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtDefault 
         Height          =   285
         Index           =   0
         Left            =   660
         Locked          =   -1  'True
         TabIndex        =   54
         Text            =   "Text1"
         Top             =   255
         Width           =   735
      End
      Begin VB.Label LabDefault 
         Caption         =   "Depth"
         Height          =   270
         Index           =   1
         Left            =   135
         TabIndex        =   56
         Top             =   630
         Width           =   450
      End
      Begin VB.Label LabDefault 
         Caption         =   "Height"
         Height          =   270
         Index           =   0
         Left            =   120
         TabIndex        =   55
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3345
      Left            =   105
      TabIndex        =   30
      Top             =   15
      Width           =   2175
      Begin VB.CommandButton cmdLoadSave 
         Caption         =   "Add"
         Height          =   285
         Index           =   1
         Left            =   1140
         TabIndex        =   69
         Top             =   555
         Width           =   885
      End
      Begin VB.CommandButton cmdPresets 
         Caption         =   "Preset FACE shapes"
         Height          =   240
         Index           =   1
         Left            =   210
         TabIndex        =   63
         Top             =   2985
         Width           =   1740
      End
      Begin VB.CommandButton cmdPresets 
         Caption         =   "Preset PLAN shapes"
         Height          =   240
         Index           =   0
         Left            =   210
         TabIndex        =   61
         Top             =   2655
         Width           =   1740
      End
      Begin VB.CommandButton cmdDefaults 
         Caption         =   "Default Height, Depth"
         Height          =   285
         Left            =   210
         TabIndex        =   52
         Top             =   2265
         Width           =   1740
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Height          =   285
         Left            =   165
         TabIndex        =   47
         Top             =   195
         Width           =   885
      End
      Begin VB.Frame Frame2 
         Caption         =   "Actions on RC-square"
         Height          =   1215
         Left            =   150
         TabIndex        =   39
         Top             =   975
         Width           =   1830
         Begin VB.CheckBox chkCopyPaste 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Paste all to RC-square"
            Height          =   270
            Index           =   1
            Left            =   75
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   840
            Width           =   1710
         End
         Begin VB.CheckBox chkCopyPaste 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Copy all in RC-square"
            Height          =   285
            Index           =   0
            Left            =   75
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   480
            Width           =   1710
         End
         Begin VB.Label LabCol 
            Caption         =   "Col ="
            Height          =   240
            Left            =   975
            TabIndex        =   41
            Top             =   225
            Width           =   705
         End
         Begin VB.Label LabRow 
            Caption         =   "Row ="
            Height          =   240
            Left            =   75
            TabIndex        =   40
            Top             =   240
            Width           =   765
         End
      End
      Begin VB.CommandButton cmdLoadSave 
         Caption         =   "Save As"
         Height          =   285
         Index           =   2
         Left            =   165
         TabIndex        =   37
         Top             =   555
         Width           =   885
      End
      Begin VB.CommandButton cmdLoadSave 
         Caption         =   "Load"
         Height          =   285
         Index           =   0
         Left            =   1140
         TabIndex        =   36
         Top             =   195
         Width           =   885
      End
   End
   Begin VB.Frame fraRC 
      Caption         =   " [1]  SELECT SQUARE"
      Height          =   3345
      Left            =   2265
      TabIndex        =   0
      Top             =   15
      Width           =   2430
      Begin VB.PictureBox picRC 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   2520
         Left            =   210
         ScaleHeight     =   168
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   123
         TabIndex        =   1
         Top             =   525
         Width           =   1845
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "X"
         Height          =   225
         Index           =   1
         Left            =   1920
         TabIndex        =   23
         Top             =   3045
         Width           =   210
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Z"
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   22
         Top             =   510
         Width           =   210
      End
      Begin VB.Label LabRCHelp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "?"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   2175
         TabIndex        =   5
         ToolTipText     =   "Left/Right button Select(Red) / Deselect "
         Top             =   525
         Width           =   195
      End
      Begin VB.Label LabRC 
         Caption         =   "R =,  C ="
         Height          =   270
         Left            =   210
         TabIndex        =   4
         Top             =   255
         Width           =   2115
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   210
         Index           =   1
         Left            =   690
         TabIndex        =   3
         Top             =   3045
         Width           =   195
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   210
         Index           =   0
         Left            =   30
         TabIndex        =   2
         Top             =   2685
         Width           =   195
      End
   End
   Begin VB.Frame fraPlan 
      Caption         =   "[2]  PLAN FOR SELECTED SQUARE"
      Height          =   4875
      Left            =   90
      TabIndex        =   6
      Top             =   3360
      Width           =   4605
      Begin VB.OptionButton optDrawResize 
         Caption         =   "ALTER BLOCKS"
         Height          =   225
         Index           =   1
         Left            =   2610
         TabIndex        =   45
         Top             =   4560
         Width           =   1590
      End
      Begin VB.OptionButton optDrawResize 
         Caption         =   "DRAW"
         Height          =   225
         Index           =   0
         Left            =   1665
         TabIndex        =   44
         Top             =   4560
         Value           =   -1  'True
         Width           =   810
      End
      Begin VB.PictureBox picPLAN 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   3975
         Left            =   315
         ScaleHeight     =   261
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   256
         TabIndex        =   7
         Top             =   360
         Width           =   3900
         Begin VB.Line LinPlanGrid 
            BorderColor     =   &H00E0E0E0&
            BorderStyle     =   3  'Dot
            Index           =   0
            X1              =   27
            X2              =   27
            Y1              =   20
            Y2              =   60
         End
      End
      Begin VB.Label LabRCHelp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "?"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   4305
         TabIndex        =   67
         ToolTipText     =   "Ctrl-Left Button on circle to select block "
         Top             =   375
         Width           =   195
      End
      Begin VB.Label LabN 
         Caption         =   "N = 99"
         Height          =   240
         Left            =   3645
         TabIndex        =   25
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "1    R   i   g   h  t"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   6
         Left            =   4335
         TabIndex        =   15
         Top             =   1620
         Width           =   120
      End
      Begin VB.Label Label2 
         Caption         =   "3    L e  f   t"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Index           =   5
         Left            =   150
         TabIndex        =   14
         Top             =   1680
         Width           =   120
      End
      Begin VB.Label Label2 
         Caption         =   "2  Back"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   4
         Left            =   1815
         TabIndex        =   13
         Top             =   195
         Width           =   585
      End
      Begin VB.Label Label2 
         Caption         =   "0  Front"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   3
         Left            =   1800
         TabIndex        =   12
         Top             =   4365
         Width           =   555
      End
      Begin VB.Label Label2 
         Caption         =   "1,1"
         Height          =   240
         Index           =   2
         Left            =   75
         TabIndex        =   11
         Top             =   4305
         Width           =   255
      End
      Begin VB.Label LabXYPLAN 
         Caption         =   "X =, Y ="
         Height          =   255
         Left            =   315
         TabIndex        =   10
         Top             =   4545
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   " Z "
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   9
         Top             =   330
         Width           =   210
      End
      Begin VB.Label Label2 
         Caption         =   "X"
         Height          =   195
         Index           =   0
         Left            =   4095
         TabIndex        =   8
         Top             =   4380
         Width           =   180
      End
   End
   Begin VB.PictureBox picViewsContainer 
      Height          =   8235
      Left            =   4695
      ScaleHeight     =   545
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   430
      TabIndex        =   16
      Top             =   0
      Width           =   6510
      Begin VB.HScrollBar HSPerspec 
         Height          =   135
         LargeChange     =   8
         Left            =   765
         Max             =   256
         TabIndex        =   70
         Top             =   7995
         Width           =   3225
      End
      Begin VB.CommandButton cmdSpecial 
         Caption         =   "S"
         Height          =   195
         Left            =   6135
         TabIndex        =   65
         ToolTipText     =   "Save with variable description "
         Top             =   7875
         Width           =   195
      End
      Begin VB.CheckBox chkGrid 
         Caption         =   "Grid"
         Height          =   270
         Left            =   4695
         TabIndex        =   50
         Top             =   7290
         Width           =   690
      End
      Begin VB.CheckBox chkPLAN 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Copy && paste red block"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   570
         Index           =   3
         Left            =   4605
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   5925
         Width           =   1650
      End
      Begin VB.CheckBox chkPLAN 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cycle thru blocks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   540
         Index           =   2
         Left            =   4605
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   5280
         Width           =   1650
      End
      Begin VB.CheckBox chkPLAN 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Undo red block"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   525
         Index           =   1
         Left            =   4605
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   4710
         Width           =   1650
      End
      Begin VB.CheckBox chkPLAN 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Clear blocks in RC-square"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   570
         Index           =   0
         Left            =   4605
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   4095
         Width           =   1650
      End
      Begin VB.OptionButton optHtRec 
         Caption         =   "DRAW"
         Height          =   195
         Index           =   1
         Left            =   4695
         TabIndex        =   32
         Top             =   2070
         Width           =   1140
      End
      Begin VB.OptionButton optHtRec 
         Caption         =   "ALTER BLOCKS"
         Height          =   195
         Index           =   0
         Left            =   4680
         TabIndex        =   31
         Top             =   1770
         Value           =   -1  'True
         Width           =   1545
      End
      Begin VB.CommandButton cmdFaces 
         Caption         =   ">"
         Height          =   240
         Index           =   1
         Left            =   5610
         TabIndex        =   21
         Top             =   1155
         Width           =   360
      End
      Begin VB.CommandButton cmdFaces 
         Caption         =   "<"
         Height          =   240
         Index           =   0
         Left            =   4860
         TabIndex        =   20
         Top             =   1155
         Width           =   360
      End
      Begin VB.PictureBox picFace 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   7740
         Left            =   330
         ScaleHeight     =   512
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   256
         TabIndex        =   18
         Top             =   225
         Width           =   3900
         Begin VB.Line LinFaceGrid 
            BorderColor     =   &H00E0E0E0&
            BorderStyle     =   3  'Dot
            Index           =   0
            X1              =   45
            X2              =   45
            Y1              =   62
            Y2              =   96
         End
      End
      Begin VB.Label LabRCHelp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "?"
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   5145
         TabIndex        =   68
         ToolTipText     =   "Also Ctrl-Left Button on circle to select block "
         Top             =   3705
         Width           =   195
      End
      Begin VB.Label LabBlocksInRC 
         Caption         =   "Blocks in RC = 0"
         Height          =   240
         Left            =   4605
         TabIndex        =   66
         Top             =   7065
         Width           =   1695
      End
      Begin VB.Shape Shape4 
         Height          =   3705
         Left            =   4485
         Top             =   3990
         Width           =   1890
      End
      Begin VB.Label LabBlockNum 
         Caption         =   "BlockNum = 0"
         Height          =   240
         Left            =   4605
         TabIndex        =   64
         Top             =   6840
         Width           =   1710
      End
      Begin VB.Shape Shape3 
         Height          =   2700
         Left            =   4485
         Top             =   75
         Width           =   1890
      End
      Begin VB.Label LabTest 
         Caption         =   "LabTest"
         Height          =   225
         Left            =   4635
         TabIndex        =   51
         Top             =   7860
         Width           =   1710
      End
      Begin VB.Label LabNumBlocks 
         Caption         =   "NumBlocks = 0"
         Height          =   240
         Left            =   4605
         TabIndex        =   48
         Top             =   6600
         Width           =   1695
      End
      Begin VB.Label LabFaceCap 
         Alignment       =   2  'Center
         Caption         =   "Top"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1485
         TabIndex        =   46
         Top             =   15
         Width           =   1290
      End
      Begin VB.Label Label5 
         Caption         =   "Altering red blocks:     Left button down to change lengths, Right button to move block."
         Height          =   1080
         Left            =   4620
         TabIndex        =   38
         Top             =   2895
         Width           =   1545
      End
      Begin VB.Shape Shape2 
         Height          =   825
         Left            =   4575
         Top             =   1590
         Width           =   1695
      End
      Begin VB.Shape Shape1 
         Height          =   870
         Left            =   4575
         Top             =   690
         Width           =   1695
      End
      Begin VB.Label LabXYFACE 
         Caption         =   "X =, Y ="
         Height          =   225
         Left            =   4620
         TabIndex        =   29
         Top             =   2475
         Width           =   1560
      End
      Begin VB.Label LabFaceAxis 
         Alignment       =   2  'Center
         Caption         =   "Y"
         Height          =   225
         Index           =   3
         Left            =   4290
         TabIndex        =   28
         Top             =   210
         Width           =   210
      End
      Begin VB.Label LabFaceAxis 
         Alignment       =   2  'Center
         Caption         =   "Y"
         Height          =   225
         Index           =   2
         Left            =   105
         TabIndex        =   27
         Top             =   210
         Width           =   165
      End
      Begin VB.Label LabFaceAxis 
         Alignment       =   2  'Center
         Caption         =   "X"
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   26
         Top             =   7785
         Width           =   210
      End
      Begin VB.Label LabFaceAxis 
         Alignment       =   2  'Center
         Caption         =   "X"
         Height          =   225
         Index           =   0
         Left            =   4305
         TabIndex        =   24
         Top             =   7800
         Width           =   210
      End
      Begin VB.Label LabViews 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   " [3]  FACES"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   4875
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.Label LabFace 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "  Front face"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4605
         TabIndex        =   17
         Top             =   795
         Width           =   1605
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Cuboid City Builder by Robert Rayment  Dec 2003

Option Explicit
Option Base 1

Private Declare Function SetCursorPos Lib "user32" _
(ByVal X As Long, ByVal Y As Long) As Long

Dim AGRID As Boolean ' Plan & Face grid
Dim aPPMD As Boolean ' To confirm picPLAN_MouseDown activated
Dim aPFMD As Boolean ' To confirm picFace_MouseDown activated
Dim aLinLoaded As Boolean  ' To confirm grid lines Loaded

' For Copy & Paste blocks
Dim CopyR As Long
Dim CopyC As Long

' For Resizing & moving blocks
Dim dX As Long
Dim dY As Long
Dim dZ As Long
Dim picFaceX As Long
Dim picFaceY As Long
Dim picFaceZ As Long
'------------------------------
' For moving fraDefaults
Dim fraX  As Single
Dim fraY As Single

Private CommonDialog1 As OSDialog


Private Sub cmdSpecial_Click()
' SPECIAL SAVING includes variable descriptions LX() etc
Dim Title$, Filt$, InDir$
   Set CommonDialog1 = New OSDialog
      If NumBlocks < 1 Then
         MsgBox " No blocks to save", vbInformation, "Saving ccs file"
         Exit Sub
      End If
      Title$ = "Save Special Cuboid File ccs"
      Filt$ = "Save ccs (*.ccs)|*.ccs"
      InDir$ = CCC_Path$
      CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, "", Me.hWnd
   
      ' Offset cursor to avoid click thru
      SetCursorPos Me.Left \ STX + 150, Me.Top \ STY + 80
      
      If Len(FileSpec$) <> 0 Then
         
         CCC_Path$ = FileSpec$
         
         Screen.MousePointer = vbHourglass
         FixExtension FileSpec$, "ccs"
         CCC_Path$ = FileSpec$
         
         SAVE_CCS_FILE
         
         Screen.MousePointer = vbDefault
      
      End If
   Set CommonDialog1 = Nothing

End Sub


Private Sub cmdNew_Click()
   If NumBlocks <> 0 Then
      res = MsgBox("Clear all" & Str$(NumBlocks) & " blocks", vbQuestion + vbYesNo, "Clear all blocks")
      If res = vbNo Then Exit Sub
   End If
   ' Clear all pics
   picPLAN.Cls
   picFace.Cls
   picViewsContainer.Cls
   LabN = "0"
   Unload frmPLANPreSets
   AfrmFACEVis = False
   Unload frmFACEPresets
  
   Form_Load
End Sub

Private Sub cmdPresets_Click(Index As Integer)
   If Index = 0 Then
      AfrmFACEVis = False
      Unload frmFACEPresets
      frmPLANPreSets.Show 0
   Else
      Unload frmPLANPreSets
      AfrmFACEVis = True
      frmFACEPresets.Show 0
   End If
End Sub

Private Sub chkRand_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARandHts = Not ARandHts
End Sub


Private Sub Form_Initialize()
   aLinLoaded = False
   
   ' Preset forms location
   frmPLANLeft = 700
   frmPLANTop = 100
   frmFACELeft = 720
   frmFACETop = 140

   optDrawResize(0).Value = True
   PlanAction = True    ' Draw plan
   
   optHtRec(0).Value = True
   FaceAction = True    ' Alter faces

   eyeX = 128
End Sub

Private Sub Form_Load()
Dim fso As New Scripting.FileSystemObject

   LabTest.Visible = False
   
   PathSpec$ = App.Path
   If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"
   
   ' Check if CCC_Files exists else App.Path
   i = InStrRev(PathSpec$, "\", Len(PathSpec$) - 1)
   CCC_Path$ = Mid$(PathSpec$, 1, i) + "CCC_Files\"
   If Not fso.FolderExists(CCC_Path$) Then
      CCC_Path$ = PathSpec$
   End If
   


' Mark RC-square
' Sub-block coords in RC-square
ReDimCoords

' Num of sub-blocks in RC-square
ReDim NumBlocksInRC(-2 To 21, -5 To 12)

   NumBlocks = 0

   STX = Screen.TwipsPerPixelX
   STY = Screen.TwipsPerPixelY
   
   SelectCul = RGB(255, 80, 80)
   GridBackColor = RGB(0, 0, 200)
   
   LOCATE_CONTROLS
   
   picFaceBacCul = picFace.Point(1, 1)
   picPLANCul = picPLAN.Point(1, 1)
   
   RedBlockNumber = 0
   LabBlockNum = "BlockNum =" & Str$(RedBlockNumber)
   LabBlocksInRC = "Blocks in RC =" & Str$(NumBlocksInRC(R, C))
   
   ' Default to Front face
   FaceIndex = 1
   cmdFaces_Click 0
   
   FaceAction = True ' Face Heights
   PlanAction = True ' Draw rectangles
   aPPMD = False
   aPFMD = False
   
   FalseALL
   
   PlanAction = optDrawResize(0).Value
   FaceAction = optHtRec(0).Value
   
   AGRID = True

   ' Init default height
   HSDefault(0).Value = 8
   txtDefault(0).Text = 8
   DummyHeight = 8
   ' Init default depth
   HSDefault(1).Value = 4
   txtDefault(1).Text = 4
   DummyDepth = 4
   ARandHts = False
   
   fraDefaults.Visible = False
   AfrmFACEVis = False
   
   HSPerspec.Visible = False

   
End Sub


Private Sub HSPerspec_Change()
If FaceIndex = 5 Then
   eyeX = HSPerspec.Value
   DRAW_Faces picFace
End If
End Sub

'#### [1] SELECT/DESELECT RC-SQUARE ###################################################

Private Sub picRC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim RR As Long, CC As Long
   RR = 21 - Y \ GridStep
   CC = X \ GridStep - 5
   LabRC = "R = " & Str$(RR) & "  C = " & Str$(CC) & "  N =" & Str$(NumBlocksInRC(RR, CC))
End Sub

Private Sub picRC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   ' Calc Row, Col
   R = 21 - Y \ GridStep
   C = X \ GridStep - 5
   
   ' Set all RC to white or black
   For j = -2 To 21  ' rows
   For i = -5 To 12  ' cols
      iy = GridStep * (21 - j)
      ix = GridStep * (i + 5)
      If NumBlocksInRC(j, i) > 0 Then
         picRC.Line (ix + 1, iy + 1)-(ix + GridStep - 1, iy + GridStep - 1), vbWhite, BF
      Else
         picRC.Line (ix + 1, iy + 1)-(ix + GridStep - 1, iy + GridStep - 1), GridBackColor, BF
      End If
   Next i
   Next j
   
   iy = GridStep * (21 - R)
   ix = GridStep * (C + 5)
   
   ' Clear all pics
   picPLAN.Cls
   If FaceIndex <> 4 Then picFace.Cls
   picViewsContainer.Cls
      
   If Button = vbLeftButton Then ' Select RC
      picRC.Line (ix + 1, iy + 1)-(ix + GridStep - 1, iy + GridStep - 1), SelectCul, BF
      
      RC_GO
      
      LabRow = " Row =" & Str$(R)
      LabCol = "Col =" & Str$(C)
      
   ElseIf Button = vbRightButton Then  ' De-select RC
      
      chkPLAN_MouseUp 0, 0, 0, 0, 0
   
   End If

   LabRC = "R = " & Str$(R) & "  C = " & Str$(C) & "  N =" & Str$(NumBlocksInRC(R, C))
   If NumBlocksInRC(R, C) = 0 Then
      LabBlockNum = "BlockNum = 0"
      LabBlocksInRC = "Blocks in RC = 0"
   End If
End Sub

Private Sub chkCopyPaste_Click(Index As Integer)
Exit Sub
End Sub

Private Sub chkCopyPaste_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   chkCopyPaste(Index).Value = 0
   
   Select Case Index
   Case 0   ' Copy from RC-square
      If NumBlocksInRC(R, C) = 0 Then
         MsgBox " No blocks here", vbInformation, "Copy Paste"
         Exit Sub
      End If
      ' Dummy ccords
      CopyR = R
      CopyC = C
   Case 1   ' Paste to RC-square
      ' Diff RC to Copy RC
      If R = CopyR And C = CopyC Then
         MsgBox " No pasting to same RC-square", vbInformation, "Pasting"
         Exit Sub
      End If
      ' Paste CopyR,CopyC to R,C
      BlockCounter = 0
      For i = 1 To NumBlocks
         If NR(i) = CopyR And NC(i) = CopyC Then
            BlockCounter = BlockCounter + 1
         End If
      Next i
      j = NumBlocks + 1
      NumBlocks = NumBlocks + BlockCounter
      
      If BlockCounter = 0 Then
         MsgBox " No blocks to paste", vbInformation, "Copy Paste"
         Exit Sub
      End If
      
      ReDimPreserve
      For i = 1 To NumBlocks
         If NR(i) = CopyR And NC(i) = CopyC Then
            LX(j) = LX(i)
            LZ(j) = LZ(i)
            LY(j) = LY(i)
            HX(j) = HX(i)
            HZ(j) = HZ(i)
            HY(j) = HY(i)
            NR(j) = R
            NC(j) = C
            j = j + 1
         End If
      Next i
      NumBlocksInRC(R, C) = NumBlocksInRC(R, C) + BlockCounter
      
      LabNumBlocks = "NumBlocks =" & Str$(NumBlocks)
      
      RC_GO
   
   End Select
End Sub

Private Sub RC_GO()
   For i = 0 To 3
      chkPLAN(i).ForeColor = SelectCul
   Next i
   chkCopyPaste(0).ForeColor = SelectCul
   chkCopyPaste(1).ForeColor = SelectCul
   
   LabRC = "R = " & Str$(R) & "  C = " & Str$(C) '& "  N =" & Str$(NumBlocksInRC(R, C))
   
   ' True all
   fraPlan.Enabled = True
   picFace.Enabled = True
   picViewsContainer.Enabled = True
   LabViews.Enabled = True
   chkCopyPaste(0).Enabled = True
   chkCopyPaste(1).Enabled = True
   For i = 0 To 3
      chkPLAN(i).Enabled = True
   Next i
   cmdPresets(0).Enabled = True
   cmdPresets(1).Enabled = True


   LabN = "N =" & Str$(NumBlocksInRC(R, C))
   If NumBlocksInRC(R, C) > 0 Then
      ' Find block number of location of current
      ' R,C selection & set RedBlockNumber
      RedBlockNumber = 0
      LabBlockNum = "BlockNum =" & Str$(RedBlockNumber)
      LabBlocksInRC = "Blocks in RC =" & Str$(NumBlocksInRC(R, C))
      BlockCounter = 0
      For i = 1 To NumBlocks
         If NR(i) = R And NC(i) = C Then
            RedBlockNumber = i
            BlockCounter = BlockCounter + 1
            If BlockCounter = NumBlocksInRC(R, C) Then Exit For
         End If
      Next i
      
      If RedBlockNumber = 0 Then FalseALL
      LabBlockNum = "BlockNum =" & Str$(RedBlockNumber)
      LabBlocksInRC = "Blocks in RC =" & Str$(NumBlocksInRC(R, C))
      REDRAW_PLAN picPLAN
      DRAW_Faces picFace
   Else  ' NumBlocksInRC(R, C) = 0
      LabBlocksInRC = "Blocks in RC =" & Str$(NumBlocksInRC(R, C))
   End If
End Sub
'#### END [1] SELECT/DESELECT RC-SQUARE ###################################################

'#### CLEAR, UNDO & CYCLE #################################################################

Private Sub chkPLAN_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim iRB As Long

chkPLAN(Index).Value = 0
   
   iy = GridStep * (21 - R)
   ix = GridStep * (C + 5)
   
   If NumBlocksInRC(R, C) = 0 Then
      picRC.Line (ix + 1, iy + 1)-(ix + GridStep - 1, iy + GridStep - 1), GridBackColor, BF
      LabN = "N =" & Str$(NumBlocksInRC(R, C))
      ' Clear all pics
      picPLAN.Cls
      picFace.Cls
      picViewsContainer.Cls
      LabRC = "R = " & Str$(R) & "  C = " & Str$(C) & "  N =" & Str$(NumBlocksInRC(R, C))
      LabNumBlocks = "NumBlocks =" & Str$(NumBlocks)
      If NumBlocks = 0 Then
         FalseALL
      ElseIf NumBlocks > 0 Then  ' >=1
         R = NR(RedBlockNumber)
         C = NC(RedBlockNumber)
         iy = GridStep * (21 - R)
         ix = GridStep * (C + 5)
         picRC.Line (ix + 1, iy + 1)-(ix + GridStep - 1, iy + GridStep - 1), SelectCul, BF
         REDRAW_PLAN picPLAN
         DRAW_Faces picFace
      End If

      Exit Sub
   End If
   
   If NumBlocksInRC(R, C) > 0 Then
      
      Select Case Index
      Case 0  ' Clear blocks in R,C
         res = MsgBox("Clear" & Str$(NumBlocksInRC(R, C)) & " blocks in this R,C square", vbYesNo, "Clear blocks")
         If res = vbNo Then
            picRC.Line (ix + 1, iy + 1)-(ix + GridStep - 1, iy + GridStep - 1), SelectCul, BF
            RC_GO
            Exit Sub
         Else
            
            Do
               For i = 1 To NumBlocks
                  If NR(i) = R And NC(i) = C Then Exit For
               Next i
               
               If i < NumBlocks Then
                  ' Squeeze out entry
                  For j = i To NumBlocks - 1
                     LX(j) = LX(j + 1)
                     LY(j) = LY(j + 1)
                     LZ(j) = LZ(j + 1)
                     HX(j) = HX(j + 1)
                     HY(j) = HY(j + 1)
                     HZ(j) = HZ(j + 1)
                     NR(j) = NR(j + 1)
                     NC(j) = NC(j + 1)
                  Next j
                  NumBlocks = NumBlocks - 1
                  ReDimPreserve
               
               ElseIf i = NumBlocks Then
                  ' Clear last block
                  If NumBlocks > 1 Then   '>=2
                     NumBlocks = NumBlocks - 1  ' >=1
                     ReDimPreserve
                     Exit Do
                  Else  'NumBlocks=0
                     NumBlocks = 0
                     ReDimCoords
                     Exit Do
                  End If
                  
               Else
                  Exit Do
               End If
            Loop
            
            NumBlocksInRC(R, C) = 0
            LabN = "N =" & Str$(NumBlocksInRC(R, C))
            iy = GridStep * (21 - R)
            ix = GridStep * (C + 5)
            picRC.Line (ix + 1, iy + 1)-(ix + GridStep - 1, iy + GridStep - 1), GridBackColor, BF
            
            ' Clear all pics
            picPLAN.Cls
            picFace.Cls
            picViewsContainer.Cls
            RedBlockNumber = NumBlocks
            LabBlockNum = "BlockNum =" & Str$(RedBlockNumber)
            LabBlocksInRC = "Blocks in RC =" & Str$(NumBlocksInRC(R, C))
            If NumBlocks = 0 Then
               FalseALL
            Else  ' Redden previous RC-square
               R = NR(RedBlockNumber)
               C = NC(RedBlockNumber)
               iy = GridStep * (21 - R)
               ix = GridStep * (C + 5)
               picRC.Line (ix + 1, iy + 1)-(ix + GridStep - 1, iy + GridStep - 1), SelectCul, BF
               REDRAW_PLAN picPLAN
               DRAW_Faces picFace
            End If
         End If
         LabRC = "R = " & Str$(R) & "  C = " & Str$(C) & "  N =" & Str$(NumBlocksInRC(R, C))
         LabNumBlocks = "NumBlocks =" & Str$(NumBlocks)
         
      Case 1  ' Undo red block @ RCIndex
         'RedBlockNumber = Num in NumBlocks
         NR(RedBlockNumber) = 0
         NC(RedBlockNumber) = 0
         If RedBlockNumber = NumBlocks Then
            If NumBlocks > 1 Then   ' >=2
               NumBlocks = NumBlocks - 1  '>=1
               ReDimPreserve
               RedBlockNumber = NumBlocks
            Else
               NumBlocks = 0
               ReDimCoords
               RedBlockNumber = NumBlocks
            End If
               
         Else  ' Squeeze out entry at RedBlockNumber
            If NumBlocks > 1 Then   '>=2
               For i = RedBlockNumber To NumBlocks - 1
                  LX(i) = LX(i + 1)
                  LY(i) = LY(i + 1)
                  LZ(i) = LZ(i + 1)
                  HX(i) = HX(i + 1)
                  HY(i) = HY(i + 1)
                  HZ(i) = HZ(i + 1)
                  NR(i) = NR(i + 1)
                  NC(i) = NC(i + 1)
               Next i
               NumBlocks = NumBlocks - 1
               ReDimPreserve
               RedBlockNumber = NumBlocks
               LabBlockNum = "BlockNum =" & Str$(RedBlockNumber)
               LabBlocksInRC = "Blocks in RC =" & Str$(NumBlocksInRC(R, C))
            End If
         End If
         
         NumBlocksInRC(R, C) = NumBlocksInRC(R, C) - 1
         LabN = "N =" & Str$(NumBlocksInRC(R, C))
         
         If NumBlocksInRC(R, C) = 0 Then
            iy = GridStep * (21 - R)
            ix = GridStep * (C + 5)
            picRC.Line (ix + 1, iy + 1)-(ix + GridStep - 1, iy + GridStep - 1), GridBackColor, BF
            ' Clear all pics
            picPLAN.Cls
            picFace.Cls
            picViewsContainer.Cls
            LabViews.Enabled = False
            
            ' Redden previous RC-square
            If NumBlocks > 0 Then   ' >=1
               R = NR(RedBlockNumber)
               C = NC(RedBlockNumber)
               iy = GridStep * (21 - R)
               ix = GridStep * (C + 5)
               picRC.Line (ix + 1, iy + 1)-(ix + GridStep - 1, iy + GridStep - 1), SelectCul, BF
               REDRAW_PLAN picPLAN
               DRAW_Faces picFace
            End If
            
         Else  ' Set RedBlockNumber to last block in R,C
            BlockCounter = 0
            RedBlockNumber = 0
            For i = 1 To NumBlocks
               If NR(i) = R And NC(i) = C Then
                  RedBlockNumber = i
                  BlockCounter = BlockCounter + 1
                  If BlockCounter = NumBlocksInRC(R, C) Then Exit For
               End If
            Next i
            LabBlockNum = "BlockNum =" & Str$(RedBlockNumber)
            LabBlocksInRC = "Blocks in RC =" & Str$(NumBlocksInRC(R, C))
            REDRAW_PLAN picPLAN
            DRAW_Faces picFace

         End If
         LabRC = "R = " & Str$(R) & "  C = " & Str$(C) & "  N =" & Str$(NumBlocksInRC(R, C))
         LabNumBlocks = "NumBlocks =" & Str$(NumBlocks)
         If NumBlocks = 0 Then FalseALL
         
      Case 2  ' Cycle red block
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
         If Button = vbLeftButton Then
            ''''''''''''''''''''''''''''
            If iRB <> 0 Then     ' iRB = RedBlockNumber
               iRB = iRB + 1
               If iRB > NumBlocksInRC(R, C) Then iRB = 1
            Else
               iRB = 1
            End If
            RedBlockNumber = BlockNumsInRC(iRB)
            ''''''''''''''''''''''''''''
         ElseIf Button = vbRightButton Then
            If iRB <> 0 Then     ' iRB = RedBlockNumber
               iRB = iRB - 1
               If iRB = 0 Then iRB = NumBlocksInRC(R, C)
            Else
               iRB = 1
            End If
            RedBlockNumber = BlockNumsInRC(iRB)
         
         End If
         LabBlockNum = "BlockNum =" & Str$(RedBlockNumber)
         LabBlocksInRC = "Blocks in RC =" & Str$(NumBlocksInRC(R, C))
         REDRAW_PLAN picPLAN
         DRAW_Faces picFace
      
      Case 3   ' Copy & Paste red block
         NumBlocks = NumBlocks + 1
         ReDimPreserve
         dX = HX(RedBlockNumber) - LX(RedBlockNumber)
         dY = HY(RedBlockNumber) - LY(RedBlockNumber)
         dZ = HZ(RedBlockNumber) - LZ(RedBlockNumber)
         
         LX(NumBlocks) = 1
         LY(NumBlocks) = 1
         LZ(NumBlocks) = 1
         HX(NumBlocks) = 1 + dX
         HY(NumBlocks) = 1 + dY
         HZ(NumBlocks) = 1 + dZ
         NR(NumBlocks) = R
         NC(NumBlocks) = C
         NumBlocksInRC(R, C) = NumBlocksInRC(R, C) + 1
         RedBlockNumber = NumBlocks
         LabBlockNum = "BlockNum =" & Str$(RedBlockNumber)
         LabBlocksInRC = "Blocks in RC =" & Str$(NumBlocksInRC(R, C))
         
         LabRC = "R = " & Str$(R) & "  C = " & Str$(C) & "  N =" & Str$(NumBlocksInRC(R, C))
         LabNumBlocks = "NumBlocks =" & Str$(NumBlocks)
         
         REDRAW_PLAN picPLAN
         DRAW_Faces picFace
      
      End Select
   
   End If
End Sub
'#### END CLEAR, UNDO & CYCLE #################################################################

Private Sub ReDimPreserve()
   ReDim Preserve LX(NumBlocks)
   ReDim Preserve LZ(NumBlocks)
   ReDim Preserve LY(NumBlocks)
   ReDim Preserve HX(NumBlocks)
   ReDim Preserve HZ(NumBlocks)
   ReDim Preserve HY(NumBlocks)
   ReDim Preserve NR(NumBlocks)
   ReDim Preserve NC(NumBlocks)
End Sub

Private Sub ReDimCoords()
   ReDim LX(1)
   ReDim LZ(1)
   ReDim LY(1)
   ReDim HX(1)
   ReDim HZ(1)
   ReDim HY(1)
   ReDim NR(1)
   ReDim NC(1)
End Sub

Private Sub FalseALL()
   LabN = " 0"
   LabRow = " Row ="
   LabCol = "Col ="
   fraPlan.Enabled = False
   picFace.Enabled = False
   picViewsContainer.Enabled = False
   LabViews.Enabled = False
   chkCopyPaste(0).Enabled = False
   chkCopyPaste(1).Enabled = False
   For i = 0 To 3
      chkPLAN(i).Enabled = True
   Next i
   cmdPresets(0).Enabled = False
   cmdPresets(1).Enabled = False
   Unload frmPLANPreSets
   AfrmFACEVis = False
   Unload frmFACEPresets
End Sub


'#### [2] PLAN @ R,C #########################################################################

Private Sub optDrawResize_Click(Index As Integer)
   If Index = 0 Then
      PlanAction = True    ' Draw
   Else
      PlanAction = False   ' Alter
   End If
End Sub

Private Sub picPLAN_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim II As Long
   
   aPPMD = True
   
   If X < 1 Then X = 1
   If X > 256 Then X = 256
   If Y < 1 Then Y = 1
   If Y > 256 Then Y = 256
   
   X1 = X
   Z1 = Y
   X2 = X
   Z2 = Y
   picFaceX = X
   picFaceZ = Y
   
   If Shift = 2 Then ' Selecting a block
      If Button = vbLeftButton Then
         If NumBlocks > 0 Then
            
            For II = 1 To NumBlocks
               If NR(II) = R Then
               If NC(II) = C Then
                  If LX(II) - 2 < X Then
                  If X < LX(II) + 2 Then
                  If LZ(II) - 2 < Y Then
                  If Y < LZ(II) + 2 Then
                     RedBlockNumber = II
                     LabBlockNum = "BlockNum =" & Str$(RedBlockNumber)
                     LabBlocksInRC = "Blocks in RC =" & Str$(NumBlocksInRC(R, C))
                     REDRAW_PLAN picPLAN
                     DRAW_Faces picFace
                  End If
                  End If
                  End If
                  End If
               End If
               End If
            Next II
         
         End If
      End If
      aPFMD = False
 
   ElseIf PlanAction Then   ' Draw recs
      If Button = vbLeftButton Then
         picPLAN.Line (X1, Z1)-(X2, Z2), RGB(255, 255, 255), B
      End If
   End If
   
   LabRC = "R = " & Str$(R) & "  C = " & Str$(C) & "  N =" & Str$(NumBlocksInRC(R, C))
   If NumBlocksInRC(R, C) = 0 Then
      LabBlockNum = "BlockNum = 0"
      LabBlocksInRC = "Blocks in RC = 0"
   End If
'LabTest = "X1 =" & Str$(X1) & " Z1 =" & Str$(Z1)
End Sub

Private Sub picPLAN_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim iRB As Long
Dim dX As Long, dZ As Long
   
   If Shift = 2 Then ' Selecting a block
      aPPMD = False
      Exit Sub
   End If

   If X < 1 Then X = 1
   If X > 256 Then X = 256
   If Y < 1 Then Y = 1
   If Y > 256 Then Y = 256
   
   LabXYPLAN = "X =" & Str$(CInt(X)) & "  Z =" & Str$(CInt(Y))
   
   If Not aPPMD Then Exit Sub
   
   If PlanAction Then   ' Draw recs
      If Button = vbLeftButton Then
         picPLAN.Line (X1, Z1)-(X2, Z2), RGB(255, 255, 255), B
         picPLAN.Line (X1, Z1)-(CInt(X), CInt(Y)), RGB(255, 255, 255), B
         X2 = CInt(X)
         Z2 = CInt(Y)
      End If
   Else  ' Alter
      
      If NumBlocksInRC(R, C) = 0 Then
         aPPMD = False
         Exit Sub
      End If
      
      If Button = vbLeftButton Then ' Change X2
         
         XOR_RedBlocks RedBlockNumber
         dX = X - picFaceX
         If dX > 0 And HX(RedBlockNumber) + dX <= 256 Then
            HX(RedBlockNumber) = HX(RedBlockNumber) + dX
         ElseIf dX < 0 And HX(RedBlockNumber) + dX - LX(RedBlockNumber) >= 1 Then
            HX(RedBlockNumber) = HX(RedBlockNumber) + dX
         End If
         dZ = Y - picFaceZ ' Change Z2
         If dZ > 0 And HZ(RedBlockNumber) + dZ <= 256 Then
            HZ(RedBlockNumber) = HZ(RedBlockNumber) + dZ
         ElseIf dZ < 0 And HZ(RedBlockNumber) + dZ - LZ(RedBlockNumber) > 1 Then
            HZ(RedBlockNumber) = HZ(RedBlockNumber) + dZ
         End If
         picFaceX = X
         picFaceZ = Y
         XOR_RedBlocks RedBlockNumber
      
      ElseIf Button = vbRightButton Then  ' Move red block
        
'''''''''' Code to move all blocks in RC-square ''''''''''''''''''''''''
'
'         'Find BlockNumbers in RC-square ( not nec in order!)
'         ReDim BlockNumsInRC(NumBlocksInRC(R, C)) '''''''''''''
'         BlockCounter = 0
'         iRB = 0 '''''''''''''''
'         For i = 1 To NumBlocks
'            If NR(i) = R And NC(i) = C Then
'               BlockCounter = BlockCounter + 1
'               BlockNumsInRC(BlockCounter) = i
'               If BlockCounter = NumBlocksInRC(R, C) Then Exit For
'            End If
'         Next i
'
'         For i = 1 To BlockCounter
'           j = BlockNumsInRC(i)
'           dX = X - picFaceX
'           If (dX > 0 And HX(j) + dX <= 256) Or _
'              (dX < 0 And LX(j) + dX >= 1) Then
'
'              LX(j) = LX(j) + dX
'              HX(j) = HX(j) + dX
'           End If
'           dZ = Y - picFaceZ
'           If (dZ > 0 And HZ(j) + dZ <= 256) Or _
'              (dZ < 0 And LZ(j) + dZ >= 1) Then
'
'              LZ(j) = LZ(j) + dZ
'              HZ(j) = HZ(j) + dZ
'           End If
'           picFaceX = X
'           picFaceZ = Y
'           ' Cls ' Plans & Face
'           MOVE_AllBlocks j 'RedBlockNumber
'         Next i
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                  
         
         XOR_RedBlocks RedBlockNumber
         dX = X - picFaceX
         If (dX > 0 And HX(RedBlockNumber) + dX <= 256) Or _
            (dX < 0 And LX(RedBlockNumber) + dX >= 1) Then

            LX(RedBlockNumber) = LX(RedBlockNumber) + dX
            HX(RedBlockNumber) = HX(RedBlockNumber) + dX
         End If
         dZ = Y - picFaceZ
         If (dZ > 0 And HZ(RedBlockNumber) + dZ <= 256) Or _
            (dZ < 0 And LZ(RedBlockNumber) + dZ >= 1) Then

            LZ(RedBlockNumber) = LZ(RedBlockNumber) + dZ
            HZ(RedBlockNumber) = HZ(RedBlockNumber) + dZ
         End If
         picFaceX = X
         picFaceZ = Y
         XOR_RedBlocks RedBlockNumber
      End If
   
   End If
End Sub

Private Sub picPLAN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Shift = 2 Then ' Selecting a block
      aPPMD = False
      Exit Sub
   End If
   If Not aPPMD Then Exit Sub
   
   If NumBlocksInRC(R, C) = 0 Then
      LabBlockNum = "BlockNum = 0"
      LabBlocksInRC = "Blocks in RC = 0"
   End If
   
   If Button = vbLeftButton Then
      If PlanAction Then   ' Draw recs
         NumBlocks = NumBlocks + 1
         LabNumBlocks = "NumBlocks =" & Str$(NumBlocks)
         NumBlocksInRC(R, C) = NumBlocksInRC(R, C) + 1
         LabN = "N =" & Str$(NumBlocksInRC(R, C))
         ReDimPreserve
         
         NR(NumBlocks) = R
         NC(NumBlocks) = C
         If X1 < X2 Then
            LX(NumBlocks) = X1
            HX(NumBlocks) = X2
         Else
            HX(NumBlocks) = X1
            LX(NumBlocks) = X2
         End If
         If Z1 < Z2 Then
            LZ(NumBlocks) = Z1
            HZ(NumBlocks) = Z2
         Else
            HZ(NumBlocks) = Z1
            LZ(NumBlocks) = Z2
         End If
         
      '        In R,C-square
      '      Z
      '      |
      '      .-------o HX(#),HZ(#)
      '      |       |
      '      |       |
      '      |       |
      '      o-------.--X
      ' LX(#),LZ(#)
            
         LabRC = "R = " & Str$(R) & "  C = " & Str$(C) & "  N =" & Str$(NumBlocksInRC(R, C))
         
         ' Redraw PLAN
         RedBlockNumber = NumBlocks
         LabBlockNum = "BlockNum =" & Str$(RedBlockNumber)
         LabBlocksInRC = "Blocks in RC =" & Str$(NumBlocksInRC(R, C))
         REDRAW_PLAN picPLAN
         DRAW_Faces picFace
      
      ElseIf Not PlanAction Then ' Alter dimensions, LeftButton
         If FaceIndex >= 4 Then   ' Whole plan & RC perspec(other faces taken care of)
            DRAW_Faces picFace
         End If
      End If
   
   ElseIf Button = vbRightButton Then  ' Alter block position, RightButton
      If FaceIndex >= 4 Then   ' Whole plan & RC perspec (other faces taken care of)
         DRAW_Faces picFace
      End If
   End If
   aPPMD = False
End Sub
'#### END PLAN @ R,C #########################################################################


'#### FACES ##################################################################################

Private Sub cmdFaces_Click(Index As Integer)
' Cycle thru' faces
   HSPerspec.Visible = False
   Select Case Index
   Case 0   '<
      If FaceIndex = 0 Then FaceIndex = 6
      FaceIndex = FaceIndex - 1
   Case 1   '>
      If FaceIndex = 5 Then FaceIndex = -1
      FaceIndex = FaceIndex + 1
   End Select
   LabFaceCap = "Top"
   Select Case FaceIndex
   Case 0: LabFace = "  Front face "
           LabFaceAxis(0) = "X"
           LabFaceAxis(1) = "1,1"
           LabFaceAxis(2) = "Y"
           LabFaceAxis(3) = ""
           DRAW_Faces picFace
   Case 1: LabFace = "  Right face "
           LabFaceAxis(0) = "Z"
           LabFaceAxis(1) = "1.1"
           LabFaceAxis(2) = "Y"
           LabFaceAxis(3) = ""
           DRAW_Faces picFace
   Case 2: LabFace = "  Back face "
           LabFaceAxis(0) = "1,1"
           LabFaceAxis(1) = "X"
           LabFaceAxis(2) = ""
           LabFaceAxis(3) = "Y"
           DRAW_Faces picFace
   Case 3: LabFace = "  Left face "
           LabFaceAxis(0) = "1,1"
           LabFaceAxis(1) = "Z"
           LabFaceAxis(2) = ""
           LabFaceAxis(3) = "Y"
           DRAW_Faces picFace
   Case 4: LabFace = "  Whole plan "
           LabFaceCap = "Far Boundary"
           LabFaceAxis(0) = "X"
           LabFaceAxis(1) = "1.1"
           LabFaceAxis(2) = "Z"
           LabFaceAxis(3) = ""
           DRAW_Faces picFace
   Case 5: LabFace = "  RC perspec "
           LabFaceCap = "Far Boundary"
           LabFaceAxis(0) = "X"
           LabFaceAxis(1) = "1.1"
           LabFaceAxis(2) = "Z"
           LabFaceAxis(3) = ""
           HSPerspec.Visible = True
           DRAW_Faces picFace
   End Select
   If FaceIndex < 4 And AfrmFACEVis Then
      SetCorrectFaceIndex
   End If
End Sub

Private Sub SetCorrectFaceIndex()
   FaceNum = FaceIndex
   With frmFACEPresets
      Select Case FaceNum
      Case 0: .LabFaces = "  Front face "
      Case 1: .LabFaces = "  Right face "
      Case 2: .LabFaces = "  Back face "
      Case 3: .LabFaces = "  Left face "
      End Select
   End With
End Sub

Private Sub optHtRec_Click(Index As Integer)
' Select face action
   If Index = 0 Then ' Alter blocks
      FaceAction = True
   Else  ' Draw on faces
      FaceAction = False
   End If
End Sub

Private Sub picFace_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim LoX As Long
Dim HiX As Long
Dim LoY As Long
Dim HiY As Long
Dim LoZ As Long
Dim HiZ As Long

Dim II As Long


Dim XX As Single
Dim YY As Single
Dim RR As Long
Dim CC As Long
   
   aPFMD = True
   
   If X < 1 Then X = 1
   If X > 256 Then X = 256
   If Y < 1 Then Y = 1
   If Y > 512 Then Y = 512
   
   X1 = X
   Y1 = Y
   X2 = X
   Y2 = Y
   
   If FaceIndex = 4 Then   ' Whole plan find RC-square
      RR = CLng((18 * Y1 - 1) / 256 + 0.5) - 2 - 1
      If RR > 21 Then RR = -2
      If RR < -2 Then RR = 21
      CC = CLng((18 * X1 - 1) / 256 + 0.5) - 5 - 1
      If CC > 12 Then CC = -5
      If CC < -5 Then CC = 12
      
      YY = (21 - RR) * GridStep
      XX = (CC + 5) * GridStep
      ' At picRC_MouseUp  - Check
      RR = 21 - CLng(YY \ GridStep)
      CC = CLng(XX \ GridStep) - 5
      
      If RR >= -2 And RR <= 21 Then
         If CC >= -5 And CC <= 12 Then
            picRC_MouseUp 1, 0, XX, YY
         End If
      End If
      
      aPFMD = False
      Exit Sub
   End If
   
   picFaceX = X
   picFaceY = Y
   picFaceZ = Y
   
   If Shift = 2 Then ' Selecting a block
      If Button = vbLeftButton Then
         If NumBlocks > 0 Then
            
            For II = 1 To NumBlocks
               If NR(II) = R Then
               If NC(II) = C Then
                  LoX = LX(II) - 2
                  HiX = LX(II) + 2
                  LoY = LY(II) - 2
                  HiY = LY(II) + 2
                  LoZ = LZ(II) - 1
                  HiZ = HZ(II) + 2
                  Select Case FaceIndex
                  Case 0   ' Front face X,Y
                     If LoX < X Then
                     If X < HiX Then
                     If LoY < Y Then
                     If Y < HiY Then
                        RedBlockNumber = II
                        LabBlockNum = "BlockNum =" & Str$(RedBlockNumber)
                        LabBlocksInRC = "Blocks in RC =" & Str$(NumBlocksInRC(R, C))
                        REDRAW_PLAN picPLAN
                        DRAW_Faces picFace
                     End If
                     End If
                     End If
                     End If
                  Case 1 ' Right face Z,Y
                     If LoZ < X Then
                     If X < HiZ Then
                     If LoY < Y Then
                     If Y < HiY Then
                        RedBlockNumber = II
                        LabBlockNum = "BlockNum =" & Str$(RedBlockNumber)
                        LabBlocksInRC = "Blocks in RC =" & Str$(NumBlocksInRC(R, C))
                        REDRAW_PLAN picPLAN
                        DRAW_Faces picFace
                     End If
                     End If
                     End If
                     End If
                  Case 2 ' Back face -X,Y
                     If LoX < 256 - X Then
                     If 256 - X < HiX Then
                     If LoY < Y Then
                     If Y < HiY Then
                        RedBlockNumber = II
                        LabBlockNum = "BlockNum =" & Str$(RedBlockNumber)
                        LabBlocksInRC = "Blocks in RC =" & Str$(NumBlocksInRC(R, C))
                        REDRAW_PLAN picPLAN
                        DRAW_Faces picFace
                     End If
                     End If
                     End If
                     End If
                  Case 3 ' Left face -Z,Y
                     If LoZ < 256 - X Then
                     If 256 - X < HiZ Then
                     If LoY < Y Then
                     If Y < HiY Then
                        RedBlockNumber = II
                        LabBlockNum = "BlockNum =" & Str$(RedBlockNumber)
                        LabBlocksInRC = "Blocks in RC =" & Str$(NumBlocksInRC(R, C))
                        REDRAW_PLAN picPLAN
                        DRAW_Faces picFace
                     End If
                     End If
                     End If
                     End If
                  End Select
               End If
               End If
            Next II
         
         End If
      End If
   
   
   ElseIf Not FaceAction Then   ' ie Draw on face
      If Button = vbLeftButton Then
         picFace.Line (X1, Y1)-(X2, Y2), RGB(255, 255, 255), B
      End If
   Else
      If RedBlockNumber < 1 Then
         aPFMD = False
         Exit Sub
      End If
   End If
   LabXYFACE = "X =" & Str$(X) & "  Y = " & Str$(Y)
End Sub

Private Sub picFace_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   If FaceIndex = 4 Or FaceIndex = 5 Then Exit Sub ' Whole plan or RC perspec
   
   If Shift = 2 Then ' Selecting a block
      aPFMD = False
      Exit Sub
   End If
   
   If X < 1 Then X = 1
   If X > 256 Then X = 256
   If Y < 1 Then Y = 1
   If Y > 512 Then Y = 512
   
   Select Case FaceIndex
   Case 0: LabXYFACE = "X =" & Str$(X) & "  Y =" & Str$(Y)
   Case 2: LabXYFACE = "X =" & Str$(257 - X) & "  Y =" & Str$(Y)
   Case 1: LabXYFACE = "Z =" & Str$(X) & "  Y =" & Str$(Y)
   Case 3: LabXYFACE = "Z =" & Str$(257 - X) & "  Y =" & Str$(Y)
   End Select
   
   If Not aPFMD Then Exit Sub
   
   If FaceAction Then   ' Alter blocks
      
      If RedBlockNumber < 1 Then Exit Sub
      If NumBlocksInRC(R, C) = 0 Then Exit Sub
      
      If Button = vbLeftButton Then  ' Change X,Y
         
         
         XOR_RedBlocks RedBlockNumber
         
         ' All faces have Y upwards
         dY = Y - picFaceY ' Change Y2
         If dY > 0 And HY(RedBlockNumber) + dY <= 512 Then
            HY(RedBlockNumber) = HY(RedBlockNumber) + dY
         ElseIf dY < 0 And HY(RedBlockNumber) + dY - LY(RedBlockNumber) >= 1 Then
            HY(RedBlockNumber) = HY(RedBlockNumber) + dY
         End If
         picFaceY = Y
         
         ' Depends on FaceIndex
         Select Case FaceIndex
         Case 0, 2  ' Front face X,Y & Back face    -X,Y
            If FaceIndex = 0 Then
               dX = X - picFaceX ' Change X2
            Else  ' FaceIndex=2 Back
               dX = -(X - picFaceX) ' Change X2
            End If
            If dX > 0 And HX(RedBlockNumber) + dX <= 256 Then
               HX(RedBlockNumber) = HX(RedBlockNumber) + dX
            ElseIf dX < 0 And HX(RedBlockNumber) + dX - LX(RedBlockNumber) >= 1 Then
               HX(RedBlockNumber) = HX(RedBlockNumber) + dX
            End If
            If FaceIndex = 0 Then
               LabXYFACE = "X =" & Str$(HX(RedBlockNumber)) & "  Y =" & Str$(HY(RedBlockNumber))
            Else
               LabXYFACE = "X =" & Str$(257 - HX(RedBlockNumber)) & "  Y =" & Str$(HY(RedBlockNumber))
            End If
            picFaceX = X
            XOR_RedBlocks RedBlockNumber
         
         Case 1, 3  ' Right face  Z,Y & Left face    -Z,Y
            If FaceIndex = 1 Then
               dZ = X - picFaceZ ' Change Z2
            Else  ' FaceIndex=3 Left
               dZ = -(X - picFaceZ) ' Change Z2
            End If
            If dZ > 0 And HZ(RedBlockNumber) + dZ <= 256 Then
               HZ(RedBlockNumber) = HZ(RedBlockNumber) + dZ
            ElseIf dZ < 0 And HZ(RedBlockNumber) + dZ - LZ(RedBlockNumber) >= 1 Then
               HZ(RedBlockNumber) = HZ(RedBlockNumber) + dZ
            End If
            If FaceIndex = 1 Then
               LabXYFACE = "Z =" & Str$(HZ(RedBlockNumber)) & "  Y =" & Str$(HY(RedBlockNumber))
            Else
               LabXYFACE = "Z =" & Str$(257 - HZ(RedBlockNumber)) & "  Y =" & Str$(HY(RedBlockNumber))
            End If
            picFaceZ = X
            XOR_RedBlocks RedBlockNumber
         End Select
      
      ElseIf Button = vbRightButton Then  ' Change block position PX1,PY1 & PX2,PY2 (RedBlockNumber)
         
         XOR_RedBlocks RedBlockNumber
         
         ' All faces have Y upwards
         dY = Y - picFaceY
         If (dY > 0 And HY(RedBlockNumber) + dY <= 512) Or _
            (dY < 0 And LY(RedBlockNumber) + dY >= 1) Then
            
            LY(RedBlockNumber) = LY(RedBlockNumber) + dY
            HY(RedBlockNumber) = HY(RedBlockNumber) + dY
         End If
         picFaceY = Y
         
         ' Depends on FaceIndex
         Select Case FaceIndex
         Case 0, 2  ' Front face X,Y & Back face  -X,Y
            If FaceIndex = 0 Then
               dX = X - picFaceX
            Else
               dX = -(X - picFaceX)
            End If
            If (dX > 0 And HX(RedBlockNumber) + dX <= 256) Or _
               (dX < 0 And LX(RedBlockNumber) + dX >= 1) Then
               
               LX(RedBlockNumber) = LX(RedBlockNumber) + dX
               HX(RedBlockNumber) = HX(RedBlockNumber) + dX
            End If
            If FaceIndex = 0 Then
               LabXYFACE = "X =" & Str$(LX(RedBlockNumber)) & "  Y =" & Str$(LY(RedBlockNumber))
            Else
               LabXYFACE = "X =" & Str$(257 - LX(RedBlockNumber)) & "  Y =" & Str$(LY(RedBlockNumber))
            End If
            picFaceX = X
            XOR_RedBlocks RedBlockNumber
         
         Case 1, 3  ' Right face Z,Y & Left face -Z,Y
            If FaceIndex = 1 Then
               dZ = X - picFaceZ
            Else
               dZ = -(X - picFaceZ)
            End If
            If (dZ > 0 And HZ(RedBlockNumber) + dZ <= 256) Or _
               (dZ < 0 And LZ(RedBlockNumber) + dZ >= 1) Then
               
               LZ(RedBlockNumber) = LZ(RedBlockNumber) + dZ
               HZ(RedBlockNumber) = HZ(RedBlockNumber) + dZ
            End If
            If FaceIndex = 1 Then
               LabXYFACE = "Z =" & Str$(LZ(RedBlockNumber)) & "  Y =" & Str$(LY(RedBlockNumber))
            Else
               LabXYFACE = "Z =" & Str$(257 - LZ(RedBlockNumber)) & "  Y =" & Str$(LY(RedBlockNumber))
            End If
            picFaceZ = X
            XOR_RedBlocks RedBlockNumber
         
         End Select
      End If
   
   Else     ' Draw on faces
         
      If Button = vbLeftButton Then
         picFace.Line (X1, Y1)-(X2, Y2), RGB(255, 255, 255), B
         picFace.Line (X1, Y1)-(CInt(X), CInt(Y)), RGB(255, 255, 255), B
         X2 = CInt(X)
         Y2 = CInt(Y)
      End If
      
   End If
End Sub

Private Sub picFace_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   If FaceIndex = 4 Or FaceIndex = 5 Then Exit Sub ' Whole plan or RC perspec
   
   If Shift = 2 Then ' Selecting a block
      aPFMD = False
      Exit Sub
   End If
   
   If Not aPFMD Then Exit Sub
   
   If Not FaceAction Then  ' Draw on face
   If Button = vbLeftButton Then
      NumBlocks = NumBlocks + 1
      LabNumBlocks = "NumBlocks =" & Str$(NumBlocks)
      NumBlocksInRC(R, C) = NumBlocksInRC(R, C) + 1
      LabN = "N =" & Str$(NumBlocksInRC(R, C))
      ReDimPreserve
      
      NR(NumBlocks) = R
      NC(NumBlocks) = C
         
      If RedBlockNumber < 1 Then
         RedBlockNumber = NumBlocks
         LabBlockNum = "BlockNum =" & Str$(RedBlockNumber)
         LabBlocksInRC = "Blocks in RC =" & Str$(NumBlocksInRC(R, C))
      End If
      
      If Y1 < Y2 Then
         LY(NumBlocks) = Y1: HY(NumBlocks) = Y2
      Else
         HY(NumBlocks) = Y1: LY(NumBlocks) = Y2
      End If
      
      Select Case FaceIndex
      Case 0   ' Front face  X,Y
         If X1 < X2 Then
            LX(NumBlocks) = X1: HX(NumBlocks) = X2
         Else
            HX(NumBlocks) = X1: LX(NumBlocks) = X2
         End If
         If LZ(RedBlockNumber) < 1 Then LZ(RedBlockNumber) = 1
         LZ(NumBlocks) = LZ(RedBlockNumber)
         HZ(NumBlocks) = LZ(NumBlocks) + DummyDepth
      Case 1   ' Right face  Z,Y
         If X1 < X2 Then
            LZ(NumBlocks) = X1: HZ(NumBlocks) = X2
         Else
            HZ(NumBlocks) = X1: LZ(NumBlocks) = X2
         End If
         If HX(RedBlockNumber) < 1 Then HX(RedBlockNumber) = 1
         HX(NumBlocks) = HX(RedBlockNumber)
         LX(NumBlocks) = HX(NumBlocks) - DummyDepth
      
      Case 2   ' Back face  -X,Y
         X1 = 256 - X1: X2 = 256 - X2
         If X1 < X2 Then
            LX(NumBlocks) = X1: HX(NumBlocks) = X2
         Else
            HX(NumBlocks) = X1: LX(NumBlocks) = X2
         End If
         If HZ(RedBlockNumber) < 1 Then HZ(RedBlockNumber) = 1
         HZ(NumBlocks) = HZ(RedBlockNumber)
         LZ(NumBlocks) = HZ(NumBlocks) - DummyDepth
      
      Case 3   ' Left face  -Z,Y
         X1 = 256 - X1: X2 = 256 - X2
         If X1 < X2 Then
            LZ(NumBlocks) = X1: HZ(NumBlocks) = X2
         Else
            HZ(NumBlocks) = X1: LZ(NumBlocks) = X2
         End If
         If LX(RedBlockNumber) < 1 Then LX(RedBlockNumber) = 1
         LX(NumBlocks) = LX(RedBlockNumber)
         HX(NumBlocks) = LX(NumBlocks) + DummyDepth
      
      End Select
      
   ' COORDS
   '      Y       B
   '      |   .-----o HX(#),HY(#),HZ(#)
   '      |  /|     /
   '      | / |    /|  Z
   '      .-------. | /
   '   L  |   .- -|-.       R
   '      |  /    | /
   '      | /     |/
   '      o-------.----X
   ' LX(#),LY(#),LZ(#)
   '
   '          F
      
      LabRC = "R = " & Str$(R) & "  C = " & Str$(C) & "  N =" & Str$(NumBlocksInRC(R, C))
      
      ' Redraw PLAN
      RedBlockNumber = NumBlocks
      LabBlockNum = "BlockNum =" & Str$(RedBlockNumber)
      LabBlocksInRC = "Blocks in RC =" & Str$(NumBlocksInRC(R, C))
      REDRAW_PLAN picPLAN
      DRAW_Faces picFace
   End If
   End If
   
   aPFMD = False
   
End Sub

'#### END FACES ##################################################################################

Private Sub XOR_RedBlocks(N As Long)
' Clear/Redraw N=RedBlockNumber or BlockCounter
Dim Cul As Long
Dim px1 As Long
Dim py1 As Long
Dim pz1 As Long
Dim px2 As Long
Dim py2 As Long
Dim pz2 As Long

'   If RedBlockNumber = 0 Then Exit Sub
'   px1 = LX(RedBlockNumber)
'   py1 = LY(RedBlockNumber)
'   pz1 = LZ(RedBlockNumber)
'   px2 = HX(RedBlockNumber)
'   py2 = HY(RedBlockNumber)
'   pz2 = HZ(RedBlockNumber)

   If N = 0 Then Exit Sub
   px1 = LX(N)
   py1 = LY(N)
   pz1 = LZ(N)
   px2 = HX(N)
   py2 = HY(N)
   pz2 = HZ(N)
   
   
   Cul = picPLANCul Xor SelectCul
   picPLAN.Line (px1, pz1)-(px2, pz2), Cul, B
   picPLAN.Circle (px1, pz1), 2, Cul
   
   Cul = picFaceBacCul Xor SelectCul
   Select Case FaceIndex
   Case 0   ' 1 Front face X,Y
      picFace.Line (px1, py1)-(px2, py2), Cul, B
      picFace.Circle (px1, py1), 2, Cul
   Case 1   ' 2 Right face  Z,Y
      picFace.Line (pz1, py1)-(pz2, py2), Cul, B
      picFace.Circle (pz1, py1), 2, Cul
   Case 2   ' 3 Back face -X,Y
      picFace.Line (256 - px1, py1)-(256 - px2, py2), Cul, B
      picFace.Circle (256 - px1, py1), 2, Cul
   Case 3   ' 4 Left face -Z,Y
      picFace.Line (256 - pz1, py1)-(256 - pz2, py2), Cul, B
      picFace.Circle (256 - pz1, py1), 2, Cul
   End Select

End Sub


'#### LOAD/SAVE CCC FILE ############################################################

Private Sub cmdLoadSave_Click(Index As Integer)
Dim Title$, Filt$, InDir$


Set CommonDialog1 = New OSDialog

   Select Case Index
   Case 0
      Title$ = "Load Cuboid File"
      Filt$ = "Load ccc (*.ccc)|*.ccc"
      InDir$ = CCC_Path$
      CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, InDir$, "", Me.hWnd
      
      ' Offset cursor to avoid click thru
      SetCursorPos Me.Left \ STX + 150, Me.Top \ STY + 80

      If Len(FileSpec$) <> 0 Then
         
         CCC_Path$ = FileSpec$
         
         Screen.MousePointer = vbHourglass
         
         READ_CCC_FILE
         
         Me.Caption = " " & FileSpec$
         
         LabNumBlocks = "NumBlocks =" & Str$(NumBlocks)

         '''''''''''''''''''''''''''''''''''''''''''
         ' Return with R & C & NumBlocksInRC(R, C)
         ' Set all RC to white
         DrawRCGrid
         For j = -2 To 21  ' Row
         For i = -5 To 12  ' Col
            ix = GridStep * (i + 5)
            iy = GridStep * (21 - j)
            If NumBlocksInRC(j, i) > 0 Then
               picRC.Line (ix + 1, iy + 1)-(ix + GridStep - 1, iy + GridStep - 1), vbWhite, BF
            End If
         Next i
         Next j
         
         ix = GridStep * (C + 5)
         iy = GridStep * (21 - R)
         
         picRC.Line (ix + 1, iy + 1)-(ix + GridStep - 1, iy + GridStep - 1), SelectCul, BF
      
         ' Clear all pics
         picPLAN.Cls
         picFace.Cls
         picViewsContainer.Cls
      
         RC_GO
         
         optDrawResize(0).Value = True
         PlanAction = True    ' Draw plan
         
         optHtRec(0).Value = True
         FaceAction = True    ' Alter faces
         
         LabRow = " Row =" & Str$(R)
         LabCol = "Col =" & Str$(C)
         '''''''''''''''''''''''''''''''''''''''''''
         Screen.MousePointer = vbDefault
      End If
   
   Case 1 ' Add
   
      If R < -2 Or R > 21 Or C < -5 Or C > 12 Then
         MsgBox "Set R,C first", vbInformation, "Adding ccc file"
         Set CommonDialog1 = Nothing
         Exit Sub
      Else
      
         Title$ = "Add Cuboid File"
         Filt$ = "Add ccc (*.ccc)|*.ccc"
         InDir$ = CCC_Path$
         CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, InDir$, "", Me.hWnd
         
         ' Offset cursor to avoid click thru
         SetCursorPos Me.Left \ STX + 150, Me.Top \ STY + 80
   
         If Len(FileSpec$) <> 0 Then
            
            CCC_Path$ = FileSpec$
            
            Screen.MousePointer = vbHourglass
            
            ADD_CCC_FILE
            
            Me.Caption = " " & FileSpec$
            
            LabNumBlocks = "NumBlocks =" & Str$(NumBlocks)
   
            '''''''''''''''''''''''''''''''''''''''''''
            ' Return with R & C & NumBlocksInRC(R, C)
            ' Set all RC to white
            DrawRCGrid
            For j = -2 To 21  ' Row
            For i = -5 To 12  ' Col
               ix = GridStep * (i + 5)
               iy = GridStep * (21 - j)
               If NumBlocksInRC(j, i) > 0 Then
                  picRC.Line (ix + 1, iy + 1)-(ix + GridStep - 1, iy + GridStep - 1), vbWhite, BF
               End If
            Next i
            Next j
            
            ix = GridStep * (C + 5)
            iy = GridStep * (21 - R)
            
            picRC.Line (ix + 1, iy + 1)-(ix + GridStep - 1, iy + GridStep - 1), SelectCul, BF
         
            ' Clear all pics
            picPLAN.Cls
            picFace.Cls
            picViewsContainer.Cls
         
            RC_GO
            
            optDrawResize(0).Value = True
            PlanAction = True    ' Draw plan
            
            optHtRec(0).Value = True
            FaceAction = True    ' Alter faces
            
            LabRow = " Row =" & Str$(R)
            LabCol = "Col =" & Str$(C)
            '''''''''''''''''''''''''''''''''''''''''''
            Screen.MousePointer = vbDefault
         End If
      
      End If
   
   Case 2
   
      If NumBlocks < 1 Then
         MsgBox " No blocks to save", vbInformation, "Saving ccc file"
         Exit Sub
      End If
      Title$ = "Save Cuboid File"
      Filt$ = "Save ccc (*.ccc)|*.ccc"
      InDir$ = CCC_Path$
      CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, "", Me.hWnd
   
      ' Offset cursor to avoid click thru
      SetCursorPos Me.Left \ STX + 150, Me.Top \ STY + 80
      
      If Len(FileSpec$) <> 0 Then
         
         CCC_Path$ = FileSpec$
         
         Screen.MousePointer = vbHourglass
         FixExtension FileSpec$, "ccc"
         CCC_Path$ = FileSpec$
         SAVE_CCC_FILE
         Screen.MousePointer = vbDefault
      
         Me.Caption = " " & FileSpec$
      
      End If
   End Select
Set CommonDialog1 = Nothing
End Sub
'#### END LOAD/SAVE CCC FILE ############################################################

Private Sub DrawRCGrid()
   GridStep = 7
   GridWidth = 18 * GridStep  ' 18 RC blocks wide
   GridHeight = 24 * GridStep ' 24 RC blocks wide
   ' Draw RC Grid
   picRC.BackColor = GridBackColor
   picRC.Cls
   picRC.Width = GridWidth * STX
   picRC.Height = GridHeight * STY
   For i = 0 To GridWidth Step GridStep
      picRC.Line (i, 0)-(i, picRC.Height), vbWhite
   Next i
   For j = 0 To GridHeight Step GridStep
      picRC.Line (0, j)-(picRC.Width, j), vbWhite
   Next j
End Sub

Private Sub LOCATE_CONTROLS()
Me.Width = 11400 '11700
Me.Height = 8700
   Me.Caption = " Cube City Coords  by Robert Rayment" ' (Start by selecting an RC-square)"
   DrawRCGrid ' RC grid
   
   chkCopyPaste(0).Enabled = False
   chkCopyPaste(1).Enabled = False
   For i = 0 To 3
      chkPLAN(i).Enabled = True
   Next i
   
   With picPLAN
      .ScaleTop = 256
      .ScaleLeft = 1
      .ScaleWidth = 256
      .ScaleHeight = -256
      .DrawMode = 7
   End With
   picPLAN.MousePointer = vbCustom
   picPLAN.MouseIcon = LoadResPicture(101, vbResCursor)
   
   With picFace
      .ScaleTop = 512
      .ScaleLeft = 1
      .ScaleWidth = 256
      .ScaleHeight = -512
      .DrawMode = 7
   End With
   picFace.MousePointer = vbCustom
   picFace.MouseIcon = LoadResPicture(101, vbResCursor)

   LabFaceAxis(0) = "X"
   LabFaceAxis(1) = "1,1"
   LabFaceAxis(2) = "Y"
   LabFaceAxis(3) = ""
   
   ' 16x16 Grid locations
   ' on picPLAN & picFACE
   If Not aLinLoaded Then
      ' LinPlanGrid() & LinFaceGrid()
      For i = 1 To 17
         Load LinPlanGrid(i)
      Next i
      For i = 0 To 8
         If i = 0 Then
            LinPlanGrid(i).X1 = 1
            LinPlanGrid(i).Y1 = 1
            LinPlanGrid(i).X2 = 1
            LinPlanGrid(i).Y2 = 256
         Else
            LinPlanGrid(i).X1 = 32 + (i - 1) * 32
            LinPlanGrid(i).Y1 = 1
            LinPlanGrid(i).X2 = 32 + (i - 1) * 32
            LinPlanGrid(i).Y2 = 256
         End If
         LinPlanGrid(i).Visible = True
      Next i
      For i = 9 To 17
         If i = 9 Then
            LinPlanGrid(i).X1 = 1
            LinPlanGrid(i).Y1 = 1
            LinPlanGrid(i).X2 = 256
            LinPlanGrid(i).Y2 = 1
         Else
            LinPlanGrid(i).X1 = 1
            LinPlanGrid(i).Y1 = 32 + (i - 10) * 32
            LinPlanGrid(i).X2 = 256
            LinPlanGrid(i).Y2 = 32 + (i - 10) * 32
         End If
         LinPlanGrid(i).Visible = True
      Next i

      For i = 1 To 25
         Load LinFaceGrid(i)
      Next i
      For i = 0 To 8
         If i = 0 Then
            LinFaceGrid(i).X1 = 1
            LinFaceGrid(i).Y1 = 1
            LinFaceGrid(i).X2 = 1
            LinFaceGrid(i).Y2 = 512
         Else
            LinFaceGrid(i).X1 = 32 + (i - 1) * 32
            LinFaceGrid(i).Y1 = 1
            LinFaceGrid(i).X2 = 32 + (i - 1) * 32
            LinFaceGrid(i).Y2 = 512
         End If
         LinFaceGrid(i).Visible = True
      Next i
      For i = 9 To 25
         If i = 9 Then
            LinFaceGrid(i).X1 = 1
            LinFaceGrid(i).Y1 = 1
            LinFaceGrid(i).X2 = 256
            LinFaceGrid(i).Y2 = 1
         Else
            LinFaceGrid(i).X1 = 1
            LinFaceGrid(i).Y1 = 32 + (i - 10) * 32
            LinFaceGrid(i).X2 = 256
            LinFaceGrid(i).Y2 = 32 + (i - 10) * 32
         End If
         LinFaceGrid(i).Visible = True
      Next i
      
      aLinLoaded = True
   End If
   
   chkGrid.Value = 1
End Sub

Private Sub chkGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim L As Long
   AGRID = Not AGRID
   For L = 0 To 17
      LinPlanGrid(L).Visible = AGRID
   Next L
   For L = 0 To 25
      LinFaceGrid(L).Visible = AGRID
   Next L
End Sub

'#### QUITTING ###############################################

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Form As Form

   If UnloadMode = 0 Then    'Close on Form1 pressed
         
      res = MsgBox("", vbQuestion + vbYesNo, "Quit Application ?")
      If res = vbNo Then
         Cancel = True
      Else
         Cancel = False
         
         Screen.MousePointer = vbDefault
         
         ' Make sure all forms cleared
         For Each Form In Forms
            Unload Form
            Set Form = Nothing
         Next Form
         End
      
      End If
   End If
End Sub

'#### DEFAULT DEPTH & HEIGHT ########################################

Private Sub HSDefault_Change(Index As Integer)
   Select Case Index
   Case 0:  ' DummyHeight
      DummyHeight = HSDefault(0).Value
   Case 1:  ' DummyDepth
      DummyDepth = HSDefault(1).Value
   End Select
   
   txtDefault(Index) = HSDefault(Index).Value
   
End Sub

Private Sub cmdCloseDefault_Click()
   fraDefaults.Visible = False
End Sub

Private Sub cmdDefaults_Click()
   fraDefaults.Visible = Not fraDefaults.Visible
End Sub

'#### FRAME MOVING #################################################################
Private Sub fraDefaults_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   fraX = X
   fraY = Y
End Sub

Private Sub fraDefaults_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
      fraMOVER Form1, fraDefaults, fraX, fraY, Button, X, Y
   End If
End Sub
'#### END FRAME MOVING #################################################################
'#### END DEFAULT DEPTH & HEIGHT ########################################

' Trebor Tnemyar

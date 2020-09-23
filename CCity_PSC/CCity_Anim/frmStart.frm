VERSION 5.00
Begin VB.Form frmStart 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   LinkTopic       =   "Form2"
   Picture         =   "frmStart.frx":0000
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "X"
      Height          =   270
      Left            =   3765
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Quit "
      Top             =   435
      Width           =   300
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Load  *.ccc  file"
      Height          =   315
      Left            =   765
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2580
      Width           =   1485
   End
   Begin VB.CommandButton cmdGO 
      BackColor       =   &H00FFC0C0&
      Caption         =   "GO"
      Height          =   315
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2580
      Width           =   870
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

'------------------------------------------------------------------------------
' Ahmad Marwan Mami's notation (PSC CodeId=42081)
Private Declare Function SWR Lib "user32" Alias "SetWindowRgn" _
   (ByVal hWnd As Long, ByVal hrgn As Long, _
    ByVal bRedraw As Boolean) As Long
    
Private Declare Function CRR Lib "gdi32" Alias "CreateRoundRectRgn" _
  (ByVal XTL As Long, ByVal YTL As Long, _
   ByVal XBR As Long, ByVal YBR As Long, _
   ByVal EW As Long, ByVal EH As Long) As Long
   
   ' XTL,YTL  XBR,YBR  Top left & Bottom right coords of rectangle
   ' EW,EH  width & height of ellipse used to create corners
'-------------------------------------------------------------------------
Private CommonDialog1 As OSDialog

Private Sub cmdExit_Click()
Dim Form As Form
   Screen.MousePointer = vbDefault
   Sleep 500
   Erase BArray, BackArray
   ' Make sure all forms cleared
   For Each Form In Forms
      Unload Form
      Set Form = Nothing
   Next Form
   End
End Sub

Private Sub cmdGO_Click()
   HITS = 0
   Form1.picHITS.Cls
   Form1.picHITS.Print "0"
   Unload frmStart
End Sub

Private Sub cmdLoad_Click()
Dim Title$, Filt$, InDir$
   HITS = 0
   Form1.picHITS.Cls
   Form1.picHITS.Print "0"
   Set CommonDialog1 = New OSDialog
   
   Title$ = "Load City File"
   Filt$ = "Load ccc (*.ccc)|*.ccc"
   InDir$ = CCC_Path$ 'Pathspec$
   CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, InDir$, "", Me.hWnd
   
   If Len(FileSpec$) <> 0 Then
      CCC_Path$ = FileSpec$
      
      READ_CCC_FILE
      
      PyramidBlockNumber = 0
      If InStr(1, LCase$(FileSpec$), "maze") <> 0 Then
         AMAZE = True
         ' Find Maze PyramidBlockNumber
         For i = 1 To NumBoxes
            If R(i) = 16 Then
            If C(i) = 7 Then
            If LX(i) = 64 Then
            If LY(i) = 1 Then
            If LZ(i) = 66 Then
            If HX(i) = 195 Then
            If HY(i) = 65 Then
            If HZ(i) = 193 Then
               PyramidBlockNumber = i
               Exit For
            End If
            End If
            End If
            End If
            End If
            End If
            End If
            End If
         Next i
      Else
         AMAZE = False
      End If
      
   End If

   Set CommonDialog1 = Nothing
   Unload frmStart

End Sub


Private Sub READ_CCC_FILE()
'FileSpec$
   On Error GoTo FErr:
   Open FileSpec$ For Input As #1
   Line Input #1, FName$
   Input #1, NumBoxes
   Close
   
   ReDim LX(NumBoxes) As Long
   ReDim LY(NumBoxes)
   ReDim LZ(NumBoxes)
   ReDim HX(NumBoxes)
   ReDim HY(NumBoxes)
   ReDim HZ(NumBoxes)
   ReDim R(NumBoxes)
   ReDim C(NumBoxes)
   DoEvents
   
   
   Open FileSpec$ For Input As #1
   Line Input #1, FName$
   Input #1, NumBoxes
   For i = 1 To NumBoxes
      Input #1, LX(i), LY(i), LZ(i), HX(i), HY(i), HZ(i), R(i), C(i)
   Next i
   Close
   Exit Sub

FErr:
Close
MsgBox FileSpec$ & "  File error - Default will be taken", vbCritical, " City file input"
FileSpec$ = vbNull

End Sub


Private Sub Form_Load()
   
   Screen.MousePointer = vbDefault
   
   SWR Me.hWnd, CRR(6, 6, (Me.Width \ STX), (Me.Height \ STY), 50, 50), True

End Sub

Attribute VB_Name = "Module2"
' Module2 (Publics.bas) by Robert Rayment

' SET UP BOXES
' MOVE BOXES
' TRANSFORM TO 2D VIEW SCREEN

Option Explicit
Option Base 1

Public PlaneZ As Long   ' Z view plane
Public NumBoxes As Long

' Input Offset Box diagonals x=1-256, z=1-256, y=1-***
' L Low, H Top opp diags
Public LX() As Long
Public LY() As Long
Public LZ() As Long
Public HX() As Long
Public HY() As Long
Public HZ() As Long
' Box coords generated from HLXYZ()
Public BoxX() As Long
Public BoxY() As Long
Public BoxZ() As Long
' Row, Column & Height position of boxes ( x 256)
Public R() As Long
Public C() As Long

' Transformed Box to 2D
Public transx() As Long
Public transy() As Long
Public zDiv As Single
Public zNum As Single
' Motion steps
Public SumStepZ As Long
Public StepZ As Long
Public ZIncr As Long
Public MaxStepZ As Long

Public SumStepX As Long
Public StepX As Long
Public XIncr As Long
Public MaxStepX As Long

Public EyeIncr As Long
Public Max_eyeY As Long
Public Min_eyeY As Long

Public StepY As Long

' Array to hold lines from transx() & transy()
Public BArray() As Byte
Public BArrayWidth As Long
Public BArrayHeight As Long
Public WindowWidth  As Long
Public WindowHeight  As Long
Public BackArray() As Byte
Public BackArrayWidth As Long
Public BackArrayHeight As Long

' Eye x,y,z
Public eyeX As Long
Public eyeY As Long
Public eyeZ As Long

Public STX As Long
Public STY As Long

' For Bresline
Public ix As Long, iy As Long
Public idx As Long, idy As Long
Public jkstep As Long
Public incx As Long
Public id As Long
Public ainc As Long, binc As Long
Public JJ As Long, kk As Long

Public PathSpec$, CCC_Path$, FileSpec$
Public FName$

Public HITS As Long
Public AMAZE As Boolean
Public PyramidBlockNumber As Long


' General public
Public i As Long
Public j As Long
Public k As Long
Public n As Long
Public m As Long
'Public a$

'Public Const pi# = 3.1415926535898
'Public Const d2r# = pi# / 180
'Public Const r2d# = 1 / d2r#

Public Sub Make_Locate_Boxes()
' DEFAULT

' Arrays used for coords - faster than Type Structure

'                   Y            RZ
'                   |            /               7
'                - -|- - - - - -/- - - - - - - -
'               |/|/|/|/|/|/|/|/|/|/|/|/|/|/|/|6
'              - - -|- - - - -/- - - - - - - -
'             |/|/|/|/|/|/|/|/|/|/|/|/|/|/|/|5
'            - - - -|- - - -/- - - - - - - -
'           |/|/|/|/|/|/|/|/|/|/|/|/|/|/|/|4
'          - - - - -|- - -/- - - - - - - -
'         |/|/|/|/|/|/|/|/|/|/|/|/|/|/|/|3
'        - - - - - -|- -/- - - - - - - -
'       |/|/|/|/|/|/2/|/|/|/|/|/|/|/|/|2
'      - - - - - - -|-/- - - - - - - -
'     |/|/|/|/|/|/|/1/|/|/|/|/|/|/|/|1
'      - - - -0-1-2-3-4-5-6- - - - - -  - - CX
'   |/|/|/|/|/|/|/|/|/|/|/|/|/|/|/|0
'  - - - - - - - - - - - - - - - -
' |/|/|/|/|/|/|/|/|/|/|/|/|/|/|/-1


' Box numbering

'  Y
'  |
'  | 8-------P7H
'  |/|       /|  Z
'  4--------3 | /
'  | |      | |/
'  | 5------|-6
'  |/       |/
' P1L-------2---- X
'
' P1 = LX,Y,Z()  P7 = HX,Y,Z()
' R Row & C Column values ( x 256)

   NumBoxes = 180

   ReDim LX(NumBoxes)
   ReDim LY(NumBoxes)
   ReDim LZ(NumBoxes)
   ReDim HX(NumBoxes)
   ReDim HY(NumBoxes)
   ReDim HZ(NumBoxes)
   ' Locate cuboids
   ReDim R(NumBoxes)
   ReDim C(NumBoxes)

   ' Made with CCBuilder special save
   
   LX(1) = 178: LY(1) = 224: LZ(1) = 1: HX(1) = 182: HY(1) = 242: HZ(1) = 5
   R(1) = 16: C(1) = 0
   LX(2) = 170: LY(2) = 242: LZ(2) = 1: HX(2) = 189: HY(2) = 245: HZ(2) = 5
   R(2) = 16: C(2) = 0
   LX(3) = 143: LY(3) = 255: LZ(3) = 1: HX(3) = 158: HY(3) = 259: HZ(3) = 5
   R(3) = 16: C(3) = 0
   LX(4) = 138: LY(4) = 254: LZ(4) = 1: HX(4) = 142: HY(4) = 278: HZ(4) = 5
   R(4) = 16: C(4) = 0
   LX(5) = 120: LY(5) = 290: LZ(5) = 1: HX(5) = 124: HY(5) = 311: HZ(5) = 5
   R(5) = 16: C(5) = 0
   LX(6) = 111: LY(6) = 290: LZ(6) = 1: HX(6) = 121: HY(6) = 294: HZ(6) = 5
   R(6) = 16: C(6) = 0
   LX(7) = 107: LY(7) = 294: LZ(7) = 1: HX(7) = 111: HY(7) = 311: HZ(7) = 5
   R(7) = 16: C(7) = 0
   LX(8) = 81: LY(8) = 336: LZ(8) = 1: HX(8) = 86: HY(8) = 338: HZ(8) = 5
   R(8) = 16: C(8) = 0
   LX(9) = 90: LY(9) = 321: LZ(9) = 1: HX(9) = 94: HY(9) = 335: HZ(9) = 5
   R(9) = 16: C(9) = 0
   LX(10) = 88: LY(10) = 336: LZ(10) = 1: HX(10) = 91: HY(10) = 347: HZ(10) = 5
   R(10) = 16: C(10) = 0
   LX(11) = 81: LY(11) = 347: LZ(11) = 1: HX(11) = 87: HY(11) = 349: HZ(11) = 5
   R(11) = 16: C(11) = 0
   LX(12) = 78: LY(12) = 335: LZ(12) = 1: HX(12) = 82: HY(12) = 346: HZ(12) = 5
   R(12) = 16: C(12) = 0
   LX(13) = 74: LY(13) = 321: LZ(13) = 1: HX(13) = 78: HY(13) = 334: HZ(13) = 5
   R(13) = 16: C(13) = 0
   LX(14) = 57: LY(14) = 363: LZ(14) = 1: HX(14) = 64: HY(14) = 368: HZ(14) = 5
   R(14) = 16: C(14) = 0
   LX(15) = 56: LY(15) = 375: LZ(15) = 1: HX(15) = 67: HY(15) = 378: HZ(15) = 5
   R(15) = 16: C(15) = 0
   LX(16) = 53: LY(16) = 351: LZ(16) = 1: HX(16) = 57: HY(16) = 377: HZ(16) = 5
   R(16) = 16: C(16) = 0
   LX(17) = 31: LY(17) = 387: LZ(17) = 1: HX(17) = 40: HY(17) = 391: HZ(17) = 5
   R(17) = 16: C(17) = 0
   LX(18) = 32: LY(18) = 398: LZ(18) = 1: HX(18) = 36: HY(18) = 401: HZ(18) = 5
   R(18) = 16: C(18) = 0
   LX(19) = 32: LY(19) = 406: LZ(19) = 1: HX(19) = 41: HY(19) = 410: HZ(19) = 5
   R(19) = 16: C(19) = 0
   LX(20) = 26: LY(20) = 387: LZ(20) = 1: HX(20) = 31: HY(20) = 410: HZ(20) = 5
   R(20) = 16: C(20) = 0
   LX(21) = 6: LY(21) = 416: LZ(21) = 1: HX(21) = 14: HY(21) = 421: HZ(21) = 5
   R(21) = 16: C(21) = 0
   LX(22) = 13: LY(22) = 422: LZ(22) = 1: HX(22) = 17: HY(22) = 444: HZ(22) = 5
   R(22) = 16: C(22) = 0
   LX(23) = 5: LY(23) = 444: LZ(23) = 1: HX(23) = 12: HY(23) = 448: HZ(23) = 5
   R(23) = 16: C(23) = 0
   LX(24) = 1: LY(24) = 416: LZ(24) = 1: HX(24) = 5: HY(24) = 448: HZ(24) = 5
   R(24) = 16: C(24) = 0
   LX(25) = 178: LY(25) = 224: LZ(25) = 1: HX(25) = 182: HY(25) = 242: HZ(25) = 5
   R(25) = 14: C(25) = 2
   LX(26) = 170: LY(26) = 242: LZ(26) = 1: HX(26) = 189: HY(26) = 245: HZ(26) = 5
   R(26) = 14: C(26) = 2
   LX(27) = 143: LY(27) = 255: LZ(27) = 1: HX(27) = 158: HY(27) = 259: HZ(27) = 5
   R(27) = 14: C(27) = 2
   LX(28) = 138: LY(28) = 254: LZ(28) = 1: HX(28) = 142: HY(28) = 278: HZ(28) = 5
   R(28) = 14: C(28) = 2
   LX(29) = 120: LY(29) = 290: LZ(29) = 1: HX(29) = 124: HY(29) = 311: HZ(29) = 5
   R(29) = 14: C(29) = 2
   LX(30) = 111: LY(30) = 290: LZ(30) = 1: HX(30) = 121: HY(30) = 294: HZ(30) = 5
   R(30) = 14: C(30) = 2
   LX(31) = 107: LY(31) = 294: LZ(31) = 1: HX(31) = 111: HY(31) = 311: HZ(31) = 5
   R(31) = 14: C(31) = 2
   LX(32) = 81: LY(32) = 336: LZ(32) = 1: HX(32) = 86: HY(32) = 338: HZ(32) = 5
   R(32) = 14: C(32) = 2
   LX(33) = 90: LY(33) = 321: LZ(33) = 1: HX(33) = 94: HY(33) = 335: HZ(33) = 5
   R(33) = 14: C(33) = 2
   LX(34) = 88: LY(34) = 336: LZ(34) = 1: HX(34) = 91: HY(34) = 347: HZ(34) = 5
   R(34) = 14: C(34) = 2
   LX(35) = 81: LY(35) = 347: LZ(35) = 1: HX(35) = 87: HY(35) = 349: HZ(35) = 5
   R(35) = 14: C(35) = 2
   LX(36) = 78: LY(36) = 335: LZ(36) = 1: HX(36) = 82: HY(36) = 346: HZ(36) = 5
   R(36) = 14: C(36) = 2
   LX(37) = 74: LY(37) = 321: LZ(37) = 1: HX(37) = 78: HY(37) = 334: HZ(37) = 5
   R(37) = 14: C(37) = 2
   LX(38) = 57: LY(38) = 363: LZ(38) = 1: HX(38) = 64: HY(38) = 368: HZ(38) = 5
   R(38) = 14: C(38) = 2
   LX(39) = 56: LY(39) = 375: LZ(39) = 1: HX(39) = 67: HY(39) = 378: HZ(39) = 5
   R(39) = 14: C(39) = 2
   LX(40) = 53: LY(40) = 351: LZ(40) = 1: HX(40) = 57: HY(40) = 377: HZ(40) = 5
   R(40) = 14: C(40) = 2
   LX(41) = 31: LY(41) = 387: LZ(41) = 1: HX(41) = 40: HY(41) = 391: HZ(41) = 5
   R(41) = 14: C(41) = 2
   LX(42) = 32: LY(42) = 398: LZ(42) = 1: HX(42) = 36: HY(42) = 401: HZ(42) = 5
   R(42) = 14: C(42) = 2
   LX(43) = 32: LY(43) = 406: LZ(43) = 1: HX(43) = 41: HY(43) = 410: HZ(43) = 5
   R(43) = 14: C(43) = 2
   LX(44) = 26: LY(44) = 387: LZ(44) = 1: HX(44) = 31: HY(44) = 410: HZ(44) = 5
   R(44) = 14: C(44) = 2
   LX(45) = 6: LY(45) = 416: LZ(45) = 1: HX(45) = 14: HY(45) = 421: HZ(45) = 5
   R(45) = 14: C(45) = 2
   LX(46) = 13: LY(46) = 422: LZ(46) = 1: HX(46) = 17: HY(46) = 444: HZ(46) = 5
   R(46) = 14: C(46) = 2
   LX(47) = 5: LY(47) = 444: LZ(47) = 1: HX(47) = 12: HY(47) = 448: HZ(47) = 5
   R(47) = 14: C(47) = 2
   LX(48) = 1: LY(48) = 416: LZ(48) = 1: HX(48) = 5: HY(48) = 448: HZ(48) = 5
   R(48) = 14: C(48) = 2
   LX(49) = 178: LY(49) = 224: LZ(49) = 1: HX(49) = 182: HY(49) = 242: HZ(49) = 5
   R(49) = 12: C(49) = 4
   LX(50) = 170: LY(50) = 242: LZ(50) = 1: HX(50) = 189: HY(50) = 245: HZ(50) = 5
   R(50) = 12: C(50) = 4
   LX(51) = 143: LY(51) = 255: LZ(51) = 1: HX(51) = 158: HY(51) = 259: HZ(51) = 5
   R(51) = 12: C(51) = 4
   LX(52) = 138: LY(52) = 254: LZ(52) = 1: HX(52) = 142: HY(52) = 278: HZ(52) = 5
   R(52) = 12: C(52) = 4
   LX(53) = 120: LY(53) = 290: LZ(53) = 1: HX(53) = 124: HY(53) = 311: HZ(53) = 5
   R(53) = 12: C(53) = 4
   LX(54) = 111: LY(54) = 290: LZ(54) = 1: HX(54) = 121: HY(54) = 294: HZ(54) = 5
   R(54) = 12: C(54) = 4
   LX(55) = 107: LY(55) = 294: LZ(55) = 1: HX(55) = 111: HY(55) = 311: HZ(55) = 5
   R(55) = 12: C(55) = 4
   LX(56) = 81: LY(56) = 336: LZ(56) = 1: HX(56) = 86: HY(56) = 338: HZ(56) = 5
   R(56) = 12: C(56) = 4
   LX(57) = 90: LY(57) = 321: LZ(57) = 1: HX(57) = 94: HY(57) = 335: HZ(57) = 5
   R(57) = 12: C(57) = 4
   LX(58) = 88: LY(58) = 336: LZ(58) = 1: HX(58) = 91: HY(58) = 347: HZ(58) = 5
   R(58) = 12: C(58) = 4
   LX(59) = 81: LY(59) = 347: LZ(59) = 1: HX(59) = 87: HY(59) = 349: HZ(59) = 5
   R(59) = 12: C(59) = 4
   LX(60) = 78: LY(60) = 335: LZ(60) = 1: HX(60) = 82: HY(60) = 346: HZ(60) = 5
   R(60) = 12: C(60) = 4
   LX(61) = 74: LY(61) = 321: LZ(61) = 1: HX(61) = 78: HY(61) = 334: HZ(61) = 5
   R(61) = 12: C(61) = 4
   LX(62) = 57: LY(62) = 363: LZ(62) = 1: HX(62) = 64: HY(62) = 368: HZ(62) = 5
   R(62) = 12: C(62) = 4
   LX(63) = 56: LY(63) = 375: LZ(63) = 1: HX(63) = 67: HY(63) = 378: HZ(63) = 5
   R(63) = 12: C(63) = 4
   LX(64) = 53: LY(64) = 351: LZ(64) = 1: HX(64) = 57: HY(64) = 377: HZ(64) = 5
   R(64) = 12: C(64) = 4
   LX(65) = 31: LY(65) = 387: LZ(65) = 1: HX(65) = 40: HY(65) = 391: HZ(65) = 5
   R(65) = 12: C(65) = 4
   LX(66) = 32: LY(66) = 398: LZ(66) = 1: HX(66) = 36: HY(66) = 401: HZ(66) = 5
   R(66) = 12: C(66) = 4
   LX(67) = 32: LY(67) = 406: LZ(67) = 1: HX(67) = 41: HY(67) = 410: HZ(67) = 5
   R(67) = 12: C(67) = 4
   LX(68) = 26: LY(68) = 387: LZ(68) = 1: HX(68) = 31: HY(68) = 410: HZ(68) = 5
   R(68) = 12: C(68) = 4
   LX(69) = 6: LY(69) = 416: LZ(69) = 1: HX(69) = 14: HY(69) = 421: HZ(69) = 5
   R(69) = 12: C(69) = 4
   LX(70) = 13: LY(70) = 422: LZ(70) = 1: HX(70) = 17: HY(70) = 444: HZ(70) = 5
   R(70) = 12: C(70) = 4
   LX(71) = 5: LY(71) = 444: LZ(71) = 1: HX(71) = 12: HY(71) = 448: HZ(71) = 5
   R(71) = 12: C(71) = 4
   LX(72) = 1: LY(72) = 416: LZ(72) = 1: HX(72) = 5: HY(72) = 448: HZ(72) = 5
   R(72) = 12: C(72) = 4
   LX(73) = 1: LY(73) = 416: LZ(73) = 1: HX(73) = 5: HY(73) = 448: HZ(73) = 5
   R(73) = 10: C(73) = 6
   LX(74) = 5: LY(74) = 444: LZ(74) = 1: HX(74) = 12: HY(74) = 448: HZ(74) = 5
   R(74) = 10: C(74) = 6
   LX(75) = 13: LY(75) = 422: LZ(75) = 1: HX(75) = 17: HY(75) = 444: HZ(75) = 5
   R(75) = 10: C(75) = 6
   LX(76) = 6: LY(76) = 416: LZ(76) = 1: HX(76) = 14: HY(76) = 421: HZ(76) = 5
   R(76) = 10: C(76) = 6
   LX(77) = 26: LY(77) = 387: LZ(77) = 1: HX(77) = 31: HY(77) = 410: HZ(77) = 5
   R(77) = 10: C(77) = 6
   LX(78) = 32: LY(78) = 406: LZ(78) = 1: HX(78) = 41: HY(78) = 410: HZ(78) = 5
   R(78) = 10: C(78) = 6
   LX(79) = 32: LY(79) = 398: LZ(79) = 1: HX(79) = 36: HY(79) = 401: HZ(79) = 5
   R(79) = 10: C(79) = 6
   LX(80) = 31: LY(80) = 387: LZ(80) = 1: HX(80) = 40: HY(80) = 391: HZ(80) = 5
   R(80) = 10: C(80) = 6
   LX(81) = 53: LY(81) = 351: LZ(81) = 1: HX(81) = 57: HY(81) = 377: HZ(81) = 5
   R(81) = 10: C(81) = 6
   LX(82) = 56: LY(82) = 375: LZ(82) = 1: HX(82) = 67: HY(82) = 378: HZ(82) = 5
   R(82) = 10: C(82) = 6
   LX(83) = 57: LY(83) = 363: LZ(83) = 1: HX(83) = 64: HY(83) = 368: HZ(83) = 5
   R(83) = 10: C(83) = 6
   LX(84) = 74: LY(84) = 321: LZ(84) = 1: HX(84) = 78: HY(84) = 334: HZ(84) = 5
   R(84) = 10: C(84) = 6
   LX(85) = 78: LY(85) = 335: LZ(85) = 1: HX(85) = 82: HY(85) = 346: HZ(85) = 5
   R(85) = 10: C(85) = 6
   LX(86) = 81: LY(86) = 347: LZ(86) = 1: HX(86) = 87: HY(86) = 349: HZ(86) = 5
   R(86) = 10: C(86) = 6
   LX(87) = 88: LY(87) = 336: LZ(87) = 1: HX(87) = 91: HY(87) = 347: HZ(87) = 5
   R(87) = 10: C(87) = 6
   LX(88) = 90: LY(88) = 321: LZ(88) = 1: HX(88) = 94: HY(88) = 335: HZ(88) = 5
   R(88) = 10: C(88) = 6
   LX(89) = 81: LY(89) = 336: LZ(89) = 1: HX(89) = 86: HY(89) = 338: HZ(89) = 5
   R(89) = 10: C(89) = 6
   LX(90) = 107: LY(90) = 294: LZ(90) = 1: HX(90) = 111: HY(90) = 311: HZ(90) = 5
   R(90) = 10: C(90) = 6
   LX(91) = 111: LY(91) = 290: LZ(91) = 1: HX(91) = 121: HY(91) = 294: HZ(91) = 5
   R(91) = 10: C(91) = 6
   LX(92) = 120: LY(92) = 290: LZ(92) = 1: HX(92) = 124: HY(92) = 311: HZ(92) = 5
   R(92) = 10: C(92) = 6
   LX(93) = 138: LY(93) = 254: LZ(93) = 1: HX(93) = 142: HY(93) = 278: HZ(93) = 5
   R(93) = 10: C(93) = 6
   LX(94) = 143: LY(94) = 255: LZ(94) = 1: HX(94) = 158: HY(94) = 259: HZ(94) = 5
   R(94) = 10: C(94) = 6
   LX(95) = 170: LY(95) = 242: LZ(95) = 1: HX(95) = 189: HY(95) = 245: HZ(95) = 5
   R(95) = 10: C(95) = 6
   LX(96) = 178: LY(96) = 224: LZ(96) = 1: HX(96) = 182: HY(96) = 242: HZ(96) = 5
   R(96) = 10: C(96) = 6
   LX(97) = 31: LY(97) = 65: LZ(97) = 1: HX(97) = 46: HY(97) = 160: HZ(97) = 5
   R(97) = 10: C(97) = 1
   LX(98) = 46: LY(98) = 151: LZ(98) = 1: HX(98) = 64: HY(98) = 160: HZ(98) = 5
   R(98) = 10: C(98) = 1
   LX(99) = 64: LY(99) = 119: LZ(99) = 1: HX(99) = 72: HY(99) = 151: HZ(99) = 5
   R(99) = 10: C(99) = 1
   LX(100) = 47: LY(100) = 112: LZ(100) = 1: HX(100) = 65: HY(100) = 119: HZ(100) = 5
   R(100) = 10: C(100) = 1
   LX(101) = 64: LY(101) = 73: LZ(101) = 1: HX(101) = 72: HY(101) = 112: HZ(101) = 5
   R(101) = 10: C(101) = 1
   LX(102) = 63: LY(102) = 66: LZ(102) = 1: HX(102) = 79: HY(102) = 74: HZ(102) = 5
   R(102) = 10: C(102) = 1
   LX(103) = 112: LY(103) = 64: LZ(103) = 1: HX(103) = 127: HY(103) = 160: HZ(103) = 5
   R(103) = 10: C(103) = 1
   LX(104) = 127: LY(104) = 151: LZ(104) = 1: HX(104) = 145: HY(104) = 160: HZ(104) = 5
   R(104) = 10: C(104) = 1
   LX(105) = 145: LY(105) = 119: LZ(105) = 1: HX(105) = 153: HY(105) = 151: HZ(105) = 5
   R(105) = 10: C(105) = 1
   LX(106) = 127: LY(106) = 112: LZ(106) = 1: HX(106) = 145: HY(106) = 119: HZ(106) = 5
   R(106) = 10: C(106) = 1
   LX(107) = 146: LY(107) = 73: LZ(107) = 1: HX(107) = 154: HY(107) = 112: HZ(107) = 5
   R(107) = 10: C(107) = 1
   LX(108) = 146: LY(108) = 65: LZ(108) = 1: HX(108) = 162: HY(108) = 73: HZ(108) = 5
   R(108) = 10: C(108) = 1
   LX(109) = 53: LY(109) = 351: LZ(109) = 1: HX(109) = 57: HY(109) = 377: HZ(109) = 5
   R(109) = 8: C(109) = 4
   LX(110) = 56: LY(110) = 375: LZ(110) = 1: HX(110) = 67: HY(110) = 378: HZ(110) = 5
   R(110) = 8: C(110) = 4
   LX(111) = 32: LY(111) = 398: LZ(111) = 1: HX(111) = 36: HY(111) = 401: HZ(111) = 5
   R(111) = 8: C(111) = 4
   LX(112) = 178: LY(112) = 224: LZ(112) = 1: HX(112) = 182: HY(112) = 242: HZ(112) = 5
   R(112) = 8: C(112) = 4
   LX(113) = 138: LY(113) = 254: LZ(113) = 1: HX(113) = 142: HY(113) = 278: HZ(113) = 5
   R(113) = 8: C(113) = 4
   LX(114) = 143: LY(114) = 255: LZ(114) = 1: HX(114) = 158: HY(114) = 259: HZ(114) = 5
   R(114) = 8: C(114) = 4
   LX(115) = 88: LY(115) = 336: LZ(115) = 1: HX(115) = 91: HY(115) = 347: HZ(115) = 5
   R(115) = 8: C(115) = 4
   LX(116) = 74: LY(116) = 321: LZ(116) = 1: HX(116) = 78: HY(116) = 334: HZ(116) = 5
   R(116) = 8: C(116) = 4
   LX(117) = 78: LY(117) = 335: LZ(117) = 1: HX(117) = 82: HY(117) = 346: HZ(117) = 5
   R(117) = 8: C(117) = 4
   LX(118) = 31: LY(118) = 387: LZ(118) = 1: HX(118) = 40: HY(118) = 391: HZ(118) = 5
   R(118) = 8: C(118) = 4
   LX(119) = 57: LY(119) = 363: LZ(119) = 1: HX(119) = 64: HY(119) = 368: HZ(119) = 5
   R(119) = 8: C(119) = 4
   LX(120) = 81: LY(120) = 347: LZ(120) = 1: HX(120) = 87: HY(120) = 349: HZ(120) = 5
   R(120) = 8: C(120) = 4
   LX(121) = 1: LY(121) = 416: LZ(121) = 1: HX(121) = 5: HY(121) = 448: HZ(121) = 5
   R(121) = 8: C(121) = 4
   LX(122) = 5: LY(122) = 444: LZ(122) = 1: HX(122) = 12: HY(122) = 448: HZ(122) = 5
   R(122) = 8: C(122) = 4
   LX(123) = 13: LY(123) = 422: LZ(123) = 1: HX(123) = 17: HY(123) = 444: HZ(123) = 5
   R(123) = 8: C(123) = 4
   LX(124) = 6: LY(124) = 416: LZ(124) = 1: HX(124) = 14: HY(124) = 421: HZ(124) = 5
   R(124) = 8: C(124) = 4
   LX(125) = 26: LY(125) = 387: LZ(125) = 1: HX(125) = 31: HY(125) = 410: HZ(125) = 5
   R(125) = 8: C(125) = 4
   LX(126) = 32: LY(126) = 406: LZ(126) = 1: HX(126) = 41: HY(126) = 410: HZ(126) = 5
   R(126) = 8: C(126) = 4
   LX(127) = 81: LY(127) = 336: LZ(127) = 1: HX(127) = 86: HY(127) = 338: HZ(127) = 5
   R(127) = 8: C(127) = 4
   LX(128) = 120: LY(128) = 290: LZ(128) = 1: HX(128) = 124: HY(128) = 311: HZ(128) = 5
   R(128) = 8: C(128) = 4
   LX(129) = 170: LY(129) = 242: LZ(129) = 1: HX(129) = 189: HY(129) = 245: HZ(129) = 5
   R(129) = 8: C(129) = 4
   LX(130) = 107: LY(130) = 294: LZ(130) = 1: HX(130) = 111: HY(130) = 311: HZ(130) = 5
   R(130) = 8: C(130) = 4
   LX(131) = 111: LY(131) = 290: LZ(131) = 1: HX(131) = 121: HY(131) = 294: HZ(131) = 5
   R(131) = 8: C(131) = 4
   LX(132) = 90: LY(132) = 321: LZ(132) = 1: HX(132) = 94: HY(132) = 335: HZ(132) = 5
   R(132) = 8: C(132) = 4
   LX(133) = 53: LY(133) = 351: LZ(133) = 1: HX(133) = 57: HY(133) = 377: HZ(133) = 5
   R(133) = 6: C(133) = 2
   LX(134) = 56: LY(134) = 375: LZ(134) = 1: HX(134) = 67: HY(134) = 378: HZ(134) = 5
   R(134) = 6: C(134) = 2
   LX(135) = 32: LY(135) = 398: LZ(135) = 1: HX(135) = 36: HY(135) = 401: HZ(135) = 5
   R(135) = 6: C(135) = 2
   LX(136) = 78: LY(136) = 335: LZ(136) = 1: HX(136) = 82: HY(136) = 346: HZ(136) = 5
   R(136) = 6: C(136) = 2
   LX(137) = 90: LY(137) = 321: LZ(137) = 1: HX(137) = 94: HY(137) = 335: HZ(137) = 5
   R(137) = 6: C(137) = 2
   LX(138) = 88: LY(138) = 336: LZ(138) = 1: HX(138) = 91: HY(138) = 347: HZ(138) = 5
   R(138) = 6: C(138) = 2
   LX(139) = 143: LY(139) = 255: LZ(139) = 1: HX(139) = 158: HY(139) = 259: HZ(139) = 5
   R(139) = 6: C(139) = 2
   LX(140) = 74: LY(140) = 321: LZ(140) = 1: HX(140) = 78: HY(140) = 334: HZ(140) = 5
   R(140) = 6: C(140) = 2
   LX(141) = 178: LY(141) = 224: LZ(141) = 1: HX(141) = 182: HY(141) = 242: HZ(141) = 5
   R(141) = 6: C(141) = 2
   LX(142) = 31: LY(142) = 387: LZ(142) = 1: HX(142) = 40: HY(142) = 391: HZ(142) = 5
   R(142) = 6: C(142) = 2
   LX(143) = 57: LY(143) = 363: LZ(143) = 1: HX(143) = 64: HY(143) = 368: HZ(143) = 5
   R(143) = 6: C(143) = 2
   LX(144) = 170: LY(144) = 242: LZ(144) = 1: HX(144) = 189: HY(144) = 245: HZ(144) = 5
   R(144) = 6: C(144) = 2
   LX(145) = 1: LY(145) = 416: LZ(145) = 1: HX(145) = 5: HY(145) = 448: HZ(145) = 5
   R(145) = 6: C(145) = 2
   LX(146) = 5: LY(146) = 444: LZ(146) = 1: HX(146) = 12: HY(146) = 448: HZ(146) = 5
   R(146) = 6: C(146) = 2
   LX(147) = 13: LY(147) = 422: LZ(147) = 1: HX(147) = 17: HY(147) = 444: HZ(147) = 5
   R(147) = 6: C(147) = 2
   LX(148) = 6: LY(148) = 416: LZ(148) = 1: HX(148) = 14: HY(148) = 421: HZ(148) = 5
   R(148) = 6: C(148) = 2
   LX(149) = 26: LY(149) = 387: LZ(149) = 1: HX(149) = 31: HY(149) = 410: HZ(149) = 5
   R(149) = 6: C(149) = 2
   LX(150) = 32: LY(150) = 406: LZ(150) = 1: HX(150) = 41: HY(150) = 410: HZ(150) = 5
   R(150) = 6: C(150) = 2
   LX(151) = 120: LY(151) = 290: LZ(151) = 1: HX(151) = 124: HY(151) = 311: HZ(151) = 5
   R(151) = 6: C(151) = 2
   LX(152) = 81: LY(152) = 336: LZ(152) = 1: HX(152) = 86: HY(152) = 338: HZ(152) = 5
   R(152) = 6: C(152) = 2
   LX(153) = 81: LY(153) = 347: LZ(153) = 1: HX(153) = 87: HY(153) = 349: HZ(153) = 5
   R(153) = 6: C(153) = 2
   LX(154) = 111: LY(154) = 290: LZ(154) = 1: HX(154) = 121: HY(154) = 294: HZ(154) = 5
   R(154) = 6: C(154) = 2
   LX(155) = 107: LY(155) = 294: LZ(155) = 1: HX(155) = 111: HY(155) = 311: HZ(155) = 5
   R(155) = 6: C(155) = 2
   LX(156) = 138: LY(156) = 254: LZ(156) = 1: HX(156) = 142: HY(156) = 278: HZ(156) = 5
   R(156) = 6: C(156) = 2
   LX(157) = 53: LY(157) = 351: LZ(157) = 1: HX(157) = 57: HY(157) = 377: HZ(157) = 5
   R(157) = 4: C(157) = 0
   LX(158) = 56: LY(158) = 375: LZ(158) = 1: HX(158) = 67: HY(158) = 378: HZ(158) = 5
   R(158) = 4: C(158) = 0
   LX(159) = 32: LY(159) = 398: LZ(159) = 1: HX(159) = 36: HY(159) = 401: HZ(159) = 5
   R(159) = 4: C(159) = 0
   LX(160) = 178: LY(160) = 224: LZ(160) = 1: HX(160) = 182: HY(160) = 242: HZ(160) = 5
   R(160) = 4: C(160) = 0
   LX(161) = 138: LY(161) = 254: LZ(161) = 1: HX(161) = 142: HY(161) = 278: HZ(161) = 5
   R(161) = 4: C(161) = 0
   LX(162) = 143: LY(162) = 255: LZ(162) = 1: HX(162) = 158: HY(162) = 259: HZ(162) = 5
   R(162) = 4: C(162) = 0
   LX(163) = 88: LY(163) = 336: LZ(163) = 1: HX(163) = 91: HY(163) = 347: HZ(163) = 5
   R(163) = 4: C(163) = 0
   LX(164) = 74: LY(164) = 321: LZ(164) = 1: HX(164) = 78: HY(164) = 334: HZ(164) = 5
   R(164) = 4: C(164) = 0
   LX(165) = 78: LY(165) = 335: LZ(165) = 1: HX(165) = 82: HY(165) = 346: HZ(165) = 5
   R(165) = 4: C(165) = 0
   LX(166) = 31: LY(166) = 387: LZ(166) = 1: HX(166) = 40: HY(166) = 391: HZ(166) = 5
   R(166) = 4: C(166) = 0
   LX(167) = 57: LY(167) = 363: LZ(167) = 1: HX(167) = 64: HY(167) = 368: HZ(167) = 5
   R(167) = 4: C(167) = 0
   LX(168) = 81: LY(168) = 347: LZ(168) = 1: HX(168) = 87: HY(168) = 349: HZ(168) = 5
   R(168) = 4: C(168) = 0
   LX(169) = 1: LY(169) = 416: LZ(169) = 1: HX(169) = 5: HY(169) = 448: HZ(169) = 5
   R(169) = 4: C(169) = 0
   LX(170) = 5: LY(170) = 444: LZ(170) = 1: HX(170) = 12: HY(170) = 448: HZ(170) = 5
   R(170) = 4: C(170) = 0
   LX(171) = 13: LY(171) = 422: LZ(171) = 1: HX(171) = 17: HY(171) = 444: HZ(171) = 5
   R(171) = 4: C(171) = 0
   LX(172) = 6: LY(172) = 416: LZ(172) = 1: HX(172) = 14: HY(172) = 421: HZ(172) = 5
   R(172) = 4: C(172) = 0
   LX(173) = 26: LY(173) = 387: LZ(173) = 1: HX(173) = 31: HY(173) = 410: HZ(173) = 5
   R(173) = 4: C(173) = 0
   LX(174) = 32: LY(174) = 406: LZ(174) = 1: HX(174) = 41: HY(174) = 410: HZ(174) = 5
   R(174) = 4: C(174) = 0
   LX(175) = 81: LY(175) = 336: LZ(175) = 1: HX(175) = 86: HY(175) = 338: HZ(175) = 5
   R(175) = 4: C(175) = 0
   LX(176) = 120: LY(176) = 290: LZ(176) = 1: HX(176) = 124: HY(176) = 311: HZ(176) = 5
   R(176) = 4: C(176) = 0
   LX(177) = 170: LY(177) = 242: LZ(177) = 1: HX(177) = 189: HY(177) = 245: HZ(177) = 5
   R(177) = 4: C(177) = 0
   LX(178) = 107: LY(178) = 294: LZ(178) = 1: HX(178) = 111: HY(178) = 311: HZ(178) = 5
   R(178) = 4: C(178) = 0
   LX(179) = 111: LY(179) = 290: LZ(179) = 1: HX(179) = 121: HY(179) = 294: HZ(179) = 5
   R(179) = 4: C(179) = 0
   LX(180) = 90: LY(180) = 321: LZ(180) = 1: HX(180) = 94: HY(180) = 335: HZ(180) = 5
   R(180) = 4: C(180) = 0

End Sub

Public Sub FillBoxes()
' R() & C() set
' From LHX,Y,Z() generate all BoxX,Y,Z(8, NumBoxes)
ReDim BoxX(8, NumBoxes)
ReDim BoxY(8, NumBoxes)
ReDim BoxZ(8, NumBoxes)
'  Y
'  |
'  | 8-------P7H
'  |/|       /|  Z
'  4--------3 | /
'  | |      | |/
'  | 5------|-6
'  |/       |/
' P1L-------2---- X
   For n = 1 To NumBoxes
      ' Front face
      ' Pt 1
      BoxX(1, n) = LX(n) + (C(n) - 1) * 256
      BoxY(1, n) = LY(n)
      BoxZ(1, n) = LZ(n) + (R(n) - 1) * 256
      ' Pt 2
      BoxX(2, n) = HX(n) + (C(n) - 1) * 256
      BoxY(2, n) = BoxY(1, n)
      BoxZ(2, n) = BoxZ(1, n)
      ' Pt 3
      BoxX(3, n) = BoxX(2, n)
      BoxY(3, n) = HY(n)
      BoxZ(3, n) = BoxZ(1, n)
      ' Pt 4
      BoxX(4, n) = BoxX(1, n)
      BoxY(4, n) = BoxY(3, n)
      BoxZ(4, n) = BoxZ(1, n)
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
      BoxX(5, n) = BoxX(1, n)
      BoxY(5, n) = LY(n)
      BoxZ(5, n) = HZ(n) + (R(n) - 1) * 256
      ' Pt 6
      BoxX(6, n) = BoxX(2, n)
      BoxY(6, n) = BoxY(5, n)
      BoxZ(6, n) = BoxZ(5, n)
      ' Pt 7
      BoxX(7, n) = BoxX(6, n)
      BoxY(7, n) = HY(n)
      BoxZ(7, n) = BoxZ(5, n)
      ' Pt 8
      BoxX(8, n) = BoxX(5, n)
      BoxY(8, n) = BoxY(7, n)
      BoxZ(8, n) = BoxZ(5, n)
   Next n
End Sub


Public Sub IncrX()
'Public StepX
   For n = 1 To NumBoxes   ' 1->1, 2->3
   For i = 1 To 8
      BoxX(i, n) = BoxX(i, n) + StepX
   Next i
   Next n
   
   SumStepX = SumStepX + StepX
   
   If SumStepX <= -256 Then     ' Boxes moving left, StepX -ve
      SumStepX = SumStepX + 256
      For n = 1 To NumBoxes
         C(n) = C(n) - 1
         If C(n) < -5 Then
            C(n) = 12
            k = (C(n) - 1) * 256 + SumStepX
            BoxX(1, n) = LX(n) + k
            BoxX(2, n) = HX(n) + k
            BoxX(3, n) = BoxX(2, n)
            BoxX(4, n) = BoxX(1, n)
            BoxX(5, n) = BoxX(1, n)
            BoxX(6, n) = BoxX(2, n)
            BoxX(7, n) = BoxX(2, n)
            BoxX(8, n) = BoxX(1, n)
         End If
      Next n
'  Y
'  | 8-------P7H
'  |/|       /|  Z
'  4--------3 | /
'  | 5------|-6
'  |/       |/
' P1L-------2---- X
   ElseIf SumStepX >= 256 Then
      SumStepX = SumStepX - 256     ' Boxes moving right, StepX +ve
      For n = 1 To NumBoxes
         C(n) = C(n) + 1
         If C(n) > 12 Then
            C(n) = -5
            k = (C(n) - 1) * 256 + SumStepX
            BoxX(1, n) = LX(n) + k
            BoxX(2, n) = HX(n) + k
            BoxX(3, n) = BoxX(2, n)
            BoxX(4, n) = BoxX(1, n)
            BoxX(5, n) = BoxX(1, n)
            BoxX(6, n) = BoxX(2, n)
            BoxX(7, n) = BoxX(2, n)
            BoxX(8, n) = BoxX(1, n)
         End If
      Next n
   End If
End Sub

Public Sub IncrZ()
'Public StepZ
   For n = 1 To NumBoxes   ' 1->1, 2->3
      For i = 1 To 8
         BoxZ(i, n) = BoxZ(i, n) + StepZ
      Next i
   Next n
   
   SumStepZ = SumStepZ + StepZ
   If SumStepZ <= -256 Then      ' Boxes going down screen, StepZ -ve
      SumStepZ = SumStepZ + 256
      For n = 1 To NumBoxes
         R(n) = R(n) - 1
         If R(n) < 0 Then
            R(n) = 21
            k = (R(n) - 1) * 256 + SumStepZ
            BoxZ(1, n) = LZ(n) + k
            BoxZ(2, n) = BoxZ(1, n)
            BoxZ(3, n) = BoxZ(1, n)
            BoxZ(4, n) = BoxZ(1, n)
            BoxZ(5, n) = HZ(n) + k
            BoxZ(6, n) = BoxZ(5, n)
            BoxZ(7, n) = BoxZ(5, n)
            BoxZ(8, n) = BoxZ(5, n)
         End If
      Next n
'  Y
'  | 8-------P7H
'  |/|       /|  Z
'  4--------3 | /
'  | 5------|-6
'  |/       |/
' P1L-------2---- X
   ElseIf SumStepZ >= 256 Then   ' Boxes going up screen, StepZ +ve
      SumStepZ = SumStepZ - 256
      For n = 1 To NumBoxes
         R(n) = R(n) + 1
         If R(n) > 21 Then
            R(n) = 0
            k = (R(n) - 1) * 256 + SumStepZ
            BoxZ(1, n) = LZ(n) + k
            BoxZ(2, n) = BoxZ(1, n)
            BoxZ(3, n) = BoxZ(1, n)
            BoxZ(4, n) = BoxZ(1, n)
            BoxZ(5, n) = HZ(n) + k
            BoxZ(6, n) = BoxZ(5, n)
            BoxZ(7, n) = BoxZ(5, n)
            BoxZ(8, n) = BoxZ(5, n)
         End If
      Next n
   
   End If
End Sub

Public Sub IncY(ByVal ny As Long)
' Shift blocks up/down
   For n = 1 To NumBoxes

'         BoxY(1, n) = 1 - BoxY(1, n)
'         BoxY(2, n) = 1 - BoxY(2, n)
'         BoxY(5, n) = 1 - BoxY(5, n)
'         BoxY(6, n) = 1 - BoxY(6, n)
'
'OR
'      If ny < 0 Then
'         BoxY(1, n) = BoxY(4, n) - (BoxY(1, n) - BoxY(4, n))
'         BoxY(2, n) = BoxY(4, n) - (BoxY(2, n) - BoxY(4, n))
'         BoxY(5, n) = BoxY(4, n) - (BoxY(5, n) - BoxY(4, n))
'         BoxY(6, n) = BoxY(4, n) - (BoxY(6, n) - BoxY(4, n))
'      Else
'         BoxY(1, n) = BoxY(4, n) + BoxY(4, n) - BoxY(1, n)
'         BoxY(2, n) = BoxY(4, n) + BoxY(4, n) - BoxY(2, n)
'         BoxY(5, n) = BoxY(4, n) + BoxY(4, n) - BoxY(5, n)
'         BoxY(6, n) = BoxY(4, n) + BoxY(4, n) - BoxY(6, n)
'      End If
' OR
'      For i = 1 To 2
'         If ny > 0 Then
'            BoxY(i, n) = BoxY(i, n) + 2
'            BoxY(i + 4, n) = BoxY(i + 4, n) + 2
'         Else
'            BoxY(i, n) = BoxY(i, n) - 2
'            BoxY(i + 4, n) = BoxY(i + 4, n) - 2
'         End If
'      Next i
' OR
'      For i = 1 To 2
'         If ny > 0 Then
'            BoxY(i, n) = BoxY(i, n) + 2
'            BoxY(i + 4, n) = BoxY(i + 4, n) + 2
'         Else
'            BoxY(i, n) = BoxY(i, n) - 2
'            BoxY(i + 4, n) = BoxY(i + 4, n) - 2
'         End If
'      Next i
' OR
'      Change all height == changing ground plane
'      For i = 1 To 8
'         If ny > 0 Then
'            BoxY(i, n) = BoxY(i, n) + 2
'         Else
'            BoxY(i, n) = BoxY(i, n) - 2
'         End If
'      Next i
'OR
      If BoxY(7, n) > 2 Then  ' Change top heights as well
         BoxY(3, n) = BoxY(3, n) + ny
         BoxY(4, n) = BoxY(4, n) + ny
         BoxY(7, n) = BoxY(7, n) + ny
         BoxY(8, n) = BoxY(8, n) + ny
      End If
      If BoxY(1, n) > 1 Then  ' Change bottom heights
         BoxY(1, n) = BoxY(1, n) + ny
         BoxY(2, n) = BoxY(2, n) + ny
         BoxY(5, n) = BoxY(5, n) + ny
         BoxY(6, n) = BoxY(6, n) + ny
      End If
      
   Next n
End Sub

Public Sub Transform()
'Public zDiv As Single
'Public zNum As Single
   ' PlaneZ = 0
   For n = 1 To NumBoxes
   For i = 1 To 8
      zDiv = CSng(BoxZ(i, n) - eyeZ)
      If zDiv > 0 Then
         zNum = CSng((BoxZ(i, n) - PlaneZ)) / zDiv
         transx(i, n) = CLng((eyeX - BoxX(i, n)) * zNum + 0.5) + BoxX(i, n)
         transy(i, n) = CLng((eyeY - BoxY(i, n)) * zNum + 0.5) + BoxY(i, n)
      End If
   Next i
   Next n
End Sub

'Public Sub TransformX()     ' Test
'   ' PlaneZ = 0
'   For n = 1 To NumBoxes
'   For i = 1 To 4
'      transx(i, n) = BoxX(i, n) + 50
'      transx(i + 4, n) = BoxX(i, n) + 50
'      If i > 2 Then
'      transy(i, n) = BoxY(i, n) + 150
'      transy(i + 4, n) = BoxY(i, n) + 150
'      Else
'      transy(i, n) = BoxY(i, n) + 20
'      transy(i + 4, n) = BoxY(i, n) + 20
'      End If
'   Next i
'   Next n
'End Sub


Public Sub DrawOnBArray()
Dim Cul As Long

ReDim BArray(BArrayWidth, BArrayHeight)
   
   ' Copy gradient background to BArray()
   CopyMemory BArray(1, 1), BackArray(1, 1), BArrayWidth * 512
   
   For n = 1 To NumBoxes
      
      ' Back plane
      ' Pt2 5-6
      
      'If transx(5, n) <> transx(6, n) Then
      'If transy(5, n) <> transy(6, n)  Then
      ' Quicker then drawing same pixel twice
      ' when added to every line drawn? -
      ' No unless all boxes of zero thicknes
      Cul = 3
      BresLine transx(5, n), transy(5, n), transx(6, n), transy(6, n), Cul
      ' Pt2 6-7
      BresLine transx(6, n), transy(6, n), transx(7, n), transy(7, n), Cul
      ' Pt2 7-8
      BresLine transx(7, n), transy(7, n), transx(8, n), transy(8, n), Cul
      ' Pt2 8-5
      BresLine transx(8, n), transy(8, n), transx(5, n), transy(5, n), Cul
      
'  Y
'  | 8-------P7H
'  |/|       /|  Z
'  4--------3 | /
'  | 5------|-6
'  |/       |/
' P1L-------2---- X
      ' Side lines
      ' Pt2 1-5
      Cul = 2
      If BoxY(7, n) - BoxY(1, n) < 2 Then Cul = 3
      BresLine transx(1, n), transy(1, n), transx(5, n), transy(5, n), Cul
      ' Pt2 2-6
      BresLine transx(2, n), transy(2, n), transx(6, n), transy(6, n), Cul
      ' Pt2 3-7
      BresLine transx(3, n), transy(3, n), transx(7, n), transy(7, n), Cul
      ' Pt2 4-8
      BresLine transx(4, n), transy(4, n), transx(8, n), transy(8, n), Cul

      ' Front plane
      ' Pt2 1-2
      Cul = 1
      If BoxY(7, n) - BoxY(1, n) < 2 Then Cul = 3
      BresLine transx(1, n), transy(1, n), transx(2, n), transy(2, n), Cul
      ' Pt2 2-3
      BresLine transx(2, n), transy(2, n), transx(3, n), transy(3, n), Cul
      ' Pt2 3-4
      BresLine transx(3, n), transy(3, n), transx(4, n), transy(4, n), Cul
      ' Pt2 4-1
      BresLine transx(4, n), transy(4, n), transx(1, n), transy(1, n), Cul

   Next n
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
   If ix1 <= BArrayWidth Or ix2 <= BArrayWidth Then
   If iy1 > 0 Or iy2 > 0 Then
   If iy1 <= BArrayHeight Or iy2 <= BArrayHeight Then
      
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
            If ix <= BArrayWidth Then
            If iy > 0 Then
            If iy <= BArrayHeight Then
               BArray(ix, iy) = Cul
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
            If iy <= BArrayWidth Then
            If ix > 0 Then
            If ix <= BArrayHeight Then
               BArray(iy, ix) = Cul
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

Public Sub MakeBackArray()
' BackArrayWidth = BArrayWidth
' BackArrayHeight = 512
ReDim BackArray(BackArrayWidth, BackArrayHeight)
   
   For i = 1 To BackArrayWidth
   For j = 1 To BackArrayHeight
      k = 5 + j \ 3
      If k > 255 Then k = 255
      m = 512 - j
      If m < 1 Then m = 1
      BackArray(i, m) = k
   Next j
   Next i
End Sub

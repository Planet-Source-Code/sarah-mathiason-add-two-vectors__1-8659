VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddVector 
   Caption         =   "Adding Vectors"
   ClientHeight    =   11565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16035
   LinkTopic       =   "Form1"
   ScaleHeight     =   11565
   ScaleWidth      =   16035
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtAngle1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   8040
      TabIndex        =   45
      Text            =   "1"
      Top             =   10560
      Width           =   615
   End
   Begin VB.TextBox txtAngle1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   8040
      TabIndex        =   44
      Text            =   "0"
      Top             =   9720
      Width           =   615
   End
   Begin VB.TextBox txtAngle1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   5040
      TabIndex        =   43
      Text            =   "1"
      Top             =   10560
      Width           =   615
   End
   Begin VB.CommandButton cmdMagMinus 
      DownPicture     =   "frmAddVector.frx":0000
      Height          =   495
      Left            =   9120
      Picture         =   "frmAddVector.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   10800
      Width           =   495
   End
   Begin VB.CommandButton cmdMagPlus 
      DownPicture     =   "frmAddVector.frx":0614
      Height          =   495
      Left            =   9120
      Picture         =   "frmAddVector.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   10200
      Width           =   495
   End
   Begin VB.TextBox txtAngle1 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   5070
      TabIndex        =   22
      Text            =   "0"
      Top             =   9720
      Width           =   615
   End
   Begin VB.Frame Frame3 
      Height          =   2535
      Left            =   9840
      TabIndex        =   20
      Top             =   9120
      Width           =   45
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   8910
      TabIndex        =   19
      Top             =   9120
      Width           =   45
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   5910
      TabIndex        =   18
      Top             =   9120
      Width           =   45
   End
   Begin MSComctlLib.Slider sldOX 
      Height          =   495
      Left            =   600
      TabIndex        =   17
      Top             =   8640
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   873
      _Version        =   393216
      Max             =   15135
      TickStyle       =   1
      TickFrequency   =   100
   End
   Begin MSComctlLib.Slider sldOY 
      Height          =   8655
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   15266
      _Version        =   393216
      Orientation     =   1
      Max             =   8415
      TickFrequency   =   100
   End
   Begin MSComctlLib.Slider sldAngle1 
      Height          =   495
      Left            =   3150
      TabIndex        =   12
      Top             =   9600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      Max             =   360
      TickStyle       =   2
      TickFrequency   =   36
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C8C8C8&
      Height          =   8415
      Left            =   720
      ScaleHeight     =   8355
      ScaleWidth      =   15075
      TabIndex        =   0
      Top             =   120
      Width           =   15135
      Begin VB.Label lblInterior3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Width           =   75
      End
      Begin VB.Label lblInterior2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   75
      End
      Begin VB.Label lblInterior1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   40
         Top             =   6720
         Width           =   75
      End
      Begin VB.Label lblResultant 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C80000&
         Height          =   240
         Left            =   12840
         TabIndex        =   36
         Top             =   5640
         Width           =   75
      End
      Begin VB.Label lblVector2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C8&
         Height          =   240
         Left            =   11160
         TabIndex        =   35
         Top             =   6120
         Width           =   75
      End
      Begin VB.Label lblVector1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000AF00&
         Height          =   240
         Left            =   9840
         TabIndex        =   34
         Top             =   7200
         Width           =   75
      End
   End
   Begin MSComctlLib.Slider sldForce1 
      Height          =   495
      Left            =   3150
      TabIndex        =   13
      Top             =   10440
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   1000
      Min             =   1
      Max             =   7500
      SelStart        =   1
      TickStyle       =   1
      TickFrequency   =   500
      Value           =   1
   End
   Begin MSComctlLib.Slider sldAngle2 
      Height          =   495
      Left            =   6150
      TabIndex        =   14
      Top             =   9600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      Max             =   360
      TickStyle       =   2
      TickFrequency   =   36
   End
   Begin MSComctlLib.Slider sldForce2 
      Height          =   495
      Left            =   6150
      TabIndex        =   15
      Top             =   10440
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   1000
      Min             =   1
      Max             =   7500
      SelStart        =   1
      TickStyle       =   1
      TickFrequency   =   500
      Value           =   1
   End
   Begin VB.Label lblMag 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "x 1"
      Height          =   255
      Left            =   9000
      TabIndex        =   21
      Top             =   9600
      Width           =   735
   End
   Begin VB.Label Label16 
      Caption         =   "Magnitude"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11070
      TabIndex        =   39
      Top             =   10200
      Width           =   2055
   End
   Begin VB.Label lblForceCalc 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10230
      TabIndex        =   38
      Top             =   10200
      Width           =   735
   End
   Begin VB.Label lblAngleBetween 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10230
      TabIndex        =   37
      Top             =   9720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Mag"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   31
      Top             =   9240
      Width           =   495
   End
   Begin VB.Label lblEndV2 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   7350
      TabIndex        =   30
      Top             =   11520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblOrV2 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   6030
      TabIndex        =   29
      Top             =   11520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblEndV1 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   4230
      TabIndex        =   28
      Top             =   11520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "End Coords"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7470
      TabIndex        =   27
      Top             =   11280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "Origin Coords"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6105
      TabIndex        =   26
      Top             =   11280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "End Coords"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4590
      TabIndex        =   25
      Top             =   11280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblOrV1 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   2910
      TabIndex        =   24
      Top             =   11520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblOrigin 
      Alignment       =   2  'Center
      Caption         =   "Origin Coords"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3030
      TabIndex        =   23
      Top             =   11280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Resultant Vector"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C80000&
      Height          =   375
      Left            =   10590
      TabIndex        =   11
      Top             =   9240
      Width           =   2295
   End
   Begin VB.Label lblResultAngle 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10230
      TabIndex        =   10
      Top             =   9720
      Width           =   735
   End
   Begin VB.Label lblResultForce 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10230
      TabIndex        =   9
      Top             =   10200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label9"
      Height          =   300
      Left            =   -1440
      TabIndex        =   8
      Top             =   -1920
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Angle of Resultant"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11070
      TabIndex        =   7
      Top             =   9720
      Width           =   2415
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Second Vector"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C8&
      Height          =   375
      Left            =   6270
      TabIndex        =   6
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "First Vector"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000AF00&
      Height          =   375
      Left            =   3270
      TabIndex        =   5
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Magnitude"
      Height          =   255
      Left            =   6990
      TabIndex        =   4
      Top             =   10920
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Angle "
      Height          =   255
      Left            =   7110
      TabIndex        =   3
      Top             =   10080
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Magnitude"
      Height          =   255
      Left            =   4110
      TabIndex        =   2
      Top             =   10920
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Angle "
      Height          =   255
      Left            =   4110
      TabIndex        =   1
      Top             =   10080
      Width           =   1095
   End
End
Attribute VB_Name = "frmAddVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'if you get errors when loading the form it is because the
'computer doesn't know where the icons for the magnification
'are.  Simply finish loading the file then re-link the
'command buttons to the icons


Dim Angle 'stores the angle for the first vector
Dim XX 'stores the ending X coord for the first vector and origin X coord for second vector
Dim YY 'stores the ending Y coord for the first vecotr and origin Y coord for second vector
Dim OX 'stores the origin X coord for the first vector and resultant vector
Dim OY 'stores the origin Y coord for the first vector and resultant vector
Dim XX2 'stores the ending X coord for the second vector and resultant vector
Dim YY2 'stores the ending Y coord for the second vector and resultant vector
Dim Angle2 'stores the angle of the second vector
Dim Rad 'will be used to convert the degree angles to radians
Dim Length 'this is the magnitude of the first vector
Dim Length2 'magnitude of the second vector
Dim Mag 'This is the Magnification factor of the display

Private Sub cmdMagMinus_Click()
'Magnifying down you have more control to better fit the diagram in the screen
Mag = Mag - 0.5
If Mag < 1 Then Mag = 1
lblMag.Caption = "x " & Mag
sldAngle1_Change
End Sub

Private Sub cmdMagPlus_Click()
Mag = Mag + 1
If Mag > 15 Then Mag = 15
lblMag.Caption = "x " & Mag
sldAngle1_Change
End Sub

Private Sub Form_Load()
'txtAngle1(0).text holds Angle for first Vector
'txtAngle1(1).text holds Magnitude for first Vector
'txtAngle1(2).text holds Angle for second Vector
'txtAngle1(3).text holds Magnitude for second Vector
Rad = 180 / 3.14159265358979
sldOX.Value = Picture1.Width / 2
sldOY.Value = Picture1.Height / 2
Mag = 1
End Sub

Private Sub sldAngle1_Change()

Dim C 'holds the magnitude of resultant vector
Dim B 'holds the angle of the resultant vector (in degrees)
Dim A 'holds interior angle of triangle with sides made by both first and second vectors
Dim Z 'holds interior angle of triangle with sides made by the resultant and one of the other vectors
Dim Q 'holds interior angle of triangle with sides made by the resultant and one of the other vectors
Dim X 'temp variable

'erase all old data
'==================
Picture1.Cls

'find out new data
'=================
OX = sldOX.Value ' when you change the sliders next to screen it changes your origin
OY = sldOY.Value ' so you can move the diagram to fit in the screen if it moves off.

txtAngle1(0).Text = sldAngle1.Value ' the text boxes text must be same as slider value
txtAngle1(1).Text = sldForce1.Value
txtAngle1(2).Text = sldAngle2.Value
txtAngle1(3).Text = sldForce2.Value

Angle = Val(sldAngle1.Value) / Rad  ' information can be updated now that old data has been
Angle2 = Val(sldAngle2.Value) / Rad ' erased from screen
Length = Val(sldForce1.Value)
Length2 = Val(sldForce2.Value)

'Calculate new verteces (coordinates) based on new angles and magnitudes
'=======================================================================
YY = (Sin(Angle) * Length * Mag) + OY
XX = (Sin((90 / Rad - Angle)) * Length * Mag) + OX
YY2 = (Sin(Angle2) * Length2 * Mag) + YY
XX2 = (Sin((90 / Rad - Angle2)) * Length2 * Mag) + XX

'draw in some axes for frames of reference at the new verteces
'=============================================================
Picture1.Line (OX, OY - 300 * Mag)-(OX, OY + 300 * Mag)
Picture1.Line (OX - 300 * Mag, OY)-(OX + 300 * Mag, OY)
Picture1.Line (XX, YY - 300 * Mag)-(XX, YY + 300 * Mag)
Picture1.Line (XX - 300 * Mag, YY)-(XX + 300 * Mag, YY)
Picture1.Line (XX2, YY2 - 300 * Mag)-(XX2, YY2 + 300 * Mag), RGB(0, 0, 200)
Picture1.Line (XX2 - 300 * Mag, YY2)-(XX2 + 300 * Mag, YY2), RGB(0, 0, 200)

'draw in new vector lines plus circles as their terminators
'===========================================================
Picture1.Line (OX, OY)-(XX, YY), RGB(0, 175, 0)
Picture1.Circle (XX, YY), (40 * Mag), RGB(0, 175, 0)
Picture1.Line (XX, YY)-(XX2, YY2), RGB(200, 0, 0)
Picture1.Circle (XX2, YY2), (40 * Mag), RGB(200, 0, 0)
Picture1.Line (OX, OY)-(XX2, YY2), RGB(0, 0, 200)

'calculate Resultant Magnitude using known verteces and the Distance Formula (not used)
'======================================================================================
'   distance formula:  distance = SQR((x2-x1)^2 + (y2-y1)^2)
lblResultForce.Caption = Int(Sqr(((OX - XX2) / Mag) ^ 2 + ((OY - YY2) / Mag) ^ 2))

'record all coordinates in labels (not visible on form but are present)
'======================================================================
lblOrV1.Caption = "(" & Int(OX) & ")" & " , " & "(" & Int(OY) & ")"
lblEndV1.Caption = "(" & Int(XX) & ")" & " , " & "(" & Int(YY) & ")"
lblOrV2.Caption = "(" & Int(XX) & ")" & " , " & "(" & Int(YY) & ")"
lblEndV2.Caption = "(" & Int(XX2) & ")" & " , " & "(" & Int(YY2) & ")"

'find angle between the two vectors being added using a formula I made up
'========================================================================
'   (notice all angles are converted to degrees)
A = Abs(180 - 180 - (270 - (Angle2 * Rad)) + (90 - (Angle * Rad)))
If A > 180 Then A = Abs((180 - A) + 180)
lblAngleBetween.Caption = Int(A) & "°"

'Calculate the Magnitude of Resultant using the Cosine Law
'=========================================================
'   Cosine Law: c^2 = a^2 + b^2 - 2*a*b*Cos(C)
C = Int(Sqr(Length ^ 2 + Length2 ^ 2 - 2 * Length * Length2 * Cos(A / Rad)))
If C < 1 Then C = 1
lblForceCalc.Caption = C

'Calculate Angle of Vector using Sin Law and VB function for Inverse Sine and conditions I made up.
'==================================================================================================
'   Sine Law:   Sin(A)/a = Sin(B)/b
'   function for inverse sine : Sin-1(X) = Atn(X / Sqr(-X * X + 1))
'   the conditions make the result match with angles measured from horizontal
'       axis and measured counter-clockwise.
X = (((YY2 - OY) / Mag) / Val(lblForceCalc.Caption))
If X = 1 Then X = 2
If X = -1 Then X = 2
B = Abs(Int(Atn(X / Sqr(Abs(-X * X + 1))) * Rad))
If XX2 < OX And YY2 > OY Then B = 180 - B
If XX2 < OX And YY2 < OY Then B = B + 180
If XX2 > OX And YY2 < OY Then B = 360 - B

'now that we have all the angles and magnitudes for all three vectors we can
'display them on the screen to better convey information
'===========================================================================
'find points halfway along vectors on which to display vector information labels
'first label point
lYY = (Sin(Angle) * Length / 2 * Mag) + OY
lxX = (Sin((90 / Rad - Angle)) * Length / 2 * Mag) + OX
'second label point
lYY2 = (Sin(Angle2) * Length2 / 2 * Mag) + YY
lXX2 = (Sin((90 / Rad - Angle2)) * Length2 / 2 * Mag) + XX
'resultant label point
rYY2 = (Sin(B / Rad) * C / 2 * Mag) + OY
rXX2 = (Sin((90 / Rad - B / Rad)) * C / 2 * Mag) + OX
lblResultAngle.Caption = B & "°"
'position vector information labels
lblVector1.Top = lYY
lblVector1.Left = lxX
lblVector2.Top = lYY2
lblVector2.Left = lXX2
lblResultant.Top = rYY2
lblResultant.Left = rXX2
'update vector position label information
lblVector1.Caption = Int(Length) & " units, @ " & Int(Angle * Rad) & "°"
lblVector2.Caption = Int(Length2) & " units, @ " & Int(Angle2 * Rad) & "°"
lblResultant.Caption = Int(C) & " units, @ " & Int(B) & "°"
'alt+01456 gives °

'draw some arcs on axes to show where angles are measured from
'=============================================================
Picture1.Circle (OX, OY), (75 * Mag), RGB(0, 175, 0), 360 / Rad - Angle, 0
Picture1.Circle (XX, YY), (75 * Mag), RGB(200, 0, 0), 360 / Rad - Angle2, 0
Picture1.Circle (OX, OY), (100 * Mag), RGB(0, 0, 200), 360 / Rad - B / Rad, 0

'calculate the other interior angles using same formula that I made up
'=====================================================================
Z = Abs(180 - 180 - (270 - (Angle2 * Rad)) + (90 - (B)))
If Z > 180 Then Z = Abs((180 - Z) + 180)
Q = Abs(180 - 180 - (270 - (Angle * Rad)) + (90 - (B)))
If Q > 180 Then Q = Abs((180 - Q) + 180)
'update the labels
lblInterior1.Caption = 180 - Int(Z) & "°"
lblInterior2.Caption = 180 - Int(Q) & "°"
lblInterior3.Caption = Int(A) & "°"
'display these interior angle labels at verteces of corresponding angles
lblInterior3.Top = YY
lblInterior3.Left = XX - lblInterior3.Width
lblInterior1.Top = YY2
lblInterior1.Left = XX2 - lblInterior1.Width
lblInterior2.Top = OY
lblInterior2.Left = OX - lblInterior2.Width

End Sub
'********************************************************************
'*for everything below:  Any change to any setting for either vector*
'*will result in an update in the screen to reflect these changes,  *
'*so a call to sldAngle1_change is made to do this.                 *
'********************************************************************
'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
Private Sub sldAngle1_Scroll()
sldAngle1_Change
End Sub

Private Sub sldAngle2_Click()
sldAngle1_Change
End Sub

Private Sub sldAngle2_Scroll()
sldAngle1_Change
End Sub

Private Sub sldForce1_Click()
sldAngle2_Click
End Sub

Private Sub sldForce1_Scroll()
sldAngle2_Click
End Sub

Private Sub sldForce2_Click()
sldAngle1_Change
End Sub

Private Sub sldForce2_Scroll()
sldAngle1_Change
End Sub

Private Sub sldOX_Click()
sldAngle1_Change
End Sub

Private Sub sldOX_Scroll()
sldAngle1_Change
End Sub

Private Sub sldOY_Click()
sldAngle1_Change
End Sub

Private Sub sldOY_Scroll()
sldAngle1_Change
End Sub
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'********************************************************************
'*for everything above:  Any change to any setting for either vector*
'*will result in an update in the screen to reflect these changes,  *
'*so a call to sldAngle1_change is made to do this.                 *
'********************************************************************

Private Sub txtAngle1_Change(Index As Integer)
'If the text is changed the slider bar will change too
'then a screen supdate is performed by calling sldAngle1_change
If txtAngle1(Index).Text <> "" Then
sldAngle1.Value = Val(txtAngle1(0).Text)
sldForce1.Value = Val(txtAngle1(1).Text)
sldAngle2.Value = Val(txtAngle1(2).Text)
sldForce2.Value = Val(txtAngle1(3).Text)
sldAngle1_Change
End If
End Sub

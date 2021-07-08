VERSION 4.00
Begin VB.Form Joystick 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Talos : Joystick - System Monitor (Press F1 for Help)"
   ClientHeight    =   6330
   ClientLeft      =   2340
   ClientTop       =   1560
   ClientWidth     =   7665
   BeginProperty Font 
      name            =   "MS Sans Serif"
      charset         =   1
      weight          =   700
      size            =   8.25
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   Height          =   6735
   Icon            =   "JOYSTICK.frx":0000
   Left            =   2280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "JOYSTICK.frx":0442
   ScaleHeight     =   6330
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   Top             =   1215
   Width           =   7785
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6765
      TabIndex        =   85
      Text            =   "10%"
      Top             =   2985
      Width           =   765
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   3300
      Top             =   1575
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Write on File"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4350
      TabIndex        =   75
      Top             =   4170
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.PictureBox term 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   510
      Index           =   3
      Left            =   3540
      MousePointer    =   10  'Up Arrow
      Picture         =   "JOYSTICK.frx":1601F4
      ScaleHeight     =   480
      ScaleWidth      =   495
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4095
      Width           =   525
   End
   Begin VB.PictureBox term 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   510
      Index           =   5
      Left            =   255
      MousePointer    =   10  'Up Arrow
      Picture         =   "JOYSTICK.frx":160636
      ScaleHeight     =   480
      ScaleWidth      =   525
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2715
      Width           =   555
   End
   Begin VB.PictureBox term 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   510
      Index           =   2
      Left            =   1440
      MousePointer    =   10  'Up Arrow
      Picture         =   "JOYSTICK.frx":160A78
      ScaleHeight     =   480
      ScaleWidth      =   495
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4110
      Width           =   525
   End
   Begin VB.PictureBox term 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   510
      Index           =   4
      Left            =   270
      MousePointer    =   10  'Up Arrow
      Picture         =   "JOYSTICK.frx":160EBA
      ScaleHeight     =   480
      ScaleWidth      =   510
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5505
      Width           =   540
   End
   Begin VB.PictureBox term 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   510
      Index           =   0
      Left            =   2460
      MousePointer    =   10  'Up Arrow
      Picture         =   "JOYSTICK.frx":1612FC
      ScaleHeight     =   480
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   2970
      Width           =   525
   End
   Begin VB.PictureBox term 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   510
      Index           =   1
      Left            =   2475
      MousePointer    =   10  'Up Arrow
      Picture         =   "JOYSTICK.frx":16173E
      ScaleHeight     =   480
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   5250
      Width           =   525
   End
   Begin VB.PictureBox term 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   465
      Index           =   6
      Left            =   1050
      MousePointer    =   10  'Up Arrow
      Picture         =   "JOYSTICK.frx":161B80
      ScaleHeight     =   435
      ScaleWidth      =   450
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3300
      Width           =   480
   End
   Begin VB.PictureBox term 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   465
      Index           =   7
      Left            =   1050
      MousePointer    =   10  'Up Arrow
      Picture         =   "JOYSTICK.frx":161FC2
      ScaleHeight     =   435
      ScaleWidth      =   450
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4905
      Width           =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   3285
      Top             =   1035
   End
   Begin VB.Label Label51 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Down"
      ForeColor       =   &H00FFFFC0&
      Height          =   270
      Left            =   855
      TabIndex        =   89
      Top             =   5820
      Width           =   465
   End
   Begin VB.Label Label43 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ζ"
      ForeColor       =   &H00FFFFC0&
      Height          =   270
      Left            =   300
      TabIndex        =   88
      Top             =   4275
      Width           =   270
   End
   Begin VB.Image Image2 
      Height          =   405
      Left            =   90
      Picture         =   "JOYSTICK.frx":162404
      Stretch         =   -1  'True
      Top             =   -15
      Width           =   450
   End
   Begin VB.Label Label50 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Scale"
      ForeColor       =   &H00FFC0C0&
      Height          =   240
      Left            =   3780
      TabIndex        =   87
      Top             =   495
      Width           =   1140
   End
   Begin VB.Label Scale_dis 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   3780
      TabIndex        =   86
      Top             =   720
      Width           =   1140
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   90
      Picture         =   "JOYSTICK.frx":162846
      Stretch         =   -1  'True
      Top             =   0
      Width           =   450
   End
   Begin VB.Label Label49 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      ForeColor       =   &H00FFFFC0&
      Height          =   270
      Left            =   1005
      TabIndex        =   84
      Top             =   4290
      Width           =   255
   End
   Begin VB.Label Label48 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Up"
      ForeColor       =   &H00FFFFC0&
      Height          =   270
      Left            =   840
      TabIndex        =   83
      Top             =   2715
      Width           =   270
   End
   Begin VB.Label Label47 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Υ"
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   3210
      TabIndex        =   82
      Top             =   4125
      Width           =   255
   End
   Begin VB.Label Label46 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Χ"
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Left            =   2715
      TabIndex        =   81
      Top             =   3615
      Width           =   255
   End
   Begin VB.Label zyg_err 
      Height          =   210
      Left            =   3300
      TabIndex        =   80
      Top             =   2175
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label Label44 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Speed"
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   6765
      TabIndex        =   79
      Top             =   2685
      Width           =   780
   End
   Begin VB.Label Label40 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "System Problems "
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   450
      TabIndex        =   78
      Top             =   45
      Width           =   915
   End
   Begin VB.Label clue 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   1350
      TabIndex        =   77
      Top             =   45
      Width           =   6180
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   2055
      Left            =   4305
      Top             =   4095
      Width           =   2760
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFC0&
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   1380
      Left            =   4305
      Top             =   2610
      Width           =   2415
   End
   Begin VB.Label Label39 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Scale Timer Enabled"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   4590
      TabIndex        =   76
      Top             =   4155
      Width           =   2385
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Operating System"
      ForeColor       =   &H00FFC0C0&
      Height          =   240
      Left            =   90
      TabIndex        =   74
      Top             =   495
      Width           =   1905
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   90
      TabIndex        =   73
      Top             =   720
      Width           =   1905
   End
   Begin VB.Label zyg_Settings 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5970
      TabIndex        =   72
      Top             =   4380
      Width           =   1005
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Scale Settings "
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   4350
      TabIndex        =   71
      Top             =   4380
      Width           =   1605
   End
   Begin VB.Label Command2 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Scale Testing"
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   4365
      TabIndex        =   70
      Top             =   3630
      Width           =   2265
   End
   Begin VB.Label Change_Digits 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Change Scale Decimals"
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   4350
      TabIndex        =   69
      Top             =   5775
      Width           =   2625
   End
   Begin VB.Label Rezero 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Scale Rezero"
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   4350
      TabIndex        =   68
      Top             =   5460
      Width           =   2625
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3765
      Top             =   2070
      _version        =   65536
      _extentx        =   847
      _extenty        =   847
      _stockprops     =   0
      cdtimeout       =   0
      commport        =   1
      ctstimeout      =   0
      dsrtimeout      =   0
      dtrenable       =   -1  'True
      handshaking     =   0
      inbuffersize    =   1024
      inputlen        =   0
      interval        =   1000
      nulldiscard     =   0   'False
      outbuffersize   =   512
      parityreplace   =   "?"
      rthreshold      =   0
      rtsenable       =   0   'False
      settings        =   "9600,n,8,1"
      sthreshold      =   0
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   7245
      TabIndex        =   67
      Top             =   2250
      Width           =   315
   End
   Begin VB.Label Label41 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Alarm 2"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   5670
      TabIndex        =   66
      Top             =   2295
      Width           =   1500
   End
   Begin VB.Label Pause 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   2850
      TabIndex        =   65
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Door 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   2025
      TabIndex        =   64
      Top             =   720
      Width           =   800
   End
   Begin VB.Label Cut_Off 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   6225
      TabIndex        =   63
      Top             =   720
      Width           =   1305
   End
   Begin VB.Label Crush 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   4950
      TabIndex        =   62
      Top             =   720
      Width           =   1245
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pause"
      ForeColor       =   &H00FFC0C0&
      Height          =   240
      Left            =   2850
      TabIndex        =   61
      Top             =   495
      Width           =   900
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Door"
      ForeColor       =   &H00FFC0C0&
      Height          =   240
      Left            =   2025
      TabIndex        =   60
      Top             =   495
      Width           =   800
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Crash"
      ForeColor       =   &H00FFC0C0&
      Height          =   240
      Left            =   6225
      TabIndex        =   59
      Top             =   495
      Width           =   1305
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Emergency"
      ForeColor       =   &H00FFC0C0&
      Height          =   240
      Left            =   4950
      TabIndex        =   58
      Top             =   495
      Width           =   1245
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Central Water"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   5670
      TabIndex        =   57
      Top             =   1965
      Width           =   1500
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   7245
      TabIndex        =   56
      Top             =   1920
      Width           =   315
   End
   Begin VB.Label Grip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Gripper"
      ForeColor       =   &H000000FF&
      Height          =   510
      Left            =   2310
      MousePointer    =   10  'Up Arrow
      TabIndex        =   8
      Top             =   4095
      Width           =   825
   End
   Begin VB.Label zyg_cont 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Scale Output Port "
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   4350
      TabIndex        =   55
      Top             =   4650
      Width           =   1605
   End
   Begin VB.Label zyg_port 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5970
      TabIndex        =   54
      Top             =   4650
      Width           =   1005
   End
   Begin VB.Label zyg_dec 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   5970
      TabIndex        =   53
      Top             =   4920
      Width           =   1005
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Scale Decimals "
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   4350
      TabIndex        =   52
      Top             =   4920
      Width           =   1605
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Scale Display "
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   4350
      TabIndex        =   51
      Top             =   5190
      Width           =   1605
   End
   Begin VB.Label zyg_display 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5970
      TabIndex        =   50
      Top             =   5190
      Width           =   1005
   End
   Begin VB.Label flag_zyg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5820
      TabIndex        =   49
      Top             =   2685
      Width           =   825
   End
   Begin VB.Label cur_tim 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5820
      TabIndex        =   48
      Top             =   3315
      Width           =   825
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5820
      TabIndex        =   47
      Top             =   3000
      Width           =   825
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Waiting Cycles"
      ForeColor       =   &H00400000&
      Height          =   315
      Left            =   4380
      TabIndex        =   46
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Time  (msec)"
      ForeColor       =   &H00400000&
      Height          =   315
      Left            =   4380
      TabIndex        =   45
      Top             =   3315
      Width           =   1455
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Measurements"
      ForeColor       =   &H00400000&
      Height          =   315
      Left            =   4380
      TabIndex        =   44
      Top             =   2685
      Width           =   1455
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Powder Motor"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3735
      TabIndex        =   43
      Top             =   1050
      Width           =   1455
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Cold Water"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   5670
      TabIndex        =   42
      Top             =   1020
      Width           =   1500
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Mixer"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3705
      TabIndex        =   41
      Top             =   1950
      Width           =   1455
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Hot Water"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   5670
      TabIndex        =   40
      Top             =   1335
      Width           =   1500
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Alarm 1"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3750
      TabIndex        =   39
      Top             =   2295
      Width           =   1455
   End
   Begin VB.Label Label32 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Right Valve"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3750
      TabIndex        =   38
      Top             =   1350
      Width           =   1515
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Left Valve"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3735
      TabIndex        =   37
      Top             =   1650
      Width           =   1455
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Mixed Water"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   5670
      TabIndex        =   36
      Top             =   1650
      Width           =   1500
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   7245
      TabIndex        =   35
      Top             =   1605
      Width           =   315
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   5265
      TabIndex        =   34
      Top             =   1605
      Width           =   315
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   5265
      TabIndex        =   33
      Top             =   1290
      Width           =   315
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   5265
      TabIndex        =   32
      Top             =   2250
      Width           =   315
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   7245
      TabIndex        =   31
      Top             =   1290
      Width           =   315
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   5265
      TabIndex        =   30
      Top             =   1920
      Width           =   315
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   7245
      TabIndex        =   29
      Top             =   975
      Width           =   315
   End
   Begin VB.Label Label38 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   5265
      TabIndex        =   28
      Top             =   975
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Axis 1"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   135
      TabIndex        =   27
      Top             =   1485
      Width           =   645
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Axis 2"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   135
      TabIndex        =   26
      Top             =   1695
      Width           =   645
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Axis 4"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   135
      TabIndex        =   25
      Top             =   2160
      Width           =   645
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Axis 3"
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   135
      TabIndex        =   24
      Top             =   1935
      Width           =   645
   End
   Begin VB.Label position 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   0
      Left            =   810
      TabIndex        =   23
      Top             =   1485
      Width           =   795
   End
   Begin VB.Label position 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   810
      TabIndex        =   22
      Top             =   1710
      Width           =   795
   End
   Begin VB.Label position 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   3
      Left            =   810
      TabIndex        =   21
      Top             =   2160
      Width           =   795
   End
   Begin VB.Label position 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   2
      Left            =   810
      TabIndex        =   20
      Top             =   1935
      Width           =   795
   End
   Begin VB.Label LBL_Encoder 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   2
      Left            =   1620
      TabIndex        =   19
      Top             =   1935
      Width           =   795
   End
   Begin VB.Label LBL_Encoder 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   3
      Left            =   1620
      TabIndex        =   18
      Top             =   2160
      Width           =   795
   End
   Begin VB.Label LBL_Encoder 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   1620
      TabIndex        =   17
      Top             =   1710
      Width           =   795
   End
   Begin VB.Label LBL_Encoder 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   0
      Left            =   1620
      TabIndex        =   16
      Top             =   1485
      Width           =   795
   End
   Begin VB.Label LBL_Velocity 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   0
      Left            =   2430
      TabIndex        =   15
      Top             =   1485
      Width           =   795
   End
   Begin VB.Label LBL_Velocity 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   2430
      TabIndex        =   14
      Top             =   1710
      Width           =   795
   End
   Begin VB.Label LBL_Velocity 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   3
      Left            =   2430
      TabIndex        =   13
      Top             =   2160
      Width           =   795
   End
   Begin VB.Label LBL_Velocity 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   2
      Left            =   2430
      TabIndex        =   12
      Top             =   1935
      Width           =   795
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Current Velocity"
      ForeColor       =   &H00FFFF00&
      Height          =   465
      Left            =   2430
      TabIndex        =   11
      Top             =   1035
      Width           =   795
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Encoder Display"
      ForeColor       =   &H00FFFF00&
      Height          =   465
      Left            =   1620
      TabIndex        =   10
      Top             =   1035
      Width           =   795
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Axis Position"
      ForeColor       =   &H00FFFF00&
      Height          =   465
      Left            =   810
      TabIndex        =   9
      Top             =   1035
      Width           =   795
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   3540
      Left            =   180
      Top             =   2610
      Width           =   4065
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFF00&
      BorderStyle     =   3  'Dot
      X1              =   2715
      X2              =   2730
      Y1              =   3480
      Y2              =   5265
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFF00&
      BorderStyle     =   3  'Dot
      X1              =   1980
      X2              =   3495
      Y1              =   4350
      Y2              =   4335
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFF00&
      BorderStyle     =   3  'Dot
      X1              =   510
      X2              =   510
      Y1              =   3300
      Y2              =   5430
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFF00&
      BorderStyle     =   3  'Dot
      X1              =   1275
      X2              =   1275
      Y1              =   4905
      Y2              =   3795
   End
End
Attribute VB_Name = "Joystick"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
 Dim clue1, clue2, clue3, clue4, clue5, clue6
Function zygisi4(X As Variant)
On Error GoTo er_ex
   Dim counter
   counter = 0: tot_counter = 0: Label10 = 0
' If system_ready = 0 Then zygisi4 = 0: Exit Function
     
    Do
     MSComm1.InBufferCount = 0
      FromModem$ = ""
   MSComm1.Output = Balance_Asking + Chr$(13)
       MilSec (40)
      dummy = DoEvents()
       If MSComm1.InBufferCount Then
           buf = MSComm1.InBufferCount
          FromModem$ = FromModem$ + MSComm1.Input
         ' asq = FromModem$
          ' asw = MSComm1.Input
          If InStr(FromModem$, Chr$(13)) Then
                If Mid$(FromModem$, 5, 1) = "." Or Mid$(FromModem$, 6, 1) = "." Or Mid$(FromModem$, 7, 1) = "." Then
                    Exit Do
                End If
          End If
      Else
       counter = counter + 1
       If counter >= 15 Then tot_counter = tot_counter + 1
            Label10 = counter
          If tot_counter >= 10 Then
          counter = 0: tot_counter = tot_counter + 1
               zygisi4 = " "
               Exit Function
           End If
          If counter >= 50 Then  ' 500
             counter = 0: tot_counter = tot_counter + 1
               MSComm1.InBufferCount = 0
             MSComm1.Output = Balance_Asking + Chr$(13)
              End If
     End If
    Loop
         zygisi4 = Left$(FromModem$, 9)
er_ex:
Exit Function
er_det:
Resume er_ex
    End Function

Function zyg_decimals()
asw = zygisi4(0)
    poin = InStr(asw, ".")
    For I = 1 To 20
       If IsNumeric(Mid(asw, poin + I, 1)) Then dec = dec + 1 Else Exit For
     Next I
 zyg_decimals = dec
 
Exit Function
er_zyg:
Err_Ck (Err)
Resume Next
End Function

Public Function CheckForNull(ByVal arg) As Integer
On Error GoTo er_chek_val
CheckForNull = 1
 If IsNull(arg) Then CheckForNull = 0
 If IsNumeric(arg) Then
     If CDbl(arg * 100000000) = 0 Then CheckForNull = 0
End If
If arg = "" Then CheckForNull = 0
 Exit Function
er_chek_val: Stop
End Function

Private Sub anix_zygaria()
On Error GoTo err_fnd
  MSComm1.CommPort = Balance_Port
    MSComm1.Settings = Balance_Settings
    If MSComm1.PortOpen = False Then MSComm1.PortOpen = True
    MSComm1.Settings = Balance_Settings
    MSComm1.InputLen = 0
   
  Me.zyg_port.BackColor = Green
    Me.zyg_port.Caption = Balance_Port
Exit Sub
err_fnd:
If Err = 8013 Or Err = 8010 Then
    zyg_err = " "
Else
MsgBox Err.Description, , "Talos "
End If
Resume Next
End Sub

Private Sub Change_Digits_Click()
On Error GoTo er_res
Screen.MousePointer = 11
If zyg_dec = "3" Then
   MSComm1.Output = Balance_Clearing + Chr$(13)
   MilSec 2000
   MSComm1.Output = Balance_2_Digits + Chr$(13)
ElseIf zyg_dec = "2" Then
   MSComm1.Output = Balance_Clearing + Chr$(13)
   MilSec 2000
   MSComm1.Output = Balance_3_Digits + Chr$(13)
Else
  zyg_dec = "-"
  zyg_dec.BackColor = Orange
  MSComm1.Output = "-"
End If
zyg_dec.Refresh
Screen.MousePointer = 1
Exit Sub
er_res:
Err_Ck (Err)
Resume Next
End Sub

Private Sub Change_Digits_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Change_Digits.BackColor = Orange '&H00FFFFFF&
End Sub

Private Sub Change_Digits_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Change_Digits.BackColor = Grey
End Sub



Private Sub check1_Click()
If check1 = 0 Then
    Timer2.Enabled = False
    zyg_dec.Enabled = False
    zyg_display.Enabled = False
     Rezero.Enabled = False
     Change_Digits.Enabled = False
Else
    Timer2.Enabled = True
    zyg_dec.Enabled = True
    zyg_display.Enabled = True
     Rezero.Enabled = True
     Change_Digits.Enabled = True
End If
End Sub

Private Sub Combo1_Click()
     Speed = Val(Combo1) / 100

If CDbl(Speed) = 0 Or CDbl(Speed) > 1.3 Then
   Speed = "0.1"
    MsgBox "Select a number between 0 and 120%   like 50% and click it with the mouse ...", , "Talos"
    Combo1.SetFocus
    asd = 1
End If
SendVeloc (CDbl(Speed))
asw = CDbl(Speed) * 100 & "%"
Speed = asw
If asd = 1 Then Combo1.ForeColor = Red Else Combo1.ForeColor = 0

End Sub

Private Sub Command2_Click()
Dim db As Database, r As Recordset
On Error GoTo er_chek
Screen.MousePointer = 11
Command2.Enabled = False
Label13.Enabled = True
Label11.Enabled = True
Label12.Enabled = True
flag_zyg.Enabled = True: flag_zyg = 0
Label10.Enabled = True
cur_tim.Enabled = True
start_tim = GetCurrentTime()
cur_tim = 0
For I = 1 To 30
    Start = GetCurrentTime()
    Me.zyg_display = zygisi4(0)
    Me.flag_zyg = I
Next I
cur_tim = GetCurrentTime() - start_tim
Screen.MousePointer = 1
Command2.Enabled = True
er_ex:
Exit Sub
er_chek:
Err_Ck (Err)
Resume er_ex
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.BackColor = Orange
End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.BackColor = Grey
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo form_exit

If Motor_Input <> "" And Motor_Preheated = 1 Then
        cmd$ = "!" + Motor_Output_off + ":"
        RobSend (cmd$)
End If
ValveReset
If zyg_err <> " " Then MSComm1.PortOpen = False
End
form_exit:
Resume Next
End Sub

Private Sub Label16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd$ = "!" + Central_Water_On + ":"
RobSend (cmd$)
Label16.BackColor = Green
End Sub

Private Sub Label16_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd$ = "!" + Central_Water_Off + ":"
RobSend (cmd$)
Label16.BackColor = Red
End Sub

Private Sub Label42_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd$ = "!" + Alarm2_on + ":"
RobSend (cmd$)
Label42.BackColor = Green
End Sub

Private Sub Label42_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd$ = "!" + Alarm2_off + ":"
RobSend (cmd$)
Label42.BackColor = Red
End Sub


Private Sub Label43_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label43.BackColor = Orange
End Sub

Private Sub Label43_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label43.BackColor = Grey
End Sub

Private Sub rezero_Click()
On Error GoTo Rez_er
'If system_ready = 0 Then Exit Sub
MSComm1.Output = Balance_Clear + Chr$(13)
MilSec 500
MSComm1.Output = Balance_Zero + Chr$(13)
Me.zyg_display = zygisi4(0)
Exit Sub
Rez_er:
Err_Ck (Err)
Resume Next
End Sub

Private Sub Form_Load()
If App.PrevInstance Then End
Screen.MousePointer = 11
SplashFrm.Show

' asw --> integer
'msg_ex:
 'MsgBox "Some programs are still running. Please close all Talos programs or close  the computer and start again ...", , "Talos : Can not  load Robot Operating System"
'End
'Declare Function GetModuleHandle Lib "Kernel" (ByVal IpModuleName As String) As Integer
'If GetModuleHandle("Coordin.exe") Then GoTo msg_ex
'If GetModuleHandle("Joystick.exe") Then GoTo msg_ex
If GetModuleHandle("Chek_Hom.exe") Then GoTo msg_ex
'If GetModuleHandle("Em6000.exe") Then GoTo msg_ex
'If GetModuleHandle("Ed6000.exe") Then GoTo msg_ex
'If GetModuleHandle("Start.exe") Then GoTo msg_ex
If GetModuleHandle("Syntages.exe") Then GoTo msg_ex
If GetModuleHandle("Syntage0.exe") Then GoTo msg_ex
If GetModuleHandle("Skones.exe") Then GoTo msg_ex



clue = ""
On Error GoTo er_form
clue = ""

FindValves
system_ready = IsOSLoaded(768)
'If FindOutput(Motor_Output_Number) = 0 Then Motor_Preheated = 1 Else Motor_Preheated = 0
zyg_Settings = Balance_Settings
  valve_mixed_on = "!" + Cold_Water_On + ":" + "!" + Hot_Water_On + ":"
  valve_mixed_off = "!" + Cold_Water_Off + ":" + "!" + Hot_Water_Off + ":"
Combo1.AddItem "1%"
Combo1.AddItem "5%"
For I = 1 To 12
Combo1.AddItem Trim(Str(I * 10)) + "%"
Next
         
If CheckForNull(Ready_Bit) = 0 Then MsgBox "Δεν είναι  δηλωμένο το Ok Bit. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός = 2 ). Η διαδικασία τερματίζεται...", , "Talos : Trouble Shooting Utility.": End
If CheckForNull(skon_motor_on) = 0 Then MsgBox "Δεν είναι  δηλωμένο το output του κινητήρα της σκόνης. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός = 4 ). Η διαδικασία τερματίζεται...", , "Talos : Trouble Shooting Utility.": End
If CheckForNull(skon_motor_off) = 0 Then MsgBox "Δεν είναι  δηλωμένο το output του κινητήρα της σκόνης. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός = 4 ). Η διαδικασία τερματίζεται...", , "Talos : Trouble Shooting Utility.": End
If CheckForNull(Cold_Water_On) = 0 Then MsgBox "Δεν είναι  δηλωμένο το output της βαλβίδας του κρύου νερού. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός = 5 ). Η διαδικασία τερματίζεται...", , "Talos : Trouble Shooting Utility.": End
If CheckForNull(Cold_Water_Off) = 0 Then MsgBox "Δεν είναι  δηλωμένο το output της βαλβίδας του κρύου νερού. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός = 5 ). Η διαδικασία τερματίζεται...", , "Talos : Trouble Shooting Utility.": End
If CheckForNull(Hot_Water_On) = 0 Then MsgBox "Δεν είναι  δηλωμένο το output της βαλβίδας του ζεστού νερού. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός = 6 ). Η διαδικασία τερματίζεται...", , "Talos : Trouble Shooting Utility.": End
If CheckForNull(Hot_Water_Off) = 0 Then MsgBox "Δεν είναι  δηλωμένο το output της βαλβίδας του ζεστού νερού. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός = 6 ). Η διαδικασία τερματίζεται...", , "Talos : Trouble Shooting Utility.": End
If CheckForNull(mixer_on) = 0 Then MsgBox "Δεν είναι  δηλωμένο το output του κινητήρα της ανάδευσης. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός = 7 ). Η διαδικασία τερματίζεται...", , "Talos : Trouble Shooting Utility.": End
If CheckForNull(mixer_off) = 0 Then MsgBox "Δεν είναι  δηλωμένο το output του κινητήρα της ανάδευσης. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός = 7 ). Η διαδικασία τερματίζεται...", , "Talos : Trouble Shooting Utility.": End
If CheckForNull(Alarm1_on) = 0 Then MsgBox "Δεν είναι  δηλωμένο το outpout του Βομβητή. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός = 11 ). Η διαδικασία τερματίζεται...", , "Talos : Trouble Shooting Utility.": End
If CheckForNull(Alarm1_off) = 0 Then MsgBox "Δεν είναι  δηλωμένο το outpout του Βομβητή. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός = 11 ). Η διαδικασία τερματίζεται...", , "Talos : Trouble Shooting Utility.": End
If CheckForNull(Balance_Port) = 0 Then MsgBox "Δεν είναι  δηλωμένη η πόρτα της ζυγαριάς. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός = 8 ).Η διαδικασία τερματίζεται...", , "Talos : Trouble Shooting Utility.": End

ValveReset

SendVeloc (0.1)
anix_zygaria
zyg_dec = zyg_decimals()
Joystick.Move 0, 0

On Error Resume Next
 asw = FindFonts()
  For I = 0 To Me.Controls.Count - 1
    Me.Controls(I).Font = asw
 Next

asw = IsOSLoaded(768)
If asw = 0 Then
  For I = 0 To 7
  term(I).Enabled = False
  term(I).BackColor = Grey
  Next I
  Grip.Enabled = False
  Grip.BackColor = Grey
  Timer1.Enabled = False
  check1.Enabled = False
  clue = "Robot Moving Programs not Loaded. Follow the Starting Procedure ..."
End If
Unload SplashFrm
Screen.MousePointer = 1
Exit Sub
er_form:
If Err = 94 Then
   MsgBox "Λείπουν στοιχεία από το αρχείο c:\Talos\coord_tb.mdb !", , "Talos"
End If
Err_Ck (Err)
Resume Next
msg_ex:
   MsgBox "Some Talos programs are still running. Can not run Joystick ...", , "Talos : Warning"
End

End Sub



Private Sub Grip_Click()
 On Error Resume Next
  If Gripper = 1 Then  '0 =ανοικτός
       cmd$ = "!" + CStr(Gripper_On) + ":"
       RobSend (cmd$)
  Else
       cmd$ = "!" + CStr(Gripper_Off) + ":"
       RobSend (cmd$)
  End If
End Sub

Private Sub Grip_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Grip.BackColor = Orange
End Sub

Private Sub Grip_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Grip.BackColor = Green
End Sub

Private Sub Rezero_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Rezero.BackColor = Orange '&H00FFFFFF&
End Sub

Private Sub Rezero_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Rezero.BackColor = Grey
End Sub

Private Sub term_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Len(Motor_Input) > 0 And FindInput(CDbl(Motor_Input)) = 0 Then
  If Motor_Preheating > 0 Then
       msg = LoadLabel(465) + Chr$(13) + LoadLabel(466) + Chr$(13) + LoadLabel(467)
      Title = LoadLabel(490)
      response = MsgBox(msg, 276, Title)
            If response = 7 Then Exit Sub
             cmd$ = "!" + Motor_Output_on + ":"
              RobSend (cmd$)
   x_hom = Shell("C:\talos\chek_hom.exe", 1)
   Do While GetModuleUsage(x_hom) <> 0: Me.WindowState = 1: DoEvents: Loop
    
   Me.WindowState = 0
           Motor_Preheated = 1
           cmd$ = "!" + Motor_Output_on + ":"
           RobSend (cmd$)
           SendVeloc (0.2)
        Exit Sub
  End If
End If
If Motor_Input <> "" And FindOutput(Motor_Output_Number) = 0 Then
        Motor_Preheated = 1
        cmd$ = "!" + Motor_Output_on + ":"
        RobSend (cmd$)
End If
Select Case Index
 Case 0
     cmd$ = "!d" + Str(-5000000) + ":"
      RobSend (cmd$)
      cmd$ = "!go1:"
      RobSend (cmd$)
Case 1
     cmd$ = "!d" + Str(5000000) + ":"
       RobSend (cmd$)
       cmd$ = "!go1:"
       RobSend (cmd$)
Case 2
    cmd$ = "!d," + Str(-5000000) + ":"
      RobSend (cmd$)
      cmd$ = "!go,1:"
      RobSend (cmd$)
Case 3
    cmd$ = "!d," + Str(5000000) + ":"
      RobSend (cmd$)
      cmd$ = "!go,1:"
      RobSend (cmd$)
Case 4
    cmd$ = "!d,," + Str(-5000000) + ":"
      RobSend (cmd$)
      cmd$ = "!go,,1:"
      RobSend (cmd$)
Case 5
     cmd$ = "!d,," + Str(5000000) + ":"
      RobSend (cmd$)
      cmd$ = "!go,,1:"
      RobSend (cmd$)
Case 6
    cmd$ = "!d,,," + Str(-5000000) + ":"
    RobSend (cmd$)
    cmd$ = "!go,,,1:"
    RobSend (cmd$)
Case 7
    cmd$ = "!d,,," + Str(5000000) + ":"
    RobSend (cmd$)
    cmd$ = "!go,,,1:"
    RobSend (cmd$)
End Select
   If term(Index).BackColor = Green Then term(Index).BackColor = Orange
End Sub
Private Sub term_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If term(Index).BackColor = Red Then
    Else
        term(Index).BackColor = Green
    End If
    cmd$ = "!s:"
   RobSend (cmd$)
   MilSec (1500)
   cmd$ = "!k:"
   RobSend (cmd$)
End Sub


Private Sub Text1_Change()

End Sub

Private Sub Timer1_Timer()
Static clue1, clue2, clue3, clue4, clue5
On Error GoTo er_tim
clue = clue1 + clue2 + clue3 + clue4 + clue5 + clue6
If clue = "" Then
    Image1.Visible = True
    Image2.Visible = False
Else
    Image1.Visible = False
    Image2.Visible = True
End If
system_ready = IsOSLoaded(Device_Address%)
If system_ready = 0 Then
Label18 = "Not Loaded"
Label18.ForeColor = Red
clue1 = " Operating System not Loaded :"
Else
Label18 = "Loaded"
Label18.ForeColor = Blue
End If

Gripper = FindInput(Gripper_input)
asw = FindInput(Door_input)
If asw = 0 Then
    Door.ForeColor = Blue
    Door = "Closed"
    clue2 = ""
Else:
    Door.ForeColor = Red
    Door = "Open"
    clue2 = " Door is Open :"
End If
asw = FindInput(Pause_input)
If asw = 0 Then
Pause.ForeColor = Blue
Pause = "Inactive"
clue3 = ""
Else
Pause.ForeColor = Red
Pause = "Active"
clue3 = " Pause Active :"
End If
Cut_Off = Crush
asw = Check_CutOff()
If asw = 0 Then
    Cut_Off.ForeColor = Blue
    Crush = "Inactive"
    Crush.ForeColor = Blue
    clue4 = ""
    clue5 = ""
ElseIf asw = 1 Then
    Crush = "Active"
    Cut_Off.ForeColor = Red
    Crush.ForeColor = Red
    clue4 = " Crash Sensor Active :"
    clue5 = " Emergency Stop Active :"
End If

If Gripper = 0 Then Grip.Caption = "Gripper Opened" Else Grip.Caption = "Gripper Closed"
   Result% = Request_Status(Device_Address%)
    ' ******* Update Motor Position ****************
    Set_Pointer Device_Address%, AXIS1_MOTOR
    For I = 0 To 3
        Read_Status Device_Address%, status_high%, status_low%, status&
        position(I).Caption = Format$(status&, "###,###")
    Next I
    ' ******* Update Encoder Position **************
    Set_Pointer Device_Address%, AXIS1_ENCODER
    For I = 0 To 3
       Read_Status Device_Address%, status_high%, status_low%, status&
        LBL_Encoder(I).Caption = Format$(status&, "###,###")
    Next I
    ' ******* Update Motor Velocity **************
    Set_Pointer Device_Address%, AXIS1_VELOCITY
    For I = 0 To 3
        Read_Status Device_Address%, status_high%, status_low%, status&
       LBL_Velocity(I).Caption = Format$(status& / 4000, "0.0")
    Next I
    ' ******* Update Limit Status ****************
    Begin_Ccw_Mask = CCW1_MASK
    Begin_Cw_Mask = CW1_MASK
    Begin_Home_Mask = HOME1_MASK
    Set_Pointer Device_Address%, LIMIT_STATUS
    Read_Status Device_Address%, status_high%, status_low%, status&
    For I = 0 To 3
        If (status_high% And Begin_Ccw_Mask) > 0 Then
            '***** Limit switch is closed *******
            If (Last_Ccw_State(I) = LIMIT_OPEN) Then
                Last_Ccw_State(I) = LIMIT_CLOSED
                term(I * 2).BackColor = Green
            End If
        Else
            '***** Limit switch is open, limit is active *******
            If (Last_Ccw_State(I) = LIMIT_CLOSED) Then
                Last_Ccw_State(I) = LIMIT_OPEN
               term(I * 2).BackColor = Red
            End If
        End If
        If (status_high% And Begin_Cw_Mask) > 0 Then
            '***** Limit switch is closed *******
            If (Last_Cw_State(I) = LIMIT_OPEN) Then
                Last_Cw_State(I) = LIMIT_CLOSED
                term(I * 2 + 1).BackColor = Green
            End If
        Else
            '***** Limit switch is open, limit is active *******
            If (Last_Cw_State(I) = LIMIT_CLOSED) Then
                Last_Cw_State(I) = LIMIT_OPEN
                term(I * 2 + 1).BackColor = Red
            End If
        End If
        Begin_Ccw_Mask = Begin_Ccw_Mask * &H4
        Begin_Cw_Mask = Begin_Cw_Mask * &H4
        Begin_Home_Mask = Begin_Home_Mask * &H2
    Next I
    
Exit Sub

er_tim:
 If Err = 2467 Then
  MsgBox LoadLabel(43), , "Talos"
 End
 End If
Resume Next

End Sub
Private Sub Label23_Click()
cmd$ = "!" + mixer_off + ":"
RobSend (cmd$)
End Sub

Private Sub Label25_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd$ = "!" + valve_mixed_on + ":"
RobSend (cmd$)
Label25.BackColor = Green
End Sub

Private Sub Label25_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd$ = "!" + valve_mixed_off + ":"
RobSend (cmd$)
Label25.BackColor = Red
End Sub

Private Sub Label26_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd$ = "!" + valve_2_on + ":"
RobSend (cmd$)
Label26.BackColor = Green
End Sub

Private Sub Label26_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd$ = "!" + valve_2_off + ":"
RobSend (cmd$)
Label26.BackColor = Red
End Sub

Private Sub Label33_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd$ = "!" + valve_1_on + ":"
RobSend (cmd$)
Label33.BackColor = Green
End Sub

Private Sub Label33_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd$ = "!" + valve_1_off + ":"
RobSend (cmd$)
Label33.BackColor = Red
End Sub

Private Sub Label34_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd$ = "!" + Alarm1_on + ":"
RobSend (cmd$)
Label34.BackColor = Green
End Sub

Private Sub Label34_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd$ = "!" + Alarm1_off + ":"
RobSend (cmd$)
Label34.BackColor = Red
End Sub

Private Sub Label35_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd$ = "!" + Hot_Water_On + ":"
RobSend (cmd$)
Label35.BackColor = Green
End Sub

Private Sub Label35_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd$ = "!" + Hot_Water_Off + ":"
RobSend (cmd$)
Label35.BackColor = Red
End Sub

Private Sub Label36_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd$ = "!" + mixer_on + ":"
RobSend (cmd$)
Label36.BackColor = Green
End Sub

Private Sub Label36_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd$ = "!" + mixer_off + ":"
RobSend (cmd$)
Label36.BackColor = Red
End Sub

Private Sub Label37_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd$ = "!" + Cold_Water_On + ":"
RobSend (cmd$)
Label37.BackColor = Green
End Sub

Private Sub Label37_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd$ = "!" + Cold_Water_Off + ":"
RobSend (cmd$)
Label37.BackColor = Red
End Sub

Private Sub Label38_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd$ = "!" + skon_motor_on + ":"
RobSend (cmd$)
Label38.BackColor = Green
End Sub

Private Sub Label38_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd$ = "!" + skon_motor_off + ":"
RobSend (cmd$)
Label38.BackColor = Red
End Sub

Private Sub zyg_timer_Timer()
zyg_dec = zyg_decimals()
If zyg_dec = "2" Then
    zyg_display = Format(zygisi4(0), "##0.00")
     zyg_display.BackColor = Green
     clue6 = ""
ElseIf zyg_dec = "3" Then
    zyg_display = Format(zygisi4(0), "##0.000")
     zyg_display.BackColor = Green
     clue6 = ""
Else
        zyg_display = "-"
        zyg_display.BackColor = Red
        clue6 = " Scale is not responding :"
End If
zyg_display.Refresh

End Sub

Private Sub Timer2_Timer()
If zyg_err <> "" Or clue6 <> "" Then
    clue6 = " Scale port is not available "
    Scale_dis.ForeColor = Red
    Scale_dis = "Inactive"
Else
    Scale_dis.ForeColor = Blue
    Scale_dis = "Active"
End If
 zyg_dec = zyg_decimals()
If zyg_dec = "2" Then
    zyg_display = Format(zygisi4(0), "##0.00")
     zyg_display.BackColor = Green
     clue6 = ""
ElseIf zyg_dec = "3" Then
    zyg_display = Format(zygisi4(0), "##0.000")
     zyg_display.BackColor = Green
     clue6 = ""
Else
        zyg_display = "-"
        zyg_display.BackColor = Red
        clue6 = " Scale is not responding "
    Scale_dis.ForeColor = Red
    Scale_dis = "Inactive"
End If
zyg_display.Refresh
 Exit Sub

End Sub



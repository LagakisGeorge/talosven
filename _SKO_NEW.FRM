VERSION 4.00
Begin VB.Form skones 
   BackColor       =   &H00808000&
   ClientHeight    =   6840
   ClientLeft      =   1170
   ClientTop       =   1470
   ClientWidth     =   10725
   ControlBox      =   0   'False
   BeginProperty Font 
      name            =   "System"
      charset         =   1
      weight          =   700
      size            =   9.75
      underline       =   0   'False
      italic          =   0   'False
      strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H008080FF&
   Height          =   7245
   Icon            =   "_SKO_NEW.frx":0000
   Left            =   1110
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6840
   ScaleWidth      =   10725
   Top             =   1125
   Width           =   10845
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Δυό μαζί διαλύματα"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   103
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox redb 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   2000
      Left            =   9750
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   500
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2955
      Top             =   2760
   End
   Begin VB.TextBox Z_Met_dis 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5700
      Locked          =   -1  'True
      TabIndex        =   101
      Top             =   525
      Width           =   525
   End
   Begin VB.TextBox Tare_Dis 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8775
      TabIndex        =   99
      Top             =   525
      Width           =   780
   End
   Begin VB.TextBox Ca_Factor_dis 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataSource      =   "=100"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   97
      Top             =   525
      Width           =   915
   End
   Begin VB.CommandButton exodos 
      Caption         =   "Εξοδος"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   1665
      TabIndex        =   96
      Top             =   3330
      Width           =   1200
   End
   Begin VB.TextBox ximiko 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataSource      =   "=100"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2355
      Locked          =   -1  'True
      TabIndex        =   95
      Top             =   1515
      Width           =   3825
   End
   Begin VB.TextBox nero_dial 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8175
      Locked          =   -1  'True
      TabIndex        =   94
      Text            =   "nero_dial"
      Top             =   15
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox s_mixtim 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8460
      Locked          =   -1  'True
      TabIndex        =   93
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox fin_con 
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5940
      Locked          =   -1  'True
      TabIndex        =   91
      Top             =   -60
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.TextBox Cur_Maker 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9000
      Locked          =   -1  'True
      TabIndex        =   89
      Text            =   "0"
      Top             =   1035
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox sol_maker 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8835
      Locked          =   -1  'True
      TabIndex        =   88
      Text            =   "20"
      Top             =   1035
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox nois 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   73
      Top             =   300
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox vim 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8655
      Locked          =   -1  'True
      TabIndex        =   72
      Top             =   30
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox upm4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9945
      Locked          =   -1  'True
      TabIndex        =   71
      Top             =   840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox var67 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9900
      Locked          =   -1  'True
      TabIndex        =   70
      Top             =   570
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Διακοπή  Προγράμματος"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   90
      TabIndex        =   69
      Top             =   3330
      Width           =   1515
   End
   Begin VB.TextBox zyg_port 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8805
      Locked          =   -1  'True
      TabIndex        =   68
      Top             =   300
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.TextBox mixtim_fin 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8820
      Locked          =   -1  'True
      TabIndex        =   67
      Top             =   30
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox nero2_off 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9780
      Locked          =   -1  'True
      TabIndex        =   65
      Top             =   855
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox nero2_on 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9735
      Locked          =   -1  'True
      TabIndex        =   64
      Top             =   570
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox baros_pot 
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2970
      TabIndex        =   63
      Top             =   1050
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox thermo1 
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3435
      Locked          =   -1  'True
      TabIndex        =   61
      Top             =   2850
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\SKON_TB.MDB"
      Exclusive       =   0   'False
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3345
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Εντολες"
      Top             =   1005
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.TextBox periek_dis 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataSource      =   "=100"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   60
      Top             =   1515
      Width           =   660
   End
   Begin VB.TextBox z_buk 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9555
      Locked          =   -1  'True
      TabIndex        =   59
      Top             =   570
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox y_buk 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9420
      Locked          =   -1  'True
      TabIndex        =   58
      Top             =   570
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox x_buk 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9285
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   570
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox x_pot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9195
      Locked          =   -1  'True
      TabIndex        =   56
      Top             =   840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox y_pot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9330
      Locked          =   -1  'True
      TabIndex        =   55
      Top             =   840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox z_pot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9465
      Locked          =   -1  'True
      TabIndex        =   54
      Top             =   840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox var88 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8955
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   570
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox var93 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8790
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   570
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox upm 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8835
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   825
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox valve_hot_off 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9570
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   300
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox valve_hot_on 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9390
      Locked          =   -1  'True
      TabIndex        =   49
      Top             =   300
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox valve_cold_off 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9165
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   30
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox valve_cold_on 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8985
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   30
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox nero1_on 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9105
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   570
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox nero1_off 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9015
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox thermo2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9705
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   30
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.CommandButton start_dosing 
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8415
      TabIndex        =   42
      Top             =   945
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox mixtim 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5250
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   2580
      Width           =   915
   End
   Begin VB.TextBox Διαρκεια 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8265
      TabIndex        =   38
      Top             =   525
      Width           =   495
   End
   Begin VB.TextBox Ληξη 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7275
      TabIndex        =   36
      Top             =   525
      Width           =   1020
   End
   Begin VB.TextBox Αρχη 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6240
      TabIndex        =   34
      Top             =   525
      Width           =   1050
   End
   Begin VB.TextBox nerostox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataSource      =   "=100"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2610
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   1035
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox nerot 
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5235
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   3585
      Width           =   930
   End
   Begin VB.TextBox fin_con1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataSource      =   "=100"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   735
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   1515
      Width           =   1545
   End
   Begin VB.TextBox nero2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5235
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   3225
      Width           =   930
   End
   Begin VB.TextBox nero1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5295
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   1890
      Width           =   885
   End
   Begin VB.TextBox periek 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataSource      =   "=100"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   400
         size            =   6
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Buk 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   520
      Width           =   750
   End
   Begin VB.TextBox Pot 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataSource      =   "=100"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3945
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   520
      Width           =   810
   End
   Begin VB.TextBox Druv_dis 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2445
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   520
      Width           =   615
   End
   Begin VB.TextBox CurQ_Dis 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   1155
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   520
      Width           =   1125
   End
   Begin VB.TextBox Target_Dis 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   520
      Width           =   1020
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Εναρξη Παρασκευής"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   120
      TabIndex        =   66
      Top             =   2040
      Width           =   1515
   End
   Begin VB.CommandButton Finish 
      Caption         =   "Επιστροφή"
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9000
      TabIndex        =   18
      Top             =   3600
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.CommandButton Βαλε_Νερο2 
      Caption         =   "Νερό Αραίωσης"
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9000
      TabIndex        =   17
      Top             =   3225
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.CommandButton Mixing 
      Caption         =   "Ανάδευση"
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9000
      TabIndex        =   15
      Top             =   2580
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.CommandButton Sendmix 
      Caption         =   "Προς Ανάδευση"
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9000
      TabIndex        =   14
      Top             =   2265
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.CommandButton Beready 
      Caption         =   "Προετοιμασία"
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9000
      TabIndex        =   13
      Top             =   1260
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.CommandButton Βαλε_Νερο1 
      Caption         =   "Νερό Διάλυσης"
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9000
      TabIndex        =   12
      Top             =   1920
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.CommandButton Dosing 
      Caption         =   "Δοσομέτρηση"
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9000
      TabIndex        =   9
      Top             =   1575
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.CommandButton Backmix 
      Caption         =   "Επιστροφή Ζυγαριά"
      Enabled         =   0   'False
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9000
      TabIndex        =   16
      Top             =   2925
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.PictureBox MSComm1 
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4710
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   104
      Top             =   1035
      Width           =   1200
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "_SKO_NEW.frx":0442
      Height          =   2955
      Left            =   0
      OleObjectBlob   =   "_SKO_NEW.frx":0450
      TabIndex        =   102
      TabStop         =   0   'False
      Top             =   3960
      Width           =   9645
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "Τάρα"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   8700
      TabIndex        =   100
      Top             =   195
      Width           =   825
   End
   Begin VB.Label Label40 
      BackColor       =   &H00808000&
      Caption         =   "Ρευστότητα Σκόνης"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4815
      TabIndex        =   98
      Top             =   75
      Width           =   1170
   End
   Begin VB.Label Label39 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Δόσεις"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   2415
      TabIndex        =   92
      Top             =   315
      Width           =   705
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Σκόνη Δοσομετρημένη"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   930
      TabIndex        =   3
      Top             =   45
      Width           =   1500
   End
   Begin VB.Label Label38 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Περιεκτικότητες ( w / w ) "
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   60
      TabIndex        =   87
      Top             =   990
      Width           =   2415
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Τελική"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1230
      TabIndex        =   86
      Top             =   1245
      Width           =   900
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Target"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   75
      TabIndex        =   85
      Top             =   1245
      Width           =   705
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ΔΙΑΔΙΚΑΣΙΕΣ"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   -1  'True
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   6210
      TabIndex        =   84
      Top             =   975
      Width           =   2150
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Νερό Αραίωσης"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   2865
      TabIndex        =   83
      Top             =   3285
      Width           =   1500
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Νερό Διάλυσης"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2910
      TabIndex        =   82
      Top             =   1920
      Width           =   1470
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Προσθήκη Σκόνης"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   6210
      TabIndex        =   81
      Top             =   1575
      Width           =   2145
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Προετοιμασία"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   6210
      TabIndex        =   80
      Top             =   1260
      Width           =   2150
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Προσθ.Νερού Διάλυσης"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   6210
      TabIndex        =   79
      Top             =   1905
      Width           =   2145
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Προς Ανάδευση"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   6210
      TabIndex        =   78
      Top             =   2265
      Width           =   2150
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ανάδευση"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   6210
      TabIndex        =   77
      Top             =   2595
      Width           =   2150
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Επιστροφή Ζυγαριά"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   6210
      TabIndex        =   76
      Top             =   2940
      Width           =   2150
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Πρ.Νερού Αραίωσης"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   6210
      TabIndex        =   75
      Top             =   3270
      Width           =   2150
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Τελ. Επιστροφή"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   6210
      TabIndex        =   74
      Top             =   3600
      Width           =   2150
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   4365
      TabIndex        =   62
      Top             =   1890
      Width           =   930
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mix.Time"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   3240
      TabIndex        =   41
      Top             =   2625
      Width           =   1980
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "(Sec)"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   8220
      TabIndex        =   39
      Top             =   225
      Width           =   570
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "Λήξη"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   7110
      TabIndex        =   37
      Top             =   210
      Width           =   915
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "Εναρξη"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   6225
      TabIndex        =   35
      Top             =   210
      Width           =   780
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Συνολο Νερου (mgr)"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   3225
      TabIndex        =   32
      Top             =   3585
      Width           =   2025
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   4350
      TabIndex        =   29
      Top             =   3225
      Width           =   900
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O.K."
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   8400
      TabIndex        =   25
      Top             =   3600
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O.K."
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   8385
      TabIndex        =   24
      Top             =   3270
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O.K."
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   8385
      TabIndex        =   23
      Top             =   2925
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   8385
      TabIndex        =   22
      Top             =   2595
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O.K."
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   8385
      TabIndex        =   21
      Top             =   2265
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O.K."
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   8385
      TabIndex        =   20
      Top             =   1905
      Width           =   585
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O.K."
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   8385
      TabIndex        =   19
      Top             =   1260
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Μέσο_Ζ 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Μέσο Ζ"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5730
      TabIndex        =   11
      Top             =   1050
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O.K."
      BeginProperty Font 
         name            =   "Arial"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   8385
      TabIndex        =   10
      Top             =   1575
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Μπουκάλι Σκόνης"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Left            =   3105
      TabIndex        =   8
      Top             =   60
      Width           =   825
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Ποτήρι Διάλυσης"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3900
      TabIndex        =   7
      Top             =   75
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Στόχος Σκόνης"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   45
      Width           =   1020
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Εκτέλεση Τελευταίας Παρασκευής"
      BeginProperty Font 
         name            =   "System"
         charset         =   1
         weight          =   700
         size            =   9.75
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   90
      TabIndex        =   90
      Top             =   2730
      Visible         =   0   'False
      Width           =   2670
   End
End
Attribute VB_Name = "skones"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Public arg As Integer, Ca_Factor As Single, nero1_ef As Long, curq As Long
Dim nero1_count   As Long, Nero2_Count  As Long, Ner_Stox As Long
Dim z_syntetag As Long, z_met As Single, Druv As Integer, Target As Long
Dim baros_pot_max As Long, baros_pot_min As Long, Tare As Long, Q_Max_Quantity As Long
Dim Last_Quantity As Long, Anevasma As Integer, Last_Dosage As Long
Dim Telos_Dialymaton  ' 1=τελείωσαν τα διαλύματα      0=υπάρχουν ακόμη διαλύματα
Dim Last_z_met As Double
Dim Proti_Fora As Boolean
Dim Skacimo

Dim nerostox1, periek1, periek_dis1, buk1, Pot1, mixtim1, s_mixtim1, mixtim_fin1, thermo1_1, thermo2_1, nero_dial_1, ximiko1
Dim nerostox2, periek2, periek_dis2, buk2, Pot2, mixtim2, s_mixtim2, mixtim_fin2, thermo1_2, thermo2_2, nero_dial_2, ximiko2

'
'
'
Function dosom_Skonh(ByVal m_Pot, ByVal m_buk, ByVal m_Target, ByVal m_nero_dial, ByVal m_mixtim, ByVal m_ximiko, ByVal m_Thermo1, ByVal m_thermo2, curq, nerot)
'-----------------------
Dim mm_Periek

Pot = m_Pot: buk = m_buk
Target = m_Target: nero_dial = m_nero_dial
mixtim = m_mixtim: ximiko = m_ximiko
thermo1 = m_Thermo1: thermo2 = m_thermo2
a = find_eidos_neron(thermo1, thermo2, nero1_on, nero1_off, nero2_on, nero2_off)
  
  
   If Pot = 0 Then Exit Function
  
  
  Rob_FeedBack (59)
  cmd$ = "!" + Prouermansh_On + ":"
  
  
  'Μεταφορά Ποτηριού
  RobSend (cmd$)
  Fere_Pot 'παράμετρος Pot
  cmd$ = "!" + Prouermansh_Off + ":"
  RobSend (cmd$)
  bar_pot = chek_pot  'οχι παραμετρος   ΕΛΕΓΧΕΙ ΑΝ ΠΟΤΗΡΙ="ok"
  
  If bar_pot = "not ok" Then
     curq = -1: CurQ_Dis = curq
            '  678 -9   "Δεν παρασκευάσθηκε  το διάλυμα  στο ποτήρι Νο ", " επειδή δεν βρέθηκε το κατάλληλο ποτήρι ... "
      Printer.Print mL_Res(678) + Str(Pot) + mL_Res(679)
     ' Printer.EndDoc
     rearange (1)   ' ΕΠΙΣΤΡΕΦΕΙ ΣΤΗΝ ΘΕΣΗ ΤΟΥ ΤΟ ΑΚΑΤΑΛΛΗΛΟ ΠΟΤΗΡΙ
     Ληξη = Time
     Διαρκεια = DateDiff("s", Αρχη, Ληξη)
     file_send   ' ΣΗΜΕΙΩΝΕΙ ΣΤΟ ΑΡΧΕΙΟ SKONES.MDB ΤΙ ΕΓΙΝΕ
     Exit Function
  End If


7
     ' If system_ready = 1 Then rezero_Click

 Rob_FeedBack (59)
 fere_buk 'παράμετρος BUK
                                    
 ' Δοσομέτρηση Σκόνης
 Command2.Enabled = False
 If InStr(1, ximiko, "*") > 0 Then
    Dosing_Click   ' ΔΙΑΔΙΚΑΣΙΑ  ΔΟΣΟΜΕΤΡΗΣΗΣ ΣΚΟΝΗΣ (Papakiriakos)
 Else
    curq = DOS_ALATI(Target)  'lagakis 24-2-2000
 End If
 On Error Resume Next
 Printer.Print ximiko + " Δοσομετρηθείσα σκόνη :" + Str(curq)
 On Error GoTo 0


 If curq <= 0 Then ' Το μπουκάλι σκόνης δεν τροφοδοτεί με σκόνη
       Rob_FeedBack (59)  ' ΠΕΡΙΜΕΝΕΙ 59 SEC NA ΤΕΛΕΙΩΣΕΙ ΤΟ ROBOT
       cmd$ = "!var12 =20000:"
       RobSend (cmd$)
       Rob_FeedBack (60)          ' (50)
       rearange (1) ' ΕΠΙΣΤΡΕΦΕΙ ΠΙΣΩ ΤΟ ΠΟΤΗΡΙ ΠΟΥ ΔΕΝ ΕΧΕΙ ΣΚΟΝΗ ,ΚΑΙ ΔΙΑΚΟΠΤΕΙ ΔΙΑΔΙΚΑΣΙΑ ΠΑΡΑΣΚΕΥΗΣ
       file_send
       Rob_FeedBack (60)
       Exit Function
End If
      
'      cmd$ = "!var12=20000:"
'      RobSend (cmd$)

'Προσθήκη Νερού Διάλυσης
Βαλε_Νερο1_Click    ' παραμετρος : nero_dial     ΠΡΟΣΘΕΤΕΙ ΤΟ ΝΕΡΟ ΔΙΑΛΥΣΗΣ

Command2.Enabled = True
 If nerot < 10000 Then
     While FindOutput(Val(Ready_Bit)) = 0: DoEvents: Wend
    secwait (1)
     rearange (1)
     curq = 0: CurQ_Dis = curq
     Ληξη = Time
    Διαρκεια = DateDiff("s", Αρχη, Ληξη)
   file_send
   While FindOutput(Val(Ready_Bit)) = 0: DoEvents: Wend
      secwait (2)
      cmd$ = "!var12=82:"   'Park
     RobSend (cmd$)
     secwait (2)
      While FindOutput(Val(Ready_Bit)) = 0: DoEvents: Wend
         cmd$ = "!" + Alarm1_on + ":"
         secwait (1)
          RobSend (cmd$)
          If nero1 < 20000 And nero1 >= 0 Then
              ' 680  "Με το άνοιγμα της βαλβίδας νερού , νερό δεν έρχεται. Η διαδικασία διακόπτεται..."
              MsgBox mL_Res(680), , "680.Talos"
         End If
         If nerot < 0 Then
                      '    "Με το κλείσιμο της βαλβίδας νερού , το νερό δεν σταματά. Η διαδικασία διακόπτεται..."
             MsgBox mL_Res(681), , "681.Talos"
        End If
        Telos_Dialymaton = 1
        Exit Function
End If
Command2.Enabled = True
Rob_FeedBack (59)

'παράμετρος αναδευση 1 ή 2
Sendmix_Click  ' ΠΑΙΡΝΕΙ ΤΟ ΠΟΤΗΡΙ ΚΑΙ ΤΟ ΠΑΕΙ ΣΤΗΝ αναδευση


DUM_DUM = Wait_Robot(0)
If DemoVersion = 1 Then
Else
   Mixing_Click   ' παράμετρος mixtim   ΜΕΤΡΑΕΙ ΤΟΝ ΧΡΟΝΟ ΑΝΑΔΕΥΣΗΣ
End If


End Function

Sub dyo_start_dosing_click()
'---------------------------
  Dim mm_Periek
  
  target1 = periek1 * nerostox1
  curq1 = dosom_Skonh(Pot1, buk1, target1, nero_dial1, mixtim1, ximiko1, thermo1_1, thermo2_1, curq, nerot)
  
  target2 = periek2 * nerostox2
  curq2 = dosom_Skonh(Pot2, buk2, target2, nero_dial2, mixtim2, ximiko2, thermo1_2, thermo2_2, curq, nerot)




'φερε από διάλυση σε ζυγό,ρίξε επιπλέον νερό,φέρε σε τελική ανάμιξη
X = 1
   nerostox = nerostox1
   teliko_nero   ' kai anadeysh

X = 2
   nerostox = nerostox2
   teliko_nero  'και ανάδευση
   
   
  Command2.Enabled = True

X = 1
  epistrofh_pothrioy   ' ΠΗΓΑΙΝΕΙ ΤΟ ΔΙΑΛΥΜΑ ΣΤΗΝ ΑΝΑΔΕΥΣΗ ΓΙΑ ΤΕΛΙΚΗ ΜΙΞΗ ΚΑΙ ΤΟ ΕΠΙΣΤΡΕΦΕΙ ΣΤΗΝ ΘΕΣΗ ΤΟΥ ΠΟΤΗΡΙΟΥ
X = 2
  epistrofh_pothrioy

vim = 0

Exit Sub




End Sub
Sub epistrofh_pothrioy():
'------------------------------------
cmd$ = "!var12=" + Str(Pot + 200) + ":"
RobSend (cmd$)
' Rob_FeedBack (59)
DUM_DUM = Wait_Robot(0)
 Label19.Visible = True
Ληξη = Now
Διαρκεια = DateDiff("s", Αρχη, Ληξη)
file_send
skones.Refresh
  Screen.MousePointer = 0


End Sub



Function find_eidos_neron(ByVal t1, ByVal t2, nero1_on, nero1_off, nero2_on, nero2_off)

Select Case Val(t1)
           Case 1
             nero1_on = "!" + Cold_Water_On + ":"
             nero1_off = "!" + Cold_Water_Off + ":"
         Case 2
            nero1_on = "!" + Cold_Water_On + ":" + "!" + Hot_Water_On + ":"
            nero1_off = "!" + Cold_Water_Off + ":" + "!" + Hot_Water_Off + ":"
          Case 3
            nero1_on = "!" + Hot_Water_On + ":"
             nero1_off = "!" + Hot_Water_Off + ":"
         End Select
        
      Select Case Val(t2)
           Case 1
             nero2_on = "!" + Cold_Water_On + ":"
             nero2_off = "!" + Cold_Water_Off + ":"
         Case 2
            nero2_on = "!" + Cold_Water_On + ":" + "!" + Hot_Water_On + ":"
            nero2_off = "!" + Cold_Water_Off + ":" + "!" + Hot_Water_Off + ":"
          Case 3
            nero2_on = "!" + Hot_Water_On + ":"
             nero2_off = "!" + Hot_Water_Off + ":"
         End Select



End Function

Sub teliko_nero(m_Curq, m_nerot, m_thermo2)
'------------------------
'Backmix_Click ' ΠΑΙΡΝΕΙ ΑΠΟ ΑΝΑΔΕΥΣΗ ΚΑΙ ΤΟ ΠΑΕΙ ΣΤΗΝ  ΖΥΓΑΡΙΑ
Dim bar_chek
 
 rezero_Click
 Rob_FeedBack (59)           ' Eλεγχος Robot στα 30 sec
 cmd$ = "!var12=101:"
 RobSend (cmd$)
 DUM_DUM = Wait_Robot(0)
 Label17.Visible = True
 bar = ZYGIS3()
 If bar < 100000 Then
        MsgBox mL_Res(682), , "Talos" '"The robot did not take the bottle from the mixer to the scale !"
        End
 End If
   If bar > Val(baros_pot) + curq Then nero1 = bar - Val(baros_pot) - curq 'ξαναμέτρα το νερο λογω εξάτμισης
   nerot = nero1
   fin_con = curq / (nerot + curq): fin_con1 = Format(fin_con, "0.0000 %")
 If Val(vim) = 0 Then
   vim = Val(var67) + upm4 * Val(Water_Bypass) 'Μετατόπιση 4ου άξονα 30 χιλ δεξιά για να μην στάζει
   cmd$ = "!d,,," + Str(vim) + ":!go,,,1:"
   RobSend (cmd$)
 End If
 vim = 0
 skones.Refresh
'-------------------



Command2.Enabled = False

Βαλε_Νερο2_Click   ' ΒΑΖΕΙ ΝΕΡΟ ΑΡΑΙΩΣΗΣ ΜΕΧΡΙ ΤΕΛΙΚΗΣ ΠΕΡΙΕΚΤΙΚΟΤΗΤΟΣ

mm_Periek = IIf(Val(nerot) > 0, curq / (curq + Val(nerot)) * 100, 0)
On Error Resume Next
   Printer.Print " συνολ.νερό :" + nerot + "   Περιεκτικότητα " + Format(mm_Periek, "0.###0") + "%"
On Error GoTo 0


' αναδευση τελικη
If mixtim_fin > 0 Then
   Sendmix_Click
   Mixing.Caption = mL_Res(683) ' "Τελική Ανάδευση"
   Label29.Caption = mL_Res(684)   '"Final Mixing  "
   Label16.Caption = "0"
   mixtim = mixtim_fin
      cmd$ = "!" + mixer_on + ":"
       RobSend (cmd$)
       mixing_make (mixtim_fin)
       cmd$ = "!" + mixer_off + ":"
       RobSend (cmd$)
End If
Rob_FeedBack (59)



End Sub

Function Zygis0_time(X As Single)
  Dim counter, starter
    ' Flush the input buffer.
       MSComm1.Output = "SEND" + Chr$(13)
       MSComm1.InBufferCount = 0
    starter = GetCurrentTime()
    Do
         dummy = DoEvents()
         If GetCurrentTime() - starter > X Then
            Zygis0_time = -999000999
            Exit Function
         End If
       ' If there is data in the buffer, then read it.
       If MSComm1.InBufferCount Then
          FromModem$ = FromModem$ + MSComm1.Input
          ' Check for "OK".
          If InStr(FromModem$, Chr$(13)) Then
                If Mid$(FromModem$, 5, 1) = "." Or Mid$(FromModem$, 6, 1) = "." Or Mid$(FromModem$, 7, 1) = "." Then
                    Exit Do
                End If
          End If
       End If
    Loop
    zyg = Val(Left$(FromModem$, 9))
    Zygis0_time = zyg * 1000
End Function

Function DOS_ALATI(M_ALATI)
Dim APOBARO, Apoklish, DIADROMH_BOLAN, ARXH, TIME_ANOXH, Z, ok, TaxyThtes
Dim counter, FIRST_TIME, Motori_On
Z = 0
Motori_On = 0
FIRST_TIME = True
counter = 0   'ΔΕΣ ΤΟ   VAR88     !!!!!!!!
If M_ALATI > 1000 Then
   Apoklish = M_ALATI * 5 / 100
Else
   Apoklish = 30
End If
DIADROMH_BOLAN = var93 - var88 ' 10000 BHMATA
TIME_ANOXH = 300000   '50SEC
'MSComm1.Output = Balance_3_Digits + Chr(13)

Set MyDB = Workspaces(0).OpenDatabase("c:\talos\coord_tb.MDB")
Set r = MyDB.OpenRecordset("Ταχυτητες")
Do Until r.EOF
     If r("id_tax") = 7 Then
        v_Z = r("Ταχυτητα")
     End If
     r.MoveNext
  Loop
r.Close
  



TaxyThtes = "!V,," + Str(v_Z) + ",0.5:"
    
    ' ΠΑΝΕ ΣΤΟ ΒΟΛΑΝ ΚΑΙ  ΠΙΑΣΤΟ   94
    'ΞΕΚΙΝΜΑ ΤΟ ΜΟΤΟΡΙ  "SKON_MOTOR_ON"
    'ΚΡΑΤΑ ΑΠΟΒΑΡΟ
    
    '1 .ANEBASE TO BOLAN ΣΙΓΑ-ΣΙΓΑ
    '    ΔΕΣ ΖΥΓΑΡΙΑ
   'ΠΛΗΣΙΑΣΕ ΠΟΛΥ    ΟΧΙ ->1
   'ΚΑΤΕΒΑ ΚΑΤΩ ΣΙΓΑ-ΣΙΓΑ
   'ΕΦΤΑΣΕΣ ΣΤΟΧΟ   ΟΧΙ =>1
   ' ΚΑΤΕΒΑΣΕ ΒΟΛΑΝ ΣΤΗΝ ΘΕΣΗ ΤΟΥ
   'ΑΝΟΙΞΕ ΓΚΡΙΠΠΕΡ
   'ΑΝΩ Ζ
    
    
    
    ' ΠΑΝΕ ΣΤΟ ΒΟΛΑΝ ΚΑΙ  ΠΙΑΣΤΟ   94
    ' dum = wait_to_understand("!var12=94:")
     
        'ΚΡΑΤΑ ΑΠΟΒΑΡΟ
     APOBARO = Zygis0(0)
     
     dum = Wait_Robot(0)


    
ARXH = GetCurrentTime()
  
If Target > 2000 Then
   xam_tax = "!v, , .1:"
   DIADROMH_BOLAN = DIADROMH_BOLAN / 10
Else
   xam_tax = "!v, , .01:"
   DIADROMH_BOLAN = DIADROMH_BOLAN / 20
End If

  RobSend (xam_tax)
   MilSec 500
  RobSend (xam_tax)


    'ΞΕΚΙΝΑ ΤΟ ΜΟΤΟΡΙ  "SKON_MOTOR_ON"
        RobSend ("!" + skon_motor_on + ":")
Motori_On = 1



ok = 0
' ΤΟ ΤΕΡΜΑ ΤΟΥ ΒΟΛΑΝ  ΕΙΝΑΙ    VAR(88)+12000 bhmata
Do While True
  
  ' ανεβασμα1
  RobSend ("!d , ," + Str(var88 + DIADROMH_BOLAN) + ":!go,,1:")
  Do While axe(3) < Int(var88 + DIADROMH_BOLAN)
       'RobSend ("!d , ," + Str(var88 + DIADROMH_BOLAN) + ":!go,,1:")
       
        'CurQ_Dis = Z
       ' RobSend ("!d , ," + Str(var88 + DIADROMH_BOLAN) + ":!go,,1:")
        Z = Zygis0_time(500) - APOBARO
        If Z > 0 Then CurQ_Dis = Z
        counter = counter + 1
        
        If GetCurrentTime() - ARXH > TIME_ANOXH Then Exit Do
        
        
        
        If Z < -999000000 Then
             RobSend ("!" + skon_motor_off + ":!k:")
              Do While Z < -999000000
                    MilSec 500
                   Z = Zygis0_time(1000) - APOBARO
              Loop
              RobSend ("!" + skon_motor_on + ":!go,,1:")
        End If
          If Z >= M_ALATI - Apoklish * 10 And FIRST_TIME = True Then
            DIADROMH_BOLAN = DIADROMH_BOLAN / 2
            'RobSend ("!" + skon_motor_off + ":!k:")
            FIRST_TIME = False

            Exit Do
        End If
        If Z >= M_ALATI - Apoklish Then
                DOS_ALATI = Z
                ok = 1
                Exit Do
        End If
        
  Loop
  
            If Target - Z > 2000 Then
                xam_tax = "!v, , .1:"
            Else
                xam_tax = "!v, , .01:"
            End If

            RobSend (xam_tax)
            MilSec 500
            RobSend (xam_tax)
  
  
  
    If ok = 1 Then Exit Do
  ' κατεβασμα
  RobSend ("!d , ," + Str(var88) + ":!go,,1:")
  Do While Abs(axe(3) - var88) > 10   ' axe(3)>var88
          Z = Zygis0_time(500) - APOBARO
          counter = counter + 1
        '
          If Z > 0 Then CurQ_Dis = Z
          If Z < -999000000 Then
               RobSend ("!" + skon_motor_off + ":!k:")
               Do While Z < -999000000
                MilSec 100
                   Z = Zygis0_time(1000) - APOBARO
              Loop
              RobSend ("!" + skon_motor_on + ":!go,,1:")
        End If

        If GetCurrentTime() - ARXH > TIME_ANOXH Then Exit Do
        
        If Z >= M_ALATI - Apoklish * 10 And FIRST_TIME = True Then
            DIADROMH_BOLAN = DIADROMH_BOLAN / 2
            FIRST_TIME = False
             RobSend ("!v, ,0.1:")
            MilSec 300
            RobSend ("!v, ,0.1:")
        End If



        If Z >= M_ALATI - Apoklish Then
                DOS_ALATI = Z
                ok = 1
                Exit Do
        End If
  Loop
     
     If GetCurrentTime() - ARXH > TIME_ANOXH Then Exit Do
  
     If ok = 1 Then Exit Do
  
     If Z < Target / 2 Then
          DIADROMH_BOLAN = DIADROMH_BOLAN + DIADROMH_BOLAN
          
          If DIADROMH_BOLAN > var93 - var88 Then
              DIADROMH_BOLAN = var93 - var88
          End If
     End If
     If Motori_On = 0 Then RobSend ("!" + skon_motor_on + ":!k:")

            If Target - Z > 2000 Then
                xam_tax = "!v, , .1:"
            Else
                xam_tax = "!v, , .01:"
            End If

            RobSend (xam_tax)
            MilSec 500
            RobSend (xam_tax)





  Loop
   
RobSend ("!" + skon_motor_off + ":!k:")

 ' RobSend ("!v, , .4:")
 ' MilSec 300
 'RobSend ("!v, , .4:")
  RobSend (TaxyThtes)
    MilSec 500
  RobSend (TaxyThtes)
  


RobSend ("!" + skon_motor_on + ":!k:")

MilSec 100
' kleisto to mpoykali
  RobSend ("!d , ," + Str(var88 - Skacimo) + ":go,,1:")
MilSec 2000
 RobSend ("!" + skon_motor_off + ":!k:")
  
  
  DOS_ALATI = Zygis0_time(1000) - APOBARO
  CurQ_Dis = DOS_ALATI

 
 'MSComm1.Output = Balance_2_Digits + Chr(13)
MilSec 100

End Function
Private Sub anix_zyg()

 On Error Resume Next
 MSComm1.CommPort = Balance_Port
 MSComm1.Settings = Balance_Settings
 MSComm1.InputLen = 0
 MSComm1.PortOpen = True
End Sub

Private Sub Fere_Pot()
 Rob_FeedBack (59)           ' Eλεγχος Robot στα 30 sec
 
 cmd$ = "!var12=" + Str(Pot + 19000) + ":" + "!var24=" + Trim(Str(x_pot)) + ":" + "!var77=" + Trim(Str(y_pot)) + ":" + "!var89=" + Trim(Str(z_pot)) + ":"
 RobSend (cmd$)
 
DUM_DUM = Wait_Robot(0)
End Sub
Private Sub Βαλε_Νερο1_Click()
      
 If system_ready = 1 Then rezero_Click
 If system_ready = 1 Then Rob_FeedBack (59)          'Eλεγχος Robot στα 59 sec
        cmd$ = nero1_on
        RobSend (cmd$)
 If Val(vim) = 0 Then
    vim = Val(var67) + upm4 * Val(Water_Bypass)      'Μετατόπιση 4ου άξονα 20-30  χιλ δεξιά για να μην στάζει
    cmd$ = "!d,,," + Str(vim) + ":!go,,,1:"
    RobSend (cmd$)
 End If
 vim = 0
If system_ready = 1 Then nero1_count_tim = Now()
Do
   nero1_count = nero1_count + 1
     nero1 = ZYGIS3() + nero1_ef
     nerot = nero1
  If (Now() - nero1_count_tim) * 100000 > 120 Then Exit Do
  Loop Until Val(nero1) > Val(nero_dial) * 1000
        
        cmd$ = nero1_off: RobSend (cmd$)
        secwait (6)
     nero1 = nero1_ef + ZYGIS3()
If Val(nero1) > Val(nero_dial) * 1000 + 20000 Then
     nerot = -1: nero1 = -1: curq = -3: CurQ_Dis = curq
     Exit Sub
End If

      nerot = nero1
          secwait (2)
     nero_try = ZYGIS3()    'ελεγχος αν εκλεισε η βαλβίδα νερου
     If nero_try > nero1 + 20000 Then
        nerot = -1: nero1 = -1: curq = -4: CurQ_Dis = curq
  End If
 Label14.Visible = True
       If system_ready = 0 Then
         nero1 = Format(nero_dial, "###,###")
         nerot = nero1
      End If
       skones.Refresh
End Sub

Private Sub Βαλε_Νερο2_Click()
'parameters  Curq,periek,nero1,
Dim Nero_chek, Fdr As Integer, Asw, Nero2_Stox, Ner_Stox
Dim Nero2_Count, Start, ner_yp, Ner, Nero_chek
'objects   nero1,Nero2,nerot, fin_con,fin_con1,nero2_on,nero2_off
'global   system_ready,cmd$,DemoVersion
        
  On Error Resume Next
        
        
        Fdr = 0
        Rob_FeedBack (59)
7
             If system_ready = 1 Then rezero_Click
             secwait (2)
             Asw = ZYGIS3()
             If Asw > 10 Or Asw < -10 Then GoTo 7
            Nero2.Text = 0:
            'curq = 212: nero1 = 112000 'για δοκιμες
            secwait (1)
            Ner_Stox = curq / periek - curq
            Nero2_Stox = Ner_Stox - 3000 - Val(nero1)
            
                                '1o γέμισμα
            cmd$ = nero2_on: RobSend (cmd$)
             Nero2_Count = 0
             Start = GetCurrentTime()
         Do
               Nero2_Count = Nero2_Count + 1
            If system_ready = 1 Then Nero2.Text = ZYGIS3()
            nerot.Text = Val(nero1.Text) + Val(Nero2.Text)
            
            If DemoVersion = 1 And nerot.Text > 150 Then Exit Do
            If GetCurrentTime() - Start > 300000 Then Exit Do

            If nerot.Text > 620000 Then Fdr = 1: Exit Do
            fin_con.Text = curq / (nerot.Text + curq): fin_con1.Text = Format(fin_con.Text, "0.0000 %")
       Loop Until nerot.Text >= Ner_Stox - 7000
             cmd$ = nero2_off: RobSend (cmd$)
             secwait (3)
           If system_ready = 1 Then
                 secwait (3)
                 Nero2 = ZYGIS3()
          End If
          
            nerot.Text = Val(nero1.Text) + Val(Nero2.Text)
            fin_con.Text = curq / (nerot.Text + curq): fin_con1.Text = Format(fin_con.Text, "0.0000 %")
            
                                         '    2ο γεμισμα
                                         
            If Fdr = 0 Then
                    ner_yp = Ner_Stox - nerot.Text
                    Nero2_Count = 70 / 100 * Nero2_Count * ner_yp / Nero2
           End If
      If ner_yp > 20 Then
        cmd$ = nero2_on: RobSend (cmd$)
        Ner = 0
         Do
         Ner = Ner + 1
             If system_ready = 1 Then Nero2 = ZYGIS3()
            nerot.Text = Val(nero1.Text) + Val(Nero2.Text)
            If nerot.Text > 620000 Then Exit Do
            If GetCurrentTime() - Start > 300000 Then Exit Do
            If DemoVersion = 1 And nerot.Text > 150 Then Exit Do
            
            fin_con.Text = curq / (nerot.Text + curq): fin_con1.Text = Format(fin_con.Text, "0.0000 %")
          Loop Until Ner >= Nero2_Count
            cmd$ = nero2_off: RobSend (cmd$)
    End If
          Label18.Visible = True
20
          secwait (2)
      If system_ready = 1 Then Nero2 = ZYGIS3()
        secwait (2)
       If system_ready = 1 Then Nero_chek = ZYGIS3()
                If system_ready = 1 And Abs(Val(Nero2) - Val(Nero_chek)) > 1000 Then GoTo 20
     nerot.Text = Val(nero1.Text) + Val(Nero2.Text)
     fin_con.Text = curq / (nerot.Text + curq): fin_con1.Text = Format(fin_con.Text, "0.0000 %")
         skones.Refresh
End Sub
Private Sub Backmix_Click()
Dim bar_chek
 rezero_Click
 Rob_FeedBack (59)           ' Eλεγχος Robot στα 30 sec
 cmd$ = "!var12=101:"
 RobSend (cmd$)
 'Rob_FeedBack (59)
 DUM_DUM = Wait_Robot(0)
 Label17.Visible = True
 bar = ZYGIS3()
 If bar < 100000 Then
        MsgBox mL_Res(682), , "Talos" '"The robot did not take the bottle from the mixer to the scale !"
        End
End If
 If bar > Val(baros_pot) + curq Then nero1 = bar - Val(baros_pot) - curq 'ξαναμέτρα το νερο λογω εξάτμισης

        
       nerot = nero1
       fin_con = curq / (nerot + curq): fin_con1 = Format(fin_con, "0.0000 %")
  If Val(vim) = 0 Then
    vim = Val(var67) + upm4 * Val(Water_Bypass) 'Μετατόπιση 4ου άξονα 30 χιλ δεξιά για να μην στάζει
    cmd$ = "!d,,," + Str(vim) + ":!go,,,1:"
    RobSend (cmd$)
 End If
 vim = 0
 skones.Refresh
 
End Sub
Private Sub Vale_Nero_Pre()
'1
   ' cmd$ = nero1_on
   ' RobSend (cmd$)
   '  MilSec 5000
   ' cmd$ = nero1_off
   ' RobSend (cmd$)
    '  MilSec 3000
'If system_ready = 1 Then nero1_ef = ZYGIS3(): nero1 = nero1_ef
'If nero1_ef < 500 Then GoTo 1
End Sub
Private Sub Beready_Click()
  Rob_FeedBack (59)           ' Eλεγχος Robot στα 30 sec
  cmd$ = "!var12=" + Str(Pot + 19000) + ":" + "!var24=" + Trim(Str(x_pot)) + ":" + "!var77=" + Trim(Str(y_pot)) + ":" + "!var89=" + Trim(Str(z_pot)) + ":"
RobSend (cmd$)
' Rob_FeedBack (59)
DUM_DUM = Wait_Robot(0)
End Sub
Function file_send()
Dim MyDB As Database, r As Recordset
If system_ready = 0 Then Exit Function
 Set MyDB = Workspaces(0).OpenDatabase("c:\talos\skon_tb.MDB")
Set r = MyDB.OpenRecordset("Διαλυματα")
    r.AddNew
 r("nerostox") = nerostox
     r("periek") = periek
    r("buk") = buk
     r("pot") = Pot
     r("mixtim") = s_mixtim
      r("mixtim_fin") = mixtim_fin
     r("thermo1") = thermo1
     r("thermo2") = thermo2
     r("nero_act") = nerot
     r("color_act") = curq
     r("duration") = Διαρκεια
     r("datum") = Now
     r("druv") = Druv
     r("nero_dial") = nero_dial
     r("ximiko") = ximiko
     r.Update
   r.Close
  
  Set r = MyDB.OpenRecordset("Εντολες")
Do Until r.EOF
entol = 0
 If IsNull(r("nero_act")) Then
    r.Edit
     r("nero_act") = nerot
     r("color_act") = curq
     r("duration") = Διαρκεια
     r("datum") = Now
     r.Update
    entol = 0
    Exit Do
 End If
  r.MoveNext
  Loop
  r.Close
End Function

Function secwait(tim As Long) 'Περιμενε σε sec
If system_ready = 0 Then secwait = tim: Exit Function
altim = 0
reftim = Second(Time): reftim1 = Time
Do
curtim1 = Second(Time)
    Do
     DoEvents
     curtim2 = Second(Time)
     time1 = Time
   Loop Until curtim2 <> curtim1
altim = altim + 1
timdif = altim - reftim
Loop Until altim >= tim
End Function

Sub Rob_FeedBack(arg As Integer) 'Εξεταζει την αποκριση του Robot
'secwait (1)
While FindOutput(Val(Ready_Bit)) = 0
     curtim = Second(Time)
     If curtim <> reftim Then
      sec = sec + 1: reftim = curtim
    If sec > arg Then
     tim = Time
     tit = "T A L O S"
     '   670    "Ωρα "
     '   671   ". Το Robot δεν ανταποκρίνεται . Η διαδικασία τερματίζεται."
     Msg = mL_Res(670) + Str(tim) + mL_Res(671)
     'Msg = "Ωρα " + Str(tim) + ". Το Robot δεν ανταποκρίνεται . Η διαδικασία τερματίζεται."
     MsgBox Msg, , "670-1.Talos"
     Telos_Dialymaton = 1
     Exit Sub
    End If
      DoEvents
     End If
  Wend
 
End Sub


Function check_system(afr)
rezero_Click
Asw = ZYGIS3()
If Asw = " " Then
  check_system = 1
  Exit Function
End If
rezero_Click
End Function

Private Sub Command3_Click()
'-------------- Δυό μαζί διαλύματα-------------------

Screen.MousePointer = 11
anix_zyg

ValveReset
Telos_Dialymaton = 0
met = 0
exodos.Enabled = False
Command1.Enabled = False
If system_ready = 1 Then
If system_ready = 1 Then rezero_Click
MilSec 200
Do
    'On Error Resume Next
     MSComm1.Output = "HI RANGE" + Chr$(13)
     secwait (1)
     met = met + 1
     dec = zyg_decimals()
        If met = 10 Then
                Exit Do
        End If
Loop Until dec = 2

End If
If dec <> 2 Then
MsgBox mL_Res(724), , "Talos" '"Μετά από 5 προσπάθειες, η ζυγαριά δεν ανταποκρίνεται.  Ανοίξτε την ζυγαριά ,  πατείστε το πλήκτρο  F  για να την ρυθμίσετε σε ζύγιση 2 δεκαδικών και ξαναδοκιμάστε."
                 exodos.Enabled = True
                Command1.Enabled = True
                Screen.MousePointer = 0
                Exit Sub
End If
Data1.Refresh
DBGrid1.Refresh
skones.Refresh
 Label25.Visible = False
 Label25 = mL_Res(674) ' 674 "Εκτέλεση Τελευταίας Παρασκευής"

If talostart() = 0 Then
 ' "Το Πρόγραμμα Κίνησης του Robot δεν μπορεί να τρέξει.Η διαδικασία τερματίζεται..."
 MsgBox mL_Res(662), , "662.Talos" ' "Το Πρόγραμμα Κίνησης του Robot δεν μπορεί να τρέξει.Η διαδικασία τερματίζεται..."
  talostart = 0
  Telos_Dialymaton = 1: Exit Sub
  fereSK.Show
End If

'   663   "Τα μπουκάλια που πρόκειται να χρησιμοποιηθούν στην παρασκευή διαλυμάτων πρέπει : "
'   664   "1. Να είναι  καθαρά και στεγνά."
'   665   "2. Να έχουν αφαιρεθεί  τα καπάκια τους."
'  666    "3. Κάθε μπουκάλι να είναι εφοδιασμένο με τον μαγνήτη του."
'  667    "4. Σε όσα μπουκάλια προβλέπεται,  να έχουν προστεθεί τα απαραίτητα χημικά ."
'  668    "     Να ξεκινήσει το Robot   την παρασκευή Διαλυμάτων ;"
Ans = MsgBox(mL_Res(663) + Chr(13) + Chr(13) + mL_Res(664) + Chr(13) + mL_Res(665) + Chr(13) + mL_Res(666) + Chr(13) + mL_Res(667) + Chr(13) + Chr(13) + Chr(13) + mL_Res(668), 20, "663-8.Talos")
'Ans = MsgBox("Τα μπουκάλια που πρόκειται να χρησιμοποιηθούν στην παρασκευή διαλυμάτων πρέπει : " + Chr(13) + "1. Να είναι  καθαρά και στεγνά." + Chr(13) + "2. Να έχουν αφαιρεθεί  τα καπάκια τους." + Chr(13) + "3. Κάθε μπουκάλι να είναι εφοδιασμένο με τον μαγνήτη του." + Chr(13) + "4. Σε όσα μπουκάλια προβλέπεται,  να έχουν προστεθεί τα απαραίτητα χημικά ." + Chr(13) + Chr(13) + "     Να ξεκινήσει το Robot   την παρασκευή Διαλυμάτων ;", 20, "Talos : Παρασκευή Διαλυμάτων")

If Ans = 7 Then   ' oxi
    Screen.MousePointer = 1
      exodos.Enabled = True
      Command1.Enabled = True
'      669  "Πρέπει να βγείτε από το πρόγραμμα  παρασκευής διαλυμάτων,πατώντας Εξοδος"
    
    MsgBox mL_Res(669), , "669.Talos"
  Exit Sub
End If
 
 
 Command1.Enabled = False

Do
Cur_Maker = Cur_Maker + 1
   If check_system(0) <> 0 Then
       Screen.MousePointer = 0
       '  670   "Παρακαλώ ανοίξτε την Ζυγαριά και πατείστε πάλι την Εναρξη..."
       MsgBox mL_Res(724), , "724.Talos" ' "Παρακαλώ ανοίξτε την Ζυγαριά και πατείστε πάλι την Εναρξη..."
       Command1.Enabled = True
       Exit Sub
   End If
   GoSub Diabase_Epomena   'φορτώνει τις μεταβλητές του διαλύματος
   Αρχη = Now
   Timer1.Enabled = True
   If Telos_Dialymaton = 1 Then Exit Do
   curq = 0: CurQ_Dis = curq
   dyo_start_dosing_click
   If Telos_Dialymaton = 1 Then Exit Do
Loop Until Cur_Maker = sol_maker

Label25 = mL_Res(725) ' "Διακοπή Προγράμματος Παρασκευής Διαλυμάτων"
     
         ' ΧΤΥΠΗΜΑ ΚΟΥΔΟΥΝΙΟΥ
 cmd$ = "!" + Alarm1_off + ":"
RobSend (cmd$)

 cmd$ = "!" + Alarm1_on + ":"
RobSend (cmd$)

Data1.Refresh
DBGrid1.Refresh
 Command1.Enabled = True
 Command2.Enabled = False
  Timer1.Enabled = False
 exodos.Enabled = True
 MSComm1.PortOpen = False
 
 '676      "Ολες οι εντολές για παρασκευή διαλυμάτων έχουν ολοκληρωθεί . Το πρόγραμμα τερματίζεται..."
  MsgBox mL_Res(726), , "726.Talos"  ' "Ολες οι εντολές για παρασκευή διαλυμάτων έχουν ολοκληρωθεί . Το πρόγραμμα τερματίζεται..."

Screen.MousePointer = 1



Exit Sub


Diabase_Epomena:
'----------------------------
Set MyDB = Workspaces(0).OpenDatabase("c:\talos\skon_tb.mdb")
Set r = MyDB.OpenRecordset("Εντολες")
Do Until r.EOF
entol = 0
 If IsNull(r("nero_act")) Then                     ' για να εκτελεί μόνο τα διαλύματα που δεν έχουν ήδη εκτελεστεί
    nerostox1 = r("nerostox")
    periek1 = r("periek")
    periek_dis1 = Format(periek, "0.##0%")
    buk1 = r("buk")
    Pot1 = r("pot")
    mixtim1 = r("mixtim")
    s_mixtim1 = r("mixtim")
    mixtim_fin1 = r("mixtim_fin")
    thermo1_1 = r("thermo1")
    thermo2_1 = r("thermo2")
    nero_dial_1 = r("nero_dial")
    Label14.Caption = nero_dial
    If r("ximiko") <> "" Then ximiko1 = r("ximiko") Else ximiko1 = "Μπουκάλα Δοκιμαστική"
    entol = 1
      
     r.MoveNext
    If r.EOF Then
       buk2 = 0
    Else
       nerostox2 = r("nerostox")
       periek2 = r("periek")
       periek_dis2 = Format(periek, "0.##0%")
       buk2 = r("buk")
       Pot2 = r("pot")
       mixtim2 = r("mixtim")
       s_mixtim2 = r("mixtim")
       mixtim_fin2 = r("mixtim_fin")
       thermo1_2 = r("thermo1")
       thermo2_2 = r("thermo2")
       nero_dial_2 = r("nero_dial")
       If r("ximiko") <> "" Then ximiko2 = r("ximiko") Else ximiko2 = "Μπουκάλα Δοκιμαστική"
    End If
      
      
      
    Exit Do
   End If
   
   
   
   
  r.MoveNext
  Loop
  
  r.Close


End Sub

Private Sub exodos_Click()
ZAP_MOVES '12/99
On Error Resume Next
  MSComm1.PortOpen = False
  Printer.EndDoc
    Unload Me
   ' fereSK.Show
End Sub

Private Sub Find_Target()

Dim posot_Θεωρ As Single, pos As Integer, Buttom_step As Integer
'pos ο στόχος σκόνης σε mgr
pos = Target - curq '+ Tare   ' η ποσότητα που θα ρίξει σε λίγο(πόσο απέχει από τον στόχο)
If pos < 0 Then
  z_met = 1 ': posot_Θεωρ = z_met * z_met * a_factor / 1000
  Exit Sub
End If

If Ca_Factor = 0 Then
  Ca_Factor = 10
  Ca_Factor_dis = Ca_Factor
  Proti_Fora = True
End If

z_met = Sqr(pos / Ca_Factor) 'μετατόπιση ντρουβανίσματος


GoSub Select_Safety_Factor

If z_met > 100 Then z_met = 100
    z_met = z_met * Safety_Factor 'παράγων ασφάλειας
GoSub Find_Buttom
If Last_z_met <= Buttom_step And Last_Dosage < pos / 3 Then
      Anevasma = Anevasma + 3 ' 5
      Buttom_step = Buttom_step + Anevasma
End If

If z_met < Buttom_step Then z_met = Buttom_step
     ' upm = 425 βήματα / mm
z_syntetag = upm * z_met * 40 / 100 'μετατρέπω το z_met απο χιλιοστά σε βήματα

'αν η μετατόπιση υπερβαίνει το άνω Ζ (var93) τότε μετατόπιση = ανωΖ(var93) - Ζβολάν(var88)
If (Val(z_syntetag) + Val(var88)) > Val(var93) Then
      z_syntetag = Val(var93) - Val(var88)
End If

'υπολογίζω το z_met σε χιλιοστά
z_met = z_syntetag / upm / 40 * 100

Exit Sub

Select_Safety_Factor:
Select Case Val(pos)
Case Is > Q_Max_Quantity
Safety_Factor = 1
Case Is > 0.9 * Q_Max_Quantity
Safety_Factor = 0.95
Case Is > 0.8 * Q_Max_Quantity
Safety_Factor = 0.9
Case Is > 0.7 * Q_Max_Quantity
Safety_Factor = 0.85
Case Is > 0.6 * Q_Max_Quantity
Safety_Factor = 0.8
Case Else
Safety_Factor = 0.9
End Select
If Val(pos) < 400 And Q_Max_Quantity = 0 Then Safety_Factor = 0.75
Return

Find_Buttom:
Select Case Ca_Factor


Case Is > 1.5
Buttom_step = 5 + Anevasma
Case Is > 0.5
Buttom_step = 10 + Anevasma
Case Is > 0.1
Buttom_step = 15 + Anevasma
Case Is > 0.01
Buttom_step = 20 + Anevasma
Case Else
Buttom_step = 25 + Anevasma
End Select
Return
End Sub

Private Sub Finish_Click()
If mixtim_fin > 0 Then
   Sendmix_Click
   Mixing.Caption = mL_Res(683) ' "Τελική Ανάδευση"
   Label29.Caption = mL_Res(684)   '"Final Mixing  "
   Label16.Caption = "0"
   mixtim = mixtim_fin
      cmd$ = "!" + mixer_on + ":"
       RobSend (cmd$)
       mixing_make (mixtim_fin)
       cmd$ = "!" + mixer_off + ":"
       RobSend (cmd$)
End If
Rob_FeedBack (59)
cmd$ = "!var12=" + Str(Pot + 200) + ":"
RobSend (cmd$)
' Rob_FeedBack (59)
DUM_DUM = Wait_Robot(0)
 Label19.Visible = True
Ληξη = Now
Διαρκεια = DateDiff("s", Αρχη, Ληξη)
file_send
skones.Refresh
  Screen.MousePointer = 0
End Sub



Private Sub Form_Load()

'On Error GoTo er_chek
Dim MyDB As Database, r As Recordset, MySet As Recordset
Ca_Factor = 0: Anevasma = 0: Last_Dosage = 0: Proti_Fora = False

Label1(0).Caption = mL_Res(760) ' "Στόχος Σκόνης"
Label3.Caption = mL_Res(761) '"Σκόνη Δοσομετρημένη"
Label39.Caption = mL_Res(762) '"Δόσεις"

Label7.Caption = mL_Res(763) ' "Μπουκάλι Σκόνης"
Label6.Caption = mL_Res(764) '"Ποτήρι Διάλυσης"
Label40.Caption = mL_Res(765) '"Ρευστότητα Σκόνης"
Label26.Caption = mL_Res(766) '"Εναρξη"
Label27.Caption = mL_Res(767) ' "Λήξη"
Label28.Caption = mL_Res(768) ' "(sec)"
Label9.Caption = mL_Res(769) '"Τάρα"
Label38.Caption = mL_Res(770) '"Περιεκτικότητες ( w / w ) "
Label23.Caption = mL_Res(771) '"Στόχος"
Label37.Caption = mL_Res(772) '"Τελική"
Label2.Caption = mL_Res(773) '"Νερό Διάλυσης"
Label29.Caption = mL_Res(774) ' "Mix.Time"
Label25.Caption = mL_Res(775) ' "Εκτέλεση Τελευταίας Παρασκευής"
Label20.Caption = mL_Res(776) ' "ΔΙΑΔΙΚΑΣΙΕΣ"
Label35.Caption = mL_Res(777) '"Προετοιμασία"
Label36.Caption = mL_Res(778) '"Προσθήκη Σκόνης"
Label34.Caption = mL_Res(779) ' "Προσθ.Νερού Διάλυσης"
Label33.Caption = mL_Res(780) ' "Προς Ανάδευση"
Label32.Caption = mL_Res(781) '"Ανάδευση"
Label30.Caption = mL_Res(782) '"Επιστροφή Ζυγαριά"
Label10.Caption = mL_Res(783) '"Πρ.Νερού Αραίωσης"
Label5.Caption = mL_Res(784) '"Τελ. Επιστροφή"
Label24.Caption = mL_Res(599) ' "Συνολο Νερου (mgr)"       ' Συνολο Νερου (mgr)
Command1.Caption = mL_Res(785) '"Εναρξη Παρασκευής"
Command2.Caption = mL_Res(786) '"Διακοπή  Προγράμματος"
exodos.Caption = mL_Res(787) '"Εξοδος"
DBGrid1.Columns(1).Caption = mL_Res(788) ' "Ποσότης"
DBGrid1.Columns(2).Caption = mL_Res(789) '"Στόχ.%"
DBGrid1.Columns(3).Caption = mL_Res(790) ' "Μπουκ."
DBGrid1.Columns(5).Caption = mL_Res(791) ' "Mixing"
DBGrid1.Columns(6).Caption = mL_Res(792) '"Fin.Mixing"
DBGrid1.Columns(7).Caption = mL_Res(793) '"Νερ.Διάλ."
DBGrid1.Columns(8).Caption = mL_Res(794) '"Νερ.Αραίωσ."
DBGrid1.Columns(9).Caption = mL_Res(795) ' "Πραγ.Νερό"
DBGrid1.Columns(10).Caption = mL_Res(796) ' "Τελ.χρώμα"
DBGrid1.Columns(11).Caption = mL_Res(797) '"Ημερ/νία"
DBGrid1.Columns(12).Caption = mL_Res(798) ' "Διάρκεια"
DBGrid1.Columns(13).Caption = mL_Res(799) '"Ποσ.Νερ.Διάλ."

DBGrid1.Columns(14).Caption = mL_Res(650) '"Περιγραφή Χημικού"
Beready.Caption = mL_Res(651) '"Προετοιμασία"
Dosing.Caption = mL_Res(652) '"Δοσομέτρηση"
Βαλε_Νερο1.Caption = mL_Res(653) '"Νερό Διάλυσης"
Sendmix.Caption = mL_Res(654) ' "Προς Ανάδευση"
Mixing.Caption = mL_Res(655) '"Ανάδευση"

Backmix.Caption = mL_Res(656) ' "Επιστροφή Ζυγαριά"
Βαλε_Νερο2.Caption = mL_Res(657) ' "Νερό Αραίωσης"
finish.Caption = mL_Res(658) '"Επιστροφή"










FindValves







system_ready = IsOSLoaded(768)
Q_Max_Quantity = 0
Last_z_met = 0

If Val(Pot) > 0 Then GoTo jump1
 'Σβήσιμο παλιων παρασκευών
Set MyDB = Workspaces(0).OpenDatabase("c:\talos\skon_tb.mdb")
Set MySet = MyDB.OpenRecordset("Εντολες", dbOpenDynaset)
sql_del = "DELETE DISTINCTROW Εντολες.*, Εντολες.nero_act FROM Εντολες WHERE ((Not Εντολες.nero_act Is Null));"
MyDB.Execute sql_del
 Timer1.Enabled = False
jump1:


Mixing.Caption = mL_Res(685)      '   "Ανάδευση"
Label29.Caption = "Mixing Time  " '  "Mixing Time  "

Ληξη = ""
Διαρκεια = ""
Druv = 0: Druv_dis = "0"
nero1 = 0
a_factor = 10
Label8.Visible = False
Label13.Visible = False
Label13.Visible = False
Label15.Visible = False
Label16.Visible = False
Label17.Visible = False
Label18.Visible = False
Label19.Visible = False
Nero2 = 0: fin_con = 0: nerot = 0: fin_con1 = "0 %"
Data1.Refresh
DBGrid1.Refresh
skones.Refresh

Set MyDB = Workspaces(0).OpenDatabase("c:\talos\skon_tb.mdb")
Set r = MyDB.OpenRecordset("Εντολες")
Do Until r.EOF
entol = 0
 If IsNull(r("nero_act")) Then                     ' για να εκτελεί μόνο τα διαλύματα που δεν έχουν ήδη εκτελεστεί
    nerostox = r("nerostox")
    periek = r("periek")
    periek_dis = Format(periek, "0.##0%")
    buk = r("buk")
    Pot = r("pot")
    mixtim = r("mixtim")
    s_mixtim = r("mixtim")
    mixtim_fin = r("mixtim_fin")
    thermo1 = r("thermo1")
    thermo2 = r("thermo2")
    nero_dial = r("nero_dial")
    Label14.Caption = nero_dial
   If r("ximiko") <> "" Then ximiko = r("ximiko") Else ximiko = "Μπουκάλα Δοκιμαστική"
    entol = 1
    Exit Do
   End If
  r.MoveNext
  Loop
  
  r.Close
  
  If entol = 0 Then
 
   
    Telos_Dialymaton = 1
    Exit Sub
 End If
 
 If check_value(nerostox) = 0 Then MsgBox mL_Res(731), , "Talos : Παρασκευή Διαλυμάτων.":    Telos_Dialymaton = 1: Exit Sub '' "Δεν υπάρχει  ποσότητα Διαλύματος.Η διαδικασία τερματίζεται..."
 If check_value(periek) = 0 Then MsgBox mL_Res(732), , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub   ' ' "Δεν υπάρχει  περιεκτικότητα Διαλύματος.Η διαδικασία τερματίζεται..."

 If check_value(buk) = 0 Then MsgBox mL_Res(687): Telos_Dialymaton = 1: Exit Sub '"Δεν υπάρχει  Μπουκάλι  Διαλύματος.Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων."
 If check_value(Pot) = 0 Then MsgBox mL_Res(688), , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub '"Δεν υπάρχει  ποτήρι  Διαλύματος.Η διαδικασία τερματίζεται..."
 If check_value(mixtim) = 0 Then MsgBox mL_Res(689), , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub '  "Δεν υπάρχει  Χρόνος  Διάλυσης.Η διαδικασία τερματίζεται..."
  If check_value(mixtim_fin) = 0 Then MsgBox mL_Res(690), , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub '  "Δεν υπάρχει  Τελικός Χρόνος  Διάλυσης.Η διαδικασία τερματίζεται..."
 If check_value(thermo1) = 0 Then MsgBox mL_Res(691), , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub ' "Δεν υπάρχει  Θερμοκρασία νερού  Διάλυσης.Η διαδικασία τερματίζεται..."
 If check_value(thermo2) = 0 Then MsgBox mL_Res(692), , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub '"Δεν υπάρχει  Θερμοκρασία νερού  Αραίωσης.Η διαδικασία τερματίζεται..."
 If check_value(nero_dial) = 0 Then MsgBox mL_Res(693), , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub    ' "Δεν δηλώθηκε η ποσότητα του νερού Διάλυσης.Η διαδικασία τερματίζεται..."
 If Val(nero_dial) < 20 Then MsgBox mL_Res(693), , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub '           "Δεν δηλώθηκε η ποσότητα του νερού Διάλυσης.Η διαδικασία τερματίζεται..."
 
Set r = MyDB.OpenRecordset("Σκονες")
Do Until r.EOF
 If buk = r("Νο_Μπουκαλιου") Then
    Ca_Factor = Val(r("Factor_A"))
    Ca_Factor_dis = Ca_Factor
    Exit Do
End If
  r.MoveNext
  Loop
  r.Close
If Factor_a = "" Then Factor_a = 2


Set MyDB = Workspaces(0).OpenDatabase("c:\talos\coord_tb.MDB")
Set r = MyDB.OpenRecordset("Ταχυτητες")
Do Until r.EOF
 If r("id_tax") = 3 Then
    upm = r("Μοναδες_μμ")
End If
 If r("id_tax") = 4 Then
    upm4 = r("Μοναδες_μμ")
End If
  r.MoveNext
  Loop
  r.Close
  
 If check_value(upm) = 0 Then MsgBox mL_Res(694), , "Talos 694": Telos_Dialymaton = 1: Exit Sub
 If check_value(upm4) = 0 Then MsgBox mL_Res(695), , "Talos 695": Telos_Dialymaton = 1: Exit Sub
     
     x_pot = "": y_pot = "": z_pot = "": x_buk = "": y_buk = "": z_buk = ""

  Set r = MyDB.OpenRecordset("Συντεταγμενες")
   Do Until r.EOF
           If r("ΑΑ_Item") = 2 Then
            If r("Νο_Τεμαχιου") = Val(Pot) Then
              x_pot = r("Συντεταγμένη_Χ_user")
              y_pot = r("Συντεταγμένη_Υ_user")
              z_pot = r("Συντεταγμένη_Ζ_user")
            End If
           End If
         If r("ΑΑ_Item") = 1 Then
            If r("Νο_Τεμαχιου") = Val(buk) Then
              x_buk = r("Συντεταγμένη_Χ_user")
              y_buk = r("Συντεταγμένη_Υ_user")
              z_buk = r("Συντεταγμένη_Ζ_user")
            End If
           End If
           r.MoveNext
          Loop
  r.Close
  
 If check_value(x_pot) = 0 Then MsgBox mL_Res(696), , "Talos": Telos_Dialymaton = 1: Exit Sub ' "Δεν υπάρχουν  οι Συντεταγμένες X Ποτηριού.Η διαδικασία τερματίζεται..."
 If check_value(y_pot) = 0 Then MsgBox mL_Res(697), , "Talos": Telos_Dialymaton = 1: Exit Sub '+ "Δεν υπάρχουν  οι Συντεταγμένες Υ Ποτηριού.Η διαδικασία τερματίζεται..."
 If check_value(z_pot) = 0 Then MsgBox mL_Res(698), , "Talos": Telos_Dialymaton = 1: Exit Sub '+ "Δεν υπάρχουν  οι Συντεταγμένες Ζ Ποτηριού.Η διαδικασία τερματίζεται..."
 If check_value(x_buk) = 0 Then MsgBox mL_Res(699), , "Talos": Telos_Dialymaton = 1: Exit Sub ' + "Δεν υπάρχουν  οι Συντεταγμένες Χ Μπουκαλιού.Η διαδικασία τερματίζεται..."
 If check_value(y_buk) = 0 Then MsgBox mL_Res(700), , "Talos": Telos_Dialymaton = 1: Exit Sub '+ "Δεν υπάρχουν  οι Συντεταγμένες Υ Μπουκαλιού.Η διαδικασία τερματίζεται..."
 If check_value(z_buk) = 0 Then MsgBox mL_Res(701), , "Talos": Telos_Dialymaton = 1: Exit Sub ' + "Δεν υπάρχουν  οι Συντεταγμένες Ζ Μπουκαλιού.Η διαδικασία τερματίζεται...
  
  
Set MyDB = Workspaces(0).OpenDatabase("c:\talos\coord_tb.MDB")
Set r = MyDB.OpenRecordset("var")
Do Until r.EOF
 If r("field_name") = "VAR93" Then
    var93 = r("field_val")
End If
 If r("field_name") = "VAR88" Then
    var88 = r("field_val")
End If
 If r("field_name") = "VAR67" Then
    var67 = r("field_val")
 End If
 
 If r("field_name") = "VAR99" Then
    Skacimo = r("field_val")
 End If

 
 
r.MoveNext
  Loop
  r.Close
  
 If check_value(var93) = 0 Then MsgBox mL_Res(702), , "Talos": Telos_Dialymaton = 1: Exit Sub  ' "Δεν υπάρχει η Παράμετρος var93.Η διαδικασία τερματίζεται..."
 If check_value(var88) = 0 Then MsgBox mL_Res(703), , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub ' "Δεν υπάρχει η Παράμετρος var88.Η διαδικασία τερματίζεται..."
  If check_value(var67) = 0 Then MsgBox mL_Res(704), , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub '"Δεν υπάρχει η Παράμετρος var67.Η διαδικασία τερματίζεται..."
 
  If entol = 0 Then Telos_Dialymaton = 1: Exit Sub
   Target = periek * Val(nerostox): Target_Dis = Str(Target)
   
  Select Case thermo1
  Case 1
     Label21 = mL_Res(705) '"Κρύο"
   Case 2
      Label21 = mL_Res(706) '"Χλιαρό"
   Case 3
      Label21 = mL_Res(707) ' "Ζεστό"
End Select

  Select Case thermo2
  Case 1
    Label22 = mL_Res(705) '"Κρύο"
   Case 2
      Label22 = mL_Res(706) '"Χλιαρό"
   Case 3
      Label22 = mL_Res(707) '"Ζεστό"
End Select

Select Case Val(thermo1)
           Case 1
             nero1_on = "!" + Cold_Water_On + ":"
             nero1_off = "!" + Cold_Water_Off + ":"
         Case 2
            nero1_on = "!" + Cold_Water_On + ":" + "!" + Hot_Water_On + ":"
            nero1_off = "!" + Cold_Water_Off + ":" + "!" + Hot_Water_Off + ":"
          Case 3
            nero1_on = "!" + Hot_Water_On + ":"
             nero1_off = "!" + Hot_Water_Off + ":"
         End Select
        
      Select Case Val(thermo2)
           Case 1
             nero2_on = "!" + Cold_Water_On + ":"
             nero2_off = "!" + Cold_Water_Off + ":"
         Case 2
            nero2_on = "!" + Cold_Water_On + ":" + "!" + Hot_Water_On + ":"
            nero2_off = "!" + Cold_Water_Off + ":" + "!" + Hot_Water_Off + ":"
          Case 3
            nero2_on = "!" + Hot_Water_On + ":"
             nero2_off = "!" + Hot_Water_Off + ":"
         End Select



         
 If check_value(Ready_Bit) = 0 Then MsgBox mL_Res(708), , "Talos": Telos_Dialymaton = 1: Exit Sub ' "Δεν είναι  δηλωμένο το Ok Bit. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός = 2 ). Η διαδικασία τερματίζεται..."
 If check_value(skon_motor_on) = 0 Then MsgBox mL_Res(709), , "Talos": Telos_Dialymaton = 1: Exit Sub '"Δεν είναι  δηλωμένο το output του κινητήρα της σκόνης. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός =4 ). Η διαδικασία τερματίζεται..."
 If check_value(skon_motor_off) = 0 Then MsgBox mL_Res(710), , "Talos": Telos_Dialymaton = 1: Exit Sub '"Δεν είναι  δηλωμένο το output του κινητήρα της σκόνης. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός =4 ). Η διαδικασία τερματίζεται..."
 If check_value(nero1_on) = 0 Then MsgBox mL_Res(711), , "Talos": Telos_Dialymaton = 1: Exit Sub '"Δεν είναι  δηλωμένο το output της βαλβίδας του νερού διάλυσης. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός =6 ). Η διαδικασία τερματίζεται..."
 If check_value(nero1_off) = 0 Then MsgBox mL_Res(712), , "Talos": Telos_Dialymaton = 1: Exit Sub '"Δεν είναι  δηλωμένο το output της βαλβίδας του νερού διάλυσης. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός =6 ). Η διαδικασία τερματίζεται..."
 If check_value(nero2_on) = 0 Then MsgBox mL_Res(713), , "Talos": Telos_Dialymaton = 1: Exit Sub '"Δεν είναι  δηλωμένο το output της βαλβίδας του νερού αραίωσης. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός =5 ). Η διαδικασία τερματίζεται..."
 If check_value(nero2_off) = 0 Then MsgBox mL_Res(714), , "Talos": Telos_Dialymaton = 1: Exit Sub '"Δεν είναι  δηλωμένο το output της βαλβίδας του νερού αραίωσης. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός =5 ).Η διαδικασία τερματίζεται..."
 If check_value(Cold_Water_On) = 0 Then MsgBox mL_Res(715), , "Talos": Telos_Dialymaton = 1: Exit Sub '"Δεν είναι  δηλωμένο το output της βαλβίδας του κρύου νερού. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός =5 ). Η διαδικασία τερματίζεται..."
 If check_value(Cold_Water_Off) = 0 Then MsgBox mL_Res(716), , "Talos": Telos_Dialymaton = 1: Exit Sub ''"Δεν είναι  δηλωμένο το output της βαλβίδας του κρύου νερού. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός =5 ). Η διαδικασία τερματίζεται..."
 If check_value(Hot_Water_On) = 0 Then MsgBox mL_Res(717), , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub '"Δεν είναι  δηλωμένο το output της βαλβίδας του ζεστού νερού. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός =6 ). Η διαδικασία τερματίζεται..."
 If check_value(Hot_Water_Off) = 0 Then MsgBox mL_Res(718), , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub ''"Δεν είναι  δηλωμένο το output της βαλβίδας του ζεστού νερού. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός =6 ). Η διαδικασία τερματίζεται..."
 If check_value(mixer_on) = 0 Then MsgBox mL_Res(719), , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub ''"Δεν είναι  δηλωμένο το output του κινητήρα της ανάδευσης. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός =7).Η διαδικασία τερματίζεται..."
 If check_value(mixer_off) = 0 Then MsgBox mL_Res(720), , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub '"Δεν είναι  δηλωμένο το output του κινητήρα της ανάδευσης. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός =7).Η διαδικασία τερματίζεται..."
 If check_value(Alarm1_on) = 0 Then MsgBox mL_Res(721), , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub '"Δεν είναι  δηλωμένο το outpout του Βομβητή. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός =11).Η διαδικασία τερματίζεται..."
 If check_value(Balance_Port) = 0 Then MsgBox mL_Res(722), , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub '"Δεν είναι  δηλωμένη η πόρτα της ζυγαριάς. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός =8).Η διαδικασία τερματίζεται..."

Me.WindowState = 2
'If system_ready = 1 Then
'anix_zyg

ex_sub:
  Exit Sub

er_chek:
 If Err = 3189 Then MsgBox mL_Res(723), , "Talos : Παρασκευή Διαλυμάτων" '"Το αρχείο Εντολές Σκόνης είναι κατειλημένο από κάποιο άλλο πρόγραμμα. To παρόν πρόγραμμα δεν μπορεί να συνεχίσει..."
          Resume Next

End Sub



Private Sub Command1_Click()
Screen.MousePointer = 11
anix_zyg

ValveReset
Telos_Dialymaton = 0
met = 0
exodos.Enabled = False
Command1.Enabled = False
If system_ready = 1 Then
If system_ready = 1 Then rezero_Click
MilSec 200
Do
    'On Error Resume Next
    MSComm1.Output = "HI RANGE" + Chr$(13)
     secwait (1)
     met = met + 1
     dec = zyg_decimals()
         If met = 10 Then
                Exit Do
        End If
 Loop Until dec = 2
End If
If dec <> 2 Then
MsgBox mL_Res(724), , "Talos" '"Μετά από 5 προσπάθειες, η ζυγαριά δεν ανταποκρίνεται.  Ανοίξτε την ζυγαριά ,  πατείστε το πλήκτρο  F  για να την ρυθμίσετε σε ζύγιση 2 δεκαδικών και ξαναδοκιμάστε."
                 exodos.Enabled = True
                Command1.Enabled = True
                Screen.MousePointer = 0
                Exit Sub
End If
Data1.Refresh
DBGrid1.Refresh
skones.Refresh
 Label25.Visible = False
 Label25 = mL_Res(674) ' 674 "Εκτέλεση Τελευταίας Παρασκευής"

If talostart() = 0 Then
 ' "Το Πρόγραμμα Κίνησης του Robot δεν μπορεί να τρέξει.Η διαδικασία τερματίζεται..."
 MsgBox mL_Res(662), , "662.Talos" ' "Το Πρόγραμμα Κίνησης του Robot δεν μπορεί να τρέξει.Η διαδικασία τερματίζεται..."
  talostart = 0
  Telos_Dialymaton = 1: Exit Sub
  fereSK.Show
End If

'   663   "Τα μπουκάλια που πρόκειται να χρησιμοποιηθούν στην παρασκευή διαλυμάτων πρέπει : "
'   664   "1. Να είναι  καθαρά και στεγνά."
'   665   "2. Να έχουν αφαιρεθεί  τα καπάκια τους."
'  666    "3. Κάθε μπουκάλι να είναι εφοδιασμένο με τον μαγνήτη του."
'  667    "4. Σε όσα μπουκάλια προβλέπεται,  να έχουν προστεθεί τα απαραίτητα χημικά ."
'  668    "     Να ξεκινήσει το Robot   την παρασκευή Διαλυμάτων ;"
Ans = MsgBox(mL_Res(663) + Chr(13) + Chr(13) + mL_Res(664) + Chr(13) + mL_Res(665) + Chr(13) + mL_Res(666) + Chr(13) + mL_Res(667) + Chr(13) + Chr(13) + Chr(13) + mL_Res(668), 20, "663-8.Talos")
'Ans = MsgBox("Τα μπουκάλια που πρόκειται να χρησιμοποιηθούν στην παρασκευή διαλυμάτων πρέπει : " + Chr(13) + "1. Να είναι  καθαρά και στεγνά." + Chr(13) + "2. Να έχουν αφαιρεθεί  τα καπάκια τους." + Chr(13) + "3. Κάθε μπουκάλι να είναι εφοδιασμένο με τον μαγνήτη του." + Chr(13) + "4. Σε όσα μπουκάλια προβλέπεται,  να έχουν προστεθεί τα απαραίτητα χημικά ." + Chr(13) + Chr(13) + "     Να ξεκινήσει το Robot   την παρασκευή Διαλυμάτων ;", 20, "Talos : Παρασκευή Διαλυμάτων")

If Ans = 7 Then   ' oxi
    Screen.MousePointer = 1
      exodos.Enabled = True
      Command1.Enabled = True
'      669  "Πρέπει να βγείτε από το πρόγραμμα  παρασκευής διαλυμάτων,πατώντας Εξοδος"
    
    MsgBox mL_Res(669), , "669.Talos"
  Exit Sub
End If
 
 
 Command1.Enabled = False

Do
Cur_Maker = Cur_Maker + 1
   If check_system(0) <> 0 Then
       Screen.MousePointer = 0
       '  670   "Παρακαλώ ανοίξτε την Ζυγαριά και πατείστε πάλι την Εναρξη..."
       MsgBox mL_Res(724), , "724.Talos" ' "Παρακαλώ ανοίξτε την Ζυγαριά και πατείστε πάλι την Εναρξη..."
       Command1.Enabled = True
       Exit Sub
   End If
   Form_Load
   Αρχη = Now
   Timer1.Enabled = True
   If Telos_Dialymaton = 1 Then Exit Do
   curq = 0: CurQ_Dis = curq
   start_dosing_Click
   If Telos_Dialymaton = 1 Then Exit Do
Loop Until Cur_Maker = sol_maker

Label25 = mL_Res(725) ' "Διακοπή Προγράμματος Παρασκευής Διαλυμάτων"
     
         ' ΧΤΥΠΗΜΑ ΚΟΥΔΟΥΝΙΟΥ
 cmd$ = "!" + Alarm1_off + ":"
RobSend (cmd$)

 cmd$ = "!" + Alarm1_on + ":"
RobSend (cmd$)

Data1.Refresh
DBGrid1.Refresh
 Command1.Enabled = True
 Command2.Enabled = False
  Timer1.Enabled = False
 exodos.Enabled = True
 MSComm1.PortOpen = False
 
 '676      "Ολες οι εντολές για παρασκευή διαλυμάτων έχουν ολοκληρωθεί . Το πρόγραμμα τερματίζεται..."
  MsgBox mL_Res(726), , "726.Talos"  ' "Ολες οι εντολές για παρασκευή διαλυμάτων έχουν ολοκληρωθεί . Το πρόγραμμα τερματίζεται..."

Screen.MousePointer = 1
End Sub

Private Sub Command2_Click()
'727 "Η διαδικασία παρασκευής διαλυμάτων θα τερματίσει με την ολοκλήρωση του τρέχοντος διαλύματος."

Ans = MsgBox(mL_Res(727) + Chr(13) + mL_Res(728), 20, "Talos")  '728 " Να τερματίσει ;"
If Ans = 6 Then
    sol_maker = Cur_Maker
    Label25.Visible = True
End If
  
 End Sub

Private Sub Dosing_Click()  'Δοσομετρηση ντρουβανισματος
Dim posot_druv As Single, z_count(10) As Single, druv_count(10) As Integer
Dim curq_count(10) As Single, curq_chek, asrw As Long



Last_Quantity = 0
Label8.Visible = False
2
'If system_ready = 1 Then rezero_Click
'If ZYGIS3() <> 0 Then GoTo 2
If system_ready = 1 Then
        MilSec 1000
6
        asrw = ZYGIS3()
        If asrw = 0 Then secwait (0): GoTo 6
      Tare = asrw: Tare_Dis = Str(Tare)
      curq = 0: CurQ_Dis = 0
End If
If system_ready = 1 Then Rob_FeedBack (59)             'Eλεγχος Robot στα 59 sec

Druv = 0
Druv_dis = "0"
arg = 0
posot_druv = 0
curq = 0
z_tot = 0
CurQ_Dis = curq

cmd$ = "!" + skon_motor_on + ":": RobSend (cmd$)
met_cubic = 0

Do
   Find_Target
   
   'druv=αριθμός ντρουβανισμάτων  & ca_factor >1  λεπτόρευστη
   If Druv = 0 And Ca_Factor > 1 Then
      z_met = z_met / 3
   End If
   
   If z_met > 5 Then
      ofel_met = z_met - 4
   Else
      ofel_met = 0
   End If
   
   met_cubic = met_cubic + (ofel_met) ^ 2
   
   'η μετατόπιση σε text box
   Z_Met_dis.Text = Format(z_met / 100, "###%")
   If Target - curq < 300 And Ca_Factor > 10 Then
      'z_syntetag = 700
   End If
   met = Val(z_syntetag) + 20000 'ανύψωση ντρουβανίσματος σε βήματα
   
   Rob_FeedBack (59)          ' Eλεγχος Robot στα 30 sec
   secwait (1)                    'Περιμενε 2 sec να πεσει η σκονη
   ' on-------------------------------------
   cmd$ = "!" + skon_motor_on + ":": RobSend (cmd$)
   cmd$ = "!var12 =" + Str(met) + ":"
   If z_met >= 3 Or z_met = 0 Then RobSend (cmd$)
''   If Ca_Factor > 100 Then
 ''   Else
       secwait (1)
   ' End If
   Rob_FeedBack (59)          'Eλεγχος Robot στα 30 sec
   'cmd$ = "!" + skon_motor_off + ":": RobSend (cmd$)
 
   'αριθμός ντρουβανισμάτων
   Druv = Druv + 1: Druv_dis = Str(Druv)
   
   If Druv > 30 Then Exit Do
   
   If Druv > IIf(Proti_Fora = True, 15, 10) And curq < Target / 10 Then
       curq = -2: CurQ_Dis = curq
       '729  "Δεν παρασκευάσθηκε  το διάλυμα  στο ποτήρι Νο "
       '730 " επειδή δεν έβγαζε  σκόνη το μπουκάλι  Νο "
       Printer.Print mL_Res(729) + Str(Pot) + mL_Res(730) + Str(buk)
       ' Printer.EndDoc
       Exit Do
   End If
   skones.Refresh
   secwait (3)
   If system_ready = 1 Then
7
       asrw = ZYGIS3()
       If asrw = 0 Then secwait (1): GoTo 7
   End If
         curq = asrw - Tare: CurQ_Dis = curq
         ' αν γυρίσει η ζυγαριά στα αρνητικά παίρνω σαν τάρα το άθροισμα
         If curq < 0 Then Tare = Tare + curq: curq = 0: CurQ_Dis = curq: Tare_Dis = Tare
         ' Q_Max_Quantity = η μέγιστη ποσότητα που έπεσε
         If curq - Last_Quantity > Q_Max_Quantity Then Q_Max_Quantity = curq - Last_Quantity
   If curq > 0 Then
       If met_cubic <> 0 Then
         ' ca_factor μέσος όρος  = ( παλιός ca_factor + νέος ca_factor=CurQ/met_cubic ) / 2
         Ca_Factor = (Ca_Factor + curq / met_cubic) / 2
       End If
   End If
   Ca_Factor_dis = Format(Ca_Factor, "##0.0000")
   Last_Dosage = curq - Last_Quantity
   Last_Quantity = curq
   Last_z_met = z_met
Loop Until curq >= Target * 97 / 100

cmd$ = "!" + skon_motor_off + ":" 'Με το στοπ του κινητήρα σκόνης σταματά η παραοχή ζεσ.νερού λογω της reset output out.3-1
 RobSend (cmd$)
    If system_ready = 1 And curq <= 0 Then
        nerot = -2: nero1 = -2
         curq = -2: CurQ_Dis = curq
          repair_powder_factors
         Exit Sub
   End If
  secwait (2)
10
    secwait (1)
     If system_ready = 1 And curq > 0 Then
          asrw = ZYGIS3()
          If asrw = 0 Then GoTo 10
           curq = asrw - Tare: CurQ_Dis = curq
    End If
 secwait (2)
 curq_chek = ZYGIS3()
 If system_ready = 1 And Abs(Val(curq_chek) - Tare - curq) > 0 Then GoTo 10
  repair_powder_factors
Label8.Visible = True
End Sub

Private Sub Mixing_Click()
    secwait (1)
     cmd$ = "!" + mixer_on + ":"
       RobSend (cmd$)
       cmd$ = nero2_off: RobSend (cmd$)
       cmd$ = nero1_off: RobSend (cmd$)
      tim_count = 0:  Label16.Visible = True
       mixing_make (mixtim)
      Rob_FeedBack (60)
      cmd$ = "!" + mixer_off + ":"
       RobSend (cmd$)
       Label16.Caption = "O.K."
       skones.Refresh
End Sub

Private Sub rezero_Click()
On Error Resume Next
If system_ready = 0 Then Exit Sub
17
MSComm1.Output = "CLEAR" + Chr$(13)
secwait (1)
MSComm1.Output = "ZERO" + Chr$(13)
secwait (1)
MSComm1.Output = "HI RANGE" + Chr$(13)
secwait (1)
asw1 = ZYGIS3()
secwait (1)
asw2 = ZYGIS3()
If Abs(asw1 - asw2) > 10 Then GoTo 17
End Sub
Function zyg_decimals()
 Dim counter
    
 If system_ready = 0 Then zyg_decimals = 0: Exit Function
    ' Flush the input buffer.
       MSComm1.Output = "SEND" + Chr$(13)
       MSComm1.InBufferCount = 0
    
    counter = 0: tot_counter = 0
    ' Wait for "OK" to come back from the modem.
    Do
       dummy = DoEvents()
       
       ' If there is data in the buffer, then read it.
       If MSComm1.InBufferCount Then
          FromModem$ = FromModem$ + MSComm1.Input
          ' Check for "OK".
          If InStr(FromModem$, Chr$(13)) Then
                If Mid$(FromModem$, 5, 1) = "." Or Mid$(FromModem$, 6, 1) = "." Or Mid$(FromModem$, 7, 1) = "." Then
                    Exit Do
                End If
          End If
       Else
            counter = counter + 1
            If tot_counter = 50 Then
               zyg_decimals = 0
             Exit Function
            End If
            If counter >= 100 Then  ' 500
                counter = 0: tot_counter = tot_counter + 1
                MSComm1.Output = "SEND" + Chr$(13)
                MSComm1.InBufferCount = 0
                FromModem$ = ""
             End If
        End If
    Loop
     zyg_decimals = 0: dec = 0
    poin = InStr(FromModem$, ".")
    For I = 1 To 20
       If IsNumeric(Mid(FromModem$, poin + I, 1)) Then dec = dec + 1 Else Exit For
     Next I
       
 
          zyg_decimals = dec
 
       
End Function
Function ZYGIS3() As Single
   Dim counter
    
 If system_ready = 0 Then ZYGIS3 = 0: Exit Function
    ' Flush the input buffer.
       MSComm1.Output = "SEND" + Chr$(13)
       MSComm1.InBufferCount = 0
    
    counter = 0: tot_counter = 0
    ' Wait for "OK" to come back from the modem.
    Do
       dummy = DoEvents()
       
       ' If there is data in the buffer, then read it.
       If MSComm1.InBufferCount Then
          FromModem$ = FromModem$ + MSComm1.Input
          ' Check for "OK".
          If InStr(FromModem$, ".") > 0 Then
             If InStr(FromModem$, Chr(13)) > 0 Then
                If Mid$(FromModem$, 5, 1) = "." Or Mid$(FromModem$, 6, 1) = "." Or Mid$(FromModem$, 7, 1) = "." Then
                    Exit Do
                End If
              End If
          End If
       Else
            counter = counter + 1
            If tot_counter = 100 Then
               ZYGIS3 = 0
             Exit Function
            End If
            If counter = 500 Then   ' 500
                counter = 0: tot_counter = tot_counter + 1
                MSComm1.Output = "SEND" + Chr$(13)
                MSComm1.InBufferCount = 0
                FromModem$ = ""
             End If
        End If
    Loop
            If Left(FromModem$, 1) = "?" Then
            zyg = Val("-" & LTrim(Mid$(FromModem$, 2, 8)))
        Else
            zyg = Val(Left$(FromModem$, 9))
        End If

         ZYGIS3 = zyg * 1000
    End Function

Private Sub Nero2_Change()

End Sub

Private Sub Sendmix_Click()
Asw = ZYGIS3()
 Rob_FeedBack (59)           ' Eλεγχος Robot στα 30 sec
 cmd$ = "!var12=100:"
 RobSend (cmd$)
 secwait (10)
 asw1 = ZYGIS3()
 If Asw - asw1 < 100000 Then
        MsgBox "The Robot did not take the Bottle !", , "Talos"
        End
End If
 Rob_FeedBack (59)
 Label15.Visible = True
 skones.Refresh
End Sub
                                       
                                         
Private Sub fere_buk()
'--------------------------------------------
If Beaker_Bottle <> "" Then
      cmd$ = "!d,,,(var68):!go,,,1:"
     RobSend (cmd$)
        Rob_FeedBack (59)
End If

 cmd$ = "!var12=" + Str(buk + 1000) + ":" + "!var36=" + Trim(Str(x_buk)) + ":" + "!var37=" + Trim(Str(y_buk)) + ":"
  RobSend (cmd$)
   'Rob_FeedBack (59)
   DUM_DUM = Wait_Robot(0)
   cmd$ = "!d,,(var88):!go,,1:"
 Label13.Visible = True

End Sub

Private Sub start_dosing_Click()
Dim mm_Periek
 
 Rob_FeedBack (59)
  
  cmd$ = "!" + Prouermansh_On + ":"
  RobSend (cmd$)
'Μεταφορά Ποτηριού
Fere_Pot
  
  cmd$ = "!" + Prouermansh_Off + ":"
  RobSend (cmd$)
  
  
  
  bar_pot = chek_pot  ' ΕΛΕΓΧΕΙ ΑΝ ΠΟΤΗΡΙ="ok"
If bar_pot = "not ok" Then
     curq = -1: CurQ_Dis = curq
            '  678 -9   "Δεν παρασκευάσθηκε  το διάλυμα  στο ποτήρι Νο ", " επειδή δεν βρέθηκε το κατάλληλο ποτήρι ... "
      Printer.Print mL_Res(678) + Str(Pot) + mL_Res(679)
     ' Printer.EndDoc
     rearange (1)   ' ΕΠΙΣΤΡΕΦΕΙ ΣΤΗΝ ΘΕΣΗ ΤΟΥ ΤΟ ΑΚΑΤΑΛΛΗΛΟ ΠΟΤΗΡΙ
     Ληξη = Time
     Διαρκεια = DateDiff("s", Αρχη, Ληξη)
     file_send   ' ΣΗΜΕΙΩΝΕΙ ΣΤΟ ΑΡΧΕΙΟ SKONES.MDB ΤΙ ΕΓΙΝΕ
     Exit Sub
Else
7
     ' If system_ready = 1 Then rezero_Click
End If

'Vale_Nero_Pre
 Rob_FeedBack (59)
fere_buk
                                    
                                    'Δοσομέτρηση Σκόνης
Command2.Enabled = False
If InStr(1, ximiko, "*") > 0 Then
   Dosing_Click   ' ΔΙΑΔΙΚΑΣΙΑ  ΔΟΣΟΜΕΤΡΗΣΗΣ ΣΚΟΝΗΣ (Papakiriakos)
Else
   curq = DOS_ALATI(Target)  'lagakis 24-2-2000
End If

On Error Resume Next
Printer.Print ximiko + " Δοσομετρηθείσα σκόνη :" + Str(curq)
On Error GoTo 0


If curq <= 0 Then ' Το μπουκάλι σκόνης δεν τροφοδοτεί με σκόνη
  Rob_FeedBack (59)  ' ΠΕΡΙΜΕΝΕΙ 59 SEC NA ΤΕΛΕΙΩΣΕΙ ΤΟ ROBOT
  cmd$ = "!var12 =20000:"
       RobSend (cmd$)
        Rob_FeedBack (60)          ' (50)
       rearange (1) ' ΕΠΙΣΤΡΕΦΕΙ ΠΙΣΩ ΤΟ ΠΟΤΗΡΙ ΠΟΥ ΔΕΝ ΕΧΕΙ ΣΚΟΝΗ ,ΚΑΙ ΔΙΑΚΟΠΤΕΙ ΔΙΑΔΙΚΑΣΙΑ ΠΑΡΑΣΚΕΥΗΣ
       file_send
        Rob_FeedBack (60)
      Exit Sub
End If
      
      cmd$ = "!var12=20000:"
      RobSend (cmd$)

                                   'Προσθήκη Νερού Διάλυσης

Βαλε_Νερο1_Click    '  ΠΡΟΣΘΕΤΕΙ ΤΟ ΝΕΡΟ ΔΙΑΛΥΣΗΣ
Command2.Enabled = True
 If nerot < 10000 Then
     While FindOutput(Val(Ready_Bit)) = 0: DoEvents: Wend
    secwait (1)
     rearange (1)
     curq = 0: CurQ_Dis = curq
     Ληξη = Time
    Διαρκεια = DateDiff("s", Αρχη, Ληξη)
   file_send
   While FindOutput(Val(Ready_Bit)) = 0: DoEvents: Wend
      secwait (2)
      cmd$ = "!var12=82:"   'Park
     RobSend (cmd$)
     secwait (2)
      While FindOutput(Val(Ready_Bit)) = 0: DoEvents: Wend
         cmd$ = "!" + Alarm1_on + ":"
         secwait (1)
          RobSend (cmd$)
          If nero1 < 20000 And nero1 >= 0 Then
              ' 680  "Με το άνοιγμα της βαλβίδας νερού , νερό δεν έρχεται. Η διαδικασία διακόπτεται..."
              MsgBox mL_Res(680), , "680.Talos"
         End If
         If nerot < 0 Then
                      '    "Με το κλείσιμο της βαλβίδας νερού , το νερό δεν σταματά. Η διαδικασία διακόπτεται..."
             MsgBox mL_Res(681), , "681.Talos"
        End If
        Telos_Dialymaton = 1
        Exit Sub
End If
Command2.Enabled = True
Rob_FeedBack (59)
Sendmix_Click  ' ΠΑΙΡΝΕΙ ΤΟ ΠΟΤΗΡΙ ΚΑΙ ΤΟ ΠΑΕΙ ΣΤΗΝ αναδευση
'Rob_FeedBack (59)
DUM_DUM = Wait_Robot(0)
If DemoVersion = 1 Then
Else
   Mixing_Click   ' ΜΕΤΡΑΕΙ ΤΟΝ ΧΡΟΝΟ ΑΝΑΔΕΥΣΗΣ
End If

Backmix_Click ' ΠΑΙΡΝΕΙ ΑΠΟ ΑΝΑΔΕΥΣΗ ΚΑΙ ΤΟ ΠΑΕΙ ΣΤΗΝ  ΖΥΓΑΡΙΑ
Command2.Enabled = False

   Βαλε_Νερο2_Click   ' ΒΑΖΕΙ ΝΕΡΟ ΑΡΑΙΩΣΗΣ ΜΕΧΡΙ ΤΕΛΙΚΗΣ ΠΕΡΙΕΚΤΙΚΟΤΗΤΟΣ

mm_Periek = IIf(Val(nerot) > 0, curq / (curq + Val(nerot)) * 100, 0)
On Error Resume Next
   Printer.Print " συνολ.νερό :" + nerot + "   Περιεκτικότητα " + Format(mm_Periek, "0.###0") + "%"
On Error GoTo 0




Command2.Enabled = True
Finish_Click   ' ΠΗΓΑΙΝΕΙ ΤΟ ΔΙΑΛΥΜΑ ΣΤΗΝ ΑΝΑΔΕΥΣΗ ΓΙΑ ΤΕΛΙΚΗ ΜΙΞΗ ΚΑΙ ΤΟ ΕΠΙΣΤΡΕΦΕΙ ΣΤΗΝ ΘΕΣΗ ΤΟΥ ΠΟΤΗΡΙΟΥ
vim = 0
End Sub

Private Sub Timer1_Timer()
Ληξη = Now
Διαρκεια = DateDiff("s", Αρχη, Now)
End Sub
Public Sub RobSend(argum As String)
      If system_ready = 0 Then Exit Sub
     
    ' καταγράφω την εντολή που έρχεται
           
         '  Set dbsJet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\watch.mdb")
         '    Set r = dbsJet.OpenRecordset("moves")
         '    r.AddNew
         '      r("movement") = left$(argum, 150)
         '        r("datetime") = Now
         '    r.update
         '    r.Close
     
     
     
     If InStr(argum, "var12") > 0 Then
         wait_to_understand (argum)
     Else
        m_Cmd$ = cmd$
        temp% = SendAT6400Block(768, argum, 0)
     End If
     



         
         
         
              
     
     
     
     
 End Sub
Function Zygis0(X As Single)
  Dim counter
    ' Flush the input buffer.
       MSComm1.Output = "SEND" + Chr$(13)
      MSComm1.InBufferCount = 0
        counter = 0
    ' Wait for "OK" to come back from the modem.
    Do
       dummy = DoEvents()
       
       ' If there is data in the buffer, then read it.
      If MSComm1.InBufferCount Then
          FromModem$ = FromModem$ + MSComm1.Input
          ' Check for "OK".
          If InStr(FromModem$, Chr$(13)) Then
                If Mid$(FromModem$, 5, 1) = "." Or Mid$(FromModem$, 6, 1) = "." Then
                    Exit Do
                End If
          End If
       Else
            counter = counter + 1
            If counter >= 500 Then  ' 500
                counter = 0
                MSComm1.Output = "SEND" + Chr$(13)
                MSComm1.InBufferCount = 0
                FromModem$ = ""
             End If
        End If
    Loop
         zyg = Val(Left$(FromModem$, 9))
             Zygis0 = zyg * 1000
End Function
Public Function chek_pot() As String
Dim bar_chek

baros_pot_max = 700000  'max βαρος μπουκαλιου με μαγνήτη και χωρίς τάπα
baros_pot_min = 100000 'min βαρος μπουκαλιου με μαγνήτη και χωρίς τάπα
13
met = met + 1
baros_pot = ZYGIS3()
If Val(baros_pot) > Val(baros_pot_max) Or Val(baros_pot) < Val(baros_pot_min) Then
      If met < 3 Then secwait (2): GoTo 13
   chek_pot = "not ok"
  Else
  chek_pot = "ok"
 End If
End Function

Private Function talostart()
   talostart = 1    '27-6
   Exit Function  '27-6


Screen.MousePointer = 11
  tim = 0
 'bit = Me.Ok_bit
 cmd$ = "!k:"
 RobSend (cmd$)
 secwait (1)

 cmd$ = "!miden:"
 RobSend (cmd$)
 secwait (1)

2 cmd$ = "!out." + Trim(Str(Ready_Bit)) + "-0:"
 RobSend (cmd$)
 secwait (1)
 If FindOutput(Ready_Bit) = 1 Then GoTo 2

   cmd$ = "!k:"
   RobSend (cmd$)

3 cmd$ = "talos:"
tim = tim + 1
If tim > 10 Then
  Screen.MousePointer = 11
 Exit Function
End If
 RobSend (cmd$)
 secwait (1)
 If FindOutput(Ready_Bit) = 0 Then GoTo 3
   talostart = 1
End Function

Public Function check_value(ByVal arg) As Integer
On Error GoTo er_chek_val
check_value = 1
 If IsNull(arg) Then check_value = 0
 If IsNumeric(arg) Then
     If Val(arg * 100000000) = 0 Then check_value = 0
End If
If arg = "" Then check_value = 0
 Exit Function
er_chek_val: Stop
End Function

Public Sub find_visco()
  cmd$ = "!var12 =" + Str(20000 + upm * 4) + ":"
  RobSend (cmd$)
  Rob_FeedBack (59)
End Sub

Private Sub Text1_Change()
        If system_ready = 1 Then rezero_Click
         Nero2 = 0: '  curq = 2000
         secwait (2)
         Ner_Stox = curq / periek - curq
         Nero2_Stox = Ner_Stox - Val(nero1)
         Nero2_Stox = Nero2_Stox - 2000
       
     If Val(nero1) > 0 Then nero2_count_stox = nero1_count * Nero2_Stox / nero1
     
                        '1o γεμισμα
              cmd$ = nero_on:   RobSend (cmd$)
                Nero2_Count = 0
         Do
            Nero2_Count = Nero2_Count + 1
            If system_ready = 1 Then Nero2 = ZYGIS3()
            nerot = Val(nero1) + Val(Nero2)
            If nerot > 560000 Then Exit Do
            fin_con = curq / (nerot + curq): fin_con1 = Format(fin_con, "0.0000 %")
       Loop Until nerot >= Nero2_Stox
       
        cmd$ = nero_off
        RobSend (cmd$)
        secwait (2)
         If system_ready = 1 Then Nero2 = ZYGIS3()
                 nerot = Val(nero1) + Val(Nero2)
                 
             '    2ο γεμισμα
                 
                 ner_yp = Nero2_Stox - Nero2
                 Nero2_Count = Nero2_Count * ner_yp / Nero2
      If ner_yp > 20 Then
         cmd$ = nero_on:   RobSend (cmd$)
         Do
         Ner = Ner + 1
             If system_ready = 1 Then Nero2 = ZYGIS3()
            nerot = Val(nero1) + Val(Nero2)
            If nerot > 600000 Then Exit Do
            fin_con = curq / (nerot + curq): fin_con1 = Format(fin_con, "0.0000 %")
          Loop Until Ner >= Nero2_Count
       End If
          Label18.Visible = True
          secwait (2)
      If system_ready = 1 Then Nero2 = ZYGIS3()
     nerot = Val(nero1) + Val(Nero2)
     fin_con = curq / (nerot + curq): fin_con1 = Format(fin_con, "0.0000 %")
         skones.Refresh
End Sub

Public Sub mixing_make(mix_tim)
tim_count = 0: Label29.Width = 15: Label29.Alignment = 0
Label29.BackStyle = 1: Label29.BorderStyle = 1: Label29.BackColor = Red
      cmd$ = "!" + mixer_on + ":"
       RobSend (cmd$)
        secwait (2)
      cmd$ = "!" + mixer_off + ":"
       RobSend (cmd$)
     secwait (1)
     cmd$ = "!" + mixer_on + ":"
       RobSend (cmd$)
  Do
     tim_count = tim_count + 1
     If tim_count / 5 = Int(tim_count / 5) Then GoSub mix_stop
      secwait (1)
      Label16 = tim_count
      Label29.Width = tim_count / mix_tim * 2000
      Label29 = Format(tim_count / mix_tim, "00%")
     Loop Until tim_count >= Val(mix_tim)
     Label29.BackStyle = 0: Label29.BorderStyle = 0: Label29.Alignment = 1
     Label29 = "Mixing Time"
Exit Sub
mix_stop:
      cmd$ = "!" + mixer_off + ":"
       RobSend (cmd$)
       secwait (0)
      cmd$ = "!" + mixer_on + ":"
       RobSend (cmd$)
       Return
End Sub

Public Sub rearange(argum)
While FindOutput(Val(Ready_Bit)) = 0: DoEvents: Wend
secwait (1)
 
If argum = 2 Then 'στείλε πίσω μπουκ & ποτήρι
   'cmd$ = "!d,,(var93):!go,,1:"
  ' RobSend (cmd$)
  ' secwait (2)
   'cmd$ = "!d,(var62):!go,1:"
  '  RobSend (cmd$)
    ' secwait (2)
   cmd$ = "!var12=" + Str(buk + 2000) + ":" + "!var36=" + Trim(Str(x_buk)) + ":" + "!var37=" + Trim(Str(y_buk)) + ":"
    RobSend (cmd$)
    secwait (1)
  While FindOutput(Val(Ready_Bit)) = 0: DoEvents: Wend
 secwait (1)
End If
  While FindOutput(Val(Ready_Bit)) = 0: DoEvents: Wend
  secwait (1)
   cmd$ = "!var12=" + Str(Pot + 4000) + ":" + "!var24=" + Trim(Str(x_pot)) + ":" + "!var77=" + Trim(Str(y_pot)) + ":" + "!var89=" + Trim(Str(z_pot)) + ":"
   RobSend (cmd$)
   secwait (1)
While FindOutput(Val(Ready_Bit)) = 0: DoEvents: Wend
secwait (1)
End Sub


Public Sub repair_powder_factors()

Set MyDB = Workspaces(0).OpenDatabase("c:\talos\skon_tb.mdb")
Set r = MyDB.OpenRecordset("Σκονες")

Do Until r.EOF
 If Val(buk) = r("Νο_Μπουκαλιου") Then
    r.Edit
      r("Datum") = Now()
      r("Factor_A") = Ca_Factor
      r("Ximiko") = ximiko
    r.Update
      r.Close
     Exit Sub
End If
  r.MoveNext
  Loop
  
  
'On Error Resume Next
     r.AddNew
       r("Νο_Μπουκαλιου") = buk
       r("Ximiko") = ximiko
       r("Factor_A") = Ca_Factor
      r("Datum") = Now()
    
     r.Update
    r.Close
'On Error GoTo 0

End Sub


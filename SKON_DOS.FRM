VERSION 4.00
Begin VB.Form skones 
   BackColor       =   &H00808000&
   Caption         =   "Talos : Αυτόματη Παρασκευή Διαλυμάτων"
   ClientHeight    =   6975
   ClientLeft      =   30
   ClientTop       =   1335
   ClientWidth     =   9660
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
   Height          =   7365
   Icon            =   "SKON_DOS.frx":0000
   Left            =   -30
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   9660
   Top             =   1005
   Width           =   9780
   WindowState     =   2  'Maximized
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
      TabIndex        =   102
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
      TabIndex        =   100
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
      TabIndex        =   98
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
      TabIndex        =   97
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
      TabIndex        =   96
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
      TabIndex        =   95
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
      TabIndex        =   94
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
      TabIndex        =   92
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
      TabIndex        =   90
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
      TabIndex        =   89
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
      TabIndex        =   74
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
      TabIndex        =   73
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
      TabIndex        =   72
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
      TabIndex        =   71
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
      TabIndex        =   70
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
      TabIndex        =   69
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
      TabIndex        =   68
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
      TabIndex        =   66
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
      TabIndex        =   65
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
      TabIndex        =   64
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
      TabIndex        =   62
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
      TabIndex        =   67
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
      TabIndex        =   101
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
      TabIndex        =   99
      Top             =   75
      Width           =   1170
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4890
      Top             =   975
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
      TabIndex        =   93
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
      TabIndex        =   88
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
      TabIndex        =   87
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
      TabIndex        =   86
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
      TabIndex        =   85
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
      TabIndex        =   84
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
      TabIndex        =   83
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
      TabIndex        =   82
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
      TabIndex        =   81
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
      TabIndex        =   80
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
      TabIndex        =   79
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
      TabIndex        =   78
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
      TabIndex        =   77
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
      TabIndex        =   76
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
      TabIndex        =   75
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
      TabIndex        =   63
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
      TabIndex        =   91
      Top             =   2730
      Visible         =   0   'False
      Width           =   2670
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "SKON_DOS.frx":0442
      Height          =   2955
      Left            =   30
      OleObjectBlob   =   "SKON_DOS.frx":0450
      TabIndex        =   61
      Top             =   3945
      Width           =   9570
   End
End
Attribute VB_Name = "skones"
Attribute VB_Creatable = False
Attribute VB_Exposed = False



Public arg As Integer, Ca_Factor As Single, nero1_ef As Long, CurQ As Long
Dim nero1_count   As Long, nero2_count  As Long, ner_stox As Long
Dim z_syntetag As Long, z_met As Single, Druv As Integer, Target As Long
Dim baros_pot_max As Long, baros_pot_min As Long, Tare As Long, Q_Max_Quantity As Long
Dim Last_Quantity As Long, Anevasma As Integer
Dim Telos_Dialymaton  ' 1=τελείωσαν τα διαλύματα      0=υπάρχουν ακόμη διαλύματα
 




Private Sub anix_zyg()
On Error Resume Next
 MSComm1.CommPort = Balance_Port
 MSComm1.Settings = Balance_Settings
 MSComm1.InputLen = 0
 MSComm1.PortOpen = True
End Sub
Private Sub Fere_Pot()
 rob_feedback (59)           ' Eλεγχος Robot στα 30 sec
 cmd$ = "!var12=" + Str(Pot + 19000) + ":" + "!var24=" + Trim(Str(x_pot)) + ":" + "!var77=" + Trim(Str(y_pot)) + ":" + "!var89=" + Trim(Str(z_pot)) + ":"
RobSend (cmd$)
 rob_feedback (59)

End Sub
Private Sub Βαλε_Νερο1_Click()
      
       If system_ready = 1 Then rezero_Click
       If system_ready = 1 Then secwait (6)           'Eλεγχος Robot στα 59 sec
        cmd$ = nero1_on
        RobSend (cmd$)
 If Val(vim) = 0 Then
    vim = Val(var67) + upm4 * Val(Water_Bypass)      'Μετατόπιση 4ου άξονα 20-30  χιλ δεξιά για να μην στάζει
    cmd$ = "!d,,," + Str(vim) + ":!go,,,1:"
    RobSend (cmd$)
 End If
 vim = 0
If system_ready = 1 Then
 nero1_count = 0
Do
   nero1_count = nero1_count + 1
     nero1 = ZYGIS3() + nero1_ef
     nerot = nero1
  If nero1_count > 2000 Then Exit Do
  Loop Until Val(nero1) > Val(nero_dial) * 1000
        cmd$ = nero1_off
        RobSend (cmd$)
        secwait (6)
     nero1 = nero1_ef + ZYGIS3()
If Val(nero1) > Val(nero_dial) * 1000 + 7000 Then
     nerot = -1: nero1 = -1: CurQ = -3: CurQ_Dis = CurQ
     Exit Sub
End If

      nerot = nero1
          secwait (2)
     nero_try = ZYGIS3()    'ελεγχος αν εκλεισε η βαλβίδα νερου
     If nero_try > nero1 + 7000 Then nerot = -1: nero1 = -1: CurQ = -4: CurQ_Dis = CurQ
  End If
 Label14.Visible = True
       If system_ready = 0 Then
         nero1 = Format(nero_dial, "###,###")
         nerot = nero1
      End If
       skones.Refresh
End Sub

Private Sub Βαλε_Νερο2_Click()
On Error Resume Next
Dim nero_chek
        If system_ready = 1 Then rezero_Click
            nero2 = 0:
            'curq = 212: nero1 = 112000 'για δοκιμες
            secwait (1)
            ner_stox = CurQ / periek - CurQ
            nero2_stox = ner_stox - 3000 - Val(nero1)
            
                                '1o γέμισμα
            cmd$ = nero2_on: RobSend (cmd$)
             nero2_count = 0
         Do
               nero2_count = nero2_count + 1
            If system_ready = 1 Then nero2 = ZYGIS3()
            nerot = Val(nero1) + Val(nero2)
            If nerot > 630000 Then Exit Do
            fin_con = CurQ / (nerot + CurQ): fin_con1 = Format(fin_con, "0.0000 %")
          Loop Until nerot >= ner_stox - 7000
        cmd$ = nero2_off: RobSend (cmd$)
        secwait (3)
         If system_ready = 1 Then
           secwait (3)
           nero2 = ZYGIS3()
          End If
            nerot = Val(nero1) + Val(nero2)
            fin_con = CurQ / (nerot + CurQ): fin_con1 = Format(fin_con, "0.0000 %")
            
             '    2ο γεμισμα
                ner_yp = ner_stox - nerot
                 nero2_count = 70 / 100 * nero2_count * ner_yp / nero2
      If ner_yp > 20 Then
        cmd$ = nero2_on: RobSend (cmd$)
        ner = 0
         Do
         ner = ner + 1
             If system_ready = 1 Then nero2 = ZYGIS3()
            nerot = Val(nero1) + Val(nero2)
            If nerot > 630000 Then Exit Do
            fin_con = CurQ / (nerot + CurQ): fin_con1 = Format(fin_con, "0.0000 %")
          Loop Until ner >= nero2_count
            cmd$ = nero2_off: RobSend (cmd$)
    End If
          Label18.Visible = True
20
          secwait (2)
      If system_ready = 1 Then nero2 = ZYGIS3()
        secwait (2)
       If system_ready = 1 Then nero_chek = ZYGIS3()
      ' If system_ready = 1 And Abs(Val(nero2) - Val(nero_chek)) > 0 Then GoTo 20
     nerot = Val(nero1) + Val(nero2)
     fin_con = CurQ / (nerot + CurQ): fin_con1 = Format(fin_con, "0.0000 %")
         skones.Refresh
End Sub



Private Sub Backmix_Click()
Dim bar_chek
 rob_feedback (59)           ' Eλεγχος Robot στα 30 sec
 cmd$ = "!var12=101:"
 RobSend (cmd$)
 rezero_Click
 rob_feedback (59)
 Label17.Visible = True
12
 
 bar = ZYGIS3()
 secwait (2)
 bar_chek = ZYGIS3()
' If Abs(Val(bar) - Val(bar_chek)) > 0 Then secwait (2): GoTo 12
 If Val(nero2) = 0 Then
       nero1 = bar - Val(baros_pot) - CurQ 'ξαναμέτρα το νερο λογω εξάτμισης
 Else
       nerot = bar - baros_pot - CurQ
       fin_con = CurQ / (nerot + CurQ): fin_con1 = Format(fin_con, "0.0000 %")
 End If
  If Val(vim) = 0 Then
    vim = Val(var67) + upm4 * Val(Water_Bypass) 'Μετατόπιση 4ου άξονα 30 χιλ δεξιά για να μην στάζει
    cmd$ = "!d,,," + Str(vim) + ":!go,,,1:"
    RobSend (cmd$)
 End If
 vim = 0
 skones.Refresh
 
End Sub
Private Sub Vale_Nero_Pre()
1
    cmd$ = nero1_on
     RobSend (cmd$)
     MilSec 1500
    cmd$ = nero1_off
     RobSend (cmd$)
      MilSec 3000
If system_ready = 1 Then nero1_ef = ZYGIS3(): nero1 = nero1_ef
If nero1_ef < 500 Then GoTo 1
End Sub
Private Sub Beready_Click()
 rob_feedback (59)           ' Eλεγχος Robot στα 30 sec
 cmd$ = "!var12=" + Str(Pot + 19000) + ":" + "!var24=" + Trim(Str(x_pot)) + ":" + "!var77=" + Trim(Str(y_pot)) + ":" + "!var89=" + Trim(Str(z_pot)) + ":"
RobSend (cmd$)
 rob_feedback (59)

      
End Sub
Function file_send()
Dim MyDB As Database, r As Recordset
If system_ready = 0 Then Exit Function
 Set MyDB = Workspaces(0).OpenDatabase("c:\talos\skon_tb.MDB")
Set r = MyDB.OpenRecordset("Διαλυματα")
    r.AddNew
 r("nerostox") = nerostox
     r("periek") = periek
    r("buk") = Buk
     r("pot") = Pot
     r("mixtim") = s_mixtim
      r("mixtim_fin") = mixtim_fin
     r("thermo1") = thermo1
     r("thermo2") = thermo2
     r("nero_act") = nerot
     r("color_act") = CurQ
     r("duration") = Διαρκεια
     r("datum") = Now
     r("druv") = Druv
     r("nero_dial") = nero_dial
     r("ximiko") = XIMIKO
     r.Update
   r.Close
  
  Set r = MyDB.OpenRecordset("Εντολες")
Do Until r.EOF
entol = 0
 If IsNull(r("nero_act")) Then
    r.Edit
     r("nero_act") = nerot
     r("color_act") = CurQ
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

Sub rob_feedback(arg As Integer) 'Εξεταζει την αποκριση του Robot
secwait (1)
While FindOutput(Val(Ready_Bit)) = 0
     curtim = Second(Time)
     If curtim <> reftim Then
      sec = sec + 1: reftim = curtim
    If sec > arg Then
     tim = Time
     tit = "T A L O S"
     msg = "Ωρα " + Str(tim) + ". Το Robot δεν ανταποκρίνεται . Η διαδικασία τερματίζεται."
     MsgBox msg, , "Talos . Αυτόματη Παρασκευή Διαλυμάτων"
     Telos_Dialymaton = 1
     Exit Sub
    End If
      DoEvents
     End If
  Wend
 
End Sub


Function check_system(afr)
rezero_Click
asw = ZYGIS3()
If asw = " " Then
  check_system = 1
  Exit Function
End If
rezero_Click
End Function

Private Sub exodos_Click()
On Error Resume Next
  MSComm1.PortOpen = False
    Unload Me
   ' fereSK.Show
End Sub

Private Sub Find_Target()
Dim posot_Θεωρ As Single, pos As Integer, Buttom_step As Integer
pos = Target - CurQ + Tare
If pos < 0 Then
  z_met = 1 ': posot_Θεωρ = z_met * z_met * a_factor / 1000
  Exit Sub
End If
If Ca_Factor = 0 Then Ca_Factor = 10: Ca_Factor_dis = Ca_Factor
z_met = Sqr(pos / Ca_Factor)
GoSub Select_Safety_Factor
If z_met > 100 Then z_met = 100
z_met = z_met * Safety_Factor 'Παράγων ασφάλειας
GoSub Find_Buttom
If CurQ <= 50 And Druv > 0 Then Anevasma = Anevasma + 10: Buttom_step = Buttom_step + Anevasma
If z_met < Buttom_step Then z_met = Buttom_step
z_syntetag = upm * z_met * 25 / 100
If (Val(z_syntetag) + Val(var88)) > Val(var93) Then z_syntetag = Val(var93) - Val(var88)
z_met = z_syntetag / upm / 25 * 100

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
Safety_Factor = 0.75
End Select
Return

Find_Buttom:
Select Case Ca_Factor
Case Is > 1.5
Buttom_step = 10
Case Is > 0.5
Buttom_step = 15
Case Is > 0.1
Buttom_step = 20
Case Is > 0.01
Buttom_step = 25
Case Else
Buttom_step = 30
End Select
Return
End Sub

Private Sub Finish_Click()
If mixtim_fin > 0 Then
   Sendmix_Click
   Mixing.Caption = "Τελική Ανάδευση"
   Label29.Caption = "Final Mixing  "
   Label16.Caption = "0"
   mixtim = mixtim_fin
      cmd$ = "!" + mixer_on + ":"
       RobSend (cmd$)
       mixing_make (mixtim_fin)
       cmd$ = "!" + mixer_off + ":"
       RobSend (cmd$)
End If
rob_feedback (59)
cmd$ = "!var12=" + Str(Pot + 200) + ":"
RobSend (cmd$)
 rob_feedback (59)
 Label19.Visible = True
Ληξη = Now
Διαρκεια = DateDiff("s", Αρχη, Ληξη)
file_send
repair_powder_factors
skones.Refresh
  Screen.MousePointer = 0
End Sub



Private Sub Form_Load()
'On Error GoTo er_chek
Dim MyDB As Database, r As Recordset, MySet As Recordset
Ca_Factor = 0
FindValves
system_ready = IsOSLoaded(768)
Q_Max_Quantity = 0
If Val(Pot) > 0 Then GoTo jump1
 'Σβήσιμο παλιων παρασκευών
Set MyDB = Workspaces(0).OpenDatabase("c:\talos\skon_tb.mdb")
Set MySet = MyDB.OpenRecordset("Εντολες", dbOpenDynaset)
sql_del = "DELETE DISTINCTROW Εντολες.*, Εντολες.nero_act FROM Εντολες WHERE ((Not Εντολες.nero_act Is Null));"
MyDB.Execute sql_del
 Timer1.Enabled = False
jump1:


Mixing.Caption = "Ανάδευση"
Label29.Caption = "Mixing Time  "
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
nero2 = 0: fin_con = 0: nerot = 0: fin_con1 = "0 %"
Data1.Refresh
DBGrid1.Refresh
skones.Refresh

Set MyDB = Workspaces(0).OpenDatabase("c:\talos\skon_tb.mdb")
Set r = MyDB.OpenRecordset("Εντολες")
Do Until r.EOF
entol = 0
 If IsNull(r("nero_act")) Then
    nerostox = r("nerostox")
    periek = r("periek")
    periek_dis = Format(periek, "0.##0%")
    Buk = r("buk")
    Pot = r("pot")
    mixtim = r("mixtim")
    s_mixtim = r("mixtim")
    mixtim_fin = r("mixtim_fin")
    thermo1 = r("thermo1")
    thermo2 = r("thermo2")
    nero_dial = r("nero_dial")
    Label14.Caption = nero_dial
   If r("ximiko") <> "" Then XIMIKO = r("ximiko") Else XIMIKO = "Μπουκάλα Δοκιμαστική"
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
 
 If check_value(nerostox) = 0 Then MsgBox "Δεν υπάρχει  ποσότητα Διαλύματος.Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.":   Telos_Dialymaton = 1: Exit Sub
 If check_value(periek) = 0 Then MsgBox "Δεν υπάρχει  περιεκτικότητα Διαλύματος.Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 If check_value(Buk) = 0 Then MsgBox "Δεν υπάρχει  Μπουκάλι  Διαλύματος.Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 If check_value(Pot) = 0 Then MsgBox "Δεν υπάρχει  ποτήρι  Διαλύματος.Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 If check_value(mixtim) = 0 Then MsgBox "Δεν υπάρχει  Χρόνος  Διάλυσης.Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
  If check_value(mixtim_fin) = 0 Then MsgBox "Δεν υπάρχει  Τελικός Χρόνος  Διάλυσης.Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 If check_value(thermo1) = 0 Then MsgBox "Δεν υπάρχει  Θερμοκρασία νερού  Διάλυσης.Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 If check_value(thermo2) = 0 Then MsgBox "Δεν υπάρχει  Θερμοκρασία νερού  Αραίωσης.Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 If check_value(nero_dial) = 0 Then MsgBox "Δεν δηλώθηκε η ποσότητα του νερού Διάλυσης.Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 If Val(nero_dial) < 20 Then MsgBox "Δεν δηλώθηκε η ποσότητα του νερού Διάλυσης.Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 
Set r = MyDB.OpenRecordset("Σκονες")
Do Until r.EOF
 If Buk = r("Νο_Μπουκαλιου") Then
    Ca_Factor = Val(r("Factor_A"))
    Ca_Factor_dis = Ca_Factor
    Exit Do
End If
  r.MoveNext
  Loop
  r.Close
If Factor_a = "" Then Factor_a = 35


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
  
 If check_value(upm) = 0 Then MsgBox "Στο πρόγραμμα Robot Setup - Χαρακτηριστικά Κίνησης Robot δεν είναι δηλωμένη   η αναλογία μμ ανά περιστροφή κινητήρα 3ου  άξονα.Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 If check_value(upm4) = 0 Then MsgBox "Στο πρόγραμμα Robot Setup - Χαρακτηριστικά Κίνησης Robot δεν είναι δηλωμένη  η αναλογία μμ ανά περιστροφή κινητήρα 4ου  άξονα.Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
     
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
            If r("Νο_Τεμαχιου") = Val(Buk) Then
              x_buk = r("Συντεταγμένη_Χ_user")
              y_buk = r("Συντεταγμένη_Υ_user")
              z_buk = r("Συντεταγμένη_Ζ_user")
            End If
           End If
           r.MoveNext
          Loop
  r.Close
  
 If check_value(x_pot) = 0 Then MsgBox "Δεν υπάρχουν  οι Συντεταγμένες Χ Ποτηριού.Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 If check_value(y_pot) = 0 Then MsgBox "Δεν υπάρχουν  οι Συντεταγμένες Υ Ποτηριού.Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 If check_value(z_pot) = 0 Then MsgBox "Δεν υπάρχουν  οι Συντεταγμένες Ζ Ποτηριού.Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 If check_value(x_buk) = 0 Then MsgBox "Δεν υπάρχουν  οι Συντεταγμένες Χ Μπουκαλιού.Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 If check_value(y_buk) = 0 Then MsgBox "Δεν υπάρχουν  οι Συντεταγμένες Υ Μπουκαλιού.Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 If check_value(z_buk) = 0 Then MsgBox "Δεν υπάρχουν  οι Συντεταγμένες Ζ Μπουκαλιού.Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
  
  
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
r.MoveNext
  Loop
  r.Close
  
 If check_value(var93) = 0 Then MsgBox "Δεν υπάρχει η Παράμετρος var93.Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 If check_value(var88) = 0 Then MsgBox "Δεν υπάρχει η Παράμετρος var88.Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
  If check_value(var67) = 0 Then MsgBox "Δεν υπάρχει η Παράμετρος var67.Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 
  If entol = 0 Then Telos_Dialymaton = 1: Exit Sub
   Target = periek * Val(nerostox): Target_Dis = Str(Target)
   
  Select Case thermo1
  Case 1
     Label21 = "Κρύο"
   Case 2
      Label21 = "Χλιαρό"
   Case 3
      Label21 = "Ζεστό"
End Select

  Select Case thermo2
  Case 1
    Label22 = "Κρύο"
   Case 2
      Label22 = "Χλιαρό"
   Case 3
      Label22 = "Ζεστό"
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



         
 If check_value(Ready_Bit) = 0 Then MsgBox "Δεν είναι  δηλωμένο το Ok Bit. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός = 2 ). Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 If check_value(skon_motor_on) = 0 Then MsgBox "Δεν είναι  δηλωμένο το output του κινητήρα της σκόνης. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός =4 ). Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 If check_value(skon_motor_off) = 0 Then MsgBox "Δεν είναι  δηλωμένο το output του κινητήρα της σκόνης. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός =4 ). Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 If check_value(nero1_on) = 0 Then MsgBox "Δεν είναι  δηλωμένο το output της βαλβίδας του νερού διάλυσης. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός =6 ). Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 If check_value(nero1_off) = 0 Then MsgBox "Δεν είναι  δηλωμένο το output της βαλβίδας του νερού διάλυσης. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός =6 ). Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 If check_value(nero2_on) = 0 Then MsgBox "Δεν είναι  δηλωμένο το output της βαλβίδας του νερού αραίωσης. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός =5 ). Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 If check_value(nero2_off) = 0 Then MsgBox "Δεν είναι  δηλωμένο το output της βαλβίδας του νερού αραίωσης. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός =5 ).Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 If check_value(Cold_Water_On) = 0 Then MsgBox "Δεν είναι  δηλωμένο το output της βαλβίδας του κρύου νερού. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός =5 ). Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 If check_value(Cold_Water_Off) = 0 Then MsgBox "Δεν είναι  δηλωμένο το output της βαλβίδας του κρύου νερού. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός =5 ). Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 If check_value(Hot_Water_On) = 0 Then MsgBox "Δεν είναι  δηλωμένο το output της βαλβίδας του ζεστού νερού. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός =6 ). Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 If check_value(Hot_Water_Off) = 0 Then MsgBox "Δεν είναι  δηλωμένο το output της βαλβίδας του ζεστού νερού. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός =6 ). Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 If check_value(mixer_on) = 0 Then MsgBox "Δεν είναι  δηλωμένο το output του κινητήρα της ανάδευσης. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός =7).Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 If check_value(mixer_off) = 0 Then MsgBox "Δεν είναι  δηλωμένο το output του κινητήρα της ανάδευσης. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός =7).Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 If check_value(Alarm1_on) = 0 Then MsgBox "Δεν είναι  δηλωμένο το outpout του Βομβητή. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός =11).Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub
 If check_value(Balance_Port) = 0 Then MsgBox "Δεν είναι  δηλωμένη η πόρτα της ζυγαριάς. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός =8).Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων.": Telos_Dialymaton = 1: Exit Sub

Me.WindowState = 2
'If system_ready = 1 Then
'anix_zyg

ex_sub:
Exit Sub
er_chek:
 If Err = 3189 Then MsgBox "Το αρχείο Εντολές Σκόνης είναι κατειλημένο από κάποιο άλλο πρόγραμμα. To παρόν πρόγραμμα δεν μπορεί να συνεχίσει...", , "Talos : Παρασκευή Διαλυμάτων"
Resume Next
End Sub



Private Sub Command1_Click()
anix_zyg

ValveReset
Telos_Dialymaton = 0
met = 0
exodos.Enabled = False
Command1.Enabled = False
If system_ready = 1 Then
Do
    MSComm1.Output = "HI RANGE" + Chr$(13)
     secwait (1)
     met = met + 1
     dec = zyg_decimals()
       If met = 5 Then Exit Do
 Loop Until dec = 2
End If
If dec <> 2 Then
 ans = MsgBox("Μετά από 5 προσπάθειες, η ζυγαριά δεν ανταποκρίνεται. Ανοίξτε την ζυγαριά και πατείστε το πλήκτρο F  για να την ρυθμίσετε σε ζύγιση 2 δεκαδικών." + Chr(13) + Chr(13) + "Πατείστε ΟΚ όταν είστε έτοιμοι να συνεχίσει το πρόγραμμα ή Cancel για ακύρωση.", 1, "Talos")
    If ans = 2 Then Command1.Enabled = True: exodos.Enabled = True: Exit Sub
  Command1_Click
End If
Data1.Refresh
DBGrid1.Refresh
skones.Refresh
 Label25.Visible = False
 Label25 = "Εκτέλεση Τελευταίας Παρασκευής"
If talostart() = 0 Then
 MsgBox "Το Πρόγραμμα Κίνησης του Robot δεν μπορεί να τρέξει.Η διαδικασία τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων"
  talostart = 0
  Telos_Dialymaton = 1: Exit Sub
  fereSK.Show
End If
ans = MsgBox("Τα μπουκάλια που πρόκειται να χρησιμοποιηθούν στην παρασκευή διαλυμάτων πρέπει : " + Chr(13) + Chr(13) + "1. Να είναι  καθαρά και στεγνά." + Chr(13) + "2. Να έχουν αφαιρεθεί  τα καπάκια τους." + Chr(13) + "3. Κάθε μπουκάλι να είναι εφοδιασμένο με τον μαγνήτη του." + Chr(13) + "4. Σε όσα μπουκάλια προβλέπεται,  να έχουν προστεθεί τα απαραίτητα χημικά ." + Chr(13) + Chr(13) + Chr(13) + "     Να ξεκινήσει το Robot   την παρασκευή Διαλυμάτων;", 20, "Talos : Παρασκευή Διαλυμάτων")
If ans = 7 Then   ' oxi
    Screen.MousePointer = 1
      exodos.Enabled = True
      Command1.Enabled = True
    MsgBox "Πρέπει να βγείτε από το πρόγραμμα  παρασκευής διαλυμάτων,πατώντας Εξοδος", , "Talos : Παρασκευή Διαλυμάτων"
  Exit Sub
End If
 
 
 Command1.Enabled = False

Do
Cur_Maker = Cur_Maker + 1
    If check_system(0) <> 0 Then
       Screen.MousePointer = 0
       MsgBox "Παρακαλώ ανοίξτε την Ζυγαριά και πατείστε πάλι την Εναρξη...", , "Talos : Παρασκευή Διαλυμάτων."
       Command1.Enabled = True
       Exit Sub
   End If
   Form_Load
   Αρχη = Now
    Timer1.Enabled = True
   If Telos_Dialymaton = 1 Then Exit Do
   CurQ = 0: CurQ_Dis = CurQ
   start_dosing_Click
   If Telos_Dialymaton = 1 Then Exit Do
Loop Until Cur_Maker = sol_maker
Label25 = "Διακοπή Προγράμματος Παρασκευής Διαλυμάτων"
Data1.Refresh
DBGrid1.Refresh
 Command1.Enabled = True
 Command2.Enabled = False
  Timer1.Enabled = False
 exodos.Enabled = True
 MSComm1.PortOpen = False
  MsgBox "Ολες οι εντολές για παρασκευή διαλυμάτων έχουν ολοκληρωθεί . Το πρόγραμμα τερματίζεται...", , "Talos : Παρασκευή Διαλυμάτων"
Screen.MousePointer = 1
 
End Sub

Private Sub Command2_Click()
ans = MsgBox("Η διαδικασία παρασκευής διαλυμάτων θα τερματίσει με την ολοκλήρωση του τρέχοντος διαλύματος." + Chr(13) + " Να τερματίσει ;", 20, "Talos : Αυτόματη Παρασκευή Διαλυμάτων ")
If ans = 6 Then
    sol_maker = Cur_Maker
    Label25.Visible = True
End If
  
 End Sub

Private Sub Dosing_Click()  'Δοσομετρηση ντρουβανισματος
Dim posot_druv As Single, z_count(10) As Single, druv_count(10) As Integer, curq_count(10) As Single, curq_chek
Last_Quantity = 0
Label8.Visible = False
2
If system_ready = 1 Then rezero_Click
If ZYGIS3() <> 0 Then GoTo 2
If system_ready = 1 Then
      Tare = ZYGIS3(): Tare_Dis = Str(Tare)
      CurQ = 0: CurQ_Dis = CurQ
End If
If system_ready = 1 Then rob_feedback (59)             'Eλεγχος Robot στα 59 sec
Druv = 0: Druv_dis = "0": arg = 0: posot_druv = 0: CurQ = 0: z_tot = 0: CurQ_Dis = CurQ
cmd$ = "!" + skon_motor_on + ":": RobSend (cmd$)
met_cubic = 0
 Do
  Find_Target
  If z_met > 10 Then ofel_met = z_met - 10 Else ofel_met = 0
  met_cubic = met_cubic + (ofel_met) ^ 2
  Z_Met_dis = Format(z_met / 100, "###%")
  met = Val(z_syntetag) + 20000
  rob_feedback (59)          ' Eλεγχος Robot στα 30 sec
  secwait (1)                    'Περιμενε 2 sec να πεσει η σκονη
cmd$ = "!" + skon_motor_on + ":": RobSend (cmd$)

  cmd$ = "!var12 =" + Str(met) + ":"
If z_met >= 3 Or z_met = 0 Then RobSend (cmd$)

 secwait (1)
 rob_feedback (20)          'Eλεγχος Robot στα 30 sec
cmd$ = "!" + skon_motor_off + ":": RobSend (cmd$)
 
 Druv = Druv + 1: Druv_dis = Str(Druv)
  If Druv > 20 Then Exit Do
  If Druv > 5 And CurQ < 190 Then
    CurQ = -2: CurQ_Dis = CurQ
  Printer.Print "Δεν παρασκευάσθηκε  το διάλυμα  στο ποτήρι Νο " + Str(Pot) + " επειδή δεν έβγαζε  σκόνη το μπουκάλι  Νο " + Str(Buk)
   Printer.EndDoc
    Exit Do
 End If
 
 skones.Refresh
 secwait (3)
 If system_ready = 1 Then CurQ = ZYGIS3() - Tare: CurQ_Dis = CurQ - Tare
 If CurQ < 0 Then Tare = Tare + CurQ: CurQ = 0: CurQ_Dis = CurQ: Tare_Dis = Tare
 If CurQ - Last_Quantity > Q_Max_Quantity Then Q_Max_Quantity = CurQ - Last_Quantity
 If CurQ > 0 And met_cubic > 0 Then Ca_Factor = (Ca_Factor + CurQ / met_cubic) / 2
  Ca_Factor_dis = Format(Ca_Factor, "##0.0000")
Last_Quantity = CurQ - Tare
Loop Until CurQ >= Target - 40 + Tare
cmd$ = "!" + skon_motor_off + ":"
 RobSend (cmd$)
    If system_ready = 1 And CurQ <= 0 Then
        nerot = -2: nero1 = -2
         CurQ = -2: CurQ_Dis = CurQ
         Exit Sub
   End If
  secwait (2)
10
    secwait (1)
     If system_ready = 1 And CurQ > 0 Then
           CurQ = ZYGIS3() - Tare: CurQ_Dis = CurQ
    End If
 secwait (2)
 curq_chek = ZYGIS3()
 If system_ready = 1 And Abs(Val(curq_chek) - Tare - CurQ) > 0 Then GoTo 10
Label8.Visible = True
End Sub

Private Sub Mixing_Click()
    secwait (1)
     cmd$ = "!" + mixer_on + ":"
       RobSend (cmd$)
      tim_count = 0:  Label16.Visible = True
       mixing_make (mixtim)
      rob_feedback (60)
      cmd$ = "!" + mixer_off + ":"
       RobSend (cmd$)
       Label16.Caption = "O.K."
       skones.Refresh
End Sub

Private Sub MSComm1_Click()

End Sub

Private Sub rezero_Click()
If system_ready = 0 Then Exit Sub
17
MSComm1.Output = "CLEAR" + Chr$(13)
secwait (1)
MSComm1.Output = "ZERO" + Chr$(13)
MSComm1.Output = "HI RANGE" + Chr$(13)
If ZYGIS3() > 10 Or ZYGIS3() < -10 Then GoTo 17
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
            If tot_counter = 10 Then
               zyg_decimals = 0
             Exit Function
            End If
            If counter >= 50 Then  ' 500
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
            If tot_counter = 20 Then
               ZYGIS3 = 0
             Exit Function
            End If
            If counter = 100 Then   ' 500
                counter = 0: tot_counter = tot_counter + 1
                MSComm1.Output = "SEND" + Chr$(13)
                MSComm1.InBufferCount = 0
                FromModem$ = ""
             End If
        End If
    Loop
         zyg = Val(Left$(FromModem$, 9))
         ZYGIS3 = zyg * 1000
    End Function

Private Sub Sendmix_Click()
 rob_feedback (59)           ' Eλεγχος Robot στα 30 sec
 cmd$ = "!var12=100:"
 RobSend (cmd$)
 rob_feedback (59)
 Label15.Visible = True
 skones.Refresh
End Sub
                                       
                                         
Private Sub fere_buk()
If Beaker_Bottle <> "" Then
      cmd$ = "!d,,,(var68):!go,,,1:"
     RobSend (cmd$)
        rob_feedback (20)
End If

 cmd$ = "!var12=" + Str(Buk + 1000) + ":" + "!var36=" + Trim(Str(x_buk)) + ":" + "!var37=" + Trim(Str(y_buk)) + ":"
  RobSend (cmd$)
 Label13.Visible = True

End Sub
Private Sub start_dosing_Click()
 rob_feedback (40)
                                       'Μεταφορά Ποτηριού
Fere_Pot
  bar_pot = chek_pot  ' ΕΛΕΓΧΕΙ ΑΝ ΠΟΤΗΡΙ="ok"
  If bar_pot = "not ok" Then
     CurQ = -1: CurQ_Dis = CurQ
      Printer.Print "Δεν παρασκευάσθηκε  το διάλυμα  στο ποτήρι Νο " + Str(Pot) + " επειδή δεν βρέθηκε το κατάλληλο ποτήρι ... "
      Printer.EndDoc
     rearange (1)   ' ΕΠΙΣΤΡΕΦΕΙ ΣΤΗΝ ΘΕΣΗ ΤΟΥ ΤΟ ΑΚΑΤΑΛΛΗΛΟ ΠΟΤΗΡΙ
     Ληξη = Time
     Διαρκεια = DateDiff("s", Αρχη, Ληξη)
     file_send   ' ΣΗΜΕΙΩΝΕΙ ΣΤΟ ΑΡΧΕΙΟ SKONES.MDB ΤΙ ΕΓΙΝΕ
     Exit Sub
Else
7
      If system_ready = 1 Then rezero_Click
End If

Vale_Nero_Pre
fere_buk

                                    'Δοσομέτρηση Σκόνης
Command2.Enabled = False
Dosing_Click   ' ΔΙΑΔΙΚΑΣΙΑ  ΔΟΣΟΜΕΤΡΗΣΗΣ ΣΚΟΝΗΣ
If CurQ <= 0 Then ' Το μπουκάλι σκόνης δεν τροφοδοτεί με σκόνη
  rob_feedback (59)  ' ΠΕΡΙΜΕΝΕΙ 59 SEC NA ΤΕΛΕΙΩΣΕΙ ΤΟ ROBOT
  cmd$ = "!var12 =20000:"
       RobSend (cmd$)
        rob_feedback (40)          ' (50)
       rearange (1) ' ΕΠΙΣΤΡΕΦΕΙ ΠΙΣΩ ΤΟ ΠΟΤΗΡΙ ΠΟΥ ΔΕΝ ΕΧΕΙ ΣΚΟΝΗ ,ΚΑΙ ΔΙΑΚΟΠΤΕΙ ΔΙΑΔΙΚΑΣΙΑ ΠΑΡΑΣΚΕΥΗΣ
       file_send
        rob_feedback (40)
      Exit Sub
End If
      cmd$ = "!var12=20000:"
      RobSend (cmd$)

                                   'Προσθήκη Νερού Διάλυσης

Βαλε_Νερο1_Click    '  ΠΡΟΣΘΕΤΕΙ ΤΟ ΝΕΡΟ ΔΙΑΛΥΣΗΣ
Command2.Enabled = True
 If nerot < 20000 Then
     While FindOutput(Val(Ready_Bit)) = 0:  Wend
    secwait (1)
     rearange (1)
     CurQ = 0: CurQ_Dis = CurQ
     Ληξη = Time
    Διαρκεια = DateDiff("s", Αρχη, Ληξη)
   file_send
   While FindOutput(Val(Ready_Bit)) = 0:  Wend
      secwait (2)
      cmd$ = "!var12=82:"   'Park
     RobSend (cmd$)
     secwait (2)
      While FindOutput(Val(Ready_Bit)) = 0: Wend
         cmd$ = "!" + Alarm1_on + ":"
         secwait (1)
          RobSend (cmd$)
          If nero1 < 20000 And nero1 >= 0 Then
              MsgBox "Με το άνοιγμα της βαλβίδας νερού , νερό δεν έρχεται. Η διαδικασία διακόπτεται...", , "Talos : Παρασκευή Διαλυμάτων."
         End If
         If nerot < 0 Then
             MsgBox "Με το κλείσιμο της βαλβίδας νερού , το νερό δεν σταματά. Η διαδικασία διακόπτεται...", , "Talos : Παρασκευή Διαλυμάτων."
        End If
        Telos_Dialymaton = 1
        Exit Sub
End If
Command2.Enabled = True
rob_feedback (59)
Sendmix_Click  ' ΠΑΙΡΝΕΙ ΤΟ ΠΟΤΗΡΙ ΚΑΙ ΤΟ ΠΑΕΙ ΣΤΗΝ αναδευση
rob_feedback (59)
Mixing_Click   ' ΜΕΤΡΑΕΙ ΤΟΝ ΧΡΟΝΟ ΑΝΑΔΕΥΣΗΣ
Backmix_Click ' ΠΑΙΡΝΕΙ ΑΠΟ ΑΝΑΔΕΥΣΗ ΚΑΙ ΤΟ ΠΑΕΙ ΣΤΗΝ  ΖΥΓΑΡΙΑ
Command2.Enabled = False
Βαλε_Νερο2_Click   ' ΒΑΖΕΙ ΝΕΡΟ ΑΡΑΙΩΣΗΣ ΜΕΧΡΙ ΤΕΛΙΚΗΣ ΠΕΡΙΕΚΤΙΚΟΤΗΤΟΣ
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
        temp% = SendAT6400Block(768, argum, 0)
 End Sub
Function Zygis0(x As Single)
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
baros_pot_min = 10000  'min βαρος μπουκαλιου με μαγνήτη και χωρίς τάπα
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
  rob_feedback (59)
End Sub

Private Sub Text1_Change()
        If system_ready = 1 Then rezero_Click
         nero2 = 0: '  curq = 2000
         secwait (2)
         ner_stox = CurQ / periek - CurQ
         nero2_stox = ner_stox - Val(nero1)
         nero2_stox = nero2_stox - 2000
       
     If Val(nero1) > 0 Then nero2_count_stox = nero1_count * nero2_stox / nero1
     
                        '1o γεμισμα
              cmd$ = nero_on:   RobSend (cmd$)
                nero2_count = 0
         Do
            nero2_count = nero2_count + 1
            If system_ready = 1 Then nero2 = ZYGIS3()
            nerot = Val(nero1) + Val(nero2)
            If nerot > 630000 Then Exit Do
            fin_con = CurQ / (nerot + CurQ): fin_con1 = Format(fin_con, "0.0000 %")
       Loop Until nerot >= nero2_stox
       
        cmd$ = nero_off
        RobSend (cmd$)
        secwait (2)
         If system_ready = 1 Then nero2 = ZYGIS3()
                 nerot = Val(nero1) + Val(nero2)
                 
             '    2ο γεμισμα
                 
                 ner_yp = nero2_stox - nero2
                 nero2_count = nero2_count * ner_yp / nero2
      If ner_yp > 20 Then
         cmd$ = nero_on:   RobSend (cmd$)
         Do
         ner = ner + 1
             If system_ready = 1 Then nero2 = ZYGIS3()
            nerot = Val(nero1) + Val(nero2)
            If nerot > 630000 Then Exit Do
            fin_con = CurQ / (nerot + CurQ): fin_con1 = Format(fin_con, "0.0000 %")
          Loop Until ner >= nero2_count
       End If
          Label18.Visible = True
          secwait (2)
      If system_ready = 1 Then nero2 = ZYGIS3()
     nerot = Val(nero1) + Val(nero2)
     fin_con = CurQ / (nerot + CurQ): fin_con1 = Format(fin_con, "0.0000 %")
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
     If tim_count / 10 = Int(tim_count / 10) Then GoSub mix_stop
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
       secwait (1)
      cmd$ = "!" + mixer_on + ":"
       RobSend (cmd$)
       Return
End Sub

Public Sub rearange(argum)
While FindOutput(Val(Ready_Bit)) = 0:  Wend
secwait (1)
 
If argum = 2 Then 'στείλε πίσω μπουκ & ποτήρι
   'cmd$ = "!d,,(var93):!go,,1:"
  ' RobSend (cmd$)
  ' secwait (2)
   'cmd$ = "!d,(var62):!go,1:"
  '  RobSend (cmd$)
    ' secwait (2)
   cmd$ = "!var12=" + Str(Buk + 2000) + ":" + "!var36=" + Trim(Str(x_buk)) + ":" + "!var37=" + Trim(Str(y_buk)) + ":"
    RobSend (cmd$)
    secwait (1)
  While FindOutput(Val(Ready_Bit)) = 0:  Wend
 secwait (1)
End If
  While FindOutput(Val(Ready_Bit)) = 0:  Wend
  secwait (1)
   cmd$ = "!var12=" + Str(Pot + 4000) + ":" + "!var24=" + Trim(Str(x_pot)) + ":" + "!var77=" + Trim(Str(y_pot)) + ":" + "!var89=" + Trim(Str(z_pot)) + ":"
   RobSend (cmd$)
   secwait (1)
While FindOutput(Val(Ready_Bit)) = 0:  Wend
secwait (1)
End Sub


Public Sub repair_powder_factors()
Set MyDB = Workspaces(0).OpenDatabase("c:\talos\skon_tb.mdb")
Set r = MyDB.OpenRecordset("Σκονες")
Do Until r.EOF
 If Buk = r("Νο_Μπουκαλιου") Then
    r.Edit
      r("Datum") = Now()
      r("Factor_A") = Ca_Factor
      r("Ximiko") = XIMIKO
    r.Update
      r.Close
     Exit Sub
End If
  r.MoveNext
  Loop
  
     r.AddNew
       r("Νο_Μπουκαλιου") = Buk
       r("Ximiko") = XIMIKO
       r("Factor_A") = 35 * (1 + percentage_value)
      r("Datum") = Now()
    r.Update
    r.Close
End Sub


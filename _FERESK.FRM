VERSION 5.00
Begin VB.Form fereSK 
   BackColor       =   &H00FF0000&
   ClientHeight    =   7080
   ClientLeft      =   1515
   ClientTop       =   420
   ClientWidth     =   9480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7080
   ScaleWidth      =   9480
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check1 
      BackColor       =   &H000080FF&
      Caption         =   "Με αραίωση"
      Height          =   255
      Left            =   120
      TabIndex        =   338
      Top             =   6480
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Continue Solution Making"
      Height          =   480
      Left            =   120
      TabIndex        =   337
      Top             =   5715
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Κατασκευή Διαλυμάτων"
      Enabled         =   0   'False
      Height          =   465
      Left            =   100
      TabIndex        =   5
      Top             =   5210
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ΕΞΟΔΟΣ"
      Height          =   465
      Left            =   2700
      TabIndex        =   4
      Top             =   5210
      Width           =   1155
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   1185
      Top             =   7275
   End
   Begin VB.Data XIMITECH 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "c:\talos\recipies.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2745
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "XIMITECH"
      Top             =   6645
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "299"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   299
      Left            =   8610
      TabIndex        =   336
      Top             =   -65
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "298"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   298
      Left            =   8145
      TabIndex        =   335
      Top             =   -65
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "297"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   297
      Left            =   7665
      TabIndex        =   334
      Top             =   -65
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "276"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   296
      Left            =   7200
      TabIndex        =   333
      Top             =   -65
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "295"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   295
      Left            =   6735
      TabIndex        =   332
      Top             =   -65
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "294"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   294
      Left            =   6255
      TabIndex        =   331
      Top             =   -65
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "293"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   293
      Left            =   5790
      TabIndex        =   330
      Top             =   -65
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "292"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   292
      Left            =   5325
      TabIndex        =   329
      Top             =   -65
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "290"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   290
      Left            =   4380
      TabIndex        =   328
      Top             =   -65
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "289"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   289
      Left            =   3915
      TabIndex        =   327
      Top             =   -65
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "288"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   288
      Left            =   3435
      TabIndex        =   326
      Top             =   -65
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "287"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   287
      Left            =   2970
      TabIndex        =   325
      Top             =   -65
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "286"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   286
      Left            =   2505
      TabIndex        =   324
      Top             =   -65
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "285"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   285
      Left            =   2025
      TabIndex        =   323
      Top             =   -65
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "284"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   284
      Left            =   1560
      TabIndex        =   322
      Top             =   -65
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "282"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   282
      Left            =   615
      TabIndex        =   321
      Top             =   -65
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "281"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   281
      Left            =   150
      TabIndex        =   320
      Top             =   -65
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "300"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   300
      Left            =   9075
      TabIndex        =   319
      Top             =   -65
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "291"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   291
      Left            =   4845
      TabIndex        =   318
      Top             =   -65
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "283"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   283
      Left            =   1095
      TabIndex        =   317
      Top             =   -65
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "30"
      Height          =   375
      Index           =   30
      Left            =   8610
      TabIndex        =   316
      Top             =   6465
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "29"
      Height          =   375
      Index           =   29
      Left            =   8145
      TabIndex        =   315
      Top             =   6465
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "28"
      Height          =   375
      Index           =   28
      Left            =   7665
      TabIndex        =   314
      Top             =   6465
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "20"
      Height          =   375
      Index           =   20
      Left            =   8610
      TabIndex        =   313
      Top             =   6915
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "19"
      Height          =   375
      Index           =   19
      Left            =   8145
      TabIndex        =   312
      Top             =   6915
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "18"
      Height          =   375
      Index           =   18
      Left            =   7665
      TabIndex        =   311
      Top             =   6915
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "10"
      Height          =   375
      Index           =   10
      Left            =   8610
      TabIndex        =   310
      Top             =   7365
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "9"
      Height          =   375
      Index           =   9
      Left            =   8145
      TabIndex        =   309
      Top             =   7365
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "8"
      Height          =   375
      Index           =   8
      Left            =   7665
      TabIndex        =   308
      Top             =   7365
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "13"
      Height          =   375
      Index           =   13
      Left            =   5325
      TabIndex        =   307
      Top             =   6915
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "1"
      Height          =   375
      Index           =   1
      Left            =   4380
      TabIndex        =   306
      Top             =   7365
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "2"
      Height          =   375
      Index           =   2
      Left            =   4845
      TabIndex        =   305
      Top             =   7365
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "3"
      Height          =   375
      Index           =   3
      Left            =   5325
      TabIndex        =   304
      Top             =   7365
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "4"
      Height          =   375
      Index           =   4
      Left            =   5790
      TabIndex        =   303
      Top             =   7365
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "5"
      Height          =   375
      Index           =   5
      Left            =   6255
      TabIndex        =   302
      Top             =   7365
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "6"
      Height          =   375
      Index           =   6
      Left            =   6735
      TabIndex        =   301
      Top             =   7365
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "7"
      Height          =   375
      Index           =   7
      Left            =   7200
      TabIndex        =   300
      Top             =   7365
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "11"
      Height          =   375
      Index           =   11
      Left            =   4380
      TabIndex        =   299
      Top             =   6915
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "12"
      Height          =   375
      Index           =   12
      Left            =   4845
      TabIndex        =   298
      Top             =   6915
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "14"
      Height          =   375
      Index           =   14
      Left            =   5790
      TabIndex        =   297
      Top             =   6915
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "15"
      Height          =   375
      Index           =   15
      Left            =   6255
      TabIndex        =   296
      Top             =   6915
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "16"
      Height          =   375
      Index           =   16
      Left            =   6735
      TabIndex        =   295
      Top             =   6915
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "17"
      Height          =   375
      Index           =   17
      Left            =   7200
      TabIndex        =   294
      Top             =   6915
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "21"
      Height          =   375
      Index           =   21
      Left            =   4380
      TabIndex        =   293
      Top             =   6465
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "22"
      Height          =   375
      Index           =   22
      Left            =   4845
      TabIndex        =   292
      Top             =   6465
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "23"
      Height          =   375
      Index           =   23
      Left            =   5325
      TabIndex        =   291
      Top             =   6465
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "24"
      Height          =   375
      Index           =   24
      Left            =   5790
      TabIndex        =   290
      Top             =   6465
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "25"
      Height          =   375
      Index           =   25
      Left            =   6255
      TabIndex        =   289
      Top             =   6465
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "26"
      Height          =   375
      Index           =   26
      Left            =   6735
      TabIndex        =   288
      Top             =   6465
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "27"
      Height          =   375
      Index           =   27
      Left            =   7200
      TabIndex        =   287
      Top             =   6465
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "263"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   263
      Left            =   1095
      TabIndex        =   286
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "271"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   271
      Left            =   4845
      TabIndex        =   285
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "280"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   280
      Left            =   9075
      TabIndex        =   284
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "261"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   261
      Left            =   150
      TabIndex        =   283
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "262"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   262
      Left            =   615
      TabIndex        =   282
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "264"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   264
      Left            =   1560
      TabIndex        =   281
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "265"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   265
      Left            =   2025
      TabIndex        =   280
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "266"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   266
      Left            =   2505
      TabIndex        =   279
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "267"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   267
      Left            =   2970
      TabIndex        =   278
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "268"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   268
      Left            =   3435
      TabIndex        =   277
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "269"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   269
      Left            =   3915
      TabIndex        =   276
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "270"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   270
      Left            =   4380
      TabIndex        =   275
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "272"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   272
      Left            =   5325
      TabIndex        =   274
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "273"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   273
      Left            =   5790
      TabIndex        =   273
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "274"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   274
      Left            =   6255
      TabIndex        =   272
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "275"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   275
      Left            =   6735
      TabIndex        =   271
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "276"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   276
      Left            =   7200
      TabIndex        =   270
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "277"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   277
      Left            =   7665
      TabIndex        =   269
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "278"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   278
      Left            =   8145
      TabIndex        =   268
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "279"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   279
      Left            =   8610
      TabIndex        =   267
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "20"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   20
      Left            =   9075
      TabIndex        =   266
      Top             =   6075
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "40"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   40
      Left            =   9075
      TabIndex        =   265
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "60"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   60
      Left            =   9075
      TabIndex        =   264
      Top             =   5190
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "80"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   80
      Left            =   9075
      TabIndex        =   263
      Top             =   4755
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "100"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   100
      Left            =   9075
      TabIndex        =   262
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "120"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   120
      Left            =   9075
      TabIndex        =   261
      Top             =   3870
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "140"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   140
      Left            =   9075
      TabIndex        =   260
      Top             =   3435
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "160"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   160
      Left            =   9075
      TabIndex        =   259
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "180"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   180
      Left            =   9075
      TabIndex        =   258
      Top             =   2550
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "200"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   200
      Left            =   9075
      TabIndex        =   257
      Top             =   2115
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "220"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   220
      Left            =   9075
      TabIndex        =   256
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "240"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   240
      Left            =   9075
      TabIndex        =   255
      Top             =   1230
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "260"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   260
      Left            =   9075
      TabIndex        =   254
      Top             =   795
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "252"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   72
      Left            =   5325
      TabIndex        =   253
      Top             =   4755
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "251"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   71
      Left            =   4845
      TabIndex        =   252
      Top             =   4755
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   70
      Left            =   4380
      TabIndex        =   251
      Top             =   4755
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "249"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   69
      Left            =   3915
      TabIndex        =   250
      Top             =   4755
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "248"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   68
      Left            =   3435
      TabIndex        =   249
      Top             =   4755
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "247"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   67
      Left            =   2970
      TabIndex        =   248
      Top             =   4755
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   66
      Left            =   2505
      TabIndex        =   247
      Top             =   4755
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   65
      Left            =   2025
      TabIndex        =   246
      Top             =   4755
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "252"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   52
      Left            =   5325
      TabIndex        =   245
      Top             =   5190
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "251"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   51
      Left            =   4845
      TabIndex        =   244
      Top             =   5190
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   50
      Left            =   4380
      TabIndex        =   243
      Top             =   5190
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "249"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   49
      Left            =   3915
      TabIndex        =   242
      Top             =   5190
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "248"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   48
      Left            =   3435
      TabIndex        =   241
      Top             =   5190
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "247"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   47
      Left            =   2970
      TabIndex        =   240
      Top             =   5190
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   46
      Left            =   2505
      TabIndex        =   239
      Top             =   5190
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   45
      Left            =   2025
      TabIndex        =   238
      Top             =   5190
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "12"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   12
      Left            =   5325
      TabIndex        =   237
      Top             =   6075
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "11"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   11
      Left            =   4845
      TabIndex        =   236
      Top             =   6075
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "10"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   10
      Left            =   4380
      TabIndex        =   235
      Top             =   6075
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "9"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   9
      Left            =   3915
      TabIndex        =   234
      Top             =   6075
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "8"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   8
      Left            =   3435
      TabIndex        =   233
      Top             =   6075
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "7"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   7
      Left            =   2970
      TabIndex        =   232
      Top             =   6075
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "6"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   6
      Left            =   2505
      TabIndex        =   231
      Top             =   6075
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "5"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   2025
      TabIndex        =   230
      Top             =   6075
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "252"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   32
      Left            =   5325
      TabIndex        =   229
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "251"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   31
      Left            =   4845
      TabIndex        =   228
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   30
      Left            =   4380
      TabIndex        =   227
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "249"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   29
      Left            =   3915
      TabIndex        =   226
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "248"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   28
      Left            =   3435
      TabIndex        =   225
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "247"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   27
      Left            =   2970
      TabIndex        =   224
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   26
      Left            =   2505
      TabIndex        =   223
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   25
      Left            =   2025
      TabIndex        =   222
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "259"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   259
      Left            =   8610
      TabIndex        =   221
      Top             =   795
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   239
      Left            =   8610
      TabIndex        =   220
      Top             =   1230
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   219
      Left            =   8610
      TabIndex        =   219
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   199
      Left            =   8610
      TabIndex        =   218
      Top             =   2115
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   179
      Left            =   8610
      TabIndex        =   217
      Top             =   2550
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   159
      Left            =   8610
      TabIndex        =   216
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   139
      Left            =   8610
      TabIndex        =   215
      Top             =   3435
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   119
      Left            =   8610
      TabIndex        =   214
      Top             =   3870
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "99"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   99
      Left            =   8610
      TabIndex        =   213
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   79
      Left            =   8610
      TabIndex        =   212
      Top             =   4755
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   59
      Left            =   8610
      TabIndex        =   211
      Top             =   5190
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   39
      Left            =   8610
      TabIndex        =   210
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "19"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   19
      Left            =   8610
      TabIndex        =   209
      Top             =   6075
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "258"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   258
      Left            =   8145
      TabIndex        =   208
      Top             =   795
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   238
      Left            =   8145
      TabIndex        =   207
      Top             =   1230
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   218
      Left            =   8145
      TabIndex        =   206
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   198
      Left            =   8145
      TabIndex        =   205
      Top             =   2115
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   178
      Left            =   8145
      TabIndex        =   204
      Top             =   2550
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   158
      Left            =   8145
      TabIndex        =   203
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   138
      Left            =   8145
      TabIndex        =   202
      Top             =   3435
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   118
      Left            =   8145
      TabIndex        =   201
      Top             =   3870
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "98"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   98
      Left            =   8145
      TabIndex        =   200
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   78
      Left            =   8145
      TabIndex        =   199
      Top             =   4755
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   58
      Left            =   8145
      TabIndex        =   198
      Top             =   5190
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   38
      Left            =   8145
      TabIndex        =   197
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "18"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   18
      Left            =   8145
      TabIndex        =   196
      Top             =   6075
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "257"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   257
      Left            =   7665
      TabIndex        =   195
      Top             =   795
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   237
      Left            =   7665
      TabIndex        =   194
      Top             =   1230
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   217
      Left            =   7665
      TabIndex        =   193
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   197
      Left            =   7665
      TabIndex        =   192
      Top             =   2115
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   177
      Left            =   7665
      TabIndex        =   191
      Top             =   2550
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   157
      Left            =   7665
      TabIndex        =   190
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   137
      Left            =   7665
      TabIndex        =   189
      Top             =   3435
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   117
      Left            =   7665
      TabIndex        =   188
      Top             =   3870
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "97"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   97
      Left            =   7665
      TabIndex        =   187
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   77
      Left            =   7665
      TabIndex        =   186
      Top             =   4755
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   57
      Left            =   7665
      TabIndex        =   185
      Top             =   5190
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   37
      Left            =   7665
      TabIndex        =   184
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "17"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   17
      Left            =   7665
      TabIndex        =   183
      Top             =   6075
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "256"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   256
      Left            =   7200
      TabIndex        =   182
      Top             =   795
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   236
      Left            =   7200
      TabIndex        =   181
      Top             =   1230
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   216
      Left            =   7200
      TabIndex        =   180
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   196
      Left            =   7200
      TabIndex        =   179
      Top             =   2115
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   176
      Left            =   7200
      TabIndex        =   178
      Top             =   2550
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   156
      Left            =   7200
      TabIndex        =   177
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   136
      Left            =   7200
      TabIndex        =   176
      Top             =   3435
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   116
      Left            =   7200
      TabIndex        =   175
      Top             =   3870
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "96"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   96
      Left            =   7200
      TabIndex        =   174
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   76
      Left            =   7200
      TabIndex        =   173
      Top             =   4755
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   56
      Left            =   7200
      TabIndex        =   172
      Top             =   5190
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   36
      Left            =   7200
      TabIndex        =   171
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "16"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   16
      Left            =   7200
      TabIndex        =   170
      Top             =   6075
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "255"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   255
      Left            =   6735
      TabIndex        =   169
      Top             =   795
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   235
      Left            =   6735
      TabIndex        =   168
      Top             =   1230
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   215
      Left            =   6735
      TabIndex        =   167
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   195
      Left            =   6735
      TabIndex        =   166
      Top             =   2115
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   175
      Left            =   6735
      TabIndex        =   165
      Top             =   2550
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   155
      Left            =   6735
      TabIndex        =   164
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   135
      Left            =   6735
      TabIndex        =   163
      Top             =   3435
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   115
      Left            =   6735
      TabIndex        =   162
      Top             =   3870
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "95"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   95
      Left            =   6735
      TabIndex        =   161
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "254"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   254
      Left            =   6255
      TabIndex        =   160
      Top             =   795
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   234
      Left            =   6255
      TabIndex        =   159
      Top             =   1230
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   214
      Left            =   6255
      TabIndex        =   158
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   194
      Left            =   6255
      TabIndex        =   157
      Top             =   2115
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   174
      Left            =   6255
      TabIndex        =   156
      Top             =   2550
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   154
      Left            =   6255
      TabIndex        =   155
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   134
      Left            =   6255
      TabIndex        =   154
      Top             =   3435
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   114
      Left            =   6255
      TabIndex        =   153
      Top             =   3870
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "94"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   94
      Left            =   6255
      TabIndex        =   152
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "253"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   253
      Left            =   5790
      TabIndex        =   151
      Top             =   795
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   233
      Left            =   5790
      TabIndex        =   150
      Top             =   1230
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   213
      Left            =   5790
      TabIndex        =   149
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   193
      Left            =   5790
      TabIndex        =   148
      Top             =   2115
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   173
      Left            =   5790
      TabIndex        =   147
      Top             =   2550
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   153
      Left            =   5790
      TabIndex        =   146
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   133
      Left            =   5790
      TabIndex        =   145
      Top             =   3435
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   113
      Left            =   5790
      TabIndex        =   144
      Top             =   3870
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "93"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   93
      Left            =   5790
      TabIndex        =   143
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "252"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   252
      Left            =   5325
      TabIndex        =   142
      Top             =   795
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   232
      Left            =   5325
      TabIndex        =   141
      Top             =   1230
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   212
      Left            =   5325
      TabIndex        =   140
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   192
      Left            =   5325
      TabIndex        =   139
      Top             =   2115
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   172
      Left            =   5325
      TabIndex        =   138
      Top             =   2550
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   152
      Left            =   5325
      TabIndex        =   137
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   132
      Left            =   5325
      TabIndex        =   136
      Top             =   3435
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   112
      Left            =   5325
      TabIndex        =   135
      Top             =   3870
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "92"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   92
      Left            =   5325
      TabIndex        =   134
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "251"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   251
      Left            =   4845
      TabIndex        =   133
      Top             =   795
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   231
      Left            =   4845
      TabIndex        =   132
      Top             =   1230
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   211
      Left            =   4845
      TabIndex        =   131
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   191
      Left            =   4845
      TabIndex        =   130
      Top             =   2115
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   171
      Left            =   4845
      TabIndex        =   129
      Top             =   2550
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   151
      Left            =   4845
      TabIndex        =   128
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   131
      Left            =   4845
      TabIndex        =   127
      Top             =   3435
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   111
      Left            =   4845
      TabIndex        =   126
      Top             =   3870
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "91"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   91
      Left            =   4845
      TabIndex        =   125
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   250
      Left            =   4380
      TabIndex        =   124
      Top             =   795
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   230
      Left            =   4380
      TabIndex        =   123
      Top             =   1230
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   210
      Left            =   4380
      TabIndex        =   122
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   190
      Left            =   4380
      TabIndex        =   121
      Top             =   2115
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   170
      Left            =   4380
      TabIndex        =   120
      Top             =   2550
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   150
      Left            =   4380
      TabIndex        =   119
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   130
      Left            =   4380
      TabIndex        =   118
      Top             =   3435
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   110
      Left            =   4380
      TabIndex        =   117
      Top             =   3870
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "90"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   90
      Left            =   4380
      TabIndex        =   116
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "249"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   249
      Left            =   3915
      TabIndex        =   115
      Top             =   795
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   229
      Left            =   3915
      TabIndex        =   114
      Top             =   1230
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   209
      Left            =   3915
      TabIndex        =   113
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   189
      Left            =   3915
      TabIndex        =   112
      Top             =   2115
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   169
      Left            =   3915
      TabIndex        =   111
      Top             =   2550
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   149
      Left            =   3915
      TabIndex        =   110
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   129
      Left            =   3915
      TabIndex        =   109
      Top             =   3435
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   109
      Left            =   3915
      TabIndex        =   108
      Top             =   3870
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "89"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   89
      Left            =   3915
      TabIndex        =   107
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "248"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   248
      Left            =   3435
      TabIndex        =   106
      Top             =   795
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   228
      Left            =   3435
      TabIndex        =   105
      Top             =   1230
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   208
      Left            =   3435
      TabIndex        =   104
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   188
      Left            =   3435
      TabIndex        =   103
      Top             =   2115
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   168
      Left            =   3435
      TabIndex        =   102
      Top             =   2550
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   148
      Left            =   3435
      TabIndex        =   101
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   128
      Left            =   3435
      TabIndex        =   100
      Top             =   3435
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   108
      Left            =   3435
      TabIndex        =   99
      Top             =   3870
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "88"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   88
      Left            =   3435
      TabIndex        =   98
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "247"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   247
      Left            =   2970
      TabIndex        =   97
      Top             =   795
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   227
      Left            =   2970
      TabIndex        =   96
      Top             =   1230
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   207
      Left            =   2970
      TabIndex        =   95
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   187
      Left            =   2970
      TabIndex        =   94
      Top             =   2115
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   167
      Left            =   2970
      TabIndex        =   93
      Top             =   2550
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   147
      Left            =   2970
      TabIndex        =   92
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   127
      Left            =   2970
      TabIndex        =   91
      Top             =   3435
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   107
      Left            =   2970
      TabIndex        =   90
      Top             =   3870
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "87"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   87
      Left            =   2970
      TabIndex        =   89
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "75"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   75
      Left            =   6735
      TabIndex        =   88
      Top             =   4755
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "55"
      ForeColor       =   &H0000002F&
      Height          =   375
      Index           =   55
      Left            =   6735
      TabIndex        =   87
      Top             =   5190
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "35"
      ForeColor       =   &H0000001B&
      Height          =   375
      Index           =   35
      Left            =   6735
      TabIndex        =   86
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "15"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   15
      Left            =   6735
      TabIndex        =   85
      Top             =   6075
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   246
      Left            =   2505
      TabIndex        =   84
      Top             =   795
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   226
      Left            =   2505
      TabIndex        =   83
      Top             =   1230
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   206
      Left            =   2505
      TabIndex        =   82
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   186
      Left            =   2505
      TabIndex        =   81
      Top             =   2115
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   166
      Left            =   2505
      TabIndex        =   80
      Top             =   2550
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   146
      Left            =   2505
      TabIndex        =   79
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   126
      Left            =   2505
      TabIndex        =   78
      Top             =   3435
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   106
      Left            =   2505
      TabIndex        =   77
      Top             =   3870
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "86"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   86
      Left            =   2505
      TabIndex        =   76
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "74"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   74
      Left            =   6255
      TabIndex        =   75
      Top             =   4755
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "54"
      ForeColor       =   &H0000002E&
      Height          =   375
      Index           =   54
      Left            =   6255
      TabIndex        =   74
      Top             =   5190
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "34"
      ForeColor       =   &H0000001A&
      Height          =   375
      Index           =   34
      Left            =   6255
      TabIndex        =   73
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "14"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   14
      Left            =   6255
      TabIndex        =   72
      Top             =   6075
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   245
      Left            =   2025
      TabIndex        =   71
      Top             =   795
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   225
      Left            =   2025
      TabIndex        =   70
      Top             =   1230
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   205
      Left            =   2025
      TabIndex        =   69
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   185
      Left            =   2025
      TabIndex        =   68
      Top             =   2115
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   165
      Left            =   2025
      TabIndex        =   67
      Top             =   2550
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   145
      Left            =   2025
      TabIndex        =   66
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   125
      Left            =   2025
      TabIndex        =   65
      Top             =   3435
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   105
      Left            =   2025
      TabIndex        =   64
      Top             =   3870
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "85"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   85
      Left            =   2025
      TabIndex        =   63
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "73"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   73
      Left            =   5790
      TabIndex        =   62
      Top             =   4755
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "53"
      ForeColor       =   &H0000002D&
      Height          =   375
      Index           =   53
      Left            =   5790
      TabIndex        =   61
      Top             =   5190
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "33"
      ForeColor       =   &H00000019&
      Height          =   375
      Index           =   33
      Left            =   5790
      TabIndex        =   60
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "13"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   13
      Left            =   5790
      TabIndex        =   59
      Top             =   6075
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   244
      Left            =   1560
      TabIndex        =   58
      Top             =   795
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   224
      Left            =   1560
      TabIndex        =   57
      Top             =   1230
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   204
      Left            =   1560
      TabIndex        =   56
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   184
      Left            =   1560
      TabIndex        =   55
      Top             =   2115
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   164
      Left            =   1560
      TabIndex        =   54
      Top             =   2550
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   144
      Left            =   1560
      TabIndex        =   53
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   124
      Left            =   1560
      TabIndex        =   52
      Top             =   3435
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   104
      Left            =   1560
      TabIndex        =   51
      Top             =   3870
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "84"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   84
      Left            =   1560
      TabIndex        =   50
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000040&
      Height          =   375
      Index           =   64
      Left            =   1560
      TabIndex        =   49
      Top             =   4755
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H0000002C&
      Height          =   375
      Index           =   44
      Left            =   1560
      TabIndex        =   48
      Top             =   5190
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000018&
      Height          =   375
      Index           =   24
      Left            =   1560
      TabIndex        =   47
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "4"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   1560
      TabIndex        =   46
      Top             =   6075
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   243
      Left            =   1095
      TabIndex        =   45
      Top             =   795
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   223
      Left            =   1095
      TabIndex        =   44
      Top             =   1230
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   203
      Left            =   1095
      TabIndex        =   43
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   183
      Left            =   1095
      TabIndex        =   42
      Top             =   2115
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   163
      Left            =   1095
      TabIndex        =   41
      Top             =   2550
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   143
      Left            =   1095
      TabIndex        =   40
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   123
      Left            =   1095
      TabIndex        =   39
      Top             =   3435
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   103
      Left            =   1095
      TabIndex        =   38
      Top             =   3870
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "83"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   83
      Left            =   1095
      TabIndex        =   37
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H0000003F&
      Height          =   375
      Index           =   63
      Left            =   1095
      TabIndex        =   36
      Top             =   4755
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H0000002B&
      Height          =   375
      Index           =   43
      Left            =   1095
      TabIndex        =   35
      Top             =   5190
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000017&
      Height          =   375
      Index           =   23
      Left            =   1095
      TabIndex        =   34
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   1095
      TabIndex        =   33
      Top             =   6075
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   242
      Left            =   615
      TabIndex        =   32
      Top             =   795
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   222
      Left            =   615
      TabIndex        =   31
      Top             =   1230
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   202
      Left            =   615
      TabIndex        =   30
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   182
      Left            =   615
      TabIndex        =   29
      Top             =   2115
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   162
      Left            =   615
      TabIndex        =   28
      Top             =   2550
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   142
      Left            =   615
      TabIndex        =   27
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   122
      Left            =   615
      TabIndex        =   26
      Top             =   3435
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   102
      Left            =   615
      TabIndex        =   25
      Top             =   3870
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "82"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   82
      Left            =   615
      TabIndex        =   24
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H0000003E&
      Height          =   375
      Index           =   62
      Left            =   615
      TabIndex        =   23
      Top             =   4755
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H0000002A&
      Height          =   375
      Index           =   42
      Left            =   615
      TabIndex        =   22
      Top             =   5190
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000016&
      Height          =   375
      Index           =   22
      Left            =   615
      TabIndex        =   21
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2"
      ForeColor       =   &H00000002&
      Height          =   375
      Index           =   2
      Left            =   615
      TabIndex        =   20
      Top             =   6075
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "241"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   241
      Left            =   150
      TabIndex        =   19
      Top             =   795
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "221"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   221
      Left            =   150
      TabIndex        =   18
      Top             =   1230
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "201"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   201
      Left            =   150
      TabIndex        =   17
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "181"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   181
      Left            =   150
      TabIndex        =   16
      Top             =   2115
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "161"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   161
      Left            =   150
      TabIndex        =   15
      Top             =   2550
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "141"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   141
      Left            =   150
      TabIndex        =   14
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "121"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   121
      Left            =   150
      TabIndex        =   13
      Top             =   3435
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "101"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   101
      Left            =   150
      TabIndex        =   12
      Top             =   3870
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "81"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   81
      Left            =   150
      TabIndex        =   11
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "61"
      ForeColor       =   &H0000003D&
      Height          =   375
      Index           =   61
      Left            =   150
      TabIndex        =   10
      Top             =   4755
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "41"
      ForeColor       =   &H00000029&
      Height          =   375
      Index           =   41
      Left            =   150
      TabIndex        =   9
      Top             =   5190
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "21"
      ForeColor       =   &H00000015&
      Height          =   375
      Index           =   21
      Left            =   150
      TabIndex        =   8
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1"
      ForeColor       =   &H00000001&
      Height          =   375
      Index           =   1
      Left            =   150
      TabIndex        =   7
      Top             =   6075
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   180
      Index           =   0
      Left            =   9090
      TabIndex        =   6
      Top             =   8280
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label SXOLIA2 
      Height          =   255
      Left            =   165
      TabIndex        =   3
      Top             =   5970
      Width           =   3795
   End
   Begin VB.Label sxolia1 
      Height          =   255
      Left            =   165
      TabIndex        =   2
      Top             =   6225
      Width           =   3795
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   375
      Index           =   0
      Left            =   -495
      TabIndex        =   1
      Top             =   8205
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   255
      TabIndex        =   0
      Top             =   6555
      Visible         =   0   'False
      Width           =   1875
   End
End
Attribute VB_Name = "fereSK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pathma_Koympioy
Dim OldColor(1 To Ar_Mpoykal) As Single   ' Public array variable.
Dim kodikos(1 To Ar_Mpoykal)     ' Public array variable.
Dim Perigrafh(1 To Ar_Mpoykal)   ' Public array variable.
Dim MHTRIKO(1 To Ar_Mpoykal)

Dim LHJH(1 To Ar_Mpoykal)   ' Public array variable.
Dim Ypoloipo(1 To Ar_Mpoykal)   ' Public array variable.
Dim periek(1 To Ar_Mpoykal)   ' Public array variable.
Dim MHTRIKO2(1 To Ar_Mpoykal)

Const Kafe = &H808080 ' &H404040
Const Keramidi = &H80C0FF  '&H80C0FF
Const Prasino = &H80FF&
Const kokkino = &HFF&
Const Portokali = &H80FF&
Const ASPRO = &HFFFFFF
Const ROZ = &HFF80FF
Function val2(n)
       s = IIf(InStr(n, "%") = 0, 6, InStr(n, "%"))
      If IsNull(s) Then
         val2 = 0
    Else
       val2 = Val(left(n, s - 1))
    End If
End Function

Function fernei(Index As Integer)
   

End Function

Private Sub back_Click()


End Sub

Private Sub DBCombo1_Click(Area As Integer)
 Dim x
   x = DBCombo1.SelectedItem
   ximitech.Recordset.Bookmark = x
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Command2_Click()

Dim mydb As Database, MyTable As Recordset
Dim SQL As String, mADD_ster
' ANANEΩΣH    ΔIAΛYMATΩN
Dim kod6400
  
' Pathma_Koympioy = 2
  
If Label3(1).BackColor = &HFFFFFF Then
    Exit Sub
End If
    Command2.Enabled = False
If system_ready = 1 Then PreHeating

  dum = begin_talos(0)
  If dum = 0 Then
        response = MsgBox(mL_Res(547), 64, "Talos")
        ' 547 "Πρόβλημα επικοινωνίας Robot-Προγράμματος "
        Exit Sub
  End If

 'Do While Pause_Check() = 1
  '     MsgBox mL_Res(548), , "Talos.548"  ' Define message.
       ' 548  "Για να συνεχίσει το πρόγραμμα γυρίστε τον διακόπτη PAUSE στην θέση 1 και πατήστε ΟΚ."
'Loop

 'Do While Is_Porta() = "α"
  ' MsgBox mL_Res(549), , "Talos"
   ' 549 "Για να συνεχίσει το πρόγραμμα κλείστε την πόρτα και πατήστε ΟΚ."  ' Define message.
'Loop
Timer1.Enabled = True
Telos_epilogon = 1 ' αν γίνει 1 τότε σταματάνε οι επιλογές μπουκαλιών

' ΠΑΤΑΕΙ ΠΡΩΤΗ ΦΟΡΑ ΤΟ ΚΟΥΜΠΙ
If Pathma_Koympioy = 0 Then
  MsgBox mL_Res(550), , "Talos"
  '550 "Eίναι οι θέσεις ποτηριών ΕΛΕΥΘΕΡΕΣ ;"
  Command2.Enabled = False
  Command1.Enabled = False
  'front.Enabled = False
  For k = 1 To Ar_Mpoykal
      Label1(k).Enabled = False
  Next
  
  For k = 1 To ar_real_pothr
    If Label3(k).BackColor = kokkino Then ' αν δεν είναι κόκκινη θέση φέρτο εδώ
       
       
       cmd$ = "!var12=" + str(6000 + Val(Label3(k).Caption)) + ":!var13=" + str(k) + ":"
       dum = wait_to_understand(cmd$)
       dum = Wait_Robot(0)
       
       
       Label3(k).BackColor = &HFFFF80
       Label1(Val(Label3(k).Caption)).BackColor = &HFFFFFF ' ασπρο
    End If
  Next
  cmd$ = "!var12=82:"   'Park
  dum = wait_to_understand(cmd$)
  dum = Wait_Robot(0)
  
  
  
  MilSec 3000

  
  
   Command2.Caption = mL_CapRes(539) ' "ΕΝΑΡΞΗ ΠΑΡΑΣΚΕΥΗΣ ΔΙΑΛΥΜΑΤΩΝ"
   Command2.Enabled = True
   Pathma_Koympioy = 1
   Exit Sub
 End If
 
' ΠΑΤΑΕΙ  ΔΕΥΤΕΡΗ ΦΟΡΑ ΤΟ ΚΟΥΜΠΙ
If Pathma_Koympioy = 1 Then
     Command2.Caption = mL_CapRes(536) '"Επιστροφή Διαλυμάτων"
     Command2.Enabled = False
           ximitech.Recordset.Index = "KOD"
           Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\SKON_TB.MDB")
           ' Open table.
           Set MyTable = mydb.OpenRecordset("entoles", dbOpenTable)
           
    '            SQL = "DELETE DISTINCTROW entoles.* FROM entoles;"
   '         MyDB.Execute SQL, dbFailOnError
           
     For k = 1 To ar_real_pothr
         If Label3(k).BackColor = &HFFFF80 Then
            ximitech.Recordset.Seek "=", kodikos(Label3(k).Caption)
            MyTable.AddNew  ' Create new record.
            MyTable("nerostox") = ximitech.Recordset("nerostox") * 1000 'ποσο διαλυμα θα παρασκευασθεί σε mgr
            MyTable("periek") = val2(ximitech.Recordset("MORFH")) / 100 'περιεκτικότης στόχου
            
            If MyTable("periek") <= 0 Then
            ' MyTable("periek") = val2(XIMITECH.Recordset("MORFH")) / 100 'περιεκτικότης στόχου
               MsgBox "Concentration = 0 . Local settings->Numbers->decimal symbol-> . "
            End If
            
            
            MyTable("buk") = Val(ximitech.Recordset("ADDR_old"))    'Νο μπουκαλιού σκόνης
            MyTable("pot") = k         'Νο ποτηριού (μπουκαλιού διαλύματος
            MyTable("mixtim") = ximitech.Recordset("MIXTIM")  'Χρόνος διαλυσης
            MyTable("mixtim_fin") = ximitech.Recordset("MIXTIM_FIN")  'Χρόνος τελικής διάλυσης
            MyTable("thermo1") = ximitech.Recordset("THERMO1") 'Θερμοκρασία νερού διάλυσης 1=κρύο ,2=χλιαρό 3= ζεστο
            MyTable("thermo2") = ximitech.Recordset("THERMO2") 'Θερμοκρασία νερού αραίωσης 1=κρύο ,2=χλιαρό 3= ζεστο
            MyTable("nero_dial") = ximitech.Recordset("nero_dial") 'νερό διάλυσης
            MyTable("ximiko") = ximitech.Recordset("kod") + "-" + ximitech.Recordset("perigrafh") 'περιγραφή
            MyTable.update          'Επιστέφεται nero_act=ποσότης νερού και color_act =ποσότης χρώματος
         End If
     Next
     mydb.Close  ' Close Database
    ' RetVal = Shell("C:\talos\skon_dos.EXE", 1)    ' Run skones
 If Check1.value = 1 Then
    araiosh.Show
 Else
'    If Balance_Type = "ADAMEQ" Then
 '      skones2.Show
  '  Else
       skones.Show
   ' End If
 End If
     
'     For k = 1 To 2000
'           DoEvents
'    Next
     
     Command2.Enabled = True
     Pathma_Koympioy = 2
     Exit Sub
End If
'****************************************************************
'****************************************************************
'***** ΠΑΤΑ 2η ΦΟΡΑ ΤΟ ΚΟΥΜΠΙ ΝΑ ΤΑ ΠΑΕΙ ΠΙΣΩ

If Pathma_Koympioy = 2 Then

   Pathma_Koympioy = 3
   MsgBox mL_Res(551), , "Talos"
   '551  "Eίναι οι θέσεις ΜΠΟΥΚΑΛΙΩΝ ΕΛΕΥΘΕΡΕΣ στο τραπέζι;"
   Command2.Enabled = False
     Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\SKON_TB.MDB")
           ' Open table.
           Set MyTable = mydb.OpenRecordset("entoles", dbOpenTable)
           
            MyTable.MoveFirst  'Π’ΕΙ ΣΤΟ 1ο record.
 Do While Not MyTable.eof
              k = MyTable("pot")
              
                            
              ximitech.Recordset.Index = "KOD"
              ximitech.Recordset.Seek "=", left$(MyTable("ximiko"), 6)   ' kodikos(Val(Label3(k).Caption))
              
              cmd$ = "!var12=" + str(12000 + ximitech.Recordset("addr_prot")) + ":!var13=" + str(k) + ":"
              dum = wait_to_understand(cmd$)
              dum = Wait_Robot(0)
              ximitech.Recordset.Index = "KOD"
              ximitech.Recordset.Seek "=", left$(MyTable("ximiko"), 6)
                ximitech.Recordset.EDIT
               'If XIMITECH.Recordset.NoMatch Then
                
                
                
                ximitech.Recordset("real_cons") = 0
                 ximitech.Recordset("ypol_PROT") = 0
                mADD_ster = ximitech.Recordset("addr_old")
                ximitech.Recordset.update
                 
           If IsNull(MyTable("COLOR_ACT")) Or IsNull(MyTable("NERO_ACT")) Then
                 TELIKO = 0
            Else
                 If IsNull(ximitech.Recordset("bohu_xhm")) Then
                                TELIKO = MyTable("color_act") + MyTable("nero_act")
                Else
                    If gXimikaDial Then  ' aytomata ximika
                        TELIKO = MyTable("color_act") + MyTable("nero_act")
                    Else
                       ' manual
                        TELIKO = MyTable("color_act") + MyTable("nero_act") + ximitech.Recordset("bohu_xhm")
                    End If
                End If
            End If
               ' teliko = MyTable("color_act") + MyTable("nero_act") + XIMITECH.Recordset("bohu_xhm")
                If TELIKO > 0 And MyTable("color_act") > 0 Then
                      ximitech.Recordset.EDIT
                      ximitech.Recordset("ypol_PROT") = TELIKO / 1000
                      ximitech.Recordset("LHJ_PROT") = Date + ximitech.Recordset("ZVH_PROT")
                      ximitech.Recordset("real_cons") = 100 * MyTable("color_act") / TELIKO
                      ximitech.Recordset.update
                      
                      ' ενημέρωση στερεού με ποσότητα που  καταναλώθηκε
                       If Check1.value = 0 Then
                          ximitech.Recordset.Seek "=", kodikos(mADD_ster)
                          ximitech.Recordset.EDIT
                          ximitech.Recordset("ypol_prot") = ximitech.Recordset("ypol_prot") - MyTable("color_act") / 1000
                          ximitech.Recordset.update
                      End If
                      
                      
                End If
                Label1(Label3(k).Caption).BackColor = OldColor(Label3(k).Caption)
                Label3(k).BackColor = &HFFFFFF  'ασπρο
                MyTable.MoveNext
   Loop
   
       
           
            SQL = "DELETE DISTINCTROW entoles.* FROM entoles;"
            mydb.Execute SQL, dbFailOnError
    
   mydb.Close  ' Close Database
    
    
    
    
     kod6400 = 82
    ' "ΠΑΡΚΑΡΙΣΕ "
     cmd$ = "!var12=" + str(kod6400) + ":"
     dum = wait_to_understand(cmd$)
     dum = Wait_Robot(0)
     Command1.Enabled = True
     
End If

End Sub

Private Sub Command3_Click()
 If Check1.value = 1 Then
    araiosh.Show
 Else
  'If Balance_Type = "ADAMEQ" Then
   '   skones2.Show
  'Else
      skones.Show
  'End If
 End If
End Sub

Private Sub Form_Load()

Dim mydb As Database, MyTable As Recordset
Dim Answer, mAD, mColorID, a, Dial As Recordset, Mydb2 As Database
Dim OlaEtoima


MDIForm1.Arxeia(10).Visible = False
MDIForm1.Syntages(20).Visible = False
MDIForm1.bohuhtika(30).Visible = False
MDIForm1.ejodos(40).Visible = False



Pathma_Koympioy = 0

Command1.Caption = mL_CapRes(534) 'εξοδος
Command2.Caption = mL_CapRes(535) 'Κατασκευή διαλυμάτων
Check1.Caption = mLResnew(905, "With dillution", 0)
  aneb = 6285 - 5210
  
 For k = ar_real_pothr + 1 To Ar_Pothr
      Label3(k).Visible = False
  Next
  
  For k = 1 To ar_real_pothr
      Label3(k).top = Label3(k).top - aneb '+ IIf(k Mod 2 <> 0, 0, -Label3(k).height / 2)
  Next
  
'  For k = 1 To Ar_Pothr
'    Label3(k).top = Label3(k).top - aneb '+ IIf(k Mod 2 <> 0, 0, -Label3(k).height / 2)
'  Next
    
  For k = 2 To Ar_Mpoykal Step 2
    Label1(k).top = Label1(k).top + Label1(k).height / 2
  Next
  '
  
   
' σχεδιάζω την νεκρή περιοχή  k_AR=0
  For k = 1 To Ar_1_mpoykal - 1
      Label1(k).Visible = False
  Next
  
  For k = Ar_tel_mpoykal + 1 To Ar_Mpoykal
      Label1(k).Visible = False
  Next
  
  
 L1 = Label1(K_AR + 1).left
 L6 = Label1(K_AR + 6).left
  
'  front.left = L1
  Command2.left = L1
  Command1.left = L6
  SXOLIA1.left = L1
  SXOLIA2.left = L1
  
  For k = Ar_1_mpoykal To Ar_tel_mpoykal
     Label1(k).BackColor = &HFFFFFF   ' ασπρο
     OldColor(k) = Label1(k).BackColor
     Label1(k).Caption = k
     MHTRIKO2(k) = 0
 Next
     
  For k = 1 To ar_real_pothr
     Label3(k).Caption = k
     OldColor(k) = Label3(k).BackColor
 Next
     
          
Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\SKON_TB.MDB")
Set MyTable = mydb.OpenRecordset("entoles", dbOpenTable)
Set Mydb2 = Workspaces(0).OpenDatabase("C:\TALOS\recipies.MDB")
Set Dial = Mydb2.OpenRecordset("ximitech", dbOpenTable)
ok = 1
                   On Error GoTo ADEIO
                   MyTable.MoveLast
                   
                    If MyTable.RecordCount >= 1 And ok = 1 Then
                          Answer = MsgBox(mL_Res(537), vbYesNo)
                          ' 537 "Υπάρχουν διαλύματα που δεν επέστρεψαν. Nα τα επιστρέψω σην θέση τους."
                          If Answer = vbYes Then
                                    Pathma_Koympioy = 2
                                    Command2.Caption = mL_CapRes(552)  '"επιστροφή"
                                    Command2.Enabled = True
                                    Command3.Visible = True
                                    Dial.Index = "KOD"
                                    MyTable.MoveFirst
                                    a = 0
                                    OlaEtoima = 1
                                    'On Error GoTo 0
                                    Do While Not MyTable.eof
                                       a = a + 1
                                       Dial.Seek "=", left$(MyTable("XIMIKO"), 6)
                                       If Not Dial.NoMatch Then
                                             Label3(a).Caption = Dial("ADDR_PROT")
                                             Label3(a).BackColor = vbRed
                                             If IsNull(MyTable("nero_act")) Then
                                                OlaEtoima = 0
                                             Else
                                                If MyTable("nero_act") > 10 Then
                                                   Label3(a).BackColor = vbGreen
                                                Else
                                                   OlaEtoima = 0
                                                End If
                                             End If
                                       End If
                                       MyTable.MoveNext
                                    Loop
                                    If OlaEtoima = 1 Then Command3.Enabled = False
                          Else
                              SQL = "DELETE  entoles.*  FROM entoles;"
                              mydb.Execute SQL
                           End If
                    End If
                    mydb.Close
                    On Error GoTo 0





Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
SQL = "select morfh,real_cons,skonh,kod,kataskeyas,addr_prot,addr_old from ximitech"
SQL = SQL + " where addr_prot>0 order by kataskeyas,skonh;"
Set x = mydb.OpenRecordset(SQL, dbOpenDynaset)
                    
' ενημερώνω όλα τα διαλύματα με την αντίστοιχη σκόνη
If x.RecordCount >= 1 Then
   
 Else
     MsgBox mL_Res(538)
     ' 538 "Δεν υπάρχουν μπουκάλια διαλυμάτων στο τραπέζι"
    Exit Sub
End If



x.MoveLast
mColorID = "ωωωωωωωωωωωωωωωω"
Do While Not x.BOF
   
   If x("skonh") = 1 Then
      mAD = x("addr_prot")
      mColorID = x("kataskeyas")
      periek(x("addr_prot")) = 0
   End If
   
   If x("skonh") = 0 Then
      x.EDIT
      If mColorID = x("kataskeyas") Then
         x("addr_old") = mAD
      Else
         x("addr_old") = 0
      End If
      If IsNull(x("real_cons")) Then
         periek(x("addr_prot")) = val2(x("morfh"))
      Else
         periek(x("addr_prot")) = x("real_cons")
      End If
      x.update
   End If
   x.MovePrevious


Loop
                    
mydb.Close
Exit Sub

ADEIO:
 ok = 0
 Resume Next
End Sub


Private Sub Form_Paint()
Dim u
 ximitech.Recordset.Index = "KOD"
 
 While Not ximitech.Recordset.eof
         u = 0
         If Not IsNull(ximitech.Recordset("ADDR_PROT")) Then
                 u = ximitech.Recordset("addr_prot")
         End If
         
         If u >= Ar_1_mpoykal And u <= Ar_tel_mpoykal Then    ' =============================================
                                  
                 If Not IsNull(ximitech.Recordset("LHJ_PROT")) Then   ' ---------------
                          hmeres = DateDiff("d", Now, ximitech.Recordset("LHJ_PROT"))
                 Else
                          hmeres = 1000
                 End If
                 
                If ximitech.Recordset("skonh") = 1 Then ' ΣΤΕΡΕΑ
                                          Label1(u).BackColor = &H808080           '   καφε
                Else  '    Δ Ι Α Λ Υ Μ Α Τ Α
                          Label1(u).BackColor = FIND_COLOR(hmeres)
                           
                           'Select Case hmeres
                           '          Case 0
                           '               Label1(u).BackColor = &H80FF&          '  πορτοκαλί
                           '           Case Is > 1
                           '               Label1(u).BackColor = &H8000& '  πράσινο
                           '           Case Else
                           '               Label1(u).BackColor = kokkino        'κοκκινο
                           ' End Select
                End If
               
                kodikos(u) = ximitech.Recordset("kod")
                Perigrafh(u) = ximitech.Recordset("perigrafh")
                LHJH(u) = ximitech.Recordset("Lhj_prot")
                Ypoloipo(u) = ximitech.Recordset("ypol_prot")
                
                 '  αν το υπόλοιπο είναι κάτω των 50 γραμμαρίων να κοκκινίζει
                 If Ypoloipo(u) < 50 Then
                 
                                 'αν το υπόλοιπο είναι κάτω των 50 γραμμαρίων να κοκκινίζει
               '  If Ypoloipo(u) < 50 Then
                '       Label1(u).BackColor = &HFF&        'κοκκινο
                ' End If

                 
                 
                   ' If XIMITECH.Recordset("skonh") = 1 Then
                   '    Label1(u).BackColor = Keramidi
                   ' Else
                       Label1(u).BackColor = kokkino       'κοκκινο
                    'End If
                 End If
                 OldColor(u) = Label1(u).BackColor
                
                MHTRIKO(u) = ximitech.Recordset("addr_old")
                 ' αναλογα με την ποσότητα να είναι το μέγεθος
               ' Label1(u).height = 100 + h0 * (Ypoloipo(u) / m0)
               ' Label1(u).top = Label1(u).top + (ho - Label1(u).height)
                
                If Not IsNull(Perigrafh(u)) Then
                     Label1(u).Caption = Format(u, "###") + " " + Perigrafh(u)
                End If

         End If ' ================================================================
         ximitech.Recordset.MoveNext
 Wend
 Screen.MousePointer = 1

'      XIMITECH.Recordset.Index = "KOD"
'      XIMITECH.Recordset.Seek "=", Data1(0).Recordset("XHM" + Right(Str(Index + 1), 1)) ' Label1(Index).Caption
'      If Not XIMITECH.Recordset.NoMatch Then
'            If IsNull(XIMITECH.Recordset("perigrafh")) Then
'                            Label11(Index).Caption = " "
'            Else
'                Label11(Index).Caption = XIMITECH.Recordset("perigrafh")
'            End If
'      End If
 
End Sub


Private Sub Form_Unload(Cancel As Integer)
MDIForm1.Arxeia(10).Visible = True
MDIForm1.Syntages(20).Visible = True
MDIForm1.bohuhtika(30).Visible = True
MDIForm1.ejodos(40).Visible = True

ValveReset

'MDIForm1.arxeia.Visible = True
'MDIForm1.Syntages.Visible = True
'MDIForm1.Bohuhtika.Visible = True
'MDIForm1.ejodos.Visible = True
End Sub



Private Sub Label1_Click(Index As Integer)
Dim ok

Check1.Enabled = False
If Label1(Index).BackColor = &HFFFFFF Or Label1(Index).BackColor = &H808080 Then         'ασπρο ή καφέ(σκόνη)  δεν κινείται
     '
ElseIf Label1(Index).BackColor = Keramidi Then
    ' τελειομενη σκονη
ElseIf Label1(Index).BackColor = ROZ Then ' αν είναι  γκρίζο να  γινει όπως ήταν
  ' το βγάζει από την θέση μπροστά
   For k = 1 To ar_real_pothr
      If Label3(k).Caption = Index Then            '  ξαναβάζω τον αρχικό του αριθμό
           If k + 1 <= ar_real_pothr Then
              If Label3(k + 1).BackColor = kokkino Then
                 MsgBox mLResnew(914, "You can Undo only the last one.", 1)
                 Exit Sub
              End If
           End If
           Label3(k).BackColor = OldColor(k)    ' ασπρο
           Label3(k).Caption = k                       '  ξαναβάζω τον αρχικό του αριθμό
          Exit For
      End If
   Next
   Label1(Index).BackColor = OldColor(Index)
   
   'Label1(Index).BackColor = OldColor(Index)
   ' το βγάζει από την θέση μπροστά
   'For k = 1 To ar_real_pothr
   '   If Label3(k).Caption = Index Then            '  ξαναβάζω τον αρχικό του αριθμό
   '        Label3(k).BackColor = &HFFFFFF     ' ασπρο
   '        Label3(k).Caption = k - 1                      '  ξαναβάζω τον αρχικό του αριθμό
   '       Exit For
   '   End If
   'Next
Else
 If Check1.value = 1 Then
      ok = 0
      For k = Ar_1_mpoykal To Ar_tel_mpoykal
         If k <> Index Then
            If MHTRIKO(Index) = MHTRIKO(k) Then  'idia mana
               If periek(k) > periek(Index) Then 'pyknotero
                  MHTRIKO2(Index) = k
                  ximitech.Recordset.Index = "KOD"
                  ximitech.Recordset.Seek "=", kodikos(Index)
                  ximitech.Recordset.EDIT
                  ximitech.Recordset("addr_old") = k
                  ximitech.Recordset.update
                  ok = 1
                  Command2.Enabled = True
                  Exit For
               End If
            End If
         End If
      Next
      If ok = 0 Then
         MsgBox mLResnew(904, "No related solution", 1)
         Exit Sub
      End If
 
 Else
  If IsNull(MHTRIKO(Index)) Then
        MsgBox mL_Res(544), , "Talos"
        '544 "Δεν έχει δηλωθεί μητρική σκόνη...  "
        Exit Sub
  ElseIf MHTRIKO(Index) < 1 Or MHTRIKO(Index) > Ar_Mpoykal Then ' είναι δηλωμένο σαν μητρικό ασχετο νούμερο
        MsgBox mL_Res(545), , "Talos"
        '545 "Δεν έχει δηλωθεί μητρική σκόνη...  "
        Exit Sub
  Else
         If Label1(MHTRIKO(Index)).BackColor = &H808080 Then ' αν είναι καφέ τότε είναι σκόνη
            ' ολα καλά
            Command2.Enabled = True
            '
         Else
             'MsgBox mL_Res(546), , "Talos"
             MsgBox mLResnew(916, "The stock of the related powder is low", 1), , "Talos"
             '546 "Δεν υπάρχει στο τραπέζι η μητρική σκόνη...  ", , "Talos"
             Exit Sub
         End If
   End If
 End If
 
   ' ψάχνει θέση για τοποθέτηση μπροστά
   For k = 1 To ar_real_pothr
      If Label3(k).BackColor <> kokkino Then ' αν δεν είναι κόκκινη θέση φέρτο εδώ
          Label3(k).Caption = Index            'παίρνει τον αριθμό του μπουκαλιού που θάρθεi
          Label3(k).BackColor = kokkino   'κοκκινίζω την μπροστά θέση

          Label1(Index).BackColor = ROZ ' &H80C0FF          'το αλλάζω χρωμα  για να μην ξαναζητηθεί
          Exit For
      End If
   Next
End If
End Sub

Private Sub Label1_DblClick(Index As Integer)
Dim kodikos


End Sub


Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)


If Not IsNull(Perigrafh(Index)) Then
     ' Me.Caption = Perigrafh(Index)
     SXOLIA1.Caption = Perigrafh(Index)
 End If
If Not IsNull(Format(Ypoloipo(Index), "###") + "              " + Format(LHJH(Index), "dd/mm/yy")) Then
     SXOLIA2.Caption = Format(Ypoloipo(Index), "###") + "              " + Format(LHJH(Index), "dd/mm/yy")
     ' Me.Caption = Me.Caption + Format(Ypoloipo(Index), "###") + "              " + Format(Lhjh(Index), "dd/mm/yy")
End If
End Sub



Private Sub Label3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim n
On Error Resume Next
Me.Caption = ""

n = Label3(Index).Caption

If n > 60 Then
   
    If Not IsNull(Perigrafh(n)) Then
      'sxolia1.Caption = Perigrafh(Index)
      If left(Perigrafh(n), 1) <> "," Then
         SXOLIA1.Caption = Perigrafh(n)
         'Me.Caption = Perigrafh(n)
      End If
    End If
    
    If Not IsNull(Format(Ypoloipo(n), "###") + "              " + Format(LHJH(n), "dd/mm/yy")) And Len(Trim(Me.Caption)) > 1 Then
        'sxolia2.Caption = Format(Ypoloipo(Index), "###") + "              " + Format(Lhjh(Index), "dd/mm/yy")
        SXOLIA1.Caption = Trim(SXOLIA1.Caption) + mL_CapRes(584) + Format(Ypoloipo(n), "###") + mL_CapRes(583) + Format(LHJH(n), "dd/mm/yy")
        ' Me.Caption = Me.Caption + mL_CapRes(584) + Format(Ypoloipo(n), "###") + mL_CapRes(583) + Format(Lhjh(n), "dd/mm/yy")
    End If
    
End If
End Sub


Private Sub Timer1_Timer()
   ' If Command2.Enabled = False Then
        If system_ready = 1 Then
              Elegxo_Katastash
        End If
   ' End If
End Sub


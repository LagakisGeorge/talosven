VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FERE 
   BackColor       =   &H00FF0000&
   ClientHeight    =   6975
   ClientLeft      =   -1365
   ClientTop       =   3255
   ClientWidth     =   9660
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6975
   ScaleWidth      =   9660
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Φερε μπουκάλια για έλεγχο"
      Enabled         =   0   'False
      Height          =   300
      Left            =   300
      TabIndex        =   288
      Top             =   5260
      Width           =   3000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ανανέωση Διαλυμάτων"
      Enabled         =   0   'False
      Height          =   300
      Left            =   300
      TabIndex        =   6
      Top             =   5710
      Width           =   3000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ΕΞΟΔΟΣ"
      Height          =   1155
      Left            =   3330
      TabIndex        =   5
      Top             =   4860
      Width           =   930
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1770
      Top             =   8145
   End
   Begin VB.Data XIMITECH 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "c:\talos\recipies.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2790
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "ximitech"
      Top             =   6855
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton front 
      Caption         =   "Οριστική Απομάκρυνση"
      Enabled         =   0   'False
      Height          =   300
      Left            =   300
      TabIndex        =   2
      Top             =   4840
      Width           =   3000
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1200
      Top             =   45
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "299"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   299
      Left            =   8550
      TabIndex        =   338
      Top             =   -155
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "298"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   298
      Left            =   8085
      TabIndex        =   337
      Top             =   -155
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "297"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   297
      Left            =   7605
      TabIndex        =   336
      Top             =   -155
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "276"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   296
      Left            =   7140
      TabIndex        =   335
      Top             =   -155
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "295"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   295
      Left            =   6675
      TabIndex        =   334
      Top             =   -155
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "294"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   294
      Left            =   6195
      TabIndex        =   333
      Top             =   -155
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "293"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   293
      Left            =   5730
      TabIndex        =   332
      Top             =   -155
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "292"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   292
      Left            =   5265
      TabIndex        =   331
      Top             =   -155
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "290"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   290
      Left            =   4320
      TabIndex        =   330
      Top             =   -155
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "289"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   289
      Left            =   3855
      TabIndex        =   329
      Top             =   -155
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "288"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   288
      Left            =   3375
      TabIndex        =   328
      Top             =   -155
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "287"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   287
      Left            =   2910
      TabIndex        =   327
      Top             =   -155
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "286"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   286
      Left            =   2445
      TabIndex        =   326
      Top             =   -155
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "285"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   285
      Left            =   1965
      TabIndex        =   325
      Top             =   -155
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "284"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   284
      Left            =   1500
      TabIndex        =   324
      Top             =   -155
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "282"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   282
      Left            =   555
      TabIndex        =   323
      Top             =   -155
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "281"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   281
      Left            =   90
      TabIndex        =   322
      Top             =   -155
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "300"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   300
      Left            =   9015
      TabIndex        =   321
      Top             =   -155
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "291"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   291
      Left            =   4785
      TabIndex        =   320
      Top             =   -155
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "283"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   283
      Left            =   1035
      TabIndex        =   319
      Top             =   -155
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "30"
      Height          =   375
      Index           =   30
      Left            =   8550
      TabIndex        =   318
      Top             =   6390
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "29"
      Height          =   375
      Index           =   29
      Left            =   8085
      TabIndex        =   317
      Top             =   6390
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "28"
      Height          =   375
      Index           =   28
      Left            =   7605
      TabIndex        =   316
      Top             =   6390
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "20"
      Height          =   375
      Index           =   20
      Left            =   8550
      TabIndex        =   315
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "19"
      Height          =   375
      Index           =   19
      Left            =   8085
      TabIndex        =   314
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "18"
      Height          =   375
      Index           =   18
      Left            =   7605
      TabIndex        =   313
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "10"
      Height          =   375
      Index           =   10
      Left            =   8550
      TabIndex        =   312
      Top             =   7290
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "9"
      Height          =   375
      Index           =   9
      Left            =   8085
      TabIndex        =   311
      Top             =   7290
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "8"
      Height          =   375
      Index           =   8
      Left            =   7605
      TabIndex        =   310
      Top             =   7290
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "13"
      Height          =   375
      Index           =   13
      Left            =   5265
      TabIndex        =   309
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "1"
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   308
      Top             =   7290
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "2"
      Height          =   375
      Index           =   2
      Left            =   4785
      TabIndex        =   307
      Top             =   7290
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "3"
      Height          =   375
      Index           =   3
      Left            =   5265
      TabIndex        =   306
      Top             =   7290
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "4"
      Height          =   375
      Index           =   4
      Left            =   5730
      TabIndex        =   305
      Top             =   7290
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "5"
      Height          =   375
      Index           =   5
      Left            =   6195
      TabIndex        =   304
      Top             =   7290
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "6"
      Height          =   375
      Index           =   6
      Left            =   6675
      TabIndex        =   303
      Top             =   7290
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "7"
      Height          =   375
      Index           =   7
      Left            =   7140
      TabIndex        =   302
      Top             =   7290
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "11"
      Height          =   375
      Index           =   11
      Left            =   4320
      TabIndex        =   301
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "12"
      Height          =   375
      Index           =   12
      Left            =   4785
      TabIndex        =   300
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "14"
      Height          =   375
      Index           =   14
      Left            =   5730
      TabIndex        =   299
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "15"
      Height          =   375
      Index           =   15
      Left            =   6195
      TabIndex        =   298
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "16"
      Height          =   375
      Index           =   16
      Left            =   6675
      TabIndex        =   297
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "17"
      Height          =   375
      Index           =   17
      Left            =   7140
      TabIndex        =   296
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "21"
      Height          =   375
      Index           =   21
      Left            =   4320
      TabIndex        =   295
      Top             =   6390
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "22"
      Height          =   375
      Index           =   22
      Left            =   4785
      TabIndex        =   294
      Top             =   6390
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "23"
      Height          =   375
      Index           =   23
      Left            =   5265
      TabIndex        =   293
      Top             =   6390
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "24"
      Height          =   375
      Index           =   24
      Left            =   5730
      TabIndex        =   292
      Top             =   6390
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "25"
      Height          =   375
      Index           =   25
      Left            =   6195
      TabIndex        =   291
      Top             =   6390
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "26"
      Height          =   375
      Index           =   26
      Left            =   6675
      TabIndex        =   290
      Top             =   6390
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "27"
      Height          =   375
      Index           =   27
      Left            =   7140
      TabIndex        =   289
      Top             =   6390
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1"
      ForeColor       =   &H00000001&
      Height          =   375
      Index           =   1
      Left            =   90
      TabIndex        =   287
      Top             =   5970
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "21"
      ForeColor       =   &H00000015&
      Height          =   375
      Index           =   21
      Left            =   90
      TabIndex        =   286
      Top             =   5535
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "41"
      ForeColor       =   &H00000029&
      Height          =   375
      Index           =   41
      Left            =   90
      TabIndex        =   285
      Top             =   5085
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "61"
      ForeColor       =   &H0000003D&
      Height          =   375
      Index           =   61
      Left            =   90
      TabIndex        =   284
      Top             =   4650
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "81"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   81
      Left            =   90
      TabIndex        =   283
      Top             =   4215
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "101"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   101
      Left            =   90
      TabIndex        =   282
      Top             =   3765
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "121"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   121
      Left            =   90
      TabIndex        =   281
      Top             =   3330
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "141"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   141
      Left            =   90
      TabIndex        =   280
      Top             =   2895
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "161"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   161
      Left            =   90
      TabIndex        =   279
      Top             =   2445
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "181"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   181
      Left            =   90
      TabIndex        =   278
      Top             =   2010
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "201"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   201
      Left            =   90
      TabIndex        =   277
      Top             =   1575
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "221"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   221
      Left            =   90
      TabIndex        =   276
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "241"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   241
      Left            =   90
      TabIndex        =   275
      Top             =   690
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2"
      ForeColor       =   &H00000002&
      Height          =   375
      Index           =   2
      Left            =   555
      TabIndex        =   274
      Top             =   5970
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000016&
      Height          =   375
      Index           =   22
      Left            =   555
      TabIndex        =   273
      Top             =   5535
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H0000002A&
      Height          =   375
      Index           =   42
      Left            =   555
      TabIndex        =   272
      Top             =   5085
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H0000003E&
      Height          =   375
      Index           =   62
      Left            =   555
      TabIndex        =   271
      Top             =   4650
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "82"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   82
      Left            =   555
      TabIndex        =   270
      Top             =   4215
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   102
      Left            =   555
      TabIndex        =   269
      Top             =   3765
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   122
      Left            =   555
      TabIndex        =   268
      Top             =   3330
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   142
      Left            =   555
      TabIndex        =   267
      Top             =   2895
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   162
      Left            =   555
      TabIndex        =   266
      Top             =   2445
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   182
      Left            =   555
      TabIndex        =   265
      Top             =   2010
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   202
      Left            =   555
      TabIndex        =   264
      Top             =   1575
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   222
      Left            =   555
      TabIndex        =   263
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   242
      Left            =   555
      TabIndex        =   262
      Top             =   690
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   1035
      TabIndex        =   261
      Top             =   5970
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000017&
      Height          =   375
      Index           =   23
      Left            =   1035
      TabIndex        =   260
      Top             =   5535
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H0000002B&
      Height          =   375
      Index           =   43
      Left            =   1035
      TabIndex        =   259
      Top             =   5085
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H0000003F&
      Height          =   375
      Index           =   63
      Left            =   1035
      TabIndex        =   258
      Top             =   4650
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "83"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   83
      Left            =   1035
      TabIndex        =   257
      Top             =   4215
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   103
      Left            =   1035
      TabIndex        =   256
      Top             =   3765
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   123
      Left            =   1035
      TabIndex        =   255
      Top             =   3330
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   143
      Left            =   1035
      TabIndex        =   254
      Top             =   2895
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   163
      Left            =   1035
      TabIndex        =   253
      Top             =   2445
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   183
      Left            =   1035
      TabIndex        =   252
      Top             =   2010
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   203
      Left            =   1035
      TabIndex        =   251
      Top             =   1575
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   223
      Left            =   1035
      TabIndex        =   250
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   243
      Left            =   1035
      TabIndex        =   249
      Top             =   690
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "4"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   1500
      TabIndex        =   248
      Top             =   5970
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000018&
      Height          =   375
      Index           =   24
      Left            =   1500
      TabIndex        =   247
      Top             =   5535
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H0000002C&
      Height          =   375
      Index           =   44
      Left            =   1500
      TabIndex        =   246
      Top             =   5085
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000040&
      Height          =   375
      Index           =   64
      Left            =   1500
      TabIndex        =   245
      Top             =   4650
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "84"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   84
      Left            =   1500
      TabIndex        =   244
      Top             =   4215
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   104
      Left            =   1500
      TabIndex        =   243
      Top             =   3765
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   124
      Left            =   1500
      TabIndex        =   242
      Top             =   3330
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   144
      Left            =   1500
      TabIndex        =   241
      Top             =   2895
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   164
      Left            =   1500
      TabIndex        =   240
      Top             =   2445
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   184
      Left            =   1500
      TabIndex        =   239
      Top             =   2010
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   204
      Left            =   1500
      TabIndex        =   238
      Top             =   1575
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   224
      Left            =   1500
      TabIndex        =   237
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   244
      Left            =   1500
      TabIndex        =   236
      Top             =   690
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "13"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   13
      Left            =   5730
      TabIndex        =   235
      Top             =   5970
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "33"
      ForeColor       =   &H00000019&
      Height          =   375
      Index           =   33
      Left            =   5730
      TabIndex        =   234
      Top             =   5535
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "53"
      ForeColor       =   &H0000002D&
      Height          =   375
      Index           =   53
      Left            =   5730
      TabIndex        =   233
      Top             =   5085
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "73"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   73
      Left            =   5730
      TabIndex        =   232
      Top             =   4650
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "85"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   85
      Left            =   1965
      TabIndex        =   231
      Top             =   4215
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   105
      Left            =   1965
      TabIndex        =   230
      Top             =   3765
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   125
      Left            =   1965
      TabIndex        =   229
      Top             =   3330
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   145
      Left            =   1965
      TabIndex        =   228
      Top             =   2895
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   165
      Left            =   1965
      TabIndex        =   227
      Top             =   2445
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   185
      Left            =   1965
      TabIndex        =   226
      Top             =   2010
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   205
      Left            =   1965
      TabIndex        =   225
      Top             =   1575
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   225
      Left            =   1965
      TabIndex        =   224
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   245
      Left            =   1965
      TabIndex        =   223
      Top             =   690
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "14"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   14
      Left            =   6195
      TabIndex        =   222
      Top             =   5970
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "34"
      ForeColor       =   &H0000001A&
      Height          =   375
      Index           =   34
      Left            =   6195
      TabIndex        =   221
      Top             =   5535
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "54"
      ForeColor       =   &H0000002E&
      Height          =   375
      Index           =   54
      Left            =   6195
      TabIndex        =   220
      Top             =   5085
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "74"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   74
      Left            =   6195
      TabIndex        =   219
      Top             =   4650
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "86"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   86
      Left            =   2445
      TabIndex        =   218
      Top             =   4215
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   106
      Left            =   2445
      TabIndex        =   217
      Top             =   3765
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   126
      Left            =   2445
      TabIndex        =   216
      Top             =   3330
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   146
      Left            =   2445
      TabIndex        =   215
      Top             =   2895
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   166
      Left            =   2445
      TabIndex        =   214
      Top             =   2445
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   186
      Left            =   2445
      TabIndex        =   213
      Top             =   2010
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   206
      Left            =   2445
      TabIndex        =   212
      Top             =   1575
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   226
      Left            =   2445
      TabIndex        =   211
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   246
      Left            =   2445
      TabIndex        =   210
      Top             =   690
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "15"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   15
      Left            =   6675
      TabIndex        =   209
      Top             =   5970
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "35"
      ForeColor       =   &H0000001B&
      Height          =   375
      Index           =   35
      Left            =   6675
      TabIndex        =   208
      Top             =   5535
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "55"
      ForeColor       =   &H0000002F&
      Height          =   375
      Index           =   55
      Left            =   6675
      TabIndex        =   207
      Top             =   5085
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "75"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   75
      Left            =   6675
      TabIndex        =   206
      Top             =   4650
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "87"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   87
      Left            =   2910
      TabIndex        =   205
      Top             =   4215
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   107
      Left            =   2910
      TabIndex        =   204
      Top             =   3765
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   127
      Left            =   2910
      TabIndex        =   203
      Top             =   3330
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   147
      Left            =   2910
      TabIndex        =   202
      Top             =   2895
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   167
      Left            =   2910
      TabIndex        =   201
      Top             =   2445
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   187
      Left            =   2910
      TabIndex        =   200
      Top             =   2010
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   207
      Left            =   2910
      TabIndex        =   199
      Top             =   1575
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   227
      Left            =   2910
      TabIndex        =   198
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "247"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   247
      Left            =   2910
      TabIndex        =   197
      Top             =   690
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "88"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   88
      Left            =   3375
      TabIndex        =   196
      Top             =   4215
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   108
      Left            =   3375
      TabIndex        =   195
      Top             =   3765
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   128
      Left            =   3375
      TabIndex        =   194
      Top             =   3330
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   148
      Left            =   3375
      TabIndex        =   193
      Top             =   2895
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   168
      Left            =   3375
      TabIndex        =   192
      Top             =   2445
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   188
      Left            =   3375
      TabIndex        =   191
      Top             =   2010
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   208
      Left            =   3375
      TabIndex        =   190
      Top             =   1575
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   228
      Left            =   3375
      TabIndex        =   189
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "248"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   248
      Left            =   3375
      TabIndex        =   188
      Top             =   690
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "89"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   89
      Left            =   3855
      TabIndex        =   187
      Top             =   4215
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   109
      Left            =   3855
      TabIndex        =   186
      Top             =   3765
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   129
      Left            =   3855
      TabIndex        =   185
      Top             =   3330
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   149
      Left            =   3855
      TabIndex        =   184
      Top             =   2895
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   169
      Left            =   3855
      TabIndex        =   183
      Top             =   2445
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   189
      Left            =   3855
      TabIndex        =   182
      Top             =   2010
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   209
      Left            =   3855
      TabIndex        =   181
      Top             =   1575
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   229
      Left            =   3855
      TabIndex        =   180
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "249"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   249
      Left            =   3855
      TabIndex        =   179
      Top             =   690
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "90"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   90
      Left            =   4320
      TabIndex        =   178
      Top             =   4215
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   110
      Left            =   4320
      TabIndex        =   177
      Top             =   3765
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   130
      Left            =   4320
      TabIndex        =   176
      Top             =   3330
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   150
      Left            =   4320
      TabIndex        =   175
      Top             =   2895
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   170
      Left            =   4320
      TabIndex        =   174
      Top             =   2445
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   190
      Left            =   4320
      TabIndex        =   173
      Top             =   2010
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   210
      Left            =   4320
      TabIndex        =   172
      Top             =   1575
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   230
      Left            =   4320
      TabIndex        =   171
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   250
      Left            =   4320
      TabIndex        =   170
      Top             =   690
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "91"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   91
      Left            =   4785
      TabIndex        =   169
      Top             =   4215
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   111
      Left            =   4785
      TabIndex        =   168
      Top             =   3765
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   131
      Left            =   4785
      TabIndex        =   167
      Top             =   3330
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   151
      Left            =   4785
      TabIndex        =   166
      Top             =   2895
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   171
      Left            =   4785
      TabIndex        =   165
      Top             =   2445
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   191
      Left            =   4785
      TabIndex        =   164
      Top             =   2010
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   211
      Left            =   4785
      TabIndex        =   163
      Top             =   1575
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   231
      Left            =   4785
      TabIndex        =   162
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "251"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   251
      Left            =   4785
      TabIndex        =   161
      Top             =   690
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "92"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   92
      Left            =   5265
      TabIndex        =   160
      Top             =   4215
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   112
      Left            =   5265
      TabIndex        =   159
      Top             =   3765
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   132
      Left            =   5265
      TabIndex        =   158
      Top             =   3330
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   152
      Left            =   5265
      TabIndex        =   157
      Top             =   2895
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   172
      Left            =   5265
      TabIndex        =   156
      Top             =   2445
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   192
      Left            =   5265
      TabIndex        =   155
      Top             =   2010
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   212
      Left            =   5265
      TabIndex        =   154
      Top             =   1575
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   232
      Left            =   5265
      TabIndex        =   153
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "252"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   252
      Left            =   5265
      TabIndex        =   152
      Top             =   690
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "93"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   93
      Left            =   5730
      TabIndex        =   151
      Top             =   4215
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   113
      Left            =   5730
      TabIndex        =   150
      Top             =   3765
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   133
      Left            =   5730
      TabIndex        =   149
      Top             =   3330
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   153
      Left            =   5730
      TabIndex        =   148
      Top             =   2895
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   173
      Left            =   5730
      TabIndex        =   147
      Top             =   2445
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   193
      Left            =   5730
      TabIndex        =   146
      Top             =   2010
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   213
      Left            =   5730
      TabIndex        =   145
      Top             =   1575
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   233
      Left            =   5730
      TabIndex        =   144
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "253"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   253
      Left            =   5730
      TabIndex        =   143
      Top             =   690
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "94"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   94
      Left            =   6195
      TabIndex        =   142
      Top             =   4215
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   114
      Left            =   6195
      TabIndex        =   141
      Top             =   3765
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   134
      Left            =   6195
      TabIndex        =   140
      Top             =   3330
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   154
      Left            =   6195
      TabIndex        =   139
      Top             =   2895
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   174
      Left            =   6195
      TabIndex        =   138
      Top             =   2445
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   194
      Left            =   6195
      TabIndex        =   137
      Top             =   2010
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   214
      Left            =   6195
      TabIndex        =   136
      Top             =   1575
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   234
      Left            =   6195
      TabIndex        =   135
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "254"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   254
      Left            =   6195
      TabIndex        =   134
      Top             =   690
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "95"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   95
      Left            =   6675
      TabIndex        =   133
      Top             =   4215
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   115
      Left            =   6675
      TabIndex        =   132
      Top             =   3765
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   135
      Left            =   6675
      TabIndex        =   131
      Top             =   3330
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   155
      Left            =   6675
      TabIndex        =   130
      Top             =   2895
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   175
      Left            =   6675
      TabIndex        =   129
      Top             =   2445
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   195
      Left            =   6675
      TabIndex        =   128
      Top             =   2010
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   215
      Left            =   6675
      TabIndex        =   127
      Top             =   1575
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   235
      Left            =   6675
      TabIndex        =   126
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "255"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   255
      Left            =   6675
      TabIndex        =   125
      Top             =   690
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "16"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   16
      Left            =   7140
      TabIndex        =   124
      Top             =   5970
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   36
      Left            =   7140
      TabIndex        =   123
      Top             =   5535
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   56
      Left            =   7140
      TabIndex        =   122
      Top             =   5085
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   76
      Left            =   7140
      TabIndex        =   121
      Top             =   4650
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "96"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   96
      Left            =   7140
      TabIndex        =   120
      Top             =   4215
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   116
      Left            =   7140
      TabIndex        =   119
      Top             =   3765
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   136
      Left            =   7140
      TabIndex        =   118
      Top             =   3330
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   156
      Left            =   7140
      TabIndex        =   117
      Top             =   2895
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   176
      Left            =   7140
      TabIndex        =   116
      Top             =   2445
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   196
      Left            =   7140
      TabIndex        =   115
      Top             =   2010
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   216
      Left            =   7140
      TabIndex        =   114
      Top             =   1575
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   236
      Left            =   7140
      TabIndex        =   113
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "256"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   256
      Left            =   7140
      TabIndex        =   112
      Top             =   690
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "17"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   17
      Left            =   7605
      TabIndex        =   111
      Top             =   5970
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   37
      Left            =   7605
      TabIndex        =   110
      Top             =   5535
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   57
      Left            =   7605
      TabIndex        =   109
      Top             =   5085
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   77
      Left            =   7605
      TabIndex        =   108
      Top             =   4650
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "97"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   97
      Left            =   7605
      TabIndex        =   107
      Top             =   4215
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   117
      Left            =   7605
      TabIndex        =   106
      Top             =   3765
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   137
      Left            =   7605
      TabIndex        =   105
      Top             =   3330
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   157
      Left            =   7605
      TabIndex        =   104
      Top             =   2895
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   177
      Left            =   7605
      TabIndex        =   103
      Top             =   2445
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   197
      Left            =   7605
      TabIndex        =   102
      Top             =   2010
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   217
      Left            =   7605
      TabIndex        =   101
      Top             =   1575
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   237
      Left            =   7605
      TabIndex        =   100
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "257"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   257
      Left            =   7605
      TabIndex        =   99
      Top             =   690
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "18"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   18
      Left            =   8085
      TabIndex        =   98
      Top             =   5970
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   38
      Left            =   8085
      TabIndex        =   97
      Top             =   5535
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   58
      Left            =   8085
      TabIndex        =   96
      Top             =   5085
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   78
      Left            =   8085
      TabIndex        =   95
      Top             =   4650
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "98"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   98
      Left            =   8085
      TabIndex        =   94
      Top             =   4215
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   118
      Left            =   8085
      TabIndex        =   93
      Top             =   3765
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   138
      Left            =   8085
      TabIndex        =   92
      Top             =   3330
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   158
      Left            =   8085
      TabIndex        =   91
      Top             =   2895
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   178
      Left            =   8085
      TabIndex        =   90
      Top             =   2445
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   198
      Left            =   8085
      TabIndex        =   89
      Top             =   2010
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   218
      Left            =   8085
      TabIndex        =   88
      Top             =   1575
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   238
      Left            =   8085
      TabIndex        =   87
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "258"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   258
      Left            =   8085
      TabIndex        =   86
      Top             =   690
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "19"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   19
      Left            =   8550
      TabIndex        =   85
      Top             =   5970
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   39
      Left            =   8550
      TabIndex        =   84
      Top             =   5535
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   59
      Left            =   8550
      TabIndex        =   83
      Top             =   5085
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   79
      Left            =   8550
      TabIndex        =   82
      Top             =   4650
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "99"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   99
      Left            =   8550
      TabIndex        =   81
      Top             =   4215
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   119
      Left            =   8550
      TabIndex        =   80
      Top             =   3765
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   139
      Left            =   8550
      TabIndex        =   79
      Top             =   3330
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   159
      Left            =   8550
      TabIndex        =   78
      Top             =   2895
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   179
      Left            =   8550
      TabIndex        =   77
      Top             =   2445
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   199
      Left            =   8550
      TabIndex        =   76
      Top             =   2010
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   219
      Left            =   8550
      TabIndex        =   75
      Top             =   1575
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   239
      Left            =   8550
      TabIndex        =   74
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "259"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   259
      Left            =   8550
      TabIndex        =   73
      Top             =   690
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   25
      Left            =   1965
      TabIndex        =   72
      Top             =   5535
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   26
      Left            =   2445
      TabIndex        =   71
      Top             =   5535
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "247"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   27
      Left            =   2910
      TabIndex        =   70
      Top             =   5535
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "248"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   28
      Left            =   3375
      TabIndex        =   69
      Top             =   5535
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "249"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   29
      Left            =   3855
      TabIndex        =   68
      Top             =   5535
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   30
      Left            =   4320
      TabIndex        =   67
      Top             =   5535
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "251"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   31
      Left            =   4785
      TabIndex        =   66
      Top             =   5535
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "252"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   32
      Left            =   5265
      TabIndex        =   65
      Top             =   5535
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "5"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   1965
      TabIndex        =   64
      Top             =   5970
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "6"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   6
      Left            =   2445
      TabIndex        =   63
      Top             =   5970
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "7"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   7
      Left            =   2910
      TabIndex        =   62
      Top             =   5970
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "8"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   8
      Left            =   3375
      TabIndex        =   61
      Top             =   5970
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "9"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   9
      Left            =   3855
      TabIndex        =   60
      Top             =   5970
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "10"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   10
      Left            =   4320
      TabIndex        =   59
      Top             =   5970
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "11"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   11
      Left            =   4785
      TabIndex        =   58
      Top             =   5970
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "12"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   12
      Left            =   5265
      TabIndex        =   57
      Top             =   5970
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   45
      Left            =   1965
      TabIndex        =   56
      Top             =   5085
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   46
      Left            =   2445
      TabIndex        =   55
      Top             =   5085
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "247"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   47
      Left            =   2910
      TabIndex        =   54
      Top             =   5085
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "248"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   48
      Left            =   3375
      TabIndex        =   53
      Top             =   5085
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "249"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   49
      Left            =   3855
      TabIndex        =   52
      Top             =   5085
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   50
      Left            =   4320
      TabIndex        =   51
      Top             =   5085
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "251"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   51
      Left            =   4785
      TabIndex        =   50
      Top             =   5085
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "252"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   52
      Left            =   5265
      TabIndex        =   49
      Top             =   5085
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   65
      Left            =   1965
      TabIndex        =   48
      Top             =   4650
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   66
      Left            =   2445
      TabIndex        =   47
      Top             =   4650
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "247"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   67
      Left            =   2910
      TabIndex        =   46
      Top             =   4650
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "248"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   68
      Left            =   3375
      TabIndex        =   45
      Top             =   4650
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "249"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   69
      Left            =   3855
      TabIndex        =   44
      Top             =   4650
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   70
      Left            =   4320
      TabIndex        =   43
      Top             =   4650
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "251"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   71
      Left            =   4785
      TabIndex        =   42
      Top             =   4650
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "252"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   72
      Left            =   5265
      TabIndex        =   41
      Top             =   4650
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "260"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   260
      Left            =   9015
      TabIndex        =   40
      Top             =   690
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "240"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   240
      Left            =   9015
      TabIndex        =   39
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "220"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   220
      Left            =   9015
      TabIndex        =   38
      Top             =   1575
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "200"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   200
      Left            =   9015
      TabIndex        =   37
      Top             =   2010
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "180"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   180
      Left            =   9015
      TabIndex        =   36
      Top             =   2445
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "160"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   160
      Left            =   9015
      TabIndex        =   35
      Top             =   2895
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "140"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   140
      Left            =   9015
      TabIndex        =   34
      Top             =   3330
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "120"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   120
      Left            =   9015
      TabIndex        =   33
      Top             =   3765
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "100"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   100
      Left            =   9015
      TabIndex        =   32
      Top             =   4215
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "80"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   80
      Left            =   9015
      TabIndex        =   31
      Top             =   4650
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "60"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   60
      Left            =   9015
      TabIndex        =   30
      Top             =   5085
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "40"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   40
      Left            =   9015
      TabIndex        =   29
      Top             =   5535
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "20"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   20
      Left            =   9015
      TabIndex        =   28
      Top             =   5970
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "279"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   279
      Left            =   8550
      TabIndex        =   27
      Top             =   255
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "278"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   278
      Left            =   8085
      TabIndex        =   26
      Top             =   255
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "277"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   277
      Left            =   7605
      TabIndex        =   25
      Top             =   255
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "276"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   276
      Left            =   7140
      TabIndex        =   24
      Top             =   255
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "275"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   275
      Left            =   6675
      TabIndex        =   23
      Top             =   255
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "274"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   274
      Left            =   6195
      TabIndex        =   22
      Top             =   255
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "273"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   273
      Left            =   5730
      TabIndex        =   21
      Top             =   255
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "272"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   272
      Left            =   5265
      TabIndex        =   20
      Top             =   255
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "270"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   270
      Left            =   4320
      TabIndex        =   19
      Top             =   255
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "269"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   269
      Left            =   3855
      TabIndex        =   18
      Top             =   255
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "268"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   268
      Left            =   3375
      TabIndex        =   17
      Top             =   255
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "267"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   267
      Left            =   2910
      TabIndex        =   16
      Top             =   255
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "266"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   266
      Left            =   2445
      TabIndex        =   15
      Top             =   255
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "265"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   265
      Left            =   1965
      TabIndex        =   14
      Top             =   255
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "264"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   264
      Left            =   1500
      TabIndex        =   13
      Top             =   255
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "262"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   262
      Left            =   555
      TabIndex        =   12
      Top             =   255
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "261"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   261
      Left            =   90
      TabIndex        =   11
      Top             =   255
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "280"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   280
      Left            =   9015
      TabIndex        =   10
      Top             =   255
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "271"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   271
      Left            =   4785
      TabIndex        =   9
      Top             =   255
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "263"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   263
      Left            =   1035
      TabIndex        =   8
      Top             =   255
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "252"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   0
      Left            =   8985
      TabIndex        =   7
      Top             =   8295
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label SXOLIA2 
      Height          =   255
      Left            =   165
      TabIndex        =   4
      Top             =   6090
      Width           =   4035
   End
   Begin VB.Label sxolia1 
      Height          =   255
      Left            =   165
      TabIndex        =   3
      Top             =   6345
      Width           =   4035
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   375
      Index           =   0
      Left            =   75
      TabIndex        =   1
      Top             =   8055
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Top             =   6765
      Visible         =   0   'False
      Width           =   1875
   End
End
Attribute VB_Name = "FERE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldColor(1 To 300) As Single   ' Public array variable.
Dim kodikos(1 To 300)     ' Public array variable.
Dim Perigrafh(1 To 300)   ' Public array variable.
Dim LHJH(1 To 300)   ' Public array variable.
Dim Ypoloipo(1 To 300)   ' Public array variable.


Function fernei(Index As Integer)
   

End Function

Private Sub back_Click()


End Sub

Private Sub DBCombo1_Click(Area As Integer)
 Dim x
   x = DBCombo1.SelectedItem
   XIMITECH.Recordset.Bookmark = x
End Sub

Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Command2_Click()
' ΟΡΙΣΤΙΚΗ ΑΠΟΜΑΚΡΥΝΣΗ
Dim kod6400
  
If Label3(1).BackColor = &HFFFFFF Then    'αν δεν είναι κόκκινο το πρώτο κουτάκι ΕΠΕΣΤΡΕΨΕ
    Exit Sub
End If
  
For k = Ar_1_mpoykal To Ar_tel_mpoykal
   Label1(k).Enabled = False
Next
  
Command3.Enabled = False  'fere gia elegxo
  
  
PreHeating
  
'----------
DUM = begin_talos(0)
  If DUM = 0 Then
        response = MsgBox(mL_Res(520), 64, "Talos")  ' no communication
        Exit Sub
  End If

 'Do While Pause_Check() = 1
  '      response = MsgBox(mL_Res(521), 64, "Talos")  ' To continue turn <<PAUSE>> to position 1 and press <<ΟΚ>>."
'Loop

 Do While Is_Porta() = "α"
        response = MsgBox(mL_Res(522), 64, "Talos")    ' klise porta
Loop

'----------------

Telos_epilogon = 1 ' αν γίνει 1 τότε σταματάνε οι επιλογές μπουκαλιών

' ΠΑΤΑΕΙ ΠΡΩΤΗ ΦΟΡΑ ΤΟ ΚΟΥΜΠΙ
If Command1.Enabled = True Then
      

         MsgBox mL_Res(519)
         ' "Eίναι οι θέσεις ποτηριών ΕΛΕΥΘΕΡΕΣ ;", , "Talos" 'Eίναι οι θέσεις ποτηριών ΕΛΕΥΘΕΡΕΣ ;"
         Command2.Enabled = False
         Command1.Enabled = False
         front.Enabled = False

For k = 1 To ar_real_pothr
    If Label3(k).BackColor = &HFF& Then ' αν δεν είναι κόκκινη θέση φέρτο εδώ
         cmd$ = "!var12=" + str(6000 + Val(Label3(k).Caption)) + ":!var13=" + str(k) + ":"
         DUM = wait_to_understand(cmd$)
         DUM = Wait_Robot(0)
          Label3(k).BackColor = &HFFFF80
    End If
Next


Command2.Enabled = True
     Command2.Caption = mL_CapRes(524)   '"ΕΠΙΣΤΡΟΦΗ ΜΠΟΥΚΑΛΙΩΝ"
      '***** ΠΑΤΑ 2η ΦΟΡΑ ΤΟ ΚΟΥΜΠΙ ΝΑ ΤΑ ΠΑΕΙ ΠΙΣΩ
Else
    Command2.Enabled = False
    For k = 1 To ar_real_pothr
      If Label3(k).BackColor = &HFFFF80 Then ' αν  είναι κόκκινη θέση πήγαινε το μέσα
        cmd$ = "!var12=" + str(12000 + Val(Label3(k).Caption)) + ":!var13=" + str(k) + ":"
        DUM = wait_to_understand(cmd$)
         DUM = Wait_Robot(0)
        XIMITECH.Recordset.Seek "=", kodikos(Val(Label3(k).Caption))
        
        XIMITECH.Recordset.Edit
            XIMITECH.Recordset("ypol_PROT") = XIMITECH.Recordset("nerostox")
            XIMITECH.Recordset("LHJ_PROT") = Date + XIMITECH.Recordset("ZVH_PROT")
            XIMITECH.Recordset("real_cons") = 0
         XIMITECH.Recordset.update
        
        Label1(Val(Label3(k).Caption)).BackColor = &H8000&        '&H8000 ' OldColor(Label3(k).Caption)
        Label3(k).BackColor = &HFFFFFF  'ασπρο
      End If
    Next

     Command1.Enabled = True
     
     kod6400 = 82
     cmd$ = "!var12=" + str(kod6400) + ":"
     DUM = wait_to_understand(cmd$)
      DUM = Wait_Robot(0)
End If

End Sub

Private Sub Command3_Click()
' ΓΙΑ ΕΛΕΓΧΟ
Dim kod6400
  
If Label3(1).BackColor = &HFFFFFF Then    'αν δεν είναι κόκκινο το πρώτο κουτάκι ΕΠΕΣΤΡΕΨΕ
    Exit Sub
End If

  For k = Ar_1_mpoykal To Ar_tel_mpoykal
       Label1(k).Enabled = False
  Next



Command3.Enabled = False
Command3.Caption = mL_CapRes(525) ' "Επιστροφή στο τραπέζι"
  PreHeating
  
'----------
DUM = begin_talos(0)
  If DUM = 0 Then
        response = MsgBox(mL_Res(520), 64, "Talos")  ' no communication
        Exit Sub
  End If

 'Do While Pause_Check() = 1
  '      response = MsgBox(mL_Res(521), 64, "Talos(521)")  ' To continue turn <<PAUSE>> to position 1 and press <<ΟΚ>>."
'Loop

 Do While Is_Porta() = "α"
     response = MsgBox(mL_Res(522))    ' klise porta
Loop

'----------------


Telos_epilogon = 1 ' αν γίνει 1 τότε σταματάνε οι επιλογές μπουκαλιών

' ΠΑΤΑΕΙ ΠΡΩΤΗ ΦΟΡΑ ΤΟ ΚΟΥΜΠΙ
If Command1.Enabled = True Then
      
  MsgBox mL_Res(523), , "Talos" 'Eίναι οι θέσεις ποτηριών ΕΛΕΥΘΕΡΕΣ ;"

Command3.Enabled = False
Command2.Enabled = False
Command1.Enabled = False
front.Enabled = False
For k = 1 To ar_real_pothr
    If Label3(k).BackColor = &HFF& Then ' αν δεν είναι κόκκινη θέση φέρτο εδώ
        cmd$ = "!var12=" + str(6000 + Val(Label3(k).Caption)) + ":!var13=" + str(k) + ":"
        DUM = wait_to_understand(cmd$)
        DUM = Wait_Robot(0)
        Label3(k).BackColor = &HFFFF80
    End If

     
Next
   Command3.Enabled = True
   Command2.Caption = mL_CapRes(524) ' "ΕΠΙΣΤΡΟΦΗ ΜΠΟΥΚΑΛΙΩΝ"
    '***** ΠΑΤΑ 2η ΦΟΡΑ ΤΟ ΚΟΥΜΠΙ ΝΑ ΤΑ ΠΑΕΙ ΠΙΣΩ
Else
'
      
    MsgBox mL_Res(551), , "Talos"
   '551  "Eίναι οι θέσεις ΜΠΟΥΚΑΛΙΩΝ ΕΛΕΥΘΕΡΕΣ στο τραπέζι;"
    
    Command3.Enabled = False
     
    For k = 1 To ar_real_pothr
      If Label3(k).BackColor = &HFFFF80 Then ' αν  είναι κόκκινη θέση πήγαινε το μέσα
        cmd$ = "!var12=" + str(12000 + Val(Label3(k).Caption)) + ":!var13=" + str(k) + ":"
        DUM = wait_to_understand(cmd$)
         DUM = Wait_Robot(0)
    
        Label1(Val(Label3(k).Caption)).BackColor = &H8000&        '&H8000 ' OldColor(Label3(k).Caption)
        Label3(k).BackColor = &HFFFFFF  'ασπρο
      End If
    Next

     
     kod6400 = 82
    ' "ΠΑΡΚΑΡΙΣΕ "
     cmd$ = "!var12=" + str(kod6400) + ":"
     DUM = wait_to_understand(cmd$)
      DUM = Wait_Robot(0)
   Command3.Caption = mL_CapRes(511) ' "Φέρε Μπουκάλια για έλεγχο"
   
   Command1.Enabled = True

   
End If








End Sub

Private Sub check_ueseis()

Timer1.Enabled = True
DUM = begin_talos(0)


For k = 84 To 210
     If k <> 287 Then
        Command4.Caption = str(k)
        ' pane
        cmd$ = "!var12=" + str(12000 + k) + ":!var13=" + str(1) + ":"
        DUM = wait_to_understand(cmd$)
        DUM = Wait_Robot(0)
       
       'fere
       cmd$ = "!var12=" + str(6000 + k) + ":!var13=" + str(1) + ":"
        DUM = wait_to_understand(cmd$)
        DUM = Wait_Robot(0)
 End If
 
Next


End Sub

Private Sub check2_ueseis()

Timer1.Enabled = True
DUM = begin_talos(0)

  Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
           ' Open table.
  Set MyTable = mydb.OpenRecordset("select *from ximitech where addr_prot > 0;", dbOpenDynaset)
           
MsgBox mL_Res(518)
'518  "Βάλτε ένα μπουκάλι στην θέση ποτηριού 1"


'For k = 81 To 300
     If k <> 287 Then
        Command4.Caption = str(k)
        MyTable.MoveFirst
        yparxei = 0
        skonh = 0
        Do While Not MyTable.eof
             If k = MyTable("addr_prot") Then
                 yparxei = 1
                 If MyTable("skonh") = 1 Then
                    skonh = 1
                 End If
                 Exit Do
              End If
              MyTable.MoveNext
         Loop
      
      
      
      If yparxei = 0 Then
        ' pane απο 1 σε θέση
         cmd$ = "var11=0:!var12=" + str(12000 + k) + ":!var13=" + str(1) + ":"
         DUM = wait_to_understand(cmd$)
         DUM = Wait_Robot(0)
       
       
         'fere
         cmd$ = "!var12=" + str(6000 + k) + ":!var13=" + str(1) + ":"
         DUM = wait_to_understand(cmd$)
         DUM = Wait_Robot(0)
     End If
     
      If yparxei = 1 And skonh = 0 Then
        ' φερε μπουκάλι σε κεφαλές
         cmd$ = "!var11=1:!var12=" + str(1000 + k) + ":"
         DUM = wait_to_understand(cmd$)
         DUM = Wait_Robot(0)
       
       
         'επεστρεψε μπουκάλι από κεφαλές
         cmd$ = "!var12=" + str(2000 + k) + ":"
         DUM = wait_to_understand(cmd$)
         DUM = Wait_Robot(0)
         
     End If
     
     
     
     
     
     
     
       
 End If
 
End Sub

Private Sub Command4_Click()
check_ueseis
End Sub

Private Sub Form_DblClick()
 Dim a
     a = save_colors(Me, Me.BackColor)
End Sub

Private Sub Form_Load()
Dim aneb

If system_ready = 1 Then
      ValveReset
End If
a = find_colors(Me)

MDIForm1.Arxeia(10).Visible = False
MDIForm1.Syntages(20).Visible = False
MDIForm1.bohuhtika(30).Visible = False
MDIForm1.ejodos(40).Visible = False

 front.Caption = mL_CapRes(510)
 Command3.Caption = mL_CapRes(511)
 Command2.Caption = mL_CapRes(512)
 Command1.Caption = mL_CapRes(513)
 

 
' σχεδιάζω την νεκρή περιοχή  k_AR=0
  For k = 1 To Ar_1_mpoykal - 1
      Label1(k).Visible = False
  Next
  
  For k = Ar_tel_mpoykal + 1 To Ar_Mpoykal
      Label1(k).Visible = False
  Next
 
  aneb = 1075
  
  For k = ar_real_pothr + 1 To Ar_Pothr
      Label3(k).Visible = False
  Next
  
  For k = 1 To ar_real_pothr
      Label3(k).top = Label3(k).top - aneb '+ IIf(k Mod 2 <> 0, 0, -Label3(k).height / 2)
  Next
  
      
If Ar_1_mpoykal = 81 Then
  For k = Ar_1_mpoykal To Ar_tel_mpoykal Step 2
      Label1(k).top = Label1(k).top + Label1(k).height / 2
  Next
Else
  For k = Ar_1_mpoykal To Ar_tel_mpoykal Step 2
      Label1(k).top = Label1(k).top - Label1(k).height
  Next
  For k = Ar_1_mpoykal + 1 To Ar_tel_mpoykal Step 2
      Label1(k).top = Label1(k).top - Label1(k).height / 2
  Next



End If '
  
  
 L1 = Label1(K_AR + 1).left
 L6 = Label1(K_AR + 6).left
  
   SXOLIA1.left = L1
   SXOLIA2.left = L1
  
 For k = Ar_1_mpoykal To Ar_tel_mpoykal
   
     Label1(k).BackColor = &HFFFFFF   ' ασπρο
     OldColor(k) = Label1(k).BackColor
     Label1(k).Caption = k
  
 Next
     
  For k = 1 To ar_real_pothr
     Label3(k).Caption = k
     Label3(k).BackColor = &HFFFF00      '&HFFFFFF   ' ασπρο
     OldColor(k) = Label3(k).BackColor
 Next

     
End Sub


Private Sub Form_Paint()
Dim u
 XIMITECH.Recordset.Index = "KOD"
 
 While Not XIMITECH.Recordset.eof
         u = 0
         If Not IsNull(XIMITECH.Recordset("ADDR_PROT")) Then
                 u = XIMITECH.Recordset("addr_prot")
         End If
         
         If u >= Ar_1_mpoykal And u <= Ar_tel_mpoykal Then    ' =============================================
                                  
                 If Not IsNull(XIMITECH.Recordset("LHJ_PROT")) Then   ' ---------------
                          hmeres = DateDiff("d", Now, XIMITECH.Recordset("LHJ_PROT"))
                 Else
                          hmeres = 1000
                 End If
                 
                If XIMITECH.Recordset("skonh") = 1 Then ' ΣΤΕΡΕΑ
                                          Label1(u).BackColor = &H808080           '   καφε
                Else  '    Δ Ι Α Λ Υ Μ Α Τ Α
                      Label1(u).BackColor = FIND_COLOR(hmeres)
                       
                End If
               ' OldColor(u) = Label1(u).BackColor
                kodikos(u) = XIMITECH.Recordset("kod")
                Perigrafh(u) = XIMITECH.Recordset("perigrafh")
                LHJH(u) = XIMITECH.Recordset("Lhj_prot")
                Ypoloipo(u) = XIMITECH.Recordset("ypol_prot")
                
                'αν το υπόλοιπο είναι κάτω των 50 γραμμαρίων να κοκκινίζει
                 If Ypoloipo(u) < 50 Then
                       Label1(u).BackColor = &HFF&        'κοκκινο
                 End If
                 OldColor(u) = Label1(u).BackColor
                
                If Not IsNull(Perigrafh(u)) Then
                     Label1(u).Caption = Format(u, "###") + " " + Perigrafh(u)
                End If

         End If ' ================================================================
         XIMITECH.Recordset.MoveNext
 Wend
 Screen.MousePointer = 1


 
End Sub


Private Sub Form_Unload(Cancel As Integer)
MDIForm1.Arxeia(10).Visible = True
MDIForm1.Syntages(20).Visible = True
MDIForm1.bohuhtika(30).Visible = True
MDIForm1.ejodos(40).Visible = True
If system_ready = 1 Then
     ValveReset
End If


End Sub

Private Sub front_Click()
'ORISTIKI APOMAKRYNSI
Dim kod6400

 
If Label3(1).BackColor <> &HFF& Then    'αν δεν είναι κόκκινο το πρώτο κουτάκι ΕΠΕΣΤΡΕΨΕ
    Exit Sub
End If
 
 PreHeating
 
  For k = Ar_1_mpoykal To Ar_tel_mpoykal
       Label1(k).Enabled = False
  Next
Command3.Enabled = False 'fere gia elegxo
Command2.Enabled = False 'ANANEOSI
'Command1.Enabled = False 'EXODOS


DUM = begin_talos(0)
  If DUM = 0 Then
        response = MsgBox(mL_Res(520), 64, "Talos")  ' no communication
        Exit Sub
  End If

' Do While Pause_Check() = 1
 '       response = MsgBox(mL_Res(521), 64, "Talos(521)")  ' To continue turn <<PAUSE>> to position 1 and press <<ΟΚ>>."
 'Loop

 Do While Is_Porta() = "α"
        response = MsgBox(mL_Res(522), 64, "Talos")    ' klise porta
 Loop

MsgBox mL_Res(517), , "Talos"
' 517 "Eίναι οι θέσεις ποτηριών ΕΛΕΥΘΕΡΕΣ ;"

front.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command1.Enabled = False
For k = 1 To ar_real_pothr
  If Label3(k).BackColor = &HFF& Then ' αν δεν είναι κόκκινη θέση φέρτο εδώ
   cmd$ = "!var12=" + str(6000 + Val(Label3(k).Caption)) + ":!var13=" + str(k) + ":"
   DUM = wait_to_understand(cmd$)
   DUM = Wait_Robot(0)
     Label3(k).BackColor = &HFFFF80
     
   XIMITECH.Recordset.Seek "=", kodikos(Val(Label3(k).Caption))
   
   XIMITECH.Recordset.Edit
   XIMITECH.Recordset("ADDR_PROT") = 0  'Format(0, "###")
   XIMITECH.Recordset("real_cons") = 0
   XIMITECH.Recordset.update
End If

     
Next

     kod6400 = 82
    ' "ΠΑΡΚΑΡΙΣΕ "
     
     cmd$ = "!var12=" + str(kod6400) + ":"
     DUM = wait_to_understand(cmd$)
     
      DUM = Wait_Robot(0)

Command1.Enabled = True



End Sub

Private Sub Label1_Click(Index As Integer)
'   fernei (index)
'                                 ασπρο                                                     γκρίζο


If Label1(Index).BackColor = &HFFFFFF Or Label1(Index).BackColor = &H80FFFF Then    'ασπρο ή κίτρινο δεν κινείται
     '
ElseIf Label1(Index).BackColor = &H80C0FF Then ' αν είναι  γκρίζο να  γινει όπως ήταν
   
   ' το βγάζει από την θέση μπροστά
   For k = 1 To ar_real_pothr
      If Label3(k).Caption = Index Then            '  ξαναβάζω τον αρχικό του αριθμό
           If k + 1 <= ar_real_pothr Then
              If Label3(k + 1).BackColor = &HFF& Then
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
Else
   ' ψάχνει θέση για τοποθέτηση μπροστά
   For k = 1 To ar_real_pothr
      If Label3(k).BackColor <> &HFF& Then ' αν δεν είναι κόκκινη θέση φέρτο εδώ
          Label3(k).Caption = Index            'παίρνει τον αριθμό του μπουκαλιού που θάρθεi
          Label3(k).BackColor = &HFF&    'κοκκινίζω την μπροστά θέση

          Label1(Index).BackColor = &H80C0FF          'το αλλάζω χρωμα  για να μην ξαναζητηθεί
          front.Enabled = True
          Command2.Enabled = True
          Command3.Enabled = True
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
     SXOLIA1.Caption = Perigrafh(Index)
   '  Me.Caption = Perigrafh(Index)
 End If
If Not IsNull(Format(Ypoloipo(Index), "###") + "              " + Format(LHJH(Index), "dd/mm/yy")) Then
     SXOLIA2.Caption = Format(Ypoloipo(Index), "###") + "              " + Format(LHJH(Index), "dd/mm/yy")
   '  Me.Caption = Me.Caption + Format(Ypoloipo(Index), "###") + "              " + Format(Lhjh(Index), "dd/mm/yy")
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
        ' Me.Caption = Perigrafh(n)
      End If
    End If
    
    If Not IsNull(Format(Ypoloipo(n), "###") + "              " + Format(LHJH(n), "dd/mm/yy")) And Len(Trim(Me.Caption)) > 1 Then
        'sxolia2.Caption = Format(Ypoloipo(Index), "###") + "              " + Format(Lhjh(Index), "dd/mm/yy")
        'sxolia1.Caption = Trim(sxolia1.Caption) + mL_CapRes(584) + Format(Ypoloipo(n), "###") + mL_CapRes(583) + Format(Lhjh(n), "dd/mm/yy")
       SXOLIA1.Caption = Me.Caption + mL_CapRes(584) + Format(Ypoloipo(n), "###") + mL_CapRes(583) + Format(LHJH(n), "dd/mm/yy")
    End If
    
End If



End Sub


Private Sub Timer1_Timer()
   ' If Command2.Enabled = False Or (front.Enabled = False) Then
       If system_ready = 1 Then
            Elegxo_Katastash
        End If
   ' End If
End Sub

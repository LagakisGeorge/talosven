VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form frmSYNT 
   BackColor       =   &H00FF0000&
   ClientHeight    =   7080
   ClientLeft      =   1335
   ClientTop       =   2400
   ClientWidth     =   9480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7080
   ScaleWidth      =   9480
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdNERO 
      Caption         =   "NERO"
      Height          =   360
      Left            =   1920
      TabIndex        =   355
      Top             =   6360
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.PictureBox StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   9420
      TabIndex        =   354
      Top             =   6825
      Width           =   9480
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   5760
      TabIndex        =   352
      Top             =   5610
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Recipy Printing"
      Height          =   270
      Left            =   2160
      TabIndex        =   351
      Top             =   6000
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.Data PROSPAU2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\recipies.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6420
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from prospau2 where status=1 and epil>0  order by epil,entolh,aa_prospau,seira"
      Top             =   5235
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Data KINHSEIS 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\COORD_TB.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   8145
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "KINHSEIS"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.CommandButton Com_file_exec 
      Caption         =   "Execute Recipies"
      Height          =   540
      Left            =   1395
      TabIndex        =   329
      Top             =   5160
      Width           =   1665
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Exit"
      Height          =   540
      Left            =   3090
      TabIndex        =   11
      Top             =   5190
      Width           =   1050
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   4725
      Top             =   5745
   End
   Begin VB.Data XIMITECH 
      Caption         =   "XIMITECH"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\recipies.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   8145
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "XIMITECH"
      Top             =   5205
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1665
      TabIndex        =   10
      Text            =   "1"
      Top             =   6000
      Width           =   375
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "_SYNTAG.frx":0000
      Height          =   2130
      Left            =   -15
      OleObjectBlob   =   "_SYNTAG.frx":0017
      TabIndex        =   330
      Top             =   0
      Width           =   9675
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8250
      Top             =   5790
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   4275
      Top             =   5745
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "c:\talos\vis4-99.rpt"
      Destination     =   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3780
      Top             =   5790
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Ref 
      BackStyle       =   0  'Transparent
      Caption         =   "Ref:"
      Height          =   270
      Left            =   30
      TabIndex        =   353
      Top             =   6280
      Width           =   2500
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "299"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   299
      Left            =   8610
      TabIndex        =   350
      Top             =   -40
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "298"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   298
      Left            =   8145
      TabIndex        =   349
      Top             =   -40
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "297"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   297
      Left            =   7665
      TabIndex        =   348
      Top             =   -40
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "276"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   296
      Left            =   7200
      TabIndex        =   347
      Top             =   -40
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "295"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   295
      Left            =   6735
      TabIndex        =   346
      Top             =   -40
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "294"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   294
      Left            =   6255
      TabIndex        =   345
      Top             =   -40
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "293"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   293
      Left            =   5790
      TabIndex        =   344
      Top             =   -40
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "292"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   292
      Left            =   5325
      TabIndex        =   343
      Top             =   -40
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "290"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   290
      Left            =   4380
      TabIndex        =   342
      Top             =   -40
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "289"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   289
      Left            =   3915
      TabIndex        =   341
      Top             =   -40
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "288"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   288
      Left            =   3435
      TabIndex        =   340
      Top             =   -40
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "287"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   287
      Left            =   2970
      TabIndex        =   339
      Top             =   -40
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "286"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   286
      Left            =   2505
      TabIndex        =   338
      Top             =   -40
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "285"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   285
      Left            =   2025
      TabIndex        =   337
      Top             =   -40
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "284"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   284
      Left            =   1560
      TabIndex        =   336
      Top             =   -40
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "282"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   282
      Left            =   615
      TabIndex        =   335
      Top             =   -40
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "281"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   281
      Left            =   150
      TabIndex        =   334
      Top             =   -40
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "300"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   300
      Left            =   9075
      TabIndex        =   333
      Top             =   -40
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "291"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   291
      Left            =   4845
      TabIndex        =   332
      Top             =   -40
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "283"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   283
      Left            =   1095
      TabIndex        =   331
      Top             =   -40
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "263"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   263
      Left            =   1095
      TabIndex        =   328
      Top             =   375
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "271"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   271
      Left            =   4845
      TabIndex        =   327
      Top             =   375
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "280"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   280
      Left            =   9075
      TabIndex        =   326
      Top             =   375
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "261"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   261
      Left            =   150
      TabIndex        =   325
      Top             =   375
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "262"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   262
      Left            =   615
      TabIndex        =   324
      Top             =   375
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "264"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   264
      Left            =   1560
      TabIndex        =   323
      Top             =   375
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "265"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   265
      Left            =   2025
      TabIndex        =   322
      Top             =   375
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "266"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   266
      Left            =   2505
      TabIndex        =   321
      Top             =   375
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "267"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   267
      Left            =   2970
      TabIndex        =   320
      Top             =   375
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "268"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   268
      Left            =   3435
      TabIndex        =   319
      Top             =   375
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "269"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   269
      Left            =   3915
      TabIndex        =   318
      Top             =   375
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "270"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   270
      Left            =   4380
      TabIndex        =   317
      Top             =   375
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "272"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   272
      Left            =   5325
      TabIndex        =   316
      Top             =   375
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "273"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   273
      Left            =   5790
      TabIndex        =   315
      Top             =   375
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "274"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   274
      Left            =   6255
      TabIndex        =   314
      Top             =   375
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "275"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   275
      Left            =   6735
      TabIndex        =   313
      Top             =   375
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "276"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   276
      Left            =   7200
      TabIndex        =   312
      Top             =   375
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "277"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   277
      Left            =   7665
      TabIndex        =   311
      Top             =   375
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "278"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   278
      Left            =   8145
      TabIndex        =   310
      Top             =   375
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "279"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   279
      Left            =   8610
      TabIndex        =   309
      Top             =   375
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "20"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   20
      Left            =   9450
      TabIndex        =   308
      Top             =   7845
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "40"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   40
      Left            =   9450
      TabIndex        =   307
      Top             =   7410
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "60"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   60
      Left            =   9450
      TabIndex        =   306
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "80"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   80
      Left            =   9075
      TabIndex        =   305
      Top             =   4770
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "100"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   100
      Left            =   9075
      TabIndex        =   304
      Top             =   4335
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "120"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   120
      Left            =   9075
      TabIndex        =   303
      Top             =   3885
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "140"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   140
      Left            =   9075
      TabIndex        =   302
      Top             =   3450
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "160"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   160
      Left            =   9075
      TabIndex        =   301
      Top             =   3015
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "180"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   180
      Left            =   9075
      TabIndex        =   300
      Top             =   2565
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "200"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   200
      Left            =   9075
      TabIndex        =   299
      Top             =   2130
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "220"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   220
      Left            =   9075
      TabIndex        =   298
      Top             =   1695
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "240"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   240
      Left            =   9075
      TabIndex        =   297
      Top             =   1245
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "260"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   260
      Left            =   9075
      TabIndex        =   296
      Top             =   810
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "252"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   72
      Left            =   5325
      TabIndex        =   295
      Top             =   4770
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "251"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   71
      Left            =   4845
      TabIndex        =   294
      Top             =   4770
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   70
      Left            =   4380
      TabIndex        =   293
      Top             =   4770
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "249"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   69
      Left            =   3915
      TabIndex        =   292
      Top             =   4770
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "248"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   68
      Left            =   3435
      TabIndex        =   291
      Top             =   4770
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "247"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   67
      Left            =   2970
      TabIndex        =   290
      Top             =   4770
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   66
      Left            =   2505
      TabIndex        =   289
      Top             =   4770
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   65
      Left            =   2025
      TabIndex        =   288
      Top             =   4770
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "252"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   52
      Left            =   5700
      TabIndex        =   287
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "251"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   51
      Left            =   5220
      TabIndex        =   286
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   50
      Left            =   4755
      TabIndex        =   285
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "249"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   49
      Left            =   4290
      TabIndex        =   284
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "248"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   48
      Left            =   3810
      TabIndex        =   283
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "247"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   47
      Left            =   3345
      TabIndex        =   282
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   46
      Left            =   2880
      TabIndex        =   281
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   45
      Left            =   2400
      TabIndex        =   280
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "12"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   12
      Left            =   5700
      TabIndex        =   279
      Top             =   7845
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "11"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   11
      Left            =   5220
      TabIndex        =   278
      Top             =   7845
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "10"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   10
      Left            =   4755
      TabIndex        =   277
      Top             =   7845
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "9"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   9
      Left            =   4290
      TabIndex        =   276
      Top             =   7845
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "8"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   8
      Left            =   3810
      TabIndex        =   275
      Top             =   7845
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "7"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   7
      Left            =   3345
      TabIndex        =   274
      Top             =   7845
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "6"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   6
      Left            =   2880
      TabIndex        =   273
      Top             =   7845
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "5"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   2400
      TabIndex        =   272
      Top             =   7845
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "252"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   32
      Left            =   5700
      TabIndex        =   271
      Top             =   7410
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "251"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   31
      Left            =   5220
      TabIndex        =   270
      Top             =   7410
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   30
      Left            =   4755
      TabIndex        =   269
      Top             =   7410
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "249"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   29
      Left            =   4290
      TabIndex        =   268
      Top             =   7410
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "248"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   28
      Left            =   3810
      TabIndex        =   267
      Top             =   7410
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "247"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   27
      Left            =   3345
      TabIndex        =   266
      Top             =   7410
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   26
      Left            =   2880
      TabIndex        =   265
      Top             =   7410
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   25
      Left            =   2400
      TabIndex        =   264
      Top             =   7410
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "259"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   259
      Left            =   8610
      TabIndex        =   263
      Top             =   810
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   239
      Left            =   8610
      TabIndex        =   262
      Top             =   1245
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   219
      Left            =   8610
      TabIndex        =   261
      Top             =   1695
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   199
      Left            =   8610
      TabIndex        =   260
      Top             =   2130
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   179
      Left            =   8610
      TabIndex        =   259
      Top             =   2565
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   159
      Left            =   8610
      TabIndex        =   258
      Top             =   3015
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   139
      Left            =   8610
      TabIndex        =   257
      Top             =   3450
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   119
      Left            =   8610
      TabIndex        =   256
      Top             =   3885
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "99"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   99
      Left            =   8610
      TabIndex        =   255
      Top             =   4335
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   79
      Left            =   8610
      TabIndex        =   254
      Top             =   4770
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   59
      Left            =   8985
      TabIndex        =   253
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   39
      Left            =   8985
      TabIndex        =   252
      Top             =   7410
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "19"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   19
      Left            =   8985
      TabIndex        =   251
      Top             =   7845
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "258"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   258
      Left            =   8145
      TabIndex        =   250
      Top             =   810
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   238
      Left            =   8145
      TabIndex        =   249
      Top             =   1245
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   218
      Left            =   8145
      TabIndex        =   248
      Top             =   1695
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   198
      Left            =   8145
      TabIndex        =   247
      Top             =   2130
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   178
      Left            =   8145
      TabIndex        =   246
      Top             =   2565
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   158
      Left            =   8145
      TabIndex        =   245
      Top             =   3015
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   138
      Left            =   8145
      TabIndex        =   244
      Top             =   3450
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   118
      Left            =   8145
      TabIndex        =   243
      Top             =   3885
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "98"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   98
      Left            =   8145
      TabIndex        =   242
      Top             =   4335
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   78
      Left            =   8145
      TabIndex        =   241
      Top             =   4770
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   58
      Left            =   8520
      TabIndex        =   240
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   38
      Left            =   8520
      TabIndex        =   239
      Top             =   7410
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "18"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   18
      Left            =   8520
      TabIndex        =   238
      Top             =   7845
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "257"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   257
      Left            =   7665
      TabIndex        =   237
      Top             =   810
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   237
      Left            =   7665
      TabIndex        =   236
      Top             =   1245
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   217
      Left            =   7665
      TabIndex        =   235
      Top             =   1695
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   197
      Left            =   7665
      TabIndex        =   234
      Top             =   2130
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   177
      Left            =   7665
      TabIndex        =   233
      Top             =   2565
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   157
      Left            =   7665
      TabIndex        =   232
      Top             =   3015
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   137
      Left            =   7665
      TabIndex        =   231
      Top             =   3450
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   117
      Left            =   7665
      TabIndex        =   230
      Top             =   3885
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "97"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   97
      Left            =   7665
      TabIndex        =   229
      Top             =   4335
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   77
      Left            =   7665
      TabIndex        =   228
      Top             =   4770
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   57
      Left            =   8040
      TabIndex        =   227
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   37
      Left            =   8040
      TabIndex        =   226
      Top             =   7410
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "17"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   17
      Left            =   8040
      TabIndex        =   225
      Top             =   7845
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "256"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   256
      Left            =   7200
      TabIndex        =   224
      Top             =   810
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   236
      Left            =   7200
      TabIndex        =   223
      Top             =   1245
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   216
      Left            =   7200
      TabIndex        =   222
      Top             =   1695
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   196
      Left            =   7200
      TabIndex        =   221
      Top             =   2130
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   176
      Left            =   7200
      TabIndex        =   220
      Top             =   2565
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   156
      Left            =   7200
      TabIndex        =   219
      Top             =   3015
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   136
      Left            =   7200
      TabIndex        =   218
      Top             =   3450
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   116
      Left            =   7200
      TabIndex        =   217
      Top             =   3885
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "96"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   96
      Left            =   7200
      TabIndex        =   216
      Top             =   4335
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   76
      Left            =   7200
      TabIndex        =   215
      Top             =   4770
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   56
      Left            =   7575
      TabIndex        =   214
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   36
      Left            =   7575
      TabIndex        =   213
      Top             =   7410
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "16"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   16
      Left            =   7575
      TabIndex        =   212
      Top             =   7845
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "255"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   255
      Left            =   6735
      TabIndex        =   211
      Top             =   810
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   235
      Left            =   6735
      TabIndex        =   210
      Top             =   1245
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   215
      Left            =   6735
      TabIndex        =   209
      Top             =   1695
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   195
      Left            =   6735
      TabIndex        =   208
      Top             =   2130
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   175
      Left            =   6735
      TabIndex        =   207
      Top             =   2565
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   155
      Left            =   6735
      TabIndex        =   206
      Top             =   3015
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   135
      Left            =   6735
      TabIndex        =   205
      Top             =   3450
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   115
      Left            =   6735
      TabIndex        =   204
      Top             =   3885
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "95"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   95
      Left            =   6735
      TabIndex        =   203
      Top             =   4335
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "254"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   254
      Left            =   6255
      TabIndex        =   202
      Top             =   810
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   234
      Left            =   6255
      TabIndex        =   201
      Top             =   1245
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   214
      Left            =   6255
      TabIndex        =   200
      Top             =   1695
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   194
      Left            =   6255
      TabIndex        =   199
      Top             =   2130
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   174
      Left            =   6255
      TabIndex        =   198
      Top             =   2565
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   154
      Left            =   6255
      TabIndex        =   197
      Top             =   3015
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   134
      Left            =   6255
      TabIndex        =   196
      Top             =   3450
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   114
      Left            =   6255
      TabIndex        =   195
      Top             =   3885
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "94"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   94
      Left            =   6255
      TabIndex        =   194
      Top             =   4335
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "253"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   253
      Left            =   5790
      TabIndex        =   193
      Top             =   810
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   233
      Left            =   5790
      TabIndex        =   192
      Top             =   1245
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   213
      Left            =   5790
      TabIndex        =   191
      Top             =   1695
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   193
      Left            =   5790
      TabIndex        =   190
      Top             =   2130
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   173
      Left            =   5790
      TabIndex        =   189
      Top             =   2565
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   153
      Left            =   5790
      TabIndex        =   188
      Top             =   3015
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   133
      Left            =   5790
      TabIndex        =   187
      Top             =   3450
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   113
      Left            =   5790
      TabIndex        =   186
      Top             =   3885
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "93"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   93
      Left            =   5790
      TabIndex        =   185
      Top             =   4335
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "252"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   252
      Left            =   5325
      TabIndex        =   184
      Top             =   810
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   232
      Left            =   5325
      TabIndex        =   183
      Top             =   1245
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   212
      Left            =   5325
      TabIndex        =   182
      Top             =   1695
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   192
      Left            =   5325
      TabIndex        =   181
      Top             =   2130
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   172
      Left            =   5325
      TabIndex        =   180
      Top             =   2565
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   152
      Left            =   5325
      TabIndex        =   179
      Top             =   3015
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   132
      Left            =   5325
      TabIndex        =   178
      Top             =   3450
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   112
      Left            =   5325
      TabIndex        =   177
      Top             =   3885
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "92"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   92
      Left            =   5325
      TabIndex        =   176
      Top             =   4335
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "251"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   251
      Left            =   4845
      TabIndex        =   175
      Top             =   810
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   231
      Left            =   4845
      TabIndex        =   174
      Top             =   1245
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   211
      Left            =   4845
      TabIndex        =   173
      Top             =   1695
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   191
      Left            =   4845
      TabIndex        =   172
      Top             =   2130
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   171
      Left            =   4845
      TabIndex        =   171
      Top             =   2565
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   151
      Left            =   4845
      TabIndex        =   170
      Top             =   3015
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   131
      Left            =   4845
      TabIndex        =   169
      Top             =   3450
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   111
      Left            =   4845
      TabIndex        =   168
      Top             =   3885
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "91"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   91
      Left            =   4845
      TabIndex        =   167
      Top             =   4335
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   250
      Left            =   4380
      TabIndex        =   166
      Top             =   810
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   230
      Left            =   4380
      TabIndex        =   165
      Top             =   1245
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   210
      Left            =   4380
      TabIndex        =   164
      Top             =   1695
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   190
      Left            =   4380
      TabIndex        =   163
      Top             =   2130
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   170
      Left            =   4380
      TabIndex        =   162
      Top             =   2565
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   150
      Left            =   4380
      TabIndex        =   161
      Top             =   3015
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   130
      Left            =   4380
      TabIndex        =   160
      Top             =   3450
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   110
      Left            =   4380
      TabIndex        =   159
      Top             =   3885
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "90"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   90
      Left            =   4380
      TabIndex        =   158
      Top             =   4335
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "249"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   249
      Left            =   3915
      TabIndex        =   157
      Top             =   810
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   229
      Left            =   3915
      TabIndex        =   156
      Top             =   1245
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   209
      Left            =   3915
      TabIndex        =   155
      Top             =   1695
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   189
      Left            =   3915
      TabIndex        =   154
      Top             =   2130
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   169
      Left            =   3915
      TabIndex        =   153
      Top             =   2565
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   149
      Left            =   3915
      TabIndex        =   152
      Top             =   3015
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   129
      Left            =   3915
      TabIndex        =   151
      Top             =   3450
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   109
      Left            =   3915
      TabIndex        =   150
      Top             =   3885
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "89"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   89
      Left            =   3915
      TabIndex        =   149
      Top             =   4335
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "248"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   248
      Left            =   3435
      TabIndex        =   148
      Top             =   810
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   228
      Left            =   3435
      TabIndex        =   147
      Top             =   1245
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   208
      Left            =   3435
      TabIndex        =   146
      Top             =   1695
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   188
      Left            =   3435
      TabIndex        =   145
      Top             =   2130
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   168
      Left            =   3435
      TabIndex        =   144
      Top             =   2565
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   148
      Left            =   3435
      TabIndex        =   143
      Top             =   3015
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   128
      Left            =   3435
      TabIndex        =   142
      Top             =   3450
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   108
      Left            =   3435
      TabIndex        =   141
      Top             =   3885
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "88"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   88
      Left            =   3435
      TabIndex        =   140
      Top             =   4335
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "247"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   247
      Left            =   2970
      TabIndex        =   139
      Top             =   810
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   227
      Left            =   2970
      TabIndex        =   138
      Top             =   1245
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   207
      Left            =   2970
      TabIndex        =   137
      Top             =   1695
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   187
      Left            =   2970
      TabIndex        =   136
      Top             =   2130
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   167
      Left            =   2970
      TabIndex        =   135
      Top             =   2565
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   147
      Left            =   2970
      TabIndex        =   134
      Top             =   3015
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   127
      Left            =   2970
      TabIndex        =   133
      Top             =   3450
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   107
      Left            =   2970
      TabIndex        =   132
      Top             =   3885
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "87"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   87
      Left            =   2970
      TabIndex        =   131
      Top             =   4335
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "75"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   75
      Left            =   6735
      TabIndex        =   130
      Top             =   4770
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "55"
      ForeColor       =   &H0000002F&
      Height          =   375
      Index           =   55
      Left            =   7110
      TabIndex        =   129
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "35"
      ForeColor       =   &H0000001B&
      Height          =   375
      Index           =   35
      Left            =   7110
      TabIndex        =   128
      Top             =   7410
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "15"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   15
      Left            =   7110
      TabIndex        =   127
      Top             =   7845
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   246
      Left            =   2505
      TabIndex        =   126
      Top             =   810
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   226
      Left            =   2505
      TabIndex        =   125
      Top             =   1245
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   206
      Left            =   2505
      TabIndex        =   124
      Top             =   1695
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   186
      Left            =   2505
      TabIndex        =   123
      Top             =   2130
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   166
      Left            =   2505
      TabIndex        =   122
      Top             =   2565
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   146
      Left            =   2505
      TabIndex        =   121
      Top             =   3015
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   126
      Left            =   2505
      TabIndex        =   120
      Top             =   3450
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   106
      Left            =   2505
      TabIndex        =   119
      Top             =   3885
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "86"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   86
      Left            =   2505
      TabIndex        =   118
      Top             =   4335
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "74"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   74
      Left            =   6255
      TabIndex        =   117
      Top             =   4770
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "54"
      ForeColor       =   &H0000002E&
      Height          =   375
      Index           =   54
      Left            =   6630
      TabIndex        =   116
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "34"
      ForeColor       =   &H0000001A&
      Height          =   375
      Index           =   34
      Left            =   6630
      TabIndex        =   115
      Top             =   7410
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "14"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   14
      Left            =   6630
      TabIndex        =   114
      Top             =   7845
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   245
      Left            =   2025
      TabIndex        =   113
      Top             =   810
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   225
      Left            =   2025
      TabIndex        =   112
      Top             =   1245
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   205
      Left            =   2025
      TabIndex        =   111
      Top             =   1695
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   185
      Left            =   2025
      TabIndex        =   110
      Top             =   2130
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   165
      Left            =   2025
      TabIndex        =   109
      Top             =   2565
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   145
      Left            =   2025
      TabIndex        =   108
      Top             =   3015
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   125
      Left            =   2025
      TabIndex        =   107
      Top             =   3450
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   105
      Left            =   2025
      TabIndex        =   106
      Top             =   3885
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "85"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   85
      Left            =   2025
      TabIndex        =   105
      Top             =   4335
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "73"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   73
      Left            =   5790
      TabIndex        =   104
      Top             =   4770
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "53"
      ForeColor       =   &H0000002D&
      Height          =   375
      Index           =   53
      Left            =   6165
      TabIndex        =   103
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "33"
      ForeColor       =   &H00000019&
      Height          =   375
      Index           =   33
      Left            =   6165
      TabIndex        =   102
      Top             =   7410
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "13"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   13
      Left            =   6165
      TabIndex        =   101
      Top             =   7845
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   244
      Left            =   1560
      TabIndex        =   100
      Top             =   810
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   224
      Left            =   1560
      TabIndex        =   99
      Top             =   1245
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   204
      Left            =   1560
      TabIndex        =   98
      Top             =   1695
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   184
      Left            =   1560
      TabIndex        =   97
      Top             =   2130
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   164
      Left            =   1560
      TabIndex        =   96
      Top             =   2565
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   144
      Left            =   1560
      TabIndex        =   95
      Top             =   3015
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   124
      Left            =   1560
      TabIndex        =   94
      Top             =   3450
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   104
      Left            =   1560
      TabIndex        =   93
      Top             =   3885
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "84"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   84
      Left            =   1560
      TabIndex        =   92
      Top             =   4335
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000040&
      Height          =   375
      Index           =   64
      Left            =   1560
      TabIndex        =   91
      Top             =   4770
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H0000002C&
      Height          =   375
      Index           =   44
      Left            =   1935
      TabIndex        =   90
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000018&
      Height          =   375
      Index           =   24
      Left            =   1935
      TabIndex        =   89
      Top             =   7410
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "4"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   1935
      TabIndex        =   88
      Top             =   7845
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   243
      Left            =   1095
      TabIndex        =   87
      Top             =   810
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   223
      Left            =   1095
      TabIndex        =   86
      Top             =   1245
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   203
      Left            =   1095
      TabIndex        =   85
      Top             =   1695
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   183
      Left            =   1095
      TabIndex        =   84
      Top             =   2130
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   163
      Left            =   1095
      TabIndex        =   83
      Top             =   2565
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   143
      Left            =   1095
      TabIndex        =   82
      Top             =   3015
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   123
      Left            =   1095
      TabIndex        =   81
      Top             =   3450
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   103
      Left            =   1095
      TabIndex        =   80
      Top             =   3885
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "83"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   83
      Left            =   1095
      TabIndex        =   79
      Top             =   4335
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H0000003F&
      Height          =   375
      Index           =   63
      Left            =   1095
      TabIndex        =   78
      Top             =   4770
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H0000002B&
      Height          =   375
      Index           =   43
      Left            =   1470
      TabIndex        =   77
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000017&
      Height          =   375
      Index           =   23
      Left            =   1470
      TabIndex        =   76
      Top             =   7410
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   1470
      TabIndex        =   75
      Top             =   7845
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   242
      Left            =   615
      TabIndex        =   74
      Top             =   810
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   222
      Left            =   615
      TabIndex        =   73
      Top             =   1245
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   202
      Left            =   615
      TabIndex        =   72
      Top             =   1695
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   182
      Left            =   615
      TabIndex        =   71
      Top             =   2130
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   162
      Left            =   615
      TabIndex        =   70
      Top             =   2565
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   142
      Left            =   615
      TabIndex        =   69
      Top             =   3015
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   122
      Left            =   615
      TabIndex        =   68
      Top             =   3450
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   102
      Left            =   615
      TabIndex        =   67
      Top             =   3885
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "82"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   82
      Left            =   615
      TabIndex        =   66
      Top             =   4335
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H0000003E&
      Height          =   375
      Index           =   62
      Left            =   615
      TabIndex        =   65
      Top             =   4770
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H0000002A&
      Height          =   375
      Index           =   42
      Left            =   990
      TabIndex        =   64
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000016&
      Height          =   375
      Index           =   22
      Left            =   990
      TabIndex        =   63
      Top             =   7410
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2"
      ForeColor       =   &H00000002&
      Height          =   375
      Index           =   2
      Left            =   990
      TabIndex        =   62
      Top             =   7845
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "241"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   241
      Left            =   150
      TabIndex        =   61
      Top             =   810
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "221"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   221
      Left            =   150
      TabIndex        =   60
      Top             =   1245
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "201"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   201
      Left            =   150
      TabIndex        =   59
      Top             =   1695
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "181"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   181
      Left            =   150
      TabIndex        =   58
      Top             =   2130
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "161"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   161
      Left            =   150
      TabIndex        =   57
      Top             =   2565
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "141"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   141
      Left            =   150
      TabIndex        =   56
      Top             =   3015
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "121"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   121
      Left            =   150
      TabIndex        =   55
      Top             =   3450
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "101"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   101
      Left            =   150
      TabIndex        =   54
      Top             =   3885
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "81"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   81
      Left            =   150
      TabIndex        =   53
      Top             =   4335
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "61"
      ForeColor       =   &H0000003D&
      Height          =   375
      Index           =   61
      Left            =   150
      TabIndex        =   52
      Top             =   4770
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "41"
      ForeColor       =   &H00000029&
      Height          =   375
      Index           =   41
      Left            =   525
      TabIndex        =   51
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "21"
      ForeColor       =   &H00000015&
      Height          =   375
      Index           =   21
      Left            =   525
      TabIndex        =   50
      Top             =   7410
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1"
      ForeColor       =   &H00000001&
      Height          =   375
      Index           =   1
      Left            =   525
      TabIndex        =   49
      Top             =   7845
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "27"
      Height          =   375
      Index           =   27
      Left            =   7200
      TabIndex        =   48
      Top             =   6285
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "26"
      Height          =   375
      Index           =   26
      Left            =   6735
      TabIndex        =   47
      Top             =   6285
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "25"
      Height          =   375
      Index           =   25
      Left            =   6255
      TabIndex        =   46
      Top             =   6285
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "24"
      Height          =   375
      Index           =   24
      Left            =   5790
      TabIndex        =   45
      Top             =   6285
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "23"
      Height          =   375
      Index           =   23
      Left            =   5325
      TabIndex        =   44
      Top             =   6285
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "22"
      Height          =   375
      Index           =   22
      Left            =   4845
      TabIndex        =   43
      Top             =   6285
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "21"
      Height          =   375
      Index           =   21
      Left            =   4380
      TabIndex        =   42
      Top             =   6285
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "17"
      Height          =   375
      Index           =   17
      Left            =   7200
      TabIndex        =   41
      Top             =   6735
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "16"
      Height          =   375
      Index           =   16
      Left            =   6735
      TabIndex        =   40
      Top             =   6735
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "15"
      Height          =   375
      Index           =   15
      Left            =   6255
      TabIndex        =   39
      Top             =   6735
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "14"
      Height          =   375
      Index           =   14
      Left            =   5790
      TabIndex        =   38
      Top             =   6735
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "12"
      Height          =   375
      Index           =   12
      Left            =   4845
      TabIndex        =   37
      Top             =   6735
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "11"
      Height          =   375
      Index           =   11
      Left            =   4380
      TabIndex        =   36
      Top             =   6735
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "7"
      Height          =   375
      Index           =   7
      Left            =   7200
      TabIndex        =   35
      Top             =   7185
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "6"
      Height          =   375
      Index           =   6
      Left            =   6735
      TabIndex        =   34
      Top             =   7185
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "5"
      Height          =   375
      Index           =   5
      Left            =   6255
      TabIndex        =   33
      Top             =   7185
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "4"
      Height          =   375
      Index           =   4
      Left            =   5790
      TabIndex        =   32
      Top             =   7185
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "3"
      Height          =   375
      Index           =   3
      Left            =   5325
      TabIndex        =   31
      Top             =   7185
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "2"
      Height          =   375
      Index           =   2
      Left            =   4845
      TabIndex        =   30
      Top             =   7185
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "1"
      Height          =   375
      Index           =   1
      Left            =   4380
      TabIndex        =   29
      Top             =   7185
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "13"
      Height          =   375
      Index           =   13
      Left            =   5325
      TabIndex        =   28
      Top             =   6735
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "8"
      Height          =   375
      Index           =   8
      Left            =   7665
      TabIndex        =   27
      Top             =   7185
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "9"
      Height          =   375
      Index           =   9
      Left            =   8145
      TabIndex        =   26
      Top             =   7185
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "10"
      Height          =   375
      Index           =   10
      Left            =   8610
      TabIndex        =   25
      Top             =   7185
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "18"
      Height          =   375
      Index           =   18
      Left            =   7665
      TabIndex        =   24
      Top             =   6735
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "19"
      Height          =   375
      Index           =   19
      Left            =   8145
      TabIndex        =   23
      Top             =   6735
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "20"
      Height          =   375
      Index           =   20
      Left            =   8610
      TabIndex        =   22
      Top             =   6735
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "28"
      Height          =   375
      Index           =   28
      Left            =   7665
      TabIndex        =   21
      Top             =   6285
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "29"
      Height          =   375
      Index           =   29
      Left            =   8145
      TabIndex        =   20
      Top             =   6285
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "30"
      Height          =   375
      Index           =   30
      Left            =   8610
      TabIndex        =   19
      Top             =   6285
      Width           =   375
   End
   Begin VB.Label Label19 
      Caption         =   ".Εκκίνηση από ποτήρι"
      Height          =   267
      Left            =   30
      TabIndex        =   18
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   165
      TabIndex        =   17
      Top             =   9380
      Width           =   9285
   End
   Begin VB.Label Label14 
      Caption         =   "Καθ"
      Height          =   255
      Left            =   15
      TabIndex        =   16
      Top             =   5700
      Width           =   495
   End
   Begin VB.Label Label13 
      Caption         =   "Μικτ"
      Height          =   210
      Left            =   15
      TabIndex        =   15
      Top             =   5175
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "Απόβ"
      Height          =   210
      Left            =   15
      TabIndex        =   14
      Top             =   5445
      Width           =   495
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF00&
      Height          =   210
      Left            =   570
      TabIndex        =   13
      Top             =   5445
      Width           =   735
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   570
      TabIndex        =   12
      Top             =   5700
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Height          =   240
      Left            =   915
      TabIndex        =   9
      Top             =   6855
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Gripper"
      Height          =   195
      Left            =   9690
      TabIndex        =   8
      Top             =   3300
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "pow der"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9750
      TabIndex        =   7
      Top             =   4125
      Visible         =   0   'False
      Width           =   375
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label12 
      BackColor       =   &H0000FFFF&
      Caption         =   "Chem  icals"
      ForeColor       =   &H0000005A&
      Height          =   375
      Left            =   9750
      TabIndex        =   6
      Top             =   3615
      Visible         =   0   'False
      Width           =   375
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Ζυγαριά"
      Height          =   210
      Left            =   9555
      TabIndex        =   0
      Top             =   2910
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label4 
      Height          =   105
      Index           =   0
      Left            =   9600
      TabIndex        =   5
      Top             =   2670
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label SXOLIA2 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2355
      TabIndex        =   4
      Top             =   7425
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label sxolia1 
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   105
      TabIndex        =   3
      Top             =   6870
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Height          =   180
      Index           =   0
      Left            =   570
      TabIndex        =   2
      Top             =   5190
      Width           =   735
   End
   Begin VB.Label Label2 
      Height          =   210
      Left            =   9600
      TabIndex        =   1
      Top             =   2340
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmSYNT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "φορμα συνταγων"

Option Explicit
Const TREXEI = 99999
Const DEN_BGAZEI = 500000

Const BLE = &HFFFF00       '&HFF0000
Const ASPRO = &HFFFFFF
Const MAYRO = &H0&
Const kokkino = &HFF&

Const ROZ = &H8080FF

Const Kafe = &H808080 ' &H404040
Const Keramidi = &H80C0FF
Const Prasino = &H80FF&
Const Portokali = &H80FF&

Const gripper_anoiktos = 0
Const gripper_kleistos = 1

Public General_Xronos
'VBREGS Flags bit values
Public xymeno As Single
Public baros As Single
Public xronos    '    , x0, y0, z0, w0
Public valve_closed
Public gINDEX
Public m_CC   ' ΑΡΙΘΜΟΣ ΣΥΝΤΑΓΩΝ
Public kod6400
Public kodTimer
Public EPIU
Public Flag_Show
Public Flag_Kefal
Public porta As String
Public a As Single
Public ZYG, Real_Alati

Public flag_paint
Public Pothri
Public Exo_entolh
Public cmd2$
Public var65_1kef
Public var66_2kef
Public var67_3kef
Public var68_4kef
Public var93_anoZ
Public var74, var88
Public AA_SYNTAG
Public flag As Integer
Public gram As Single
Dim OldColor(1 To 300) As Single   ' Public array variable.
Dim kodikos(1 To 300)     ' Public array variable.
Dim Perigrafh(1 To 300)   ' Public array variable.
Dim LHJH(1 To 300)   ' Public array variable.
Dim Ypoloipo(1 To 300)    ' Public array variable.
Dim mTop(1 To 300)   ' Public array variable.
Dim mLeft(1 To 300)   ' Public array variable.
Dim pTop(1 To 300)   ' Public array variable.
Dim pLeft(1 To 300)   ' Public array variable.
Dim XHM(1 To 300)
Dim gr(1 To 300)
Dim per(1 To 300)
Dim ADDR(1 To 300)
Dim Antistoix(1 To 300)
Dim APOBARO
Dim XALASE_NERO
Dim pin_Syntag(1 To Ar_Pothr)
Dim Feygei
Dim Erxetai
Dim L1
Dim DOSEIS(1 To 30)   ' ΧΡΕΙΑ΄ΖΟΝΤΑΙ ΓΙΑ ΤΙΣ ΔΟΚΙΜΕΣ ΔΟΣΟΜΕΤΡΗΣΕΩΝ
Dim STADIA(1 To 30) ' ΧΡΕΙΑ΄ΖΟΝΤΑΙ ΓΙΑ ΤΙΣ ΔΟΚΙΜΕΣ ΔΟΣΟΜΕΤΡΗΣΕΩΝ
Dim katofli(1 To 300)
Dim blabes(1 To 10)
Dim tkin(1 To 10)
Dim KEF(1 To 2), KEFALH
Dim m_syntag(1 To Ar_Pothr)

Dim a_Syntagh(1 To 5)
Dim a_Fere(1 To 5)
Dim a_Kinhsh(1 To 5)
Dim a_Kefalh(1 To 5)
Dim a_Mgrs(1 To 5)
Dim a_Pane(1 To 5)
Dim a_Xronos(1 To 5)
Dim a_Tkinhsh(1 To 5)
Dim a_Kef1(1 To 5)
Dim a_Kef2(1 To 5)
Dim a_Ar_Syntag(1 To 5)
Dim Ektelesueises(1 To 40)
Dim m_ximika(1 To 50, 1 To 5)
Dim a00x(1 To 5), b00x(1 To 5)
Dim f_poso_apor(1 To 5)  'apo poy kai pano theoreitai apotyxhmenh
Dim Water_On, Water_Off



Function calc_ximika()
'-------------------------------------------------
Dim mydb, r, counter
Dim use_valve(1 To 5) As Integer
Dim mepil, r2, k, L

calc_ximika = 1 'OK

For L = 1 To aritmos_ximikon: use_valve(L) = 0: Next

For k = 1 To 50
   For L = 1 To aritmos_ximikon
       m_ximika(k, L) = 0
   Next
Next

'     Exit Sub

Set mydb = OpenDatabase("c:\talos\recipies.mdb")
Set r = mydb.OpenRecordset("select *from prospau2 where status=1 and epil>0")
r.MoveFirst
Set r2 = mydb.OpenRecordset("XIMITECH")
r2.Index = "KOD"

Do While Not r.eof
    r2.Seek "=", r("kod")
    If Not r2.NoMatch Then
       If r2("addr_prot") <= 51 + aritmos_ximikon And r2("addr_prot") >= 51 Then
          If r("epil") > 0 And r("epil") <= 50 Then
             m_ximika(r("epil"), r2("addr_prot") - 51) = r("GR")
             use_valve(r2("addr_prot") - 51) = 1 ' i use this valve
             If left(r("perigr"), 2) = "~~" Then
             Else
                r.EDIT
                r("perigr") = "~~" + left(r("perigr"), 20)
                r("apor") = r2("addr_prot") - 51
                r.update
             End If
          Else
             If r("epil") > 50 Then
                MsgBox mLResnew(934, "Not dispense chemicals for selection >50 ", 1)
             End If
          End If
       End If
   End If
   r.MoveNext
Loop

r.Close
r2.Close



Set r = mydb.OpenRecordset("parametroi")
r.MoveFirst
r.MoveNext
r.MoveNext
If IsNull(r("Colormatchdir")) Or IsNull(r("colormatchfilename")) Then
       r.Close
       MsgBox mLResnew(933, "Calibrate valves for chemicals*", 1)
       Exit Function
Else
       a00x(5) = Val(r("Colormatchdir")) 'nero
       b00x(5) = Val(r("colormatchfilename"))
       If IsEmpty(a00x(5)) Or a00x(5) = 0 Then calc_ximika = 0 'OK
End If

For k = 1 To aritmos_ximikon
    r.MoveNext
    If IsNull(r("Colormatchdir")) Or IsNull(r("colormatchfilename")) Then
       r.Close
       MsgBox mLResnew(933, "Calibrate valves for chemicals*", 1)
       Exit Function
    Else
       a00x(k) = Val(r("Colormatchdir"))
       b00x(k) = Val(r("colormatchfilename"))
       If IsEmpty(a00x(k)) Or a00x(k) = 0 Then
          If use_valve(k) = 1 Then
             calc_ximika = 0 'OK
          End If
       End If
    End If
Next
r.Close

' Set db = OpenDatabase("c:\talos\recipies.mdb")


prospau2.RecordSource = "select * from prospau2 where status=1 and epil>0 and not perigr like '~~*' order by epil,entolh,aa_prospau,seira"
prospau2.Refresh



End Function




Sub dispense_old_Ximika()

Dim dr As Long, c, quantity, a00(1 To 5), b00(1 To 5), a2
Dim mepil, xr(1 To 5), for_use, dokXr(1 To 5), sXr(1 To 5)
Dim k, L, megas, amax, Off_On(1 To 5), per
Dim max_xronos, mz, mZ1, mz2
Dim sum_gr, q_ner, xr_ner

Dim Max_apoklish

Max_apoklish = 500 'mgr
' m_ximika(epil, 1 To 5)  posothtes

max_xronos = 50000
For k = 1 To 5: xr(k) = 0: Off_On(k) = 0: Next



If IsNull(prospau2.Recordset("epil")) Then Exit Sub
If Not IsNumeric(prospau2.Recordset("epil")) Then Exit Sub
If prospau2.Recordset("epil") > 50 Then Exit Sub

 mepil = Int(prospau2.Recordset("epil"))

'Quantity = xr
for_use = 0
sum_gr = 0
For k = 1 To aritmos_ximikon
  If m_ximika(mepil, k) > 0 Then
      
      sum_gr = sum_gr + m_ximika(mepil, k)
      
      If m_ximika(mepil, k) < 1600 Then
         m_ximika(mepil, k) = m_ximika(mepil, k) - 70
         for_use = for_use + 1
      End If
      
      If a00x(k) = 0 Then
         xr(k) = 0
      Else
         xr(k) = (m_ximika(mepil, k) - b00x(k)) / a00x(k) 'xrhsimopoionta 2, 0.5 sec
      End If
      
  End If
Next
 


If Not IsNull(prospau2.Recordset("nero")) Then  ' Exit Sub
  
  q_ner = prospau2.Recordset("nero")
  sum_gr = sum_gr + q_ner
  If a00x(5) = 0 Then
      xr_ner = 0
  Else
      xr_ner = (q_ner - b00x(5)) / a00x(5) 'xrhsimopoionta 2, 0.5 sec
  End If
  
End If
 
 
's[-arr
For k = 1 To 4: dokXr(k) = xr(k): Next

For L = 1 To 4
    amax = -2
    megas = 1
   For k = 1 To 4
      If dokXr(k) > amax Then
         amax = dokXr(k)
         megas = k
      End If
   Next
   dokXr(megas) = -1
   sXr(L) = amax
Next
 
 
 
 MilSec 1000
 
 
 
 
mz = Zygis0(0)
Do While Abs(mz) > 30
'  G_REZERO MSComm1
  
If Balance_Type = "ADAM" Then
      MSComm1.Output = "ZERO" + Chr(13)
Else
      MSComm1.Output = "Z" + Chr$(13) + Chr(10)
End If
 
 'MSComm1.Output = Zero
  
  
'    MSComm1.Output = "Z " + Chr$(13) + Chr(10)
     MilSec 200
     mz = Zygis0(0)
Loop





c = GetCurrentTime()

If aritmos_ximikon >= 1 Then        ' nero
    If xr(1) > 0 Then
       RobSend ("!" + Xim1_on + ":")
       Off_On(1) = 1
    End If
End If

If aritmos_ximikon >= 2 Then
   If xr(2) > 0 Then
      RobSend ("!" + Xim2_on + ":")
      Off_On(2) = 1
   End If
End If
  
If aritmos_ximikon >= 3 Then
   If xr(3) > 0 Then
      RobSend ("!" + Xim3_on + ":")
      Off_On(3) = 1
   End If
End If
 
If aritmos_ximikon >= 4 Then
   If xr(4) > 0 Then
      RobSend ("!" + Xim4_on + ":")
      Off_On(4) = 1
   End If
End If
 
If XimNer = 1 Then
   If xr_ner > 0 Then
      RobSend ("!" + XimNer_on + ":")
      Off_On(5) = 1
   End If
End If
 
'=====================================

If xr_ner > sXr(1) Then    '  sxr(1)=3    xr_ner=27   max_xronos=13
   sXr(1) = xr_ner
End If



If sXr(1) > max_xronos Then
   sXr(1) = max_xronos
End If


Do While GetCurrentTime() - c < sXr(1)
     DoEvents
     per = GetCurrentTime() - c
           
     If Off_On(1) = 1 Then
        If per > xr(1) Then
           RobSend ("!" + xim1_Off + ":")
           Off_On(1) = 0
        End If
     End If
     
     If Off_On(2) = 1 Then
        If per > xr(2) Then
           RobSend ("!" + xim2_Off + ":")
           Off_On(2) = 0
        End If
     End If
     If Off_On(3) = 1 Then
        If per > xr(3) Then
           RobSend ("!" + xim3_Off + ":")
           Off_On(3) = 0
        End If
     End If
     If Off_On(4) = 1 Then
        If per > xr(4) Then
           RobSend ("!" + xim4_Off + ":")
           Off_On(4) = 0
        End If
     End If
     
     If Off_On(5) = 1 Then
        If per > xr_ner Then
           RobSend ("!" + XimNer_Off + ":")
           Off_On(5) = 0
        End If
     End If

Loop

'---------------------------------------------
' τα  ξανακλεινω για λόγους ασφαλείας
If aritmos_ximikon >= 1 Then RobSend ("!" + xim1_Off + ":")
If aritmos_ximikon >= 2 Then RobSend ("!" + xim2_Off + ":")
If aritmos_ximikon >= 3 Then RobSend ("!" + xim3_Off + ":")
If aritmos_ximikon >= 4 Then RobSend ("!" + xim4_Off + ":")
If XimNer = 1 Then RobSend ("!" + XimNer_Off + ":")
  
  MilSec 2000
        
dispense_ximika2:
mZ1 = Zygis0(0)
mz2 = Zygis0(0)
If Abs(mZ1 - mz2) > 10 Then GoTo dispense_ximika2
mz = (mZ1 + mz2) / 2

If Abs(mz - sum_gr) <= Max_apoklish Then
   'ok
Else
   m_ximika(50, 5) = -31 ' Ektel_kinhsh = -31
   'MsgBox mLResnew(932, "Check the level of Chemicals or Calibrate them.*", 1)
End If


    

End Sub

Function DISPENSE_XIMIKA(m_synt)

Dim dr As Long, c, quantity, a00, b00, a2
Dim mepil, xr(1 To 5), for_use, dokXr(1 To 5), sXr(1 To 5)
Dim k, L, megas, amax, Off_On(1 To 5), per
Dim max_xronos, mz, mZ1, mz2
Dim sum_gr, q_ner, xr_ner, dum

Dim Max_apoklish
Dim ENA_ENA As Integer

ENA_ENA = 1

Timer1.Enabled = False


Max_apoklish = 3000 'mgr
' m_ximika(epil, 1 To 5)  posothtes

max_xronos = 50000
For k = 1 To 5: xr(k) = 0: Off_On(k) = 0: Next



If IsNull(prospau2.Recordset("epil")) Then Exit Function
If Not IsNumeric(prospau2.Recordset("epil")) Then Exit Function
If prospau2.Recordset("epil") > 50 Then Exit Function

 mepil = Int(prospau2.Recordset("epil"))

'Quantity = xr
for_use = 0
sum_gr = 0
For k = 1 To aritmos_ximikon
  If m_ximika(mepil, k) > 0 Then
      quantity = m_ximika(mepil, k)
      sum_gr = sum_gr + m_ximika(mepil, k)
      
      If m_ximika(mepil, k) < 1600 Then
         m_ximika(mepil, k) = m_ximika(mepil, k) - 70
         for_use = for_use + 1
      End If
      
      If a00x(k) = 0 Then
         xr(k) = 0
      Else
        a00 = a00x(k): b00 = b00x(k)
        dum = DISP_ENA_XIMIKO(k, quantity, a00, b00, m_synt)
        'xr(k) = (m_ximika(mepil, k) - b00x(k)) / a00x(k) 'xrhsimopoionta 2, 0.5 sec
      End If
      
  End If
Next
 
'If Len(XimNer_on) > 5 Then

       



'End If


667  'DUM = DISP_ENA_XIMIKO(5, 5000, a00, b00, m_synt)
 If IsNull(prospau2.Recordset("ximikon")) Then
 Else
     If prospau2.Recordset("ximikon") > 0 Then
     
        a00 = a00x(k): b00 = b00x(k)
        dum = DISP_ENA_XIMIKO(1, prospau2.Recordset("ximikon"), a00, b00, m_synt)
     End If
  End If
 
 
 DISPENSE_XIMIKA = dum
 Exit Function
 
' GoTo 667
 
 If Not IsNull(prospau2.Recordset("nero")) Then  ' Exit Sub
  
  q_ner = prospau2.Recordset("nero")
  sum_gr = sum_gr + q_ner
  If a00x(5) = 0 Then
      xr_ner = 0
  Else
      a00 = a00x(5): b00 = b00x(5)
      dum = DISP_ENA_XIMIKO(5, q_ner, a00, b00, m_synt)
      xr_ner = (q_ner - b00x(5)) / a00x(5) 'xrhsimopoionta 2, 0.5 sec
  End If
  
End If
 

DISPENSE_XIMIKA = dum
    




End Function


Function DISP_ENA_XIMIKO(BALBIDA, quantity, a00, b00, m_synt)
Dim X_ON, X_OFF, c, a2, a0, q1, XR2
Dim cmd$, mydb

Select Case BALBIDA
   Case 5
      X_ON = XimNer_on
      X_OFF = XimNer_Off
   Case 1
      X_ON = Xim1_on
      X_OFF = xim1_Off
      
   
      
      
      
   Case 2
      X_ON = Xim2_on  '18-1,13-0
      X_OFF = xim2_Off
   Case 3
      X_ON = Xim3_on
      X_OFF = xim3_Off
   Case 4
      X_ON = Xim4_on
      X_OFF = xim4_Off
End Select
       
' X_ON = XimNer_on
 'X_OFF = XimNer_Off
       
       
       
a0 = Zygis0(0)
q1 = quantity
'If quantity > 6000 Then q1 = quantity - 4000 Else q1 = quantity
If a00 <> 0 Then
    XR2 = (q1) / 600
Else
   XR2 = 0
End If

XR2 = XR2 * 1000

RobSend ("!" + XimNer_on + ":")
c = GetCurrentTime()
Do While GetCurrentTime() - c < XR2
    DoEvents
Loop
RobSend ("!" + XimNer_Off + ":")

MilSec 3000
a2 = Zygis0(0) - a0
    
'B GYROS --------------------------------------
If quantity > 6000 Then
    q1 = quantity - a2
    If a00 <> 0 Then
       XR2 = (q1 - b00) / a00 'xrhsimopoionta 2, 0.5 sec
    Else
       XR2 = 0
    End If
    
    RobSend ("!" + X_ON + ":")
    c = GetCurrentTime()
    Do While GetCurrentTime() - c < XR2
       DoEvents
    Loop
    cmd$ = "!" + X_OFF + ":"
    RobSend (cmd$)
    MilSec 3000
    a2 = Zygis0(0) - a0
    
End If
If BALBIDA <> 5 Then
   Set mydb = OpenDatabase("c:\talos\recipies.mdb")
   cmd$ = "update prospau2 set regr=" + str(a2) + " where apor=" + str(BALBIDA) + " and  entolh='" + left(m_synt, 6) + "' and aa_prospau='" + Mid$(m_synt, 8, 2) + "';"
   mydb.Execute cmd$
   mydb.Close
End If

Dim dum
dum = NER_DOSOMETRHSH(1000, 99)


DISP_ENA_XIMIKO = a2

'-------------------------------------









End Function

'
'
'
Sub dispenseMONO_Ximika()

Dim dr As Long, c, quantity, a00(1 To 5), b00(1 To 5), a2
Dim mepil, xr(1 To 5), for_use, dokXr(1 To 5), sXr(1 To 5)
Dim k, L, megas, amax, Off_On(1 To 5), per
Dim max_xronos, mz, mZ1, mz2
Dim sum_gr
Dim Max_apoklish

Max_apoklish = 200 'mgr
' m_ximika(epil, 1 To 5)  posothtes

max_xronos = 50000
For k = 1 To 5: Off_On(k) = 0: Next



If IsNull(prospau2.Recordset("epil")) Then Exit Sub
If Not IsNumeric(prospau2.Recordset("epil")) Then Exit Sub
If prospau2.Recordset("epil") > 50 Then Exit Sub

 mepil = Int(prospau2.Recordset("epil"))

'Quantity = xr
for_use = 0
sum_gr = 0
For k = 1 To aritmos_ximikon
  If m_ximika(mepil, k) > 0 Then
      sum_gr = sum_gr + m_ximika(mepil, k)
      If m_ximika(mepil, k) < 1600 Then
         m_ximika(mepil, k) = m_ximika(mepil, k) - 70
         for_use = for_use + 1
      End If
      xr(k) = (m_ximika(mepil, k) - b00x(k)) / a00x(k) 'xrhsimopoionta 2, 0.5 sec
  End If
Next
 
 
 
 
's[-arr
For k = 1 To 5: dokXr(k) = xr(k): Next

For L = 1 To 5
   amax = 0
   For k = 1 To 5
      If dokXr(k) > amax Then
         amax = dokXr(k)
         megas = k
      End If
   Next
   dokXr(megas) = -1
   sXr(L) = amax
Next
 
 
 
 MilSec 1000
 
 
 
 
mz = Zygis0(0)
Do While Abs(mz) > 30
'  G_REZERO MSComm1
  
If Balance_Type = "ADAM" Then
      MSComm1.Output = "ZERO" + Chr(13)
Else
      MSComm1.Output = "Z" + Chr$(13) + Chr(10)
End If
 
 'MSComm1.Output = Zero
  
  
'    MSComm1.Output = "Z " + Chr$(13) + Chr(10)
     MilSec 200
     mz = Zygis0(0)
Loop





c = GetCurrentTime()

If aritmos_ximikon >= 1 Then
    If xr(1) > 0 Then
       RobSend ("!" + Xim1_on + ":")
       Off_On(1) = 1
    End If
End If

If aritmos_ximikon >= 2 Then
   If xr(2) > 0 Then
      RobSend ("!" + Xim2_on + ":")
      Off_On(2) = 1
   End If
End If
  
If aritmos_ximikon >= 3 Then
   If xr(3) > 0 Then
      RobSend ("!" + Xim3_on + ":")
      Off_On(3) = 1
   End If
End If
 
If aritmos_ximikon >= 4 Then
   If xr(4) > 0 Then
      RobSend ("!" + Xim4_on + ":")
      Off_On(4) = 1
   End If
End If
 

 
'=====================================
If sXr(1) > max_xronos Then
   sXr(1) = max_xronos
End If

Do While GetCurrentTime() - c < sXr(1)
     per = GetCurrentTime() - c
     If Off_On(1) = 1 Then
        If per > xr(1) Then
           RobSend ("!" + xim1_Off + ":")
           Off_On(1) = 0
        End If
     End If
     
     If Off_On(2) = 1 Then
        If per > xr(2) Then
           RobSend ("!" + xim2_Off + ":")
           Off_On(2) = 0
        End If
     End If
     If Off_On(3) = 1 Then
        If per > xr(3) Then
           RobSend ("!" + xim3_Off + ":")
           Off_On(3) = 0
        End If
     End If
     If Off_On(4) = 1 Then
        If per > xr(4) Then
           RobSend ("!" + xim4_Off + ":")
           Off_On(4) = 0
        End If
     End If
     

Loop

'---------------------------------------------
' τα  ξανακλεινω για λόγους ασφαλείας
If aritmos_ximikon >= 1 Then RobSend ("!" + xim1_Off + ":")
If aritmos_ximikon >= 2 Then RobSend ("!" + xim2_Off + ":")
If aritmos_ximikon >= 3 Then RobSend ("!" + xim3_Off + ":")
If aritmos_ximikon >= 4 Then RobSend ("!" + xim4_Off + ":")
'If aritmos_ximikon >= 5 Then RobSend ("!" + xim5_Off + ":")
  MilSec 2000
        
dispense_ximika2:
mZ1 = Zygis0(0)
mz2 = Zygis0(0)
If Abs(mZ1 - mz2) > 10 Then GoTo dispense_ximika2
mz = (mZ1 + mz2) / 2

If Abs(mz - sum_gr) <= Max_apoklish Then
   'ok
Else
   m_ximika(50, 5) = -31 ' Ektel_kinhsh = -31
   'MsgBox mLResnew(932, "Check the level of Chemicals or Calibrate them.*", 1)
End If


    
End Sub
'
'
'
'
'
'
Function DOS_ALATI_DOK(M_ALATI)
Dim APOBARO, Apoklish, DIADROMH_BOLAN, Arxh, TIME_ANOXH, Z, ok, TaxyThtes
Dim counter, FIRST_TIME, Motori_On, Xam_Ligo, Xam_Poly, reystothta
Dim FLAG1, D1, D2, BHMA
Dim ZYG1, ZYG2, FLAG2, FLAG3, Target
Dim mydb, r, var93, Skacimo, v_Z
Dim dum, xam_tax, CurQ_Dis
Dim ARX_DIADROMH, m_axe3, DZYG, max_zyg
 'ΑΛΛΑΓΗ   DIADROMH_BOLAN = DIADROMH_BOLAN / 60 => 40   17/04/2001
' XRONIKO PERIUVRIO SE 10MIN
ZYG1 = 0
ZYG2 = 0
FLAG1 = 0
FLAG2 = 0
FLAG3 = 0
BHMA = 10

Label9.Caption = M_ALATI

Z = 0
Motori_On = 0
FIRST_TIME = True
counter = 0   'ΔΕΣ ΤΟ   VAR88     !!!!!!!!

'M_ALATI = 7000
Target = M_ALATI

If Target < 1100 Then
   reystothta = 20
   BHMA = 5
Else
   reystothta = 0
End If
Apoklish = Target * 1 / 100 ' 20

Set mydb = Workspaces(0).OpenDatabase("c:\talos\coord_tb.MDB")
Set r = mydb.OpenRecordset("var")
Do Until r.eof
 If r("field_name") = "VAR93" Then
    var93 = r("field_val")
End If
 If r("field_name") = "VAR88" Then
    var88 = r("field_val")
End If
 
If r("field_name") = "VAR99" Then
    Skacimo = r("field_val")
End If

 
 
r.MoveNext
  Loop
  r.Close


Set mydb = Workspaces(0).OpenDatabase("c:\talos\recipies.MDB")
Set r = mydb.OpenRecordset("parametroi")
  If r.RecordCount = 1 Then
       r.AddNew
       r("FORNEXT_BIG") = 400
       r("FORNEXT_SMALL") = 250
       r.update
  End If
  r.MoveFirst
  r.MoveNext
   D1 = IIf(IsNull(r("FORNEXT_BIG")), 400, r("FORNEXT_BIG"))
   
   D2 = IIf(IsNull(r("FORNEXT_SMALL")), 250, r("FORNEXT_SMALL")) ' r("FORNEXT_SMALL")
  
  
       r.Close






DIADROMH_BOLAN = var93 - var88 ' 10000 BHMATA

If DIADROMH_BOLAN > 16000 Then '=40 MM = megisto ανοιγμα μπουκαλιού σκόνης
    DIADROMH_BOLAN = 16000
End If

TIME_ANOXH = 600000   '10 MIN
'MSComm1.Output = Balance_3_Digits + Chr(13)
'Target = 2000
Set mydb = Workspaces(0).OpenDatabase("c:\talos\coord_tb.MDB")
Set r = mydb.OpenRecordset("taxythtes")
Do Until r.eof
     If r("id_tax") = 7 Then
        v_Z = r("taxythta")
     End If
     r.MoveNext
  Loop
r.Close
  



TaxyThtes = "!V,," + str(v_Z) + ":"
    
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
   
     dum = wait_to_understand("!var12=94:")
     
     
     APOBARO = Zygis0(0)
     
     dum = Wait_Robot(0)


    
Arxh = GetCurrentTime()
'   DUM = wait_to_understand("!var12=94:")
     
  
Xam_Ligo = "!v, , .1:"    '0.1
Xam_Poly = "!v, , .01:"   '0.01


'   DUM = wait_to_understand("!var12=94:")


If Target > 1100 And reystothta = 0 Then
   xam_tax = Xam_Ligo      '  "!v, , .1:"
 '  DIADROMH_BOLAN = DIADROMH_BOLAN / 60    17/04/2001
   
   ' radaeli  big=800   small=500
   If Target > 1500 Then  'radaeli 1500  kanonika itan 4500  18-12-2002
       DIADROMH_BOLAN = D1  ' 400
      ' Apoklish = 20 ' radaeli  2019-09-09
       
   Else
       DIADROMH_BOLAN = D2 ' 250
   End If
Else

   xam_tax = Xam_Poly    '  "!v, , .01:"
   If reystothta <= 10 Then
      DIADROMH_BOLAN = 250
   ElseIf reystothta > 10 And reystothta < 20 Then
       DIADROMH_BOLAN = 25
   Else
       DIADROMH_BOLAN = 800
   End If
   If reystothta >= 10 And Target < 1100 Then
       DIADROMH_BOLAN = 10
       
   End If

End If

ARX_DIADROMH = DIADROMH_BOLAN
  
  RobSend (xam_tax)
   MilSec 500
  RobSend (xam_tax)
' var88 = axe(3)  2/10/2001

    'ΞΕΚΙΝΑ ΤΟ ΜΟΤΟΡΙ  "SKON_MOTOR_ON"
        RobSend ("!" + skon_motor_on + ":")
Motori_On = 1

var88 = axe(3) ' giati exei kai skacimo kai se mperdeyei
ok = 0

Open "C:\TALOS\ALATI.TXT" For Output As #1
Dim ANEBASMATA As Long
Dim KATEBASMATA As Long



' ΤΟ ΤΕΡΜΑ ΤΟΥ ΒΟΛΑΝ  ΕΙΝΑΙ    VAR(88)+12000 bhmata
DIADROMH_BOLAN = IIf(IsNull(DIADROMH_BOLAN), 0, DIADROMH_BOLAN)
'======================================================================================
'==========================================================================================
'=========================================================================================


Do While True
  
  ANEBASMATA = 0

  
  '----ANEBASMA ------------------------ anebasma -----------ανεβασμα1
  RobSend ("!d , ," + str(var88 + DIADROMH_BOLAN) + ":!go,,1:")
  Do While axe(3) < Int(var88 + DIADROMH_BOLAN)
        Z = Zygis0(0) - APOBARO
        Label8.Caption = Z
        If Z > 0 Then frmSYNT.Caption = Format(Z, "#####")
        counter = counter + 1
        ANEBASMATA = ANEBASMATA + 1
        Write #1, frmSYNT.Caption + " ANEBASMA " + str(ANEBASMATA) + " DIADROMH_BOLAN=" + str(DIADROMH_BOLAN)
        If GetCurrentTime() - Arxh > TIME_ANOXH Then Exit Do
        If Z < -999000000 Then   '--- ZYGOS BLOCK
             RobSend ("!" + skon_motor_off + ":!k:")
              Do While Z < -999000000
                    MilSec 500
                   Z = Zygis0(0) - APOBARO
              Loop
              RobSend ("!" + skon_motor_on + ":!go,,1:")
        End If
          
       
        If Z >= M_ALATI - Apoklish * 5 Then  'HTAN 5 2019-09-09  ME 7 PHGE STO 10%
                DOS_ALATI_DOK = Z
                RobSend ("!k:")
                ok = 1
                Exit Do
        End If
        If FLAG1 = 0 And (M_ALATI - Z) < 1000 Then   ' 20191101  APO 500=>1000
            RobSend ("!k:")
            Exit Do
        End If
        
        
        
  Loop
  'TELOS LOOP ----ANEBASMA ------------------------ anebasma -----------ανεβασμα1
  
    If ok = 1 Then
         ' kleisto to mpoykali  ΓΙΑ ΝΑ ΞΑΝΑΖΥΓΙΣΩ
         RobSend ("!d , ," + str(var88 - Skacimo - 200) + ":go,,1:")
         MilSec 500
         RobSend ("!" + skon_motor_off + ":!k:")
          RobSend (TaxyThtes)
           MilSec 500
          RobSend (TaxyThtes)
          ' MilSec 500
                Z = Zygis0(0) - APOBARO
                If Z >= M_ALATI - Apoklish * 2 Then  'HTAN 5 2019-09-09  ME 7 PHGE STO 10%
                   DOS_ALATI_DOK = Z
                   RobSend ("!k:")
                   Exit Do
                Else
                        
                        RobSend (xam_tax)
                        MilSec 500
                        RobSend (xam_tax)
                        RobSend ("!" + skon_motor_on + ":")
                        Motori_On = 1
                End If
    End If
  
  
  ' κατεβασμα
  m_axe3 = 0
  RobSend ("!d , ," + str(var88) + ":!go,,1:")
  
  '---------------------------- katebasma  -----------ανεβασμα1
  KATEBASMATA = 0
  Do While Abs(axe(3) - var88) > 10   ' axe(3)>var88
          Z = Zygis0(0) - APOBARO
           Label8.Caption = Z
          
          KATEBASMATA = KATEBASMATA + 1
          Write #1, Format(Z, "#####0") + " kateBASMA " + str(KATEBASMATA) + " DIADROMH_BOLAN=" + str(DIADROMH_BOLAN)
          
          
          If m_axe3 = axe(3) Then
              RobSend ("!d , ," + str(var88) + ":!go,,1:")
          End If
          counter = counter + 1
          m_axe3 = axe(3)
          
          If Z > 0 Then CurQ_Dis = Format(Z, "#####")
          If Z < -999000000 Then  ' ZYGOS BLOCK
               RobSend ("!" + skon_motor_off + ":!k:")
               Do While Z < -999000000
                MilSec 100
                   Z = Zygis0(0) - APOBARO
              Loop
              RobSend ("!" + skon_motor_on + ":!go,,1:")
        End If


        If GetCurrentTime() - Arxh > TIME_ANOXH Then Exit Do
        If Z >= M_ALATI - Apoklish * 9 Then  'HTAN 5 TO EKANA 9  20191101
                DOS_ALATI_DOK = Z
                ok = 1
                Exit Do
        End If
  Loop
'''''''''''''' END OF KATEBASMA ---------------------------------------

     
     If GetCurrentTime() - Arxh > TIME_ANOXH Then
         Exit Do
     End If
     Ref.Caption = (GetCurrentTime() - Arxh) / 1000
     
     
     
     
     
     
    If ok = 1 Then
         ' kleisto to mpoykali  ΓΙΑ ΝΑ ΞΑΝΑΖΥΓΙΣΩ
         RobSend ("!d , ," + str(var88 - Skacimo - 200) + ":go,,1:")
         MilSec 500
         RobSend ("!" + skon_motor_off + ":!k:")
          RobSend (TaxyThtes)
           MilSec 500
          RobSend (TaxyThtes)
          ' MilSec 500
                Z = Zygis0(0) - APOBARO
                If Z >= M_ALATI - Apoklish * 2 Then  'HTAN 5 2019-09-09  ME 7 PHGE STO 10%
                   DOS_ALATI_DOK = Z
                   RobSend ("!k:")
                   Exit Do
                Else
                        
                        RobSend (xam_tax)
                        MilSec 500
                        RobSend (xam_tax)
                        RobSend ("!" + skon_motor_on + ":")
                        Motori_On = 1
                End If
    End If
     
'     If ok = 1 Then
'         Exit Do
 '    End If
     
     
DZYG = Z - ZYG1
If DZYG > max_zyg And DZYG <= M_ALATI * 0.2 Then
     max_zyg = DZYG
End If


ZYG1 = Z
' HTAN EDV







If DZYG < (M_ALATI - Z) / 2 Then   ' ANEBASE
    
    If DIADROMH_BOLAN < D1 * 3 / 4 Then '300
         If (M_ALATI - Z) < 800 Then
              If D1 >= 800 Then  'xontro alati gia skagia
                  DIADROMH_BOLAN = DIADROMH_BOLAN + 150 '+ 15  ' Agmond Canada
              Else
                  DIADROMH_BOLAN = DIADROMH_BOLAN + 200 '+ 2  'Carolina Cotton
              End If
         Else
              DIADROMH_BOLAN = DIADROMH_BOLAN + 200 ' + 20
         End If
    Else  '2019
          DIADROMH_BOLAN = DIADROMH_BOLAN + 200
    End If

Else
     
     DIADROMH_BOLAN = DIADROMH_BOLAN - 50
     If DIADROMH_BOLAN < 100 Then
         DIADROMH_BOLAN = 50
     End If

End If
     
     
     
     
     
'DIADROMH_BOLAN = 0
If FLAG1 = 0 And (M_ALATI - Z) < 800 Then   '1000
   FLAG1 = 1
   DIADROMH_BOLAN = 50
         Write #1, Format(Z, "#####0") + " <1000 DIADR=50 " + " DIADROMH_BOLAN=" + str(DIADROMH_BOLAN)
   RobSend (xam_tax)
   MilSec 500
   RobSend (xam_tax)
Else


    If FLAG1 = 0 And (M_ALATI - Z) < 2500 Then   '1500  '20191101  APO 1000 TO EKANA 1500
       FLAG1 = 1
       DIADROMH_BOLAN = 200
       RobSend (xam_tax)
                Write #1, Format(Z, "#####0") + " <1500 DIADR=50 " + " DIADROMH_BOLAN=" + str(DIADROMH_BOLAN)
       MilSec 500
       RobSend (xam_tax)
    End If
End If

     
     
     
     
     
     
     
     
     
     
     
' PERIORISMOI
          If DIADROMH_BOLAN > var93 - var88 Then
              DIADROMH_BOLAN = var93 - var88
          End If

     If M_ALATI < 2500 Then
        If DIADROMH_BOLAN > 990 Then
             DIADROMH_BOLAN = 800
         End If
         
         If (M_ALATI - Z) < 300 And DIADROMH_BOLAN > 400 Then
            DIADROMH_BOLAN = 400
         End If
         
         
     End If
     
     
     
     
     If Motori_On = 0 Then RobSend ("!" + skon_motor_on + ":!k:")

            If Target > 1000 Then   '-Z
                xam_tax = Xam_Ligo    '  "!v, , .1:"
            Else
                xam_tax = Xam_Poly    ' "!v, , .01:"
            End If

            RobSend (xam_tax)
            MilSec 500
            RobSend (xam_tax)





        If (M_ALATI - (Zygis0(0) - APOBARO)) < 0.03 * M_ALATI Then
                Write #1, Format(Zygis0(0) - APOBARO, "######") + " exit do " + str(DIADROMH_BOLAN)
            Exit Do
            
         End If


    '  Command1.Caption = DIADROMH_BOLAN
      Write #1, Format(Zygis0(0) - APOBARO, "######") + " xanagyrisma sto megalo loop DIADROMH_BOLAN=" + str(DIADROMH_BOLAN)

  Loop


'================ KENTRIKO LOOP =================================================================

Write #1, Format(Z, "######") + " telos megaloy loop "



'Set mydb = Workspaces(0).OpenDatabase("c:\talos\recipies.MDB")
'Set r = mydb.OpenRecordset("ximitech")
'r.Index = "kod"
'r.Seek "=", left(XIMIKO, 6)
'If Not r.NoMatch Then
'     r.Edit
'     r("min_sygk") = diad_Max_Pos
'     r.update
'End If
'r.Close




RobSend ("!" + skon_motor_on + ":!k:")

MilSec 100
' kleisto to mpoykali
  RobSend ("!d , ," + str(var88 - Skacimo - 200) + ":go,,1:")
MilSec 500
 RobSend ("!" + skon_motor_off + ":!k:")
  
  
  RobSend (TaxyThtes)
    MilSec 500
  RobSend (TaxyThtes)
  



MilSec 100
' kleisto to mpoykali
'  RobSend ("!d , ," + str(var88 - Skacimo - 100) + ":go,,1:")
'MilSec 3000
' RobSend ("!" + skon_motor_off + ":!k:")
  
  
  Write #1, Format(Z, "######") + " proteleytais zygish "
  
  
  DOS_ALATI_DOK = Zygis0(0) - APOBARO
  
  Do While DOS_ALATI_DOK < -APOBARO / 2  ' ΛΑΘΟΣ ΖΥΓΙΣΗ
          MilSec 100
          DOS_ALATI_DOK = Zygis0(0) - APOBARO
          If GetCurrentTime() - Arxh > TIME_ANOXH Then Exit Do
  Loop
  
  
  frmSYNT.Caption = Format(DOS_ALATI_DOK, "#####")
  Label8.Caption = Format(DOS_ALATI_DOK, "#####")
 
 
 Write #1, Format(DOS_ALATI_DOK, "######") + " teliki zygish "
 
 Close #1
 
 
Open "C:\TALOS\ALATITEL.TXT" For Append As #2
Write #2, str(M_ALATI) + "REAL=>" + Format(DOS_ALATI_DOK, "######") + " teliki zygish " + str(GetCurrentTime() - Arxh)
 Close #2
 
 
 
 
 'MSComm1.Output = Balance_2_Digits + Chr(13)
MilSec 100

End Function









Sub ANIXE()

Dim a

a = InStr(Gripper_On, "!")
If a > 0 Then
   Gripper_On = Mid$(Gripper_On, 1, a - 1) + Mid$(Gripper_On, a + 1, Len(Gripper_On) - a)
End If
RobSend (Gripper_On + ":wait(in.1=b0):t0.3:")





'If FindInput(Gripper_input) = gripper_kleistos Then  ' kleistos
'     RobSend (Gripper_On + ":")
'     MilSec 100
'     Do While FindInput(Gripper_input) = 0
'       DoEvents
'       MilSec 50
'        RobSend (Gripper_On + ":")
'    Loop
'    MilSec 500
'End If

End Sub

Function dos_alati_tests(Target)
Dim a(1 To 10, 1 To 10), k, j
Dim SUMA, MEG, Z
'μηδενισμος ζυγου
For k = 1 To 10
    For j = 1 To 10
       'ΚΙΝΗΣΗ Π ΚΑΤΑ 10ΜΜ
       'ΖΥΓΙΣΗ  Ζ=
       a(k, j) = Z - a(k, j)
    Next
Next

For k = 1 To 10
    SUMA = 0: MEG = 0
    For j = 1 To 10
       SUMA = SUMA + a(k, j)
       MEG = IIf(a(k, j) > MEG, a(k, j), MEG)
    Next
   ' mo(k) = SUMA / 10
   ' MEGISTOS(k) = MEG
Next


End Function

Function Load_kinhsh(Index As Integer)
    a_Syntagh(Index) = KINHSEIS.Recordset("SYNTAGH")
    a_Fere(Index) = KINHSEIS.Recordset("FERE")
    a_Kinhsh(Index) = KINHSEIS.Recordset("kinhsh")
    a_Kefalh(Index) = KINHSEIS.Recordset("kefalh")
    a_Mgrs(Index) = KINHSEIS.Recordset("mgrs")
    a_Pane(Index) = KINHSEIS.Recordset("pane")
    a_Xronos(Index) = KINHSEIS.Recordset("xronos")
    a_Tkinhsh(Index) = KINHSEIS.Recordset("tkinhsh")
    a_Kef1(Index) = KINHSEIS.Recordset("KEF1")
    a_Kef2(Index) = KINHSEIS.Recordset("KEF2")
    a_Ar_Syntag(Index) = KINHSEIS.Recordset("ar_syntag")
End Function


Function Mhdenismos_Zygoy(dum)
Dim Can_Zero, mz, counter_num
Screen.MousePointer = 11
Can_Zero = 1
' ΑΝΟΙΓΩ ΣΕΙΡΙΑΚΗ ΠΟΡΤΑ ΚΑΙ ΜΗΔΕΝΙΖΩ ΖΥΓΑΡΙΑ
On Error Resume Next
       MSComm1.commport = Balance_Port
       MSComm1.Settings = Balance_Settings
       MSComm1.InputLen = 0
       MSComm1.PortOpen = False
       MSComm1.PortOpen = True
       
       MSComm1.Output = Balance_3_Digits + Chr$(13)
      MilSec 500
       MSComm1.Output = Balance_Clear + Chr$(13)
         MilSec 500
       MSComm1.Output = Balance_3_Digits + Chr$(13)
         
       mz = Zygis0(0)
       counter_num = 0
         
       Do While Abs(mz) > 2
            counter_num = counter_num + 1
            If counter_num >= 2 Then
                  MsgBox "Παρακαλώ ανοίξτε τη Ζυγαριά !", , "Talos"
                  MilSec 2000
                  MSComm1.Output = Balance_Zero + Chr$(13)
            End If
            MilSec 100
            MSComm1.Output = Balance_Zero + Chr$(13)
            mz = Zygis0(0)
            If counter_num >= 3 And Abs(mz) > 2 Then
                Can_Zero = 0
                Exit Do
            End If
       Loop

On Error GoTo 0

If Can_Zero = 1 Then
   Mhdenismos_Zygoy = 1
Else
    If mz < 900000000 Then
        Mhdenismos_Zygoy = -1
        MsgBox "Η ζυγαριά παραμένει κλειστή"
    Else
        Mhdenismos_Zygoy = 0
    End If
End If
Screen.MousePointer = 1
End Function
         



Sub move_w(thesi)
'        cmd$ = "!var12=90:"
                   
     RobSend ("!d,,," + str(thesi) + ":!go,,,1:")
     Ref.Caption = "!move " + str(thesi)
     Do While Abs(thesi - Uesh_4axona()) > 20
        DoEvents
     Loop



End Sub

Sub sort_recipes()
' ταξινομει τις συνταγές για να έχει λίγες μετακινήσεις
Dim kod(1 To 30, 1 To 10), seira(1 To 30, 1 To 10)
Dim posaX(1 To 30)
Dim k, r, L, mepil, c_s, c_sys, zygo, mono1, mono2

Dim k0, l0, pano, kato, koina, mydb

mydb = OpenDatabase("c:\talos\recipies.mdb")
Set r = mydb.OpenRecordset("select...")
For k = 1 To 30: For L = 1 To 10: kod(k, L) = "**": seira(k, L) = 18: Next: Next

'Set r = data1.Recordset
If r.RecordCount < 3 Then Exit Sub
r.MoveFirst


Do While Not r.eof
     r.EDIT
     r("EPIL") = Int(r("EPIL"))
     r.update
     r.MoveNext
Loop

     
' Exit Sub



' fortoma toy pinaka
r.MoveFirst
mepil = r("epil")
c_s = 1
c_sys = 1
Do While Not r.eof
     kod(c_s, c_sys) = r("kod"): seira(c_s, c_sys) = 18
'     List1.AddItem Str(c_s) + " - " + Str(c_sys) + " - " + kod(c_s, c_sys)
     posaX(mepil) = c_sys
     r.MoveNext
     If r.eof Then Exit Do
     If mepil = r("epil") Then
        c_sys = c_sys + 1
     Else
     
        mepil = r("epil")
        c_sys = 1
        c_s = c_s + 1
     End If
 Loop
 
 
 
 
 
 
 
'piano thn proth zygh syntagh kai blepo ta koina xromata

' to kauena apo thn 2 me poio tairiazei me thn 1
'ZYGO = 4
'MONO1 = 1

For zygo = 2 To c_s Step 2   'TO C_S
    mono1 = zygo - 1
    mono2 = zygo + 1


For k = 1 To posaX(zygo)
     For L = 1 To posaX(mono1)    ' EXOYN IDIO KODIKO KAI DEN EINAI KAPAROMENO
          If kod(zygo, k) = kod(mono1, L) And seira(mono1, L) = 18 Then
                seira(zygo, k) = 5
               ' seira(mono1, l) = 5
                '
                Exit For
          End If
      Next
Next
          
' MONO2 = MONO1 + 2

'BLEPO PINAKA    - DEBUG -
'List1.AddItem "============POSA ME APOPANO =========="
For k0 = 1 To c_s: For l0 = 1 To posaX(k)
      ' List1.AddItem str(k0) + "/" + str(l0) + "/" + kod(k0, l0) + "--" + str(seira(k0, l0))
Next: Next
'List1.AddItem "======================"









For k = 1 To posaX(zygo)
     For L = 1 To posaX(mono2)
          If kod(zygo, k) = kod(mono2, L) Then
                
                seira(zygo, k) = IIf(seira(zygo, k) = 5, 15, 20)
               ' seira(mono2, l) = 10
                Exit For
          End If
      Next
Next
          
'BLEPO PINAKA    - DEBUG -
'List1.AddItem "============POSA ME APOKATO =========="
For k0 = 1 To c_s: For l0 = 1 To posaX(k)
       'List1.AddItem str(k0) + "/" + str(l0) + "/" + kod(k0, l0) + "--" + str(seira(k0, l0))
Next: Next
'List1.AddItem "======================"
          
          
          
          
          
          
'aderfiko moirasma
pano = 0
kato = 0
koina = 0
For k = 1 To posaX(zygo)
      If seira(zygo, k) = 5 Then pano = pano + 1
      If seira(zygo, k) = 20 Then kato = kato + 1
      If seira(zygo, k) = 15 Then koina = koina + 1
Next


Do While pano > 2
    For k = 1 To posaX(zygo)
        If seira(zygo, k) = 5 Then
           seira(zygo, k) = 18
           pano = pano - 1
           Exit For
        End If
    Next
Loop

Do While kato > 2
    For k = 1 To posaX(zygo)
        If seira(zygo, k) = 20 Then
           seira(zygo, k) = 18
           kato = kato - 1
           Exit For
        End If
    Next
Loop










If koina > 0 Then
   If pano = 0 Then
       'pare ena koino
       GoSub parePano: koina = koina - 1: pano = 1
   End If
   If kato = 0 And koina > 0 Then
       ' pare ena koino
       GoSub pareKato: kato = 1: koina = koina - 1
   End If
   If pano = 1 And koina > 0 Then
       'pare ena koino
       GoSub parePano: pano = 2: koina = koina - 1
   End If
   If kato = 1 And koina > 0 Then
       'pare ena koino
       GoSub pareKato: kato = 2: koina = koina - 1
   End If
End If



For k = 1 To posaX(zygo)
   ' List1.AddItem str(seira(zygo, k))
Next



'dino sta mona pano thn telikh ariumhsh
For k = 1 To posaX(zygo)
   If seira(zygo, k) = 5 Then
          For L = 1 To posaX(mono1)
                If kod(zygo, k) = kod(mono1, L) Then
                   seira(mono1, L) = 20
                   Exit For
                End If
          Next
   End If
Next

'dino sta mona Kato thn telikh ariumhsh
For k = 1 To posaX(zygo)
   If seira(zygo, k) = 20 Then
          For L = 1 To posaX(mono2)
                If kod(zygo, k) = kod(mono2, L) Then
                   seira(mono2, L) = 5
                   Exit For
                End If
          Next
   End If
Next

    
'BLEPO PINAKA
'For K0 = 1 To C_S
'    For L0 = 1 To posaX(K)
'       List1.AddItem Str(K0) + "/" + Str(L0) + "/" + kod(K0, L0) + "--" + Str(seira(K0, L0))
'    Next
'Next
'List1.AddItem "======================"
    





Next zygo


'BLEPO PINAKA
For k = 1 To c_s
    For L = 1 To posaX(k)
      ' List1.AddItem kod(k, L) + "--" + str(seira(k, L))
    Next
    Next

'Exit Sub




' ENHMERONO THN DATABASE
' JE-fortoma toy pinaka
r.MoveFirst
For k = 1 To c_s
    For L = 1 To posaX(k)
         r.EDIT
         r("EPIL") = r("EPIL") + seira(k, L) / 1000 '+ IIf(K Mod 2 = 0, 1000000 - Val(R("KOD")), Val(R("KOD"))) / 1000000000000#
         r.update
         r.MoveNext
         If r.eof Then Exit For
     Next
     If r.eof Then Exit For
 Next




' data1.Refresh










Exit Sub
parePano:
For k = 1 To posaX(zygo)
    If seira(zygo, k) = 15 Then
       seira(zygo, k) = 5
       Exit For
    End If
Next

Return



pareKato:

For k = 1 To posaX(zygo)
    If seira(zygo, k) = 15 Then
       seira(zygo, k) = 20
       Exit For
    End If
Next

Return

End Sub

'


Function Zyg_KAUARO(dum)
   Zyg_KAUARO = Zygis0(0) - APOBARO
   Label8.Caption = IIf(Zyg_KAUARO < 0, 0, Zyg_KAUARO)
End Function




Function NER_DOSOMETRHSH(ByVal gr As Single, ByVal porta As String)

'πιάνω το 95% με την μία
Dim ssdxx As Single, Y As Double, Z As Double, w As Double, fores
Dim x As Single, tmax As Double, tmin As Double, t As Double, texcel As Double, stoxos As Double
Dim xronos, mxronos
Dim Poses_Zyg, starter, DEYTEROLEPTA
Dim proth_fora As Byte, ZYG1, DT
Dim posa_loop, mz
Dim arx_xron, cmd$
Dim ll, ii$, zyg0, TEMP%, t_a0
Dim abs_zyg, anoikth, mgr_sec, a0
Dim mS, xr
Dim Water
Dim orio_2 As Long

  DEYTEROLEPTA = 20
  posa_loop = 0
  texcel = 0
 Label9.Caption = gr
 
' MSComm1.Output = "Z " + Chr$(13) + Chr(10)

' G_REZERO MSComm1
     
     MilSec 200
 
mz = Zygis0(0)

Do While Abs(Val(mz)) > 40


  If Balance_Type = "ADAM" Then
    ' ΕΙΝΑΙ ΣΒΗΣΤΗ ΤΗΝ ΑΝΑΒΩ ΚΑΙ ΤΗΝ ΠΑΩ ΣΤΟ 0
    If Zygis0(0) < -99900 Then
        MSComm1.Output = "ON" + Chr$(13)
        MilSec 2000
    End If
  End If
  
   G_REZERO MSComm1
'    MSComm1.Output = "Z " + Chr$(13) + Chr(10)
     
     MilSec 2000
     mz = Zygis0(0)
Loop

APOBARO = mz
Label9.Caption = APOBARO


 'MSComm1.Output = "Z " + Chr$(13) + Chr(10)

'Qpragm = 799


 
EPIU = gr

Label7.Caption = EPIU

arx_xron = GetCurrentTime()

Water_On = Cold_Water_On
Water_Off = Cold_Water_Off





If Val(porta) = 99 Then ' ximiko dipla sto nero
  Water_On = XimNer_on
  Water_Off = XimNer_Off
End If




If EPIU < 5000 Or Val(porta) = 99 Then   '300
  a = 0
Else

   ll = 0
   ii$ = ""
  ZYG = 0
  zyg0 = Zyg_KAUARO(0)
  
  
  Water_On = Cold_Water_On
  Water_Off = Cold_Water_Off
  
If Val(porta) = 99 Then ' ximiko dipla sto nero
  Water_On = XimNer_on
  Water_Off = XimNer_Off
End If
  
  
  
  
  starter = GetCurrentTime()
  cmd$ = "!" + Water_On + ":"
  TEMP% = SendATBlock(768, cmd$, 0)
 
  
 
 abs_zyg = 0
 anoikth = 1
'=================================================
 Do While True
          abs_zyg = Zygis0(1500)
          ' δεν απαντά γρήγορα η ζυγαριά
          Do While abs_zyg < -999000000
                cmd$ = "!" + Water_Off + ":"
                TEMP% = SendATBlock(768, cmd$, 0)
                abs_zyg = Zygis0(3000)
                anoikth = 0
                If GetCurrentTime() - starter > 30000 Then
                         cmd$ = "!" + Water_Off + ":"
                         TEMP% = SendATBlock(768, cmd$, 0)
                         MilSec 3000
                         NER_DOSOMETRHSH = 0
                         Exit Function
                End If
          Loop
          
      If FindInput(OverFlowInput) = 0 Then
                  RobSend ("!" + Water_Off + ":")
                  Stamathma_Talos "OVERFLOW WATER"
                  RobSend ("!" + Alarm1_off + ":")
                  RobSend ("!" + Alarm1_on + ":")
                  MsgBox "WATER ON SCALE"
                  End
       End If
          
         
          ZYG = abs_zyg - APOBARO
          
          If ZYG > 0 Then Label3(0).Caption = ZYG
          
      Poses_Zyg = Poses_Zyg + 1
      If Val(EPIU) - ZYG < Avance_Online_Water Then                 '  3500       2500=kalo
          Exit Do
       End If
       If Poses_Zyg Mod 500 = 0 Then
           If ZYG - zyg0 < 5 Then
                  Exit Do
           Else
                 zyg0 = ZYG
            End If
       End If
       DoEvents
       If anoikth = 0 Then
          anoikth = 1
          cmd$ = "!" + Water_On + ":"
          TEMP% = SendATBlock(768, cmd$, 0)
       End If
       
       If GetCurrentTime() - starter > MaxTimeWater * 1000 / 2 Then
                         cmd$ = "!" + Water_Off + ":"
                         TEMP% = SendATBlock(768, cmd$, 0)
                         MilSec 3000
                         NER_DOSOMETRHSH = 0
                         Exit Function
       End If
 Loop
'=================================================
     
     
     cmd$ = "!" + Water_Off + ":"
     TEMP% = SendATBlock(768, cmd$, 0)
     mxronos = GetCurrentTime() - starter

     Label8.Caption = GetCurrentTime() - starter

      ' τελευταία Ζύγιση
      ZYG = Zygis0(0)
      Label3(0).Caption = ZYG
      a = ZYG
End If


MilSec 2000

     'A = Zyg_TRIPLO(115, 40)
a = Zyg_KAUARO(0)     '  Zyg_TRIPLO(40, 40)  '    gis1(0)

If mxronos = 0 Then
   mgr_sec = 10000
Else
   mgr_sec = 1000 * a / mxronos
End If

If mgr_sec > 30000 Then mgr_sec = 30000
If mgr_sec < 2000 Then mgr_sec = 2000

'mgr_sec = 0 ' IIf(mxronos = 0, 0, A / mxronos)
starter = GetCurrentTime()

proth_fora = 1: ZYG1 = 0: DT = 0

If Val(porta) = 99 Then orio_2 = 40 Else orio_2 = 500
Do While EPIU - a > orio_2
     a0 = a
     t_a0 = GetCurrentTime()
     
     If mgr_sec = 0 Then
        a = 0
     Else
     
        xr = DT + 1000 * (EPIU - a) / mgr_sec - 67
        If xr > 1000 Then xr = 1000
        If (xr + GetCurrentTime()) - starter > MaxTimeWater * 1000 Then
           xr = MaxTimeWater * 1000 - (GetCurrentTime() - starter)
        Else
          If Val(porta) = 99 Then
               a = metrhs3_ner(xr + 100, 1)
          Else
               a = metrhs3_ner(xr, 1) ' /2 mono ston griba
          End If
        End If
     End If
     mS = GetCurrentTime()
     
     If FindInput(OverFlowInput) = 0 Then
                  RobSend ("!" + Water_Off + ":")
                  Stamathma_Talos "OVERFLOW WATER"
                  RobSend ("!" + Alarm1_off + ":")
                  RobSend ("!" + Alarm1_on + ":")
                  MsgBox "WATER ON SCALE"
                  End
     End If
     
     
     
     
   ' τελευταία Ζύγιση
     MilSec 2500
     ZYG = Zyg_KAUARO(0)     '  Zyg_TRIPLO(40, 40)  '    gis1(0)
     If ZYG1 >= ZYG Then
         If proth_fora = 1 Then
            proth_fora = 2
            DT = 90
         End If
         DT = DT + 5
     End If
     
     ZYG1 = ZYG
     a = ZYG
    ' υπολογίζω τα mgr/sec
     If (mS - t_a0) > 0 And a - a0 > 0 Then
          mgr_sec = 1000 * (a - a0) / (mS - t_a0)
          If mgr_sec > 30000 Then mgr_sec = 30000
          If mgr_sec < 5000 Then mgr_sec = 5000
     End If
    
    
    If GetCurrentTime() - starter > MaxTimeWater * 1000 Then
        NER_DOSOMETRHSH = -999000999
        cmd$ = "!" + Water_Off + ":"
        TEMP% = SendATBlock(768, cmd$, 0)
        Exit Do
    End If


Loop

If Abs(EPIU - a) > 10000 Then
     NER_DOSOMETRHSH = a '-999000999
Else
    NER_DOSOMETRHSH = a
End If

End Function







'
'




Function zyg_ADAM(Zygis_Kind As Variant) As Long
On Error GoTo er_det
   Dim counter
  Dim tot_counter
  Dim start
  Dim FromModem As String
  Dim dummy
  Dim BUF
  
   counter = 0: tot_counter = 0: Label10 = 0
' If system_ready = 0 Then zyg_ADAM = "System not Working": Exit Function
  start = GetCurrentTime()
  
  
    Do
    If GetCurrentTime() - start > 10001 Then zyg_ADAM = -999000999: Exit Function
     MSComm1.InBufferCount = 0
      FromModem$ = ""
           MSComm1.Output = Balance_Asking + Chr$(13)
          MilSec (50)
          dummy = DoEvents()
          If MSComm1.InBufferCount Then
           BUF = MSComm1.InBufferCount
          FromModem$ = FromModem$ + MSComm1.Input
        '  List1.AddItem FromModem$, 0
        
        
        
         
              If InStr(UCase(FromModem$), "OVER") > 0 Then
                'MsgBox "Scale Overload ...", , "Talos"
                zyg_ADAM = -7900000  ' "OL"
                Exit Do
              End If
              If InStr(UCase(FromModem$), "UNDER") > 0 Then
                 'MsgBox "Scale Overload ...", , "Talos"
                 zyg_ADAM = -9900000 ' "UL"
                 Exit Do
              End If
        
        
        
           
'           If InStr(FromModem$, "OL") > 0 Then
'                'MsgBox "Scale Overload ...", , "Talos"
'                zyg_ADAM = "OL"
'                Exit Do
'            End If
'           If InStr(FromModem$, "UL") > 0 Then
'                'MsgBox "Scale Overload ...", , "Talos"
'                zyg_ADAM = "UL"
'                Exit Do
'            End If
            If GetCurrentTime() - start > 5000 And Zygis_Kind = "OK" Then Zygis_Kind = 0: tot_counter = 0
          
          
          
          
          
    If InStr(FromModem$, ".") > 0 Then
          
            Dim NN As Integer
            NN = InStr(FromModem$, ".")
            Dim CC3 As String
            Dim DD As String
            
           If NN > 5 Then
            
            
            '    CC = Mid(FromModem$, NN - 3, 7)
             '   DD = Mid(FromModem$, NN, 4)
              '  If InStr(DD, "?") > 0 Then
               '    DD = Replace(DD, "?", "0")
               ' End If
                
                
                CC3 = Mid(FromModem$, NN - 5, 9)
                
                If InStr(CC3, "?") > 0 Then
                 
                   
                Else
                
                  '  CC = Replace(CC, "?", "0")
                    zyg_ADAM = Val(CC3) * 1000
                    Exit Do
                End If
            Else
               zyg_ADAM = -999000999
            End If
            
               
          
     End If

          
          
          
          
          
          
          
          
          
          
          
          
          
          
          
          
          
          
          
          
          
          If InStr(FromModem$, Chr$(13)) >= 5 Then
                 If InStr(FromModem$, ".") >= 5 Then
                       Exit Do
                End If
          End If
       
       
       
       
       counter = counter + 1
       If counter >= 15 Then tot_counter = tot_counter + 1
            Label10 = counter
          If tot_counter >= 20 And Zygis_Kind <> "OK" Then
          counter = 0: tot_counter = tot_counter + 1
               zyg_ADAM = " "
               Exit Function
           End If
          If counter >= 50 Then  ' 500
             counter = 0: tot_counter = tot_counter + 1
               MSComm1.InBufferCount = 0
              MSComm1.Output = Balance_Asking + Chr$(13)
              End If
     End If
     If Zygis_Kind = "OK" Then
             'Me.Zyg_Show.Width = (GetCurrentTime() - start) / 1000 * 650
              'Me.Zyg_Show.Caption = Int(((GetCurrentTime() - start) / 1000) + 0.5) & " Sec"
             'Me.Zyg_Show.Refresh
      End If
      
    Loop
         
       '  zyg_ADAM = Mid$(FromModem$, InStr(FromModem$, ".") - 4, 8)
        ' If InStr(FromModem$, "-") > 0 Or InStr(FromModem$, "?") > 0 Then
         '    zyg_ADAM = Str(-Val(LTrim(zyg_ADAM)))
        ' End If
er_ex:
Exit Function
er_det:
 zyg_ADAM = -999000999  '" "
Resume er_ex
End Function






Function Zygis_ADAM(x As Single)  'OLD ONE
  Dim counter, START1, FromModem$, dummy, ZYG
    ' Flush the input buffer.
       MSComm1.Output = Balance_Asking + Chr$(13)
       MSComm1.InBufferCount = 0
       counter = 0
    ' Wait for "OK" to come back from the modem.
    START1 = GetCurrentTime()
    Do
       MilSec 40
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
            If counter >= 500 Then  ' 500
                counter = 0
                MSComm1.Output = "SEND" + Chr$(13)
                MSComm1.InBufferCount = 0
                FromModem$ = ""
             End If
       End If
        If GetCurrentTime() - START1 > 15000 Then
            Zygis_ADAM = -999000999
            Exit Function
       End If
    Loop
    ZYG = Val(left$(FromModem$, 9))
    Zygis_ADAM = ZYG * 1000
End Function
'
'
'

Function Zygis0(x As Single)
  Dim counter, starter, FromModem$, x0, y0
  
   If Balance_Type = "ADAM" Then
        Zygis0 = zyg_ADAM(0)
        Label3(0).Caption = Zygis0
        Exit Function
   Else
'        Zygis0 = zyg_neo(MSComm1) * 1000
 '       Me.Caption = Zygis0
  '      Exit Function
   End If
  
  
  
    ' Flush the input buffer.
       MSComm1.Output = "SI " + Chr$(13) + Chr$(10)
'       MSComm1.Output = "UPD 20 " + Chr$(13) + Chr$(10)
       MSComm1.InBufferCount = 0
    starter = GetCurrentTime()
    counter = 0
    ' Wait for "OK" to come back from the modem.
    Do
       MilSec 40
       DoEvents
       If GetCurrentTime() - starter > 2000 Then
            Zygis0 = -999000999
            Exit Function
       End If
       ' If there is data in the buffer, then read it.
      ' MilSec 10
       If MSComm1.InBufferCount Then
          FromModem$ = FromModem$ + MSComm1.Input
          ' Check for "OK".
          Me.Caption = FromModem$
          If InStr(FromModem$, Chr$(13)) Then
              If (InStr(FromModem$, ".") > 5 Or InStr(6, FromModem$, ".") > 5) And InStr(FromModem$, "g") > 10 Then
                    Exit Do
              End If
          End If
       Else
            counter = counter + 1
            If counter = 100 Then  ' 500
                 MilSec 500
                 counter = 0
                 MSComm1.Output = "SIR " + Chr$(13) + Chr(10)
                 MSComm1.InBufferCount = 0
                 FromModem$ = ""
             End If
        End If
    Loop
      If InStr(FromModem$, "UL") > 0 Then
       Zygis0 = -9900000
       Exit Function
   End If
   
    If InStr(FromModem$, "OL") > 0 Then
       Zygis0 = -7900000
       Exit Function
   End If
   
   'zygisi0 = FromModem$
     
      If InStr(FromModem$, ".") < 5 Then
         x0 = InStr(6, FromModem$, ".")
         y0 = InStr(x, FromModem$, "g")
      Else
         x0 = InStr(FromModem$, ".")
         y0 = InStr(FromModem$, "g")
      End If
      Zygis0 = Val(Mid$(FromModem$, IIf(x0 - 6 < 1, 1, x0 - 6), 10)) * 1000 'If Val(Label2) < 1 Then
    
      ' Me.Caption = Format(Zygis0, "###0.000")
    
    '     write_zygish "*" + FromModem$, Zygis0
'         zyg = Val(left$(FromModem$, 9))
         'Label3(0).Caption = Zygis0
 '        Zygis0 = zyg * 1000
End Function


Function Zygis0OLD(x As Single)
  Dim counter, starter, dummy, FromModem$
  Dim Y
  Ref.Caption = left(Ref.Caption + Space(15), 15) + "Sc"

    ' Flush the input buffer.
      
    '   MSComm1.Output = "UPD 20 " + Chr$(13) + Chr$(10)
    
    If Balance_Type = "ADAM" Then
       Zygis0OLD = Zygis0_ADAM(MSComm1, 0)
       Label3(0).Caption = Zygis0OLD
       'ZYGIS0OLD = Zygis0_ADAM(x)
       Exit Function
    Else
    
       Zygis0OLD = zyg_neo(MSComm1)
       Label3(0).Caption = Zygis0OLD
    
    
    End If
    
     MSComm1.Output = "SI " + Chr$(13) + Chr$(10)
    
    
       MSComm1.InBufferCount = 0
    starter = GetCurrentTime()
    counter = 0
    ' Wait for "OK" to come back from the modem.
    Do
       dummy = DoEvents()
          If GetCurrentTime() - starter > 2000 Then
            Zygis0OLD = -999000999
            Gwrite_zygish "---OverTime---", z0
            Exit Function
      End If
       ' If there is data in the buffer, then read it.
       MilSec 40
       If MSComm1.InBufferCount Then
          FromModem$ = FromModem$ + MSComm1.Input
          ' Check for "OK".
          Me.Caption = FromModem$
          If InStr(FromModem$, Chr$(13)) Then
                FromModem$ = LTrim(Right$(FromModem$, 20))
                If (InStr(FromModem$, ".") Or InStr(6, FromModem$, ".")) And InStr(6, FromModem$, "g") > 10 Then
                       Exit Do
                End If
          End If
          
          
          'If InStr(FromModem$, Chr$(13)) Then
           '   If (InStr(FromModem$, ".") > 5 Or InStr(6, FromModem$, ".") > 5) And InStr(FromModem$, "g") > 10 Then
            '        Exit Do
             ' End If
          'End If
       Else
            counter = counter + 1
            If counter = 500 Then  ' 500
                 counter = 0
                 MSComm1.Output = "SIR" + Chr$(13) + Chr(10)
                 MSComm1.InBufferCount = 0
                 FromModem$ = ""
             End If
        End If
    Loop
      
      
    
    
 
      If InStr(FromModem$, ".") < 5 Then
         x = InStr(6, FromModem$, ".")
         Y = InStr(x, FromModem$, "g")
      Else
         x = InStr(FromModem$, ".")
         Y = InStr(FromModem$, "g")
      End If
      Zygis0OLD = Val(Mid$(FromModem$, x - 6, 10)) * 1000

    
'    Gwrite_zygish FromModem$, ZYGIS0OLD
    Label3(0).Caption = Zygis0OLD
End Function



'
'
Function fernei(Index As Integer)
 '   Dim A
 '   If Label1(Index).BackColor = &HFFFFFF Then  ' ΑΝ ΕΙΝΑΙ ΑΣΠΡΟ ΠΑΝΤΟ ΠΙΣΩ
 '       A = 4000 + Index
 '      Label1(Index).BackColor = &HFF&
 '   Else
 '       A = 3000 + Index
 '       Label1(Index).BackColor = &HFFFFFF
 '   End If
 '   cmd$ = "!var12=" + Str(A) + ":"
 '    temp% = SendATBlock(768, cmd$, 0)
 '    Label2.Caption = cmd$

End Function


Function Update2_PROSPAU(mpoykali, UESH, Zygisma)
    Dim m_ek, M_gl, m_gr, m_regr
    Dim KOD_XHM
        ' ΕΝΗΜΕΡΏΝΩ ΤΙΣ ΠΡΑΓΜΑΤΙΚΈΣ ΠΟΣΌΤΗΤΕΣ ΠΟΥ ΔΟΣΟΜΈΤΡΗΣΑ
        prospau2.Recordset.EDIT
        prospau2.Recordset("REGR") = Zygisma        ' αρχείο prospau
        ' υπολογίζω τα πραγματικά ποσοστά
        m_ek = prospau2.Recordset("EK")
        M_gl = prospau2.Recordset("gl")
        m_gr = prospau2.Recordset("gr")
        m_regr = prospau2.Recordset("Regr")
        If m_gr > 0 Then
           prospau2.Recordset("REk") = (m_ek + M_gl) * IIf(m_gr = 0, 1, (m_regr / m_gr))
        End If
       
        prospau2.Recordset("status") = 2
        prospau2.Recordset("usere") = User_ID
        KOD_XHM = prospau2.Recordset("KOD")
        prospau2.Recordset.update
        Ypoloipo(mpoykali) = Ypoloipo(mpoykali) - Zygisma / 1000   ' πίνακας
        
        ximitech.Recordset.Seek "=", KOD_XHM
         
        On Error Resume Next
        ximitech.Recordset.EDIT
        ximitech.Recordset("ypol_prot") = ximitech.Recordset("ypol_prot") - Zygisma / 1000
           ximitech.Recordset("last_dosom") = CStr(Now)
                 'XIMITECH.Recordset("katofli") = katofli(MPOYKALI)
        ximitech.Recordset.update
        On Error GoTo 0
End Function
'
'
Function NER_metrhs2(xr As Double, ByVal DEYTEROLEPTA As Integer) As Double
 '*************************************

End Function








Function gyrise_zygaria(chfia)
Dim dummy

' ΖΥΓΙΖΕΙ ΣΤΟ    "HIGH"  2 δεκαδικά ακρίβεια
 Dim counter, starter, Proeid, FromModem$, a
    Proeid = 0
    ' Flush the input buffer.
       MSComm1.Output = "SEND" + Chr$(13)
       MSComm1.InBufferCount = 0
    starter = GetCurrentTime()
    counter = 0
    ' Wait for "OK" to come back from the modem.
    Do
       dummy = DoEvents()
       If GetCurrentTime() - starter > 27000 Then
          '
             MsgBox "Ανοίξτε την ζυγαριά KAI ΓΥΡΙΣΤΕ ΤΗΝ ΣΤΑ " + str(chfia) + " ΠΑΤΩΝΤΑΣ ΤΟ   F "
             starter = GetCurrentTime()
        End If
        Exit Do
       ' If there is data in the buffer, then read it.
       If MSComm1.InBufferCount Then
          FromModem$ = FromModem$ + MSComm1.Input
          ' Check for "OK".
          If InStr(FromModem$, Chr$(13)) Then
                 
                 If InStr(FromModem$, ".") Then
                      a = InStr(FromModem$, ".")
                      
                      ' 2ο  ψηφίο ΜΕΤΑ ΤΗΝ  .  είναι ΑΡΙΘΜΗΤΙΚΟ
                      If IsNumeric(Mid$(FromModem$, a + 2, 1)) Then
                         'εινα γυρισμ ένο στα 3 ψηφία
                         If chfia = 2 Then
                              gyrise_zygaria = 1
                              Exit Function
                         End If
                         If chfia = 3 Then
                              MilSec 100
                              MSComm1.Output = Balance_3_Digits + Chr$(13)
                              FromModem$ = "0000000000000"
                         End If
                      End If
                      
                        
                      ' τριτο ψηφίο ΜΕΤΑ ΤΗΝ  .  είναι ΑΡΙΘΜΗΤΙΚΟ
                      If IsNumeric(Mid$(FromModem$, a + 3, 1)) Then
                         'εινα γυρισμ ένο στα 3 ψηφία
                         If chfia = 2 Then
                              MilSec 100
                              MSComm1.Output = Balance_2_Digits + Chr$(13)
                              FromModem$ = "0000000000000"
                         End If
                         If chfia = 3 Then
                              gyrise_zygaria = 1
                              Exit Function
                         End If
                      End If
                End If
                
                If Mid$(FromModem$, 6, 1) = "." Or Mid$(FromModem$, 7, 1) = "." Then
                    Exit Do
                End If
          End If
       Else
            counter = counter + 1
            If counter = 500 Then  ' 500
                counter = 0
                MSComm1.Output = "SEND" + Chr$(13)
                MSComm1.InBufferCount = 0
                FromModem$ = ""
             End If
        End If
    Loop
    

End Function
Function metrhs2(xr As Double, rixies As Integer) As Double
 '*************************************
Dim ll, ii$, cmd$, TEMP%
'***  online  ΖΥΓΙΣΗ
Dim Poses_Zyg, start, zyg0, blab
 ll = 0
ii$ = ""
ZYG = 0
blab = 0
 Poses_Zyg = 0
 cmd$ = "!out." + porta + "-1:" + Open_Valve
         TEMP% = SendATBlock(768, cmd$, 0)
 zyg0 = Zyg_KAUARO(0)
 Do While True
     ZYG = Zyg_KAUARO(0)
     Poses_Zyg = Poses_Zyg + 1
     If Val(EPIU) - ZYG < 1200 Then
        Exit Do
     End If
     If Poses_Zyg Mod 500 = 0 Then
           If ZYG - zyg0 < 5 Then
               blab = 1
               Exit Do
           Else
               zyg0 = ZYG
            End If
     End If
 Loop
  cmd$ = Close_Valve + "!out." + porta + "-0:"
 
         TEMP% = SendATBlock(768, cmd$, 0)


' τελευταία Ζύγιση
ZYG = Zyg_TRIPLO(100, 2)
  '  Label3(0).Caption = zyg
 If blab = 1 Then
     metrhs2 = DEN_BGAZEI
 Else
     metrhs2 = ZYG
 End If

End Function


'
'
Private Sub print_rep(mpot, m_synt)

' print_rep str(a_Syntagh(Move)), a_Ar_Syntag(Move)
Dim k, JOBLIST, x, r, ok, a, mydb, ddd, Diff, stoxos
Dim rkod As String
Dim m_err

'dim ddd
If Check1.value <> 1 Then Exit Sub



On Error Resume Next

   'Kill "C:\TALOS\dokprosp.dbf"


If prospau2.Recordset.eof Then prospau2.Recordset.MoveLast

Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
mydb.Execute "DELETE *FROM DOKPROS2;"


mydb.Execute "INSERT INTO DOKPROS2 SELECT prospau2.*  FROM PROSPAU2 where entolh='" + left(m_synt, 6) + "' and aa_prospau='" + Mid$(m_synt, 8, 2) + "';"

mydb.Execute "update prospau2 set status=2 where  entolh='" + left(m_synt, 6) + "' and aa_prospau='" + Mid$(m_synt, 8, 2) + "';"

mydb.Execute "UPDATE dokpros2 SET epil=" + str(mpot) + ";"


mydb.Execute "DELETE *FROM dokprosp;"
mydb.Execute "INSERT INTO DOKPROSP select dokpros2.*  FROM DOKPROS2 order by seira ;"


'A = InStr(Time$, ":")
'x = Mid$(Time$, 1, A - 1) + Mid$(Time$, A + 1, 2)

'mydb.Execute "UPDATE dokprosP SET APOR=" + x + "/100;"
ok = 1

Set r = mydb.OpenRecordset("dokprosp")
r.MoveFirst

' r("apor") = Val(Format(r("date_disp"), "hh.mm"))

ddd = Val(left((Format(r("date_disp"), "hh.mm")), 2)) + Val(Right((Format(r("date_disp"), "hh.mm")), 2)) / 100

rkod = r("kod")
m_err = "-----------------------"
Do While Not r.eof
   Diff = Abs(r("GR") - r("regr"))
   stoxos = r("gr")
     
     If r("SEIRA") = 1 Then
        If Abs(r("NERO") - r("RENERO")) > f_poso_apor(5) Then
                ok = 0
                m_err = "------- Water --------------"
                Ref.Caption = "ref:Water not proper"
        End If
     End If
     
     
     If stoxos <= 2000 Then
        If Diff > f_poso_apor(1) Then
           ok = 0 ':   If ok = 0 Then r("perigr") = "-------------- ERROR"
           m_err = "-------- " + r("perigr")
        End If

     ElseIf stoxos <= 8000 And stoxos > 2000 Then
        If Diff > f_poso_apor(2) Then
            m_err = "---------- " + r("perigr")
           ok = 0  ':   If ok = 0 Then r("perigr") = "-------------- ERROR"
        End If
     ElseIf stoxos <= 16000 And stoxos > 8000 Then
        If Diff > f_poso_apor(3) Then
           ok = 0 ' :   If ok = 0 Then r("perigr") = "-------------- ERROR"
            m_err = "---------- " + r("perigr")
            

        End If
     Else
        If Diff > f_poso_apor(4) Then
           ok = 0  ' :   If ok = 0 Then r("perigr") = "-------------- ERROR"
           m_err = "---------- " + r("perigr")

        End If
     End If
   
   If r("regr") < 0 Then
      ok = 0
      m_err = "---------- " + r("perigr")
   End If
   
   r.EDIT
   r("apor") = ddd
   r.update
   r.MoveNext
Loop

If ok = 0 Then
    r.AddNew
    r("kod") = rkod
    r("perigr") = left(m_err, 28)
    r.update
     '       mydb.Execute "UPDATE dokprosP SET perigr='" + "------ERROR-------';"
End If

r.Close

If ok = 0 Then
    mydb.Execute "UPDATE dokprosp SET epil=" + str(mpot) + ";"
End If

mydb.Close

 Report1.Action = 1

On Error GoTo 0

End Sub
Private Sub back()
Dim kod6400, k, cmd$, TEMP%, L
   
For k = 2 To 19
  If Label3(k).BackColor = &HFF& Then ' αν δεν είναι κόκκινη θέση φέρτο εδώ
   kod6400 = 12000 + Ar_Mpoykal * k + Val(Label3(k).Caption)
   cmd$ = "!var12=" + str(kod6400) + ":"
     TEMP% = SendATBlock(768, cmd$, 0)
     For L = 1 To 1000
     Next
     ' περίμενε μέχρι να κλείσει ο gripper
     While Not (out_bit(4) = 1 And out_bit(5) = 0) 'περίμενε μέχρι να πάρεις σήμα ότι τελείωσε το ρομποτ την εργασία του
        DoEvents
     Wend
     Label3(k).BackColor = &HFFFF80       ' χρωμάτισε το μπουάλι άσπρο
     
     
     While out_bit(bit_ok) = 0
        DoEvents
     Wend
     Label1(Val(Label3(k).Caption)).BackColor = &HFF&   ' κοκκινο
     For L = 1 To 100000
     Next
     
End If
Next


End Sub


Function metrhs3_ner(ByVal XR2 As Double, rixies As Integer) As Single
'**********************************************
Dim dr As Long, c, cmd$, TEMP%, xr
xr = XR2
xr = IIf(xr <= 0, 0, xr)

xr = IIf(xr > MaxTimeWater * 1000 / 2, MaxTimeWater * 1000 / 2, xr)

c = GetCurrentTime()
 
 cmd$ = "!" + Water_On + ":"
 TEMP% = SendATBlock(768, cmd$, 0)
  
Do While GetCurrentTime() - c < xr
Loop
  
  
  cmd$ = "!" + Water_Off + ":"
  TEMP% = SendATBlock(768, cmd$, 0)
         
    metrhs3_ner = ZYG
    
End Function




Function Uesh_4axona()
Dim result%, status_high%, status_low%, status&
   ' result% = Request_Status(768)
    ' ******* Update Motor Position ****************
   ' Set_Pointer 768, AXIS4_MOTOR
   ' Read_Status 768, status_high%, status_low%, status&
    Uesh_4axona = axe(4)
End Function
Function uesh_ajonon()
     Dim x, result%, status_high%, status_low%, status&
     Dim Y, Z
       uesh_ajonon = axe(1) + axe(2) + axe(3)
       
End Function





Private Sub Check1_Click()
'  If Check1.Value = 0 Then
'      DBGrid1.Visible = False
'Else
'      DBGrid1.Visible = True
'End If
End Sub

Private Sub cmdNERO_Click()
447

 Dim SSS As Long
 SSS = GetCurrentTime()
 G_REZERO MSComm1
 MsgBox GetCurrentTime() - SSS
 
GoTo 447
'  NER_DOSOMETRHSH 8000, MSComm1

If MSComm1.PortOpen = False Then

Else
  MSComm1.PortOpen = False
End If

' G_REZERO MSComm1
  
   Syn_Dos.asked_q = 400   ' a_Mgrs(Move) '  ZHTOYMENH ΠΟΣΟTHTA ΣE MGRS
               Fornext_small = 2000
                Syn_Dos.Level = 500  ' Ypoloipo(a_Fere(Move))      'ΓΡΑΜΑΡΙΑ ΜΠΟΥΚΑΛΙΟΥ
                Syn_Dos.Valve_On = valve_1_on    '  , valve_2_on) ' ΑΝΟΙΓΜΑ ΒΑΛΒΙΔΑΣ
                Syn_Dos.Valve_Off = valve_1_off    '  , valve_2_off)   ' ΚΛΕΙΣΙΜΟ ΒΑΛΒΙΔΑΣ
                Syn_Dos.Caption = "TEST"   '    str(a_Fere(Move)) + " : " + Perigrafh(a_Fere(Move)) 'περιγραφη χρώματος που δοσομετρείται
                Syn_Dos.Show 1
  
  
  
  
  
  
  
  
  
End Sub

Private Sub Command1_Click()
Timer1.Enabled = False

'  PROSPAU.Recordset.MoveFirst
'  Do While Not PROSPAU.Recordset.EOF
'        If PROSPAU.Recordset("regr9") > 0 Then
'           PROSPAU.Recordset.Edit
'           PROSPAU.Recordset("STATUS") = "2"
'           PROSPAU.Recordset.Update
'        End If
'        PROSPAU.Recordset.MoveNext
'  Loop
    
    
ZAP_MOVES '12/99
    
     Unload Me
    If Menu = 3 Then
    
      MDIForm1.Arxeia(10).Visible = True
      MDIForm1.Syntages(20).Visible = True
      MDIForm1.bohuhtika(30).Visible = True
      MDIForm1.ejodos(40).Visible = True
    
    
    
    Else
      WAIT.Show
    End If
End Sub


Private Function MAKE_FILE()
Dim m(1 To 30, 1 To 199), MilliGrs(1 To 30, 1 To 199), UeshMpoyk(1 To 30, 1 To 199)
Dim Ar_Syntag, k, MSYN, Arxh, m_nero(1 To 30)
Dim Mpoyk_Synt, Count, mS, mKEF, mydb, MySet
Dim sql_del, mgr, L, aa, bb, mXR, kef_doso, kef_kin
   tkin(1) = "To scale pot....."
   tkin(2) = "Return pot......."
   tkin(3) = "To head bottle..."
   tkin(4) = "Return&bring....."
   tkin(5) = "Right Head Scale."
   tkin(6) = "Right Head scale."
   tkin(7) = "Left head scale.."
   tkin(8) = "Water Dosing....."
   tkin(9) = "Dosing..........."
   tkin(10) = "....."
On Error GoTo problem
Set mydb = Workspaces(0).OpenDatabase("c:\talos\coord_tb.mdb")
Set MySet = mydb.OpenRecordset("KINHSEIS", dbOpenDynaset)
 
' sql_del = "insert into KINHSALL in 'c:\talos\coord_tb.mdb' select [kinhseis].*  from [kinhseis];"
'  mydb.Execute sql_del

 
sql_del = "DELETE DISTINCTROW KINHSEIS.* FROM KINHSEIS"
mydb.Execute sql_del
mydb.Close
   
   Ar_Syntag = IIf(Val(Text1.text) > Ar_Pothr Or Val(Text1.text) < 1, 0, Val(Text1.text) - 1)
   ' διαβάζω τον πίνακα από τις συνταγές
   
   ximitech.Recordset.Index = "KOD"
   prospau2.Recordset.MoveFirst
   
   Do While Not prospau2.Recordset.eof
       Ar_Syntag = Ar_Syntag + 1
       m_syntag(Ar_Syntag) = Trim(prospau2.Recordset("ENTOLH")) + "-" + LTrim(prospau2.Recordset("aa_PROSPAU"))
       Count = 0
       Do While Not prospau2.Recordset.eof And m_syntag(Ar_Syntag) = prospau2.Recordset("ENTOLH") + "-" + LTrim(prospau2.Recordset("aa_prospau"))
          If prospau2.Recordset("SEIRA") = 1 Then
             m_nero(Ar_Syntag) = prospau2.Recordset("nero")
          End If
          If Not IsNull(prospau2.Recordset("GR")) Then
              mgr = prospau2.Recordset("GR")
              If mgr > 0 Then
                   'XIMITECH.Recordset.Index = "KOD"
                   ximitech.Recordset.Seek "=", prospau2.Recordset("KOD")
                   If Not IsNull(ximitech.Recordset("ADDr_prot")) Then
                     Count = Count + 1
                     m(Ar_Syntag, Count) = ximitech.Recordset("ADDr_prot")
                     MilliGrs(Ar_Syntag, Count) = Val(prospau2.Recordset("gr"))
                     UeshMpoyk(Ar_Syntag, Count) = k 'θέση που θα ενημερώσω μετά την ζύγιση
                     End If
              End If
           End If
           prospau2.Recordset.MoveNext
           If prospau2.Recordset.eof Then Exit Do
       Loop
       If prospau2.Recordset.eof Then Exit Do
      ' prospau2.Recordset.MoveNext
   Loop
   
   
   
   
   
   
KEF(1) = 0: KEF(2) = 0
xronos = 1
 KEFALH = 1
Arxh = IIf(Val(Text1.text) > Ar_Pothr Or Val(Text1.text) < 1, 1, Val(Text1.text))
For mS = Arxh To Ar_Syntag
     'βρίσκω πόσα μπουκάλια έχει η συνταγή
     Mpoyk_Synt = 0
     For L = 1 To 30 ' 25-10-2000
       If Not IsNull(m(mS, L)) Then
          If m(mS, L) > 0 Then
            Mpoyk_Synt = Mpoyk_Synt + 1
          End If
       End If
     Next
    
       'φέρε ποτήρι
       Upd_kin mS, 1, mS, 0, 0, 0, 1
       
       
        xronos = xronos + 1
       ' δοσομέτρηση νερού
       Upd_kin mS, 8, 0, 0, m_nero(mS), 0, xronos
     
       'φέρε 1ο μπουκάλι στην ζυγαριά
       Select Case m(mS, 1)
            Case KEF(1) 'βρίσκεται στην κεφαλή 1
               xronos = xronos + 1
               KEFALH = 1
               'κεφαλή 1 σε ζυγαριά
               Upd_kin mS, 6, m(mS, 1), KEFALH, 0, 0, xronos
            Case KEF(2) 'βρίσκεται στην κεφαλή 2
               xronos = xronos + 1
               KEFALH = 2
               'κεφαλή 2 σε ζυγαριά
               Upd_kin mS, 7, m(mS, 1), KEFALH, 0, 0, xronos
            Case Else
               '--------------
               KEF(1) = m(mS, 1) 'αποθηκεύω στην κεφαλή τον αρ.μπουκαλιού
               
               Upd_kin mS, 3, m(mS, 1), KEFALH, 0, 0, xronos
               'KEF(1) = m(mS, 1) 'αποθηκεύω στην κεφαλή τον αρ.μπουκαλιού
       End Select
             
             
             ' kef_dos = KEFALH
             xronos = xronos + 1
       
       'φέρε 2o μπουκάλι δοσομέτρησε το 1ο
       If Mpoyk_Synt > 1 Then
             'kef_dos = KEFALH
             KEFALH = IIf(KEFALH = 1, 2, 1)
             If KEF(KEFALH) = 0 Then 'ηταν άδεια ( η προηγούμενη είχε μόνο ένα μπουκαλι)
                'φέρε 2 ο  μπουκάλι
                KEF(KEFALH) = m(mS, 2)
                Upd_kin mS, 3, m(mS, 2), KEFALH, 0, 0, xronos
             Else
                 'έχει η κεφαλή ΗΔΗ το 2ο μπουκάλι
                 If KEF(KEFALH) = m(mS, 2) Then 'ήδη υπάρχει στην άλλη κεφαλή το 2ο μπουκάλι
                    'μην κάνεις τίποτα
                 Else
                    'πανε πίσω το τελ.μπουκάλι της προηγούμενης συνταγής
                    '& φέρε το 2ο
                    mKEF = KEF(KEFALH)  'αυτό που είχα από την προηγούμενη συνταγή
                    KEF(KEFALH) = m(mS, 2)
                    Upd_kin mS, 4, m(mS, 2), KEFALH, 0, mKEF, xronos
       
                 End If
             End If
             '  στην δοσομέτρηση σαν PANE βάζω την θέση στην αρχική συνταγή για να μπορώ να ενημερώσω το REGRi
             Upd_kin mS, 9, m(mS, 1), 0, MilliGrs(mS, 1), UeshMpoyk(mS, 1), xronos 'δοσομετρώ το 1ο μπουκάλι
       End If
             
    '---------------------------stop για σήμερα 30-8-98  --------------

For L = 3 To Mpoyk_Synt
             xronos = xronos + 1
             KEFALH = IIf(KEFALH = 1, 2, 1)
      '   O2    O1d
          'πανε πίσω το  L-2 μπουκάλι & φέρε το L
             Upd_kin mS, 4, m(mS, L), KEFALH, 0, KEF(KEFALH), xronos
              KEF(KEFALH) = m(mS, L)
              '  στην δοσομέτρηση σαν PANE βάζω την θέση στην αρχική συνταγή για να μπορώ να ενημερώσω το REGRi
              'δοσομετρώ το L-1
              Upd_kin mS, 9, m(mS, L - 1), 0, MilliGrs(mS, L - 1), UeshMpoyk(mS, L - 1), xronos 'δοσομετρώ το 1ο μπουκάλι
Next
     
     
     
     If mS = Ar_Syntag Then 'τελευταία συνταγή ................
             xronos = xronos + 1
             
             KEFALH = IIf(KEFALH = 1, 2, 1)
             
             'επιστρέφω το μπουκάλι τελ-1
             If KEF(KEFALH) > 0 Then
                mKEF = KEF(KEFALH)
                KEF(KEFALH) = 0
                Upd_kin mS, 5, mKEF, KEFALH, 0, 0, xronos
             End If
             
             KEFALH = IIf(KEFALH = 1, 2, 1)
             '  στην δοσομέτρηση σαν PANE βάζω την θέση στην αρχική συνταγή για να μπορώ να ενημερώσω το REGRi
             Upd_kin mS, 9, m(mS, Mpoyk_Synt), 0, MilliGrs(mS, Mpoyk_Synt), UeshMpoyk(mS, Mpoyk_Synt), xronos 'δοσ τελευταίο
             
             xronos = xronos + 1
             'επιστρέφω το τελευταίο μπουκάλι
             KEF(KEFALH) = 0
             Upd_kin mS, 5, m(mS, Mpoyk_Synt), KEFALH, 0, 0, xronos 'Πάνε
         
             
             
     Else 'αυτη η συνταγή δεν είναι η τελευταία
           xronos = xronos + 1
           
       'ελέγχω μήπως το τελευταίο είναι ίδιο με το 1ο της επόμενης
       If m(mS + 1, 1) <> m(mS, Mpoyk_Synt) Then
           If Mpoyk_Synt > 1 Then
                   ' επιστρέφω το τελ-1 φέρνω το 1ο της 2ης συνταγής
                   KEFALH = IIf(KEFALH = 1, 2, 1)
                   KEF(KEFALH) = m(mS + 1, 1)
                   Upd_kin mS, 4, m(mS + 1, 1), KEFALH, 0, m(mS, Mpoyk_Synt - 1), xronos 'Πάνε-φερε
           Else
              KEFALH = IIf(KEFALH = 1, 2, 1)
              '  φέρνω το 1ο της 2ης συνταγής
              If KEF(KEFALH) = 0 Then
                   KEF(KEFALH) = m(mS + 1, 1)
                   Upd_kin mS, 3, m(mS + 1, 1), KEFALH, 0, 0, xronos 'φερε
              Else
                   mKEF = KEF(KEFALH)
                   KEF(KEFALH) = m(mS + 1, 1)
                   Upd_kin mS, 4, m(mS + 1, 1), KEFALH, 0, mKEF, xronos 'φερε
              End If
           End If
       End If
       
           'δοσομετρώ το τελευταίο
           '  στην δοσομέτρηση σαν PANE βάζω την θέση στην αρχική συνταγή για να μπορώ να ενημερώσω το REGRi
           Upd_kin mS, 9, m(mS, Mpoyk_Synt), 0, MilliGrs(mS, Mpoyk_Synt), UeshMpoyk(mS, Mpoyk_Synt), xronos 'δοσ ν-1
     End If
             
             'πάνε  ποτήρι
             xronos = xronos + 1
             Upd_kin mS, 2, mS, 0, 0, 0, xronos


Next   'επομενη συνταγή
On Error GoTo 0
KINHSEIS.Refresh
'DBGrid1.Visible = True

KINHSEIS.Recordset.MoveFirst
Load_kinhsh (1)


MAKE_FILE = 0
Exit Function

problem:
   MAKE_FILE = mS
   MsgBox "874-875" + mL_CapRes(873) + m_syntag(mS) + Chr(13) + mL_CapRes(874)
   



Exit Function


Printer.FontSize = 8
Printer.FontName = "COURIER"

aa = "Συνταγή α/α Ενέργεια       Φέρε        Πάνε          χρόνος          Κεφ1         Κεφ2         Κεφαλή         Κεφ.δοσ"
bb = "---------------------------------------------------------------------------------------------------------------------"
Printer.Print aa
Printer.Print bb

Do While Not KINHSEIS.Recordset.eof
    MSYN = a_Syntagh(1)
    mXR = a_Xronos(1)
    Load_kinhsh (1)
    If MSYN <> a_Syntagh(1) Then
       Printer.Print bb
       Printer.Print aa
       Printer.Print bb

    End If
    kef_doso = IIf(a_Kinhsh(1) = 9, Format(a_Kefalh(1), "###"), "   ")
    kef_kin = IIf(a_Kinhsh(1) <> 9, Format(a_Kefalh(1), "###"), " ")
    Printer.Print a_Ar_Syntag(1); Format(a_Syntagh(1), "##"); left(a_Tkinhsh(1) + "            ", 15), Format(a_Fere(1), "###"), Format(a_Pane(1), "###"), Format(a_Xronos(1), "###"), Format(a_Kef1(1), "###"), Format(a_Kef2(1), "###"), kef_kin, kef_doso
    If mXR = a_Xronos(1) Then
      Printer.Print "    "
    End If
    KINHSEIS.Recordset.MoveNext
Loop
Printer.EndDoc
End Function
Sub Upd_kin(syntagh, kinhsh, FERE, KEFALH, MGRS, PANE, xronos)
             
    KINHSEIS.Recordset.AddNew
    KINHSEIS.Recordset("SYNTAGH") = syntagh
    KINHSEIS.Recordset("FERE") = FERE
    KINHSEIS.Recordset("kinhsh") = kinhsh
    If kinhsh = 9 Then 'δοσομέτρηση
       KINHSEIS.Recordset("kefalh") = IIf(KEF(1) = FERE, 1, 2)
    Else
       KINHSEIS.Recordset("kefalh") = KEFALH
    End If
    KINHSEIS.Recordset("mgrs") = MGRS
    KINHSEIS.Recordset("pane") = PANE
    KINHSEIS.Recordset("xronos") = xronos
    
    KINHSEIS.Recordset("tkinhsh") = left(tkin(kinhsh), 15)
    KINHSEIS.Recordset("KEF1") = KEF(1)
    KINHSEIS.Recordset("KEF2") = KEF(2)
    KINHSEIS.Recordset("ar_syntag") = m_syntag(syntagh)
    KINHSEIS.Recordset.update
    
    
End Sub



Function Ektel_kinhsh(Move)
  
  Dim Ok_s, Llk, All_Null, dum, alati, m_Counter, mCoun, time_a, time_x
  Dim mgen_xr, dum_x, mz, cmd$, DDDDD, real_dos
  Dim mpoykali, Zygisma, m_synt
Dim UESH
Dim metr
  
  alati = 0
  

mgen_xr = tkin(a_Kinhsh(Move))
Get_Gen_xronos mgen_xr

  
  
Ref.Caption = "--"

  
   Ektel_kinhsh = 1
  Select Case a_Kinhsh(Move)
     Case 1 'ΦΕΡΕ ΠΟΤΗΡΙ
        
    '   On Error Resume Next
    '    Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\watch.mdb")
    '    Set r = dbsjet.OpenRecordset("moves")
    '        r.AddNew
    '        r("datetime") = Now
    '        r("ENTOLH") = Trim(prospau2.Recordset("ENTOLH")) + prospau2.Recordset("AA_PROSPAU")
    '        r.update
    '        r.Close
    '    On Error GoTo 0
     
     
     
       If Not IsNull(prospau2.Recordset("ALATIGR")) Then
          If prospau2.Recordset("ALATIGR") > 0 And UESH_ALATIOY > 0 Then
               alati = 1
         End If
       End If
     
     
     
     
       frmSYNT.Caption = a_Tkinhsh(Move) + str(a_Syntagh(Move))


         
'       dum_x=getcurrenttime():while MilSec 1000
         
        
'If Balance_Type = "ADAM" Then 'metler
        Ref.Caption = "ref:Rez_met 1"
        G_REZERO MSComm1
        
        MilSec 1000
           mz = Zygis0(0)
           metr = 0
                Do While Abs(Val(mz)) > 60
                     Ref.Caption = "ref:Rez_met 2"
                     G_REZERO MSComm1
                      MilSec 1000
                      mz = Zygis0(0)
                      metr = metr + 1
                      If metr > 10 Then
                         Ektel_kinhsh = -21
                         Exit Function
                      End If
                      Ref.Caption = "ref:Rezero Metler"
                Loop
       MilSec 50

'End If
        
        
        
   Ref.Caption = "ref:Bring Beak1"
     
   If alati = 1 Then
           kod6400 = 19000 + a_Syntagh(Move)
   Else
        kod6400 = 3000 + a_Syntagh(Move)
   End If
   
        cmd$ = "!var12=" + str(kod6400) + ":"
   Ref.Caption = "ref:Wait Beak1"
        dum = wait_to_understand(cmd$)
        

        
        
        
        '----------------
        Label3(a_Syntagh(Move)).BackColor = kokkino
              mz = Zygis0(0)
        
        
'If Balance_Type <> "ADAM" Then   '16-9-2002 to metefera parapano mazi me to ADAM
'        Do While Abs(mz) > 60
'             MilSec 2000
'             G_REZERO MSComm1
'              MilSec 400
'              mz = Zygis0(0)
'              If mz <= -9900000 Then
'                    MilSec 1000
'                    mz = Zygis0(0)
'                    If mz = -9900000 Then
'                         Ektel_kinhsh = -21
'                         Exit Function
'                    End If
'             End If
'             Ref.Caption = "ref:Rezero ADAM"'
'
'       Loop
'       MilSec 50
'End If
       
       
       G_Balance_Digits MSComm1, 2
'         MSComm1.Output = Balance_2_Digits + Chr$(13)
         
         MilSec 1000
     
     dum = Wait_Robot(0)
   If alati = 1 Then
      wait_to_understand ("!var12=93:")  ' φερε κεφαλή σκόνης σε ζυγό
              
      Ref.Caption = "ref:Wait_kef_sk"
      dum = Wait_Robot(0)
   End If


       mCoun = 0
51101
        MilSec 500     '  50         500 mono ston griba
              
        Ref.Caption = "ref:Zyg_beak1"

        mz = Zygis0(0)
        If (mz) / 1000 < 170 Then
                MilSec 1000
                mz = Zygis0(0)
                If mz / 1000 < 170 Then
                    mCoun = mCoun + 1
                    If mCoun < 3 Then
                        MilSec 1000
                        GoTo 51101
                    Else
                        Ektel_kinhsh = -20
                    End If
                End If
        End If

    
    
    
    If alati = 1 And Ektel_kinhsh <> -20 Then
51100
    
           Real_Alati = DOS_ALATI_DOK(prospau2.Recordset("ALATIGR"))
         '  GoTo 51100
              prospau2.Recordset.EDIT
              prospau2.Recordset("AlatiGr") = Real_Alati
                 Real_Alati = 0
              prospau2.Recordset.update
         
         
         ' aneba στο πανω Ζ
           ANIXE
         
         RobSend ("d,," + str(var93_anoZ) + ":")
         Find_Axes
         z0 = var93_anoZ
         Ref.Caption = "ref:Wait_v93"
         RobSend_w ("go,,1:")
         
         ' πηγαινε το gripper μπροστά και τον 4 άξονα στο νερό
     ' ANIXE
      
      
      RobSend ("!d" + str(var74) + ",,," + str(var67_3kef) + ":")
         
         Find_Axes
         x0 = var74
         w0 = var67_3kef
         Ref.Caption = "ref:Wait_67_3"

         RobSend_w ("!go1,,,1:")
    
    
    
    End If

     
     
     
Case 2 'ΠΑΝΕ ΠΟΤΗΡΙ
         
        dum = Wait_Robot(0)
         
         
         frmSYNT.Caption = a_Tkinhsh(Move) + str(a_Syntagh(Move))
         kod6400 = 4000 + a_Syntagh(Move)
         cmd$ = "!var12=" + str(kod6400) + ":"
         dum = wait_to_understand(cmd$)
         
               Label3(a_Syntagh(Move)).BackColor = Prasino
  
         MilSec 100
         
         
          G_Balance_Digits MSComm1, 2
          '      MSComm1.Output = Balance_2_Digits + Chr$(13)
        
        
         If Ektelesueises(a_Syntagh(Move)) = 1 Then
            print_rep str(a_Syntagh(Move)), a_Ar_Syntag(Move)
         End If

       MilSec 100
         
                
          G_Balance_Digits MSComm1, 2
  
          '  MSComm1.Output = Balance_2_Digits + Chr$(13)
       
     Case 3 'ΦΕΡΕ ΜΠΟΥΚΑΛΙ
      frmSYNT.Caption = a_Tkinhsh(Move) + str(a_Fere(Move))
       'frmSYNT.caption = ml_CapRes(365) + Str(a_Fere(Move))
       kod6400 = 1000 + a_Fere(Move)
       cmd$ = "!var12=" + str(kod6400) + ":"
       dum = wait_to_understand(cmd$)

     Case 4 'ΠΑΝΕ & ΦΕΡΕ ΜΠΟΥΚΑΛΙ
          If a_Syntagh(Move) = 4 Then
               DDDDD = 0
         End If
          
           If a_Pane(Move) <> a_Fere(Move) Then
              '       frmSYNT.caption = ml_CapRes(382) + Str$(a_Pane(Move)) + mL_Res(383) + Str$(a_Fere(Move))
              Erxetai = a_Fere(Move)
              Feygei = a_Pane(Move)
              frmSYNT.Caption = a_Tkinhsh(Move) + str(a_Pane(Move)) + " bring " + str(a_Fere(Move))
              kod6400 = 100000 + a_Pane(Move)
              cmd$ = "!var12=" + str(kod6400) + ":!var13=" + str(a_Fere(Move)) + ":"
              Ref.Caption = "ref:Wait_panfer"
               
              dum = wait_to_understand(cmd$)
          End If
     Case 5 'ΠΑΝΕ ΜΠΟΥΚΑΛΙ
          frmSYNT.Caption = " Return  " + str(a_Fere(Move))
           ' frmSYNT.caption = ml_CapRes(369) + PER(kke) + Str(kod6400 - 2000)
           kod6400 = 2000 + a_Fere(Move)
          cmd$ = "!var12=" + str(kod6400) + ":"
          Ref.Caption = "ref:Wait_pane"

          dum = wait_to_understand(cmd$)
     Case 6 'ΦΕΡΕ ΚΕΦΑΛΗ 1 ΣΕ ΖΥΓΑΡΙΑ (ΔΕΞΙΑ)
            frmSYNT.Caption = "ΚΕΦΑΛΗ 1 ΣΕ ΖΥΓΑΡΙΑ "
            If var65_1kef = Uesh_4axona() Then
            Else
                 porta = Porta_Dex      '  "1"  ' δεξιά κεφαλ
                cmd$ = "!var12=90:"
                
            End If
          Ref.Caption = "ref:WaitU_kef1"
          
          dum = wait_to_understand(cmd$)
          Ref.Caption = "ref:Wait_kef1"

          dum = Wait_Robot(0)
        

     Case 7 'ΦΕΡΕ ΚΕΦΑΛΗ 2 ΣΕ ΖΥΓΑΡΙΑ   (ARISTERH)
        frmSYNT.Caption = "ΚΕΦΑΛΗ 2 ΣΕ ΖΥΓΑΡΙΑ "
        If var66_2kef = Uesh_4axona() Then
        Else
            porta = Porta_Dex      '  "1"  ' δεξιά κεφαλ
          
          cmd$ = "!var12=91:"
          Ref.Caption = "ref:WaitU_kef2"
       
          dum = wait_to_understand(cmd$)
          Ref.Caption = "ref:Wait_kef2"

          dum = Wait_Robot(0)

       End If

     Case 8 'ΔΟΣΟΜΕΤΡΗΣΗ ΝΕΡΟΥ
            ' MSComm1.Output = "Z " + Chr$(13) + Chr(10)
            G_REZERO MSComm1
           MilSec 2500
           ' G_REZERO MSComm1
           
           If aritmos_ximikon > 0 Or Len(XimNer_on) > 5 Then ' h exo ximika dipla sto nero
                 m_synt = a_Ar_Syntag(Move)
                 real_dos = DISPENSE_XIMIKA(m_synt)
                 prospau2.Recordset.EDIT
                 prospau2.Recordset("ximikon") = real_dos
                 ' prospau2.Recordset("date_disp") = Now '  time_x / 100
                 prospau2.Recordset.update
                 
                 If m_ximika(50, 5) = -31 Then
                    Ektel_kinhsh = -31   ' Ektel_kinhsh = -31 Then
                    Exit Function
                 End If
                 
            '     Exit Function
           End If
If Abs(Zygis0(0)) > 5 Then
           G_REZERO MSComm1
           MilSec 500
End If




          ' MSComm1.Output = Balance_3_Digits + Chr$(13)
'          dum = gyrise_zygaria(3)
           MilSec 500
           real_dos = 0
           If Not IsNull(prospau2.Recordset("nero")) Then
33:
               If system_ready = 0 Then
                  real_dos = prospau2.Recordset("nero") + 1
               Else
                  real_dos = NER_DOSOMETRHSH(prospau2.Recordset("nero"), Porta_Ner)
               End If
                
               Get_Gen_xronos "telos neroy"
                
                
                
                
                ' GoTo 33
           End If
            frmSYNT.Caption = frmSYNT.Caption + " dispense "
   
           time_a = InStr(Time$, ":")
           time_x = Mid$(Time$, 1, time_a - 1) + Mid$(Time$, time_a + 1, 2)
           If real_dos < -111000222 Then
                Ektel_kinhsh = -30
                prospau2.Recordset.EDIT
                prospau2.Recordset("RENERO") = 999999
                prospau2.Recordset("date_disp") = Now   ' time_x / 100
                prospau2.Recordset.update
           Else
                prospau2.Recordset.EDIT
                prospau2.Recordset("RENERO") = real_dos
                prospau2.Recordset("date_disp") = Now '  time_x / 100
                prospau2.Recordset.update
           End If
      Case 9 'ΔΟΣΟΜΕΤΡΗΣΗ ΧΡΩΜΑΤΟΣ
             frmSYNT.Caption = frmSYNT.Caption + " dispense  " + str(a_Fere(Move))
             Timer1.Enabled = False
             Label8 = ""
             MSComm1.PortOpen = False
             
             If system_ready = 0 Then
                Label8 = a_Mgrs(Move) + 1
             Else

m_Counter = 0
111:
                
          ' If Balance_Type = "ADAM" Then
          '      Syn_Dos2.Asked_Q = a_Mgrs(Move) '  ZHTOYMENH ΠΟΣΟTHTA ΣE MGRS
          '      Syn_Dos2.Level = Ypoloipo(a_Fere(Move))      'ΓΡΑΜΑΡΙΑ ΜΠΟΥΚΑΛΙΟΥ
          '      Syn_Dos2.Valve_on = IIf(a_Kefalh(Move) = 1, valve_1_on, valve_2_on) ' ΑΝΟΙΓΜΑ ΒΑΛΒΙΔΑΣ
          '      Syn_Dos2.Valve_off = IIf(a_Kefalh(Move) = 1, valve_1_off, valve_2_off)   ' ΚΛΕΙΣΙΜΟ ΒΑΛΒΙΔΑΣ
          '      Syn_Dos2.Caption = str(a_Fere(Move)) + " : " + Perigrafh(a_Fere(Move)) 'περιγραφη χρώματος που δοσομετρείται
          '      Syn_Dos2.Show 1
          ' Else
                Syn_Dos.asked_q = a_Mgrs(Move) '  ZHTOYMENH ΠΟΣΟTHTA ΣE MGRS
                Syn_Dos.Level = Ypoloipo(a_Fere(Move))      'ΓΡΑΜΑΡΙΑ ΜΠΟΥΚΑΛΙΟΥ
                Syn_Dos.Valve_On = IIf(a_Kefalh(Move) = 1, valve_1_on, valve_2_on) ' ΑΝΟΙΓΜΑ ΒΑΛΒΙΔΑΣ
                Syn_Dos.Valve_Off = IIf(a_Kefalh(Move) = 1, valve_1_off, valve_2_off)   ' ΚΛΕΙΣΙΜΟ ΒΑΛΒΙΔΑΣ
                Syn_Dos.Caption = str(a_Fere(Move)) + " : " + Perigrafh(a_Fere(Move)) 'περιγραφη χρώματος που δοσομετρείται
                Syn_Dos.Show 1
           'End If
            
         '  GoTo 111
               If Val(Label8.Caption) = -5 Or Label8.Caption = "0" Then
                  m_Counter = m_Counter + 1
                  DoEvents
                  '   If m_Counter < 30 Then GoTo 111
               End If
               
             End If
           
            Ref.Caption = "ref:Dispensing"
            
            Do
               DoEvents
            Loop Until Label8 <> ""
             
             
             Get_Gen_xronos "telos dosom"

             
             
             On Error Resume Next
             MSComm1.PortOpen = True
             Timer1.Enabled = True
             On Error GoTo 0
             mpoykali = a_Fere(Move)
             UESH = a_Pane(Move)
             Zygisma = Val(Label8.Caption)
             Label9.Caption = Label8.Caption
             a = Update2_PROSPAU(mpoykali, UESH, Zygisma)
             Ektel_kinhsh = IIf(Val(Label8.Caption) < 0, Val(Label8.Caption), 1)
End Select

End Function

Function DOS_ALATI(M_ALATI)

    '------dont use it!!!!!!!!!!!------------------
    Dim APOBARO, Apoklish, DIADROMH_BOLAN, Arxh, TIME_ANOXH, Z, ok, TaxyThtes

    Dim counter, FIRST_TIME, Motori_On, dum

    Motori_On = 0
    FIRST_TIME = True
    counter = 0   'ΔΕΣ ΤΟ   VAR88     !!!!!!!!
    Apoklish = 30
    DIADROMH_BOLAN = 12000 ' BHMATA
    TIME_ANOXH = 500000   '50SEC
    TaxyThtes = "!V,,28,0.5:"
    Label3(0).Caption = M_ALATI
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
    dum = wait_to_understand("!var12=94:")
     
    'ΚΡΑΤΑ ΑΠΟΒΑΡΟ
    APOBARO = Zygis0(0)
     
    dum = Wait_Robot(0)
    
    Arxh = GetCurrentTime()
  
    RobSend ("!v, , .01:")
    MilSec 500
  
    RobSend ("!v, , .01:")

    'ΞΕΚΙΝΑ ΤΟ ΜΟΤΟΡΙ  "SKON_MOTOR_ON"
    RobSend ("!" + skon_motor_on + ":")
    Motori_On = 1

    DIADROMH_BOLAN = DIADROMH_BOLAN / 30
    ok = 0

    ' ΤΟ ΤΕΡΜΑ ΤΟΥ ΒΟΛΑΝ  ΕΙΝΑΙ    VAR(88)+12000 bhmata
    Do While True
  
        ' ανεβασμα1
        RobSend ("!d , ," + str((-500 + var88) + DIADROMH_BOLAN) + ":!go,,1:")

        Do While axe(3) < (-500 + var88) + DIADROMH_BOLAN
            Z = Zygis0(0) - APOBARO
            counter = counter + 1
            frmSYNT.Caption = str(counter) + "  ^   " + str(Z)

            If Z < -999000000 Then
                RobSend ("!" + skon_motor_off + ":!k:")
              
                Do While Z < -999000000
                    Z = Zygis0(0) - APOBARO
                Loop
              
                RobSend ("!" + skon_motor_on + ":!go,,1:")
            End If

            If Z >= M_ALATI - Apoklish * 10 And FIRST_TIME = True Then
                DIADROMH_BOLAN = DIADROMH_BOLAN / 10
                RobSend ("!" + skon_motor_off + ":!k:")
                FIRST_TIME = False
                RobSend ("!v, ,0.1:")
                MilSec 300
                RobSend ("!v, ,0.1:")
                Motori_On = 0

                Exit Do

            End If

            If Z >= M_ALATI - Apoklish Then
                DOS_ALATI = Z
                ok = 1

                Exit Do

            End If

        Loop
  
        If ok = 1 Then Exit Do
        ' κατεβασμα
        RobSend ("!d , ," + str(-500 + var88) + ":!go,,1:")

        Do While Abs(axe(3) - (-500 + var88)) > 10 ' axe(3)>var88
            Z = Zygis0(0) - APOBARO
            counter = counter + 1
            frmSYNT.Caption = str(counter) + "     " + str(Z)
          
            If Z < -999000000 Then
                RobSend ("!" + skon_motor_off + ":!k:")

                Do While Z < -999000000
                    Z = Zygis0(0) - APOBARO
                Loop

                RobSend ("!" + skon_motor_on + ":!go,,1:")
            End If
        
            If Z >= M_ALATI - Apoklish * 10 And FIRST_TIME = True Then
                DIADROMH_BOLAN = DIADROMH_BOLAN / 10
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
  
        If ok = 1 Then Exit Do
  
        If Motori_On = 0 Then RobSend ("!" + skon_motor_on + ":!k:")

    Loop
   
    RobSend ("!" + skon_motor_off + ":!k:")
  
    RobSend (TaxyThtes)
    MilSec 300
    RobSend (TaxyThtes)

    ' kleisto to mpoykali   GMPA@STIAFILCO.COM   DONTAS
    RobSend ("!d , ," + str(-500 + var88) + ":go,,1:")
  
    MilSec 1500
    RobSend ("!" + skon_motor_off + ":!k:")
    DOS_ALATI = Zygis0(0) - APOBARO
    Label8.Caption = DOS_ALATI

End Function

'
'
'
Private Sub com_file_exec_Click()
Dim Last_Epil, Message, SYN_ALATIOY
Dim m_Counter, dum, Bohu(1 To 19), Taytoxronismos, result, DOSO_MPOYK, aas, k
Dim response, cmd$, TEMP%, PHDHMENA_POTHRIA, mkin
Dim mm_Synt

'On Error Resume Next
 General_Xronos = GetCurrentTime()
'grid.RecordSource = "select prospau2.*,ximitech.perigrafh,ximitech.addr_prot from prospau2 inner join ximitech on prospau2.kod=ximitech.kod where entolh='" + prospau2.Recordset("entolh") + "'and aa_prospau='" + prospau2.Recordset("aa_prospau") + "'  order by epil,entolh,aa_prospau,seira"

 Com_file_exec.Enabled = False
 
 Do While Is_Porta() = "α"
        response = MsgBox(mL_Res(522), 64, "Talos")    ' klise porta
 Loop
 If aritmos_ximikon > 0 Then
    If calc_ximika() = 0 Then
       MsgBox mLResnew(920, "Chemicals without calibration*", 1)
       Exit Sub
    End If
 End If
 
 ' Sort_Recipes
 prospau2.Refresh
 MilSec 1000
 
 
 
 SYN_ALATIOY = 0
 prospau2.Recordset.MoveFirst
 Do While Not prospau2.Recordset.eof
    If Not IsNull(prospau2.Recordset("ALATIGR")) Then
        SYN_ALATIOY = SYN_ALATIOY + prospau2.Recordset("ALATIGR")
    End If
    prospau2.Recordset.MoveNext
    'Grid.Refresh
 Loop
 If DOYLEYO_ALATI = 0 Or UESH_ALATIOY = 0 Then SYN_ALATIOY = 0
 
 If SYN_ALATIOY > 0 Then
       ' πανε τον  W στην θέση της σκόνης
         dum = wait_to_understand("!var12=93:")
         dum = Wait_Robot(0)
         '
         ' φέρε το αλάτι στην κεφαλή της σκόνης
         dum = wait_to_understand("!var12=" + LTrim(str(1000 + UESH_ALATIOY)) + ":!var13=1:")
         dum = Wait_Robot(0)
         ANIXE
         ' aneba στο πανω Ζ
         RobSend ("!d,," + str(var93_anoZ) + ":")
         Find_Axes
         z0 = var93_anoZ
         RobSend_w ("!go,,1:")
         
         
         ' πηγαινε το gripper μπροστά
         RobSend ("!d" + str(var74) + ":")
         
         Find_Axes
         x0 = var74
         RobSend_w ("!go1:")
         
' -------------------  ΔΟΚΙΜΑΣΤΙΚΟ    ---------------------------------------
' prospau2.Recordset.MoveFirst
' Do While Not prospau2.Recordset.EOF
'    If Not IsNull(prospau2.Recordset("ALATIGR")) Then
'          If prospau2.Recordset("ALATIGR") > 0 Then
'                DOS_ALATI (prospau2.Recordset("ALATIGR"))
'         End If
'    End If
'    prospau2.Recordset.MoveNext
' Loop
' -------------------  ΔΟΚΙΜΑΣΤΙΚΟ    ---------------------------------------
End If
'
 
   Printer.ForeColor = BLE
 
 
 
 Com_file_exec.Enabled = False
 
 
 aas = MAKE_FILE()
 If aas > 0 Then
    Exit Sub
 End If
  Command1.Enabled = False
 
 For k = 1 To 40
   Ektelesueises(k) = 0
 Next
  
  
  'start.Enabled = False
  Text1.Enabled = False
  FindValves
  'system_ready = IsOSLoaded(768)
If system_ready = 1 Then
  'If Len(Motor_Input) > 0 And FindInput(Val(Motor_Input)) = 0 Then
      '   MsgBox "Οι κινητήρες είναι εκτός λειτουργίας. " + Chr(13) + "Βγείτε από το παρόν και ακολουθήστε τη διαδικασία σύντομο ΗΟΜΕ !", , "Talos : ΠΡΟΣΟΧΗ ! "
   '   MsgBox mL_Res(349), , "Talos"
   '   Command1.Enabled = True
    '  Exit Sub
  'End If
  'If Motor_Input <> "" Then
   '     cmd$ = "!" + Motor_Output_on + ":"
    '     TEMP% = SendATBlock(768, cmd$, 0)
  'End If
  'DUM = begin_talos(0)
End If
Timer1.Enabled = True
XALASE_NERO = 0

   'ΜΑΡΚΑΡΩ ΤΙΣ ΚΙΝΗΣΕΙΣ ΠΟΥ ΕΚΤΕΛΟΥΝΤΑΙ ΤΑΥΤΟΧΡΟΝΑ
 prospau2.Recordset.MoveFirst
 Last_Epil = prospau2.Recordset("epil")

PHDHMENA_POTHRIA = Val(Text1.text) - 1
PHDHMENA_POTHRIA = IIf(PHDHMENA_POTHRIA < 0, 0, PHDHMENA_POTHRIA)
PHDHMENA_POTHRIA = IIf(PHDHMENA_POTHRIA >= Ar_Pothr, 0, PHDHMENA_POTHRIA)
   
DOSO_MPOYK = 0  'ΔΟΣΟΜΕΤΡΟΥΜΕΝΟ ΜΠΟΥΚΑΛΙ ΣΤΟ ΟΠΟΙΟ ΣΥΝΕΒΗ ΒΛΑΒΗ
   
   KINHSEIS.Recordset.MoveFirst
   Do While Not KINHSEIS.Recordset.eof
       Pothri = KINHSEIS.Recordset("SYNTAGH")
       mm_Synt = KINHSEIS.Recordset("ar_syntag")
       'ελέγχω την επόμενη κίνηση μήπως δουλεύουν ταυτόχρονα
    
       Taytoxronismos = 0
       Load_kinhsh (1)
       mkin = KINHSEIS.Recordset("kinhsh")
       
    
       
       KINHSEIS.Recordset.MoveNext
       If KINHSEIS.Recordset.eof Then
          KINHSEIS.Recordset.MoveLast
          Taytoxronismos = 0
       Else
          Load_kinhsh (2)
          If a_Xronos(1) = a_Xronos(2) Then
             Taytoxronismos = 1
           Else
            Taytoxronismos = 0
            KINHSEIS.Recordset.MovePrevious
          End If
       End If
       
       If Taytoxronismos = 1 Then
          
          'πρώτα εκτελώ την κίνηση & μετά την δοσομέτρηση
           If a_Kinhsh(1) = 9 Then 'ΔΟΣΟΜΕΤΡΗΣΗ ΧΡΩΜΑΤΟΣ ΣΤΟ (1)
               
               
                 If a_Fere(1) = a_Kef1(1) Then  'ara θα δοσομετρήσω την 1η κεφαλή
                      'κινησε τον 4ο στην 1η κεφαλή
                     If var65_1kef = Uesh_4axona() Then
                     Else
                          cmd$ = "var12=90:"
                          dum = wait_to_understand(cmd$)
                          dum = Wait_Robot(0)
                    End If
                 Else
                      'κινησε τον 4ο στην 2η κεφαλή
                        If var66_2kef = Uesh_4axona() Then
                        Else
                               cmd$ = "var12=91:"
                               dum = wait_to_understand(cmd$)
                               dum = Wait_Robot(0)
                        End If
                 End If
                    result = Ektel_kinhsh(2)
                    result = Ektel_kinhsh(1) 'ΔΟΣΟΜΕΤΡΗΣΗ
                    
                           Last_Epil = prospau2.Recordset("epil")
                           prospau2.Recordset.MoveNext
                           'Grid.Refresh
                    
                    
                    
                    
                    DOSO_MPOYK = a_Fere(1)
                    dum = Wait_Robot(0)
           
           
           ElseIf a_Kinhsh(2) = 9 Then   ' ΔΟΣΟΜΕΤΡΗΣΗ ΧΡΩΜΑΤΟΣ ΣΤΟ (2)
                 
                 If a_Fere(2) = a_Kef1(2) Then  'ara θα δοσομετρήσω την 1η κεφαλή
                      'κινησε τον 4ο στην 1η κεφαλή
                     If var65_1kef = Uesh_4axona() Then
                     Else
                      cmd$ = "var12=90:"
                      dum = wait_to_understand(cmd$)
                      dum = Wait_Robot(0)
                    End If
                 Else
                      'κινησε τον 4ο στην 2η κεφαλή
                     If var66_2kef = Uesh_4axona() Then
                     Else
                         cmd$ = "var12=91:"
                         dum = wait_to_understand(cmd$)
                         dum = Wait_Robot(0)
                      End If
                 End If
           
           
                   result = Ektel_kinhsh(1)
                   result = Ektel_kinhsh(2) 'ΔΟΣΟΜΕΤΡΗΣΗ
                   Last_Epil = prospau2.Recordset("epil")
                   prospau2.Recordset.MoveNext
                   'Grid.Refresh
                   DOSO_MPOYK = a_Fere(2)
                   dum = Wait_Robot(0)
           Else  ' ΔΕΝ ΕΧΩ ΔΟΣΟΜΕΤΡΗΣΗ ΧΡΩΜΑΤΟΣ
           
          
              If a_Kinhsh(1) = 8 Then 'ΔΟΣΟΜΕΤΡΗΣΗ ΝΕΡΟΥ
                  result = Ektel_kinhsh(2)
                  result = Ektel_kinhsh(1)
              Else
                  result = Ektel_kinhsh(1)
                  result = Ektel_kinhsh(2)
              End If
              dum = Wait_Robot(0)
           End If
       
       Else '---------- ΟΧΙ ΤΑΥΤΟΧΡΟΝΙΣΜΟΣ ------------------------------------------
              If a_Kinhsh(1) = 9 Then  'αν είναι ΔΟΣΟΜΕΤΡΗΣΗ
                 If a_Fere(1) = a_Kef1(1) Then  'ara θα δοσομετρήσω την 1η κεφαλή
                      'κινησε τον 4ο στην 1η κεφαλή
                    
                       If var65_1kef = Uesh_4axona() Then
                       Else
                           cmd$ = "var12=90:"
                           dum = wait_to_understand(cmd$)
                           dum = Wait_Robot(0)
                       End If
                 Else
                      'κινησε τον 4ο στην 2η κεφαλή
                        If var66_2kef = Uesh_4axona() Then
                        Else
                           cmd$ = "var12=91:"
                           dum = wait_to_understand(cmd$)
                           dum = Wait_Robot(0)
                        End If
                 End If
               End If
               
               
               
                result = Ektel_kinhsh(1)
                 DOSO_MPOYK = a_Fere(1)

               If a_Kinhsh(1) <> 9 Then
                    dum = Wait_Robot(0)
               Else
                    Last_Epil = prospau2.Recordset("epil")
                    prospau2.Recordset.MoveNext
                    'Grid.Refresh
               End If
       End If
       
       If result < -1 And result <> -7 Then     ' -7=απετυχε ο αλγόριθμος ζύγισης
            
          On Error Resume Next
              ' print_rep str(a_Syntagh(Move)), a_Ar_Syntag(Move)
               print_rep Pothri, mm_Synt
          On Error GoTo 0
            
            Ektelesueises(Pothri) = -1  ' φλαγκ που λέει ότι εκτελείται κανονικά η συνταγή
            Exit Do
       End If
       
       
       frmSYNT.Caption = ""
       
       Get_Gen_xronos "telos syntaghs"

       
       
       KINHSEIS.Recordset.MoveNext
       If KINHSEIS.Recordset.eof Then
             result = 0

             Exit Do
       Else
          
          If a_Syntagh(2) > a_Syntagh(1) Then  'ΑΝ ΠΗΓΕ ΣΤΗΝ ΕΠΟΜΕΝΗ ΣΥΝΤΑΓΗ

          End If
       End If
   
       Ektelesueises(Pothri) = 1  ' φλαγκ που λέει ότι εκτελείται κανονικά η συνταγή
   
   Loop
     
   

' βαζω status=2 για τις εκτελεσθείσες συνταγές
'  prospau2.Recordset.MoveLast
'  Do While Not prospau2.Recordset.BOF
'       m2_syntag = Trim(prospau2.Recordset("ENTOLH")) + "-" + LTrim(prospau2.Recordset("aa_PROSPAU"))
'       If prospau2.Recordset("userE") > 0 Then Updating = 1 Else Updating = 0
'       Do While Not prospau2.Recordset.BOF And m2_syntag = prospau2.Recordset("ENTOLH") + "-" + LTrim(prospau2.Recordset("aa_prospau"))
'         If Updating = 1 Then
'             prospau2.Recordset.Edit
'                prospau2.Recordset("status") = 2
'             prospau2.Recordset.update
'           End If
'          prospau2.Recordset.MovePrevious
'           If prospau2.Recordset.BOF Then Exit Do
'       Loop
'       If prospau2.Recordset.BOF Then Exit Do
'   Loop
   
     
     
     
' ΠΡΟΣΟΧΗ ΜΕΤΑ ΤΙΣ ΔΟΚΙΜΕΣ ΝΑ ΓΙΝΕΙ  "< 0 Then"
   
     Message = mL_Res(875) ' "ΟΛΕΣ ΟΙ ΣΥΝΤΑΓΕΣ ΟΛΟΚΛΗΡΩΘΗΚΑΝ ΜΕ ΕΠΙΤΥΧΙΑ"
   
   If result < -1 Then
    
    a_Syntagh(1) = 0
    a_Kinhsh(1) = KINHSEIS.Recordset("kinhsh")
    a_Kefalh(1) = KINHSEIS.Recordset("kefalh")
    a_Mgrs(1) = 0
    a_Pane(1) = 0
    a_Xronos(1) = 0
    a_Tkinhsh(1) = "Return of.."
    
    a_Kef1(1) = KINHSEIS.Recordset("kef1")
    a_Kef2(1) = KINHSEIS.Recordset("kef2")
    
    ' ΕΠΙΣΤΡΟΦΗ ΜΠΟΥΚΑΛΑΣ 1ης ΚΕΦΑΛΗΣ
      If a_Kef1(1) > 0 Then
        
                       'κινησε τον 4ο στην 2η κεφαλή
                       If var66_2kef = Uesh_4axona() Then
                       Else
                            cmd$ = "var12=91:"
                           dum = wait_to_understand(cmd$)
                           dum = Wait_Robot(0)
                       End If
            
        
        
        a_Kinhsh(1) = 5
         a_Fere(1) = a_Kef1(1)
         dum = Ektel_kinhsh(1)
         dum = Wait_Robot(0)
      End If
      
    ' ΕΠΙΣΤΡΟΦΗ ΜΠΟΥΚΑΛΑΣ 2ης ΚΕΦΑΛΗΣ
      If a_Kef2(1) > 0 Then
                      'κινησε τον 4ο στην 1η κεφαλή
                       If var65_1kef = Uesh_4axona() Then
                       Else
                            cmd$ = "var12=90:"
                           dum = wait_to_understand(cmd$)
                           dum = Wait_Robot(0)
                       End If
       
         
         a_Kinhsh(1) = 5
         a_Fere(1) = a_Kef2(1)
          dum = Ektel_kinhsh(1)
          dum = Wait_Robot(0)
      End If
         
      ' επιστροφη ποτηριού
         a_Kinhsh(1) = 2
         a_Syntagh(1) = Pothri
         dum = Ektel_kinhsh(1)
         dum = Wait_Robot(0)
       
       
       
       Select Case result
            
          Case -31
             '
             Message = mLResnew(932, "Check the level of Chemicals or Calibrate them.*", 1)
          Case -20
             ' MsgBox "ΔΕΝ ΒΡΗΚΑ ΤΟ ΠΟΤΗΡΙ " + Str(POTHRI)
             Message = mL_Res(876) + str(Pothri) '"ΔΕΝ ΒΡΗΚΑ ΤΟ ΠΟΤΗΡΙ "
          
          Case -21
             '"ΔΙΑΚΟΠΗ ΣΥΝΤΑΓΩΝ ΣΤΟ ΠΟΤΗΡΙ  "
             Message = "877-878" + mL_CapRes(877) + str(Pothri) + mL_CapRes(878) '  " ΛΟΓΩ ΜΗ ΛΕΙΤΟΥΡΓΙΑΣ ΤΟΥ ΖΥΓΟΥ "
          Case -2
             Message = mL_Res(879) + str(DOSO_MPOYK) ' "η βαλβίδα δεν τρέχει στο μπουκάλι "
          Case -3
             Message = mL_Res(880) ' " η ζυγαρια υπερφορτώθηκε"
          Case -4
             Message = mL_Res(881) ' " δεν έγινε διαδικασία Εναρξης"
          Case -5
             Message = mL_Res(882) '"Η ζυγαριά δεν ανταποκρίνεται"
           Case -6
             Message = mL_Res(883) ' " Η ζυγαριά εκτός περιοχής"
           Case -7
             Message = mL_Res(884) ' " Η ζύγιση έχει αποτύχει"
             ' MsgBox
           Case -8
             Message = mL_Res(885) ' " Παρουσιάσθηκε απροσδόκητο λάθος στον αλγόρυθμο ζύγισης"
             'MsgBox
           Case -9
              Message = mL_Res(886) ' " Η ζύγιση ακυρώθηκε από τον χρήστη"
              '   MsgBox " Η ζύγιση ακυρώθηκε από τον χρήστη"
           Case -10
              Message = mL_Res(887) '" Η στάθμη του μπουκαλιού είναι πολύ χαμηλή <50 gr " + Str(DOSO_MPOYK)
              ' MsgBox
           Case -11
              Message = mL_Res(888) ' " Υπερβολική  ζητούμενη ποσότητα  (max. 100 gr)"
              '           MsgBox
        End Select
        
       
       Printer.Font = "Tahoma"
       Printer.FontSize = 12

       Printer.Print Message
       Printer.EndDoc
        
       On Error Resume Next
       
       Last_Epil = prospau2.Recordset("epil")
        
       prospau2.Recordset.MoveFirst
        Do While Not prospau2.Recordset.eof
            If Last_Epil = prospau2.Recordset("epil") Then
                prospau2.Recordset.EDIT
                  prospau2.Recordset("status") = 1
                prospau2.Recordset.update
            End If
            prospau2.Recordset.MoveNext
        Loop
        
       On Error GoTo 0
        
   End If
     
     
    dum = Wait_Robot(0)
     
If SYN_ALATIOY > 0 Then
         
       ' πανε τον  W στην θέση της σκόνης
         dum = wait_to_understand("!var12=93:")
         dum = Wait_Robot(0)
         
         
         
         ' φέρε το αλάτι στην κεφαλή της σκόνης
         dum = wait_to_understand("!var12=" + LTrim(str(2000 + UESH_ALATIOY)) + ":")
         dum = Wait_Robot(0)
     
     
     
End If
     
     kod6400 = 82  ' parking
     cmd$ = "var12=82:"
     
    dum = wait_to_understand(cmd$)
    dum = Wait_Robot(0)
     
     
     
         ' ΧΤΥΠΗΜΑ ΚΟΥΔΟΥΝΙΟΥ
 cmd$ = Alarm2_off + ":"
RobSend (cmd$)

cmd$ = Alarm2_on + ":"
RobSend (cmd$)


 Command1.Enabled = True

 MsgBox Message


'On Error GoTo 0
End Sub

Private Sub com_file_exec_old_Click()
Dim Last_Epil, Message, SYN_ALATIOY
Dim m_Counter, dum, Bohu(1 To 19), Taytoxronismos, result, DOSO_MPOYK, aas, k
Dim response, cmd$, TEMP%, PHDHMENA_POTHRIA, mkin
Dim mm_Synt

'On Error Resume Next
 General_Xronos = GetCurrentTime()
'grid.RecordSource = "select prospau2.*,ximitech.perigrafh,ximitech.addr_prot from prospau2 inner join ximitech on prospau2.kod=ximitech.kod where entolh='" + prospau2.Recordset("entolh") + "'and aa_prospau='" + prospau2.Recordset("aa_prospau") + "'  order by epil,entolh,aa_prospau,seira"

 Com_file_exec.Enabled = False
 
 Do While Is_Porta() = "α"
        response = MsgBox(mL_Res(522), 64, "Talos")    ' klise porta
 Loop
 If aritmos_ximikon > 0 Then
    If calc_ximika() = 0 Then
       MsgBox mLResnew(920, "Chemicals without calibration*", 1)
       Exit Sub
    End If
 End If
 
 ' Sort_Recipes
 prospau2.Refresh
 MilSec 1000
 
 
 
 SYN_ALATIOY = 0
 prospau2.Recordset.MoveFirst
 Do While Not prospau2.Recordset.eof
    If Not IsNull(prospau2.Recordset("ALATIGR")) Then
        SYN_ALATIOY = SYN_ALATIOY + prospau2.Recordset("ALATIGR")
    End If
    prospau2.Recordset.MoveNext
    'Grid.Refresh
 Loop
 If DOYLEYO_ALATI = 0 Or UESH_ALATIOY = 0 Then SYN_ALATIOY = 0
 
 If SYN_ALATIOY > 0 Then
       ' πανε τον  W στην θέση της σκόνης
         dum = wait_to_understand("!var12=93:")
         dum = Wait_Robot(0)
         '
         ' φέρε το αλάτι στην κεφαλή της σκόνης
         dum = wait_to_understand("!var12=" + LTrim(str(1000 + UESH_ALATIOY)) + ":!var13=1:")
         dum = Wait_Robot(0)
         ANIXE
         ' aneba στο πανω Ζ
         RobSend ("!d,," + str(var93_anoZ) + ":")
         Find_Axes
         z0 = var93_anoZ
         RobSend_w ("!go,,1:")
         
         
         ' πηγαινε το gripper μπροστά
         RobSend ("!d" + str(var74) + ":")
         
         Find_Axes
         x0 = var74
         RobSend_w ("!go1:")
         
' -------------------  ΔΟΚΙΜΑΣΤΙΚΟ    ---------------------------------------
' prospau2.Recordset.MoveFirst
' Do While Not prospau2.Recordset.EOF
'    If Not IsNull(prospau2.Recordset("ALATIGR")) Then
'          If prospau2.Recordset("ALATIGR") > 0 Then
'                DOS_ALATI (prospau2.Recordset("ALATIGR"))
'         End If
'    End If
'    prospau2.Recordset.MoveNext
' Loop
' -------------------  ΔΟΚΙΜΑΣΤΙΚΟ    ---------------------------------------
End If
'
 
   Printer.ForeColor = BLE
 
 
 
 Com_file_exec.Enabled = False
 
 
 aas = MAKE_FILE()
 If aas > 0 Then
    Exit Sub
 End If
  Command1.Enabled = False
 
 For k = 1 To 40
   Ektelesueises(k) = 0
 Next
  
  
  'start.Enabled = False
  Text1.Enabled = False
  FindValves
  ' system_ready = IsOSLoaded(768)
If system_ready = 1 Then
  If Len(Motor_Input) > 0 And FindInput(Val(Motor_Input)) = 0 Then
      '   MsgBox "Οι κινητήρες είναι εκτός λειτουργίας. " + Chr(13) + "Βγείτε από το παρόν και ακολουθήστε τη διαδικασία σύντομο ΗΟΜΕ !", , "Talos : ΠΡΟΣΟΧΗ ! "
      MsgBox mL_Res(349), , "Talos"
      Command1.Enabled = True
      Exit Sub
  End If
  If Motor_Input <> "" Then
        cmd$ = "!" + Motor_Output_on + ":"
         TEMP% = SendATBlock(768, cmd$, 0)
  End If
  dum = begin_talos(0)
End If
Timer1.Enabled = True
XALASE_NERO = 0

   'ΜΑΡΚΑΡΩ ΤΙΣ ΚΙΝΗΣΕΙΣ ΠΟΥ ΕΚΤΕΛΟΥΝΤΑΙ ΤΑΥΤΟΧΡΟΝΑ
 prospau2.Recordset.MoveFirst
 Last_Epil = prospau2.Recordset("epil")

PHDHMENA_POTHRIA = Val(Text1.text) - 1
PHDHMENA_POTHRIA = IIf(PHDHMENA_POTHRIA < 0, 0, PHDHMENA_POTHRIA)
PHDHMENA_POTHRIA = IIf(PHDHMENA_POTHRIA >= Ar_Pothr, 0, PHDHMENA_POTHRIA)
   
DOSO_MPOYK = 0  'ΔΟΣΟΜΕΤΡΟΥΜΕΝΟ ΜΠΟΥΚΑΛΙ ΣΤΟ ΟΠΟΙΟ ΣΥΝΕΒΗ ΒΛΑΒΗ
   
   KINHSEIS.Recordset.MoveFirst
   Do While Not KINHSEIS.Recordset.eof
       Pothri = KINHSEIS.Recordset("SYNTAGH")
       mm_Synt = KINHSEIS.Recordset("ar_syntag")
       'ελέγχω την επόμενη κίνηση μήπως δουλεύουν ταυτόχρονα
    
       Taytoxronismos = 0
       Load_kinhsh (1)
       mkin = KINHSEIS.Recordset("kinhsh")
       
    
       
       KINHSEIS.Recordset.MoveNext
       If KINHSEIS.Recordset.eof Then
          KINHSEIS.Recordset.MoveLast
          Taytoxronismos = 0
       Else
          Load_kinhsh (2)
          If a_Xronos(1) = a_Xronos(2) Then
             Taytoxronismos = 1
           Else
            Taytoxronismos = 0
            KINHSEIS.Recordset.MovePrevious
          End If
       End If
       
       If Taytoxronismos = 1 Then
          
          'πρώτα εκτελώ την κίνηση & μετά την δοσομέτρηση
           If a_Kinhsh(1) = 9 Then 'ΔΟΣΟΜΕΤΡΗΣΗ ΧΡΩΜΑΤΟΣ ΣΤΟ (1)
               
               
                 If a_Fere(1) = a_Kef1(1) Then  'ara θα δοσομετρήσω την 1η κεφαλή
                      'κινησε τον 4ο στην 1η κεφαλή
                     If var65_1kef = Uesh_4axona() Then
                     Else
 
                             move_w var65_1kef
 
 
                          ' DUM = wait_to_understand(cmd$)
                          ' DUM = Wait_Robot(0)
                    End If
                 Else
                      'κινησε τον 4ο στην 2η κεφαλή
                        If var66_2kef = Uesh_4axona() Then
                        Else
                               move_w var66_2kef
                               'cmd$ = "!var12=91:"
                               
                               
                                'RobSend (cmd$)
                            'Ref.Caption = cmd$
                            'Do While Abs(var66_2kef - Uesh_4axona()) > 20
                            '    DoEvents
                            'Loop
                               
                               
                               
                               

                               
                               
                               
                               
                               
                               
                               
                               'DUM = wait_to_understand(cmd$)
                               'DUM = Wait_Robot(0)
                        End If
                 End If
                    result = Ektel_kinhsh(2)
                    result = Ektel_kinhsh(1) 'ΔΟΣΟΜΕΤΡΗΣΗ
                    
                           Last_Epil = prospau2.Recordset("epil")
                           prospau2.Recordset.MoveNext
                           'Grid.Refresh
                    
                    
                    
                    
                    DOSO_MPOYK = a_Fere(1)
                    dum = Wait_Robot(0)
           
           
           ElseIf a_Kinhsh(2) = 9 Then   ' ΔΟΣΟΜΕΤΡΗΣΗ ΧΡΩΜΑΤΟΣ ΣΤΟ (2)
                 
                 If a_Fere(2) = a_Kef1(2) Then  'ara θα δοσομετρήσω την 1η κεφαλή
                      'κινησε τον 4ο στην 1η κεφαλή
                     If var65_1kef = Uesh_4axona() Then
                     Else
                       move_w var65_1kef
                      'cmd$ = "!var12=90:"
                      'RobSend (cmd$)
                      'Ref.Caption = cmd$
                      'Do While Abs(var65_1kef - Uesh_4axona()) > 20
                      '      DoEvents
                      'Loop
                      
                      
                      
                     ' DUM = wait_to_understand(cmd$)
                      'DUM = Wait_Robot(0)
                    End If
                 Else
                      'κινησε τον 4ο στην 2η κεφαλή
                     If var66_2kef = Uesh_4axona() Then
                     Else
                                                    move_w var66_2kef
                         'cmd$ = "!var12=91:"
                         
                          '  RobSend (cmd$)
                           ' Ref.Caption = cmd$
                            'Do While Abs(var66_2kef - Uesh_4axona()) > 20
                             '   DoEvents
                            'Loop
                         
                         
                         'DUM = wait_to_understand(cmd$)
                         'DUM = Wait_Robot(0)
                      End If
                 End If
           
           
                   result = Ektel_kinhsh(1)
                   result = Ektel_kinhsh(2) 'ΔΟΣΟΜΕΤΡΗΣΗ
                   Last_Epil = prospau2.Recordset("epil")
                   prospau2.Recordset.MoveNext
                   'Grid.Refresh
                   DOSO_MPOYK = a_Fere(2)
                   dum = Wait_Robot(0)
           Else  ' ΔΕΝ ΕΧΩ ΔΟΣΟΜΕΤΡΗΣΗ ΧΡΩΜΑΤΟΣ
           
          
              If a_Kinhsh(1) = 8 Then 'ΔΟΣΟΜΕΤΡΗΣΗ ΝΕΡΟΥ
                  result = Ektel_kinhsh(2)
                  result = Ektel_kinhsh(1)
              Else
                  result = Ektel_kinhsh(1)
                  result = Ektel_kinhsh(2)
              End If
              dum = Wait_Robot(0)
           End If
       
       Else '---------- ΟΧΙ ΤΑΥΤΟΧΡΟΝΙΣΜΟΣ ------------------------------------------
              If a_Kinhsh(1) = 9 Then  'αν είναι ΔΟΣΟΜΕΤΡΗΣΗ
                 If a_Fere(1) = a_Kef1(1) Then  'ara θα δοσομετρήσω την 1η κεφαλή
                      'κινησε τον 4ο στην 1η κεφαλή
                    
                       If var65_1kef = Uesh_4axona() Then
                       Else
                          move_w var65_1kef
                         '  cmd$ = "!var12=90:"
                          '  RobSend (cmd$)
                           ' Ref.Caption = cmd$
                            'Do While Abs(var65_1kef - Uesh_4axona()) > 20
                             '   DoEvents
                          '  Loop
                           
                           
                           
                           
                           'DUM = wait_to_understand(cmd$)
                           'DUM = Wait_Robot(0)
                       End If
                 Else
                      'κινησε τον 4ο στην 2η κεφαλή
                        If var66_2kef = Uesh_4axona() Then
                        Else
                            move_w var66_2kef
                           'cmd$ = "!var12=91:"
                           
                            'RobSend (cmd$)
                          '  Ref.Caption = cmd$
                           ' Do While Abs(var66_2kef - Uesh_4axona()) > 20
                            '    DoEvents
                         '   Loop
                           
                           
                           'DUM = wait_to_understand(cmd$)
                           'DUM = Wait_Robot(0)
                        End If
                 End If
               End If
               
               
               
                result = Ektel_kinhsh(1)
                 DOSO_MPOYK = a_Fere(1)

               If a_Kinhsh(1) <> 9 Then
                    dum = Wait_Robot(0)
               Else
                    Last_Epil = prospau2.Recordset("epil")
                    prospau2.Recordset.MoveNext
                    'Grid.Refresh
               End If
       End If
       
       If result < -1 And result <> -7 Then     ' -7=απετυχε ο αλγόριθμος ζύγισης
            
          On Error Resume Next
              ' print_rep str(a_Syntagh(Move)), a_Ar_Syntag(Move)
               print_rep Pothri, mm_Synt
          On Error GoTo 0
            
            Ektelesueises(Pothri) = -1  ' φλαγκ που λέει ότι εκτελείται κανονικά η συνταγή
            Exit Do
       End If
       
       
       frmSYNT.Caption = ""
       
       Get_Gen_xronos "telos syntaghs"

       
       
       KINHSEIS.Recordset.MoveNext
       If KINHSEIS.Recordset.eof Then
             result = 0

             Exit Do
       Else
          
          If a_Syntagh(2) > a_Syntagh(1) Then  'ΑΝ ΠΗΓΕ ΣΤΗΝ ΕΠΟΜΕΝΗ ΣΥΝΤΑΓΗ

          End If
       End If
   
       Ektelesueises(Pothri) = 1  ' φλαγκ που λέει ότι εκτελείται κανονικά η συνταγή
   
   Loop
     
   

' βαζω status=2 για τις εκτελεσθείσες συνταγές
'  prospau2.Recordset.MoveLast
'  Do While Not prospau2.Recordset.BOF
'       m2_syntag = Trim(prospau2.Recordset("ENTOLH")) + "-" + LTrim(prospau2.Recordset("aa_PROSPAU"))
'       If prospau2.Recordset("userE") > 0 Then Updating = 1 Else Updating = 0
'       Do While Not prospau2.Recordset.BOF And m2_syntag = prospau2.Recordset("ENTOLH") + "-" + LTrim(prospau2.Recordset("aa_prospau"))
'         If Updating = 1 Then
'             prospau2.Recordset.Edit
'                prospau2.Recordset("status") = 2
'             prospau2.Recordset.update
'           End If
'          prospau2.Recordset.MovePrevious
'           If prospau2.Recordset.BOF Then Exit Do
'       Loop
'       If prospau2.Recordset.BOF Then Exit Do
'   Loop
   
     
     
     
' ΠΡΟΣΟΧΗ ΜΕΤΑ ΤΙΣ ΔΟΚΙΜΕΣ ΝΑ ΓΙΝΕΙ  "< 0 Then"
   
     Message = mL_Res(875) ' "ΟΛΕΣ ΟΙ ΣΥΝΤΑΓΕΣ ΟΛΟΚΛΗΡΩΘΗΚΑΝ ΜΕ ΕΠΙΤΥΧΙΑ"
   
   If result < -1 Then
    
    a_Syntagh(1) = 0
    a_Kinhsh(1) = KINHSEIS.Recordset("kinhsh")
    a_Kefalh(1) = KINHSEIS.Recordset("kefalh")
    a_Mgrs(1) = 0
    a_Pane(1) = 0
    a_Xronos(1) = 0
    a_Tkinhsh(1) = "Return of.."
    
    a_Kef1(1) = KINHSEIS.Recordset("kef1")
    a_Kef2(1) = KINHSEIS.Recordset("kef2")
    
    ' ΕΠΙΣΤΡΟΦΗ ΜΠΟΥΚΑΛΑΣ 1ης ΚΕΦΑΛΗΣ
      If a_Kef1(1) > 0 Then
        
                       'κινησε τον 4ο στην 2η κεφαλή
                       If var66_2kef = Uesh_4axona() Then
                       Else
                             move_w var66_2kef
                     '       cmd$ = "!var12=91:"
                      '      RobSend (cmd$)
                       '     Ref.Caption = cmd$
                        '    Do While Abs(var66_2kef - Uesh_4axona()) > 20
                         '       DoEvents
                          '  Loop
                           'DUM = wait_to_understand(cmd$)
                           'DUM = Wait_Robot(0)
                       End If
            
        
        
        a_Kinhsh(1) = 5
         a_Fere(1) = a_Kef1(1)
         dum = Ektel_kinhsh(1)
      End If
      
    ' ΕΠΙΣΤΡΟΦΗ ΜΠΟΥΚΑΛΑΣ 2ης ΚΕΦΑΛΗΣ
      If a_Kef2(1) > 0 Then
                      'κινησε τον 4ο στην 1η κεφαλή
                       If var65_1kef = Uesh_4axona() Then
                       Else
                            
                             move_w var65_1kef
                            'cmd$ = "!var12=90:"
                            
                            
                            'RobSend (cmd$)
                            'Ref.Caption = cmd$
                            'Do While Abs(var65_1kef - Uesh_4axona()) > 20
                            '    DoEvents
                            'Loop

                            
                            
                           'DUM = wait_to_understand(cmd$)
                           'DUM = Wait_Robot(0)
                       End If
       
         
         a_Kinhsh(1) = 5
         a_Fere(1) = a_Kef2(1)
          dum = Ektel_kinhsh(1)
          dum = Wait_Robot(0)
      End If
         
      ' επιστροφη ποτηριού
         a_Kinhsh(1) = 2
         a_Syntagh(1) = Pothri
         dum = Ektel_kinhsh(1)
         dum = Wait_Robot(0)
       
       
       
       Select Case result
            
          Case -31
             '
             Message = mLResnew(932, "Check the level of Chemicals or Calibrate them.*", 1)
          Case -20
             ' MsgBox "ΔΕΝ ΒΡΗΚΑ ΤΟ ΠΟΤΗΡΙ " + Str(POTHRI)
             Message = mL_Res(876) + str(Pothri) '"ΔΕΝ ΒΡΗΚΑ ΤΟ ΠΟΤΗΡΙ "
          
          Case -21
             '"ΔΙΑΚΟΠΗ ΣΥΝΤΑΓΩΝ ΣΤΟ ΠΟΤΗΡΙ  "
             Message = "877-878" + mL_CapRes(877) + str(Pothri) + mL_CapRes(878) '  " ΛΟΓΩ ΜΗ ΛΕΙΤΟΥΡΓΙΑΣ ΤΟΥ ΖΥΓΟΥ "
          Case -2
             Message = mL_Res(879) + str(DOSO_MPOYK) ' "η βαλβίδα δεν τρέχει στο μπουκάλι "
          Case -3
             Message = mL_Res(880) ' " η ζυγαρια υπερφορτώθηκε"
          Case -4
             Message = mL_Res(881) ' " δεν έγινε διαδικασία Εναρξης"
          Case -5
             Message = mL_Res(882) '"Η ζυγαριά δεν ανταποκρίνεται"
           Case -6
             Message = mL_Res(883) ' " Η ζυγαριά εκτός περιοχής"
           Case -7
             Message = mL_Res(884) ' " Η ζύγιση έχει αποτύχει"
             ' MsgBox
           Case -8
             Message = mL_Res(885) ' " Παρουσιάσθηκε απροσδόκητο λάθος στον αλγόρυθμο ζύγισης"
             'MsgBox
           Case -9
              Message = mL_Res(886) ' " Η ζύγιση ακυρώθηκε από τον χρήστη"
              '   MsgBox " Η ζύγιση ακυρώθηκε από τον χρήστη"
           Case -10
              Message = mL_Res(887) '" Η στάθμη του μπουκαλιού είναι πολύ χαμηλή <50 gr " + Str(DOSO_MPOYK)
              ' MsgBox
           Case -11
              Message = mL_Res(888) ' " Υπερβολική  ζητούμενη ποσότητα  (max. 100 gr)"
              '           MsgBox
        End Select
        
       
       Printer.Font = "Tahoma"
       Printer.FontSize = 12

       Printer.Print Message
       Printer.EndDoc
        
       On Error Resume Next
       
       Last_Epil = prospau2.Recordset("epil")
        
       prospau2.Recordset.MoveFirst
        Do While Not prospau2.Recordset.eof
            If Last_Epil = prospau2.Recordset("epil") Then
                prospau2.Recordset.EDIT
                  prospau2.Recordset("status") = 1
                prospau2.Recordset.update
            End If
            prospau2.Recordset.MoveNext
        Loop
        
       On Error GoTo 0
        
   End If
     
     
    dum = Wait_Robot(0)
     
If SYN_ALATIOY > 0 Then
         
       ' πανε τον  W στην θέση της σκόνης
         dum = wait_to_understand("!var12=93:")
         dum = Wait_Robot(0)
         
         
         
         ' φέρε το αλάτι στην κεφαλή της σκόνης
         dum = wait_to_understand("!var12=" + LTrim(str(2000 + UESH_ALATIOY)) + ":")
         dum = Wait_Robot(0)
     
     
     
End If
     
     kod6400 = 82
     cmd$ = "!var12=82:"
     
    dum = wait_to_understand(cmd$)
    dum = Wait_Robot(0)
     
     
     
         ' ΧΤΥΠΗΜΑ ΚΟΥΔΟΥΝΙΟΥ
 cmd$ = "!" + Alarm1_off + ":"
RobSend (cmd$)

cmd$ = "!" + Alarm1_on + ":"
RobSend (cmd$)


 Command1.Enabled = True

 MsgBox Message


'On Error GoTo 0
End Sub

Private Sub Command3_Click()

    a = MAKE_FILE()
    If a > 0 Then
        Exit Sub
    End If
End Sub

Private Sub Form_DblClick()
 Dim a
      'CommonDialog1.ShowColor
      'Me.BackColor = CommonDialog1.Color
      a = save_colors(Me, Me.BackColor)
End Sub

'select * from prospau2 where status=1 and epil>0  order by epil,seira
Private Sub Form_Load()
Dim k, a, aneb, L1, L6
Dim mydb As Database
Dim r As Recordset
Dim co As Integer
'Me.BackColor =
a = find_colors(Me)
If system_ready <> 0 Then
   ValveReset
End If
 Flag_Kefal = 1
 Flag_Show = -1    ' αδιάφορο
 flag_paint = 0
 Real_Alati = 0

      MDIForm1.Arxeia(10).Visible = False
      MDIForm1.Syntages(20).Visible = False
      MDIForm1.bohuhtika(30).Visible = False
      MDIForm1.ejodos(40).Visible = False





On Error Resume Next

 Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
 Set r = mydb.OpenRecordset("parametroi")
 r.MoveFirst
 r.MoveNext
 
 
 
 If IsNull(r("lsneg")) Then
    For k = 1 To 5:   f_poso_apor(co) = 100: Next
 Else
    co = 0
    For k = 1 To 21 Step 5
        co = co + 1
        f_poso_apor(co) = Val(Mid$(r("lsneg"), k, 5))
    Next
 End If
 
 mydb.Close

Ref.Caption = Format(f_poso_apor(1), "####") + "/" + Format(f_poso_apor(2), "####") + "/" + Format(f_poso_apor(3), "####") + "/" + Format(f_poso_apor(4), "####")



' ΑΝΟΙΓΩ ΣΕΙΡΙΑΚΗ ΠΟΡΤΑ ΚΑΙ ΜΗΔΕΝΙΖΩ ΖΥΓΑΡΙΑ
 MSComm1.commport = Balance_Port
 MSComm1.Settings = Balance_Settings
 MSComm1.InputLen = 0

MSComm1.PortOpen = False
MSComm1.PortOpen = True


'MSComm1.Output = Balance_2_Digits + Chr$(13)
 
  
' σχεδιάζω την νεκρή περιοχή  k_AR=0
  For k = 1 To 20
    Label1(K_AR + k).Visible = False
    Label1(K_AR + k + 20).Visible = False
    Label1(K_AR + k + 40).Visible = False
    If k <= 9 Then
      Label1(K_AR + k + 60).Visible = False
    End If
 Next
  aneb = 6285 - 5210
  
  For k = 1 To 30
    Label3(k).top = Label3(k).top - aneb   ' + IIf(k Mod 2 <> 0, 0, -Label3(k).height / 2)
  Next
  
'  For k = 8 To 14
'    Label3(k).top = Label3(k).top - aneb + IIf(k Mod 2 = 0, 0, -Label3(k).height / 2)
'  Next
  

  
  
  
  
  For k = 1 To Ar_Mpoykal Step 2
    Label1(k).top = Label1(k).top - Label1(k).height / 2
  Next
  '
  
  
 L1 = Label1(K_AR + 1).left
 L6 = Label1(K_AR + 6).left
  
  
  
    L1 = Label1(K_AR + 1).left     ' η θέση του πρώτου κρυμμένου κουτιού
'  L6 = Label1(K_AR + 6).left
  
  ' start.left = l1 + 630
  
'  Command1.left = L1 + 630
  
  
 Label2.left = 1440 + L1
'  Label2.left = 2040 + L1
  
'  Label3(0).left = 2760 + L1
  Label4(0).left = 1320 + L1
  
  Label5.left = 2760 + L1
  
  Label6.left = 720 + L1 ' GRIPPER
  
  Label11.left = (3240 - 1800) + L1 ' powder
  
  Label12.left = 1920 + L1
'======================================
aneb = 200
 Label2.top = Label2.top - aneb
 Label4(0).top = Label4(0).top - aneb
 Label5.top = Label5.top - aneb
 Label6.top = Label6.top - aneb
 Label11.top = Label11.top - aneb
 Label12.top = Label12.top - aneb
  Label5.left = 2760 + L1
  
  Label6.left = 720 + L1 ' GRIPPER
  
  Label11.left = (3240 - 1800) + L1 ' powder
  
  Label12.left = 1920 + L1
  
  
  
  For k = 1 To Ar_Mpoykal
    
        Label1(k).BackColor = &HFFFFFF     ' ασπρο          '  &HFF0000    μπλέ  φόντο      '
        'Label1(k).ForeColor = &HFF0000         ' μπλέ γραμματα
        OldColor(k) = Label1(k).BackColor
        Label1(k).Caption = k
        mTop(k) = Label1(k).top
        mLeft(k) = Label1(k).left
    
 Next
     
  For k = 1 To Ar_Pothr
     Label3(k).Caption = k
     Label3(k).BackColor = &HFFFF00 'SIEL      &HFFFFFF   ' ασπρο&H00FFFF00&
     pTop(k) = Label3(k).top
     pLeft(k) = Label3(k).left
     
 Next
 Label13.Caption = mL_CapRes(330) '   "Μικτό"
Label10.Caption = mL_CapRes(331) '   "Απόβ"
Label14.Caption = mL_CapRes(332) '   "Καθ"
Label19.Caption = mL_CapRes(333)  ' εκκίνηση απο ποτήρι
Label5.Caption = mL_CapRes(334)  'ζυγαρια
Label6.Caption = mL_CapRes(335)  'Gripper
'start.Caption = mL_CapRes(340) '    "ΕΚΚΙΝΗΣΗ"
Command1.Caption = mL_CapRes(341) '    "ΕΞΟΔΟΣ"




' σχεδιάζω την νεκρή περιοχή  k_AR=0
  For k = 1 To Ar_1_mpoykal - 1
      Label1(k).Visible = False
  Next
  For k = Ar_1_mpoykal To Ar_tel_mpoykal
      Label1(k).Visible = True
  Next
  
  For k = Ar_tel_mpoykal + 1 To Ar_Mpoykal
      Label1(k).Visible = False
  Next





End Sub





Private Sub Form_Paint()
Dim u, DIEYUYNSH(1 To 8), PROBLHMA, mmm_D, m_Counter, v, dbsjet
Dim k, hmeres
If flag_paint = 1 Then
   Exit Sub
Else
  flag_paint = 1 ' για να μην περνά 2 φορές
End If
   m_CC = 1
  frmSYNT.Visible = False
  
  PROBLHMA = 0
  For k = 1 To 8
     DIEYUYNSH(k) = "Ω" + str(k)
  Next
  
  
  Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\coord_tb.mdb")
  Set v = dbsjet.OpenRecordset("var")
  
  
  v.MoveFirst
  While Not v.eof
         If v("field_name") = "VAR65" Then
            var65_1kef = v("field_val")
         End If
         If v("field_name") = "VAR66" Then
            var66_2kef = v("field_val")
         End If
         If v("field_name") = "VAR67" Then
            var67_3kef = v("field_val")
         End If
         If v("field_name") = "VAR68" Then
            var68_4kef = v("field_val")
         End If
         
         If v("field_name") = "VAR93" Then
            var93_anoZ = v("field_val")
         End If
         If v("field_name") = "VAR74" Then
            var74 = v("field_val")
         End If
           If v("field_name") = "VAR88" Then
            var88 = v("field_val")
         End If
         
         
         v.MoveNext
  Wend
  
  v.Close
  
  
 ximitech.Recordset.Index = "KOD"
 ximitech.Recordset.MoveFirst
 While Not ximitech.Recordset.eof
         u = 0
         If Not IsNull(ximitech.Recordset("ADDR_PROT")) Then
                 u = ximitech.Recordset("addr_prot")
         End If
         
         If u > 0 And u <= Ar_Mpoykal Then    ' =============================================
                                  
                 If Not IsNull(ximitech.Recordset("LHJ_PROT")) Then   ' ---------------
                          hmeres = DateDiff("d", Now, ximitech.Recordset("LHJ_PROT"))
                 Else
                          hmeres = 1000
                 End If
                 
                If ximitech.Recordset("skonh") = 1 Then ' ΣΤΕΡΕΑ
                                          Label1(u).BackColor = &H808080           '   καφε
                Else  '    Δ Ι Α Λ Υ Μ Α Τ Α
                      Label1(u).BackColor = FIND_COLOR(hmeres)
                End If
                
                OldColor(u) = Label1(u).BackColor
                kodikos(u) = IIf(Not IsNull(ximitech.Recordset("kod")), ximitech.Recordset("kod"), "0000")
                Perigrafh(u) = IIf(Not IsNull(ximitech.Recordset("perigrafh")), ximitech.Recordset("perigrafh"), "  ")
                LHJH(u) = IIf(Not IsNull(ximitech.Recordset("Lhj_prot")), ximitech.Recordset("Lhj_prot"), "00/00/00")
                Ypoloipo(u) = IIf(Not IsNull(ximitech.Recordset("ypol_prot")), ximitech.Recordset("ypol_prot"), 0)
                 'αν το υπόλοιπο είναι κάτω των 50 γραμμαρίων να κοκκινίζει
                 If Ypoloipo(u) < 50 Then
                       Label1(u).BackColor = &HFF&        'κοκκινο
                 End If
                Ypoloipo(u) = IIf(Ypoloipo(u) < 0, 100, Ypoloipo(u))
                katofli(u) = IIf(Not IsNull(ximitech.Recordset("katofli")), ximitech.Recordset("katofli"), 0)
                'katofli(u) = IIf(katofli(u) > 15500, katofli(u), 15500)
                If Not IsNull(Perigrafh(u)) Then
                     Label1(u).Caption = Format(u, "###") + " " + Perigrafh(u)
                End If
                
            '  μικρεμα & μεγάλωμα μπουκαλιών ανάλογα την ποσότητα
               ' Label1(u).top = find_top(Label1(u).top, Label1(u).height, Ypoloipo(u))
               ' Label1(u).height = find_height(Label1(u).top, Label1(u).height, Ypoloipo(u))
         End If ' ================================================================
         ximitech.Recordset.MoveNext
 Wend


prospau2.Recordset.MoveLast

If prospau2.Recordset.RecordCount = 0 Then
   MsgBox mL_Res(889) ' "Δεν έχουν περάσει οι συνταγές από την ζύγιση πανιών "
   Com_file_exec.Enabled = False
End If

 prospau2.Recordset.MoveFirst
m_CC = 0
Do While Not prospau2.Recordset.eof
    If prospau2.Recordset("seira") = 1 Then
        m_CC = m_CC + 1
    End If
    prospau2.Recordset.MoveNext
Loop

 








        
'        Else
'            If GR(k) = 0 Then
'                GR(k) = -1
'            Else
        '        PROSPAU.Recordset.Edit
         '       PROSPAU.Recordset("ADD" + Right(Str(k), 1)) = XIMITECH.Recordset("ADDR_PROT")
          '      PROSPAU.Recordset("P" + Right(Str(k), 1)) = IIf(Not IsNull(XIMITECH.Recordset("perigrafh")), XIMITECH.Recordset("perigrafh"), "  ")
           '     PROSPAU.Recordset.update
            '    DIEYUYNSH(k) = PROSPAU.Recordset("ADD" + Right(Str(k), 1))
                 
                 
                 '  ελέγχω μην έχει πάρει τρελλή διεύθυνση
  '               mmm_D = Val(DIEYUYNSH(k))
   '                   If mmm_D < 1 Then
    '                       PROBLHMA = 1
     '                  ElseIf mmm_D > 908 Then
      '                     PROBLHMA = 1
       '                 ElseIf mmm_D < 906 And mmm_D > Ar_Mpoykal Then
 '                          PROBLHMA = 1
'                        ElseIf (mmm_D > K_AR And mmm_D < K_AR + 9) Or (mmm_D > K_AR + 20 And mmm_D < K_AR + 29) Or (mmm_D > K_AR + 40 And mmm_D < K_AR + 49) Or (mmm_D > K_AR + 60 And mmm_D < K_AR + 69) Or mmm_D Mod 20 = 0 Then
    '                       PROBLHMA = 1
   '                  End If
  '                   If PROBLHMA = 1 Then
 '                       Beep
                        
                       ' frmSYNT.caption = ml_CapRes(358) + DIEYUYNSH(k) + "   για την  συνταγή   " + PROSPAU.Recordset("entolh") + Str$(PROSPAU.Recordset("aa_prospau"))
                           '"ΑΝΥΠΑΡΚΤΗ ΘΕΣΗ  ΜΠΟΥΚΑΛΑΣ " + DIEYUYNSH(k) + "   για την  συνταγή   " + PROSPAU.Recordset("entolh") + Str$(PROSPAU.Recordset("aa_prospau"))
       '                 Command1.Enabled = True
      '                  start.Enabled = False
     '                   PROBLHMA = 0
    '                   End If
                     
                     
                     
   '         End If
  '      End If
        
        
 '    End If
'  Else
  '   GR(k) = -1
 ' End If
'Next

   ' ελέγχω τον πίνακα Κod_Χhmika για διπλοεγγραφές
'   For k = 1 To 8
 '      For kk = k + 1 To 8
  '         If DIEYUYNSH(k) = DIEYUYNSH(kk) Then
   '             fault = 1
    '            Beep
     '           frmSYNT.caption = ml_CapRes(359) + Str$(k) + mL_Res(360) + Str$(kk) + mL_Res(361) + PROSPAU.Recordset("entolh") + Str$(PROSPAU.Recordset("aa_prospau"))
'                Command1.Enabled = True
 '               start.Enabled = False
  '          End If
   '    Next
'   Next

 '  PROSPAU.Recordset.MoveNext


'Wend

'End If


'If PROSPAU.Recordset.RecordCount > 0 Then
 
 '  PROSPAU.Recordset.MoveFirst
'   m_Counter = 0
 '  Do While Not PROSPAU.Recordset.EOF
  '         m_Counter = m_Counter + 1
   '        pin_Syntag(m_Counter) = PROSPAU.Recordset("ENTOLH") + "-" + Str(PROSPAU.Recordset("aa_prospau")) + "  " + PROSPAU.Recordset("kod_pel")
           
    '       If Not IsNull(PROSPAU.Recordset("P1")) Then
     '         pin_Syntag(m_Counter) = pin_Syntag(m_Counter) + " ** " + PROSPAU.Recordset("P1") + "-"
      '     End If
       '    If Not IsNull(PROSPAU.Recordset("EK1")) Then
        '      pin_Syntag(m_Counter) = pin_Syntag(m_Counter) + Str(PROSPAU.Recordset("EK1")) + "---"
         '  End If
           
'           If Not IsNull(PROSPAU.Recordset("P2")) Then
 '             pin_Syntag(m_Counter) = pin_Syntag(m_Counter) + PROSPAU.Recordset("P2") + "-"
  '         End If
   '        If Not IsNull(PROSPAU.Recordset("EK2")) Then
    '          pin_Syntag(m_Counter) = pin_Syntag(m_Counter) + Str(PROSPAU.Recordset("EK2")) + "---"
     '      End If
           
'           If Not IsNull(PROSPAU.Recordset("P3")) Then
 '             pin_Syntag(m_Counter) = pin_Syntag(m_Counter) + PROSPAU.Recordset("P3") + "-"
  '         End If
   '        If Not IsNull(PROSPAU.Recordset("EK3")) Then
    '          pin_Syntag(m_Counter) = pin_Syntag(m_Counter) + Str(PROSPAU.Recordset("EK3")) + "---"
     '      End If
 '          PROSPAU.Recordset.MoveNext
  ' Loop
   
   
 '  PROSPAU.Recordset.MoveFirst

'End If

 ' δίνω το αρχικό χρώμα στα ποτήρια
   For k = 1 To Ar_Pothr
         Label3(k).BackColor = ASPRO    ' &HFFFFFF   ' ασπρο
   Next
   
   For k = m_Counter + 1 To Ar_Pothr
       Label3(k).BackColor = BLE      '   &HFF0000       ' mple
       pin_Syntag(k) = "   "
   Next
   
   For k = 1 To m_CC
       Label3(k).BackColor = ROZ '   &HFF0000       ' mple
       pin_Syntag(k) = "   "
   Next

   
   Screen.MousePointer = 1

    frmSYNT.Visible = True

'End If

End Sub
Function find_top(Top0, H0, ypol)
   Select Case ypol
      Case Is > 600
         find_top = Top0
      Case Is > 400
         find_top = Top0 + H0 / 5
      Case Is > 200
         find_top = Top0 + H0 / 3
      Case Is > 0
         find_top = Top0 + 3 * H0 / 5
      End Select
End Function
Function find_height(ByVal Top0 As Integer, ByVal H0 As Integer, ByVal ypol As Single) As Integer
   Select Case ypol
      Case Is > 600
         find_height = H0
      Case Is > 400
         find_height = H0 - H0 / 5
      Case Is > 200
         find_height = H0 - H0 / 3
      Case Is > 0
         find_height = H0 - 3 * H0 / 5
   End Select
End Function



Function Zyg_TRIPLO(Kauysterhsh As Single, akrivia) As Single
Dim Z1, z2, z3, mo, metr
'not in use 31-8-2001


metr = 0
 Do While True
 
 
   If metr = 1 Then
       'For k = 1 To kauysterhsh: DoEvents: Next
       MilSec Kauysterhsh / 13.7
   End If
 
   Z1 = Zyg_KAUARO(0)
   If Z1 > TREXEI And z2 > TREXEI Then
        mo = TREXEI
         Exit Do
     End If
   
   If metr = 1 Then
        mo = (Z1 + z2 + z3) / 3
        If Abs(Z1 - mo) < akrivia And Abs(z2 - mo) < akrivia And Abs(z3 - mo) < akrivia Then
           Exit Do
        End If
   End If
     
     
     MilSec Kauysterhsh / 13.7
     'For k = 1 To kauysterhsh: DoEvents: Next
     z2 = Zyg_KAUARO(0)
     If z2 > TREXEI And z3 > TREXEI Then
         mo = TREXEI
         Exit Do
     End If
   
   If metr = 1 Then
      mo = (Z1 + z2 + z3) / 3
      If Abs(Z1 - mo) < akrivia And Abs(z2 - mo) < akrivia And Abs(z3 - mo) < akrivia Then
          Exit Do
      End If
   End If
     
   'For k = 1 To kauysterhsh: DoEvents: Next
   MilSec Kauysterhsh / 13.7
   z3 = Zyg_KAUARO(0)
   If z3 > TREXEI And Z1 > TREXEI Then
         mo = TREXEI
         Exit Do
     End If
   mo = (Z1 + z2 + z3) / 3
   If Abs(Z1 - mo) < akrivia And Abs(z2 - mo) < akrivia And Abs(z3 - mo) < akrivia Then
         Exit Do
   End If
   If mo - EPIU > 17000 Then ' υπερχείλιση
       Exit Do
   End If
   metr = 1
 Loop
Zyg_TRIPLO = Int(mo + 0.5)

If Zyg_TRIPLO < 0 Then
     APOBARO = APOBARO + Zyg_TRIPLO
     Zyg_TRIPLO = 0
End If
   
   If mo - EPIU > 17000 Then ' υπερχείλιση
       Zyg_TRIPLO = TREXEI
   End If

End Function

Function zygisabs(Kauysterhsh As Single, akrivia) As Single
Dim Z1, z2, z3, mo
'not in use 31-8-2001
 Do While True
   Z1 = Zygis0(0)
     MilSec Kauysterhsh / 13.7
     'For k = 1 To kauysterhsh: DoEvents: Next
   z2 = Zygis0(0)
     ' For k = 1 To kauysterhsh: DoEvents: Next
     MilSec Kauysterhsh / 13.7
   z3 = Zygis0(0)
   mo = (Z1 + z2 + z3) / 3
   If Abs(Z1 - mo) < akrivia And Abs(z2 - mo) < akrivia And Abs(z3 - mo) < akrivia Then
         Exit Do
   End If
 Loop
zygisabs = Int(mo + 0.5)

End Function




Private Sub Form_Unload(Cancel As Integer)
     MSComm1.PortOpen = False
     ValveReset
    ' Reset
     'XIMITECH.Recordset.Close
     'PROSPAU.Recordset.Close
     'Var.Recordset.Close
End Sub

Private Sub Label1_Click(Index As Integer)
'   fernei (index)
'                                 ασπρο                                                     γκρίζο
Dim k

'If Label1(Index).BackColor = &HFFFFFF Then   'ασπρο δεν κινείται
     '
'ElseIf Label1(Index).BackColor = &H80C0FF Then ' αν είναι  γκρίζο να  γινει όπως ήταν
 '  Label1(Index).BackColor = OldColor(Index)
   ' το βγάζει από την θέση μπροστά
 '  For k = 2 To 20
  '    If Label3(k).Caption = Index Then            '  ξαναβάζω τον αρχικό του αριθμό
   '        Label3(k).BackColor = &HFFFFFF     ' ασπρο
    '       Label3(k).Caption = k                        '  ξαναβάζω τον αρχικό του αριθμό
     '     Exit For
'      End If
 '  Next
'Else
   ' ψάχνει θέση για τοποθέτηση μπροστά
 '  For k = 2 To 20
  '    If Label3(k).BackColor <> &HFF& Then ' αν δεν είναι κόκκινη θέση φέρτο εδώ
   '       Label3(k).Caption = Index           'παίρνει τον αριθμό του μπουκαλιού που θάρθει
    '      Label3(k).BackColor = &HFF&    'κοκκινίζω την μπροστά θέση

     '     Label1(Index).BackColor = &H80C0FF          'το αλλάζω χρωμα  για να μην ξαναζητηθεί
      '    Exit For
    '  End If
 '  Next
'End If

End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
If Not IsNull(Perigrafh(Index)) Then
     SXOLIA1.Caption = Perigrafh(Index)
 End If
If Not IsNull(Format(Ypoloipo(Index), "###") + "              " + Format(LHJH(Index), "dd/mm/yy")) Then
     SXOLIA2.Caption = Format(Ypoloipo(Index), "###") + "              " + Format(LHJH(Index), "dd/mm/yy")
End If
End Sub

'
'




Private Sub Label3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
If Index > 0 And Index <= Ar_Pothr Then
     Label18.Caption = IIf(IsNull(pin_Syntag(Index)), "", pin_Syntag(Index))
End If
End Sub


   
   
   
   
   
   
   
   
   
   



Private Sub Text1_LostFocus()
Dim k
If Not IsNumeric(Text1) Then Text1 = "1"
If Val(Text1) > Ar_Pothr Then Text1 = "1"
   
   
   For k = 1 To Ar_Pothr
       Label3(k).BackColor = BLE        '   &HFF0000       ' mple
       pin_Syntag(k) = "   "
   Next
  
  
  For k = Val(Text1) To m_CC + Val(Text1) - 1
       Label3(k).BackColor = ROZ        '   &HFF0000       ' mple
       pin_Syntag(k) = "   "
   Next
End Sub

Private Sub Timer1_Timer()
 
  Dim mkod, grip_top, grip_left, a, mtimer, Arist_Kef_Left, Dej_Kef_Left
  Dim kef_top, zyg_top
  
  
  
   If Com_file_exec.Enabled = False And system_ready <> 0 Then
        Elegxo_Katastash
  End If
 Exit Sub
     
   a = 0
             ' gripper object
   grip_top = Label6.top - Label1(181 + K_AR).height '  4950
   grip_left = Label6.left '  3600
                 
                 ' chemicals
   kef_top = Label12.top  ' 4680  ok
                  
                  ' ζυγός
   Dej_Kef_Left = Label12.left + Label12.width + Label12.width + 100  ' 5640  ok
   
                   ' ζυγος           Chemicals
   Arist_Kef_Left = Label12.left + Label12.width + 50       '5160
             
             ' zygos
   zyg_top = Label5.top



     '           kod6400 = gia_epistrofh + ADDR(k) * AR_MPOYKAL + 100000
     ' ===== επέστρεψε μπουκάλι  & φέρε μπουκάλι   =============
     If kod6400 > 100000 Then
           ' mkod = kod6400 - 100000
           ' Feygei = mkod Mod AR_MPOYKAL
           ' Erxetai = Int(mkod / AR_MPOYKAL)
            
            If out_bit(grip_c) = 1 And out_bit(grip_o) = 0 And Flag_Show = -1 Then    'κλειστός gripper
                    Flag_Show = 1  ' flag ότι έκλεισε ο gripper για πρώτη φορά
                        mkod = Feygei
                        Label1(mkod).Visible = False
                    Label1(mkod).top = grip_top  ' εξαφανίζει το κουτάκι του μπουκαλιού
                    Label1(mkod).left = grip_left
                        Label1(mkod).Visible = True
            End If
            
            If out_bit(grip_o) = 1 And out_bit(grip_c) = 0 And Flag_Show = 1 Then    'ανοικτός gripper
                    mkod = Feygei
                    Label1(mkod).Visible = False
                    Label1(mkod).top = mTop(mkod)
                    Label1(mkod).left = mLeft(mkod)
                    Label1(mkod).Visible = True
                    Flag_Show = 2
            End If
            
            If out_bit(grip_c) = 1 And out_bit(grip_o) = 0 And Flag_Show = 2 Then    'κλειστός gripper
                    Flag_Show = 3  ' flag ότι έκλεισε ο gripper για  2η φορά
                        mkod = Erxetai
                        Label1(mkod).Visible = False
                    Label1(mkod).top = grip_top  ' εξαφανίζει το κουτάκι του μπουκαλιού
                    Label1(mkod).left = grip_left
                        Label1(mkod).Visible = True
            End If
            
            If out_bit(grip_o) = 1 And out_bit(grip_c) = 0 And Flag_Show = 3 Then    'ανοικτός gripper
                   
                   mkod = Erxetai
                   Label1(mkod).Visible = False
                  Label1(mkod).top = kef_top
                  If Flag_Kefal = 1 Then  ' δεξιά κεφαλή
                       Label1(mkod).left = Dej_Kef_Left
                       Flag_Kefal = 2
                 Else  ' αριστερή κεφαλή
                       Label1(mkod).left = Arist_Kef_Left
                       Flag_Kefal = 1
                 End If
                 Flag_Show = -1
                  Label1(mkod).Visible = True
                    
            End If
     
     End If
     
     
     
     
     '  ========  φερε ποτήρι =============
     If kod6400 > 3000 And kod6400 < 3999 Then
            mkod = kod6400 - 3000
            
            If out_bit(grip_c) = 1 And out_bit(grip_o) = 0 And Flag_Show < 0 Then     'κλειστός gripper
                    Flag_Show = 1  ' flag ότι έκλεισε ο gripper για πρώτη φορά
                    Label3(mkod).Visible = False
                    Label3(mkod).top = grip_top  ' εξαφανίζει το κουτάκι του ποτηριού
                    Label3(mkod).left = grip_left
                    Label3(mkod).Visible = True
            End If
            
            If out_bit(grip_o) = 1 And out_bit(grip_c) = 0 And Flag_Show = 1 Then    'ανοικτός gripper
                    Label3(mkod).Visible = False
                    Label3(mkod).left = Dej_Kef_Left
                    Label3(mkod).top = zyg_top - Label3(mkod).height
                    Flag_Show = -1
                     Label3(mkod).Visible = True
            End If
     End If
     
     ' ===== επέστρεψε ποτήρι =============
     If kod6400 > 4000 And kod6400 < 4999 Then
            mkod = kod6400 - 4000
            
            If out_bit(grip_c) = 1 And out_bit(grip_o) = 0 And Flag_Show < 0 Then     'κλειστός gripper
                    Flag_Show = 1  ' flag ότι έκλεισε ο gripper για πρώτη φορά
                     Label3(mkod).Visible = False
                    Label3(mkod).top = grip_top  ' εξαφανίζει το κουτάκι του ποτηριού
                    Label3(mkod).left = grip_left
                     Label3(mkod).Visible = True
            End If
            
            If out_bit(grip_o) = 1 And out_bit(grip_c) = 0 And Flag_Show = 1 Then    'ανοικτός gripper
                    Label3(mkod).Visible = False
                    Label3(mkod).top = pTop(mkod)
                    Label3(mkod).left = pLeft(mkod)
                    Flag_Show = -1
                     Label3(mkod).Visible = True
            End If
     End If
     
     '  ========  φερε μπουκάλι  =============
     If kod6400 > 1000 And kod6400 < 1999 Then
            mkod = kod6400 - 1000
            
            ' τσιμπάει το μπουκάλι
            If out_bit(grip_c) = 1 And out_bit(grip_o) = 0 And Flag_Show = -1 Then    'κλειστός gripper
                    Flag_Show = 1  ' flag ότι έκλεισε ο gripper για πρώτη φορά
                    Label1(mkod).Visible = False
                    Label1(mkod).top = grip_top  ' εξαφανίζει το κουτάκι του μπουκαλιού
                    Label1(mkod).left = grip_left
                    Label1(mkod).Visible = True
            End If
            ' ξεφορτώνει το μπουκάλι
            If out_bit(grip_o) = 1 And out_bit(grip_c) = 0 And Flag_Show = 1 Then    'ανοικτός gripper
                   Label1(mkod).Visible = False
                  Label1(mkod).top = kef_top
                  If Flag_Kefal = 1 Then  ' δεξιά κεφαλή
                       Label1(mkod).left = Dej_Kef_Left
                       Flag_Kefal = 2
                 Else  ' αριστερή κεφαλή
                       Label1(mkod).left = Arist_Kef_Left
                       Flag_Kefal = 1
                 End If
                 Flag_Show = -1
                  Label1(mkod).Visible = True
            End If
     End If
     
     ' ===== επέστρεψε μπουκάλι  =============
     If kod6400 > 2000 And kod6400 < 2999 Then
            mkod = kod6400 - 2000
            
            If out_bit(grip_c) = 1 And out_bit(grip_o) = 0 And Flag_Show = -1 Then    'κλειστός gripper
                    Flag_Show = 1  ' flag ότι έκλεισε ο gripper για πρώτη φορά
                    Label1(mkod).Visible = False
                    Label1(mkod).top = grip_top  ' εξαφανίζει το κουτάκι του μπουκαλιού
                    Label1(mkod).left = grip_left
                    Label1(mkod).Visible = True
            End If
            
            If out_bit(grip_o) = 1 And out_bit(grip_c) = 0 And Flag_Show = 1 Then    'ανοικτός gripper
                    Label1(mkod).Visible = False
                    Label1(mkod).top = mTop(mkod)
                    Label1(mkod).left = mLeft(mkod)
                    Label1(mkod).Visible = True
                    Flag_Show = -1
            End If
     End If
     '2H KEFALH
       If var66_2kef = Uesh_4axona() Then
         '        For I = 1 To Form1.Controls.Count - 1   ola ta control
      
            'Label5.Visible = False 'ZYGARIA
            'Label3(POTHRI).Visible = False  'POTHRI POY DOSOMETREITAI
            
            Label5.left = Arist_Kef_Left      '    L1 + 1920
            Label3(Pothri).left = Arist_Kef_Left   ' (L1 + 1920) + 120
            
            'Label5.Visible = True
            'Label3(POTHRI).Visible = True
       End If
     ' 1 KEFALH
      If var65_1kef = Uesh_4axona() Then
           ' Label5.Visible = False
           ' Label3(POTHRI).Visible = False
             
                       
            Label5.left = Dej_Kef_Left      '    L1 + 2400
            Label3(Pothri).left = Dej_Kef_Left  ' (L1 + 2400) + 120
            
           ' Label5.Visible = True
           ' Label3(POTHRI).Visible = True
     End If
     
     
     
     
End Sub




Public Sub Get_Gen_xronos(stringato)
Exit Sub
Dim mydb, rRd
Set mydb = OpenDatabase("C:\Talos\katoflia.mdb")
Set rRd = mydb.OpenRecordset("dosom")
rRd.AddNew
rRd("string") = stringato
rRd("mpoykali") = Int((GetCurrentTime() - General_Xronos) / 1000)
rRd("hme") = Now
rRd.update

End Sub

Sub G_REZERO(porta)

Dim k As Integer


    Dim Zero

    Dim ZZ As Long

    On Error Resume Next

    If system_ready = 0 Then Exit Sub
 
    If Balance_Type = "ADAM" Then  ' OVER =zyg_ADAM = -7900000  '
    
        ZZ = zyg_ADAM(0)

        If Abs(ZZ) = 9900000 Then  ' uNDER
            porta.Output = "ON" + Chr(13)  'KLEINO KAI ANOOIGV
            MilSec 4000  '4 O
            '
            porta.Output = "ON" + Chr(13)
            MilSec 4000  '4 O
                
        Else
            If Abs(ZZ) = 999000999 Then ' ΣΒΗΣΤΗ Ή ΧΑΛΑΣΜΕΝΗ
                For k = 1 To 4
                    porta.Output = "ON" + Chr(13)  'ΤΗΝ ΑΝΑΒΩ
                    MilSec 4000  '4 OK
                    porta.Output = "T" + Chr(13)
                    MilSec 2000

                    If Abs(zyg_ADAM(0)) < 2 Then

                        ' OK
                        Exit For

                    Else
                        ' ΗΤΑΝ ΣΒΗΣΤΗ ΚΑΙ ΤΗΝ ΑΝΟΙΓΕΙ
                        porta.Output = "ON" + Chr(13)
                        MilSec 4000  '4 OK
                        porta.Output = "T" + Chr(13)
                    End If

                Next
       
        
        
      
            Else  '     Then ZZ > 0 And ZZ < 0 Then    ' ΔΕΙΧΝΕΙ ΤΙΜΗ ΑΛΛΑ ΘΕΤΙΚΗ Ή ΑΡΜΗΤΙΚΗ  ΟΧΙ UNDER KAI OVER
                'Else
                'Zero = "Z" + Chr$(13) + Chr(10)
                'porta.Output = Zero
                
                porta.Output = "T" + Chr(13)
                   MilSec 4000
                For k = 1 To 10
                
                    If Abs(zyg_ADAM(0)) < 2 Then
                        ' OK
                        Exit For
                    End If
                    porta.Output = "T" + Chr(13)
                    If k > 5 Then
                       MilSec 6000
                    Else
                        MilSec 4000
                    End If
                    
                    
                Next
         
            '     Else ' ΣΒΗΣΤΗ
       
                
            End If
       
        End If



   End If '  Balance_Type = "ADAM" Then

        'secwait (1)
        'MSComm1.Output = Zero
        'secwait (1)'

        'secwait (1)
        'asw1 = ZYGIS3()
        'secwait (1)
        'asw2 = ZYGIS3()
        'If Abs(asw1) + Abs(asw2) > 10 Then
        '   MilSec 1000
        '   GoTo 17
        'End If
  
    End Sub


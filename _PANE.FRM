VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form PANE 
   BackColor       =   &H00008000&
   ClientHeight    =   7404
   ClientLeft      =   2040
   ClientTop       =   516
   ClientWidth     =   9660
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7404
   ScaleWidth      =   9660
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1695
      TabIndex        =   289
      Top             =   6315
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Χρήση Barcode"
      Height          =   255
      Left            =   105
      TabIndex        =   288
      Top             =   6345
      Width           =   1515
   End
   Begin VB.Data XIMITECH 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\recipies.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   945
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "XIMITECH"
      Top             =   7155
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ΕΞΟΔΟΣ"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   5235
      Width           =   1335
   End
   Begin VB.Data mpoyk 
      Caption         =   "mpoyk"
      Connect         =   "dBASE III;"
      DatabaseName    =   "c:\talos"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   330
      Left            =   210
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "MPOYK"
      Top             =   8025
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3135
      Top             =   8055
   End
   Begin VB.Data SQLXIMITECH 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "c:\talos\RECIPIES.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1830
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from ximitech where  (addr_prot=0) or (addr_prot is null) order by mperigrafh"
      Top             =   8055
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton back 
      Caption         =   "πήγαινε πίσω"
      Height          =   315
      Left            =   105
      TabIndex        =   2
      Top             =   5235
      Width           =   3315
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "_PANE.frx":0000
      DataSource      =   "XIMITECH"
      Height          =   315
      Left            =   105
      TabIndex        =   3
      Top             =   5895
      Width           =   3315
      _ExtentX        =   5842
      _ExtentY        =   508
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      ListField       =   "MPERIGRAFH"
      BoundColumn     =   "MPERIGRAFH"
      Text            =   "DBCombo1"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   -225
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "299"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   299
      Left            =   8505
      TabIndex        =   341
      Top             =   -15
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "298"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   298
      Left            =   8040
      TabIndex        =   340
      Top             =   -15
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "297"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   297
      Left            =   7560
      TabIndex        =   339
      Top             =   -15
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "276"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   296
      Left            =   7095
      TabIndex        =   338
      Top             =   -15
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "295"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   295
      Left            =   6630
      TabIndex        =   337
      Top             =   -15
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "294"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   294
      Left            =   6150
      TabIndex        =   336
      Top             =   -15
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "293"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   293
      Left            =   5685
      TabIndex        =   335
      Top             =   -15
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "292"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   292
      Left            =   5220
      TabIndex        =   334
      Top             =   -15
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "290"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   290
      Left            =   4275
      TabIndex        =   333
      Top             =   -15
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "289"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   289
      Left            =   3810
      TabIndex        =   332
      Top             =   -15
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "288"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   288
      Left            =   3330
      TabIndex        =   331
      Top             =   -15
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "287"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   287
      Left            =   2865
      TabIndex        =   330
      Top             =   -15
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "286"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   286
      Left            =   2400
      TabIndex        =   329
      Top             =   -15
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "285"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   285
      Left            =   1920
      TabIndex        =   328
      Top             =   -15
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "284"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   284
      Left            =   1455
      TabIndex        =   327
      Top             =   -15
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "282"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   282
      Left            =   510
      TabIndex        =   326
      Top             =   -15
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "281"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   281
      Left            =   45
      TabIndex        =   325
      Top             =   -15
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "300"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   300
      Left            =   8970
      TabIndex        =   324
      Top             =   -15
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "291"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   291
      Left            =   4740
      TabIndex        =   323
      Top             =   -15
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "283"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   283
      Left            =   990
      TabIndex        =   322
      Top             =   -15
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Label6"
      Enabled         =   0   'False
      Height          =   180
      Index           =   0
      Left            =   3780
      TabIndex        =   321
      Top             =   6675
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "30"
      Height          =   375
      Index           =   30
      Left            =   8970
      TabIndex        =   320
      Top             =   6525
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "29"
      Height          =   375
      Index           =   29
      Left            =   8505
      TabIndex        =   319
      Top             =   6525
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "28"
      Height          =   375
      Index           =   28
      Left            =   8025
      TabIndex        =   318
      Top             =   6525
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "20"
      Height          =   375
      Index           =   20
      Left            =   8970
      TabIndex        =   317
      Top             =   6975
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "19"
      Height          =   375
      Index           =   19
      Left            =   8505
      TabIndex        =   316
      Top             =   6975
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "18"
      Height          =   375
      Index           =   18
      Left            =   8025
      TabIndex        =   315
      Top             =   6975
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "10"
      Height          =   375
      Index           =   10
      Left            =   8970
      TabIndex        =   314
      Top             =   7425
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "9"
      Height          =   375
      Index           =   9
      Left            =   8505
      TabIndex        =   313
      Top             =   7425
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "8"
      Height          =   375
      Index           =   8
      Left            =   8025
      TabIndex        =   312
      Top             =   7425
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "13"
      Height          =   375
      Index           =   13
      Left            =   5685
      TabIndex        =   311
      Top             =   6975
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "1"
      Height          =   375
      Index           =   1
      Left            =   4740
      TabIndex        =   310
      Top             =   7425
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "2"
      Height          =   375
      Index           =   2
      Left            =   5205
      TabIndex        =   309
      Top             =   7425
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "3"
      Height          =   375
      Index           =   3
      Left            =   5685
      TabIndex        =   308
      Top             =   7425
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "4"
      Height          =   375
      Index           =   4
      Left            =   6150
      TabIndex        =   307
      Top             =   7425
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "5"
      Height          =   375
      Index           =   5
      Left            =   6615
      TabIndex        =   306
      Top             =   7425
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "6"
      Height          =   375
      Index           =   6
      Left            =   7095
      TabIndex        =   305
      Top             =   7425
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "7"
      Height          =   375
      Index           =   7
      Left            =   7560
      TabIndex        =   304
      Top             =   7425
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "11"
      Height          =   375
      Index           =   11
      Left            =   4740
      TabIndex        =   303
      Top             =   6975
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "12"
      Height          =   375
      Index           =   12
      Left            =   5205
      TabIndex        =   302
      Top             =   6975
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "14"
      Height          =   375
      Index           =   14
      Left            =   6150
      TabIndex        =   301
      Top             =   6975
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "15"
      Height          =   375
      Index           =   15
      Left            =   6615
      TabIndex        =   300
      Top             =   6975
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "16"
      Height          =   375
      Index           =   16
      Left            =   7095
      TabIndex        =   299
      Top             =   6975
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "17"
      Height          =   375
      Index           =   17
      Left            =   7560
      TabIndex        =   298
      Top             =   6975
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "21"
      Height          =   375
      Index           =   21
      Left            =   4740
      TabIndex        =   297
      Top             =   6525
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "22"
      Height          =   375
      Index           =   22
      Left            =   5205
      TabIndex        =   296
      Top             =   6525
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "23"
      Height          =   375
      Index           =   23
      Left            =   5685
      TabIndex        =   295
      Top             =   6525
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "24"
      Height          =   375
      Index           =   24
      Left            =   6150
      TabIndex        =   294
      Top             =   6525
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "25"
      Height          =   375
      Index           =   25
      Left            =   6615
      TabIndex        =   293
      Top             =   6525
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "26"
      Height          =   375
      Index           =   26
      Left            =   7095
      TabIndex        =   292
      Top             =   6525
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "27"
      Height          =   375
      Index           =   27
      Left            =   7560
      TabIndex        =   291
      Top             =   6525
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      Height          =   315
      Left            =   3480
      TabIndex        =   290
      Top             =   6285
      Visible         =   0   'False
      Width           =   3030
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Ονομα Χημικοτεχνικού"
      Height          =   240
      Left            =   105
      TabIndex        =   5
      Top             =   5625
      Width           =   3315
   End
   Begin VB.Label SXOLIA2 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   1660
      TabIndex        =   287
      Top             =   6840
      Width           =   3885
   End
   Begin VB.Label SXOLIA1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1660
      TabIndex        =   286
      Top             =   6360
      Width           =   3165
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "263"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   263
      Left            =   990
      TabIndex        =   285
      Top             =   405
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "271"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   271
      Left            =   4740
      TabIndex        =   284
      Top             =   405
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "280"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   280
      Left            =   8970
      TabIndex        =   283
      Top             =   405
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "261"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   261
      Left            =   45
      TabIndex        =   282
      Top             =   405
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "262"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   262
      Left            =   510
      TabIndex        =   281
      Top             =   405
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "264"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   264
      Left            =   1455
      TabIndex        =   280
      Top             =   405
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "265"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   265
      Left            =   1920
      TabIndex        =   279
      Top             =   405
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "266"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   266
      Left            =   2400
      TabIndex        =   278
      Top             =   405
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "267"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   267
      Left            =   2865
      TabIndex        =   277
      Top             =   405
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "268"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   268
      Left            =   3330
      TabIndex        =   276
      Top             =   405
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "269"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   269
      Left            =   3810
      TabIndex        =   275
      Top             =   405
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "270"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   270
      Left            =   4275
      TabIndex        =   274
      Top             =   405
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "272"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   272
      Left            =   5220
      TabIndex        =   273
      Top             =   405
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "273"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   273
      Left            =   5685
      TabIndex        =   272
      Top             =   405
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "274"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   274
      Left            =   6150
      TabIndex        =   271
      Top             =   405
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "275"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   275
      Left            =   6630
      TabIndex        =   270
      Top             =   405
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "276"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   276
      Left            =   7095
      TabIndex        =   269
      Top             =   405
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "277"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   277
      Left            =   7560
      TabIndex        =   268
      Top             =   405
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "278"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   278
      Left            =   8040
      TabIndex        =   267
      Top             =   405
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "279"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   279
      Left            =   8505
      TabIndex        =   266
      Top             =   405
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "20"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   20
      Left            =   8970
      TabIndex        =   265
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "40"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   40
      Left            =   8970
      TabIndex        =   264
      Top             =   5685
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "60"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   60
      Left            =   8970
      TabIndex        =   263
      Top             =   5235
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "80"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   80
      Left            =   8970
      TabIndex        =   262
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "100"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   100
      Left            =   8970
      TabIndex        =   261
      Top             =   4365
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "120"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   120
      Left            =   8970
      TabIndex        =   260
      Top             =   3915
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "140"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   140
      Left            =   8970
      TabIndex        =   259
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "160"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   160
      Left            =   8970
      TabIndex        =   258
      Top             =   3045
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "180"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   180
      Left            =   8970
      TabIndex        =   257
      Top             =   2595
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "200"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   200
      Left            =   8970
      TabIndex        =   256
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "220"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   220
      Left            =   8970
      TabIndex        =   255
      Top             =   1725
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "240"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   240
      Left            =   8970
      TabIndex        =   254
      Top             =   1275
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "260"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   260
      Left            =   8970
      TabIndex        =   253
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "252"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   72
      Left            =   5220
      TabIndex        =   252
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "251"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   71
      Left            =   4740
      TabIndex        =   251
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   70
      Left            =   4275
      TabIndex        =   250
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "249"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   69
      Left            =   3810
      TabIndex        =   249
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "248"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   68
      Left            =   3330
      TabIndex        =   248
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "247"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   67
      Left            =   2865
      TabIndex        =   247
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   66
      Left            =   2400
      TabIndex        =   246
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   65
      Left            =   1920
      TabIndex        =   245
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "252"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   52
      Left            =   5220
      TabIndex        =   244
      Top             =   5235
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "251"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   51
      Left            =   4740
      TabIndex        =   243
      Top             =   5235
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   50
      Left            =   4275
      TabIndex        =   242
      Top             =   5235
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "249"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   49
      Left            =   3810
      TabIndex        =   241
      Top             =   5235
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "248"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   48
      Left            =   3330
      TabIndex        =   240
      Top             =   5235
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "247"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   47
      Left            =   2865
      TabIndex        =   239
      Top             =   5235
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   46
      Left            =   2400
      TabIndex        =   238
      Top             =   5235
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   45
      Left            =   1920
      TabIndex        =   237
      Top             =   5235
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "12"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   12
      Left            =   5220
      TabIndex        =   236
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "11"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   11
      Left            =   4740
      TabIndex        =   235
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "10"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   10
      Left            =   4275
      TabIndex        =   234
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "9"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   9
      Left            =   3810
      TabIndex        =   233
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "8"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   8
      Left            =   3330
      TabIndex        =   232
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "7"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   7
      Left            =   2865
      TabIndex        =   231
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "6"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   6
      Left            =   2400
      TabIndex        =   230
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "5"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   5
      Left            =   1920
      TabIndex        =   229
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "252"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   32
      Left            =   5220
      TabIndex        =   228
      Top             =   5685
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "251"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   31
      Left            =   4740
      TabIndex        =   227
      Top             =   5685
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   30
      Left            =   4275
      TabIndex        =   226
      Top             =   5685
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "249"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   29
      Left            =   3810
      TabIndex        =   225
      Top             =   5685
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "248"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   28
      Left            =   3330
      TabIndex        =   224
      Top             =   5685
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "247"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   27
      Left            =   2865
      TabIndex        =   223
      Top             =   5685
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   26
      Left            =   2400
      TabIndex        =   222
      Top             =   5685
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   25
      Left            =   1920
      TabIndex        =   221
      Top             =   5685
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "259"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   259
      Left            =   8505
      TabIndex        =   220
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   239
      Left            =   8505
      TabIndex        =   219
      Top             =   1275
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   219
      Left            =   8505
      TabIndex        =   218
      Top             =   1725
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   199
      Left            =   8505
      TabIndex        =   217
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   179
      Left            =   8505
      TabIndex        =   216
      Top             =   2595
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   159
      Left            =   8505
      TabIndex        =   215
      Top             =   3045
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   139
      Left            =   8505
      TabIndex        =   214
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   119
      Left            =   8505
      TabIndex        =   213
      Top             =   3915
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "99"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   99
      Left            =   8505
      TabIndex        =   212
      Top             =   4365
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   79
      Left            =   8505
      TabIndex        =   211
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   59
      Left            =   8505
      TabIndex        =   210
      Top             =   5235
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   39
      Left            =   8505
      TabIndex        =   209
      Top             =   5685
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "19"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   19
      Left            =   8505
      TabIndex        =   208
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "258"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   258
      Left            =   8040
      TabIndex        =   207
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   238
      Left            =   8040
      TabIndex        =   206
      Top             =   1275
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   218
      Left            =   8040
      TabIndex        =   205
      Top             =   1725
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   198
      Left            =   8040
      TabIndex        =   204
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   178
      Left            =   8040
      TabIndex        =   203
      Top             =   2595
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   158
      Left            =   8040
      TabIndex        =   202
      Top             =   3045
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   138
      Left            =   8040
      TabIndex        =   201
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   118
      Left            =   8040
      TabIndex        =   200
      Top             =   3915
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "98"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   98
      Left            =   8040
      TabIndex        =   199
      Top             =   4365
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   78
      Left            =   8040
      TabIndex        =   198
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   58
      Left            =   8040
      TabIndex        =   197
      Top             =   5235
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   38
      Left            =   8040
      TabIndex        =   196
      Top             =   5685
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "18"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   18
      Left            =   8040
      TabIndex        =   195
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "257"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   257
      Left            =   7560
      TabIndex        =   194
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   237
      Left            =   7560
      TabIndex        =   193
      Top             =   1275
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   217
      Left            =   7560
      TabIndex        =   192
      Top             =   1725
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   197
      Left            =   7560
      TabIndex        =   191
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   177
      Left            =   7560
      TabIndex        =   190
      Top             =   2595
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   157
      Left            =   7560
      TabIndex        =   189
      Top             =   3045
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   137
      Left            =   7560
      TabIndex        =   188
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   117
      Left            =   7560
      TabIndex        =   187
      Top             =   3915
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "97"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   97
      Left            =   7560
      TabIndex        =   186
      Top             =   4365
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   77
      Left            =   7560
      TabIndex        =   185
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   57
      Left            =   7560
      TabIndex        =   184
      Top             =   5235
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   37
      Left            =   7560
      TabIndex        =   183
      Top             =   5685
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "17"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   17
      Left            =   7560
      TabIndex        =   182
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "256"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   256
      Left            =   7095
      TabIndex        =   181
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   236
      Left            =   7095
      TabIndex        =   180
      Top             =   1275
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   216
      Left            =   7095
      TabIndex        =   179
      Top             =   1725
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   196
      Left            =   7095
      TabIndex        =   178
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   176
      Left            =   7095
      TabIndex        =   177
      Top             =   2595
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   156
      Left            =   7095
      TabIndex        =   176
      Top             =   3045
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   136
      Left            =   7095
      TabIndex        =   175
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   116
      Left            =   7095
      TabIndex        =   174
      Top             =   3915
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "96"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   96
      Left            =   7095
      TabIndex        =   173
      Top             =   4365
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   76
      Left            =   7095
      TabIndex        =   172
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   56
      Left            =   7095
      TabIndex        =   171
      Top             =   5235
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   36
      Left            =   7095
      TabIndex        =   170
      Top             =   5685
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "16"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   16
      Left            =   7095
      TabIndex        =   169
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "255"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   255
      Left            =   6630
      TabIndex        =   168
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   235
      Left            =   6630
      TabIndex        =   167
      Top             =   1275
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   215
      Left            =   6630
      TabIndex        =   166
      Top             =   1725
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   195
      Left            =   6630
      TabIndex        =   165
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   175
      Left            =   6630
      TabIndex        =   164
      Top             =   2595
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   155
      Left            =   6630
      TabIndex        =   163
      Top             =   3045
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   135
      Left            =   6630
      TabIndex        =   162
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   115
      Left            =   6630
      TabIndex        =   161
      Top             =   3915
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "95"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   95
      Left            =   6630
      TabIndex        =   160
      Top             =   4365
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "254"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   254
      Left            =   6150
      TabIndex        =   159
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   234
      Left            =   6150
      TabIndex        =   158
      Top             =   1275
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   214
      Left            =   6150
      TabIndex        =   157
      Top             =   1725
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   194
      Left            =   6150
      TabIndex        =   156
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   174
      Left            =   6150
      TabIndex        =   155
      Top             =   2595
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   154
      Left            =   6150
      TabIndex        =   154
      Top             =   3045
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   134
      Left            =   6150
      TabIndex        =   153
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   114
      Left            =   6150
      TabIndex        =   152
      Top             =   3915
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "94"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   94
      Left            =   6150
      TabIndex        =   151
      Top             =   4365
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "253"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   253
      Left            =   5685
      TabIndex        =   150
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   233
      Left            =   5685
      TabIndex        =   149
      Top             =   1275
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   213
      Left            =   5685
      TabIndex        =   148
      Top             =   1725
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   193
      Left            =   5685
      TabIndex        =   147
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   173
      Left            =   5685
      TabIndex        =   146
      Top             =   2595
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   153
      Left            =   5685
      TabIndex        =   145
      Top             =   3045
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   133
      Left            =   5685
      TabIndex        =   144
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   113
      Left            =   5685
      TabIndex        =   143
      Top             =   3915
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "93"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   93
      Left            =   5685
      TabIndex        =   142
      Top             =   4365
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "252"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   252
      Left            =   5220
      TabIndex        =   141
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   232
      Left            =   5220
      TabIndex        =   140
      Top             =   1275
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   212
      Left            =   5220
      TabIndex        =   139
      Top             =   1725
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   192
      Left            =   5220
      TabIndex        =   138
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   172
      Left            =   5220
      TabIndex        =   137
      Top             =   2595
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   152
      Left            =   5220
      TabIndex        =   136
      Top             =   3045
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   132
      Left            =   5220
      TabIndex        =   135
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   112
      Left            =   5220
      TabIndex        =   134
      Top             =   3915
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "92"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   92
      Left            =   5220
      TabIndex        =   133
      Top             =   4365
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "251"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   251
      Left            =   4740
      TabIndex        =   132
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   231
      Left            =   4740
      TabIndex        =   131
      Top             =   1275
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   211
      Left            =   4740
      TabIndex        =   130
      Top             =   1725
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   191
      Left            =   4740
      TabIndex        =   129
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   171
      Left            =   4740
      TabIndex        =   128
      Top             =   2595
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   151
      Left            =   4740
      TabIndex        =   127
      Top             =   3045
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   131
      Left            =   4740
      TabIndex        =   126
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   111
      Left            =   4740
      TabIndex        =   125
      Top             =   3915
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "91"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   91
      Left            =   4740
      TabIndex        =   124
      Top             =   4365
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   250
      Left            =   4275
      TabIndex        =   123
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   230
      Left            =   4275
      TabIndex        =   122
      Top             =   1275
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   210
      Left            =   4275
      TabIndex        =   121
      Top             =   1725
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   190
      Left            =   4275
      TabIndex        =   120
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   170
      Left            =   4275
      TabIndex        =   119
      Top             =   2595
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   150
      Left            =   4275
      TabIndex        =   118
      Top             =   3045
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   130
      Left            =   4275
      TabIndex        =   117
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   110
      Left            =   4275
      TabIndex        =   116
      Top             =   3915
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "90"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   90
      Left            =   4275
      TabIndex        =   115
      Top             =   4365
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "249"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   249
      Left            =   3810
      TabIndex        =   114
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   229
      Left            =   3810
      TabIndex        =   113
      Top             =   1275
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   209
      Left            =   3810
      TabIndex        =   112
      Top             =   1725
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   189
      Left            =   3810
      TabIndex        =   111
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   169
      Left            =   3810
      TabIndex        =   110
      Top             =   2595
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   149
      Left            =   3810
      TabIndex        =   109
      Top             =   3045
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   129
      Left            =   3810
      TabIndex        =   108
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   109
      Left            =   3810
      TabIndex        =   107
      Top             =   3915
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "89"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   89
      Left            =   3810
      TabIndex        =   106
      Top             =   4365
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "248"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   248
      Left            =   3330
      TabIndex        =   105
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   228
      Left            =   3330
      TabIndex        =   104
      Top             =   1275
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   208
      Left            =   3330
      TabIndex        =   103
      Top             =   1725
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   188
      Left            =   3330
      TabIndex        =   102
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   168
      Left            =   3330
      TabIndex        =   101
      Top             =   2595
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   148
      Left            =   3330
      TabIndex        =   100
      Top             =   3045
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   128
      Left            =   3330
      TabIndex        =   99
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   108
      Left            =   3330
      TabIndex        =   98
      Top             =   3915
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "88"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   88
      Left            =   3330
      TabIndex        =   97
      Top             =   4365
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "247"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   247
      Left            =   2865
      TabIndex        =   96
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   227
      Left            =   2865
      TabIndex        =   95
      Top             =   1275
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   207
      Left            =   2865
      TabIndex        =   94
      Top             =   1725
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   187
      Left            =   2865
      TabIndex        =   93
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   167
      Left            =   2865
      TabIndex        =   92
      Top             =   2595
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   147
      Left            =   2865
      TabIndex        =   91
      Top             =   3045
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   127
      Left            =   2865
      TabIndex        =   90
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   107
      Left            =   2865
      TabIndex        =   89
      Top             =   3915
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "87"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   87
      Left            =   2865
      TabIndex        =   88
      Top             =   4365
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "75"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   75
      Left            =   6630
      TabIndex        =   87
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "55"
      ForeColor       =   &H0000002F&
      Height          =   375
      Index           =   55
      Left            =   6630
      TabIndex        =   86
      Top             =   5235
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "35"
      ForeColor       =   &H0000001B&
      Height          =   375
      Index           =   35
      Left            =   6630
      TabIndex        =   85
      Top             =   5685
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "15"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   15
      Left            =   6630
      TabIndex        =   84
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   246
      Left            =   2400
      TabIndex        =   83
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   226
      Left            =   2400
      TabIndex        =   82
      Top             =   1275
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   206
      Left            =   2400
      TabIndex        =   81
      Top             =   1725
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   186
      Left            =   2400
      TabIndex        =   80
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   166
      Left            =   2400
      TabIndex        =   79
      Top             =   2595
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   146
      Left            =   2400
      TabIndex        =   78
      Top             =   3045
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   126
      Left            =   2400
      TabIndex        =   77
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   106
      Left            =   2400
      TabIndex        =   76
      Top             =   3915
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "86"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   86
      Left            =   2400
      TabIndex        =   75
      Top             =   4365
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "74"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   74
      Left            =   6150
      TabIndex        =   74
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "54"
      ForeColor       =   &H0000002E&
      Height          =   375
      Index           =   54
      Left            =   6150
      TabIndex        =   73
      Top             =   5235
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "34"
      ForeColor       =   &H0000001A&
      Height          =   375
      Index           =   34
      Left            =   6150
      TabIndex        =   72
      Top             =   5685
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "14"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   14
      Left            =   6150
      TabIndex        =   71
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   245
      Left            =   1920
      TabIndex        =   70
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   225
      Left            =   1920
      TabIndex        =   69
      Top             =   1275
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   205
      Left            =   1920
      TabIndex        =   68
      Top             =   1725
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   185
      Left            =   1920
      TabIndex        =   67
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   165
      Left            =   1920
      TabIndex        =   66
      Top             =   2595
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   145
      Left            =   1920
      TabIndex        =   65
      Top             =   3045
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   125
      Left            =   1920
      TabIndex        =   64
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   105
      Left            =   1920
      TabIndex        =   63
      Top             =   3915
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "85"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   85
      Left            =   1920
      TabIndex        =   62
      Top             =   4365
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "73"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   73
      Left            =   5685
      TabIndex        =   61
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "53"
      ForeColor       =   &H0000002D&
      Height          =   375
      Index           =   53
      Left            =   5685
      TabIndex        =   60
      Top             =   5235
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "33"
      ForeColor       =   &H00000019&
      Height          =   375
      Index           =   33
      Left            =   5685
      TabIndex        =   59
      Top             =   5685
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "13"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   13
      Left            =   5685
      TabIndex        =   58
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   244
      Left            =   1455
      TabIndex        =   57
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   224
      Left            =   1455
      TabIndex        =   56
      Top             =   1275
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   204
      Left            =   1455
      TabIndex        =   55
      Top             =   1725
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   184
      Left            =   1455
      TabIndex        =   54
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   164
      Left            =   1455
      TabIndex        =   53
      Top             =   2595
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   144
      Left            =   1455
      TabIndex        =   52
      Top             =   3045
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   124
      Left            =   1455
      TabIndex        =   51
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   104
      Left            =   1455
      TabIndex        =   50
      Top             =   3915
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "84"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   84
      Left            =   1455
      TabIndex        =   49
      Top             =   4365
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000040&
      Height          =   375
      Index           =   64
      Left            =   1455
      TabIndex        =   48
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H0000002C&
      Height          =   375
      Index           =   44
      Left            =   1455
      TabIndex        =   47
      Top             =   5235
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000018&
      Height          =   375
      Index           =   24
      Left            =   1455
      TabIndex        =   46
      Top             =   5685
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "4"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   1455
      TabIndex        =   45
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   243
      Left            =   990
      TabIndex        =   44
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   223
      Left            =   990
      TabIndex        =   43
      Top             =   1275
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   203
      Left            =   990
      TabIndex        =   42
      Top             =   1725
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   183
      Left            =   990
      TabIndex        =   41
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   163
      Left            =   990
      TabIndex        =   40
      Top             =   2595
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   143
      Left            =   990
      TabIndex        =   39
      Top             =   3045
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   123
      Left            =   990
      TabIndex        =   38
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   103
      Left            =   990
      TabIndex        =   37
      Top             =   3915
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "83"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   83
      Left            =   990
      TabIndex        =   36
      Top             =   4365
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H0000003F&
      Height          =   375
      Index           =   63
      Left            =   990
      TabIndex        =   35
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H0000002B&
      Height          =   375
      Index           =   43
      Left            =   990
      TabIndex        =   34
      Top             =   5235
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000017&
      Height          =   375
      Index           =   23
      Left            =   990
      TabIndex        =   33
      Top             =   5685
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   990
      TabIndex        =   32
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   242
      Left            =   510
      TabIndex        =   31
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   222
      Left            =   510
      TabIndex        =   30
      Top             =   1275
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   202
      Left            =   510
      TabIndex        =   29
      Top             =   1725
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   182
      Left            =   510
      TabIndex        =   28
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   162
      Left            =   510
      TabIndex        =   27
      Top             =   2595
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   142
      Left            =   510
      TabIndex        =   26
      Top             =   3045
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   122
      Left            =   510
      TabIndex        =   25
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   102
      Left            =   510
      TabIndex        =   24
      Top             =   3915
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "82"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   82
      Left            =   510
      TabIndex        =   23
      Top             =   4365
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H0000003E&
      Height          =   375
      Index           =   62
      Left            =   510
      TabIndex        =   22
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H0000002A&
      Height          =   375
      Index           =   42
      Left            =   510
      TabIndex        =   21
      Top             =   5235
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "250"
      ForeColor       =   &H00000016&
      Height          =   375
      Index           =   22
      Left            =   510
      TabIndex        =   20
      Top             =   5685
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2"
      ForeColor       =   &H00000002&
      Height          =   375
      Index           =   2
      Left            =   510
      TabIndex        =   19
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "241"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   241
      Left            =   45
      TabIndex        =   18
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "221"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   221
      Left            =   45
      TabIndex        =   17
      Top             =   1275
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "201"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   201
      Left            =   45
      TabIndex        =   16
      Top             =   1725
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "181"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   181
      Left            =   45
      TabIndex        =   15
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "161"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   161
      Left            =   45
      TabIndex        =   14
      Top             =   2595
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "141"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   141
      Left            =   45
      TabIndex        =   13
      Top             =   3045
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "121"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   121
      Left            =   45
      TabIndex        =   12
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "101"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   101
      Left            =   45
      TabIndex        =   11
      Top             =   3915
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "81"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   81
      Left            =   45
      TabIndex        =   10
      Top             =   4365
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "61"
      ForeColor       =   &H0000003D&
      Height          =   375
      Index           =   61
      Left            =   45
      TabIndex        =   9
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "41"
      ForeColor       =   &H00000029&
      Height          =   375
      Index           =   41
      Left            =   45
      TabIndex        =   8
      Top             =   5235
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "21"
      ForeColor       =   &H00000015&
      Height          =   375
      Index           =   21
      Left            =   45
      TabIndex        =   7
      Top             =   5685
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1"
      ForeColor       =   &H00000001&
      Height          =   375
      Index           =   1
      Left            =   45
      TabIndex        =   6
      Top             =   6120
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   375
      Index           =   0
      Left            =   3045
      TabIndex        =   1
      Top             =   7980
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   3735
      TabIndex        =   0
      Top             =   8055
      Visible         =   0   'False
      Width           =   1875
   End
End
Attribute VB_Name = "PANE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldColor(1 To Ar_Mpoykal) As Single   ' Public array variable.
Dim kodikos(1 To Ar_Mpoykal)     ' Public array variable.
Dim Perigrafh(1 To Ar_Mpoykal)   ' Public array variable.
Dim LHJH(1 To Ar_Mpoykal)   ' Public array variable.
Dim ebale
Dim M_PERIGRAFH
Dim Kafe, Keramidi, Prasino, kokkino, Portokali, ASPRO
Dim Ypoloipo(1 To Ar_Mpoykal)   ' Public array variable.




Function fernei(Index As Integer)

End Function

Sub AKYR_TOPOU(Index)
' ακύρωση τοποθέτησης μέσα
      Dim mADDRESS
      
      If Label3(Index).BackColor = kokkino Then ' αν  είναι κόκκινη θέση μπορώ να το ακυρώσω
         If Index + 1 <= ar_real_pothr Then
              If Label3(Index + 1).BackColor = kokkino Then
                 MsgBox mLResnew(913, "You can Undo only the last one.", 1)
                 Exit Sub
              End If
         End If
            mADDRESS = Label3(Index).Caption           'παίρνει τον αριθμό του μπουκαλιού
            Label1(Val(mADDRESS)).BackColor = OldColor(Val(mADDRESS)) 'ASPRO    ' το ασπρίζω
            Label1(Val(mADDRESS)).Caption = mADDRESS
            Label3(Index).Caption = Index
            Label3(Index).BackColor = OldColor(Index)   ' ασπρίζω το ποτήρι
           XIMITECH.Recordset.Index = "kod"
           XIMITECH.Recordset.Seek "=", kodikos(Val(mADDRESS))
           XIMITECH.Recordset.EDIT
           XIMITECH.Recordset("ADDR_PROT") = 0 'Format(0, "###")
           XIMITECH.Recordset("YPOL_PROT") = 0
           XIMITECH.Recordset.update
           Perigrafh(Val(mADDRESS)) = " "
      
      End If
End Sub
Private Sub back_Click()
Dim DUM, Ar_Mpoykal
   
 If system_ready = 0 Then
      MsgBox mL_Res(582)
      ' "Δεν έχει φορτωθεί το λειτουργικό σύστημα"
      Exit Sub
 End If
 
   
   back.Enabled = False
   Command1.Enabled = False
  PreHeating
On Error Resume Next
For i = 0 To PANE.Controls.Count - 1
     PANE.Controls(i).Enabled = False
Next i
On Error GoTo 0



 XIMITECH.Recordset.Index = "kod"
  
  
 
DUM = begin_talos(0)
  If DUM = 0 Then
        response = MsgBox(mL_Res(570), 64, "Talos")  ' no communication
        Exit Sub
  End If

' Do While Pause_Check() = 1
 '       response = MsgBox(mL_Res(571), 64, "Talos(571)")  ' To continue turn <<PAUSE>> to position 1 and press <<ΟΚ>>."
'Loop

 Do While Is_Porta() = "α"
          response = MsgBox(mL_Res(572), 64, "Talos")    ' klise porta
Loop

MsgBox mL_Res(573), , "Talos"  '"Eίναι οι θέσεις ΜΠΟΥΚΑΛΙΩΝ ΕΛΕΥΘΕΡΕΣ στο τραπέζι;"
Timer1.Enabled = True
Ar_Mpoykal = 0
For k = 1 To ar_real_pothr
  If Label3(k).BackColor = kokkino Then ' αν είναι κόκκινη θέση πάν'το μέσα
   Ar_Mpoykal = Ar_Mpoykal + 1
   cmd$ = "!var12=" + str(12000 + Val(Label3(k).Caption)) + ":!var13=" + str(k) + ":"
   DUM = wait_to_understand(cmd$)
   DUM = Wait_Robot(0)


     Label3(k).BackColor = &HFFFF80
       
   XIMITECH.Recordset.Seek "=", kodikos(Val(Label3(k).Caption))
   
   XIMITECH.Recordset.EDIT
        XIMITECH.Recordset("ADDR_PROT") = Val(Label3(k).Caption)
   XIMITECH.Recordset.update
   
   End If
Next

If Ar_Mpoykal = 0 Then
   MsgBox mL_Res(575), , "Talos(575)"  ' Define message.
   back.Enabled = True
Else

End If


Command1.Enabled = True

End Sub

Private Sub Check1_Click()
  If Check1 Then
     DBCombo1.Enabled = False
     DBCombo1.Visible = False
     Label4.Visible = False
     Text1.Visible = True
  Else
     DBCombo1.Enabled = True
     DBCombo1.Visible = True
     Label4.Visible = True
     Text1.Visible = False
  End If
   

End Sub

Private Sub Command1_Click()
 M_PERIGRAFH = "....................."
 If Label3(1).BackColor = kokkino Then
     ans = MsgBox(mLResnew(912, "Exit without transfer the bottles to the table.", 1), vbYesNo)
     If ans = vbNo Then
        Exit Sub
     End If
 End If
 Unload Me
End Sub

Private Sub DBCombo1_Click(Area As Integer)
   Dim x
  ' x = DBCombo1.SelectedItem
  ' XIMITECH.Recordset.Bookmark = x
   ebale = 1
   M_PERIGRAFH = DBCombo1
End Sub

Private Sub Form_DblClick()
 Dim a, b
      'CommonDialog1.ShowColor
      b = Me.BackColor
      a = save_colors(Me, Me.BackColor)
      If a = Prasino Or a = Kafe Or a = Portokali Or a = ASPRO Then
          Me.BackColor = b
      End If
      
      
End Sub

Private Sub Form_Load()
Dim L1, L6

'Me.BackColor =
a = find_colors(Me)
Kafe = &H808080 ' &H404040
Keramidi = &H80C0FF
Prasino = &H8000&
kokkino = &HFF&
Portokali = &H80FF&
ASPRO = &HFFFFFF


                           'Select Case MERES
                            '        Case 0
                             '            FIND_COLOR = &H80FF&           '  πορτοκαλί
                              '       Case Is > 0
                               '          FIND_COLOR = &H8000& '  πράσινο
                                '      Case Else
                                 '         FIND_COLOR = &HFF&         'κοκκινο
                        '    End Select










back.Caption = mL_CapRes(560)  'πήγαινε πίσω
Command1.Caption = mL_CapRes(561)  ' ejodos
Label4.Caption = mL_CapRes(565)  'Onoma ximikotexnikoy
Check1.Caption = mL_CapRes(566) ' ";;Χρήση Barcode"


If system_ready = 1 Then
     ValveReset
End If
MDIForm1.Arxeia(10).Visible = False
MDIForm1.Syntages(20).Visible = False
MDIForm1.bohuhtika(30).Visible = False
MDIForm1.ejodos(40).Visible = False

ebale = 0
flag_paint = 0
 

  aneb = 6285 - 5210
  
  
  For k = ar_real_pothr + 1 To Ar_Pothr
      Label3(k).Visible = False
  Next
  For k = 1 To ar_real_pothr
      Label3(k).top = Label3(k).top - aneb '+ IIf(k Mod 2 <> 0, 0, -Label3(k).height / 2)
  Next
  
 ' For k = 1 To Ar_Pothr
 '   Label3(k).top = Label3(k).top - aneb
 ' Next
  

  
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
 
  
  
  For k = 2 To Ar_Mpoykal Step 2
    Label1(k).top = Label1(k).top + Label1(k).height / 2
  Next
 ' DBCombo1.left = L1
 ' Label4.left = L1
  For k = Ar_1_mpoykal To Ar_tel_mpoykal
   
     OldColor(k) = Label1(k).BackColor
     Label1(k).Caption = k
 Next
     
  For k = 1 To ar_real_pothr
     Label3(k).Caption = k
     OldColor(k) = Label3(k).BackColor
 Next

     
End Sub


Private Sub Form_Paint()
Dim u, UU
If flag_paint = 0 Then
    flag_paint = 1

For k = 1 To Ar_Mpoykal
     Perigrafh(k) = ",,"
Next


 'XIMITECH.Recordset.Index = "KOD"
 'mpoyk.Recordset.Index = "ADDRESS"
 
 mpoyk.Recordset.MoveFirst
 Do While Not mpoyk.Recordset.eof
   If mpoyk.Recordset("Enabled") = 0 Then
       UU = Val(mpoyk.Recordset("address"))
       Label1(UU).BackColor = PANE.BackColor
       OldColor(UU) = Label1(UU).BackColor

   End If
    If mpoyk.Recordset("Enabled") = 2 Then
       UU = Val(mpoyk.Recordset("address"))
       Label1(UU).BackColor = Keramidi
       OldColor(UU) = Label1(UU).BackColor

   End If
    If mpoyk.Recordset("Enabled") = 1 Then
       UU = Val(mpoyk.Recordset("address"))
       Label1(UU).BackColor = ASPRO ' ασπρο
       OldColor(UU) = Label1(UU).BackColor
   End If
   mpoyk.Recordset.MoveNext
 Loop
 If XIMITECH.Recordset.RecordCount >= 1 Then
 Else
   MsgBox mL_Res(580)
   ' 580 "Δεν υπάρχουν διαθέσιμα χημικοτεχνικά για τοποθέτηση στο τραπέζι"
   Exit Sub
 End If
 XIMITECH.Recordset.MoveFirst
 While Not XIMITECH.Recordset.eof
         u = 0
         If Not IsNull(XIMITECH.Recordset("ADDR_PROT")) Then
                 u = XIMITECH.Recordset("addr_prot")
         Else
           XIMITECH.Recordset.EDIT
           XIMITECH.Recordset("addr_prot") = 0 '"   "
           XIMITECH.Recordset.update
         End If
         
         If u > 0 And u <= Ar_Mpoykal Then   ' =============================================
                                  
                 If Not IsNull(XIMITECH.Recordset("LHJ_PROT")) Then   ' ---------------
                          hmeres = DateDiff("d", Now, XIMITECH.Recordset("LHJ_PROT"))
                 Else
                          hmeres = 1000
                 End If
                 
                If XIMITECH.Recordset("skonh") = 1 Then ' ΣΤΕΡΕΑ
                                          Label1(u).BackColor = Kafe  ' &H808080           '   καφε
                Else  '    Δ Ι Α Λ Υ Μ Α Τ Α
                           Label1(u).BackColor = FIND_COLOR(hmeres)
                           'Select Case hmeres
                            '         Case 0
                             '             Label1(u).BackColor = Portokali          '  πορτοκαλί
                              '        Case Is > 1
                               '           Label1(u).BackColor = Prasino '  πράσινο
                                '      Case Else
                                 '         Label1(u).BackColor = kokkino    'κοκκινο
                            'End Select
                End If
                
                
                
                
                OldColor(u) = Label1(u).BackColor
                kodikos(u) = XIMITECH.Recordset("kod")
                Perigrafh(u) = XIMITECH.Recordset("mperigrafh")
                LHJH(u) = XIMITECH.Recordset("Lhj_prot")
                Ypoloipo(u) = XIMITECH.Recordset("ypol_prot")
                If Not IsNull(Perigrafh(u)) Then
                     Label1(u).Caption = Format(u, "###") + " " + Perigrafh(u)
                End If
                
                                'αν το υπόλοιπο είναι κάτω των 50 γραμμαρίων να κοκκινίζει
                 If Ypoloipo(u) < 50 Then
                       Label1(u).BackColor = &HFF&        'κοκκινο
                 End If

                
                
                

         End If ' ================================================================
         XIMITECH.Recordset.MoveNext
 Wend
'     XIMITECH.Recordset.Index = "PER"
     Screen.MousePointer = 1
End If
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


Private Sub front_Click()
Dim kodikos
   
For k = 2 To ar_real_pothr
   If Label3(k).BackColor = kokkino Then ' αν δεν είναι κόκκινη θέση φέρτο εδώ
      kodikos = 6000 + Ar_Mpoykal * k + Val(Label3(k).Caption)
      cmd$ = "!var12=" + str(kodikos) + ":"
      DUM = wait_to_understand(cmd$)
      While out_bit(bit_ok) = 0
           DoEvents
      Wend
      Label3(k).BackColor = &HFFFF80
   End If
Next



End Sub

Private Sub Form_Unload(Cancel As Integer)

MDIForm1.Arxeia(10).Visible = True
MDIForm1.Syntages(20).Visible = True
MDIForm1.bohuhtika(30).Visible = True
MDIForm1.ejodos(40).Visible = True

M_PERIGRAFH = "..............."
 
 If system_ready = 1 Then
       ValveReset
 End If

End Sub

Private Sub Label1_Click(Index As Integer)
Dim ok, tel_Pot
'   fernei (index)
'                                 ασπρο                                                     γκρίζο
tel_Pot = 1
ok = 1
' If Not IsNull(XIMITECH.Recordset("addr_PROT")) Then
'       If Val(XIMITECH.Recordset("addr_PROT")) > 0 Then
'            ok = 0
'            Print Chr$(7)
'            PANE.Caption = "υπάρχει ήδη"
'        End If
' Else
         
' End If

For k = Ar_Pothr To 1 Step -1
  If k = Val(Label3(k).Caption) Then
  Else
     If Index = Val(Label3(k).Caption) Then
          tel_Pot = 0
          AKYR_TOPOU k
          Exit Sub
     End If
     
     If Index = Val(Label3(k).Caption) Then
          MsgBox mLResnew(913, "You can Undo only the last one.", 1)
          Exit Sub
     End If
     
     
     'Exit For
  End If
Next

For k = 1 To Ar_Mpoykal
      If Perigrafh(k) = M_PERIGRAFH Then
          MsgBox mL_Res(576), , "Talos"  '  "Δεν μπορεί το ίδιο μπουκάλι να είναι σε δύο θέσεις"
          Exit Sub
      End If
Next





If Label1(Index).BackColor = Portokali Or ok = 0 Or ebale = 0 Then      ' Or Label1(Index).BackColor = &H80FFFF Then    'AN OXI ασπρο δεν κινείται
ElseIf Label1(Index).BackColor = Prasino Or ok = 0 Or ebale = 0 Then
ElseIf Label1(Index).BackColor = kokkino Or ok = 0 Or ebale = 0 Then
ElseIf Label1(Index).BackColor = Kafe Or ok = 0 Or ebale = 0 Then
ElseIf Label1(Index).BackColor = PANE.BackColor Or ok = 0 Or ebale = 0 Then

     '
Else
   If Check1 Then
     XIMITECH.Recordset.Index = "kod"
     XIMITECH.Recordset.Seek "=", Text1.text
   Else
     XIMITECH.Recordset.Index = "mPER"
     XIMITECH.Recordset.Seek "=", M_PERIGRAFH
   End If
   If XIMITECH.Recordset.NoMatch Then
          Exit Sub
   End If
     
   
   
   If XIMITECH.Recordset("skonh") = 1 Then ' ΣΤΕΡΕΑ
       If Label1(Index).BackColor = ASPRO Then
          MsgBox mL_Res(578), , "Talos" '"Αυτή είναι θέση για διαλύματα"
          Exit Sub
       End If
   End If
   If XIMITECH.Recordset("skonh") = 0 Then ' dialymata
       If Label1(Index).BackColor = Keramidi Then
          MsgBox mL_Res(579), , "Talos" ' "Αυτή είναι θέση για σκόνες"
          Exit Sub
       End If
   End If
   
   
   
   
   
   ' ψάχνει θέση για τοποθέτηση μπροστά
   For k = 1 To ar_real_pothr
      If Label3(k).BackColor <> kokkino Then ' αν δεν είναι κόκκινη θέση φέρτο εδώ
          Label3(k).Caption = Index            'παίρνει τον αριθμό του μπουκαλιού που θάρθεi
          Label3(k).BackColor = kokkino    'κοκκινίζω την μπροστά θέση
          Label1(Index).BackColor = &H80C0FF          'το αλλάζω χρωμα  για να μην ξαναζητηθεί
          Exit For
      End If
   Next
   
 'mpoyk.Recordset.Seek "=", Format(Index, "000")
   
 
   XIMITECH.Recordset.EDIT
'   XIMITECH.Recordset("ADDR_PROT") = Format(Index, "###")
   XIMITECH.Recordset("LHJ_PROT") = Date + XIMITECH.Recordset("ZVH_PROT")
   XIMITECH.Recordset("YPOL_PROT") = XIMITECH.Recordset("nerostox")
   XIMITECH.Recordset("KATOFLI") = 1000  ' mpoyk.Recordset("KATOFLI")
   XIMITECH.Recordset.update
'   XIMITECH.Recordset.Seek "=", "TRTRTRTR"
   
   
   
       
   u = Index
                                  
                 If Not IsNull(XIMITECH.Recordset("LHJ_PROT")) Then   ' ---------------
                          hmeres = DateDiff("d", Now, XIMITECH.Recordset("LHJ_PROT"))
                 Else
                          hmeres = 1000
                 End If
                 
                If XIMITECH.Recordset("skonh") = 1 Then ' ΣΤΕΡΕΑ
                                          Label1(u).BackColor = Kafe ' &H808080           '   καφε
                Else  '    Δ Ι Α Λ Υ Μ Α Τ Α
                           Select Case hmeres
                                     Case 0
                                          Label1(u).BackColor = Portokali         '  πορτοκαλί
                                      Case Is > 1
                                          Label1(u).BackColor = Prasino '  πράσινο
                                      Case Else
                                          Label1(u).BackColor = kokkino     'κοκκινο
                            End Select
                End If
                ' OldColor(u) = Label1(u).BackColor
                kodikos(u) = XIMITECH.Recordset("kod")
                Perigrafh(u) = XIMITECH.Recordset("mperigrafh")
                LHJH(u) = XIMITECH.Recordset("Lhj_prot")
                Ypoloipo(u) = XIMITECH.Recordset("ypol_prot")
                If Not IsNull(Perigrafh(u)) Then
                     Label1(u).Caption = Format(u, "###") + " " + Perigrafh(u)
                End If
End If
End Sub

Private Sub se_kefales(Index As Integer)
Dim kodikos

End Sub


Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
PANE.Caption = ""

If Not IsNull(Perigrafh(Index)) Then
     ' sxolia1.Caption = Perigrafh(Index)
    If left(Perigrafh(Index), 1) <> "," Then
        Label1(Index).ToolTipText = Perigrafh(Index)
       ' Me.Caption = Perigrafh(Index)
    End If
End If
If Not IsNull(Format(Ypoloipo(Index), "###") + "              " + Format(LHJH(Index), "dd/mm/yy")) And Len(Trim(PANE.Caption)) > 1 Then
     Label1(Index).ToolTipText = Format(Ypoloipo(Index), "###") + "              " + Format(LHJH(Index), "dd/mm/yy")
     SXOLIA1.Caption = Trim(SXOLIA1.Caption) + mL_CapRes(584) + Format(Ypoloipo(Index), "###") + mL_CapRes(583) + Format(LHJH(Index), "dd/mm/yy")
    ' Me.Caption = Me.Caption + mL_CapRes(584) + Format(Ypoloipo(Index), "###") + mL_CapRes(583) + Format(Lhjh(Index), "dd/mm/yy")
End If
End Sub



Private Sub Label3_Click(Index As Integer)
' ακύρωση τοποθέτησης μέσα
      Dim mADDRESS
      
      If Label3(Index).BackColor = kokkino Then ' αν  είναι κόκκινη θέση μπορώ να το ακυρώσω
         If Index + 1 <= ar_real_pothr Then
              If Label3(Index + 1).BackColor = kokkino Then
                 MsgBox mLResnew(913, "You can Undo only the last one.", 1)
                 Exit Sub
              End If
         End If
            mADDRESS = Label3(Index).Caption           'παίρνει τον αριθμό του μπουκαλιού
            Label1(Val(mADDRESS)).BackColor = OldColor(Val(mADDRESS)) 'ASPRO    ' το ασπρίζω
            Label1(Val(mADDRESS)).Caption = mADDRESS
            Label3(Index).Caption = Index
            Label3(Index).BackColor = OldColor(Index)   ' ασπρίζω το ποτήρι
           XIMITECH.Recordset.Index = "kod"
           XIMITECH.Recordset.Seek "=", kodikos(Val(mADDRESS))
           XIMITECH.Recordset.EDIT
           XIMITECH.Recordset("ADDR_PROT") = 0 'Format(0, "###")
           XIMITECH.Recordset("YPOL_PROT") = 0
           XIMITECH.Recordset.update
           Perigrafh(Val(mADDRESS)) = " "
      
      End If



End Sub

Private Sub Label3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim n
PANE.Caption = ""

n = Label3(Index).Caption

If n > 60 Then
   
    If Not IsNull(Perigrafh(n)) Then
      'sxolia1.Caption = Perigrafh(Index)
      If left(Perigrafh(n), 1) <> "," Then
         'sxolia1.Caption = Perigrafh(n)
         Me.Caption = Perigrafh(n)
      End If
    End If
    
    If Not IsNull(Format(Ypoloipo(n), "###") + "              " + Format(LHJH(n), "dd/mm/yy")) And Len(Trim(PANE.Caption)) > 1 Then
        'sxolia2.Caption = Format(Ypoloipo(Index), "###") + "              " + Format(Lhjh(Index), "dd/mm/yy")
        'sxolia1.Caption = Trim(sxolia1.Caption) + mL_CapRes(584) + Format(Ypoloipo(n), "###") + mL_CapRes(583) + Format(Lhjh(n), "dd/mm/yy")
        Me.Caption = Me.Caption + mL_CapRes(584) + Format(Ypoloipo(n), "###") + mL_CapRes(583) + Format(LHJH(n), "dd/mm/yy")
    End If
    
End If
End Sub


Private Sub Text1_LostFocus()
   
   XIMITECH.Recordset.Index = "kod"
   XIMITECH.Recordset.Seek "=", left(Text1.text, 6)
   If XIMITECH.Recordset.eof Then
       MsgBox mL_Res(581)
       ' 581 "Δεν βρέθηκε τέτοιο χημικό "
       Label5.Caption = ""
  Else
     Label5.Caption = XIMITECH.Recordset("mperigrafh")
     M_PERIGRAFH = XIMITECH.Recordset("mperigrafh")
     ebale = 1
  End If
  
  End Sub

Private Sub Timer1_Timer()
                If back.Enabled = False Then
                      If system_ready = 1 Then
                            Elegxo_Katastash
                      End If
                End If
End Sub

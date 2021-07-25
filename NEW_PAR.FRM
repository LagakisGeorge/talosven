VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form PARAMETROI 
   Caption         =   "Form1"
   ClientHeight    =   6975
   ClientLeft      =   2340
   ClientTop       =   720
   ClientWidth     =   9660
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6975
   ScaleWidth      =   9660
   WindowState     =   2  'Maximized
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\RECIPIES.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   360
      Left            =   150
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "parametroi"
      Top             =   6090
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\RECIPIES.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   330
      Left            =   135
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "users"
      Top             =   5745
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Εξοδος"
      Height          =   690
      Left            =   6990
      TabIndex        =   21
      Top             =   5715
      Width           =   2400
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5730
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   10107
      _Version        =   393216
      Tabs            =   6
      TabHeight       =   529
      BackColor       =   12640511
      TabCaption(0)   =   "Γενικά Συνταγές"
      TabPicture(0)   =   "NEW_PAR.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label11"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label12(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label13"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label12(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label12(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label12(3)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label14"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label15"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label12(5)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text4"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Check4"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Check1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text3"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Check6"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text11"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text12(1)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text12(6)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text12(11)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text12(16)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text13"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text14"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text12(21)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).ControlCount=   28
      TabCaption(1)   =   "Νερό Συνταγών"
      TabPicture(1)   =   "NEW_PAR.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text8"
      Tab(1).Control(1)=   "Text10"
      Tab(1).Control(2)=   "Label8"
      Tab(1).Control(3)=   "Label10"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Διαλύματα"
      TabPicture(2)   =   "NEW_PAR.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text7"
      Tab(2).Control(1)=   "Text6"
      Tab(2).Control(2)=   "Text5"
      Tab(2).Control(3)=   "Text9"
      Tab(2).Control(4)=   "Label7"
      Tab(2).Control(5)=   "Label6"
      Tab(2).Control(6)=   "Label5"
      Tab(2).Control(7)=   "Label9"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Προγραμματος Γενικά"
      TabPicture(3)   =   "NEW_PAR.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Check5"
      Tab(3).Control(1)=   "Check3"
      Tab(3).Control(2)=   "Check2"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Users"
      TabPicture(4)   =   "NEW_PAR.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "DBGrid1"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Ρυθμισεις για αλάτι"
      TabPicture(5)   =   "NEW_PAR.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Command3"
      Tab(5).Control(1)=   "DBGrid2"
      Tab(5).Control(2)=   "Command2"
      Tab(5).ControlCount=   3
      Begin VB.CommandButton Command3 
         Caption         =   "ΔΙΟΡΘΩΣΗ "
         Height          =   360
         Left            =   -74760
         TabIndex        =   49
         Top             =   1380
         Width           =   990
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "NEW_PAR.frx":00A8
         Height          =   3345
         Left            =   -73230
         OleObjectBlob   =   "NEW_PAR.frx":00BC
         TabIndex        =   48
         Top             =   1020
         Width           =   6345
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Δημιουργία Πίνακα"
         Height          =   360
         Left            =   -74790
         TabIndex        =   47
         Top             =   720
         Width           =   990
      End
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   21
         Left            =   7575
         MaxLength       =   7
         TabIndex        =   45
         Top             =   3480
         Width           =   1245
      End
      Begin VB.TextBox Text14 
         Height          =   360
         Left            =   3315
         TabIndex        =   42
         Top             =   5130
         Width           =   1245
      End
      Begin VB.TextBox Text13 
         Height          =   360
         Left            =   3315
         TabIndex        =   41
         Top             =   4695
         Width           =   1245
      End
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   16
         Left            =   7560
         MaxLength       =   5
         TabIndex        =   39
         Top             =   3000
         Width           =   1245
      End
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   7560
         MaxLength       =   5
         TabIndex        =   37
         Top             =   2520
         Width           =   1245
      End
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   7560
         MaxLength       =   5
         TabIndex        =   35
         Top             =   2040
         Width           =   1245
      End
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   7560
         MaxLength       =   5
         TabIndex        =   32
         Top             =   1560
         Width           =   1245
      End
      Begin VB.TextBox Text11 
         DataField       =   "uesh_alatioy"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   3315
         TabIndex        =   29
         Top             =   1680
         Width           =   1245
      End
      Begin VB.CheckBox Check6 
         Alignment       =   1  'Right Justify
         Caption         =   "Χρησιμοποιώ ξηρό αλάτι"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   870
         TabIndex        =   28
         Top             =   1335
         Width           =   2640
      End
      Begin VB.TextBox Text7 
         DataField       =   "WEIGHT_BOTTLE_MAX"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   -71850
         TabIndex        =   25
         Top             =   3750
         Width           =   1245
      End
      Begin VB.TextBox Text6 
         DataField       =   "WEIGHT_BOTTLE_MIN"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   -71865
         TabIndex        =   24
         Top             =   3360
         Width           =   1245
      End
      Begin VB.TextBox Text3 
         Height          =   360
         Left            =   3315
         TabIndex        =   22
         Top             =   4260
         Width           =   1245
      End
      Begin VB.TextBox Text5 
         DataField       =   "metatopish_W"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   -71850
         TabIndex        =   18
         Top             =   2955
         Width           =   1245
      End
      Begin VB.TextBox Text9 
         DataField       =   "TIME_PREHEAT"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   -71865
         TabIndex        =   17
         Top             =   2565
         Width           =   1245
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "Αλλάζω χρώματα"
         DataField       =   "allazo_xromata"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   -74145
         TabIndex        =   15
         Top             =   1395
         Width           =   2610
      End
      Begin VB.CheckBox Check3 
         Alignment       =   1  'Right Justify
         Caption         =   "Αλλάζω fonts"
         DataField       =   "allazo_fonts"
         DataSource      =   "Data1"
         Height          =   405
         Left            =   -74130
         TabIndex        =   14
         Top             =   1860
         Width           =   2610
      End
      Begin VB.CheckBox Check5 
         Alignment       =   1  'Right Justify
         Caption         =   "Numbers on Labels"
         Height          =   435
         Left            =   -74115
         TabIndex        =   13
         Top             =   2355
         Width           =   2610
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Προζυγισμένα πανιά"
         DataField       =   "prozygismena"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   885
         TabIndex        =   12
         Top             =   2130
         Width           =   2640
      End
      Begin VB.CheckBox Check4 
         Alignment       =   1  'Right Justify
         Caption         =   "Δοκιμαστική συνταγή"
         DataField       =   "dokimastikh_syntagh"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   900
         TabIndex        =   11
         Top             =   2415
         Width           =   2640
      End
      Begin VB.TextBox Text8 
         DataField       =   "MAXTIMEWATER"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   -70995
         TabIndex        =   8
         Top             =   1185
         Width           =   1245
      End
      Begin VB.TextBox Text10 
         Height          =   360
         Left            =   -70995
         TabIndex        =   7
         Top             =   1620
         Width           =   1245
      End
      Begin VB.TextBox Text1 
         DataField       =   "AXRHSIMOPOIHTES_ORES"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   3315
         TabIndex        =   3
         Top             =   2865
         Width           =   1245
      End
      Begin VB.TextBox Text2 
         DataField       =   "default_baros_panioy"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   3315
         TabIndex        =   2
         Top             =   3330
         Width           =   1245
      End
      Begin VB.TextBox Text4 
         DataField       =   "MPANIO"
         DataSource      =   "Data1"
         Height          =   360
         Left            =   3315
         TabIndex        =   1
         Top             =   3855
         Width           =   1245
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "NEW_PAR.frx":0A8F
         Height          =   1500
         Left            =   -73830
         OleObjectBlob   =   "NEW_PAR.frx":0AA3
         TabIndex        =   16
         Top             =   1170
         Width           =   2670
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Water Tolerance  in mgrs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   4860
         TabIndex        =   46
         Top             =   3480
         Width           =   2700
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Disp.Parameter A ( 2g-)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         TabIndex        =   44
         Top             =   4770
         Width           =   2655
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Disp.Parameter B(-2g)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   615
         TabIndex        =   43
         Top             =   5190
         Width           =   2655
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "16001-  mgrs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   4800
         TabIndex        =   40
         Top             =   3000
         Width           =   2700
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "8001-16000 mgrs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   4800
         TabIndex        =   38
         Top             =   2520
         Width           =   2700
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "2001 - 8000 mgrs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   4800
         TabIndex        =   36
         Top             =   2040
         Width           =   2700
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Tolerance in mgrs to considered OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   34
         Top             =   960
         Width           =   3855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Up to 2000mgrs "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   4800
         TabIndex        =   33
         Top             =   1560
         Width           =   2700
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Θέση Αλατιού"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   345
         TabIndex        =   30
         Top             =   1710
         Width           =   2655
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Μέγιστο Βάρος Αδειου Μπουκαλιού"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74640
         TabIndex        =   27
         Top             =   3765
         Width           =   2655
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ελάχιστο Βάρος Αδειου Μπουκαλιού"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74685
         TabIndex        =   26
         Top             =   3420
         Width           =   2760
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Max.Capacity of Beaker"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   570
         TabIndex        =   23
         Top             =   4335
         Width           =   2655
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Displacement of W"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74640
         TabIndex        =   20
         Top             =   2970
         Width           =   2655
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Preheating time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74685
         TabIndex        =   19
         Top             =   2625
         Width           =   2760
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Max.Time in Recipe Water Dispense (s)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74760
         TabIndex        =   10
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Stop dispensing before (mgr)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -73785
         TabIndex        =   9
         Top             =   1665
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Default βάρος πανιού"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         TabIndex        =   6
         Top             =   3435
         Width           =   2655
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Οριο ωρών για δοκιμαστικές συνταγές"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   630
         TabIndex        =   5
         Top             =   2805
         Width           =   2655
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Default Liquor Ratio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   585
         TabIndex        =   4
         Top             =   3930
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6885
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   9705
      Begin VB.Data Data3 
         Caption         =   "ALATI"
         Connect         =   "Access"
         DatabaseName    =   "C:\TALOS\RECIPIES.MDB"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   360
         Left            =   0
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT * FROM PARAMETROI"
         Top             =   0
         Visible         =   0   'False
         Width           =   1800
      End
   End
End
Attribute VB_Name = "PARAMETROI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check5_Click()
  If Check5 Then numbers_on_labels = 1
End Sub


Private Sub Command1_Click()
Dim SS As String
 Data1.Recordset.MoveFirst
 Data1.UpdateControls
 
'-------------------------------------------
 Data1.Recordset.EDIT
 If Check6.value = 1 Then
   Data1.Recordset("doyleyo_alati") = 1
 Else
   Data1.Recordset("doyleyo_alati") = 0
 End If
 Data1.Recordset("fornext_big") = Text13.text
 Data1.Recordset("fornext_small") = Text14.text
 Data1.Recordset.update
'-------------------------------------------
 
 
 
 Set r = Data1.Recordset
  
  r.MoveNext
  r.EDIT
  r("maxtimewater") = Val(Text10.text)
  r("Default_baros_panioy") = Text3.text

SS = ""
For k = 1 To 21 Step 5
   SS = SS + left(Text12(k).text + Space(5), 5)
Next
  
  
  r("lsneg") = SS
'  Avance_Online_Water = Val(Text10.text)
'  YperXeilish = Text3.text
  
  
  r.update
r.Close



Load_Parametroi

MDIForm1.Arxeia(10).Visible = True
MDIForm1.Syntages(20).Visible = True
MDIForm1.bohuhtika(30).Visible = True
MDIForm1.ejodos(40).Visible = True
  
  
  
  
  
  
  
  
  
  Unload Me
End Sub


Private Sub Command2_Click()
'default μπάνιο
Set mydb = Workspaces(0).OpenDatabase("c:\TALOS\RECIPIES.MDB", False, False)

On Error Resume Next
mydb.Execute "create table ALATI (APO SINGLE,EOS SINGLE ,TIMH1 SINGLE,TIMH2 SINGLE,TIMH3 SINGLE)"

'Set r = mydb.OpenRecordset("parametroi")
'ok = 0
'For k = 0 To r.Fields.Count - 1
'     If UCase(r.Fields(k).Name) = "MPANIO" Then M_AP = 1
'Next
End Sub

Private Sub Form_Load()

a = find_colors(Me)
MDIForm1.Arxeia(10).Visible = False
MDIForm1.Syntages(20).Visible = False
MDIForm1.bohuhtika(30).Visible = False
MDIForm1.ejodos(40).Visible = False





On Error Resume Next
Data3.RecordSource = "SELECT * FROM ALATI"
Data3.Refresh









Check1.Caption = mL_CapRes(750)  ' "Προζυγισμένα πανιά"
Check2.Caption = mL_CapRes(751)  'Αλλάζω χρώματα
Check3.Caption = mL_CapRes(752)  'Αλλάζω fonts
Check4.Caption = mL_CapRes(753)  'Δοκιμαστική συνταγή
Check5.Caption = mLResnew(924, "Numbers on Labels", 0)
DBGrid1.Caption = mL_CapRes(754) 'Χειριστές
DBGrid1.Columns(0).Caption = mL_CapRes(755) ' χειριστής
Label2.Caption = mL_CapRes(757) ' Οριο ωρών για δοκιμαστικές συνταγές
Label1.Caption = mL_CapRes(756) 'Default βάρος πανιού
Label9.Caption = mLResnew(925, "Preheating Time ms(1000=1sec*", 0)
Label10.Caption = mLResnew(926, "Water Dispensing Parameter*", 0)

'3-3
Check6.Caption = mLResnew(933, "Dry salt in Use *", 0)
Label11.Caption = mLResnew(934, "Address of Dry salt*", 0)
Label6.Caption = mLResnew(935, "Minimum weight of empty container*", 0)
Label7.Caption = mLResnew(936, "Maximum weight of empty container*", 0)



 SSTab1.TabCaption(0) = mLResnew(937, "Recipes General*", 0)
 SSTab1.TabCaption(1) = mLResnew(938, "Water of Recipes *", 0)
 SSTab1.TabCaption(2) = mLResnew(939, "Solutions*", 0)
 SSTab1.TabCaption(3) = mLResnew(940, "General*", 0)
 SSTab1.TabCaption(4) = mLResnew(941, "Users*", 0)
 DBGrid1.Caption = mLResnew(942, "Users*", 0)
 DBGrid1.Columns(0).Caption = mLResnew(943, "User*", 0)
 Label4.Caption = mLResnew(944, "Default Liquor ratio*", 0)
 Label3.Caption = mLResnew(945, "Maximum capacity of beakers*", 0)
 Label8.Caption = mLResnew(946, "Max.Time in Recipe Water Dispense (s)*", 0)
 Label5.Caption = mLResnew(947, "Displacement of W*", 0)

'Label10.Caption = mLResnew(926, "Water Dispensing Parameter*", 0)


'3-3


Command1.Caption = mL_CapRes(759) 'Εξοδος

If DOYLEYO_ALATI = 0 Then
     Text3.Visible = False
     Label3.Visible = False
     Check6.value = 0 'False
Else
     Check6.value = 1 'True

End If







'default μπάνιο
Set mydb = Workspaces(0).OpenDatabase("c:\TALOS\RECIPIES.MDB", False, False)
Set r = mydb.OpenRecordset("parametroi")
ok = 0
For k = 0 To r.Fields.Count - 1
     If UCase(r.Fields(k).Name) = "MPANIO" Then M_AP = 1
Next

Text13.text = r("Fornext_Big")
Text14.text = r("Fornext_small")


r.MoveNext

On Error GoTo addmore

Text10.text = r("maxtimewater")



'Avance_Online_Water = r("maxtimewater")
Text3.text = YperXeilish
For k = 1 To 21 Step 5
   Text12(k).text = Mid$(r("lsneg"), k, 5)
Next
  ' Text12.text = r("MPANIO") 'POTE EINAI OK H SYNTAGH
r.Close
On Error Resume Next
   If M_AP = 0 Then mydb.Execute "ALTER TABLE PARAMETROI ADD COLUMN MPANIO TEXT(10)"
'------------
Exit Sub

addmore:
       r.AddNew
       r.update
       Resume Next
'End Sub

End Sub



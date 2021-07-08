VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmXHMedit 
   ClientHeight    =   6015
   ClientLeft      =   195
   ClientTop       =   405
   ClientWidth     =   9660
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6015
   ScaleWidth      =   9660
   WindowState     =   2  'Maximized
   Begin VB.ComboBox dbCombo3 
      Height          =   315
      Left            =   8160
      TabIndex        =   44
      Text            =   "Combo1"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.ComboBox dbCombo2 
      Height          =   315
      Left            =   8160
      TabIndex        =   43
      Text            =   "Combo1"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox Reyst2 
      DataField       =   "MAX_SYGK"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2760
      TabIndex        =   35
      Top             =   2040
      Width           =   975
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   "c:\talos\RECIPIES.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   2370
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "NERA"
      Top             =   6315
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   "c:\talos\RECIPIES.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   405
      Left            =   945
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "NERA"
      Top             =   6270
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\RECIPIES.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   450
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from ximitech"
      Top             =   6585
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Σκόνη"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   5970
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1350
      Width           =   1020
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Διάλυμα"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   5955
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   375
      Width           =   990
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\recipies.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   330
      Left            =   5160
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   0  'Table
      RecordSource    =   "EIDDIAL"
      Top             =   7275
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox txtFields9 
      Alignment       =   2  'Center
      DataField       =   "KATASKEYAS"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   9
      Left            =   2775
      MaxLength       =   14
      TabIndex        =   0
      Top             =   1140
      Width           =   960
   End
   Begin VB.TextBox txtFields 
      Alignment       =   2  'Center
      DataField       =   "NEROSTOX"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   8
      Left            =   2760
      TabIndex        =   6
      Top             =   3585
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      Alignment       =   2  'Center
      DataField       =   "ZVH_PROT"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   2
      Left            =   2760
      TabIndex        =   5
      Top             =   3255
      Width           =   1935
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      DataField       =   "BOHU_XHM"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   8190
      TabIndex        =   11
      Text            =   "0"
      Top             =   3660
      Width           =   1005
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      DataField       =   "MIXTIM_FIN"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   8190
      TabIndex        =   10
      Text            =   "40"
      Top             =   2625
      Width           =   1005
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      DataField       =   "MIXTIM"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   8190
      TabIndex        =   9
      Text            =   "60"
      Top             =   2265
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Εξοδος"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7290
      TabIndex        =   13
      Top             =   5205
      Width           =   1215
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Caption         =   "data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\RECIPIES.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "XHMEDIT"
      Top             =   5670
      Visible         =   0   'False
      Width           =   9660
   End
   Begin VB.TextBox txtFields 
      Alignment       =   2  'Center
      DataField       =   "ADDR_PROT"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   2760
      MaxLength       =   3
      TabIndex        =   8
      Top             =   4230
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      Alignment       =   2  'Center
      DataField       =   "YPOL_PROT"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   2760
      TabIndex        =   7
      Top             =   3915
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      Alignment       =   2  'Center
      DataField       =   "LHJ_PROT"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   2760
      TabIndex        =   4
      Top             =   2955
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   2760
      TabIndex        =   3
      Top             =   2655
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      Alignment       =   2  'Center
      DataField       =   "PERIGRAFH"
      DataSource      =   "Data1"
      Height          =   315
      Index           =   1
      Left            =   2775
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1425
      Width           =   2865
   End
   Begin VB.TextBox txtFields 
      Alignment       =   2  'Center
      DataField       =   "KOD"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   2760
      MaxLength       =   6
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   855
      Width           =   1455
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      DataField       =   "NERO_DIAL"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   8190
      TabIndex        =   12
      Text            =   "0"
      Top             =   4005
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Height          =   1470
      Left            =   5910
      TabIndex        =   14
      Top             =   225
      Width           =   1410
   End
   Begin VB.Frame Frame2 
      Height          =   1470
      Left            =   7365
      TabIndex        =   37
      Top             =   225
      Width           =   1845
      Begin VB.OptionButton Option2 
         Caption         =   "Chem 5"
         Enabled         =   0   'False
         Height          =   255
         Index           =   5
         Left            =   100
         TabIndex        =   42
         Top             =   1155
         Width           =   1665
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Chem 2"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   100
         TabIndex        =   41
         Top             =   405
         Width           =   1665
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Chem 4"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   100
         TabIndex        =   40
         Top             =   900
         Width           =   1665
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Chem 3"
         Height          =   255
         Index           =   3
         Left            =   100
         TabIndex        =   39
         Top             =   650
         Width           =   1665
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Chem 1"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   100
         TabIndex        =   38
         Top             =   165
         Width           =   1665
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   225
      Top             =   6195
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Bindings        =   "_XHMEDIT.frx":0000
      Left            =   3225
      Top             =   6375
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "c:\TALOS\xhm1.rpt"
      Destination     =   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "_XHMEDIT.frx":0014
      DataField       =   "MORFH"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   2760
      TabIndex        =   2
      Top             =   2325
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      ListField       =   "EIDOS"
      BoundColumn     =   "EIDOS"
      Text            =   "DBCombo1"
   End
   Begin VB.Label Reyst1 
      BackStyle       =   0  'Transparent
      Caption         =   "Flow Speed"
      Height          =   255
      Left            =   1080
      TabIndex        =   36
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1050
      TabIndex        =   34
      Top             =   45
      Width           =   4845
   End
   Begin VB.Shape Shape1 
      Height          =   5925
      Left            =   270
      Top             =   15
      Width           =   9240
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Ταυτότητα"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   14
      Left            =   1785
      TabIndex        =   31
      Top             =   1200
      Width           =   930
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Αρχικό νερό διάλυσης (gr)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5865
      TabIndex        =   30
      Top             =   4005
      Width           =   2250
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ποσότητα ανανέωσ(gr)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   525
      TabIndex        =   29
      Top             =   3555
      Width           =   2190
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Μέρες Ζωής"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   885
      TabIndex        =   28
      Top             =   3225
      Width           =   1830
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Βοηθητικά Χημικά (mgr)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5865
      TabIndex        =   27
      Top             =   3660
      Width           =   2250
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Είδος νερού διάλυσης"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   5865
      TabIndex        =   26
      Top             =   2985
      Width           =   2250
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Είδος νερού αραίωσης"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   5865
      TabIndex        =   25
      Top             =   3345
      Width           =   2250
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Χρόνος τελ. ανάμειξης sec"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   5865
      TabIndex        =   24
      Top             =   2625
      Width           =   2250
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Χρόνος διάλυσης σε sec"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5865
      TabIndex        =   23
      Top             =   2310
      Width           =   2250
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Διεύθυνση"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   7
      Left            =   900
      TabIndex        =   22
      Top             =   4200
      Width           =   1830
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Υπόλοιπο (gr)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   6
      Left            =   900
      TabIndex        =   21
      Top             =   3885
      Width           =   1830
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Ημερομηνία Λήξης"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   5
      Left            =   885
      TabIndex        =   20
      Top             =   2940
      Width           =   1830
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Πραγμ.συγκέντρωση"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   4
      Left            =   885
      TabIndex        =   19
      Top             =   2625
      Width           =   1830
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Συγκέντρωση  (w/w)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   18
      Top             =   2325
      Width           =   1905
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Περιγραφή"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   855
      TabIndex        =   17
      Top             =   1440
      Width           =   1860
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Κωδικός"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   855
      TabIndex        =   15
      Top             =   885
      Width           =   1860
   End
End
Attribute VB_Name = "frmXHMedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XX(1 To 3)
Dim real_new
Dim mtxtFields2, txtFields8

Dim proth_fora

Sub load_fields()
'On Error Resume Next

'           Data3.Recordset("kod") = Data1.Recordset("kod")
 '          Data3.Recordset.Update
          ' Data3.RecordSource = "select *from ximitech where [kod]='" + b + "'"
          ' Data3.Refresh







'Data3.Refresh
'If Data3.Recordset.RecordCount = 0 Then
'   MsgBox "Οι αλλαγές δεν καταχωρήθηκαν"
'   Exit Sub
'End If

'Data3.Recordset.MoveFirst
'Data1.Recordset.MoveFirst
             

'  Data3.Recordset.Edit
  Data3.Recordset("kod") = Data1.Recordset("kod")
  Data3.Recordset("perigrafh") = Data1.Recordset("perigrafh")
  Data3.Recordset("addr_prot") = Data1.Recordset("addr_prot")
  
   Data3.Recordset("max_sygk") = Data1.Recordset("max_sygk")
  
  
 
      Data3.Recordset("lhj_prot") = Data1.Recordset("lhj_prot")
  Data3.Recordset("zvh_prot") = Data1.Recordset("zvh_prot")
  Data3.Recordset("ypol_prot") = Data1.Recordset("ypol_prot")
  Data3.Recordset("kataskeyas") = LTrim(Data1.Recordset("kataskeyas"))

  Data3.Recordset("BOHU_XHM") = Data1.Recordset("BOHU_XHM")
  
 ' Data3.Recordset("real_cons") = Data1.Recordset("real_cons")
  Data3.Recordset("nero_dial") = Data1.Recordset("nero_dial")
  Data3.Recordset("mixtim") = Data1.Recordset("mixtim")
  Data3.Recordset("mixtim_fin") = Data1.Recordset("mixtim_fin")
  Data3.Recordset("thermo1") = Data1.Recordset("thermo1")
  Data3.Recordset("thermo2") = Data1.Recordset("thermo2")
  Data3.Recordset("nerostox") = Data1.Recordset("nerostox")
  Data3.Recordset("skonh") = IIf(Option1(1) = True, 1, 0)


If frmXHM.cmdAdd.Enabled = False Then 'add rec
   
   For k = 1 To 5
       If Option2(k) Then
           Data3.Recordset("timh") = k
           Data3.Recordset("addr_prot") = 50 + k
       End If
   Next
   
   If Option1(1) Then    ' skonh
      Data3.Recordset("morfh") = "      "
   Else
      Data3.Recordset("morfh") = Data1.Recordset("morfh")
   End If

End If




Data3.Recordset("mperigrafh") = left(Trim(Data3.Recordset("perigrafh")) + "-" + Trim(Data3.Recordset("morfh")) + Space(35), 35)
Data3.Recordset.update
On Error GoTo 0
End Sub



Private Sub cmdDelete_Click()
  
 If Not IsNull(Data1.Recordset("addr_prot")) Then
     If Data1.Recordset("addr_prot") > 0 Then
           MsgBox mL_Res(162), , "Talos"
           '"Δεν μπορεί να διαγραφεί επειδή το αντίστοιχο δοχείο καταλαμβάνει μία διεύθυνση στο καρτεσιανό τραπέζι.Να γίνει πρώτα απομάκρυνση με την χρήση του Robot ", , "Talos"
           Exit Sub
    End If
End If
  'this may produce an error if you delete the last
  'record or the only record in the recordset
 ans = MsgBox(mL_Res(163), 20, "Talos")
 '"Θέλετε να διαγράψετε αυτή την εγγραφή ;", 20, "Talos")
 If ans = 7 Then
     Exit Sub
Else
   Data1.Recordset.DELETE
        
   Data1.Recordset.MoveNext
End If
End Sub


Function Group_Enable(DUM)
'DBCombo1.Enabled = dum
  txtFields(1).Enabled = DUM
  txtFields(8).Enabled = DUM
'  txtFields(6).Enabled = dum
'  txtFields(3).Enabled = dum
   txtFields(2).Enabled = DUM
'  Text3.Enabled = dum
'  Text4.Enabled = dum
  Text5.Enabled = DUM
  Text6.Enabled = DUM
 ' Text7.Enabled = dum
 ' Text8.Enabled = dum
  Text9.Enabled = DUM
  Text10.Enabled = DUM
  
End Function
Private Sub axrhsto_cmdUpdate_Click()
' κανω disable τα κουτάκια για να μην τα πειράζουν κατά λάθος
If cmdUpdate.Caption = mL_CapRes(106) Then ' cmdUpdate.Caption = "Διόρθωση"
   Group_Enable (True)
   cmdUpdate.Caption = mL_CapRes(103) ' "ενημέρωση"
   cmdClose.Enabled = False
   cmdAdd.Enabled = False
   cmdDelete.Enabled = False
Else
   cmdUpdate.Caption = mL_CapRes(106)  ' "Διόρθωση"
   cmdClose.Enabled = True
   Group_Enable (False)
   cmdAdd.Enabled = True
   cmdDelete.Enabled = True
  
  If IsNull(txtFields(0)) Or Val(txtFields(0)) = 0 Then
       MsgBox mL_Res(150), , "Talos"
      'MsgBox "Δεν συμπληρώσατε τον κωδικό και η ενημέρωση ακυρώνεται"
      'MsgBox "Δεν συμπληρώσατε τον κωδικό και η ενημέρωση ακυρώνεται"
      Exit Sub
 End If
  If IsNull(txtFields(1)) Or Len(Trim(txtFields(1))) = 0 Then
       
       MsgBox mL_Res(151), , "Talos"
      'MsgBox "Δεν συμπληρώσατε την περιγραφή  και η ενημέρωση ακυρώνεται"
      'MsgBox "Δεν συμπληρώσατε την περιγραφή  και η ενημέρωση ακυρώνεται"
      Exit Sub
 End If
 If IsNull(DBCombo1.text) Or (Val(left(DBCombo1.text, 3)) = 0 And InStr(1, DBCombo1.text, "POWDER") = 0) Then
       MsgBox mL_Res(152), , "Talos"
      'MsgBox "Δεν συμπληρώσατε την συγκέντρωση και η ενημέρωση ακυρώνεται"
      'MsgBox "Δεν συμπληρώσατε την συγκέντρωση και η ενημέρωση ακυρώνεται"
      Exit Sub
 End If
  If InStr(1, DBCombo1, "POWDER") > 0 Then
     txtFields(3) = 0
     Text5 = 0
     Text6 = 0
     Text7 = 0
     Text8 = 0
     Text9 = 0
     Text10 = 0
  End If
  
  
  
  
  Data1.UpdateRecord
  Data1.Recordset.Bookmark = Data1.Recordset.LastModified
End If
End Sub

Sub notes()
'στο show_dial  14/4/2002
' οταν άλλαζες διάλυμα-σκόνη δεν φαίνονταν
' REYST1.Visible = True
'   Reyst2.Visible = True

End Sub

Function show_Dial(DUM)
'-------------------------------------
Dim D

Data4.Recordset.MoveFirst
Do While Not Data4.Recordset.eof
         D = D + 1
         dbCombo2.AddItem Data4.Recordset("onoma")
         XX(D) = Data4.Recordset("onoma")
    Data4.Recordset.MoveNext
Loop

Data5.Recordset.MoveFirst
Do While Not Data5.Recordset.eof
        dbCombo3.AddItem Data5.Recordset("onoma")
    Data5.Recordset.MoveNext
Loop




Data4.Recordset.MoveFirst
Do While Not Data4.Recordset.eof
    If Data4.Recordset("kod") = Data1.Recordset("thermo1") Then
        dbCombo2 = Data4.Recordset("onoma")
        Exit Do
    End If
    Data4.Recordset.MoveNext
Loop


Data5.Recordset.MoveFirst
Do While Not Data5.Recordset.eof
    If Data5.Recordset("kod") = Data1.Recordset("thermo2") Then
        dbCombo3 = Data5.Recordset("onoma")
        Exit Do
    End If
    Data5.Recordset.MoveNext
Loop





If DUM = False Then
   lblLabels(3).Visible = False
   DBCombo1.Visible = False
   lblLabels(4).Visible = False
   txtFields(4).Visible = False
   dbCombo2.Visible = False
   dbCombo3.Visible = False
   Reyst1.Visible = False
   Reyst2.Visible = False
   
Else
   lblLabels(3).Visible = True
   DBCombo1.Visible = True
   dbCombo2.Visible = True
   dbCombo3.Visible = True
   Reyst1.Visible = True
   Reyst2.Visible = True

   
   lblLabels(4).Visible = True
   txtFields(4).Visible = True
   
End If



Text5.Visible = DUM
Text6.Visible = DUM
'Text7.Visible = dum
'Text8.Visible = dum
Text9.Visible = DUM
Text10.Visible = DUM

lblLabels(2).Visible = DUM
lblLabels(9).Visible = DUM
lblLabels(10).Visible = DUM
lblLabels(11).Visible = DUM

Label5.Visible = DUM
Label6.Visible = DUM
'Label7.Visible = dum
'Label4.Visible = dum
'Shape2.Visible = dum


End Function

Private Sub cmdClose_Click()

Dim merr, mtops(1 To 20), kato, mydb, MtTable
Dim b As String, problem, ll
problem = 0



If frmXHM.cmdAdd.Enabled = False And Option2(1).Enabled = True Then 'add rec
   For k = 1 To 5
       If Option2(k) Then
          ll = k
       End If
   Next
   Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
   Set MyTable = mydb.OpenRecordset("select *from ximitech where timh=" + str(ll), dbOpenDynaset)
   If MyTable.RecordCount > 0 Then
       MsgBox mLResnew(1181, "Allready exist*", 1)
       problem = 1
   End If
   mydb.Close
End If

On Error Resume Next
Data3.Recordset.MoveLast

'ελέγχω αν είναι στερεό μήπως έχω δώσει ξανά το ίδιο ColorID και σε άλλη σκόνη
If Data3.Recordset.RecordCount > 1 And Option1(1) Then
     Data3.Recordset.MoveFirst
     Do While Not Data3.Recordset.eof
           If txtFields9(9) = Data3.Recordset("kataskeyas") Then
               If Data3.Recordset("skonh") = 1 And Data3.Recordset("kod") <> Data1.Recordset("kod") Then
                      MsgBox "176-177." + mL_CapRes(176) + Data3.Recordset("perigrafh") + mL_CapRes(177)
                      '176 "Υπάρχει ξανά η ίδια ταυτότητα στό  "
                      '177 " Δώσε άλλο Color ID. "
                      problem = 1
                      Exit Do
               End If
            End If
            Data3.Recordset.MoveNext
     Loop
End If

On Error GoTo 0


 If Val(Text5.text) = 0 And Option1(1) = False Then
       MsgBox mL_Res(178) ' 178 "Δεν συμπληρώσατε το πεδίο χρόνος διάλυσης."
      
       problem = 1
 End If

 If Val(Text6.text) = 0 And Option1(1) = False Then
       MsgBox mL_Res(179) ' 179 "Δεν συμπληρώσατε το πεδίο χρόνος τελικής ανάμειξης."
      
       problem = 1
 End If
 If Val(Text10.text) = 0 And Option1(1) = False Then
       MsgBox mL_Res(180)
       ' 180 "Δεν συμπληρώσατε το πεδίο νερό διάλυσης."
     
       problem = 1
 End If
 
 If Len(Trim(dbCombo2)) = 0 And Option1(1) = False Then
       MsgBox mL_Res(181)
       ' 181 "Δεν συμπληρώσατε το πεδίο είδος νερού διάλυσης."
      'MsgBox "Δεν συμπληρώσατε την περιγραφή  και η ενημέρωση ακυρώνεται"
       problem = 1
 End If
 If Len(Trim(dbCombo3)) = 0 And Option1(1) = False Then
       MsgBox mL_Res(182)
       ' 182 "Δεν συμπληρώσατε το πεδίο είδος νερού αραίωσης."
      'MsgBox "Δεν συμπληρώσατε την περιγραφή  και η ενημέρωση ακυρώνεται"
       problem = 1
 End If






If IsNull(txtFields(0)) Or Val(txtFields(0)) = 0 Then
       MsgBox mL_Res(150), , "Talos"
      'MsgBox "Δεν συμπληρώσατε τον κωδικό και η ενημέρωση ακυρώνεται"
       problem = 1
 End If
  
 If IsNull(txtFields(1)) Or Len(Trim(txtFields(1))) = 0 Then
       MsgBox mL_Res(151), , "Talos"
      'MsgBox "Δεν συμπληρώσατε την περιγραφή  και η ενημέρωση ακυρώνεται"
       problem = 1
 End If
  If IsNull(txtFields9(9)) Or Len(Trim(txtFields9(9))) = 0 Then
       MsgBox mL_Res(183)
       ' 183 "Δεν συμπληρώσατε την ταυτότητα"
      
       problem = 1
 End If
 
  If Val(txtFields(8)) = 0 Then
       MsgBox mL_Res(184) '184 "Δεν συμπληρώσατε την ποσότητα ανανέωσης"
      
       problem = 1
 End If
  
  If Val(txtFields(2)) = 0 Then
       MsgBox mL_Res(185)
       '185 "Δεν συμπληρώσατε τις ημέρες ζωής."
      
       problem = 1
 End If
 
 
 
 
 If (IsNull(DBCombo1.text) Or Val(left(DBCombo1.text, 3)) = 0) And Option1(0) = True Then
       MsgBox mL_Res(152), , "Talos"
      'MsgBox "Δεν συμπληρώσατε την συγκέντρωση και η ενημέρωση ακυρώνεται"
      'MsgBox "Δεν συμπληρώσατε την συγκέντρωση και η ενημέρωση ακυρώνεται"
      problem = 1
 End If
 
               
 If Data1.Recordset("perigrafh") <> txtFields(1) Then ' 6
               Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
               Set MyTable = mydb.OpenTable("XIMITECH")
               MyTable.Index = "per"
               MyTable.Seek "=", txtFields(1)
               If Not MyTable.NoMatch Then '7
                  MsgBox mLResnew(908, "*Not allowed the same description two times.", 1)
                  problem = 1
               End If '7
 End If
 
 
 
 
 b = Data1.Recordset("kod")
            
If problem = 1 Then
   response = MsgBox(mL_Res(186), vbYesNo)
   ' 186 "Διορθώνω τις λάθη ή ελλείψεις;"
   If response = vbYes Then
      Exit Sub
   Else
       MsgBox mL_Res(187)
       '187  "Οι αλλαγές δεν καταχωρήθηκαν λόγω ελλειπών στοιχείων"
       frmXHM.Data1.Refresh
       frmXHM.cmdUpdate.Enabled = True
       frmXHM.cmdAdd.Enabled = True
       MDIForm1.Arxeia(10).Visible = False
      MDIForm1.Syntages(20).Visible = False
      MDIForm1.bohuhtika(30).Visible = False
       MDIForm1.ejodos(40).Visible = False
       Unload Me
       Exit Sub
   End If
End If



response = MsgBox(mL_Res(188), vbYesNo)
' 188 "Να καταχωρηθούν οι αλλαγές;"
' On Error Resume Next
If response = vbYes Then
    
    'DBCombo2.DataSource = "" 'Enabled = False
    'DBCombo3.DataSource = "" 'Enabled = False
    
    For k = 1 To 3
       If dbCombo2.text = XX(k) Then
            Data1.Recordset.EDIT
            Data1.Recordset("THERMO1") = k
            Data1.Recordset.update
            
       End If
    Next
    For k = 1 To 3
       If dbCombo3.text = XX(k) Then
            Data1.Recordset.EDIT
            Data1.Recordset("THERMO2") = k
            Data1.Recordset.update
       End If
    Next
    
    
    
    
    
    
    If frmXHM.cmdAdd.Enabled = False Then
    
            Data1.Recordset.EDIT
            Data1.Recordset.update
            Data1.Recordset.MoveFirst
            Data3.Recordset.AddNew
    
    Else
           Data1.Recordset.EDIT
           Data1.Recordset.update
           
           Data3.RecordSource = "SELECT *FROM XIMITECH WHERE KOD='" + Data1.Recordset("kod") + "';"
           Data3.Refresh
           
           Data3.Recordset.EDIT
    End If
    load_fields
End If
    
    frmXHM.Data1.Refresh
    frmXHM.cmdUpdate.Enabled = True
    frmXHM.cmdAdd.Enabled = True
    Unload Me
 

End Sub


Private Sub Command2_Click()
    
End Sub

Private Sub Command3_Click()
   CrystalReport1.Action = 1
End Sub



Private Sub Form_DblClick()
 Dim a
      'CommonDialog1.ShowColor
      'Me.BackColor = CommonDialog1.Color
      a = save_colors(Me, Me.BackColor)
End Sub

Private Sub Form_Load()
Dim asto
MDIForm1.Arxeia(10).Visible = False
MDIForm1.Syntages(20).Visible = False
MDIForm1.bohuhtika(30).Visible = False
MDIForm1.ejodos(40).Visible = False
Group_Enable (True)
proth_fora = 1


If frmXHM.cmdAdd.Enabled = False Then
 
   Label1.Caption = mL_CapRes(108) '"Εγγραφή νέου χημικοτεχνικού"
   If (IsNull(Xim1_on) Or IsEmpty(Xim1_on)) Then Option2(1).Enabled = False Else Option2(1).Enabled = True
   If (IsNull(Xim2_on) Or IsEmpty(Xim2_on)) Then Option2(2).Enabled = False Else Option2(2).Enabled = True
   If (IsNull(Xim3_on) Or IsEmpty(Xim3_on)) Then Option2(3).Enabled = False Else Option2(3).Enabled = True
   If (IsNull(Xim4_on) Or IsEmpty(Xim4_on)) Then Option2(4).Enabled = False Else Option2(4).Enabled = True
   If (IsNull(Xim5_on) Or IsEmpty(Xim5_on)) Then Option2(5).Enabled = False Else Option2(5).Enabled = True
   
ElseIf frmXHM.cmdUpdate.Enabled = False Then
    Label1.Caption = mL_CapRes(106) ' "Διόρθωση χημικοτεχνικού"
    Option1(1).Enabled = False
    Option1(0).Enabled = False
    For k = 1 To 4
        Option2(k).Enabled = False
    Next
    DBCombo1.Enabled = False
    Frame2.Enabled = False
End If

a = find_colors(Me)
 
 cmdClose.Caption = mL_CapRes(105)  ' "Εξοδος"
 
 
 Label6.Caption = mL_CapRes(115) '   "Αρχικό νερό διάλυσης (gr)"
 lblLabels(13).Caption = mL_CapRes(116) '   "Ποσότητα ανανέωσ(gr)"
 lblLabels(12).Caption = mL_CapRes(117) '   "Μέρες Ζωής"
 Label5.Caption = mL_CapRes(118) '  "Βοηθητικά Χημικά (mgr)"

 lblLabels(11).Caption = mL_CapRes(120) '   "Είδος νερού διάλυσης"
 lblLabels(10).Caption = mL_CapRes(121) '   "Είδος νερού αραίωσης"
 lblLabels(9).Caption = mL_CapRes(122) '   "Χρόνος τελ. ανάμειξης sec"
 lblLabels(2).Caption = mL_CapRes(123) '   "Χρόνος διάλυσης σε sec"

 
 lblLabels(7).Caption = mL_CapRes(127) '   "Διεύθυνση"
 lblLabels(6).Caption = mL_CapRes(128) '   "Υπόλοιπο (gr)"
 lblLabels(5).Caption = mL_CapRes(129) '   "Ημερομηνία Λήξης"

 lblLabels(4).Caption = mL_CapRes(130) '  "Πραγμ.συγκέντρωση"
 lblLabels(3).Caption = mL_CapRes(131) '  "Συγκέντρωση  (w/w)"
 lblLabels(1).Caption = mL_CapRes(132) '  "Περιγραφή"
 lblLabels(0).Caption = mL_CapRes(133) '  "Κωδικός"
          
 lblLabels(14).Caption = mL_CapRes(172) '
 
 
 
 
 'Label1.caption = ml_CapRes(173)
 
 Option1(0).Caption = mL_CapRes(174)
 Option1(1).Caption = mL_CapRes(175)
 
 Option2(1).Caption = mLResnew(927, "chem1*", 0)
 Option2(2).Caption = mLResnew(928, "Chem2*", 0)
 Option2(3).Caption = mLResnew(929, "Chem3*", 0)
 Option2(4).Caption = mLResnew(930, "Chem4*", 0)
 Option2(5).Caption = mLResnew(931, "Chem5*", 0)
 
End Sub

Private Sub Form_Paint()
  
  Screen.MousePointer = 1
If proth_fora = 1 Then
          mtxtFields2 = txtFields(2)
          txtFields8 = txtFields(8)
          txtFields(4).text = Format(Data1.Recordset("real_cons"), "#0.#####")
   proth_fora = 2
  If Data1.Recordset("skonh") = 1 Then
      Option1(1) = True
  Else
      Option1(0) = True
      If dbCombo2 = mL_CapRes(124) Or dbCombo2 = mL_CapRes(125) Or dbCombo2 = mL_CapRes(126) Then
         ' δεν χρειάζονται μετάφραση τα κρυο,ζεστό,χλιαρό
      Else
         translate_kryo
      End If
  End If
  
  If Not IsNull(Data1.Recordset("timh")) Then
      If Data1.Recordset("timh") > 0 And Data1.Recordset("timh") <= 5 Then
         Option2(Data1.Recordset("timh")) = True
      End If
  End If
  
 End If
 
 
 
   
End Sub
Sub translate_kryo()
L = Data4.Recordset("onoma").Size
        
Data4.Recordset.MoveFirst

        Data4.Recordset.EDIT
        Data4.Recordset("onoma") = left(mL_CapRes(124), L)
        Data4.Recordset.update
        
        Data4.Recordset.MoveNext
        Data4.Recordset.EDIT
        Data4.Recordset("onoma") = left(mL_CapRes(125), L)
        Data4.Recordset.update
        
        Data4.Recordset.MoveNext
        Data4.Recordset.EDIT
        Data4.Recordset("onoma") = left(mL_CapRes(126), L)
        Data4.Recordset.update
End Sub

Private Sub Option1_Click(Index As Integer)
   If Option1(0) Then   'διαλυμα
       show_Dial (True)
   Else
       show_Dial (False)
   End If
     If frmXHM.cmdAdd.Enabled = False And Option1(1) Then    ' nea eggrafh kai skonh
         If mtxtFields2 = txtFields(2) Then
              txtFields(2) = 180
        End If
        If txtFields8 = txtFields(8) Then
              txtFields(8) = 200
        End If
     End If
End Sub

Private Sub Text10_LostFocus()
 If Val(Text10) < 20 Or Val(Text10) > 200 Then
         MsgBox mL_Res(155), , "Talos"
        ' MsgBox "Απαράδεκτη ποσότητα νερού διάλυσης...", , "Talos"
        ' MsgBox "Απαράδεκτη ποσότητα νερού διάλυσης...", , "Talos"
         Text10 = 190
    End If
End Sub

Private Sub Text3_LostFocus()
     Text3.text = Val(Text3.text)
End Sub

Private Sub Text5_LostFocus()
 If Val(Text5) < 20 Or Val(Text5) > 500 Then
          MsgBox mL_Res(156), , "Talos"
         'MsgBox "Απαράδεκτος χρόνος διάλυσης...", , "Talos"
         'MsgBox "Απαράδεκτος χρόνος διάλυσης...", , "Talos"
         Text5 = 60
    End If
End Sub

Private Sub Text6_LostFocus()
 If Val(Text6) < 15 Or Val(Text6) > 60 Then
         
         MsgBox mL_Res(160), , "Talos"
         Text6 = 60
    End If
End Sub


Private Sub Text8_LostFocus()
 If Val(Text8) > 3 Or Val(Text8) < 1 Then
         MsgBox mL_Res(161), , "Talos"
         '"Δεν υπάρχει τέτοιο νερό διάλυσης...", , "Talos"
         Text8 = 0
    End If
End Sub

Private Sub txtFields_Change(Index As Integer)
    If Index = 2 Or Index = 8 Then '1
       If Not IsNumeric(txtFields(Index)) Then '2
           MsgBox mL_Res(189) '189 "Mόνο αριθμοί επιτρέπονται"
           txtFields(Index) = Space(Len(txtFields(Index)))
        End If '2
    End If '1
    If Index = 1 Then ' onoma ximikoy     3
      
      If left$(txtFields(1), 1) = " " Then  '4
          MsgBox mL_Res(190)
          '190 "Δεν επιτρέπεται το πρώτο ψηφίο να είναι κενό"
          txtFields(1) = Mid(txtFields(1), 2, Data1.Recordset("perigrafh").Size)
      End If  '4
      
     
    
    
    End If '3
       
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
Dim L
     If Index = 2 Then
        If Val(txtFields(2)) > 180 Or Val(txtFields(2)) <= 0 Then
      
           MsgBox mL_Res(191) ' 191 "Ημέρες εκτός ορίων 1-180"    ' 158 res"
            txtFields(2) = 180
           'MsgBox "Δεν υπάρχει τέτοια διεύθυνση μπουκαλιού...", , "Talos"
           'MsgBox "Δεν υπάρχει τέτοια διεύθυνση μπουκαλιού...", , "Talos"
           End If
        End If
     If Index = 3 Then
        If Val(txtFields(3)) > Ar_Mpoykal Or Val(txtFields(3)) < 1 Then
      
           MsgBox mL_Res(158), , "Talos"
           'MsgBox "Δεν υπάρχει τέτοια διεύθυνση μπουκαλιού...", , "Talos"
           'MsgBox "Δεν υπάρχει τέτοια διεύθυνση μπουκαλιού...", , "Talos"
            txtFields(3) = 0
        End If
    End If
    
    If Index = 8 Then  ' ποσότητα ανανέωσης
        If Val(txtFields(8)) > 600 Or Val(txtFields(8)) < 100 Then
            
            MsgBox mL_Res(159), , "Talos"
           'MsgBox "Ποσότητα εκτός ορίων (250-600)...", , "Talos"
           'MsgBox "Ποσότητα εκτός ορίων (250-600)...", , "Talos"
             txtFields(8) = 500
        End If
    End If
    If Index = 4 Or Index = 6 Then
        txtFields(Index).text = Val(txtFields(Index).text)
     End If
    
    If Index = 1 Then
         L = Data1.Recordset("perigrafh").Size
         If Len(txtFields(1)) > L Then
              MsgBox "192-193." + mL_CapRes(192) + str(L) + mL_CapRes(193)
              '192 "Η περιγραφή δέχεται μέχρι "
              '193 "   χαρακτήρες"
              txtFields(1) = left(txtFields(1), L)
         End If
         
         If (Not IsNull(Data3.Recordset("perigrafh"))) And (Not IsNull(txtFields(1))) Then '5
          
            On Error GoTo nocurrent
            If Data3.Recordset("perigrafh") <> txtFields(1) Then ' 6
               Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
               Set MyTable = mydb.OpenTable("XIMITECH")
               MyTable.Index = "per"
               MyTable.Seek "=", txtFields(1)
               If Not MyTable.NoMatch Then '7
                  MsgBox mLResnew(908, "*Not allowed the same description two times.", 1)
               End If '7
            End If '6
            On Error GoTo 0
        End If '5
         
         
         
         
         
    End If
    
     Exit Sub
nocurrent:
Resume Next


End Sub

Private Sub txtFields9_LostFocus(Index As Integer)
Dim db
If IsNull(txtFields(1).text) Or txtFields(1).text = "" Then
   Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
   Set MyTable = mydb.OpenRecordset("select *from ximitech where kataskeyas='" + txtFields9(9).text + "' and skonh=1", dbOpenDynaset)
   If MyTable.RecordCount > 0 Then
       MyTable.MoveFirst
       txtFields(1).text = LCase(MyTable("perigrafh"))
   End If
   mydb.Close
End If

End Sub



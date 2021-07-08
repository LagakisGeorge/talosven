VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form userinput 
   ClientHeight    =   6975
   ClientLeft      =   1905
   ClientTop       =   1725
   ClientWidth     =   6690
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6975
   ScaleWidth      =   6690
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Height          =   390
      Left            =   1815
      LinkItem        =   "LABEL1"
      LinkTimeout     =   500
      LinkTopic       =   "MOVEMENT|FORM2"
      MaxLength       =   180
      TabIndex        =   3
      Top             =   6225
      Visible         =   0   'False
      Width           =   3285
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "Συνέχεια  >>>"
      DragIcon        =   "_USERINP.frx":0000
      Height          =   555
      Left            =   4770
      TabIndex        =   2
      Top             =   3915
      Width           =   1515
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\RECIPIES.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4215
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "userinout"
      Top             =   4815
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\RECIPIES.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   435
      Left            =   2355
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "users"
      Top             =   4785
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBCtls.DBList DBList1 
      Bindings        =   "_USERINP.frx":0442
      DataField       =   "kod"
      DataSource      =   "Data2"
      Height          =   1230
      Left            =   1965
      TabIndex        =   1
      Top             =   1995
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2170
      _Version        =   393216
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ListField       =   "user_name"
      BoundColumn     =   "kod"
   End
   Begin VB.Label Label2 
      Caption         =   "vers.5.06.6K4.v6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1590
      TabIndex        =   4
      Top             =   75
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Χειριστής"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1950
      TabIndex        =   0
      Top             =   1230
      Width           =   3030
   End
End
Attribute VB_Name = "userinput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim proth_fora
Private Sub Command1_Click()
     
 Data2.Recordset.MoveLast
 User_ID = Data2.Recordset("kod")

 If IsNull(User_ID) Or User_ID = 0 Then
          MsgBox mL_Res(2) '"Διάλεξε User(χειριστή)"
          Exit Sub
 End If

Data2.Recordset.EDIT
    Data2.Recordset("userIN1") = Now
Data2.Recordset.update

'Unload Me
userinput.Visible = False



MDIForm1.Caption = MDIForm1.Caption + str(User_ID)


MDIForm1.Arxeia(10).Visible = True
MDIForm1.Syntages(20).Visible = True
MDIForm1.bohuhtika(30).Visible = True
MDIForm1.ejodos(40).Visible = True
  



End Sub


Private Sub Form_Load()

MDIForm1.Arxeia(10).Visible = False
MDIForm1.Syntages(20).Visible = False
MDIForm1.bohuhtika(30).Visible = False
MDIForm1.ejodos(40).Visible = False
Label1.Caption = mL_CapRes(3)   '3 "Χειριστής"
Command1.Caption = mL_CapRes(4) '4  "Συνέχεια  >>>"
  proth_fora = 1
  
 mday = LTrim(str(Day(Date)))
On Error Resume Next
mystamp = FileDateTime("\backup\" + mday + ".zip") ' Returns "2/12/93 4:35:47 PM".
If DateDiff("d", mystamp, Date) > 0 Then
    MsgBox mLResnew(899, "Να πάρετε πρώτα αντίγραφα.Αδύνατη η συνέχεια", 1)
'    End
End If
On Error GoTo 0
  
  
  
  
End Sub


Private Sub Form_Paint()
 Dim fors
 If proth_fora = 1 Then
   proth_fora = 2
   Data2.Recordset.AddNew
   Data2.Recordset.update
   Data2.Recordset.MoveLast
   
   s = GetCurrentTime()
 '  For k = 1 To 100: Next
'   s = GetCurrentTime() - s
   ForNexts_Milsec = 1000 '10000000 / s
 End If

End Sub


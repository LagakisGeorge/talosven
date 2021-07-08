VERSION 4.00
Begin VB.Form Syn_Hist 
   Caption         =   "Δοσομετρημένες Συνταγές"
   ClientHeight    =   6390
   ClientLeft      =   840
   ClientTop       =   1545
   ClientWidth     =   8595
   ControlBox      =   0   'False
   Height          =   6795
   Icon            =   "SYN_HIST.frx":0000
   Left            =   780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   8595
   Top             =   1200
   Width           =   8715
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Sort"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2490
      TabIndex        =   3
      Top             =   -30
      Width           =   1365
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sort"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   570
      TabIndex        =   2
      Top             =   -30
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Εξοδος"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6750
      TabIndex        =   1
      Top             =   -30
      Width           =   1365
   End
   Begin VB.Data Data1 
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\COORD_TB.MDB"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   -15
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Syntages"
      Top             =   6645
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "SYN_HIST.frx":0442
      Height          =   6345
      Left            =   45
      OleObjectBlob   =   "SYN_HIST.frx":0450
      TabIndex        =   0
      Top             =   330
      Width           =   8070
   End
End
Attribute VB_Name = "Syn_Hist"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Data1.Recordset.Sort = "[Description]"
Data1.Refresh
DBGrid1.Refresh
End Sub

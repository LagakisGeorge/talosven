VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4755
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   7620
   Height          =   5160
   Left            =   1080
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   7620
   Top             =   1170
   Width           =   7740
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5640
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1680
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\COORD_TB.MDB"
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from [Συντεταγμενες]  order by [Νο_Τεμαχιου];"
      Top             =   4200
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "Μπουκάλια"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Text1_LostFocus()

Set r = Data1.Recordset
r.MoveFirst
Do While Not r.EOF
   If r("Νο_Τεμαχιου") = Val(Text1.Text) Then
      Combo1.AddItem r("Συντεταγμένη_Ζ_user")
      x = r("Συντεταγμένη_Ζ_user")
      Exit Do
   End If
   r.MoveNext
Loop

r.MoveFirst
Do While Not r.EOF
   If r("Νο_Τεμαχιου") > Val(Text1) And (r("Νο_Τεμαχιου") - Val(Text1)) Mod 20 = 0 And r("Νο_Τεμαχιου") < 221 Then
      Combo1.AddItem Str(r("Νο_Τεμαχιου")) + "*** " + Str(r("Συντεταγμένη_Ζ_user"))
      x = r("Συντεταγμένη_Ζ_user")
   End If
   r.MoveNext
Loop




End Sub


Private Sub Text2_LostFocus()

Data1.RecordSource = "select *from [Συντεταγμενες] where [ΑΑ_Item]=2 order by [Νο_Τεμαχιου];"
Data1.Refresh
Set r = Data1.Recordset
r.MoveFirst
Do While Not r.EOF
   If r("Νο_Τεμαχιου") = Val(Text1.Text) Then
      Combo1.AddItem r("Συντεταγμένη_Ζ_user")
      x = r("Συντεταγμένη_Ζ_user")
      Exit Do
   End If
   r.MoveNext
Loop

r.MoveFirst
Do While Not r.EOF
   If r("Νο_Τεμαχιου") > Val(Text2) And (r("Νο_Τεμαχιου") - Val(Text2)) < 10 And r("Νο_Τεμαχιου") < 221 Then
      Combo1.AddItem Str(r("Νο_Τεμαχιου")) + "*** " + Str(r("Συντεταγμένη_Ζ_user"))
      x = r("Συντεταγμένη_Ζ_user")
   End If
   r.MoveNext
Loop



End Sub



VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5940
   ClientLeft      =   3150
   ClientTop       =   1410
   ClientWidth     =   6690
   Height          =   6345
   Left            =   3090
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   6690
   Top             =   1065
   Width           =   6810
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   540
      Left            =   3105
      TabIndex        =   8
      Top             =   1425
      Width           =   2250
   End
   Begin VB.TextBox Text3 
      Height          =   360
      Left            =   4125
      TabIndex        =   5
      Text            =   "61"
      Top             =   510
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\KATOFLIA.MDB"
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "katofli"
      Top             =   5700
      Width           =   3240
   End
   Begin VB.TextBox Text2 
      Height          =   360
      Left            =   2580
      TabIndex        =   2
      Text            =   "600"
      Top             =   525
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   960
      TabIndex        =   1
      Text            =   "500"
      Top             =   540
      Width           =   1425
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   450
      TabIndex        =   0
      Top             =   1425
      Width           =   2025
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "KATOFLIA.frx":0000
      Height          =   3090
      Left            =   450
      OleObjectBlob   =   "KATOFLIA.frx":000E
      TabIndex        =   7
      Top             =   2640
      Width           =   5925
   End
   Begin VB.Label Label3 
      Caption         =   "Mpoyk;ali"
      Height          =   360
      Left            =   4050
      TabIndex        =   6
      Top             =   90
      Width           =   1380
   End
   Begin VB.Label Label2 
      Caption         =   "Παροχή"
      Height          =   360
      Left            =   2520
      TabIndex        =   4
      Top             =   105
      Width           =   1380
   End
   Begin VB.Label Label1 
      Caption         =   "σταθμη"
      Height          =   345
      Left            =   945
      TabIndex        =   3
      Top             =   105
      Width           =   1470
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Function katax_katofli(mpoykali, katofli, paroxh)
Set r = Data1.Recordset
r.Index = "katofli"
r.Seek "=", mpoykali, katofli

If r.NoMatch Then
   r.AddNew
   r("mpoykali") = mpoykali
   r("staumh") = Int(katofli)
   r("paroxh") = Int(paroxh)
   r("hme") = Now
   r.Update
Else
   r.Edit
   r("paroxh") = Int(paroxh)
   r("hme") = Now
   r.Update

End If
katax_katofli = 1
End Function

Private Function anaz_parox(mpoykali, katofli)
Set r = Data1.Recordset
r.Index = "katofli"
r.Seek "<=", mpoykali, katofli

If r.NoMatch Then
Else
   
   
   s1 = r("staumh")
   p1 = r("paroxh")
   
   r.MoveNext
   
   s2 = r("staumh")
   p2 = r("paroxh")
   
   
   
dx = (p2 - p1) / (s2 - s1)

   
   
   
   

End If
anaz_parox = p1 + (katofli - s1) * dx


End Function

Private Sub Command1_Click()
  a = katax_katofli(Text3.Text, Text1.Text, Text2.Text)
End Sub


Private Sub Command2_Click()
  a = anaz_parox(Text3.Text, Text1.Text)
End Sub



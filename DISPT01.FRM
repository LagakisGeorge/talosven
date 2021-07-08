VERSION 4.00
Begin VB.Form frmSynt 
   Caption         =   "For Next Small"
   ClientHeight    =   5295
   ClientLeft      =   780
   ClientTop       =   975
   ClientWidth     =   8880
   Height          =   5700
   Left            =   720
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5295
   ScaleWidth      =   8880
   Top             =   630
   Width           =   9000
   Begin VB.CommandButton Command4 
      Caption         =   "Gemisma arxeioy"
      Height          =   375
      Left            =   7200
      TabIndex        =   18
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Height          =   2010
      Left            =   6480
      TabIndex        =   17
      Top             =   4200
      Width           =   2835
   End
   Begin VB.CommandButton Command3 
      Caption         =   "times synartiseon"
      Height          =   375
      Left            =   6480
      TabIndex        =   16
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2400
      TabIndex        =   15
      Text            =   "550"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2400
      TabIndex        =   13
      Text            =   "70"
      Top             =   960
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   4350
      Left            =   120
      TabIndex        =   12
      Top             =   1710
      Width           =   4470
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   885
      Left            =   4800
      TabIndex        =   11
      Top             =   2880
      Width           =   1470
   End
   Begin VB.TextBox FornextSmall 
      Height          =   360
      Left            =   7185
      TabIndex        =   4
      Text            =   "1000"
      Top             =   570
      Width           =   2000
   End
   Begin VB.TextBox FornextBig 
      Height          =   360
      Left            =   7155
      TabIndex        =   3
      Text            =   "10000"
      Top             =   60
      Width           =   2000
   End
   Begin VB.TextBox Quantity 
      Height          =   360
      Left            =   2385
      TabIndex        =   2
      Text            =   "500"
      Top             =   105
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ƒÔÛÔÏÂÙÒﬁÛÂÈÚ"
      Height          =   1080
      Left            =   4800
      TabIndex        =   1
      Top             =   1680
      Width           =   1470
   End
   Begin VB.TextBox Times 
      Height          =   360
      Left            =   2385
      TabIndex        =   0
      Text            =   "2"
      Top             =   555
      Width           =   3135
   End
   Begin VB.Label Label6 
      Caption         =   "ÃÔıÍ‹ÎÈ ÕÔ"
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label8 
      Height          =   405
      Left            =   750
      TabIndex        =   10
      Top             =   3615
      Width           =   1620
   End
   Begin VB.Label Label5 
      Caption         =   "Level"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "ForNext Small"
      Height          =   345
      Left            =   5640
      TabIndex        =   8
      Top             =   630
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "ForNext Big"
      Height          =   345
      Left            =   5700
      TabIndex        =   7
      Top             =   75
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Times of Dispensing"
      Height          =   390
      Left            =   555
      TabIndex        =   6
      Top             =   495
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Quantity to Dispense"
      Height          =   390
      Left            =   450
      TabIndex        =   5
      Top             =   90
      Width           =   1575
   End
End
Attribute VB_Name = "frmSynt"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Command1_Click()

For k = 1 To Times
       frmSynt.Caption = k
       Command2.Enabled = True
       Fornext_big = FornextBig.Text
       Fornext_small = FornextSmall.Text
                A = GetCurrentTime()
                Syn_Dos.Asked_Q = Quantity.Text '   a_Mgrs(Move) '  ZHTOYMENH –œ”œTHTA ”E MGRS
                Syn_Dos.Level = Val(Text2.Text)   'Ypoloipo(a_Fere(Move))      '√—¡Ã¡—…¡ Ã–œ’ ¡À…œ’
                Syn_Dos.Valve_on = valve_1_on ' ¡Õœ…√Ã¡ ¬¡À¬…ƒ¡”
                Syn_Dos.Valve_off = valve_1_off     '  À≈…”…Ãœ ¬¡À¬…ƒ¡”
                Syn_Dos.Caption = Text1.Text 'tr(a_Fere(Move)) + " : " + Perigrafh(a_Fere(Move)) 'ÂÒÈ„Ò·ˆÁ ˜Ò˛Ï·ÙÔÚ Ôı ‰ÔÛÔÏÂÙÒÂﬂÙ·È
                Syn_Dos.Show 1
                
                Do
                  DoEvents
                
                Loop Until Label8 <> ""
                
                List1.AddItem "ı=" + Str(Int(Val(Text2.Text))) + " ‰ÔÛ=" + (Label8.Caption) + "sp=" + Str(Int(0.795 * Val(Text2.Text) + 216.8)) + " ir=" + Str(Int(0.665 * Val(Text2.Text) + 294.4)) + ":˜Ò=" + Str(Int((GetCurrentTime() - A) / 1000))
                Text2.Text = Val(Text2.Text) - Val(Label8.Caption) / 1000
                If Command2.Enabled = False Then Exit For
                
                
Next


End Sub


Private Sub Command2_Click()
 List1.Clear
End Sub


Private Sub Command3_Click()

A = (560.47 - 420.79) / (400 - 190)
b = 560.47 - A * 400



For k = 670 To 50 Step -10
   Y1 = 0.795 * k + 216.8 'spiti
   y = 0.665 * k + 294.4 'iris
   List2.AddItem Str(k) + "---" + Str(y) + "sp" + Str(Y1)
Next
End Sub


Private Sub Command4_Click()
 Set mydb = OpenDatabase("c:\talos\katoflia.mdb")
 mydb.Execute "delete *from katofli"
 Set r = mydb.OpenRecordset("katofli")
 
 For k = 61 To 220 Step 2
   For l = 620 To 50 Step -10
       r.AddNew
       r("mpoykali") = k
       r("staumh") = l
       r("paroxh") = 0.795 * l + 216.8
       r.Update
   Next l
Next k
mydb.Close
End Sub



VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6975
   ClientLeft      =   1020
   ClientTop       =   135
   ClientWidth     =   7950
   Height          =   7380
   Left            =   960
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   7950
   Top             =   -210
   Width           =   8070
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check2 
      Caption         =   "Νερό"
      Height          =   240
      Left            =   4485
      TabIndex        =   14
      Top             =   6060
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Από Ιστορικό αρχείο"
      Height          =   285
      Left            =   4485
      TabIndex        =   13
      Top             =   6300
      Width           =   1845
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Διαλυμ.(αποκλ. περιεκτικότητας)"
      Height          =   420
      Left            =   5715
      TabIndex        =   12
      Top             =   5760
      Width           =   2100
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Text            =   "0"
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2745
      TabIndex        =   8
      Text            =   "30000"
      Top             =   6015
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   5280
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2745
      TabIndex        =   4
      Top             =   5655
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "c:\talos\recipies.mdb"
      Exclusive       =   0   'False
      Height          =   615
      Left            =   -120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from prospau2;"
      Top             =   7800
      Width           =   1140
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Στατιστικά συνταγών"
      Height          =   405
      Left            =   4440
      TabIndex        =   2
      Top             =   5295
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Διαλυμ.(αποκλ.σκόνης)"
      Height          =   420
      Left            =   5700
      TabIndex        =   1
      Top             =   5295
      Width           =   2100
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\SKON_TB.MDB"
      Exclusive       =   0   'False
      Height          =   375
      Left            =   465
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT *FROM Διαλυματα"
      Top             =   8160
      Width           =   1845
   End
   Begin VB.Label Label4 
      Caption         =   "Από mgr"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Μέχρι mgr"
      Height          =   375
      Left            =   2265
      TabIndex        =   10
      Top             =   6015
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Μέχρι"
      Height          =   375
      Left            =   2265
      TabIndex        =   7
      Top             =   5655
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Από"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5640
      Width           =   495
   End
   Begin GraphLib.Graph G 
      Height          =   5055
      Left            =   -1560
      TabIndex        =   0
      Top             =   240
      Width           =   10965
      _version        =   65536
      _extentx        =   19341
      _extenty        =   8916
      _stockprops     =   96
      borderstyle     =   1
      background      =   3
      bottomtitle     =   "Απόκλιση % από στόχο"
      foreground      =   14
      graphtype       =   2
      indexstyle      =   1
      legendstyle     =   1
      palette         =   1
      printstyle      =   1
      randomdata      =   1
      colordata       =   0
      extradata       =   0
      extradata[]     =   0
      fontfamily      =   4
      fontsize        =   4
      fontsize[0]     =   100
      fontsize[1]     =   150
      fontsize[2]     =   100
      fontsize[3]     =   100
      fontstyle       =   4
      graphdata       =   0
      graphdata[]     =   0
      labeltext       =   0
      legendtext      =   0
      patterndata     =   0
      symboldata      =   0
      xposdata        =   0
      xposdata[]      =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim akr(1 To 11)

For k = 1 To 11: akr(k) = 0: Next

synt = ""
'SELECT *FROM Διαλυματα
If IsDate(Text1.Text) And IsDate(Text2.Text) Then
   hm2 = DateAdd("d", 1, Text2.Text)
   synt = "datum>=#" + Format(Text1.Text, "mm/dd/yyyy") + "# and datum<= #" + Format(hm2, "mm/dd/yyyy") + "#"
Else
   MsgBox "Οχι σωστές ημερομηνίες"
   Exit Sub
End If

Data1.RecordSource = "select *from Διαλυματα where " + synt
Data1.Refresh






Set r = Data1.Recordset

r.MoveFirst
Do While Not r.EOF
    apait = r("nerostox") * r("periek")
    If apait > 0 Then
         apokl = Abs((r("color_act") - apait) / apait * 10)
         If apokl > 10 Then
            akr(11) = akr(11) + 1
         Else
            akr(CInt(apokl) + 1) = akr(CInt(apokl) + 1) + 1
         End If
    End If
    r.MoveNext
Loop



G.NumSets = 1
G.NumPoints = 10

'For k = 1 To 10
'   G.ThisSet = 1
'   G.ThisPoint = k
'   G.GraphData = akr(k)
'Next

s = 0
For k = 1 To 10
   G.ThisSet = 1
   G.ThisPoint = k
   G.GraphData = akr(k)
   s = s + akr(k)
Next
Combo1.Clear
For k = 1 To 10
   Combo1.AddItem Format(10 * k, "###") + "%   " + Format(akr(k) / s * 100, "###.#") + "%"
Next







G.AutoInc = 1
For k = 1 To 10
   G.LabelText = 10 * k
Next
G.ThisSet = 1
' G.LegendText = "Ακρίβεια διαλυμάτων"
G.GraphStyle = 6
G.DrawMode = 2
End Sub

Private Sub Command2_Click()
Dim akr(1 To 11), r, synt

For k = 1 To 11: akr(k) = 0: Next

synt = ""
'SELECT *FROM Διαλυματα
If IsDate(Text1.Text) And IsDate(Text2.Text) Then
   hm2 = DateAdd("d", 1, Text2.Text)
   synt = "hme>=#" + Format(Text1.Text, "mm/dd/yyyy") + "# and hme<= #" + Format(hm2, "mm/dd/yyyy") + "#"
Else
   MsgBox "Οχι σωστές ημερομηνίες"
   Exit Sub
End If

If IsNumeric(Text3.Text) And IsNumeric(Text3.Text) Then
   synt = synt + " and gr>=" + Text4.Text + " and gr<= " + Text3.Text + ";"
Else
   MsgBox "Οχι σωστά mgr"
   Exit Sub
End If




If Check1 Then
    Data2.RecordSource = "select *from prospau3 where " + synt
    Data2.Refresh
Else
    Data2.RecordSource = "select *from prospau2 where " + synt
    Data2.Refresh
End If



x = Now
Do 'repeat
    DoEvents
Loop Until DateDiff("s", x, Now) > 2








Set r = Data2.Recordset



    

On Error GoTo den_exo
r.MoveFirst
On Error GoTo 0
a = 0
Do While Not r.EOF
  If check2 Then
     Real = r("renero")
     zht = r("nero")
  Else
     Real = r("regr")
     zht = r("gr")
  End If
  
  
  
  If Not IsNull(Real + zht) Then
  
    a = a + 1
    Form1.Caption = "Δείγματα " + Format(a, "###,###")
  
    apait = zht ' r("gr")
    If apait > 0 Then
         apokl = Abs((Real - apait) / apait * 100)
         
         If apokl > 9 Then
            akr(10) = akr(10) + 1
         Else
            akr(CInt(apokl) + 1) = akr(CInt(apokl) + 1) + 1
          End If
          
    End If
  End If
    r.MoveNext
Loop



G.NumSets = 1
G.NumPoints = 10
s = 0
For k = 1 To 10
   G.ThisSet = 1
   G.ThisPoint = k
   G.GraphData = akr(k)
   s = s + akr(k)
Next



Combo1.Clear
For k = 1 To 10
   Combo1.AddItem Format(k, "##") + "%   " + Format(akr(k) / s * 100, "###.#") + "%"
Next


G.AutoInc = 1

For k = 1 To 10
   G.LabelText = 1 * k
Next
G.ThisSet = 1
'G.LegendText = "Ακρίβεια διαλυμάτων"
G.GraphStyle = 6
G.DrawMode = 2

Exit Sub

den_exo:

If r.RecordCount < 2 Then
   MsgBox " δεν υπάρχουν εγγραφές"
   Exit Sub
End If

End Sub


Private Sub Command3_Click()
' περιεκτικοτητα
Dim akr(1 To 11)
For k = 1 To 11: akr(k) = 0: Next

synt = ""
'SELECT *FROM Διαλυματα
If IsDate(Text1.Text) And IsDate(Text2.Text) Then
   hm2 = DateAdd("d", 1, Text2.Text)
   synt = "datum>=#" + Format(Text1.Text, "mm/dd/yyyy") + "# and datum<= #" + Format(hm2, "mm/dd/yyyy") + "#"
Else
   MsgBox "Οχι σωστές ημερομηνίες"
   Exit Sub
End If



On Error Resume Next
If Check1 Then
   Data1.DatabaseName = "C:\paldisk\TAL2001\skon_tb.mdb"
Else
   Data1.DatabaseName = "c:\talos\skon_tb.mdb"
End If

Data1.RecordSource = "select *from Διαλυματα where " + synt
Data1.Refresh






Set r = Data1.Recordset
a = 0
r.MoveFirst
Do While Not r.EOF
    a = a + 1
    Form1.Caption = "Δείγματα " + Format(a, "###,###")
    ' If a > 10 Then Exit Do
    apait = r("periek")
    If apait > 0 Then
         apokl = Abs(apait - r("color_act") / (r("color_act") + r("nero_act"))) / apait * 100
         If apokl > 50 Then
            akr(10) = akr(10) + 1
         ElseIf apokl > 9 And apokl <= 50 Then
            akr(9) = akr(9) + 1
         Else
            akr(CInt(apokl) + 1) = akr(CInt(apokl) + 1) + 1
         End If
    End If
    r.MoveNext
Loop



G.NumSets = 1
G.NumPoints = 10

'For k = 1 To 10
'   G.ThisSet = 1
'   G.ThisPoint = k
'   G.GraphData = akr(k)
'Next

s = 0
For k = 1 To 10
   G.ThisSet = 1
   G.ThisPoint = k
   G.GraphData = akr(k)
   s = s + akr(k)
Next
Combo1.Clear
For k = 1 To 8
   Combo1.AddItem Format(1 * k, "###") + "%   " + Format(akr(k) / s * 100, "###.#") + "%"
Next
   Combo1.AddItem "9%-49 " + Format(akr(k) / s * 100, "###.#") + "%"
   Combo1.AddItem "50%-  " + Format(akr(k) / s * 100, "###.#") + "%"







G.AutoInc = 1
For k = 1 To 8
   G.LabelText = Str(k) + "%"
Next
G.LabelText = "9%-49%"
G.LabelText = "50%-"

G.ThisSet = 1
' G.LegendText = "Ακρίβεια διαλυμάτων"
G.GraphStyle = 6
G.DrawMode = 2
End Sub

Private Sub Form_Load()
Text1.Text = Format(Now - 30, "dd/mm/yyyy")
Text2.Text = Format(Now, "dd/mm/yyyy")

End Sub




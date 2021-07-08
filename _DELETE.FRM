VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   5940
   ClientLeft      =   -2100
   ClientTop       =   1440
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5940
   ScaleWidth      =   9660
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "Μηδενισμός  αρχείου καταγραφής κινήσεων"
      Height          =   345
      Left            =   3015
      TabIndex        =   5
      Top             =   2640
      Width           =   5805
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Διαγραφή Χημικοτεχνικών"
      Height          =   345
      Left            =   3015
      TabIndex        =   4
      Top             =   1860
      Width           =   5805
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Εξοδος"
      Height          =   840
      Left            =   7200
      TabIndex        =   3
      Top             =   4635
      Width           =   1740
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   6960
      TabIndex        =   1
      Top             =   465
      Width           =   1845
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Διαγραφή συνταγών και εργασιών"
      Height          =   345
      Left            =   3000
      TabIndex        =   0
      Top             =   960
      Width           =   5805
   End
   Begin VB.Label Label1 
      Caption         =   "Nα σβηστούν οι συνταγές μέχρι και την :"
      Height          =   285
      Left            =   3030
      TabIndex        =   2
      Top             =   495
      Width           =   3840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
'******
If prosbash() <> "ok" Then
   MsgBox mL_Res(800) '"ΔΕΝ ΕΧΕΤΕ ΠΡΟΣΒΑΣΗ "
   Exit Sub
End If
   
If Val(Text1.text) = 0 Then
  response = MsgBox(mL_Res(801), 20, "talos")   '"ΔΕΝ ΔΗΛΩΣΑΤΕ ΜΕΧΡΙ ΠΟΙΟ ΑΡΙΘΜΟ ΣΥΝΤΑΓΗΣ ΝΑ ΣΒΗΣΩ."
  Exit Sub
Else
 response = MsgBox(mL_Res(802), 20, "talos") ' "ΝΑ ΣΒΗΣΤΟΥΝ ΟΛΕΣ ΟΙ ΣΥΝΤΑΓΕΣ"
 If response = 7 Then ' akyro
     Exit Sub
 End If
End If

Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)


Set MyTableDef = mydb.TableDefs("Joblist3")
' Create new Field object.
Set MyField = MyTableDef.Fields("kod_erg")
' Set another property of MyField.
MyField.AllowZeroLength = True   ' Greater than longest phone number.

Set MyField = MyTableDef.Fields("parat")
' Set another property of MyField.
MyField.AllowZeroLength = True   ' Greater than longest phone number.


ans = MsgBox("Delete old Historical File?", vbYesNo)
If ans = vbNo Then
   sql_del = "insert into joblist3  select [joblist].*  from [joblist] where entolh<='" + Right("000000" + Text1.text, 6) + "';"
   mydb.Execute sql_del
   sql_del = "DELETE DISTINCTROW JOBLIST.* FROM JOBLIST where entolh<='" + Right("000000" + Text1.text, 6) + "'"
   mydb.Execute sql_del
   mydb.Close
   
   Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
   Set MySet = mydb.OpenRecordset("PROSPAU2")
   sql_del = "insert into prospau3  select [prospau2].*  from [prospau2] where entolh<='" + Right("000000" + Text1.text, 6) + "';"
   mydb.Execute sql_del
   sql_del = "DELETE DISTINCTROW PROSPAU2.* FROM PROSPAU2 where entolh<='" + Right("000000" + Text1.text, 6) + "';"
   mydb.Execute sql_del
   mydb.Close
Else
   mydb.Execute "drop table joblist3;"
   mydb.Execute "drop table prospau3;"
   
   sql_del = "select [joblist].* into joblist3  from [joblist] where entolh<='" + Right("000000" + Text1.text, 6) + "';"
   mydb.Execute sql_del
   
   sql_del = "DELETE DISTINCTROW JOBLIST.* FROM JOBLIST where entolh<='" + Right("000000" + Text1.text, 6) + "'"
   mydb.Execute sql_del
   
   sql_del = "select [prospau2].*  into prospau3  from [prospau2] where entolh<='" + Right("000000" + Text1.text, 6) + "';"
   mydb.Execute sql_del
   
   sql_del = "DELETE DISTINCTROW PROSPAU2.* FROM PROSPAU2 where entolh<='" + Right("000000" + Text1.text, 6) + "';"
   mydb.Execute sql_del
   
   mydb.Close


End If

MsgBox mL_Res(803) ' "θα γίνει αναδιοργάνωση των αρχείων"
      ' m_err = "reindex.exe,c:\talos"
 Reindexing (0)

MsgBox mL_Res(804) '  "ΕΧΟΥΝ ΣΒΗΣΤΕΙ ΟΙ ΣΥΝΤΑΓΕΣ ΠΟΥ ΔΗΛΩΣΑΤΕ"

End Sub


Private Function prosbash()
Dim default, title, myvalue
prosbash = "no"
Message = mL_Res(812) ' "Δώστε κωδικό πρόσβασης "   ' Set prompt.
title = mL_Res(813) ' "Εργασία κρίσιμη" ' Set title.
default = "0000"   ' Set default.
' Display message, title, and default value.
myvalue = InputBox(Message, title, default)
If myvalue = "3631" Then prosbash = "ok"

End Function


Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Command3_Click()

If prosbash <> "ok" Then
   MsgBox mL_Res(805) ' "ΔΕΝ ΕΧΕΤΕ ΠΡΟΣΒΑΣΗ "
   Exit Sub
End If
   
 response = MsgBox(mL_Res(806), 20, "talos") ' ("ΝΑ ΣΒΗΣΤΟΥΝ ΟΛΑ ΤΑ ΧΗΜΙΚΟΤΕΧΝΙΚΑ;", 20, "talos")
 If response = 7 Then ' akyro
     Exit Sub
 End If


Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
Set MySet = mydb.OpenRecordset("XIMITECH")
sql_del = "DELETE DISTINCTROW XIMITECH.* FROM XIMITECH"
mydb.Execute sql_del
mydb.Close
MsgBox mL_Res(807) ' "θα γίνει αναδιοργάνωση των αρχείων"
    'm_err = "reindex.exe,c:\talos"
Reindexing (0)
MsgBox mL_Res(808) ' "ΕΧΟΥΝ ΣΒΗΣΤΕΙ ΟΛΑ ΤΑ ΧΗΜΙΚΟΤΕΧΝΙΚΑ."

End Sub


Private Sub Command4_Click()


If prosbash <> "ok" Then
   MsgBox mL_Res(809) '"ΔΕΝ ΕΧΕΤΕ ΠΡΟΣΒΑΣΗ "
   Exit Sub
End If
   
 response = MsgBox(mL_Res(810), 20, "talos")
 If response = 7 Then ' akyro
     Exit Sub
 End If


 Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\watch.mdb")
             Set r = dbsjet.OpenRecordset("moves")
sql_del = "DELETE MOVES.* FROM MOVES"
dbsjet.Execute sql_del
dbsjet.Close
   
MsgBox mL_Res(811) '"ΕΧEI ΜΗΔΕΝΙΣΤΕΙ ΤΟ ΑΡΧΕΙΟ ΚΑΤΑΓΡΑΦΗΣ ΚΙΝΗΣΕΩΝ."

End Sub


Private Sub Form_Load()
Label1.Caption = mL_CapRes(814) '  Nα σβηστούν οι συνταγές μέχρι και την :"

Command1.Caption = mL_CapRes(815) ' Διαγραφή συνταγών και εργασιών

Command3.Caption = mL_CapRes(816) 'Διαγραφή Χημικοτεχνικών
Command4.Caption = mL_CapRes(817) 'Μηδενισμός  αρχείου καταγραφής κινήσεων
Command2.Caption = mL_CapRes(818) 'Διαγραφή Χημικοτεχνικών
End Sub


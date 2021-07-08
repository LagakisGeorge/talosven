VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form mJOBLIST 
   BackColor       =   &H00C0FFFF&
   ClientHeight    =   6855
   ClientLeft      =   990
   ClientTop       =   420
   ClientWidth     =   10800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6855
   ScaleWidth      =   10800
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text10 
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   1245
      TabIndex        =   37
      Top             =   2610
      Width           =   1125
   End
   Begin VB.TextBox Text9 
      Height          =   300
      Left            =   8340
      TabIndex        =   8
      Top             =   495
      Width           =   465
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFF00&
      Height          =   315
      Left            =   75
      TabIndex        =   34
      Top             =   2625
      Width           =   1125
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   5400
      MaxLength       =   20
      TabIndex        =   5
      Top             =   75
      Width           =   735
   End
   Begin VB.Data prospau2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\recipies.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   360
      Left            =   8175
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "PROSPAU2"
      Top             =   7195
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Παρατηρήσεις"
      Height          =   300
      Left            =   120
      TabIndex        =   31
      Top             =   6150
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Ημερ.Παράδοσης"
      Height          =   300
      Left            =   120
      TabIndex        =   30
      Top             =   5700
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Κατά Εντολή"
      Height          =   300
      Left            =   120
      TabIndex        =   29
      Top             =   5250
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   5400
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   1875
      Width           =   2250
   End
   Begin VB.Data pania 
      Caption         =   "pania"
      Connect         =   "Access"
      DatabaseName    =   "c:\talos\recipies.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   465
      Left            =   2625
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT *FROM pania ORDER BY PANI"
      Top             =   5055
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Κατασκευή Συνταγής"
      DragIcon        =   "_JOBLIST.frx":0000
      Height          =   615
      Left            =   6795
      TabIndex        =   28
      Top             =   2300
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ΕΞΟΔΟΣ"
      Height          =   615
      Left            =   8175
      TabIndex        =   27
      Top             =   2300
      Width           =   1335
   End
   Begin VB.Data EIDH_BAF 
      Caption         =   "EIDH_BAFHS"
      Connect         =   "Access"
      DatabaseName    =   "c:\talos\recipies.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6315
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT *FROM EIDH_BAF ORDER BY ONO_BAF"
      Top             =   5475
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data pelates 
      Caption         =   "pelates"
      Connect         =   "Access"
      DatabaseName    =   "c:\talos\recipies.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   4440
      Options         =   0
      ReadOnly        =   -1  'True
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select * from pelates order by onoma;"
      Top             =   5505
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Διαγραφή"
      Height          =   615
      Left            =   5400
      TabIndex        =   25
      Top             =   2300
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Διόρθωση"
      Height          =   615
      Left            =   4020
      TabIndex        =   13
      Top             =   2300
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nέα Εργασία"
      Height          =   615
      Left            =   2625
      TabIndex        =   12
      Top             =   2300
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      Height          =   300
      Left            =   5400
      MaxLength       =   10
      TabIndex        =   10
      Top             =   1425
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      Height          =   300
      Left            =   5400
      MaxLength       =   10
      TabIndex        =   9
      Top             =   975
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   300
      Left            =   5400
      MaxLength       =   14
      TabIndex        =   7
      Top             =   525
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   300
      Left            =   8325
      TabIndex        =   6
      Text            =   "8.5"
      Top             =   45
      Width           =   465
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   1455
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1425
      Width           =   1815
   End
   Begin VB.Data Joblist 
      Caption         =   "JOBLIST"
      Connect         =   "Access"
      DatabaseName    =   "c:\talos\recipies.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   420
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "joblist"
      Top             =   4815
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "_JOBLIST.frx":0442
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Top             =   1845
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      ListField       =   "ONO_BAF"
      BoundColumn     =   "ONO_BAF"
      Text            =   "DBCombo2"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3555
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1200
      Top             =   5850
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "_JOBLIST.frx":0459
      Height          =   315
      Left            =   1455
      TabIndex        =   2
      Top             =   975
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      ListField       =   "PANI"
      BoundColumn     =   "PANI"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "_JOBLIST.frx":046D
      Height          =   315
      Left            =   1455
      TabIndex        =   1
      Top             =   525
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      ListField       =   "ONOMA"
      BoundColumn     =   "ONOMA"
      Text            =   "DBCombo1"
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "_JOBLIST.frx":0483
      Height          =   3465
      Left            =   -120
      OleObjectBlob   =   "_JOBLIST.frx":0499
      TabIndex        =   0
      Top             =   3000
      Width           =   9720
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Chemicals by Hand"
      Height          =   315
      Left            =   6585
      TabIndex        =   36
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Jump to task:"
      Height          =   300
      Left            =   120
      TabIndex        =   35
      Top             =   2310
      Width           =   2295
   End
   Begin VB.Label Label14 
      Caption         =   "Ταξινόμιση "
      Height          =   300
      Left            =   120
      TabIndex        =   33
      Top             =   4800
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Βάρος πανιού (mgr)"
      Height          =   300
      Left            =   3800
      TabIndex        =   32
      Top             =   75
      Width           =   1520
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Παρατηρήσεις"
      Height          =   300
      Left            =   3800
      TabIndex        =   26
      Top             =   1875
      Width           =   1520
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Ημερ.Παράδοσης"
      Height          =   300
      Left            =   3800
      TabIndex        =   24
      Top             =   1425
      Width           =   1520
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Ημερ.Παραγγελίας"
      Height          =   300
      Left            =   3800
      TabIndex        =   23
      Top             =   975
      Width           =   1520
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Αριθ.Παραγγελίας"
      Height          =   255
      Left            =   3800
      TabIndex        =   22
      Top             =   525
      Width           =   1520
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Μπάνιο 1:"
      Height          =   315
      Left            =   6975
      TabIndex        =   21
      Top             =   30
      Width           =   1305
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Διεργασία"
      Height          =   300
      Left            =   100
      TabIndex        =   20
      Top             =   1875
      Width           =   1200
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Απόχρωση"
      Height          =   300
      Left            =   100
      TabIndex        =   19
      Top             =   1425
      Width           =   1200
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Πανί"
      Height          =   300
      Left            =   100
      TabIndex        =   18
      Top             =   975
      Width           =   1200
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Πελάτης"
      Height          =   300
      Left            =   100
      TabIndex        =   17
      Top             =   525
      Width           =   1200
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   300
      Left            =   100
      TabIndex        =   16
      Top             =   75
      Width           =   1200
   End
   Begin VB.Label Label2 
      DataField       =   "ENTOLH"
      DataSource      =   "Data1"
      Height          =   300
      Left            =   1455
      TabIndex        =   15
      Top             =   75
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ΕΡΓΑΣΙΕΣ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   8040
      TabIndex        =   14
      Top             =   6195
      Visible         =   0   'False
      Width           =   3660
   End
End
Attribute VB_Name = "mJOBLIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_edit
Dim Caption_update, Caption_add, Caption_edit, book
Dim DEFAULT_MPANIO
Dim proth_fora
Private Sub Command3_Click()
Dim MyWorkspace As Workspace, mydb As Database

If JOBLIST.Recordset.eof Or JOBLIST.Recordset.BOF Then
    ' TIPOTA
Else
  ' If IsNull(JOBLIST.Recordset("ENTOLH")) Then
      '     ΓΙΑ ΝΑ ΜΗΝ ΤΙΣ ΣΒΗΣΕΙ ΟΛΕΣ
   ' Else
       'If Val(JOBLIST.Recordset("ENTOLH")) = 0 Then
       'Else
       If IsNull(JOBLIST.Recordset("ENTOLH")) Then
           ans = MsgBox("266-267." + mL_CapRes(266) + " " + mL_CapRes(267), vbYesNo)
       Else
          ans = MsgBox("266-267." + mL_CapRes(266) + JOBLIST.Recordset("ENTOLH") + mL_CapRes(267), vbYesNo)
        End If
         '  266   "Προσοχή θα σβηστεί η εργασία "
         '  267   " μαζί με τις αντίστοιχες συνταγές.Να σβησθεί;"
       If ans = vbNo Then
          Exit Sub
       End If
       
       m_entolh = JOBLIST.Recordset("ENTOLH")
       
       
       
 If "IRIS.TXT" = Dir("C:\TALOS\IRIS.TXT") Then
   Set MyWorkspace = Workspaces(0)   ' Get default Workspace.
   Set mydb = MyWorkspace.OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
   
   Set r = mydb.OpenRecordset("JOBLIST3")
   r.Index = "ENTOLH"
   r.Seek "=", m_entolh
   'If r.NoMatch Then
      ' MyWorkspace.BeginTrans
       '------------------------------------------
       On Error GoTo JOB_COMM3
       sql_del = "insert into joblist3  select [joblist].*  from [joblist] where entolh='" + m_entolh + "';"
       mydb.Execute sql_del
       sql_del = "DELETE DISTINCTROW JOBLIST.* FROM JOBLIST where entolh='" + m_entolh + "'"
       mydb.Execute sql_del
       sql_del = "insert into prospau3  select [prospau2].*  from [prospau2] where entolh='" + m_entolh + "';"
       mydb.Execute sql_del, dbFailOnError
       sql_del = "DELETE DISTINCTROW PROSPAU2.* FROM PROSPAU2 where entolh='" + m_entolh + "';"
       mydb.Execute sql_del
       mydb.Close
       On Error GoTo 0
       JOBLIST.Refresh
       Form_Paint
      ' MyWorkspace.CommitTrans
       Exit Sub
   'End If
End If
       
       
       
       
       
       
       
       
       
       
       
       
       
       
       
          JOBLIST.Recordset.DELETE
          m_entolh = IIf(IsNull(m_entolh), Space(JOBLIST.Recordset("ENTOLH").Size), m_entolh)
           Set mydb = Workspaces(0).OpenDatabase("c:\talos\recipies.mdb")
           ' Open table.
           Set MyTable = mydb.OpenRecordset("prospau2", dbOpenTable)
           
            SQL = "DELETE DISTINCTROW prospau2.* FROM prospau2 where status<>1 and ENTOLH='" + m_entolh + "';"
            mydb.Execute SQL, dbFailOnError
          
          
'          prospau2.Recordset.Index = "ENTOLH"
 '         prospau2.Recordset.Seek "=", m_entolh
  '        Do While Not prospau2.Recordset.EOF And m_entolh = prospau2.Recordset("ENTOLH")
   '           prospau2.Recordset.Delete
    '          prospau2.Recordset.MoveNext
     '     Loop
          JOBLIST.Recordset.MoveLast
       'End If
    'End If
End If

Form_Paint

Exit Sub
JOB_COMM3:

MyWorkspace.Rollback
Exit Sub


End Sub

Sub Update_Controls()
DBGrid1.Enabled = False
Text4.Enabled = False
Text10.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False

 Label2.Caption = CH_NULL(JOBLIST.Recordset("entolh"))
 DBCombo1.text = CH_NULL(JOBLIST.Recordset("kod_pel"))
 DBCombo3.text = CH_NULL(JOBLIST.Recordset("eidos_pani"))
 DBCombo2.text = CH_NULL(JOBLIST.Recordset("eidos_baf"))
 Text5.text = IIf(IsNull(JOBLIST.Recordset("sxesh_mpan")), 10, JOBLIST.Recordset("sxesh_mpan"))
 Text6.text = CH_NULL(JOBLIST.Recordset("kod_erg"))
 Text7.text = IIf(IsNull(JOBLIST.Recordset("hme_entol")), Date, JOBLIST.Recordset("hme_entol"))
 Text8.text = IIf(IsNull(JOBLIST.Recordset("hme_parad")), Date, JOBLIST.Recordset("hme_parad"))
 Text1.text = CH_NULL(JOBLIST.Recordset("parat"))
 Text9.text = CH_NULL(JOBLIST.Recordset("xhmika"))
 Text3.text = CH_NULL(JOBLIST.Recordset("apoxrvsh"))
 Text2.text = CH_NULL(JOBLIST.Recordset("baros_pani"))
End Sub

Function CH_NULL(x)
If IsNull(x) Then
   CH_NULL = Space(2)
Else
   CH_NULL = x
End If
End Function

Function Updating()
     
'ΕΛΕΓΧΩ ΑΝ ΣΥΜΠΛΗΡΩΘΗΚΑΝ ΤΑ ΑΠΑΡΑΙΤΗΤΑ ΠΕΔΙΑ
If m_edit = 2 Or m_edit = 1 Then
  DBGrid1.Enabled = False
  If Len(Trim(DBCombo1.text)) = 0 Or Len(Trim(DBCombo3.text)) = 0 Or Val(Text5.text) = 0 Then
     MsgBox mL_Res(270) '"Πρέπει να συμπληρωθούν τα πεδία Πελάτης,Πανί,βάρος πανιού και Σχέση μπάνιου"
     Updating = 0
     Exit Function
  End If
  If Not IsNumeric(Text2.text) Then
     MsgBox mL_Res(265) ' "Πρέπει να συμπληρωθεί το πεδίο βάρος πανιού."
     Updating = 0
     Exit Function
  End If
End If

If m_edit = 2 Then 'add
  
  On Error GoTo 3000
  JOBLIST.Recordset.AddNew
     Update_Fields  ' απο τα controls
  JOBLIST.Recordset.update
  JOBLIST.Recordset.MoveLast
  Rezero_controls
  m_edit = 0
  Command1.Caption = Caption_add
  DBGrid1.Enabled = True
  Text4.Enabled = True
  Text10.Enabled = True

  Option1.Enabled = True
  Option2.Enabled = True
  Option3.Enabled = True
  

ElseIf m_edit = 1 Then 'edit
     On Error GoTo 3000
     
     JOBLIST.Recordset.EDIT
     Update_Fields  ' απο τα controls
     
     JOBLIST.Recordset.update
     Rezero_controls
     m_edit = 0
     Command2.Caption = Caption_edit
     DBGrid1.Enabled = True
     Text4.Enabled = True
     Text10.Enabled = True
     Option1.Enabled = True
     Option2.Enabled = True
     Option3.Enabled = True
     
End If
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command5.Enabled = True

DBGrid1.Refresh
Updating = 1

Exit Function
3000
If Err.Number = 3186 Or Err.Number = 3260 Then
    ans = MsgBox(mLResnew(903, "The file is updated by another program.You can wait a little and after <Yes> to update the record  or <No> to cancel the update.", 1), vbYesNo)
    If ans = vbYes Then
       Resume
    Else
       On Error Resume Next
    End If
End If
Resume
End Function

Sub Rezero_controls()
 Label2.Caption = "  "
 DBCombo1.text = " "
 DBCombo3.text = " "
 DBCombo2.text = "  "
 Text5.text = "10"
 Text6.text = " "
 Text7.text = " "
 Text8.text = " "
 Text1.text = " "
 Text9.text = " "
 Text3.text = " "
 Text2.text = Default_baros_panioy

End Sub

Private Sub Command1_Click()
'Νεα εργασία

Dim m
If m_edit = 2 Then ' ΗΔΗ ΕΧΕΙ ΠΑΤΗΣΕΙ ΝΕΑ ΕΓΓΡΑΦΗ και ενημερώνει
   m = Updating
   If m = 1 Then
      Staueropoihse_Arxeio 0
   End If
   If Command5.Enabled Then Command5.SetFocus
   Exit Sub
End If

DBCombo1.SetFocus

' πρωτη φορά πατάει το κουμπί
m_edit = 2 'add
If JOBLIST.Recordset.RecordCount = 0 Then
   m = 1
Else
  On Error Resume Next
  JOBLIST.Recordset.MoveLast
  If IsNull(JOBLIST.Recordset("ENTOLH")) Then
     m = 1
   Else
      m = Val(JOBLIST.Recordset("ENTOLH")) + 1
   End If
   m = IIf(m = 0, 1, m)
   
   On Error GoTo 0
End If

Label2.Caption = Right$(str(m + 1000000), 6)

' καντα όλα enabled
Enable_Controls

Text5.text = DEFAULT_MPANIO ' 8.5
Text2.text = Default_baros_panioy
Command1.Caption = Caption_update
Command2.Enabled = False
Command3.Enabled = False
Command5.Enabled = False

Staueropoihse_Arxeio 1

End Sub

Private Sub Command2_Click()
' Διορθωση

Dim Lathos, a
Lathos = 0
On Error GoTo chech_current
   a = JOBLIST.Recordset("entolh")
On Error GoTo 0


If m_edit = 1 Then ' ΗΔΗ ΕΧΕΙ ΠΑΤΗΣΕΙ διορθωση και τώρα καταχωρεί
    m = Updating
    If m = 1 Then
       Staueropoihse_Arxeio 0
    End If
   Exit Sub
End If
   
 
   
   Command2.Caption = Caption_update
   Command1.Enabled = False
   Command3.Enabled = False
   Command5.Enabled = False
   
   m_edit = 1  'editing
   Update_Controls
    Staueropoihse_Arxeio 1
   
   Exit Sub
   
chech_current:
If Err = 3021 Then
  Lathos = 1
  Exit Sub
End If
 Resume Next
End Sub

Sub Staueropoihse_Arxeio(flag)
  
  
  If flag = 1 Then 'ακινητοποιεί όλα τα control που κινούν το αρχείο
     DBGrid1.Enabled = False
     Text4.Enabled = False
     Text10.Enabled = False
     Option1.Enabled = False
     Option2.Enabled = False
     Option3.Enabled = False
  Else
     ' επαναφέρει όλα τα control που κινούν το αρχείο
     DBGrid1.Enabled = True
     Text4.Enabled = True
     Text10.Enabled = True
     Option1.Enabled = True
     Option2.Enabled = True
     Option3.Enabled = True
  End If
End Sub



Private Sub NEW_Command3_Click()
Dim eof, aaa, ans, MyTable, mydb, SQL

If JOBLIST.Recordset.eof Or JOBLIST.Recordset.BOF Then
    ' TIPOTA
Else
 
       If IsNull(JOBLIST.Recordset("ENTOLH")) Then
           ans = MsgBox("266-267." + mL_CapRes(266) + " " + mL_CapRes(267), vbYesNo)
       Else
          ans = MsgBox("266-267." + mL_CapRes(266) + JOBLIST.Recordset("ENTOLH") + mL_CapRes(267), vbYesNo)
       End If
       
       If ans = vbNo Then
          Exit Sub
       End If
       m_entolh = JOBLIST.Recordset("ENTOLH")
          
          
          
       'krato thn thesi
       JOBLIST.Recordset.MoveNext
       If JOBLIST.Recordset.eof Then
          eof = 1
       Else
          aaa = JOBLIST.Recordset.Bookmark
       End If
          
          m_entolh = IIf(IsNull(m_entolh), Space(JOBLIST.Recordset("ENTOLH").Size), m_entolh)
           Set mydb = Workspaces(0).OpenDatabase("c:\talos\recipies.mdb")
           ' Open table.
           Set MyTable = mydb.OpenRecordset("prospau2", dbOpenTable)
           
          
            SQL = "insert into prospau3 select prospau2.* FROM prospau2 where ENTOLH='" + m_entolh + "';"
            mydb.Execute SQL, dbFailOnError
          
            SQL = "DELETE DISTINCTROW prospau2.* FROM prospau2 where ENTOLH='" + m_entolh + "';"
            mydb.Execute SQL, dbFailOnError
          
            
            SQL = "insert into joblist3 select joblist.* FROM joblist where ENTOLH='" + m_entolh + "';"
            mydb.Execute SQL, dbFailOnError
          
            SQL = "DELETE DISTINCTROW joblist.* FROM joblist where ENTOLH='" + m_entolh + "';"
            mydb.Execute SQL, dbFailOnError
          
          
          
End If

JOBLIST.Refresh
If eof = 1 Then
   JOBLIST.Recordset.MoveLast
Else
   JOBLIST.Recordset.Bookmark = aaa
End If



' Form_Paint
End Sub

Private Sub Command4_Click()
sel_menu = 0

If m_edit > 0 Then
   response = MsgBox(mL_Res(268), vbYesNo)
   '268 ("Να καταχωρηθούν οι αλλαγές; "
   If response = vbYes Then
      If Updating = 0 Then
         MsgBox mL_Res(269) ' "Οι αλλαγές δεν καταχωρήθηκαν"
         '269  "Οι αλλαγές δεν καταχωρήθηκαν"
      End If
   End If
End If
Unload Me
End Sub

Private Sub Command5_Click()
   
   On Error Resume Next
   m_entolh = JOBLIST.Recordset("ENTOLH") 'Label2.Caption  ' ΕΝΤΟΛΉ
  If Not IsNull(m_entolh) Then
      If Val(m_entolh) > 0 Then
      mJOBLIST.Hide
      FRMSYNT2.Show
    End If

 End If
 On Error GoTo 0
End Sub

Private Sub Command6_Click()
  mJOBLIST.Hide
  WAIT.Show
  
End Sub

Sub Update_Fields() ' απο τα controls
   
   JOBLIST.Recordset("ENTOLH") = Label2.Caption
   
   JOBLIST.Recordset("kod_pel") = left(CH_NULL(DBCombo1.text), JOBLIST.Recordset("kod_pel").Size)
   
   BBB = left(CH_NULL(DBCombo3.text), JOBLIST.Recordset("eidos_pani").Size)
   JOBLIST.Recordset("eidos_pani") = BBB
   JOBLIST.Recordset("eidos_baf") = left(CH_NULL(DBCombo2.text), JOBLIST.Recordset("eidos_baf").Size)
   JOBLIST.Recordset("sxesh_mpan") = Val(Text5.text)
   JOBLIST.Recordset("kod_erg") = left(CH_NULL(Text6.text), JOBLIST.Recordset("kod_erg").Size)
   JOBLIST.Recordset("hme_entol") = IIf(IsDate(Text7.text), Text7, Date)
   JOBLIST.Recordset("hme_parad") = IIf(IsDate(Text8.text), Text8, Date)
   On Error Resume Next
   JOBLIST.Recordset("parat") = left(CH_NULL(Text1.text), JOBLIST.Recordset("parat").Size)
   On Error GoTo 0
   JOBLIST.Recordset("xhmika") = Val(left(CH_NULL(Text9.text), JOBLIST.Recordset("xhmika").Size))
   JOBLIST.Recordset("apoxrvsh") = left(CH_NULL(Text3.text), JOBLIST.Recordset("apoxrvsh").Size)
   JOBLIST.Recordset("baros_pani") = Val(left(CH_NULL(Text2.text), JOBLIST.Recordset("baros_pani").Size))

   pania.Recordset.MoveFirst
   Do While Not pania.Recordset.eof
      If Trim(BBB) = Trim(pania.Recordset("pani")) Then
          found = 1
          JOBLIST.Recordset("APOR") = pania.Recordset("APOR")
          Exit Do
      End If
      pania.Recordset.MoveNext
   Loop
   If found = 0 Then
       JOBLIST.Recordset("APOR") = 0
   End If
End Sub

Private Sub Enable_Controls()
   DBCombo1.Enabled = True
   DBCombo3.Enabled = True
   DBCombo2.Enabled = True
   
   Text3.Enabled = True
   Text2.Enabled = True
   Text5.Enabled = True
   Text6.Enabled = True
   Text7.Enabled = True
   Text8.Enabled = True
   Text1.Enabled = True
   Text9.Enabled = True
   
End Sub
Sub Disable_Controls()
   DBCombo1.Enabled = False
   DBCombo3.Enabled = False
   DBCombo2.Enabled = False
   
   Text3.Enabled = False
   Text2.Enabled = False
   Text5.Enabled = False
   Text6.Enabled = False
   Text7.Enabled = False
   Text8.Enabled = False
   Text1.Enabled = False
   Text9.Enabled = False
End Sub

Private Sub DBCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
   
End Sub


Private Sub DBCombo2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub


Private Sub DBCombo3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub


Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
   MORDER = DBGrid1.Columns(ColIndex).DataField

End Sub


Private Sub Form_DblClick()
Dim a
    ' CommonDialog1.ShowColor
    ' Me.BackColor = CommonDialog1.Color
     a = save_colors(Me, Me.BackColor)
      
'
      
End Sub

Private Sub Form_Load()
proth_fora = 1
m_edit = 0
MDIForm1.Arxeia(10).Visible = False
MDIForm1.Syntages(20).Visible = False
MDIForm1.bohuhtika(30).Visible = False
MDIForm1.ejodos(40).Visible = False




'Me.BackColor =
a = find_colors(Me)


Caption_add = mL_CapRes(230) '   "Νέα Εργασία"
Command1.Caption = Caption_add

Caption_update = mL_CapRes(231) '   "Ενημέρωση"
Caption_edit = mL_CapRes(229) ' "Διόρθωση"
Command2.Caption = Caption_edit


Command3.Caption = mL_CapRes(232) '   "Διαγραφή"
Command4.Caption = mL_CapRes(233)   '"ΕΞΟΔΟΣ"
Command5.Caption = mL_CapRes(234) '  "Κατασκευή Συνταγής"
' grid
DBGrid1.Caption = mL_CapRes(236) '   "ΕΡΓΑΣΙΕΣ"
DBGrid1.Columns(0).Caption = mL_CapRes(237) '   "Eργασία"
DBGrid1.Columns(1).Caption = mL_CapRes(238) '  "Απόχρωση"


DBGrid1.Columns(3).Caption = mL_CapRes(240) '   "Πελάτης"

DBGrid1.Columns(7).Caption = mL_CapRes(241) '
DBGrid1.Columns(2).Caption = mL_CapRes(245) '
DBGrid1.Columns(8).Caption = mL_CapRes(239) '   "Βαφή"

DBGrid1.Columns(6).Caption = mL_CapRes(242) '  "Ημερ.Παράδοσης"

DBGrid1.Columns(4).Caption = mL_CapRes(243) '   parathrhseis
DBGrid1.Columns(5).Caption = mL_CapRes(244) '   eidos panioy



'// labels
Option1.Caption = mL_CapRes(248) '   "Κατά Εντολή"
Option2.Caption = mL_CapRes(249) '   "Κατά Ημερ.Παράδοσης"
Option3.Caption = mL_CapRes(250) '   "Κατά Παρατηρήσεις"


Label16.Caption = mLResnew(907, "Chemicals by hand", 0)
Label8.Caption = mL_CapRes(251) '   "Παρατηρήσεις"
Label4.Caption = mL_CapRes(252) '   "Πελάτης"
Label5.Caption = mL_CapRes(253) '   "Πανί"
Label6.Caption = mL_CapRes(254) '   "Απόχρωση"
Label7.Caption = mL_CapRes(255) '   "Διεργασία"
Label10.Caption = mL_CapRes(256) '   "Σχέση Μπάνιου"
Label11.Caption = mL_CapRes(257) '   "Αριθ.Παραγγελίας"
Label12.Caption = mL_CapRes(258) '   "Ημερ.Παραγγελίας"
Label13.Caption = mL_CapRes(259) '  "Ημερ.Παράδοσης"

Label14.Caption = mL_CapRes(263) '  "Ταξινόμηση"
Label15.Caption = mL_CapRes(246) '  "jump to task"

Label9.Caption = mL_CapRes(264) '  "baros panioy

mJOBLIST.Caption = mL_CapRes(260) '   "ΕΡΓΑΣΙΕΣ"
Label1.Caption = mL_CapRes(260) '   "ΕΡΓΑΣΙΕΣ"
Label1.FontSize = 18
Label3.Caption = mL_CapRes(261) '   "Παρατηρήσεις"





'default μπάνιο
Set mydb = Workspaces(0).OpenDatabase("c:\TALOS\RECIPIES.MDB", False, False)
Set r = mydb.OpenRecordset("parametroi")
If IsNull(r("MPANIO")) Then
  DEFAULT_MPANIO = 10
Else
  DEFAULT_MPANIO = r("MPANIO")
End If
r.Close
Text5.text = DEFAULT_MPANIO





End Sub

Private Sub Form_Paint()
  If JOBLIST.Recordset.RecordCount = 0 Then
      JOBLIST.Recordset.AddNew
         JOBLIST.Recordset("entolh") = "000001"
      JOBLIST.Recordset.update
   End If
On Error GoTo error
   'Joblist.Refresh
   If Command1.Enabled = True And Command2.Enabled = True Then
        If proth_fora = 1 Then
           JOBLIST.Recordset.MoveLast
        End If
   End If
   Screen.MousePointer = 1
proth_fora = 0
Command1.SetFocus




   Exit Sub


error:
   If Err = 3021 Then
      Command2.Enabled = False
      Command3.Enabled = False
      Command5.Enabled = False
   End If
   Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
'MDIForm1.arxeia.Visible = True
'MDIForm1.Syntages.Visible = True
'MDIForm1.Bohuhtika.Visible = True
'MDIForm1.ejodos.Visible = True

MDIForm1.Arxeia(10).Visible = True
MDIForm1.Syntages(20).Visible = True
MDIForm1.bohuhtika(30).Visible = True
MDIForm1.ejodos(40).Visible = True



End Sub

Private Sub Label8_Click()
' Label8.Caption = JobList.Recordset.Bookmarkable
End Sub

Private Sub Option1_Click()
       JOBLIST.Recordset.Index = "ENTOLH"
   
End Sub

Private Sub Option2_Click()
     JOBLIST.Recordset.Index = "HME_PARAD"
End Sub

Private Sub Option3_Click()
      JOBLIST.Recordset.Index = "PARAT"
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub


Private Sub Text10_LostFocus()
Dim m
On Error Resume Next
On Error GoTo 0
If Len(Trim(Text10.text)) > 0 Then
Else
  Exit Sub
End If

JOBLIST.Recordset.Index = "APOXRVSH"

JOBLIST.Recordset.Seek ">=", Text10.text
m = JOBLIST.Recordset.Bookmark
  JOBLIST.Recordset.Index = "ENTOLH"
JOBLIST.Recordset.Bookmark = m



'Exit Sub
'ERR:

'mydb = OpenDatabase("C:\TALOS\RECIPIES.MDB")

'mydb.Execute "CREATE INDEX APOXRVSH ON JOBLIST(APOXRVSH);"
'mydb.Close
End Sub


Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub


Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub


Private Sub Text4_LostFocus()
JOBLIST.Recordset.Index = "ENTOLH"
On Error Resume Next
JOBLIST.Recordset.Seek "=", Right("000000" + Text4.text, 6)
On Error GoTo 0
'DBGrid1.SetFocus
End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Text5_LostFocus()
   If Not IsNumeric(Text5) Then
      Text5 = 10
   End If
End Sub


Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"
End Sub


Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Text7_LostFocus()
     If Not IsDate(Text7.text) Then
           Text7.text = Date
     End If
End Sub

Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub Text8_LostFocus()
    If Not IsDate(Text8.text) Then
           Text8.text = Date
     End If
End Sub

Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub



VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form EPANEKTYP 
   BackColor       =   &H00FFFF00&
   ClientHeight    =   6375
   ClientLeft      =   510
   ClientTop       =   2160
   ClientWidth     =   9660
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6375
   ScaleWidth      =   9660
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text8 
      Height          =   372
      Left            =   5520
      TabIndex        =   26
      Text            =   "200"
      Top             =   4800
      Width           =   1452
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Disp.Recipes"
      Height          =   372
      Left            =   120
      TabIndex        =   25
      Top             =   6000
      Width           =   1092
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      Left            =   6960
      TabIndex        =   24
      Text            =   "Combo1"
      Top             =   5760
      Width           =   1932
   End
   Begin VB.CommandButton Command3 
      Caption         =   "dispen"
      Height          =   252
      Left            =   6000
      TabIndex        =   23
      Top             =   5760
      Width           =   612
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Detail Printing"
      Height          =   345
      Left            =   5400
      TabIndex        =   22
      Top             =   5295
      Width           =   1905
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Current Recipies"
      Height          =   315
      Index           =   1
      Left            =   5235
      TabIndex        =   21
      Top             =   4200
      Value           =   -1  'True
      Width           =   1905
   End
   Begin VB.OptionButton Option1 
      Caption         =   "¡Ò˜ÂﬂÔı"
      Height          =   350
      Index           =   0
      Left            =   5220
      TabIndex        =   20
      Top             =   3810
      Width           =   1905
   End
   Begin VB.CommandButton Command2 
      Caption         =   "≈–¡Õ¡÷œ—¡ ”≈ À…”‘¡ ¡Õ¡ÃœÕ«”"
      Height          =   405
      Left            =   7380
      TabIndex        =   19
      Top             =   4800
      Width           =   2145
   End
   Begin VB.TextBox Text7 
      Height          =   315
      Left            =   3600
      TabIndex        =   17
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   300
      Left            =   3630
      TabIndex        =   14
      Top             =   5055
      Width           =   1320
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   3615
      TabIndex        =   13
      Top             =   4485
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "≈Œœƒœ”"
      Height          =   420
      Left            =   7365
      TabIndex        =   12
      Top             =   5250
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "¡Õ¡∆«‘«”«"
      Height          =   405
      Left            =   285
      TabIndex        =   9
      Top             =   5430
      Width           =   4710
   End
   Begin VB.Data PELATES 
      Caption         =   "–≈À¡‘≈”"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\recipies.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4020
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT *FROM pelates ORDER BY ONOMA"
      Top             =   6780
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      DataField       =   "APOXRVSH"
      DataSource      =   "mem2"
      Height          =   285
      Left            =   270
      TabIndex        =   6
      Top             =   4980
      Width           =   2160
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      DataField       =   "AA_PROSPAU"
      DataSource      =   "mem2"
      Height          =   285
      Left            =   8565
      TabIndex        =   3
      Top             =   4005
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      DataField       =   "ENTOLH"
      DataSource      =   "mem2"
      Height          =   285
      Left            =   7365
      TabIndex        =   2
      Top             =   4005
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "≈ ‘’–Ÿ”«"
      Height          =   405
      Left            =   7395
      TabIndex        =   1
      Top             =   4365
      Width           =   2145
   End
   Begin VB.Data SQLPROSPAU 
      Caption         =   "SQLPROSPAU"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\recipies.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   930
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT prospau2.* FROM PROSPAU2   where prospau2.SEIRA=1 AND prospau2.STATUS<2  ORDER BY prospau2.entolh;"
      Top             =   6690
      Visible         =   0   'False
      Width           =   2655
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   6165
      Top             =   6945
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "c:\talos\vis4-99.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "_EPANEKT.frx":0000
      Height          =   3630
      Left            =   105
      OleObjectBlob   =   "_EPANEKT.frx":0019
      TabIndex        =   11
      Top             =   120
      Width           =   9465
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "_EPANEKT.frx":1758
      DataField       =   "PELATHS"
      DataSource      =   "mem2"
      Height          =   315
      Left            =   270
      TabIndex        =   10
      Top             =   4230
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   0
      ListField       =   "ONOMA"
      BoundColumn     =   "ONOMA"
      Text            =   "DBCombo1"
   End
   Begin VB.Label apotelesma 
      Height          =   252
      Left            =   5280
      TabIndex        =   27
      Top             =   4560
      Width           =   1332
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "◊ÂÈÒÈÛÙﬁÚ"
      Height          =   285
      Left            =   2700
      TabIndex        =   18
      Top             =   3975
      Width           =   840
   End
   Begin VB.Shape Shape1 
      Height          =   2145
      Left            =   105
      Top             =   3810
      Width           =   5040
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "≈˘Ú"
      Height          =   285
      Left            =   2715
      TabIndex        =   16
      Top             =   5055
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      Height          =   285
      Left            =   2715
      TabIndex        =   15
      Top             =   4500
      Width           =   840
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "¡¸˜Ò˘ÛÁ"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   285
      TabIndex        =   8
      Top             =   4740
      Width           =   2145
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "–ÂÎ‹ÙÁÚ"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   270
      TabIndex        =   7
      Top             =   3975
      Width           =   1965
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "–ÒÔÛ‹ËÂÈ·"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8565
      TabIndex        =   5
      Top             =   3780
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "≈Ò„·Ûﬂ·"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7365
      TabIndex        =   4
      Top             =   3780
      Width           =   1095
   End
   Begin VB.Label text1 
      Height          =   465
      Left            =   60
      TabIndex        =   0
      Top             =   6015
      Width           =   8985
   End
End
Attribute VB_Name = "EPANEKTYP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ayjon
Dim j_katal  'Í·Ù‹ÎÁÓÁ  "2"=ÒÔÛ·Ë2   " "=prospau3
Dim p_katal  'Í·Ù‹ÎÁÓÁ  "2"=ÒÔÛ·Ë2   "3"=prospau3

Sub print99_rep()

On Error Resume Next


Set mydb = Workspaces(0).OpenDatabase("c:\TALOS\recipies.mdb", False, False)


'MyDB.Execute "DROP TABLE dokpros2;"
'MyDB.Execute "SELECT prospau" + p_katal + ".*,XIMITECH.PERIGRAFH INTO DOKPROS2 from prospau" + p_katal + " INNER JOIN XIMITECH ON prospau" + p_katal + ".KOD=XIMITECH.KOD where entolh='" + Right$("000000" + Trim(Text2.Text), 6) + "' and aa_prospau='" + Right$("00" + Trim(Text3.Text), 2) + "';"
'MyDB.Execute "DROP TABLE dokprosp;"
'MyDB.Execute "select dokpros2.*,joblist.* into dokprosp from dokpros2 inner join joblist on dokpros2.entolh=joblist.entolh;"
'MyDB.Close


 ' Kill "C:\TALOS\dokprosp.dbf"


'If prospau2.Recordset.EOF Then prospau2.Recordset.MoveLast

'Set MyDB = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)

mydb.Execute "drop table DOKPROS2;"


If Option1(0).value = True Then ' istoriko
   mydb.Execute "SELECT prospau3.*  INTO DOKPROS2  from prospau3  where entolh='" + Right$("000000" + Trim(Text2.text), 6) + "' and aa_prospau='" + Right$("00" + Trim(Text3.text), 2) + "';"
Else
   mydb.Execute "SELECT prospau2.*  INTO DOKPROS2  from prospau2  where entolh='" + Right$("000000" + Trim(Text2.text), 6) + "' and aa_prospau='" + Right$("00" + Trim(Text3.text), 2) + "';"
End If
mydb.Execute "DROP TABLE dokprosp;"

mydb.Execute "select dokpros2.*  into dokprosp from dokpros2  ;"
 
mydb.Close
 
 Report1.Action = 1



End Sub

Private Sub Command1_Click()
sel_menu = 0
 Unload EPANEKTYP
' mJOBLIST.Show
End Sub

Private Sub Command2_Click()

Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
SQL = "UPDATE prospau" + p_katal + " SET STATUS=1  where entolh='" + Right$("000000" + Trim(Text2.text), 6) + "' and aa_prospau='" + Right$("00" + Trim(Text3.text), 2) + "';"
mydb.Execute SQL
 
 
 
 
End Sub



Private Sub Command3_Click()

Set db = OpenDatabase("c:\talos\recipies.mdb")
Set r = db.OpenRecordset("ximitech")
'r.FindFirst ("addr_prot=91")
r.MoveFirst
Do While Not r.eof
    If r("addr_prot") = 91 Then
        
        Exit Do
    End If
    r.MoveNext
Loop
       
    
'





If r.NoMatch Then
   MsgBox "no find"
   Exit Sub
End If




For k = 1 To 5

    

                Me.Caption = k

                Syn_DosXorPin.asked_q = Val(Text8.text) '  ZHTOYMENH –œ”œTHTA ”E MGRS
                Syn_DosXorPin.Level = r("ypol_prot")     '√—¡Ã¡—…¡ Ã–œ’ ¡À…œ’
                Syn_DosXorPin.Valve_On = valve_1_on  ', valve_2_on) ' ¡Õœ…√Ã¡ ¬¡À¬…ƒ¡”
                Syn_DosXorPin.Valve_Off = valve_1_off ', valve_2_off)   '  À≈…”…Ãœ ¬¡À¬…ƒ¡”
                Syn_DosXorPin.Caption = "91: kokkino " 'ÂÒÈ„Ò·ˆÁ ˜Ò˛Ï·ÙÔÚ Ôı ‰ÔÛÔÏÂÙÒÂﬂÙ·È
                Syn_DosXorPin.Show 1
                
                r.EDIT
                r("ypol_prot") = r("ypol_prot") - Val(apotelesma.Caption) / 1000
                r.update
                
                
                
Next
k = 1
           'End If
  r.Close
            ' GoTo 111
End Sub

Private Sub Command4_Click()
   Dim m
   On Error Resume Next
   
   print99_rep
If Check1 Then  'megalo xarti
   Report1.ReportFileName = "c:\talos\vis4-99b.rpt"
Else
   Report1.ReportFileName = "c:\talos\vis4-99.rpt"
End If
   Report1.Action = 1
   Exit Sub
   
End Sub


Private Sub Command5_Click()
Dim u(1 To 10), D
On Error Resume Next
           If IsDate(Text5) And IsDate(Text6) Then
              Where = " seira=1 and  hme>=datevalue('" + Text5 + "') and hme<=datevalue('" + Text6 + "')"
              SQL = "SELECT prospau" + p_katal + ".* from prospau" + p_katal + "  where " + Where + "ORDER BY prospau" + p_katal + ".entolh;"
              SQLPROSPAU.RecordSource = SQL
              SQLPROSPAU.Refresh
                             
              Set D = SQLPROSPAU.Recordset
              D.MoveFirst
              Do While Not D.eof
                If D("usere") > 0 And D("usere") < 10 Then
                   u(D("usere")) = u(D("usere")) + 1
                End If
                D.MoveNext
              Loop
              msg = ""
              For k = 1 To 10
                 If u(k) > 0 Then
                    msg = msg + "User " + LTrim(str(k)) + " :" + LTrim(str(u(k))) + Chr(13)
                 End If
              Next
              If Len(msg) > 1 Then
                 MsgBox msg
              End If
              
           End If
    
End Sub

Private Sub Command6_Click()
'--------------------- ·Ì·ÊﬁÙÁÛÁ -------------------
  Dim SQL, user
'       mem2.Recordset.Edit
 '      mem2.Recordset.update
  
  
    'DBGrid1.Columns(0).DataField = "ENTOLH"
    'DBGrid1.Columns(4).DataField = "APOXRVSH"
    'DBGrid1.Columns(5).DataField = "kod_pel"
    'DBGrid1.Columns(6).DataField = "BAROS_PANI"
    'DBGrid1.Columns(4).DataField = "HME"
  ' If Len(Text4.text) = 0 And DBCombo1 = "" Then     '¡–œ◊—  color AND CUSTOMER = NOTHING
  '
  '      Where = " seira=1 and  hme>=datevalue('" + Text5 + "') and hme<=datevalue('" + Text6 + "')"
  '       sql = "SELECT prospau" + p_katal + ".* from prospau" + p_katal + "  where " + Where + "ORDER BY prospau" + p_katal + ".entolh;"
  '       SQLPROSPAU.RecordSource = sql
  '
  ' End If
  If Len(Text4.text) = 0 Then  '¡–œ◊—  color
      If DBCombo1 = "" Then   'no customer
           If IsDate(Text5) And IsDate(Text6) Then
              Where = " seira=1 and  hme>=datevalue('" + Text5 + "') and hme<=datevalue('" + Text6 + "')"
              SQL = "SELECT prospau" + p_katal + ".* from prospau" + p_katal + "  where " + Where + "ORDER BY prospau" + p_katal + ".entolh;"
              SQLPROSPAU.RecordSource = SQL
           Else
               'LAGAKIS
               SQL = "SELECT prospau" + p_katal + ".* FROM PROSPAU" + p_katal + "  where SEIRA=1  ORDER BY entolh;"
               SQLPROSPAU.RecordSource = SQL
           End If
      Else
     If IsDate(Text5) And IsDate(Text6) Then
         Where = " seira=1 and  kod_pel='" + DBCombo1 + "' and hme>=datevalue('" + Text5 + "') and hme<=datevalue('" + Text6 + "')"
         SQL = "SELECT prospau" + p_katal + ".* from prospau" + p_katal + "  where " + Where + "ORDER BY prospau" + p_katal + ".entolh;"
         SQLPROSPAU.RecordSource = SQL
     Else
         Where = " seira=1 and kod_pel='" + DBCombo1 + "'"
         SQL = "SELECT prospau" + p_katal + ".* FROM PROSPAU" + p_katal + " where " + Where + "ORDER BY prospau" + p_katal + ".entolh;"
         SQLPROSPAU.RecordSource = SQL
     End If
    End If
  Else  ' ÏÂ ·Ô˜Ò˘ÛÁ
  
       If DBCombo1 = "" Then   'no customer
          If IsDate(Text5) And IsDate(Text6) Then
             Where = " seira=1 and  APOXRVSH='" + Text4.text + "'  and hme>=datevalue('" + Text5 + "') and hme<=datevalue('" + Text6 + "')"
             SQL = "SELECT prospau" + p_katal + ".* FROM PROSPAU" + p_katal + " where " + Where + "ORDER BY prospau" + p_katal + ".entolh;"
             SQLPROSPAU.RecordSource = SQL
           Else
             Where = " seira=1   and APOXRVSH='" + Text4.text + "' "
             SQL = "SELECT prospau" + p_katal + ".* FROM PROSPAU" + p_katal + "  where " + Where + "ORDER BY prospau" + p_katal + ".entolh;"
             SQLPROSPAU.RecordSource = SQL
           End If
       Else
           If IsDate(Text5) And IsDate(Text6) Then
             Where = " seira=1 and kod_pel='" + DBCombo1 + "' and APOXRVSH='" + Text4.text + "'  and hme>=datevalue('" + Text5 + "') and hme<=datevalue('" + Text6 + "')"
             SQL = "SELECT prospau" + p_katal + ".* FROM PROSPAU" + p_katal + " where " + Where + "ORDER BY prospau" + p_katal + ".entolh;"
             SQLPROSPAU.RecordSource = SQL
           Else
             Where = " seira=1 and  kod_pel='" + DBCombo1 + "'  and APOXRVSH='" + Text4.text + "' "
             SQL = "SELECT prospau" + p_katal + ".* FROM PROSPAU" + p_katal + "  where " + Where + "ORDER BY prospau" + p_katal + ".entolh;"
             SQLPROSPAU.RecordSource = SQL
             '  SQLPROSPAU.RecordSource = "SELECT  prospau" + p_katal + ".*   FROM prospau" + p_katal + ",MEM2   WHERE prospau" + p_katal + ".kod_pel=MEM2.PELATHS and prospau" + p_katal + ".APOXR=MEM2.APOXRVSH;"
           End If
        End If
  End If
   DBGrid1.Visible = False
    DBGrid1.Columns(0).DataField = "ENTOLH"
   ' DBGrid1.Columns(4).DataField = "apoxrvsh"
   ' DBGrid1.Columns(5).DataField = "kod_pel"
    DBGrid1.Columns(6).DataField = "BAROS_PANI"
    
    DBGrid1.Columns(7).DataField = "HME"
    
    DBGrid1.Columns(8).DataField = "USERW"
    DBGrid1.Columns(9).DataField = "USERE"
'SQLPROSPAU.RecordSource = "SELECT prospau" + p_katal + ".*,joblist" + j_katal + ".kod_pel,joblist" + j_katal + ".apoxrvsh FROM PROSPAU" + p_katal + " inner join joblist" + j_katal + " on joblist" + j_katal + ".entolh=prospau" + p_katal + ".entolh  where prospau" + p_katal + ".SEIRA=1 ORDER BY prospau" + p_katal + ".entolh;"
'SQLPROSPAU.Refresh
    




       SQLPROSPAU.Refresh
     '  DBGrid1.Visible = False
       DBGrid1.Visible = True
       
       
End Sub



Private Sub DBGrid1_DblClick()
Dim e, p, a
' Set p = SQLPROSPAU.Recordset
'  If Not IsNull(SQLPROSPAU.Recordset("prospau" + p_katal + ".entolh")) Then
'   Text2.text = SQLPROSPAU.Recordset("prospau" + p_katal + ".entolh")
'  End If
  
'  If Not IsNull(SQLPROSPAU.Recordset("aa_prospau")) Then
 '  Text3.text = SQLPROSPAU.Recordset("aa_prospau")
 ' End If
 'DBGrid1.Refresh
 On Error GoTo 0
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
On Error Resume Next
    Text2.text = SQLPROSPAU.Recordset("entolh")
    Text3.text = SQLPROSPAU.Recordset("aa_prospau")
    
End Sub

Private Sub Form_Load()
Option1_Click (1)

'Me.BackColor = wait_c
MDIForm1.Arxeia(10).Visible = False
MDIForm1.Syntages(20).Visible = False
MDIForm1.bohuhtika(30).Visible = False
MDIForm1.ejodos(40).Visible = False
Option1(0).Caption = mL_CapRes(851)
Option1(1).Caption = mL_CapRes(852)


 ' DBGrid1.caption = ml_CapRes(236) '   "≈—√¡”…≈”"
DBGrid1.Columns(0).Caption = mL_CapRes(430) '   "ar.EÒ„·Ûﬂ·"
DBGrid1.Columns(1).Caption = mL_CapRes(431) '  prosp
DBGrid1.Columns(2).Caption = mL_CapRes(432) '   "epilogi
DBGrid1.Columns(3).Caption = mL_CapRes(433) '   "”˜›ÛÁ Ï‹ÌÈÔı" '
DBGrid1.Columns(4).Caption = mL_CapRes(434)  '  ·¸˜Ò˘ÛÁ
DBGrid1.Columns(5).Caption = mL_CapRes(435) '  –ÂÎ‹ÙÁÚ
DBGrid1.Columns(6).Caption = mL_CapRes(436) '  ¬‹ÒÔÚ ·ÌÈÔı
DBGrid1.Columns(7).Caption = mL_CapRes(437) ' «ÏÂÒ.–·Ò‹‰ÔÛÁÚ
      

Label4.Caption = mL_CapRes(840) ' "–ÂÎ‹ÙÁÚ"  //l4
Label8.Caption = mL_CapRes(841) ' "◊ÂÈÒÈÛÙﬁÚ"  //Ù7
Label6.Caption = mL_CapRes(842) '842 "¡¸˜Ò˘ÛÁ" // l6
Label1.Caption = mL_CapRes(843) '843 "¡¸" //l1
Label7.Caption = mL_CapRes(844) '844 "≈˘Ú" //Î7
Label2.Caption = mL_CapRes(845) '845 "≈Ò„·Ûﬂ·"  //l2
Label3.Caption = mL_CapRes(846) '846 "–ÒÔÛ‹ËÂÈ·"  //Î3
Command6.Caption = mL_CapRes(847) ' "¡Õ¡∆«‘«”«"  //c6
Command4.Caption = mL_CapRes(848) ' "≈ ‘’–Ÿ”«"  //c4
Command2.Caption = mL_CapRes(849) ' "≈–¡Õ¡÷œ—¡ ”≈ À…”‘¡ ¡Õ¡ÃœÕ«”"  //c2
Command1.Caption = mL_CapRes(850) ' "≈Œœƒœ”"  //c1




'End

      
       DBEngine.IniPath = "C:\talos\talos.INI"
End Sub

Private Sub Form_Paint()
          Screen.MousePointer = 1
          ayjon = 0
      

End Sub

Private Sub Form_Unload(Cancel As Integer)



MDIForm1.Arxeia(10).Visible = True
MDIForm1.Syntages(20).Visible = True
MDIForm1.bohuhtika(30).Visible = True
MDIForm1.ejodos(40).Visible = True
End Sub


Private Sub Option1_Click(Index As Integer)

If Option1(1) Then
     Command2.Enabled = True
    j_katal = "" 'Í·Ù‹ÎÁÓÁ  "2"=ÒÔÛ·Ë2   " "=prospau3
    p_katal = "2" 'Í·Ù‹ÎÁÓÁ  "2"=ÒÔÛ·Ë2   "3"=prospau3
   ' DBGrid1.Columns(4).DataField = "prospau2.apoxrvsh"
   ' DBGrid1.Columns(5).DataField = "prospau2.kod_pel"
    DBGrid1.Columns(0).DataField = "ENTOLH"
    DBGrid1.Columns(7).DataField = "HME"
Else  'palies
     Command2.Enabled = False
    j_katal = "3" 'Í·Ù‹ÎÁÓÁ  "2"=ÒÔÛ·Ë2   " "=prospau3
    p_katal = "3" 'Í·Ù‹ÎÁÓÁ  "2"=ÒÔÛ·Ë2   "3"=prospau3
   ' DBGrid1.Columns(4).DataField = "prospau3.apoxrvsh"
   ' DBGrid1.Columns(5).DataField = "prospau3.kod_pel"
    DBGrid1.Columns(0).DataField = "ENTOLH"
    DBGrid1.Columns(7).DataField = "HME"
End If

SQLPROSPAU.RecordSource = "SELECT prospau" + p_katal + ".* FROM PROSPAU" + p_katal + " where prospau" + p_katal + ".SEIRA=1 ORDER BY prospau" + p_katal + ".entolh;"
SQLPROSPAU.Refresh

End Sub

Private Sub Text5_LostFocus()
If Not IsDate(Text5) Then
       Text5 = " "  'date
   End If
End Sub


Private Sub Text6_LostFocus()
   If Not IsDate(Text6) Then
       Text6 = " "  ' Date
   End If
End Sub



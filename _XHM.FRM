VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmXHM 
   BackColor       =   &H00FF0000&
   ClientHeight    =   6696
   ClientLeft      =   228
   ClientTop       =   1200
   ClientWidth     =   9660
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6696
   ScaleWidth      =   9660
   WindowState     =   2  'Maximized
   Begin VB.TextBox MaskedBox3 
      Height          =   345
      Left            =   4230
      TabIndex        =   9
      Top             =   465
      Width           =   570
   End
   Begin VB.TextBox MaskedBox2 
      Height          =   345
      Left            =   1215
      TabIndex        =   18
      Top             =   465
      Width           =   2970
   End
   Begin VB.TextBox Maskedbox1 
      Height          =   345
      Left            =   315
      TabIndex        =   6
      Top             =   465
      Width           =   870
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Κατά Ταυτότητα"
      Height          =   330
      Left            =   7650
      TabIndex        =   16
      Top             =   2130
      Width           =   1800
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Κατά Διεύθυνση"
      Height          =   330
      Left            =   7650
      TabIndex        =   15
      Top             =   1695
      Width           =   1800
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Κατά Περιγραφή"
      Height          =   330
      Left            =   7650
      TabIndex        =   14
      Top             =   1290
      Value           =   -1  'True
      Width           =   1800
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\recipies.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3675
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "XIMITECH"
      Top             =   7140
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data3 
      Caption         =   "data3"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\recipies.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2010
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "XHMEDIT"
      Top             =   6840
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   7545
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   4605
      Width           =   2000
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "_XHM.frx":0000
      Left            =   7545
      List            =   "_XHM.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3255
      Width           =   2000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Εκτύπωση αρχείου"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   4950
      TabIndex        =   3
      Top             =   6030
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Αναζήτηση"
      Height          =   300
      Left            =   330
      TabIndex        =   7
      Top             =   75
      Width           =   4485
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Εξοδος"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   6225
      TabIndex        =   0
      Top             =   6030
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Διόρθωση"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   3675
      TabIndex        =   10
      Top             =   6030
      Width           =   1230
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Διαγραφή"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   2400
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6030
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Νέα Εγγραφή"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   1125
      TabIndex        =   1
      Top             =   6030
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\recipies.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   -150
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT *FROM XIMITECH"
      Top             =   6855
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5220
      Top             =   135
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   7545
      Top             =   5070
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   262150
      ReportFileName  =   "c:\talos\vis5.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "_XHM.frx":0039
      Height          =   5040
      Left            =   0
      OleObjectBlob   =   "_XHM.frx":004D
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   900
      Width           =   7464
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ταξινόμηση"
      Height          =   240
      Left            =   7575
      TabIndex        =   17
      Top             =   900
      Width           =   1995
   End
   Begin VB.Shape Shape1 
      Height          =   1350
      Left            =   7545
      Top             =   1200
      Width           =   2000
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Κατάσταση  Xημ/κού"
      Height          =   255
      Left            =   7545
      TabIndex        =   13
      Top             =   4320
      Width           =   1995
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Θέση Χημικοτεχνικού"
      Height          =   225
      Left            =   7575
      TabIndex        =   12
      Top             =   3000
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   45
      TabIndex        =   11
      Top             =   6645
      Width           =   9525
   End
End
Attribute VB_Name = "frmXHM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim real_new
Dim Oldcolor_b
Dim Oldcolor_f
Option Compare Binary
Sub LoadFields()
'On Error Resume Next
If Data3.Recordset.RecordCount = 0 Then
   Data3.Recordset.AddNew
   Data3.Recordset.update
End If
Data3.Recordset.MoveFirst
  Data3.Recordset.EDIT
  Data3.Recordset("kod") = Data1.Recordset("kod")
  Data3.Recordset("perigrafh") = Data1.Recordset("perigrafh")
  Data3.Recordset("addr_prot") = Data1.Recordset("addr_prot")
  Data3.Recordset("morfh") = Data1.Recordset("morfh")
  Data3.Recordset("kataskeyas") = Data1.Recordset("kataskeyas")
  
 Data3.Recordset("max_sygk") = Data1.Recordset("max_sygk")
  
  Data3.Recordset("zvh_prot") = Data1.Recordset("zvh_prot")
    Data3.Recordset("lhj_prot") = Data1.Recordset("lhj_prot")
  Data3.Recordset("ypol_prot") = Data1.Recordset("ypol_prot")
 Data3.Recordset("morfh") = Data1.Recordset("morfh")
  Data3.Recordset("real_cons") = Data1.Recordset("real_cons")
  Data3.Recordset("nero_dial") = Data1.Recordset("nero_dial")
  Data3.Recordset("mixtim") = Data1.Recordset("mixtim")
  Data3.Recordset("mixtim_fin") = Data1.Recordset("mixtim_fin")
  Data3.Recordset("thermo1") = Data1.Recordset("thermo1")
  Data3.Recordset("thermo2") = Data1.Recordset("thermo2")
  Data3.Recordset("nerostox") = Data1.Recordset("nerostox")
  Data3.Recordset("skonh") = Data1.Recordset("skonh")
  Data3.Recordset("timh") = Data1.Recordset("timh")
  Data3.Recordset("BOHU_XHM") = Data1.Recordset("BOHU_XHM")
Data3.Recordset.update
On Error GoTo 0


End Sub
Private Sub cmdAdd_Click()
Dim mS
mS = Data1.RecordSource
   
  DBGrid1.Enabled = False
 
  Data1.Refresh
  Data1.RecordSource = "SELECT *FROM XIMITECH where [kod]<>null order by [kod];"
  
  Data1.Refresh

If Data1.Recordset.RecordCount = 0 Then
    x = 1
Else
  Data1.Recordset.MoveLast
  x = Val(Data1.Recordset("kod")) + 1
End If

Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
Set MySet = mydb.OpenRecordset("xhmedit")

  
sql_del = "DELETE DISTINCTROW xhmedit.* FROM xhmedit"
mydb.Execute sql_del
MySet.Close
  
  
  Data3.Recordset.AddNew
  
  Data3.Recordset("kod") = Right$("000000" + LTrim(str(x)), 6)
  Data3.Recordset("mixtim") = 180
  Data3.Recordset("mixtim_fin") = 20
  Data3.Recordset("thermo1") = 3
  Data3.Recordset("thermo2") = 1
  Data3.Recordset("nero_dial") = 40
  Data3.Recordset("nerostox") = 550
  Data3.Recordset("nero_dial") = 40
  Data3.Recordset("Zvh_Prot") = 3
  Data3.Recordset.update
  
   
   
   cmdAdd.Enabled = False
   DBGrid1.Enabled = True
    
    frmXHMedit.Show
    
    
      Data1.RecordSource = mS

  
  Data1.Refresh
    
    
    
End Sub

Private Sub cmdClose_Click()
   frmXHM.cmdAdd.Enabled = True
   frmXHM.cmdUpdate.Enabled = True
   Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim problem As Integer
problem = 0

If Data1.Recordset.RecordCount = 0 Then
   MsgBox mL_Res(164) ' "Δεν υπάρχει εγγραφή για διαγραφή"
   Exit Sub
End If


 If Not IsNull(Data1.Recordset("addr_prot")) Then
     If Data1.Recordset("addr_prot") > 0 And IsNull(Data1.Recordset("timh")) Then
           MsgBox mL_Res(162), , "Talos"
           '"Δεν μπορεί να διαγραφεί επειδή το αντίστοιχο δοχείο καταλαμβάνει μία διεύθυνση στο καρτεσιανό τραπέζι.Να γίνει πρώτα απομάκρυνση με την χρήση του Robot ", , "Talos"
           Exit Sub
     End If
End If
If IsNull(Data1.Recordset("kod")) Then
   m_mhn = "  "
Else
   m_mhn = Data1.Recordset("kod")
End If

ans = MsgBox("163." + m_mhn + " " + mL_CapRes(163), 20, "Talos")
 '"Θέλετε να διαγράψετε αυτή την εγγραφή ;", 20, "Talos")
 If ans = 7 Then
     Exit Sub
 End If
 
 Data1.Recordset.EDIT
 Data1.Recordset.DELETE

  On Error Resume Next
  Data1.Recordset.MoveNext
  On Error GoTo 0

End Sub




Private Sub cmdUpdate_Click()
   Dim mS

If Data1.Recordset.RecordCount = 0 Then
   MsgBox mL_Res(165) ' "Δεν υπάρχει εγγραφή για διόρθωση"
   Exit Sub
End If

mS = Data1.RecordSource
   
   
   cmdUpdate.Enabled = False
   LoadFields
   frmXHMedit.Show

 Data1.RecordSource = mS
 Data1.Refresh

End Sub

Private Sub Combo1_Click()
Dim m
m = Data1.RecordSource
Data1.Refresh
  MORDER = "[perigrafh]"

If Option1 Then
   MORDER = "[perigrafh]"
End If
If Option2 Then
   MORDER = "[addr_prot]"
End If
If Option3 Then
   MORDER = "[kataskeyas]"
End If
   
        
'Data1.RecordSource = "SELECT *FROM XIMITECH"

'Data1.Refresh
If Combo1 = Combo1.List(0) Then  'OLA
      If Combo2 = Combo2.List(0) Then ' ola
          Data1.RecordSource = "SELECT *FROM XIMITECH where kod like '*' order by " + MORDER
      ElseIf Combo2 = Combo2.List(1) Then ' μονο σκόνες
          Data1.RecordSource = "SELECT *FROM XIMITECH where [skonh]=1 ORDER BY " + MORDER
      Else  ' dialymata
          Data1.RecordSource = "SELECT *FROM XIMITECH where [skonh]<>1 ORDER BY " + MORDER
      End If
ElseIf Combo1 = Combo1.List(1) Then     ' STO TRAPEZI
       
      If Combo2 = Combo2.List(0) Then ' ola
          Data1.RecordSource = "SELECT *FROM XIMITECH WHERE [ADDR_PROT]>0 ORDER BY " + MORDER
      ElseIf Combo2 = Combo2.List(1) Then ' SKONES
          Data1.RecordSource = "SELECT *FROM XIMITECH WHERE [SKONH]=1 AND [ADDR_PROT]>0 ORDER BY " + MORDER
      Else ' DIALYMATA
          Data1.RecordSource = "SELECT *FROM XIMITECH WHERE [SKONH]<>1 AND [ADDR_PROT]>0 ORDER BY " + MORDER
      End If
ElseIf Combo1 = Combo1.List(2) Then          ' EKTOS   TRAPEZI
      If Combo2 = Combo2.List(0) Then ' ola
          Data1.RecordSource = "SELECT *FROM XIMITECH WHERE [addr_prot]=0 or [addr_prot]=null  ORDER BY " + MORDER
      ElseIf Combo2 = Combo2.List(1) Then ' SKONES
          Data1.RecordSource = "SELECT *FROM XIMITECH WHERE [SKONH]=1 AND ([addr_prot]=0 or [addr_prot]=null)  ORDER BY " + MORDER
      Else ' DIALYMATA
          Data1.RecordSource = "SELECT *FROM XIMITECH WHERE [SKONH]=0 AND ( [addr_prot]=0 or [addr_prot]=null) ORDER BY " + MORDER
      End If
End If


'If m <> Data1.RecordSource Then
       Data1.Refresh
       
       
       
       DBGrid1.Refresh
'End If


' MsgBox "click"
End Sub




Private Sub Combo2_Click()
   Combo1_Click
End Sub

Private Sub Command1_Click()
'"     Data1.Recordset.Index = "PER"
 '    Data1.Recordset.Seek "<=", Text2.Text
End Sub

Private Sub Command3_Click()
   Dim m
   m = Data1.RecordSource
   m = "INSERT INTO DOKX " + m
   
Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
'Set MySet = MyDB.OpenRecordset("XIMITECH")
SQL = "DELETE DISTINCTROW DOKX.* FROM DOKX"
mydb.Execute SQL
'"DELETE DISTINCTROW JOBLIST.* FROM JOBLIST"
mydb.Execute m
mydb.Close
     
CrystalReport1.Action = 1
End Sub





Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
Dim m
m = Data1.RecordSource
Data1.Refresh
MORDER = "[perigrafh]"


MORDER = DBGrid1.Columns(ColIndex).DataField
'If Option1 Then
'   MORDER = "[perigrafh]"
'End If
'If Option2 Then
'   MORDER = "[addr_prot]"
'End If
'If Option3 Then
'   MORDER = "[kataskeyas]"
'End If
   
        
'Data1.RecordSource = "SELECT *FROM XIMITECH"

'Data1.Refresh


If Combo1 = Combo1.List(0) Then  'OLA
      If Combo2 = Combo2.List(0) Then ' ola
          Data1.RecordSource = "SELECT *FROM XIMITECH where kod like '*' order by " + MORDER
      ElseIf Combo2 = Combo2.List(1) Then ' μονο σκόνες
          Data1.RecordSource = "SELECT *FROM XIMITECH where [skonh]=1 ORDER BY " + MORDER
      Else  ' dialymata
          Data1.RecordSource = "SELECT *FROM XIMITECH where [skonh]<>1 ORDER BY " + MORDER
      End If
ElseIf Combo1 = Combo1.List(1) Then     ' STO TRAPEZI
       
      If Combo2 = Combo2.List(0) Then ' ola
          Data1.RecordSource = "SELECT *FROM XIMITECH WHERE [ADDR_PROT]>0 ORDER BY " + MORDER
      ElseIf Combo2 = Combo2.List(1) Then ' SKONES
          Data1.RecordSource = "SELECT *FROM XIMITECH WHERE [SKONH]=1 AND [ADDR_PROT]>0 ORDER BY " + MORDER
      Else ' DIALYMATA
          Data1.RecordSource = "SELECT *FROM XIMITECH WHERE [SKONH]<>1 AND [ADDR_PROT]>0 ORDER BY " + MORDER
      End If
ElseIf Combo1 = Combo1.List(2) Then          ' EKTOS   TRAPEZI
      If Combo2 = Combo2.List(0) Then ' ola
          Data1.RecordSource = "SELECT *FROM XIMITECH WHERE [addr_prot]=0 or [addr_prot]=null  ORDER BY " + MORDER
      ElseIf Combo2 = Combo2.List(1) Then ' SKONES
          Data1.RecordSource = "SELECT *FROM XIMITECH WHERE [SKONH]=1 AND ([addr_prot]=0 or [addr_prot]=null)  ORDER BY " + MORDER
      Else ' DIALYMATA
          Data1.RecordSource = "SELECT *FROM XIMITECH WHERE [SKONH]=0 AND ( [addr_prot]=0 or [addr_prot]=null) ORDER BY " + MORDER
      End If
End If


'If m <> Data1.RecordSource Then
       Data1.Refresh
       
       
       
       DBGrid1.Refresh
'End If


' MsgBox "click"

End Sub


Private Sub Form_DblClick()
  Dim a
      'CommonDialog1.ShowColor
      'Me.BackColor = CommonDialog1.Color
      a = save_colors(Me, Me.BackColor)
End Sub

Private Sub Form_Load()

MDIForm1.Arxeia(10).Visible = False
MDIForm1.Syntages(20).Visible = False
MDIForm1.bohuhtika(30).Visible = False
MDIForm1.ejodos(40).Visible = False




a = find_colors(Me)


Label1.BackColor = Me.BackColor
Oldcolor_b = Maskedbox1.BackColor
Oldcolor_f = Maskedbox1.ForeColor
 
On Error Resume Next
 
 Command2.Caption = mL_CapRes(101) '  "Αναζήτηση"
 cmdClose.Caption = mL_CapRes(105)  ' "Εξοδος"
 cmdUpdate.Caption = mL_CapRes(106) '   "Διόρθωση"  αν το πατήσω μία φορά τότε παίρνει το 103 (Ενημέρωση)
 
 
 cmdDelete.Caption = mL_CapRes(107) '  "Διαγραφή"
 
 cmdAdd.Caption = mL_CapRes(108) '   "Νέα Εγγραφή"
 Command3.Caption = mL_CapRes(109) '   "Εκτύπωση εγγραφών "
 


 DBGrid1.Columns(0).Caption = mL_CapRes(140) '  "Κωδικός"
 DBGrid1.Columns(1).Caption = mL_CapRes(141) '  "ΠΕΡΙΓΡΑΦΗ"
 DBGrid1.Columns(2).Caption = mL_CapRes(142) '  "ΔΙΕΥΘΥΝΣΗ"
 
 DBGrid1.Columns(9).Caption = mL_CapRes(143) '
 DBGrid1.Columns(10).Caption = mL_CapRes(144) '
 DBGrid1.Columns(13).Caption = mL_CapRes(145) '
 ' prosoxh ta local settings na einai Agglika inomenon politeion
 DBGrid1.Columns(13).NumberFormat = "###.####" '
 '. 143 ";;Ταυτότητα"
 '. 144 ";;Συγκέντρωση"
 '. 145 ";;Υπόλοιπο"
 
 
Label4.Caption = mL_CapRes(149) ' 149 ";;Ταξινόμηση"

Option1.Caption = mL_CapRes(146) ' 146 ";;Κατά Περιγραφή"
Option2.Caption = mL_CapRes(147) ' 147 ";;Κατά Διεύθυνση"
Option3.Caption = mL_CapRes(148) ' 148 ";;Κατά Ταυτότητα"

Label2.Caption = mL_CapRes(135) ' 135 ";;Θέση Χημικοτεχνικού"
Label3.Caption = mL_CapRes(136) ' 136 ";;Κατάσταση Χημ/κού"
 
 
 
Combo1.List(0) = mL_CapRes(90) '90 ";;ΟΛΑ"
Combo1.List(1) = mL_CapRes(91) '91 ";;ΜΟΝΟ ΣΤΟ ΤΡΑΠΕΖΙ"
Combo1.List(2) = mL_CapRes(92) '92 ";;ΕΚΤΟΣ ΤΡΑΠΕΖΙΟΥ"
Combo1.text = Combo1.List(0)

Combo2.List(0) = mL_CapRes(93) '93 ";;ΟΛΑ"
Combo2.List(1) = mL_CapRes(94) '94 ";;ΜΟΝΟ ΣΚΟΝΕΣ"
Combo2.List(2) = mL_CapRes(95) '95  ";;ΜΟΝΟ dialymata"
Combo2.text = Combo1.List(0)
On Error GoTo 0

End Sub

Private Sub Form_Paint()
  Screen.MousePointer = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIForm1.Arxeia(10).Visible = True
MDIForm1.Syntages(20).Visible = True
MDIForm1.bohuhtika(30).Visible = True
MDIForm1.ejodos(40).Visible = True
End Sub









Private Sub MaskEdBox1_GotFocus()
 '&H80000005&   aspro
  Maskedbox1.BackColor = &HFF0000      'mple
  Maskedbox1.ForeColor = &H80000005    ' aspro
  Label1.Caption = mL_CapRes(166) ' "Πλητρολόγησε τον κωδικό που θέλεις.Π.χ. 17 ή 000017"

End Sub

Private Sub MaskEdBox1_LostFocus()
Dim b As String
b = LTrim(Maskedbox1.text)
Data1.RecordSource = "select *from ximitech where [kod]='" + Right$("000000" + b, 6) + "'"
       Data1.Refresh
       DBGrid1.Refresh
'End If
 
 '&H80000005&   aspro
  Maskedbox1.BackColor = Oldcolor_b   ' &HFF0000      'mple
  Maskedbox1.ForeColor = Oldcolor_f

  Label1.Caption = "  "
     
     
End Sub

Private Sub axrhstoMaskEdBox2_GotFocus()
  
 


End Sub

Private Sub MaskEdBox2_GotFocus()
  MaskedBox2.BackColor = &HFF0000      'mple
  MaskedBox2.ForeColor = &H80000005    'ασπρο
  Label1.Caption = mL_CapRes(167) ' "Πλητρολόγησε το όνομα του χημικοτεχνικού "



End Sub

Private Sub MaskEdBox2_LostFocus()
Dim b As String
b = UCase(LTrim(MaskedBox2.text))

Data1.RecordSource = "select *from ximitech where [perigrafh] like '" + b + "*';"

Data1.Refresh
DBGrid1.Refresh

  MaskedBox2.BackColor = Oldcolor_b   ' &HFF0000      'mple
  MaskedBox2.ForeColor = Oldcolor_f
  Label1.Caption = "  "


End Sub


Private Sub MaskEdBox3_GotFocus()
  MaskedBox3.BackColor = &HFF0000       'mple
  MaskedBox3.ForeColor = &H80000005    'ασπρο
  Label1.Caption = mL_CapRes(168) ' "Πλητρολόγησε την διεύθυνση του χημικοτεχνικού "
End Sub

Private Sub MaskEdBox3_LostFocus()
'  fff
 MaskedBox3.BackColor = &H80000005   'ασπρο
  MaskedBox3.ForeColor = &HFF0000          'mple
 Data1.RecordSource = "SELECT *FROM XIMITECH where addr_prot=" + str(Val(MaskedBox3)) + ";"
 Data1.Refresh
 DBGrid1.Refresh
 
End Sub

Private Sub Option1_Click()
   Combo1_Click
End Sub


Private Sub Option2_Click()
   Combo1_Click

End Sub


Private Sub Option3_Click()
   Combo1_Click

End Sub



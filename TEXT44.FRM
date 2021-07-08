VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form Text44 
   BackColor       =   &H00808080&
   Caption         =   "Form1"
   ClientHeight    =   6975
   ClientLeft      =   1800
   ClientTop       =   2295
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6975
   ScaleWidth      =   9660
   WindowState     =   2  'Maximized
   Begin VB.Data Data2 
      Caption         =   "eidos_bafhs"
      Connect         =   "Access"
      DatabaseName    =   "C:\talos\Recipies.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "pania"
      Top             =   0
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.CommandButton DHMAS 
      Caption         =   "DHMAS"
      Height          =   405
      Left            =   8040
      TabIndex        =   42
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   5760
      TabIndex        =   39
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton TreePoint 
      Caption         =   "TreePoint"
      Height          =   405
      Left            =   8040
      TabIndex        =   38
      Top             =   3075
      Width           =   1335
   End
   Begin VB.CommandButton FARBETEX 
      Caption         =   "farbetex"
      Height          =   405
      Left            =   8025
      TabIndex        =   37
      Top             =   4215
      Width           =   1335
   End
   Begin VB.CommandButton Polydye 
      Caption         =   "Datacolor"
      Height          =   405
      Left            =   8040
      TabIndex        =   35
      Top             =   1935
      Width           =   1335
   End
   Begin VB.CommandButton ColorTe 
      Caption         =   "ColorTec"
      Height          =   405
      Left            =   8040
      TabIndex        =   34
      Top             =   1395
      Width           =   1335
   End
   Begin VB.Data eidos_bafhs 
      Caption         =   "eidos_bafhs"
      Connect         =   "Access"
      DatabaseName    =   "C:\talos\Recipies.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "eidh_baf"
      Top             =   6480
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   5760
      TabIndex        =   29
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton damapong 
      Caption         =   "Clarion"
      Height          =   405
      Left            =   8040
      TabIndex        =   28
      Top             =   765
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "c:\talos\recipies.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   540
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "parametroi"
      Top             =   4200
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   5760
      TabIndex        =   26
      Top             =   4065
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   5760
      TabIndex        =   23
      Top             =   3645
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   390
      Left            =   4140
      TabIndex        =   22
      Top             =   4770
      Width           =   1440
   End
   Begin VB.TextBox Filename 
      DataField       =   "colormatchfilename"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2070
      TabIndex        =   20
      Text            =   "b"
      Top             =   5325
      Width           =   1620
   End
   Begin VB.TextBox Path 
      DataField       =   "colormatchdir"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   315
      TabIndex        =   18
      Text            =   "C:\talos"
      Top             =   5325
      Width           =   1620
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5625
      TabIndex        =   7
      Top             =   4770
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5760
      TabIndex        =   3
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5760
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Data pelates 
      Caption         =   "pelates"
      Connect         =   "Access"
      DatabaseName    =   "C:\talos\Recipies.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2100
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "pelates"
      Top             =   5880
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Silver"
      Height          =   420
      Left            =   8055
      TabIndex        =   0
      Top             =   2475
      Width           =   1320
   End
   Begin MSDBCtls.DBCombo DBCombo3 
      Bindings        =   "TEXT44.frx":0000
      Height          =   315
      Left            =   5745
      TabIndex        =   43
      Top             =   1260
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   16777215
      ForeColor       =   0
      ListField       =   "PANI"
      BoundColumn     =   "PANI"
      Text            =   "DBCombo3"
   End
   Begin MSDBCtls.DBCombo DBCombo2 
      Bindings        =   "TEXT44.frx":0014
      Height          =   315
      Left            =   5760
      TabIndex        =   32
      Top             =   3120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   16777215
      ForeColor       =   0
      ListField       =   "ONO_BAF"
      BoundColumn     =   "ONO_BAF"
      Text            =   "DBCombo2"
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "TEXT44.frx":002E
      Height          =   315
      Left            =   5730
      TabIndex        =   17
      Top             =   1635
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
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Artikle"
      Height          =   375
      Left            =   4500
      TabIndex        =   44
      Top             =   1290
      Width           =   1215
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dry Salt"
      Height          =   255
      Left            =   3960
      TabIndex        =   41
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   7620
      TabIndex        =   40
      Top             =   4485
      Width           =   2025
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   7620
      TabIndex        =   36
      Top             =   4095
      Width           =   2025
   End
   Begin VB.Label Label7 
      Height          =   255
      Index           =   4
      Left            =   2640
      TabIndex        =   33
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Procedure"
      Height          =   255
      Left            =   4200
      TabIndex        =   31
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Shade"
      Height          =   255
      Left            =   4200
      TabIndex        =   30
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Chemicals "
      Height          =   255
      Left            =   3960
      TabIndex        =   27
      Top             =   4065
      Width           =   1695
   End
   Begin VB.Label Label11 
      Caption         =   "Import from ColorMatch machine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4005
      TabIndex        =   25
      Top             =   255
      Width           =   4410
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Chemicals by hand %"
      Height          =   255
      Left            =   3960
      TabIndex        =   24
      Top             =   3645
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "Text FileName"
      Height          =   300
      Left            =   2100
      TabIndex        =   21
      Top             =   4980
      Width           =   1605
   End
   Begin VB.Label Label8 
      Caption         =   "Path of Text File"
      Height          =   300
      Left            =   345
      TabIndex        =   19
      Top             =   4980
      Width           =   1605
   End
   Begin VB.Label Label7 
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   16
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label7 
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   15
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label7 
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   14
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label7 
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   13
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   " "
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   2500
   End
   Begin VB.Label Label6 
      Caption         =   " "
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   2500
   End
   Begin VB.Label Label6 
      Caption         =   " "
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   2500
   End
   Begin VB.Label Label6 
      Caption         =   " "
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   2500
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      Caption         =   "Label5"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   555
      Width           =   3735
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   300
      Width           =   3750
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Weight of Fabrik"
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Liquor Ratio"
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "Text44"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim x, Perigrafh


Function FIND_XHM(M_KOD)


Dim x1, x2, m1, m2, p1, p2, mo
x1 = "[["
x2 = "[["
Perigrafh = "---"
FIND_XHM = "   "
x.MoveFirst
Do While Not x.eof
   
   If Trim(x("kataskeyas")) = M_KOD And x("skonh") = 0 And val2(x("MORFH")) > 0 Then '10
           x1 = x("kod")
           m1 = x("real_cons")   ' x("morfh")
           p1 = x("perigrafh")
           Exit Do
   End If '10
   x.MoveNext
Loop

If x1 <> "[[" Then  'an brike mono ena
      FIND_XHM = x1
      Perigrafh = p1
Else
    MsgBox M_KOD + " Not found."
End If





End Function

Private Sub Command1_Click()
' carolina coton
Dim mm, mm0, work, mydb, r, text, j, p, k, m, prosp, m_aa, m_Seira, m_ek, M_KOD, x_kod, ans
Command3.Enabled = False

'default μπάνιο
Command1.Enabled = False
Set work = Workspaces(0)
Set mydb = work.OpenDatabase("c:\TALOS\RECIPIES.MDB", False, False)
Set r = mydb.OpenRecordset("parametroi")
    Text1.text = r("MPANIO")
    Text2.text = r("default_baros_panioy")
     r.Close

'Data2.DatabaseName = Path.text
'Data2.RecordSource = filename.text
'Data2.Refresh
'DBGrid1.Columns(0).Name = "robolab"
'DBGrid1.Refresh

Set x = mydb.OpenRecordset("ximitech")
x.Index = "KOD"







'Set r = Data1.Recordset

Set text = OpenDatabase(Path.text, False, False, "Text;")
Set r = text.OpenRecordset(Filename.text)

Set j = mydb.OpenRecordset("Joblist")
j.Index = "ENTOLH"

Set p = mydb.OpenRecordset("prospau2")
p.Index = "ENTOLH"

'Set m = PROSMACB.Recordset
If r.RecordCount < 2 Then
   MsgBox "The file is empty."
   Command3.Enabled = True

   Exit Sub
End If
Do While Not r.eof
 If Mid$(r(0), 65, 2) = "NN" Then 'new recipy
     
     'work.BeginTrans
     
     m_entolh = Mid$(r(0), 25, 6)
     Label4.Caption = m_entolh
     
     
     
     mm = r.Bookmark
     On Error Resume Next
     mm0 = 0
     For k = 0 To 3
            r.MoveNext
            If Not r.eof Then
              If Len(Trim(Mid$(r(0), 25, 8))) > 2 Then
                 Label7(k).Caption = Val(Mid$(r(0), 33, 20)) '0.79432
                 Label6(k).Caption = Mid$(r(0), 25, 8) 'code
                 mm0 = mm0 + Val(Label7(k).Caption)
              End If
            End If
     Next
     Label7(4).Caption = mm0
     On Error GoTo 0
     
     
     
     r.Bookmark = mm
     
     
     
     If IsNumeric(m_entolh) Then
         m_entolh = Right("000000" + LTrim(str(Val(m_entolh))), 6) 'last
     Else
         
     End If
     
     m_entolh = Val(m_entolh)
          If m_entolh = 0 Then
      Else
          j.Seek "=", Right("000000" + LTrim(str(Val(m_entolh))), 6) 'last
     End If
     If j.NoMatch Or m_entolh = 0 Then
         j.MoveLast
         If IsNull(j("ENTOLH")) Then
             m = 1
         Else
             m = Val(j("ENTOLH")) + 1
         End If
         Label5.Caption = m
         Command2.Enabled = True
         Do While Command2.Enabled = True
           DoEvents
         Loop
         m_entolh = Right("000000" + LTrim(str(Val(m))), 6) 'last
         
       On Error Resume Next
         j.AddNew
         j("entolh") = m_entolh
         Form1.Caption = m_entolh
         j("kod_pel") = DBCombo1.text
         Form1.Caption = DBCombo1.text
         j("baros_pani") = Val(Text2.text)
         Form1.Caption = Val(Text2.text)
         j("sxesh_mpan") = Val(Text1.text)
         Form1.Caption = Val(Text2.text)
         j("kod_erg") = 0
         Form1.Caption = "kod_erg"
         j("xhmika") = Val(Text3.text)
          
         j("apoxrvsh") = left(Text5.text, 20)
         j("eidos_baf") = left(DBCombo2.text, 15)
         j.update
         On Error GoTo 0
         
         
         
         Form1.Caption = "J/update"
         prosp = 1
      Else
         
         p.Seek "=", j("entolh") 'BRISKO TON MEGALITERO ARIUMO PROSPAUIAS
         m_aa = p("AA_PROSPAU")
         Form1.Caption = "m_aa"
         Do While p.ENTOLH = m_entolh And Not p.eof
            If p("AA_PROSPAU") > m_aa Then
               m_aa = p("AA_PROSPAU")
               Form1.Caption = "r/m_aa2"
            End If
            p.MoveNext
         Loop
         prosp = Val(m_aa)
         Form1.Caption = "r/m_aa3"
     End If
     
     For k = 0 To 3
            Label6(k).Caption = " "
            Label7(k).Caption = " "
     Next
     
          r.MoveNext
          m_Seira = 0
     Do While Not r.eof
         If Mid$(r(0), 65, 2) = "NN" Then Exit Do 'new recipy
         
         m_ek = Val(Mid$(r(0), 33, 20)) '0.79432
         Form1.Caption = "r/m_ek"
         M_KOD = Mid$(r(0), 25, 8) 'code
         Form1.Caption = "r/code"
         If m_ek <= 0 Or Len(Trim(M_KOD)) < 2 Then
           x_kod = " "
         Else
           x_kod = find_kod(M_KOD, m_ek)
         End If
         
         Form1.Caption = "r/x_kod:"
       
       If left(x_kod, 1) <> " " Then
         m_Seira = m_Seira + 1
         p.AddNew
         p("entolh") = m_entolh 'last
'         Form1.Caption = "p/" + m_entolh
         p("aa_prospau") = Right("00" + LTrim(str(prosp)), 2)
 '        Form1.Caption = "p/aa"
         p("kod_pel") = DBCombo1.text
  '       Form1.Caption = "p/kodpel"
         p("baros_pani") = Val(Text2.text)
   '      Form1.Caption = "p/baros_pani"
         p("sxesh_mpan") = Val(Text1.text)
    '     Form1.Caption = "p/sxesh_mpan"
         p("ek") = m_ek
     '    Form1.Caption = "p/m_ek"
         p("gl") = 0
      '   Form1.Caption = "p/gl"
         p("kod") = x_kod
       '  Form1.Caption = "p/x_kod"
         p("SEIRA") = m_Seira
        ' Form1.Caption = "p/m_seira"
         p("ayjon") = m_Seira
         p("apoxrvsh") = left(Text5.text, 20)
        ' p("eidos_baf") = j("eidos_baf")
         'Form1.Caption = "p/ayjon"
         p("status") = 0
         p("APOR") = 0
         p("alati") = Val(Text4.text)
         
         p.update
         Form1.Caption = "p/update"
         If m_Seira <= 4 And Not IsNull(x("perigrafh")) Then
            Label6(m_Seira - 1).Caption = Perigrafh
            Label7(m_Seira - 1).Caption = Format(m_ek, "#0.0000")
         End If
        End If
         
         r.MoveNext
     Loop
     
     ' salt
       
         
'         m_ek = Val(Text4.text) '0.79432
 '        m_kod = "SALT...." 'code
  '
   '      If m_ek <= 0 Or Len(Trim(m_kod)) < 2 Then
    '       x_kod = " "
     '    Else
      '     x_kod = find_kod(m_kod, m_ek)
       '  End If
       
       
       
       
'       If Val(Text4.text) > 0 And x_kod <> " " Then
'          m_Seira = m_Seira + 1
 '         p.AddNew
  '       p("entolh") = m_entolh 'last
   '      p("aa_prospau") = Right("00" + LTrim(Str(prosp)), 2)
    '     p("kod_pel") = DBCombo1.text
'         p("baros_pani") = Val(Text2.text)
 '        p("sxesh_mpan") = Val(Text1.text)
  '       p("ek") = 0
   '      p("gl") = m_ek
    '     p("kod") = x_kod
     '    p("SEIRA") = m_Seira
      '   p("ayjon") = m_Seira
       '  p("status") = 0
        ' p("APOR") = 0
         'p.update
'       End If
     
     
     
     ans = MsgBox("Save this recipe ?", vbYesNo)
     If ans = vbYes Then
        'work.CommitTrans
     Else
        'work.Rollback
        mydb.Execute "delete *from prospau2 where entolh='" + m_entolh + "'"
        mydb.Execute "delete *from joblist where entolh='" + m_entolh + "'"
     End If
     
     
     For k = 0 To 3
            Label6(k).Caption = " "
            Label7(k).Caption = " "
     Next
     
     
     
     If r.eof Then Exit Do
  Else
     r.MoveNext
  End If
Loop
r.Close
ans = MsgBox(" Do you want to delete the text file ? ", vbYesNo, "The import is finished.")
If ans = vbYes Then
   'r.Close
   On Error GoTo cantdelete
   FileCopy "c:\talos\emptytxt.txt", Path.text + "\" + Filename.text
   On Error GoTo 0
End If

Unload Me
Exit Sub

cantdelete:
MsgBox "The file can not deleted. Another user uses the same file"
Resume Next

End Sub

Private Sub ColorTe_Click()
'   Colibri
Dim mm, mm0, work, mydb, r, text, j, p, k, m, prosp, m_aa, m_Seira, m_ek, M_KOD, x_kod, ans
Dim DUM, nFile


For nFile = 1 To 30


Command3.Enabled = False
'If Len(Dir("e:\My Documents\" + LTrim(Str(nFile)) + ".prn")) < 4 Then GoTo 9088
If Len(Dir("e:\MyDocu~1\" + LTrim(str(nFile)) + ".prn")) < 4 Then GoTo 9088



DUM = Shell("c:\talos\merge.bat  " + LTrim(str(nFile)), 0)
MilSec 1000
'default μπάνιο
Command1.Enabled = False
Set work = Workspaces(0)
Set mydb = work.OpenDatabase("c:\TALOS\RECIPIES.MDB", False, False)
Set r = mydb.OpenRecordset("parametroi")
    Text1.text = r("MPANIO")
    Text2.text = r("default_baros_panioy")
     r.Close


Set x = mydb.OpenRecordset("ximitech")
x.Index = "KOD"

Set text = OpenDatabase(Path.text, True, False, "Text;")
Set r = text.OpenRecordset(Filename.text)
If r.RecordCount < 5 Then
   mydb.Close
   text.Close
   GoTo 9088
End If
'debug'





Set j = mydb.OpenRecordset("Joblist")
j.Index = "ENTOLH"

Set p = mydb.OpenRecordset("prospau2")
p.Index = "ENTOLH"

'Set m = PROSMACB.Recordset
If r.RecordCount < 2 Then
   MsgBox "The file is empty."
   Command3.Enabled = True

   Exit Sub
End If

Do While Not r.eof   'until find 1 recipe
 If Mid$(r(0), 7, 5) = "Match" Or Mid$(r(0), 7, 5) = "Corre" Then 'new recipy
     
     'work.BeginTrans
     
     m_entolh = Mid$(r(0), 18, 18)
     Label4.Caption = LTrim(str(nFile)) + "  " + m_entolh
     Text5.text = Mid$(r(0), 18, 18)
     
     
     mm = r.Bookmark
     On Error Resume Next
     mm0 = 0
     For k = 0 To 3
            r.MoveNext
            If Not r.eof Then
              If Val(Mid$(r(0), 27, 8)) > 0 And Val(Mid$(r(0), 1, 4)) > 0 Then
                 Label7(k).Caption = Val(Mid$(r(0), 27, 8)) '0.79432
                 Label6(k).Caption = Mid$(r(0), 1, 20) 'code
                 mm0 = mm0 + Val(Label7(k).Caption)
              End If
            End If
     Next
     Label7(4).Caption = mm0
     On Error GoTo 0
     
     
     
     r.Bookmark = mm
     
     
     
     If IsNumeric(m_entolh) Then
         m_entolh = Right("000000" + LTrim(str(Val(m_entolh))), 6) 'last
     Else
         
     End If
     
     m_entolh = Val(m_entolh)
          If m_entolh = 0 Then
      Else
          j.Seek "=", Right("000000" + LTrim(str(Val(m_entolh))), 6) 'last
     End If
     If j.NoMatch Or m_entolh = 0 Then
         j.MoveLast
         If IsNull(j("ENTOLH")) Then
             m = 1
         Else
             m = Val(j("ENTOLH")) + 1
         End If
         Label5.Caption = m
         Command2.Enabled = True
         Do While Command2.Enabled = True
           DoEvents
         Loop
         m_entolh = Right("000000" + LTrim(str(Val(m))), 6) 'last
         
       On Error Resume Next
         j.AddNew
         j("entolh") = m_entolh
         Form1.Caption = m_entolh
         j("kod_pel") = DBCombo1.text
         Form1.Caption = DBCombo1.text
         j("baros_pani") = Val(Text2.text)
         Form1.Caption = Val(Text2.text)
         j("sxesh_mpan") = Val(Text1.text)
         Form1.Caption = Val(Text2.text)
         j("kod_erg") = 0
         Form1.Caption = "kod_erg"
         j("xhmika") = Val(Text3.text)
          
         j("apoxrvsh") = left(Text5.text, 20)
         j("eidos_baf") = left(DBCombo2.text, 15)
         j.update
         On Error GoTo 0
         
         
         
         Form1.Caption = "J/update"
         prosp = 1
      Else
         
         p.Seek "=", j("entolh") 'BRISKO TON MEGALITERO ARIUMO PROSPAUIAS
         m_aa = p("AA_PROSPAU")
         Form1.Caption = "m_aa"
         Do While p.ENTOLH = m_entolh And Not p.eof
            If p("AA_PROSPAU") > m_aa Then
               m_aa = p("AA_PROSPAU")
               Form1.Caption = "r/m_aa2"
            End If
            p.MoveNext
         Loop
         prosp = Val(m_aa)
         Form1.Caption = "r/m_aa3"
     End If
     
     For k = 0 To 3
            Label6(k).Caption = " "
            Label7(k).Caption = " "
     Next
     
          r.MoveNext
          m_Seira = 0
     Do While Not r.eof   ' loop sta systatika ths syntaghs
         If Val(Mid$(r(0), 27, 8)) = 0 And Val(Mid$(r(0), 1, 4)) = 0 Then Exit Do 'new recipy
         
         m_ek = Val(Mid$(r(0), 27, 8))  '0.79432
         Form1.Caption = "r/m_ek"
         M_KOD = Mid$(r(0), 1, 4) 'code
         Form1.Caption = "r/code"
         If m_ek <= 0 Or Len(Trim(M_KOD)) < 2 Then
           x_kod = " "
         Else
           x_kod = find_kod(M_KOD, m_ek)
         End If
         
         Form1.Caption = "r/x_kod:"
       
       If left(x_kod, 1) <> " " And Val(x_kod) > 0 Then
         m_Seira = m_Seira + 1
         p.AddNew
         p("entolh") = m_entolh 'last
'         Form1.Caption = "p/" + m_entolh
         p("aa_prospau") = Right("00" + LTrim(str(prosp)), 2)
 '        Form1.Caption = "p/aa"
         p("kod_pel") = DBCombo1.text
  '       Form1.Caption = "p/kodpel"
         p("baros_pani") = Val(Text2.text)
   '      Form1.Caption = "p/baros_pani"
         p("sxesh_mpan") = Val(Text1.text)
    '     Form1.Caption = "p/sxesh_mpan"
         p("ek") = m_ek
     '    Form1.Caption = "p/m_ek"
         p("gl") = 0
      '   Form1.Caption = "p/gl"
         p("kod") = x_kod
       '  Form1.Caption = "p/x_kod"
         p("SEIRA") = m_Seira
        ' Form1.Caption = "p/m_seira"
         p("alati") = Val(Text4.text)
         p("ayjon") = m_Seira
         p("apoxrvsh") = left(Text5.text, 20)
        ' p("eidos_baf") = j("eidos_baf")
         'Form1.Caption = "p/ayjon"
         p("status") = 0
         p("APOR") = 0
         p.update
         Form1.Caption = "p/update"
         If m_Seira <= 4 And Not IsNull(x("perigrafh")) Then
            Label6(m_Seira - 1).Caption = Perigrafh
            Label7(m_Seira - 1).Caption = Format(m_ek, "#0.0000")
         End If
        End If
         
         r.MoveNext
     Loop
     
     ' salt
       
         
      '   m_ek = Val(Text4.text) '0.79432
      '   m_kod = "SALT...." 'code
         
      '   If m_ek <= 0 Or Len(Trim(m_kod)) < 2 Then
      '     x_kod = " "
      '   Else
      '     x_kod = find_kod(m_kod, m_ek)
      '   End If
      '
       
       
       
      ' If Val(Text4.text) > 0 And x_kod <> " " Then
      '    m_Seira = m_Seira + 1
      '    p.AddNew
      '   p("entolh") = m_entolh 'last
      '   p("aa_prospau") = Right("00" + LTrim(Str(prosp)), 2)
      '   p("kod_pel") = DBCombo1.text
      '   p("baros_pani") = Val(Text2.text)
      '   p("sxesh_mpan") = Val(Text1.text)
      '   p("ek") = 0
      '   p("gl") = m_ek
      '   p("kod") = x_kod
      '   p("SEIRA") = m_Seira
      '   p("ayjon") = m_Seira
      '   p("status") = 0
      '   p("APOR") = 0
      '   p.update
      ' End If
     
     
     
     ans = MsgBox("Save this recipe ?", vbYesNo)
     If ans = vbYes Then
        'work.CommitTrans
     Else
        'work.Rollback
        mydb.Execute "delete *from prospau2 where entolh='" + m_entolh + "'"
        mydb.Execute "delete *from joblist where entolh='" + m_entolh + "'"
     End If
   
     
     
     For k = 0 To 3
            Label6(k).Caption = " "
            Label7(k).Caption = " "
     Next
     
     Exit Do
     
     'If r.EOF Then Exit Do
  Else
     r.MoveNext
  End If

Loop
r.Close
ans = MsgBox(" Do you want to delete the text file ? ", vbYesNo, "The import is finished.")
If ans = vbYes Then
   'r.Close
   On Error GoTo cantdelete
       Kill "e:\MyDocu~1\" + LTrim(str(nFile)) + ".prn"
   On Error GoTo 0
End If

9088
Next nFile



Unload Me
Exit Sub

cantdelete:
MsgBox "The file can not deleted. Another user uses the same file"
Resume Next

End Sub

Private Sub oldPolydye_Click()
'  polydye  / datacolor
Dim mm, mm0, work, mydb, r, text, j, p, k, m, prosp, m_aa, m_Seira, m_ek, M_KOD, x_kod, ans
Dim DUM, nFile, nspace, nc


For nFile = 1 To 5


Command3.Enabled = False
 
If Len(Dir("a:\" + LTrim(str(nFile)) + ".prn")) < 4 Then GoTo 9081



DUM = Shell("c:\talos\merge.bat  " + LTrim(str(nFile)), 0)
MilSec 1000
'default μπάνιο
Command1.Enabled = False


'DUM = Shell("c:\talos\merge.bat  1", 1)   '+ LTrim(Str(nFile))

Set work = Workspaces(0)
Set mydb = work.OpenDatabase("c:\TALOS\RECIPIES.MDB", False, False)
Set r = mydb.OpenRecordset("parametroi")
    Text1.text = r("MPANIO")
    Text2.text = r("default_baros_panioy")
     r.Close


Set x = mydb.OpenRecordset("ximitech")
x.Index = "KOD"

Set text = OpenDatabase(Path.text, True, False, "Text;")
Set r = text.OpenRecordset(Filename.text)
If r.RecordCount < 5 Then
   mydb.Close
   text.Close
   GoTo 9081
End If

Set j = mydb.OpenRecordset("Joblist")
j.Index = "ENTOLH"

Set p = mydb.OpenRecordset("prospau2")
p.Index = "ENTOLH"

'Set m = PROSMACB.Recordset
If r.RecordCount < 2 Then
   MsgBox "The file is empty."
   Command3.Enabled = True

   Exit Sub
End If

Do While Not r.eof   'until find 1 recipe
 If IsNull(r(0)) Then
    r.MoveNext: GoTo 9082
 End If
 
 If InStr(1, r(0), "Recipe") >= 1 And InStr(1, r(0), "Recipe") < 12 Then 'new recipy


' If Mid$(r(0), 11, 6) = "Liquor" Then 'new recipy
     
     'work.BeginTrans
     
     'm_entolh = Mid$(r(0), 18, 18)
     m_entolh = Mid$(r(0), 17, 10)
     Label4.Caption = LTrim(str(nFile)) + "  " + m_entolh
     Text5.text = Mid$(r(0), 17, 10)
     
     
     mm = r.Bookmark
     On Error Resume Next
     mm0 = 0
     r.Move 1
     nc = -1
     
'On Error GoTo 0
     
     For k = 0 To 18
            r.MoveNext
            If Not IsNull(r(0)) Then
              If Not r.eof And InStr(10, r(0), "  %  ") Then
                 nc = nc + 1
                 nspace = InStr(7, r(0), " ")
                 Label6(nc).Caption = Mid$(r(0), nspace, 20) 'code  'If Val(Mid$(r(0), 27, 8)) > 0 And Val(Mid$(r(0), 1, 4)) > 0 Then
                 nspace = InStr(10, r(0), "  %  ")
                 Label7(nc).Caption = Val(Mid$(r(0), nspace - 10, 10)) '0.79432
                 mm0 = mm0 + Val(Label7(nc).Caption)
              End If
            End If
     Next
     Label7(4).Caption = mm0
     On Error GoTo 0
     
     
     
     r.Bookmark = mm
     
     r.Move 1
     
     
     If IsNumeric(m_entolh) Then
         m_entolh = Right("000000" + LTrim(str(Val(m_entolh))), 6) 'last
     Else
         
     End If
     
     m_entolh = Val(m_entolh)
          If m_entolh = 0 Then
      Else
          j.Seek "=", Right("000000" + LTrim(str(Val(m_entolh))), 6) 'last
     End If
     If j.NoMatch Or m_entolh = 0 Then
         j.MoveLast
         If IsNull(j("ENTOLH")) Then
             m = 1
         Else
             m = Val(j("ENTOLH")) + 1
         End If
         Label5.Caption = m
         Command2.Enabled = True
         Do While Command2.Enabled = True
           DoEvents
         Loop
         m_entolh = Right("000000" + LTrim(str(Val(m))), 6) 'last
         
       On Error Resume Next
         j.AddNew
         j("entolh") = m_entolh
         Form1.Caption = m_entolh
         j("kod_pel") = DBCombo1.text
         Form1.Caption = DBCombo1.text
         j("baros_pani") = Val(Text2.text)
         Form1.Caption = Val(Text2.text)
         j("sxesh_mpan") = Val(Text1.text)
         Form1.Caption = Val(Text2.text)
         j("kod_erg") = 0
         Form1.Caption = "kod_erg"
         j("xhmika") = Val(Text3.text)
          
         j("apoxrvsh") = left(Text5.text, 20)
         j("eidos_baf") = left(DBCombo2.text, 15)
         j.update
         On Error GoTo 0
         
         
         
         Form1.Caption = "J/update"
         prosp = 1
      Else
         
         p.Seek "=", j("entolh") 'BRISKO TON MEGALITERO ARIUMO PROSPAUIAS
         m_aa = p("AA_PROSPAU")
         Form1.Caption = "m_aa"
         Do While p.ENTOLH = m_entolh And Not p.eof
            If p("AA_PROSPAU") > m_aa Then
               m_aa = p("AA_PROSPAU")
               Form1.Caption = "r/m_aa2"
            End If
            p.MoveNext
         Loop
         prosp = Val(m_aa)
         Form1.Caption = "r/m_aa3"
     End If
     
     For k = 0 To 3
            Label6(k).Caption = " "
            Label7(k).Caption = " "
     Next
     
          r.MoveNext
          m_Seira = 0
     Do While Not r.eof   ' loop sta systatika ths syntaghs
       If Not IsNull(r(0)) Then   'If Val(Mid$(r(0), 27, 8)) = 0 And Val(Mid$(r(0), 1, 4)) = 0 Then Exit Do 'new recipy
       If InStr(10, r(0), "  %  ") > 0 Then
         nspace = InStr(1, r(0), " ")
         M_KOD = Mid$(r(0), 1, nspace - 1) 'code
         
         nspace = InStr(10, r(0), "  %  ")
         m_ek = Val(Mid$(r(0), nspace - 10, 10)) '0.79432
        ' m_ek = Val(Mid$(r(0), 27, 8))  '0.79432
         
         If m_ek <= 0 Or Len(Trim(M_KOD)) < 2 Then
           x_kod = " "
         Else
           x_kod = find_kod(M_KOD, m_ek)
         End If
         
       
       If left(x_kod, 1) <> " " And Val(x_kod) > 0 Then
         m_Seira = m_Seira + 1
         p.AddNew
         p("entolh") = m_entolh 'last
         p("aa_prospau") = Right("00" + LTrim(str(prosp)), 2)
         p("kod_pel") = DBCombo1.text
         p("baros_pani") = Val(Text2.text)
         p("sxesh_mpan") = Val(Text1.text)
         p("ek") = m_ek
         p("gl") = 0
         p("kod") = x_kod
         p("SEIRA") = m_Seira
         p("alati") = Val(Text4.text)
         p("ayjon") = m_Seira
         p("apoxrvsh") = left(Text5.text, 20)
         p("status") = 0
         p("APOR") = 0
         p.update
         Form1.Caption = "p/update"
         If m_Seira <= 4 And Not IsNull(x("perigrafh")) Then
            Label6(m_Seira - 1).Caption = Perigrafh
            Label7(m_Seira - 1).Caption = Format(m_ek, "#0.0000")
         End If
        End If
       End If ' instr
       End If ' not isnull
         r.MoveNext
     Loop
     
     ' salt
       
         
      '   m_ek = Val(Text4.text) '0.79432
      '   m_kod = "SALT...." 'code
         
      '   If m_ek <= 0 Or Len(Trim(m_kod)) < 2 Then
      '     x_kod = " "
      '   Else
      '     x_kod = find_kod(m_kod, m_ek)
      '   End If
      '
       
       
       
      ' If Val(Text4.text) > 0 And x_kod <> " " Then
      '    m_Seira = m_Seira + 1
      '    p.AddNew
      '   p("entolh") = m_entolh 'last
      '   p("aa_prospau") = Right("00" + LTrim(Str(prosp)), 2)
      '   p("kod_pel") = DBCombo1.text
      '   p("baros_pani") = Val(Text2.text)
      '   p("sxesh_mpan") = Val(Text1.text)
      '   p("ek") = 0
      '   p("gl") = m_ek
      '   p("kod") = x_kod
      '   p("SEIRA") = m_Seira
      '   p("ayjon") = m_Seira
      '   p("status") = 0
      '   p("APOR") = 0
      '   p.update
      ' End If
     
     
     
     ans = MsgBox("Save this recipe ?", vbYesNo)
     If ans = vbYes Then
        'work.CommitTrans
     Else
        'work.Rollback
        mydb.Execute "delete *from prospau2 where entolh='" + m_entolh + "'"
        mydb.Execute "delete *from joblist where entolh='" + m_entolh + "'"
     End If
   
     
     
     For k = 0 To 3
            Label6(k).Caption = " "
            Label7(k).Caption = " "
     Next
     
     Exit Do
     
     'If r.EOF Then Exit Do
  Else

     r.MoveNext
  End If
9082
Loop
r.Close
ans = MsgBox(" Do you want to delete the text file ? ", vbYesNo, "The import is finished.")
If ans = vbYes Then
   On Error GoTo cantdelete
      ' Kill "e:\MyDocu~1\" + LTrim(Str(nFile)) + ".prn"
   On Error GoTo 0
End If

9081
Next nFile



Unload Me
Exit Sub

cantdelete:
MsgBox "The file can not deleted. Another user uses the same file"
Resume Next

End Sub

Private Sub DHMAS_Click()
' DHMAS
'1 η γραμμή
'6ψήφιος αύξων αριθμός δείγματος - 20ψήφια περιγραφή απόχρωσης

'2 η γραμμή
'2ψήφιος κωδικός είδους νήματος - 2ψήφιος κωδικός σύνθεσης νήματος - 3ψήφιος κωδικός ποιότητας νήματος

'3 η γραμμή
'4ψήφιο (##.#) πεδίο σχέσης μπάνιου -  4ψήφιο (##.#) πεδίο χημικών με το χέρι

'4 η γραμμή
'20ψήφιο πεδίο παρατηρήσεων

'5η-12η γραμμή
'13ψήφιος κωδικός χρώματος - 35ψήφια περιγραφή χρώματος - 6ψήφιος κωδικός μπουκαλιού ΤΑΛΟΣ - 8ψήφιο (##.#####) ποσοστό χρώματος

'Το αρχείο που επισυνάπτεται έχει μεταφερθεί από το Croma με το batch file "merge".

'Έχω χρησιμοποιήσει ως κωδικό μπουκαλιού ΤΑΛΟΣ τον κωδικό του χρώματος και όχι τη θέση του μπουκαλιού στο τραπέζι. Αν αυτό δεν είναι σωστό ειδοποιήστε με εγκαίρως να το αλλάξω.

'Πέραν της εισαγωγής των συνταγών στο ΤΑΛΟΣ από το Croma επιθυμούμε και την εξαγωγή με το ίδιο format της συνταγής που τελικά ζύγισε το ΤΑΛΟΣ στην εφαρμογή Croma.




Dim DUM, mcounter, m_parat
Dim mm, mm0, work, mydb, r, text, j, p, k, m, prosp, m_aa, m_Seira, m_ek, M_KOD, x_kod, ans
Command3.Enabled = False
Polydye.Enabled = False
ColorTe.Enabled = False
damapong.Enabled = False
Command1.Enabled = False
DHMAS.Enabled = False

Me.Caption = "The position for chemicals is 71"
'default μπάνιο
Text4.text = 1
Screen.MousePointer = 11
 Label15.Caption = "Recognize ID=CHEMIC"
DUM = Shell("c:\talos\merge.bat")
MilSec 10000
Screen.MousePointer = 1
Command1.Enabled = False
Set work = Workspaces(0)
Set mydb = work.OpenDatabase("c:\TALOS\RECIPIES.MDB", False, False)
Set r = mydb.OpenRecordset("parametroi")
    Text1.text = r("MPANIO")
    Text2.text = r("default_baros_panioy")
     r.Close

'Data2.DatabaseName = Path.text
'Data2.RecordSource = filename.text
'Data2.Refresh
'DBGrid1.Columns(0).Name = "robolab"
'DBGrid1.Refresh

Set x = mydb.OpenRecordset("ximitech")
x.Index = "KOD"







'Set r = Data1.Recordset

Set text = OpenDatabase(Path.text, False, False, "Text;")
Set r = text.OpenRecordset(Filename.text)

Set j = mydb.OpenRecordset("Joblist")
j.Index = "ENTOLH"

Set p = mydb.OpenRecordset("prospau2")
p.Index = "ENTOLH"

'Set m = PROSMACB.Recordset
If r.RecordCount < 2 Then
   MsgBox "The file is empty."
   Command3.Enabled = True

   Exit Sub
End If
Do While Not r.eof
 
 
 
 
 
 If IsNull(r(0)) Then Exit Sub
 
' r.MovePrevious
 
' If Len(r(0)) > 20 And Len(Mid$(r(0), 51, 6)) = 0 Then 'new recipy
     
If Val(Mid$(r(0), 1, 6)) > 0 Then 'new recipy
     'work.BeginTrans
     
     
     m_entolh = Mid$(r(0), 1, 6)
     Label4.Caption = m_entolh
   '  r.MoveNext
     Text5.text = Mid$(r(0), 7, 15)
     If r.eof Then Exit Do
     
     mm = r.Bookmark
     On Error Resume Next
     mm0 = 0
     mcounter = 0
     'r.MoveNext
     For k = 0 To 11
            r.MoveNext
            If r.eof Then Exit For
        If Not IsNull(r(0)) Then
            If Mid$(r(0), 51, 6) = "WATER" Then Exit For
            If Not r.eof Then
              If Len(LTrim(Trim(Mid$(r(0), 51, 6)))) > 2 And Val(Mid$(r(0), 58, 10)) > 0 Then
                 Label7(mcounter).Caption = Val(Mid$(r(0), 58, 10)) '0.79432
                 Label6(mcounter).Caption = Mid$(r(0), 51, 6) 'code
                 mm0 = mm0 + Val(Label7(mcounter).Caption)
                 mcounter = mcounter + 1
                 
              End If
            End If
        End If
        If k = 1 Then
            Text1.text = Mid$(r(0), 1, 5)
            Text3.text = Mid$(r(0), 6, 6)
        End If
        If k = 2 Then
            m_parat = "."
           If Not IsNull(r(0)) Then
              m_parat = left(r(0), 20)
           End If
        End If
     Next
     If r.eof Then Exit Do
     
'     Exit Sub
     
     
     Label7(4).Caption = mm0
     On Error GoTo 0
     
     
     
     r.Bookmark = mm
     
     
     
     If IsNumeric(m_entolh) Then
         m_entolh = Right("000000" + LTrim(str(Val(m_entolh))), 6) 'last
     Else
         
     End If
     
     m_entolh = Val(m_entolh)
          If m_entolh = 0 Then
      Else
          j.Seek "=", Right("000000" + LTrim(str(Val(m_entolh))), 6) 'last
     End If
     If j.NoMatch Or m_entolh = 0 Then
         j.MoveLast
         If IsNull(j("ENTOLH")) Then
             m = 1
         Else
             m = Val(j("ENTOLH")) + 1
         End If
         Label5.Caption = m
         Command2.Enabled = True
         Do While Command2.Enabled = True
           DoEvents
         Loop
         m_entolh = Right("000000" + LTrim(str(Val(m))), 6) 'last
         
       On Error Resume Next
         j.AddNew
         j("entolh") = m_entolh
         Me.Caption = m_entolh
         j("kod_pel") = DBCombo1.text
         Me.Caption = DBCombo1.text
         j("baros_pani") = Val(Text2.text)
         Me.Caption = Val(Text2.text)
         j("sxesh_mpan") = Val(Text1.text)
         Me.Caption = Val(Text2.text)
         j("kod_erg") = 0
         Me.Caption = "kod_erg"
         j("xhmika") = Val(Text3.text)
         j("parat") = m_parat
         j("apoxrvsh") = left(Text5.text, 20)
         j("eidos_baf") = left(DBCombo2.text, 15)
         j.update
         On Error GoTo 0
         
         
         
         Me.Caption = "J/update"
         prosp = 1
      Else
         
         p.Seek "=", j("entolh") 'BRISKO TON MEGALITERO ARIUMO PROSPAUIAS
         m_aa = p("AA_PROSPAU")
         Me.Caption = "m_aa"
         Do While p.ENTOLH = m_entolh And Not p.eof
            If p("AA_PROSPAU") > m_aa Then
               m_aa = p("AA_PROSPAU")
               Me.Caption = "r/m_aa2"
            End If
            p.MoveNext
         Loop
         prosp = Val(m_aa)
         Me.Caption = "r/m_aa3"
     End If
     
     For k = 0 To 3
            Label6(k).Caption = " "
            Label7(k).Caption = " "
     Next
     
          r.MoveNext
          r.MoveNext
          r.MoveNext
          r.MoveNext

          m_Seira = 0
     Do While Not r.eof
         If IsNull(r(0)) Then Exit Do 'new recipy
         If Mid$(r(0), 1, 2) = "--" Then Exit Do
         
         m_ek = Val(Mid$(r(0), 58, 10)) '0.79432
         Me.Caption = "r/m_ek"
         M_KOD = Trim(Mid$(r(0), 51, 6)) 'code
         Me.Caption = "r/code"
         If m_ek <= 0 Or Len(Trim(M_KOD)) < 2 Then
           x_kod = " "
         Else
           x_kod = find_kod(M_KOD, m_ek)
         End If
                '           x_kod = " "
         Me.Caption = "r/x_kod:"
       
       If left(x_kod, 1) <> " " Then
         m_Seira = m_Seira + 1
         p.AddNew
         p("entolh") = m_entolh 'last
'         Form1.Caption = "p/" + m_entolh
         p("aa_prospau") = Right("00" + LTrim(str(prosp)), 2)
 '        Form1.Caption = "p/aa"
         p("kod_pel") = DBCombo1.text
  '       Form1.Caption = "p/kodpel"
         p("baros_pani") = Val(Text2.text)
   '      Form1.Caption = "p/baros_pani"
         p("sxesh_mpan") = Val(Text1.text)
    '     Form1.Caption = "p/sxesh_mpan"
         p("ek") = m_ek
     '    Form1.Caption = "p/m_ek"
         p("gl") = 0
      '   Form1.Caption = "p/gl"
         p("kod") = x_kod
       '  Form1.Caption = "p/x_kod"
         p("SEIRA") = m_Seira
        ' Form1.Caption = "p/m_seira"
         p("ayjon") = m_Seira
         p("apoxrvsh") = left(Text5.text, 20)
        ' p("eidos_baf") = j("eidos_baf")
         'Form1.Caption = "p/ayjon"
          p("alati") = Val(Text6.text)
         p("status") = 0
         p("APOR") = 0
         p.update
         Me.Caption = "p/update"
         If m_Seira <= 4 And Not IsNull(x("perigrafh")) Then
            Label6(m_Seira - 1).Caption = Perigrafh
            If Not IsNull(x("real_cons")) Then
               If x("real_cons") > 100 * m_ek Then
                  Label7(m_Seira - 1).BackColor = vbRed
               End If
            Else
                Label7(m_Seira - 1).BackColor = &H8000000F
            End If
            
            Label7(m_Seira - 1).Caption = Format(m_ek, "#0.0000")
         End If
        End If
         
         r.MoveNext
     Loop
     
     ' salt
       
         
         m_ek = Val(Text4.text) '0.79432
         M_KOD = "CHEMIC" 'code
         
         If M_KOD = "CHEMIC" Then
           x_kod = FIND_XHM(M_KOD)
         ElseIf m_ek <= 0 Or Len(Trim(M_KOD)) < 2 Then
           x_kod = " "
         Else
           x_kod = find_kod(M_KOD, m_ek)
         End If
       
       
       
       
       If Val(Text4.text) > 0 And Len(Trim(x_kod)) > 2 Then
          m_Seira = m_Seira + 1
          p.AddNew
         p("entolh") = m_entolh 'last
         p("aa_prospau") = Right("00" + LTrim(str(prosp)), 2)
         p("kod_pel") = DBCombo1.text
         p("baros_pani") = Val(Text2.text)
         p("sxesh_mpan") = Val(Text1.text)
         p("ek") = 0
         p("gl") = m_ek
         p("kod") = x_kod
         p("SEIRA") = m_Seira
         p("ayjon") = m_Seira
         p("status") = 0
         p("APOR") = 0
         p.update
       End If
     
     
     
     ans = MsgBox("Save this recipe ?", vbYesNo)
     If ans = vbYes Then
        'work.CommitTrans
     Else
        'work.Rollback
        mydb.Execute "delete *from prospau2 where entolh='" + m_entolh + "'"
        mydb.Execute "delete *from joblist where entolh='" + m_entolh + "'"
     End If
     
     
     For k = 0 To 3
            Label6(k).Caption = " "
            Label7(k).Caption = " "
     Next
     
     
     r.MoveNext
     If r.eof Then Exit Do
  Else
     r.MoveNext
  End If
  ' r.MoveNext
Loop
r.Close
ans = MsgBox(" Do you want to delete the text file ? ", vbYesNo, "The import is finished.")
If ans = vbYes Then
   'r.Close
   On Error GoTo cantdelete
   FileCopy "c:\talos\emptytxt.txt", Path.text + "\" + Filename.text
   On Error GoTo 0
End If

Unload Me
Exit Sub

cantdelete:
MsgBox "The file can not deleted. Another user uses the same file"
Resume Next


End Sub

Private Sub damapong_Click()

Dim work, mydb, r, text, j, p, m, prosp, m_aa, k, m_Seira, kk, m_ek, M_KOD, x_kod, ans
Command3.Enabled = False

'default μπάνιο
Command1.Enabled = False
Set work = Workspaces(0)
Set mydb = work.OpenDatabase("c:\TALOS\RECIPIES.MDB", False, False)
Set r = mydb.OpenRecordset("parametroi")
    Text1.text = r("MPANIO")
    Text2.text = r("default_baros_panioy")
     r.Close

'Data2.DatabaseName = Path.text
'Data2.RecordSource = filename.text
'Data2.Refresh


Set x = mydb.OpenRecordset("ximitech")
x.Index = "KOD"




Set text = OpenDatabase(Path.text, False, False, "Text;")
Set r = text.OpenRecordset(Filename.text)

Set j = mydb.OpenRecordset("Joblist")
j.Index = "ENTOLH"

Set p = mydb.OpenRecordset("prospau2")
p.Index = "ENTOLH"

'Set m = PROSMACB.Recordset
If r.RecordCount < 2 Then
   MsgBox "The file is empty."
   Exit Sub
End If

'--------------------------------
r.MoveFirst
Do While Not r.eof
 
     
     m_entolh = r(0)
     Label4.Caption = m_entolh
     If IsNumeric(m_entolh) Then
         m_entolh = Right("000000" + LTrim(str(Val(m_entolh))), 6) 'last
     Else
         
     End If
     
     m_entolh = Val(m_entolh)
          If m_entolh = 0 Then
      Else
          j.Seek "=", Right("000000" + LTrim(str(Val(m_entolh))), 6) 'last
     End If
     If j.NoMatch Or m_entolh = 0 Then
         j.MoveLast
         If IsNull(j("ENTOLH")) Then
             m = 1
         Else
             m = Val(j("ENTOLH")) + 1
         End If
         Label5.Caption = m
         Command2.Enabled = True
         Do While Command2.Enabled = True
           DoEvents
         Loop
         m_entolh = Right("000000" + LTrim(str(Val(m))), 6) 'last
         j.AddNew
         j("entolh") = m_entolh
         j("kod_pel") = DBCombo1.text
         j("baros_pani") = Val(Text2.text)
         j("sxesh_mpan") = Val(Text1.text)
         j("kod_erg") = 0
         j("xhmika") = Val(Text3.text)
         j.update
         prosp = 1
      Else
         
         p.Seek "=", j("entolh") 'BRISKO TON MEGALITERO ARIUMO PROSPAUIAS
         m_aa = p("AA_PROSPAU")
         Form1.Caption = "m_aa"
         Do While p.ENTOLH = m_entolh And Not p.eof
            If p("AA_PROSPAU") > m_aa Then
               m_aa = p("AA_PROSPAU")
            End If
            p.MoveNext
         Loop
         prosp = Val(m_aa)
     End If
     
     For k = 0 To 3
            Label6(k).Caption = " "
            Label7(k).Caption = " "
     Next
     
         ' r.MoveNext
          m_Seira = 0



'xhmika ths syntaghs
     'If r.EOF Then Exit Do
     For kk = 1 To r(1) 'Step 2 ' Do While Not r.EOF
         m_ek = r(kk * 2 + 1) '0.79432
         M_KOD = left(r(kk * 2), 3) 'code
         If m_ek <= 0 Or Val(M_KOD) = 0 Then
           x_kod = " "
         Else
           x_kod = find_kod(M_KOD, m_ek)
         End If
       
       If left(x_kod, 1) <> " " Then
         m_Seira = m_Seira + 1
         p.AddNew
         p("entolh") = m_entolh 'last
         p("aa_prospau") = Right("00" + LTrim(str(prosp)), 2)
         p("kod_pel") = DBCombo1.text
         p("baros_pani") = Val(Text2.text)
         p("sxesh_mpan") = Val(Text1.text)
         p("ek") = m_ek
         p("gl") = 0
         p("kod") = x_kod
         p("SEIRA") = m_Seira
         p("ayjon") = m_Seira
         p("status") = 0
         p("APOR") = 0
         p.update
         If m_Seira <= 4 And Not IsNull(x("perigrafh")) Then
            Label6(m_Seira - 1).Caption = Perigrafh
            Label7(m_Seira - 1).Caption = Format(m_ek, "#0.0000")
         End If
        End If
         
         'r.MoveNext
     Next
     
     ' salt
       
         
         m_ek = Val(Text4.text) '0.79432
         M_KOD = "SALT...." 'code
         
         If m_ek <= 0 Or Len(Trim(M_KOD)) < 2 Then
           x_kod = " "
         Else
           x_kod = find_kod(M_KOD, m_ek)
         End If
       
       
       
       
       If Val(Text4.text) > 0 And x_kod <> " " Then
          m_Seira = m_Seira + 1
          p.AddNew
         p("entolh") = m_entolh 'last
         p("aa_prospau") = Right("00" + LTrim(str(prosp)), 2)
         p("kod_pel") = DBCombo1.text
         p("baros_pani") = Val(Text2.text)
         p("sxesh_mpan") = Val(Text1.text)
         p("ek") = 0
         p("gl") = m_ek
         p("kod") = x_kod
         p("SEIRA") = m_Seira
         p("ayjon") = m_Seira
         p("status") = 0
         p("APOR") = 0
         p.update
       End If
     
     
     
     ans = MsgBox("Save this recipe ?", vbYesNo)
     If ans = vbYes Then
     Else
        mydb.Execute "delete *from prospau2 where entolh='" + m_entolh + "'"
        mydb.Execute "delete *from joblist where entolh='" + m_entolh + "'"
     End If
     
     
     For k = 0 To 3
            Label6(k).Caption = " "
            Label7(k).Caption = " "
     Next
     r.MoveNext
     If r.eof Then Exit Do
     
         ' Η ΕΠΟΜΕΝΗ ΣΥΝΤΑΓΗ
          For kk = 1 To r(1) 'Step 2 ' Do While Not r.EOF
             m_ek = r(kk * 2 + 1) '0.79432
             M_KOD = left(r(kk * 2), 3) 'code
             Label6(kk - 1).Caption = r(kk * 2)
             Label7(kk - 1).Caption = r(kk * 2 + 1)
          Next
     
     
     
     
Loop
r.Close
ans = MsgBox(" Do you want to delete the text file ? ", vbYesNo, "The import is finished.")
If ans = vbYes Then
   'r.Close
   FileCopy "c:\talos\emptytxt.txt", Path.text + "\" + Filename.text
End If

Unload Me
End Sub

Private Sub TreePoint_Click()
'   Colibri

Dim mm, mm0, work, mydb, r, text, j, p, k, m, prosp, m_aa, m_Seira, m_ek, M_KOD, x_kod, ans
Dim DUM, nFile, MyPath, MyName, AFiles(1 To 100), ELEM(1 To 30), KL, KE
Dim Narx, rr$, m_entolh, reco(1 To 200), aaReco
' βρίσκω εγώ όλα τα files που τελειώνουν σε *.txt KAI TA APOTHIKEYO SE PINAKA
' to diabazo kai meta to sbhno
TreePoint.Enabled = False

' Display the names in C:\ that represent directories.
k = 0
MyPath = "c:\talos\temp\"  ' Set the path.
MyName = Dir(MyPath, vbNormal)   ' Retrieve the first entry.
Do While MyName <> ""   ' Start the loop.
    If UCase(Right(MyName, 3)) = "TXT" Then
       k = k + 1
       If k <= 100 Then
          AFiles(k) = MyName
       End If
    End If
    MyName = Dir    ' Get next entry.
Loop



Narx = k

If Narx = 0 Then
    MsgBox "No files for recipes"
End If


'DUM = Shell("c:\talos\merge.bat  " + LTrim(AFiles(nFile)), 0) 'prosuetei mia grammh
'MilSec 3000



For nFile = 1 To Narx


Command3.Enabled = False


'DUM = Shell("c:\talos\merge.bat  " + LTrim(AFiles(nFile)), 0) 'prosuetei mia grammh
'MilSec 3000


Open MyPath + LTrim(AFiles(nFile)) For Input As #1
' Close before reopening in another mode.
aaReco = 0
Do While Not eof(1)
   Input #1, rr$
   aaReco = aaReco + 1
   reco(aaReco) = rr$
Loop
Close #1



'default μπάνιο
Command1.Enabled = False

Set work = Workspaces(0)
Set mydb = work.OpenDatabase("c:\TALOS\RECIPIES.MDB", False, False)
Set r = mydb.OpenRecordset("parametroi")
    Text1.text = r("MPANIO")
    Text2.text = r("default_baros_panioy")
     r.Close


Set x = mydb.OpenRecordset("ximitech")
x.Index = "KOD"





'Set text = OpenDatabase(Path.text, True, False, "Text;")
'Set r = text.OpenRecordset(filename.text)
'If r.RecordCount < 5 Then
'   mydb.Close
'   text.Close
'   GoTo 19088
'End If





Set j = mydb.OpenRecordset("Joblist")
j.Index = "ENTOLH"

Set p = mydb.OpenRecordset("prospau2")
p.Index = "ENTOLH"

'Set m = PROSMACB.Recordset
If aaReco < 2 Then
   MsgBox "The file is empty."
   Command3.Enabled = True

   Exit Sub
End If

'X;H;P-ID;[P-ZID];[P-DDESIG];weight;unit of weight<LF>
'--------------------------------------------------------------------------------
'X                       corresponds:    Extra Line
'H                       corresponds:    Header Line
'P-ID                    corresponds:    Production ID (Same as file name)
'[P-ZID]         corresponds:    Production ZID (can be empty then ;; is into t
'[P-DESIG]               corresponds:    Production Designation (can be empty ;
'120;                    corresponds:    Sample weight
'g                       corresponds:    Unit of sample
'<LF>                    corresponds:    End of Line


' ;=CHR(9)  TAB
'X;H;00123;;Red A034;50.00;g<LF>
'P;PR0321;D<LF>
'A;H;1;2;Dyeliquor;500.00;ml<LF>
'A;F;PROCION 003;001003;Yellow H-4G;0.2100;%;0.105;g;1<LF>
'A;F;PROCION 007;001007;Red H-E7B;1.32;%;0.660;g;1<LF>
'A;C;SOULFATE;100020;Soulfate;5.00;g/l;2.500;g;0<LF>
'A;B;------------------------------------------------- <LF>
'A;H;2;1;Add in Dyeliquor;500;ml<LF>
'A;B;After dispensing set pH to 5.5 - 6.0 with formic acid<LF>
'A;B;Addition of oxidol (Ludigol)<LF>
'A;C;LUDIGOL;200045;Ludigol;10.00;ml/l;0.550;g;1<LF>
'A;B;-------------------------------------------------- <LF>
'A;H;3;1;Manual Add;P-VOLUME;ml<LF>
'A;B;Set alkali in 2 steps<LF>
'A;C;SODAASH;100035;Soda Ash;15.00;g/l;7.500;g;0<LF>



r = 1



Do While r <= aaReco ' 2 --------------------------------    until find 1 recipe
 
 If Mid$(reco(r), 1, 1) = "X" Then 'new recipy
     
    
    '  DIABAZO SE PINAKA OLA TA STOIXEIA THS GRAMHS
         DUM = FETES_DELIM(reco(r), ELEM)
    
     
     
     
     
     m_entolh = ELEM(3)
     
     Label4.Caption = LTrim(str(nFile)) + "  " + m_entolh
     
     Text5.text = ELEM(5)
     
     
     mm = r
'     On Error Resume Next
     mm0 = 0
     
     r = r + 1
     k = 0
     Do While r <= aaReco
            
            '  DIABAZO SE PINAKA OLA TA STOIXEIA THS GRAMHS
            DUM = FETES_DELIM(reco(r), ELEM)
            
           ' If Not r.EOF Then
              ' If ELEM(1) = "A" And (ELEM(2) = "F" Or ELEM(2) = "C") Then
               If ELEM(1) = "A" And (ELEM(2) = "F") Then   'mono xromata
                    Label7(k).Caption = ELEM(6) '0.79432
                    Label6(k).Caption = ELEM(4) + ELEM(5) 'code
                    mm0 = mm0 + Val(Label7(k).Caption)
                    k = k + 1
               End If
            'End If
            r = r + 1 '.MoveNext
            
     Loop
     
     
     Label7(4).Caption = mm0
     On Error GoTo 0
     
     
     
     r = mm
     
     
     
     If IsNumeric(m_entolh) Then
         m_entolh = Right("000000" + LTrim(str(Val(m_entolh))), 6) 'last
     Else
         
     End If
     
     m_entolh = Val(m_entolh)
     
     If m_entolh = 0 Then
     Else
          j.Seek "=", Right("000000" + LTrim(str(Val(m_entolh))), 6) 'last
     End If
     
     If j.NoMatch Or m_entolh = 0 Then
         j.MoveLast
         If IsNull(j("ENTOLH")) Then
             m = 1
         Else
             m = Val(j("ENTOLH")) + 1
         End If
         Label5.Caption = m
         Command2.Enabled = True
         Do While Command2.Enabled = True
           DoEvents
         Loop
         m_entolh = Right("000000" + LTrim(str(Val(m))), 6) 'last
         
       On Error Resume Next
         j.AddNew
         j("entolh") = m_entolh
         ' Form1.Caption = m_entolh
         j("kod_pel") = DBCombo1.text
         'Form1.Caption = DBCombo1.text
         j("baros_pani") = Val(Text2.text)
         ' Form1.Caption = Val(Text2.text)
         j("sxesh_mpan") = Val(Text1.text)
         'Form1.Caption = Val(Text2.text)
         j("kod_erg") = 0
         'Form1.Caption = "kod_erg"
         j("xhmika") = Val(Text3.text)
          
         j("apoxrvsh") = left(Text5.text, 20)
         j("eidos_baf") = left(DBCombo2.text, 15)
         j.update
         On Error GoTo 0
         
         
         
         'Form1.Caption = "J/update"
         prosp = 1
      Else
         
         p.Seek "=", j("entolh") 'BRISKO TON MEGALITERO ARIUMO PROSPAUIAS
         m_aa = p("AA_PROSPAU")
         'Form1.Caption = "m_aa"
         Do While p.ENTOLH = m_entolh And Not p.eof
            If p("AA_PROSPAU") > m_aa Then
               m_aa = p("AA_PROSPAU")
               'Form1.Caption = "r/m_aa2"
            End If
            p.MoveNext
         Loop
         prosp = Val(m_aa)
         ' Form1.Caption = "r/m_aa3"
     End If
     
     
     
     For k = 0 To 3
            Label6(k).Caption = " "
            Label7(k).Caption = " "
     Next
     
     
          r = r + 1
          m_Seira = 0
     
     
     
     Do While r <= aaReco ' Not r.EOF   ' loop sta systatika ths syntaghs
         
          If IsNull(reco(r)) Then Exit Do
         'If Val(Mid$(r(0), 27, 8)) = 0 And Val(Mid$(r(0), 1, 4)) = 0 Then Exit Do 'new recipy
                    '  DIABAZO SE PINAKA OLA TA STOIXEIA THS GRAMHS
            DUM = FETES_DELIM(reco(r), ELEM)
 
          If Mid$(reco(r), 1, 1) = "X" Then Exit Do 'new recipy
'                    If ELEM(1) = "A" And (ELEM(2) = "F" Or ELEM(2) = "C") Then
          
             m_ek = Val(ELEM(6)) '0.79432
             
             M_KOD = ELEM(4) 'code
             ' If m_ek <= 0 Or Len(Trim(M_KOD)) < 2  Then
             If m_ek <= 0 Or Len(Trim(M_KOD)) < 2 Or ELEM(2) = "C" Then 'bgazo ta ximika
                   x_kod = " "
             Else
                   x_kod = find_kod(M_KOD, m_ek)
             End If
       
             If left(x_kod, 1) <> " " And Val(x_kod) > 0 Then
                m_Seira = m_Seira + 1
                
                p.AddNew
                p("entolh") = m_entolh 'last
                p("aa_prospau") = Right("00" + LTrim(str(prosp)), 2)
                p("kod_pel") = DBCombo1.text
                p("baros_pani") = Val(Text2.text)
                p("sxesh_mpan") = Val(Text1.text)
                
                If InStr(ELEM(7), "%") Then
                    p("ek") = m_ek
                    p("gl") = 0
                Else
                    p("ek") = 0
                    p("gl") = m_ek
                End If
                
                
                p("kod") = x_kod
                p("SEIRA") = m_Seira
                p("alati") = Val(Text4.text)
                p("ayjon") = m_Seira
                p("apoxrvsh") = left(Text5.text, 20)
                p("status") = 0
                p("APOR") = 0
                p.update
                'Form1.Caption = "p/update"
                If m_Seira <= 4 And Not IsNull(x("perigrafh")) Then
                      Label6(m_Seira - 1).Caption = Perigrafh
                      Label7(m_Seira - 1).Caption = Format(m_ek, "#0.0000")
                End If
             End If
         
             r = r + 1 '.MoveNext
     Loop
     
     
     
     ans = MsgBox("Save this recipe ?", vbYesNo)
     If ans = vbYes Then
        'work.CommitTrans
     Else
        'work.Rollback
        mydb.Execute "delete *from prospau2 where entolh='" + m_entolh + "'"
        mydb.Execute "delete *from joblist where entolh='" + m_entolh + "'"
     End If
   
     Kill MyPath + LTrim(AFiles(nFile))
     
     For k = 0 To 3
            Label6(k).Caption = " "
            Label7(k).Caption = " "
     Next
     
     Exit Do
     
  Else
     r = r + 1
  End If

Loop  '  2  MEGALO LOOP ME 1 SYNTAGH ANA KYKLO
'r.Close




ans = vbNo ' MsgBox(" Do you want to delete the text file ? ", vbYesNo, "The import is finished.")
If ans = vbYes Then
   'r.Close
   On Error GoTo cantdelete
       Kill "e:\MyDocu~1\" + LTrim(str(nFile)) + ".prn"
   On Error GoTo 0
End If

19088
Next nFile



Unload Me
Exit Sub

cantdelete:
MsgBox "The file can not deleted. Another user uses the same file"
Resume Next


End Sub

'
Private Sub TreePoint_old_Click()  ' 13/5/2002
'   Colibri

Dim mm, mm0, work, mydb, r, text, j, p, k, m, prosp, m_aa, m_Seira, m_ek, M_KOD, x_kod, ans
Dim DUM, nFile, MyPath, MyName, AFiles(1 To 100), ELEM(1 To 30), KL, KE
Dim Narx
' βρίσκω εγώ όλα τα files που τελειώνουν σε *.txt KAI TA APOTHIKEYO SE PINAKA
' to diabazo kai meta to sbhno


' Display the names in C:\ that represent directories.
k = 0
MyPath = "c:\talos\temp\"  ' Set the path.
MyName = Dir(MyPath, vbNormal)   ' Retrieve the first entry.
Do While MyName <> ""   ' Start the loop.
    If UCase(Right(MyName, 3)) = "TXT" Then
       k = k + 1
       If k <= 100 Then
          AFiles(k) = MyName
       End If
    End If
    MyName = Dir    ' Get next entry.
Loop



Narx = k

If Narx = 0 Then
    MsgBox "No files for recipes"
End If

For nFile = 1 To Narx


Command3.Enabled = False


DUM = Shell("c:\talos\merge.bat  " + LTrim(AFiles(nFile)), 0) 'prosuetei mia grammh
MilSec 3000


'default μπάνιο
Command1.Enabled = False

Set work = Workspaces(0)
Set mydb = work.OpenDatabase("c:\TALOS\RECIPIES.MDB", False, False)
Set r = mydb.OpenRecordset("parametroi")
    Text1.text = r("MPANIO")
    Text2.text = r("default_baros_panioy")
     r.Close


Set x = mydb.OpenRecordset("ximitech")
x.Index = "KOD"





Set text = OpenDatabase(Path.text, True, False, "Text;")
Set r = text.OpenRecordset(Filename.text)
If r.RecordCount < 5 Then
   mydb.Close
   text.Close
   GoTo 19088
End If
'debug'





Set j = mydb.OpenRecordset("Joblist")
j.Index = "ENTOLH"

Set p = mydb.OpenRecordset("prospau2")
p.Index = "ENTOLH"

'Set m = PROSMACB.Recordset
If r.RecordCount < 2 Then
   MsgBox "The file is empty."
   Command3.Enabled = True

   Exit Sub
End If

'X;H;P-ID;[P-ZID];[P-DDESIG];weight;unit of weight<LF>
'--------------------------------------------------------------------------------
'X                       corresponds:    Extra Line
'H                       corresponds:    Header Line
'P-ID                    corresponds:    Production ID (Same as file name)
'[P-ZID]         corresponds:    Production ZID (can be empty then ;; is into t
'[P-DESIG]               corresponds:    Production Designation (can be empty ;
'120;                    corresponds:    Sample weight
'g                       corresponds:    Unit of sample
'<LF>                    corresponds:    End of Line


' ;=CHR(9)  TAB
'X;H;00123;;Red A034;50.00;g<LF>
'P;PR0321;D<LF>
'A;H;1;2;Dyeliquor;500.00;ml<LF>
'A;F;PROCION 003;001003;Yellow H-4G;0.2100;%;0.105;g;1<LF>
'A;F;PROCION 007;001007;Red H-E7B;1.32;%;0.660;g;1<LF>
'A;C;SOULFATE;100020;Soulfate;5.00;g/l;2.500;g;0<LF>
'A;B;------------------------------------------------- <LF>
'A;H;2;1;Add in Dyeliquor;500;ml<LF>
'A;B;After dispensing set pH to 5.5 - 6.0 with formic acid<LF>
'A;B;Addition of oxidol (Ludigol)<LF>
'A;C;LUDIGOL;200045;Ludigol;10.00;ml/l;0.550;g;1<LF>
'A;B;-------------------------------------------------- <LF>
'A;H;3;1;Manual Add;P-VOLUME;ml<LF>
'A;B;Set alkali in 2 steps<LF>
'A;C;SODAASH;100035;Soda Ash;15.00;g/l;7.500;g;0<LF>



Do While Not r.eof   ' 2 --------------------------------    until find 1 recipe
 If Mid$(r(0), 1, 1) = "X" Then 'new recipy
     
    
    '  DIABAZO SE PINAKA OLA TA STOIXEIA THS GRAMHS
         DUM = FETES_DELIM(r(0), ELEM)
    
     
     
     
     
     m_entolh = ELEM(3)
     
     Label4.Caption = LTrim(str(nFile)) + "  " + m_entolh
     
     Text5.text = ELEM(5)
     
     
     mm = r.Bookmark
     On Error Resume Next
     mm0 = 0
     
     r.MoveNext
     k = 0
     Do While Not r.eof
            
            '  DIABAZO SE PINAKA OLA TA STOIXEIA THS GRAMHS
            DUM = FETES_DELIM(r(0), ELEM)
            
           ' If Not r.EOF Then
              ' If ELEM(1) = "A" And (ELEM(2) = "F" Or ELEM(2) = "C") Then
               If ELEM(1) = "A" And (ELEM(2) = "F") Then   'mono xromata
                    Label7(k).Caption = ELEM(6) '0.79432
                    Label6(k).Caption = ELEM(4) + ELEM(5) 'code
                    mm0 = mm0 + Val(Label7(k).Caption)
                    k = k + 1
               End If
            'End If
            r.MoveNext
            
     Loop
     
     
     Label7(4).Caption = mm0
     On Error GoTo 0
     
     
     
     r.Bookmark = mm
     
     
     
     If IsNumeric(m_entolh) Then
         m_entolh = Right("000000" + LTrim(str(Val(m_entolh))), 6) 'last
     Else
         
     End If
     
     m_entolh = Val(m_entolh)
     
     If m_entolh = 0 Then
     Else
          j.Seek "=", Right("000000" + LTrim(str(Val(m_entolh))), 6) 'last
     End If
     
     If j.NoMatch Or m_entolh = 0 Then
         j.MoveLast
         If IsNull(j("ENTOLH")) Then
             m = 1
         Else
             m = Val(j("ENTOLH")) + 1
         End If
         Label5.Caption = m
         Command2.Enabled = True
         Do While Command2.Enabled = True
           DoEvents
         Loop
         m_entolh = Right("000000" + LTrim(str(Val(m))), 6) 'last
         
       On Error Resume Next
         j.AddNew
         j("entolh") = m_entolh
         Form1.Caption = m_entolh
         j("kod_pel") = DBCombo1.text
         Form1.Caption = DBCombo1.text
         j("baros_pani") = Val(Text2.text)
         Form1.Caption = Val(Text2.text)
         j("sxesh_mpan") = Val(Text1.text)
         Form1.Caption = Val(Text2.text)
         j("kod_erg") = 0
         Form1.Caption = "kod_erg"
         j("xhmika") = Val(Text3.text)
          
         j("apoxrvsh") = left(Text5.text, 20)
         j("eidos_baf") = left(DBCombo2.text, 15)
         j.update
         On Error GoTo 0
         
         
         
         Form1.Caption = "J/update"
         prosp = 1
      Else
         
         p.Seek "=", j("entolh") 'BRISKO TON MEGALITERO ARIUMO PROSPAUIAS
         m_aa = p("AA_PROSPAU")
         Form1.Caption = "m_aa"
         Do While p.ENTOLH = m_entolh And Not p.eof
            If p("AA_PROSPAU") > m_aa Then
               m_aa = p("AA_PROSPAU")
               Form1.Caption = "r/m_aa2"
            End If
            p.MoveNext
         Loop
         prosp = Val(m_aa)
         Form1.Caption = "r/m_aa3"
     End If
     
     
     
     For k = 0 To 3
            Label6(k).Caption = " "
            Label7(k).Caption = " "
     Next
     
     
          r.MoveNext
          m_Seira = 0
     
     
     
     Do While Not r.eof   ' loop sta systatika ths syntaghs
         
          If IsNull(r(0)) Then Exit Do
         'If Val(Mid$(r(0), 27, 8)) = 0 And Val(Mid$(r(0), 1, 4)) = 0 Then Exit Do 'new recipy
                    '  DIABAZO SE PINAKA OLA TA STOIXEIA THS GRAMHS
            DUM = FETES_DELIM(r(0), ELEM)
 
          If Mid$(r(0), 1, 1) = "X" Then Exit Do 'new recipy
'                    If ELEM(1) = "A" And (ELEM(2) = "F" Or ELEM(2) = "C") Then
          
             m_ek = Val(ELEM(6)) '0.79432
             
             M_KOD = ELEM(4) 'code
             ' If m_ek <= 0 Or Len(Trim(M_KOD)) < 2  Then
             If m_ek <= 0 Or Len(Trim(M_KOD)) < 2 Or ELEM(2) = "C" Then 'bgazo ta ximika
                   x_kod = " "
             Else
                   x_kod = find_kod(M_KOD, m_ek)
             End If
       
             If left(x_kod, 1) <> " " And Val(x_kod) > 0 Then
                m_Seira = m_Seira + 1
                
                p.AddNew
                p("entolh") = m_entolh 'last
                p("aa_prospau") = Right("00" + LTrim(str(prosp)), 2)
                p("kod_pel") = DBCombo1.text
                p("baros_pani") = Val(Text2.text)
                p("sxesh_mpan") = Val(Text1.text)
                
                If InStr(ELEM(7), "%") Then
                    p("ek") = m_ek
                    p("gl") = 0
                Else
                    p("ek") = 0
                    p("gl") = m_ek
                End If
                
                
                p("kod") = x_kod
                p("SEIRA") = m_Seira
                p("alati") = Val(Text4.text)
                p("ayjon") = m_Seira
                p("apoxrvsh") = left(Text5.text, 20)
                p("status") = 0
                p("APOR") = 0
                p.update
                Form1.Caption = "p/update"
                If m_Seira <= 4 And Not IsNull(x("perigrafh")) Then
                      Label6(m_Seira - 1).Caption = Perigrafh
                      Label7(m_Seira - 1).Caption = Format(m_ek, "#0.0000")
                End If
             End If
         
             r.MoveNext
     Loop
     
     ' salt
       
         
      '   m_ek = Val(Text4.text) '0.79432
      '   m_kod = "SALT...." 'code
         
      '   If m_ek <= 0 Or Len(Trim(m_kod)) < 2 Then
      '     x_kod = " "
      '   Else
      '     x_kod = find_kod(m_kod, m_ek)
      '   End If
      '
       
       
       
      ' If Val(Text4.text) > 0 And x_kod <> " " Then
      '    m_Seira = m_Seira + 1
      '    p.AddNew
      '   p("entolh") = m_entolh 'last
      '   p("aa_prospau") = Right("00" + LTrim(Str(prosp)), 2)
      '   p("kod_pel") = DBCombo1.text
      '   p("baros_pani") = Val(Text2.text)
      '   p("sxesh_mpan") = Val(Text1.text)
      '   p("ek") = 0
      '   p("gl") = m_ek
      '   p("kod") = x_kod
      '   p("SEIRA") = m_Seira
      '   p("ayjon") = m_Seira
      '   p("status") = 0
      '   p("APOR") = 0
      '   p.update
      ' End If
     
     
     
     ans = MsgBox("Save this recipe ?", vbYesNo)
     If ans = vbYes Then
        'work.CommitTrans
     Else
        'work.Rollback
        mydb.Execute "delete *from prospau2 where entolh='" + m_entolh + "'"
        mydb.Execute "delete *from joblist where entolh='" + m_entolh + "'"
     End If
   
     
     
     For k = 0 To 3
            Label6(k).Caption = " "
            Label7(k).Caption = " "
     Next
     
     Exit Do
     
     'If r.EOF Then Exit Do
  Else
     r.MoveNext
  End If

Loop  '  2  MEGALO LOOP ME 1 SYNTAGH ANA KYKLO
r.Close




ans = vbNo ' MsgBox(" Do you want to delete the text file ? ", vbYesNo, "The import is finished.")
If ans = vbYes Then
   'r.Close
   On Error GoTo cantdelete
       Kill "e:\MyDocu~1\" + LTrim(str(nFile)) + ".prn"
   On Error GoTo 0
End If

19088
Next nFile



Unload Me
Exit Sub

cantdelete:
MsgBox "The file can not deleted. Another user uses the same file"
Resume Next


End Sub



Function to437(s_)
'********************************************************************
Dim k, a, b
Dim s, s437, s869
s = ""
s437 = "€‚ƒ„…†‡‰‹‘’“”•–—™› ΅Ά£¤¥¦§¨©«¬­®―ΰª" + Chr(34)
s869 = "ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩαβγδεζηθικλμνξοπρστυφχψως" + Chr(32)
For k = 1 To Len(s_)
    a = Mid$(s_, k, 1)
    b = InStr(s437, a)
    If b > 0 Then '&& ηταν ελληνικό 437
       s = s + Mid(s869, b, 1)
    Else
       s = s + a
    End If
Next
to437 = s
End Function


'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'

'
'
Function FETES_DELIM(LINE, ELEM)
Dim KL, KE
 '  DIABAZO SE PINAKA OLA TA STOIXEIA THS GRAMHS
     For KE = 1 To 30
       ELEM(KE) = ""
     Next
     
     KL = 0   ' metraei xaraktires
     KE = 0   ' metritis toy pinaka  ELEMENT
     Do
        KE = KE + 1  ' metritis toy pinaka  ELEMENT
        KL = KL + 1  ' metraei xaraktires
        
        Do While Mid$(LINE, KL, 1) <> Chr(9) ' tab
            ELEM(KE) = ELEM(KE) + Mid$(LINE, KL, 1)
            KL = KL + 1  ' metraei xaraktires
            If KL > Len(LINE) Then Exit Do
        Loop
        
        
            If KL > Len(LINE) Then Exit Do
     
     Loop Until Mid$(LINE, KL, 1) = Chr(10) 'OLO TO MHKOS THS GRAMMHS

     FETES_DELIM = 0

End Function
'
'
'
'
'
Private Sub farbetex_Click()
Dim work, mydb, r(0 To 2), text, j, p, m, prosp, m_aa, k, m_Seira, kk, m_ek, M_KOD, x_kod, ans
Dim dis, mcounter, m_prosp, DUM, yparxei, reco(1 To 200), aaReco, aaReco2
Dim r0, rr$
Command3.Enabled = False
On Error Resume Next
'default μπάνιο

'DUM = Shell("c:\talos\merge.bat")
'MilSec 5000

disable_imports





Command1.Enabled = False
Set work = Workspaces(0)
Set mydb = work.OpenDatabase("c:\TALOS\RECIPIES.MDB", False, False)
Set r0 = mydb.OpenRecordset("parametroi")
    Text1.text = r0("MPANIO")
    Text2.text = r0("default_baros_panioy")
     r0.Close

'Data2.DatabaseName = Path.text
'Data2.RecordSource = filename.text
'Data2.Refresh


Set x = mydb.OpenRecordset("ximitech")
x.Index = "KOD"




' 27-5 Set text = OpenDatabase(Path.text, False, False, "Text;")
'27-5 Set r = text.OpenRecordset(filename.text)


'-------- 27-5
Open LTrim(Path.text + Filename.text) For Input As #1
' Close before reopening in another mode.
aaReco = 0
Do While Not eof(1)
   Input #1, rr$
   aaReco = aaReco + 1
   reco(aaReco) = rr$
Loop
Close #1
'-------- 27-5





Set j = mydb.OpenRecordset("Joblist")
j.Index = "ENTOLH"

Set p = mydb.OpenRecordset("prospau2")
p.Index = "ENTOLH"

'Set m = PROSMACB.Recordset
If aaReco < 1 Then
   MsgBox "The file is empty."
   Exit Sub
End If
aaReco2 = 0
'--------------------------------
'r.MoveFirst
Do While aaReco2 < aaReco
 
     aaReco2 = aaReco2 + 1
     r(0) = to437(reco(aaReco2))
     
     m_entolh = left(r(0), 6)
     Label4.Caption = m_entolh
     Command2.Enabled = True
     
    DBCombo1.text = Mid$(r(0), 36, 15)
     
     m_prosp = Mid$(r(0), 8, 2)
     
     Text1.text = Val(Mid$(r(0), 85, 4))
     Text5.text = Mid$(r(0), 11, 12)
     
     If IsNumeric(m_entolh) Then
         m_entolh = Right("000000" + LTrim(str(Val(m_entolh))), 6) 'last
     Else
         
     End If
     
     
     
     m_entolh = Val(m_entolh)
      
     If m_entolh = 0 Then
     Else
          j.Seek "=", Right("000000" + LTrim(str(Val(m_entolh))), 6) 'last
     End If
     
     m = m_entolh
     
     
'         j.MoveLast
 '        If IsNull(j("ENTOLH")) Then
  '           m = 1
   '      Else
    '         m = Val(j("ENTOLH")) + 1
     '    End If
      '   Label5.Caption = m
       '  Command2.Enabled = True
         
'     For k = 0 To 11
 '           If Not r.eof Then
  '            If Len(LTrim(Trim(Mid$(r(0), 51, 6)))) > 2 And Val(Mid$(r(0), 58, 10)) > 0 Then
   '              Label7(mCounter).Caption = Val(Mid$(r(0), 58, 10)) '0.79432
    '             Label6(mCounter).Caption = Mid$(r(0), 51, 6) 'code
     '            mm0 = mm0 + Val(Label7(mCounter).Caption)
      ''           mCounter = mCounter + 1
        '
         '     End If
          '  End If
     '   End If
         
              dis = 0
              mcounter = 0
          For k = 0 To 3
             
              Label7(k).Caption = Val(Mid$(r(0), dis + 113, 7)) '0.79432
              Label6(k).Caption = Mid$(r(0), dis + 90, 6) + "-" + Mid$(r(0), dis + 97, 15) 'code
              ' mcounter = mcounter + 1
             
             
             ' REPL ENTOLH     WITH LEFT(TRANS->S,6)
             ' REPL aa_prospau WITH strzero(min(99,val(subs(TRANS->S,8,2))),2)
             ' REPL APOR       WITH 0
             ' REPL SXESH_MPAN WITH MIN(99,val( SUBS(TRANS->S,85,4) ))
             ' REPL STATUS WITH 0
             
             '   SELE XIMITECH
             '   SEEK subs(TRANS->S,dis+90,6)
             '   IF .NOT. EOF() .AND. LEN( TRIM ( subs(TRANS->S,dis+90,6)  ) ) > 0
             '       periek1 = val(subs(TRANS->S,dis+113,7))
             '       m_reco = recn()
             '       m_reco_morfh = val(XIMITECH->morfh)
             '       skip

               '     * §¨ε§«ΰ© ¤ β®ΰ ›η¦ §¨ ΅« ΅ζ««ª
              '      if ximitech->kataskeyas = subs(TRANS->S,dis+90,6)


               '     repl kod        with XIMITECH->KOD
                '    repl ek      with val(subs(TRANS->S,dis+113,7))
                '    repl gl      with val(subs(TRANS->S,dis+121,7))


                dis = dis + 39

             Next

         
         
         Do While Command2.Enabled = True
              DoEvents
         Loop
         
         m_entolh = Right("000000" + LTrim(str(Val(m))), 6) 'last
          
         j.Seek "=", Right("000000" + LTrim(str(Val(m_entolh))), 6) 'last
     If j.NoMatch Then
         j.AddNew
         j("entolh") = m_entolh
         j("kod_pel") = DBCombo1.text
         j("baros_pani") = Val(Text2.text)
         j("sxesh_mpan") = Val(Text1.text)
         j("apoxrvsh") = left(Text5.text, 20)
         j("eidos_baf") = left(DBCombo2.text, 15)
         
         j("kod_erg") = 0
         j("xhmika") = Val(Text3.text)
         j.update
      End If
      
         prosp = Val(m_prosp)
'      Else
         
 '        p.Seek "=", j("entolh") 'BRISKO TON MEGALITERO ARIUMO PROSPAUIAS
  '       m_aa = p("AA_PROSPAU")
   '      Form1.Caption = "m_aa"
    '     Do While p.ENTOLH = m_entolh And Not p.eof
     '       If p("AA_PROSPAU") > m_aa Then
      '         m_aa = p("AA_PROSPAU")
       '     End If
        '    p.MoveNext
'         Loop
 '        prosp = Val(m_aa)
  '   End If
     
     For k = 0 To 3
            Label6(k).Caption = " "
            Label7(k).Caption = " "
     Next
     
         ' r.MoveNext
          m_Seira = 0
'xhmika ths syntaghs
     'If r.EOF Then Exit Do
     
     
   dis = 0
        p.Index = "entolh"
        p.Seek "=", m_entolh, Right("00" + LTrim(str(prosp)), 2)
     
            
   If p.NoMatch Then
      yparxei = False
     
     For kk = 0 To 3 'Step 2 ' Do While Not r.EOF
        Label7(kk).Caption = Val(Mid$(r(0), dis + 113, 7)) '0.79432
        Label6(kk).Caption = Mid$(r(0), dis + 90, 6) 'code
         m_ek = Val(Mid$(r(0), dis + 113, 7)) '0.79432
         M_KOD = Mid$(r(0), dis + 90, 6) 'code
         
         If m_ek <= 0 Or Val(M_KOD) = 0 Then
           x_kod = " "
         Else
           x_kod = find_kod(M_KOD, m_ek)
         End If
       
       If left(x_kod, 1) <> " " Then
         m_Seira = m_Seira + 1
         
         
   
           p.AddNew
           p("entolh") = m_entolh 'last
           p("aa_prospau") = Right("00" + LTrim(str(prosp)), 2)
           p("kod_pel") = DBCombo1.text
           p("baros_pani") = Val(Text2.text)
           p("sxesh_mpan") = Val(Text1.text)
           p("ek") = m_ek
           p("gl") = 0
           p("kod") = x_kod
           p("SEIRA") = m_Seira
           p("ayjon") = m_Seira
           p("status") = 0
           p("apoxrvsh") = left(Text5.text, 20)
           
           p("APOR") = 0
'           j("apoxrvsh") = left(Text5.text, 20)
 '         j("eidos_baf") = left(DBCombo2.text, 15)
           
           
           
           p.update
         
         If m_Seira <= 4 And Not IsNull(x("perigrafh")) Then
            Label6(m_Seira - 1).Caption = Perigrafh
            Label7(m_Seira - 1).Caption = Format(m_ek, "#0.0000")
         End If
        End If
        dis = dis + 39
         
         
         
         'r.MoveNext
     Next
     
            yparxei = False
 Else
            yparxei = True
 End If
     
     
     
     
     
     
     
     
     
     
     ' salt
       
         
         m_ek = Val(Text4.text) '0.79432
         M_KOD = "SALT...." 'code
         
         If m_ek <= 0 Or Len(Trim(M_KOD)) < 2 Then
           x_kod = " "
         Else
           x_kod = find_kod(M_KOD, m_ek)
         End If
       
       
       
       
       If Val(Text4.text) > 0 And x_kod <> " " And yparxei = False Then
          m_Seira = m_Seira + 1
          p.AddNew
         p("entolh") = m_entolh 'last
         p("aa_prospau") = Right("00" + LTrim(str(prosp)), 2)
         p("kod_pel") = DBCombo1.text
         p("baros_pani") = Val(Text2.text)
         p("sxesh_mpan") = Val(Text1.text)
         p("ek") = 0
         p("gl") = m_ek
         p("kod") = x_kod
         p("SEIRA") = m_Seira
         p("ayjon") = m_Seira
         p("status") = 0
         p("APOR") = 0
         p.update
       End If
     
     
     If yparxei = False Then
        ans = MsgBox("Nα αποθηκευτεί η συνταγή " + m_entolh + "-" + Right("00" + LTrim(str(prosp)), 2) + " ; ", vbYesNo)
     Else
        MsgBox "Yπάρχει ήδη η συνταγή " + m_entolh + "-" + Right("00" + LTrim(str(prosp)), 2)
     End If
     
     
    If ans = vbYes Then
    Else
        mydb.Execute "delete *from prospau2 where entolh='" + m_entolh + "'"
        mydb.Execute "delete *from joblist where entolh='" + m_entolh + "'"
    End If
     
     
     For k = 0 To 3
            Label6(k).Caption = " "
            Label7(k).Caption = " "
     Next
    '27-5 r.MoveNext
     '27-5      If r.eof Then Exit Do
     
         ' Η ΕΠΟΜΕΝΗ ΣΥΝΤΑΓΗ
       '   For kk = 1 To r(1) 'Step 2 ' Do While Not r.EOF
        '     m_ek = r(kk * 2 + 1) '0.79432
         '    M_KOD = left(r(kk * 2), 3) 'code
          '   Label6(kk - 1).Caption = r(kk * 2)
           '  Label7(kk - 1).Caption = r(kk * 2 + 1)
          'Next
     
     
      
     
Loop
' r.Close
ans = MsgBox(" Do you want to delete the text file ? ", vbYesNo, "The import is finished.")
If ans = vbYes Then
   'r.Close
   Kill LTrim(Path.text + Filename.text)
   Kill "f:\talos\syntages.txt"
'   FileCopy "c:\talos\emptytxt.txt", Path.text + "\" + filename.text
End If

Unload Me
End Sub


Function find_kod(M_KOD, m_ek)
    Dim x1, x2, m1, m2, p1, p2, mo, Mega_sel
x1 = "[["
x2 = "[["
Perigrafh = "---"
   find_kod = " "
x.MoveFirst
Do While Not x.eof
   
   If Trim(x("kataskeyas")) = M_KOD And x("skonh") = 0 Then
     
     If IsNull(x("real_cons")) Then    '10
         x.EDIT
            x("real_cons") = val2(x("morfh"))
         x.update
     End If
     
     If x("real_cons") > 0 Then  '11
       If x1 = "[[" Then  '12
           x1 = x("kod")
           m1 = x("real_cons")   ' x("morfh")
           p1 = x("perigrafh")
       Else
           x2 = x("kod")
           m2 = x("real_cons")  ' x("morfh")
           p2 = x("perigrafh")
           Exit Do
       End If '12
      End If '11
   End If '10
   x.MoveNext
Loop
If x1 <> "[[" Then  'an brike mono ena
   If x2 <> "[[" Then
          
         If m1 > m2 Then
               If m_ek > m1 / 20 Then
                  Mega_sel = 1
               Else
                  Mega_sel = 0
               End If
          
         Else
               If m_ek > m2 / 50 Then
                  Mega_sel = 1
               Else
                  Mega_sel = 0
               End If
         End If
      
         If Mega_sel = 1 Then
            find_kod = IIf(m1 > m2, x1, x2)
            Perigrafh = IIf(m1 > m2, p1, p2)
         Else
            find_kod = IIf(m2 < m1, x2, x1)
            Perigrafh = IIf(m2 < m1, p2, p1)
         End If
   Else
         find_kod = x1
         Perigrafh = p1
   End If
   x.Seek "=", find_kod
Else
    MsgBox M_KOD + " Not found."
End If


End Function


Private Sub Command2_Click()
  Command2.Enabled = False
End Sub

Private Sub Command3_Click()

   Unload Me
End Sub


Function val2(n)
Dim s
       s = IIf(InStr(n, "%") = 0, 6, InStr(n, "%"))
      If IsNull(s) Then
         val2 = 0
    Else
       val2 = Val(left(n, s - 1))
    End If
End Function



Private Sub old2Polydye_Click()
'----------   POLYDYE  ---------------
Dim mm, mm0, work, mydb, r, text, j, p, k, m, prosp, m_aa, m_Seira, m_ek, M_KOD, x_kod, ans
Dim DUM, nFile, DIRECT, mcc, m1%


For nFile = 1 To 30

DIRECT = "C:\TALOS\POLYDYE\"  '"e:\MyDocu~1\"
Command3.Enabled = False
'If Len(Dir("e:\My Documents\" + LTrim(Str(nFile)) + ".prn")) < 4 Then GoTo 9088
If Len(Dir(DIRECT + LTrim(str(nFile)) + ".prn")) < 4 Then GoTo 9088



DUM = Shell("c:\talos\merge.bat  " + LTrim(str(nFile)), 0)
MilSec 2000
'default μπάνιο
Command1.Enabled = False
Set work = Workspaces(0)
Set mydb = work.OpenDatabase("c:\TALOS\RECIPIES.MDB", False, False)
Set r = mydb.OpenRecordset("parametroi")
    Text1.text = r("MPANIO")
    Text2.text = r("default_baros_panioy")
     r.Close


Set x = mydb.OpenRecordset("ximitech")
x.Index = "KOD"

Set text = OpenDatabase(Path.text, True, False, "Text;")
Set r = text.OpenRecordset(Filename.text)
If r.RecordCount < 5 Then
   mydb.Close
   text.Close
   GoTo 9088
End If
'debug'





Set j = mydb.OpenRecordset("Joblist")
j.Index = "ENTOLH"

Set p = mydb.OpenRecordset("prospau2")
p.Index = "ENTOLH"

'Set m = PROSMACB.Recordset
If r.RecordCount < 2 Then
   MsgBox "The file is empty."
   Command3.Enabled = True

   Exit Sub
End If

Do While Not r.eof   'until find 1 recipe
 If Not IsNull(r(0)) Then 'new recipy
   If InStr(1, r(0), "Recipe") > 0 And InStr(1, r(0), "Standard") > 0 Then
     'work.BeginTrans
     
     m_entolh = Mid$(r(0), 18, 18)
     Label4.Caption = LTrim(str(nFile)) + "  " + Mid$(r(0), InStr(1, r(0), "Recipe") + 7, 10)
     Text5.text = Mid$(r(0), 18, 18)
     
     
     mm0 = 0
     r.Move 5
     mm = r.Bookmark
    ' On Error Resume Next
     mcc = -1
     For k = 0 To 20
            r.MoveNext
            If InStr(1, r(0), "Total Concentration") > 0 Then Exit For
            If Not r.eof Then
              If Not IsNull(r(0)) Then
                 m1% = InStr(1, r(0), "  %  ")
                If m1% > 1 Then
                   mcc = mcc + 1
                   Label7(mcc).Caption = Val(Mid$(r(0), m1% - 6, 8)) '0.79432
                   
                   Label6(mcc).Caption = Mid$(r(0), IIf(Asc(Mid$(r(0), 1, 10)) = 10, 8, 1), 20) 'code
                   mm0 = mm0 + Val(Label7(mcc).Caption)
                End If
              End If
            End If
     Next
     Label7(4).Caption = mm0
     On Error GoTo 0
     
     
     
     r.Bookmark = mm
    ' r.Move 5
     
     
     If IsNumeric(m_entolh) Then
         m_entolh = Right("000000" + LTrim(str(Val(m_entolh))), 6) 'last
     Else
         
     End If
     
     m_entolh = Val(m_entolh)
          If m_entolh = 0 Then
      Else
          j.Seek "=", Right("000000" + LTrim(str(Val(m_entolh))), 6) 'last
     End If
     If j.NoMatch Or m_entolh = 0 Then
         j.MoveLast
         If IsNull(j("ENTOLH")) Then
             m = 1
         Else
             m = Val(j("ENTOLH")) + 1
         End If
         Label5.Caption = m
         Command2.Enabled = True
         Do While Command2.Enabled = True
           DoEvents
         Loop
         m_entolh = Right("000000" + LTrim(str(Val(m))), 6) 'last
         
       On Error Resume Next
         j.AddNew
         j("entolh") = m_entolh
         Form1.Caption = m_entolh
         j("kod_pel") = DBCombo1.text
         Form1.Caption = DBCombo1.text
         j("baros_pani") = Val(Text2.text)
         Form1.Caption = Val(Text2.text)
         j("sxesh_mpan") = Val(Text1.text)
         Form1.Caption = Val(Text2.text)
         j("kod_erg") = 0
         Form1.Caption = "kod_erg"
         j("xhmika") = Val(Text3.text)
          
         j("apoxrvsh") = left(Text5.text, 20)
         j("eidos_baf") = left(DBCombo2.text, 15)
         j.update
         On Error GoTo 0
         
         
         
         Form1.Caption = "J/update"
         prosp = 1
      Else
         
         p.Seek "=", j("entolh") 'BRISKO TON MEGALITERO ARIUMO PROSPAUIAS
         m_aa = p("AA_PROSPAU")
         Form1.Caption = "m_aa"
         Do While p.ENTOLH = m_entolh And Not p.eof
            If p("AA_PROSPAU") > m_aa Then
               m_aa = p("AA_PROSPAU")
               Form1.Caption = "r/m_aa2"
            End If
            p.MoveNext
         Loop
         prosp = Val(m_aa)
         Form1.Caption = "r/m_aa3"
     End If
     
     For k = 0 To 3
            Label6(k).Caption = " "
            Label7(k).Caption = " "
     Next
     
          r.MoveNext
          m_Seira = 0
     '----------------------------------
  Do While Not r.eof   ' loop sta systatika ths syntaghs
    If Not IsNull(r(0)) Then
      If InStr(1, r(0), "Total Concentration") > 0 Then Exit Do
      If InStr(1, r(0), "  %  ") Then 'new recipy
         'm_ek = Val(Mid$(r(0), 27, 8))  '0.79432
         
         
                m1% = InStr(1, r(0), "  %  ")
                m_ek = Val(Mid$(r(0), m1% - 6, 8)) '0.79432
                   
                For k = 1 To 20
                   If Asc(Mid$(r(0), k, 1)) >= 48 Then Exit For
                Next
                m1% = InStr(k, r(0), " ")
                If m1% = 0 Or m1% > 12 Then m1% = k + 4
                M_KOD = Mid$(r(0), k, m1% - k)
            
         
         
         
         
         'm_kod = Mid$(r(0), 1, 4) 'code
         If m_ek <= 0 Or Len(Trim(M_KOD)) < 2 Then
           x_kod = " "
         Else
           x_kod = find_kod(M_KOD, m_ek)
         End If
       If left(x_kod, 1) <> " " And Val(x_kod) > 0 Then
         m_Seira = m_Seira + 1
         p.AddNew
         p("entolh") = m_entolh 'last
         p("aa_prospau") = Right("00" + LTrim(str(prosp)), 2)
         p("kod_pel") = DBCombo1.text
         p("baros_pani") = Val(Text2.text)
         p("sxesh_mpan") = Val(Text1.text)
         p("ek") = m_ek
         p("gl") = 0
         p("kod") = x_kod
         p("SEIRA") = m_Seira
         p("alati") = Val(Text4.text)
         p("ayjon") = m_Seira
         p("apoxrvsh") = left(Text5.text, 20)
         p("status") = 0
         p("APOR") = 0
         p.update
         If m_Seira <= 4 And Not IsNull(x("perigrafh")) Then
            Label6(m_Seira - 1).Caption = Perigrafh
            Label7(m_Seira - 1).Caption = Format(m_ek, "#0.0000")
         End If
        End If
      End If 'instr
    End If 'isnull(r(0))
        r.MoveNext
  Loop
     '----------------------------------
     
     ' salt
     ans = MsgBox("Save this recipe ?", vbYesNo)
     If ans = vbYes Then
        'work.CommitTrans
     Else
        'work.Rollback
        mydb.Execute "delete *from prospau2 where entolh='" + m_entolh + "'"
        mydb.Execute "delete *from joblist where entolh='" + m_entolh + "'"
     End If
     For k = 0 To 3
            Label6(k).Caption = " "
            Label7(k).Caption = " "
     Next
     
     Exit Do
    Else
      r.MoveNext
    End If ' instr(recipe
  Else
     r.MoveNext
  End If

Loop
r.Close
ans = MsgBox(" Do you want to delete the text file ? ", vbYesNo, "The import is finished.")
If ans = vbYes Then
   'r.Close
   On Error GoTo cantdelete
       Kill DIRECT + LTrim(str(nFile)) + ".prn"
   On Error GoTo 0
End If

9088
Next nFile



Unload Me
Exit Sub

cantdelete:
MsgBox "The file can not deleted. Another user uses the same file"
Resume Next

End Sub


Private Sub Polydye_Click()
' POLYDYE LTD
Dim DUM, latos
Dim mm, mm0, work, mydb, r, text, j, p, k, m, prosp, m_aa, m_Seira, m_ek, M_KOD, x_kod, ans
Dim Pelaths
Pelaths = "AGMONT"
On Error Resume Next

Command3.Enabled = False
Polydye.Enabled = False
ColorTe.Enabled = False
damapong.Enabled = False
Command1.Enabled = False
TreePoint.Enabled = False
DHMAS.Enabled = False
FARBETEX.Enabled = False

Form1.Caption = "The position for chemicals is 71"
'default μπάνιο
If Pelaths = "AGMONT" Then
    Text4.text = 0
Else
    Text4.text = 1
    Label15.Caption = "Recognize ID=CHEMIC"
End If

Screen.MousePointer = 11
DUM = Shell("c:\talos\merge.bat")
MilSec 10000
Screen.MousePointer = 1
Command1.Enabled = False
Set work = Workspaces(0)
Set mydb = work.OpenDatabase("c:\TALOS\RECIPIES.MDB", False, False)
Set r = mydb.OpenRecordset("parametroi")
    Text1.text = r("MPANIO")
    Text2.text = r("default_baros_panioy")
     r.Close

'Data2.DatabaseName = Path.text
'Data2.RecordSource = filename.text
'Data2.Refresh
'DBGrid1.Columns(0).Name = "robolab"
'DBGrid1.Refresh

Set x = mydb.OpenRecordset("ximitech")
x.Index = "KOD"







'Set r = Data1.Recordset

Set text = OpenDatabase(Path.text, False, False, "Text;")
Set r = text.OpenRecordset(Filename.text)

Set j = mydb.OpenRecordset("Joblist")
j.Index = "ENTOLH"

Set p = mydb.OpenRecordset("prospau2")
p.Index = "ENTOLH"

'Set m = PROSMACB.Recordset
If r.RecordCount < 2 Then
   MsgBox "The file is empty."
   Command3.Enabled = True

   Exit Sub
End If
Do While Not r.eof
 
 If Len(r(0)) > 43 Then 'new recipy
     
     'work.BeginTrans
     
     m_entolh = Mid$(r(0), 25, 16)
     Label4.Caption = m_entolh
     
   '  r.MoveNext
     
     If r.eof Then Exit Do
     
     mm = r.Bookmark
     On Error Resume Next
     mm0 = 0
     'r.MoveNext
     For k = 0 To 3
            r.MoveNext
            If r.eof Then Exit For
            If Mid$(r(0), 25, 5) = "WATER" Then Exit For
            If Not r.eof Then
              If Len(Trim(Mid$(r(0), 25, 8))) > 2 And Len(Trim(r(0))) < 50 Then
                 Label7(k).Caption = Val(Mid$(r(0), 35, 7)) '0.79432
                 Label6(k).Caption = Mid$(r(0), 25, 8) 'code
                 mm0 = mm0 + Val(Label7(k).Caption)
              End If
            End If
     Next
     If r.eof Then Exit Do
     Label7(4).Caption = mm0
     On Error GoTo 0
     
     
     
     r.Bookmark = mm
     
     
     
     If IsNumeric(m_entolh) Then
         m_entolh = Right("000000" + LTrim(str(Val(m_entolh))), 6) 'last
     Else
         
     End If
     
     m_entolh = Val(m_entolh)
          If m_entolh = 0 Then
      Else
          j.Seek "=", Right("000000" + LTrim(str(Val(m_entolh))), 6) 'last
     End If
     If j.NoMatch Or m_entolh = 0 Then
         
         If j.RecordCount = 0 Then
            m = 1
         Else
            latos = 0
            On Error GoTo pane_proti
            j.MoveLast
            On Error GoTo 0
            
            If latos = 1 Then
                m = 1
            Else
               If IsNull(j("ENTOLH")) Then
                 m = 1
               Else
                  m = Val(j("ENTOLH")) + 1
               End If
            End If
            
            
         End If
         
         Label5.Caption = m
         Command2.Enabled = True
         Do While Command2.Enabled = True
           DoEvents
         Loop
         m_entolh = Right("000000" + LTrim(str(Val(m))), 6) 'last
         
       On Error Resume Next
         j.AddNew
         j("entolh") = m_entolh
         Form1.Caption = m_entolh
         j("kod_pel") = DBCombo1.text
         Form1.Caption = DBCombo1.text
         j("baros_pani") = Val(Text2.text)
         Form1.Caption = Val(Text2.text)
         j("sxesh_mpan") = Val(Text1.text)
         j("eidos_pani") = DBCombo3.text
         Form1.Caption = Val(Text2.text)
         j("kod_erg") = 0
         Form1.Caption = "kod_erg"
         j("xhmika") = Val(Text3.text)
          
         j("apoxrvsh") = left(Text5.text, 20)
         j("eidos_baf") = left(DBCombo2.text, 15)
         j.update
         On Error GoTo 0
         
         
         
         Form1.Caption = "J/update"
         prosp = 1
      Else
         
         p.Seek "=", j("entolh") 'BRISKO TON MEGALITERO ARIUMO PROSPAUIAS
         m_aa = p("AA_PROSPAU")
         Form1.Caption = "m_aa"
         Do While p.ENTOLH = m_entolh And Not p.eof
            If p("AA_PROSPAU") > m_aa Then
               m_aa = p("AA_PROSPAU")
               Form1.Caption = "r/m_aa2"
            End If
            p.MoveNext
         Loop
         prosp = Val(m_aa)
         Form1.Caption = "r/m_aa3"
     End If
     
     For k = 0 To 3
            Label6(k).Caption = " "
            Label7(k).Caption = " "
     Next
     
          r.MoveNext
          m_Seira = 0
     Do While Not r.eof
         If IsNull(r(0)) Then Exit Do 'new recipy
         If Mid$(r(0), 25, 5) = "WATER" Then Exit Do
         
         m_ek = Val(Mid$(r(0), 35, 7)) '0.79432
         Form1.Caption = "r/m_ek"
         M_KOD = Trim(Mid$(r(0), 25, 8)) 'code
         Form1.Caption = "r/code"
         If m_ek <= 0 Or Len(Trim(M_KOD)) < 2 Then
           x_kod = " "
         Else
           x_kod = find_kod(M_KOD, m_ek)
         End If
                '           x_kod = " "
         Form1.Caption = "r/x_kod:"
       
       If left(x_kod, 1) <> " " Then
         m_Seira = m_Seira + 1
         p.AddNew
         p("entolh") = m_entolh 'last
'         Form1.Caption = "p/" + m_entolh
         p("aa_prospau") = Right("00" + LTrim(str(prosp)), 2)
 '        Form1.Caption = "p/aa"
         p("kod_pel") = DBCombo1.text
  '       Form1.Caption = "p/kodpel"
         p("baros_pani") = Val(Text2.text)
   '      Form1.Caption = "p/baros_pani"
         p("sxesh_mpan") = Val(Text1.text)
    '     Form1.Caption = "p/sxesh_mpan"
         p("ek") = m_ek
     '    Form1.Caption = "p/m_ek"
         p("gl") = 0
      '   Form1.Caption = "p/gl"
         p("kod") = x_kod
       '  Form1.Caption = "p/x_kod"
         p("SEIRA") = m_Seira
        ' Form1.Caption = "p/m_seira"
         p("ayjon") = m_Seira
         p("apoxrvsh") = left(Text5.text, 20)
        ' p("eidos_baf") = j("eidos_baf")
         'Form1.Caption = "p/ayjon"
          p("alati") = Val(Text6.text)
         p("status") = 0
         p("APOR") = 0
         p.update
         Form1.Caption = "p/update"
         If m_Seira <= 4 And Not IsNull(x("perigrafh")) Then
            Label6(m_Seira - 1).Caption = Perigrafh
            If Not IsNull(x("real_cons")) Then
               If x("real_cons") > 100 * m_ek Then
                  Label7(m_Seira - 1).BackColor = vbRed
               End If
            Else
                Label7(m_Seira - 1).BackColor = &H8000000F
            End If
            
            Label7(m_Seira - 1).Caption = Format(m_ek, "#0.0000")
         End If
        End If
         
         r.MoveNext
     Loop
     
     
     
     
     ' salt
      If Val(Text4.text) > 0 Then
         
         m_ek = Val(Text4.text) '0.79432
         M_KOD = "CHEMIC" 'code
         
         If M_KOD = "CHEMIC" Then
           x_kod = FIND_XHM(M_KOD)
         ElseIf m_ek <= 0 Or Len(Trim(M_KOD)) < 2 Then
           x_kod = " "
         Else
           x_kod = find_kod(M_KOD, m_ek)
         End If
         
       End If
       
       
       
       If Val(Text4.text) > 0 And Len(Trim(x_kod)) > 2 Then
          m_Seira = m_Seira + 1
          p.AddNew
         p("entolh") = m_entolh 'last
         p("aa_prospau") = Right("00" + LTrim(str(prosp)), 2)
         p("kod_pel") = DBCombo1.text
         p("baros_pani") = Val(Text2.text)
         p("sxesh_mpan") = Val(Text1.text)
         p("ek") = 0
         p("gl") = m_ek
         p("kod") = x_kod
         p("SEIRA") = m_Seira
         p("ayjon") = m_Seira
         p("status") = 0
         p("APOR") = 0
         p.update
       End If
     
     
     
     ans = MsgBox("Save this recipe ?", vbYesNo)
     If ans = vbYes Then
        'work.CommitTrans
     Else
        'work.Rollback
        mydb.Execute "delete *from prospau2 where entolh='" + m_entolh + "'"
        mydb.Execute "delete *from joblist where entolh='" + m_entolh + "'"
     End If
     
     
     For k = 0 To 3
            Label6(k).Caption = " "
            Label7(k).Caption = " "
     Next
     
     
     
     If r.eof Then Exit Do
  Else
     r.MoveNext
  End If
Loop
r.Close
ans = MsgBox(" Do you want to delete the text file ? ", vbYesNo, "The import is finished.")

If ans = vbYes Then
     On Error GoTo cantdelete
   Kill LTrim(Path.text + "\" + Filename.text)
End If




'If ans = vbYes Then
   'r.Close
 '  On Error GoTo cantdelete
  ' FileCopy "c:\talos\emptytxt.txt", Path.text + "\" + filename.text
   'On Error GoTo 0
'End If

Unload Me
Exit Sub


pane_proti:
   latos = 1
Resume Next
Return

cantdelete:
MsgBox "The file can not deleted. Another user uses the same file"
Resume Next

End Sub






Private Static Sub disable_imports()
FARBETEX.Enabled = False
DHMAS.Enabled = False
TreePoint.Enabled = False
Command1.Enabled = False
Polydye.Enabled = False
ColorTe.Enabled = False
damapong.Enabled = False

End Sub

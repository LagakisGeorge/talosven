VERSION 4.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   2250
   ClientWidth     =   9420
   Height          =   7380
   Icon            =   "_PARALHL.frx":0000
   Left            =   0
   LinkTopic       =   "MDIForm1"
   Picture         =   "_PARALHL.frx":0442
   Top             =   1620
   Width           =   9540
   WindowState     =   2  'Maximized
   Begin VB.Menu Arxeia 
      Caption         =   "������"
      Enabled         =   0   'False
      Index           =   10
      Begin VB.Menu ������ 
         Caption         =   "������"
         Index           =   11
      End
      Begin VB.Menu ������� 
         Caption         =   "�������"
         Index           =   12
      End
      Begin VB.Menu ����������� 
         Caption         =   "�����������"
         Index           =   13
      End
      Begin VB.Menu pania 
         Caption         =   "�����"
         Index           =   14
      End
      Begin VB.Menu Reindex 
         Caption         =   "������������ �������"
         Index           =   15
      End
      Begin VB.Menu Mhded_joblist 
         Caption         =   "���������� ������� ��������"
         Index           =   16
      End
   End
   Begin VB.Menu Syntages 
      Caption         =   "��������"
      Index           =   20
      Begin VB.Menu �������� 
         Caption         =   "��������"
         Index           =   21
      End
      Begin VB.Menu ���� 
         Caption         =   "����� ��������"
         Enabled         =   0   'False
         Index           =   22
      End
      Begin VB.Menu Import 
         Caption         =   "Import"
      End
   End
   Begin VB.Menu bohuhtika 
      Caption         =   "���������"
      Enabled         =   0   'False
      Index           =   30
      Begin VB.Menu ferta 
         Caption         =   "���� ���������"
         Index           =   32
      End
      Begin VB.Menu phgaine 
         Caption         =   "������� ���������"
         Index           =   33
      End
      Begin VB.Menu make_dial 
         Caption         =   "��������� ����������"
         Index           =   34
      End
      Begin VB.Menu check_kentra 
         Caption         =   "������� ��������������"
         Index           =   36
      End
      Begin VB.Menu AXRHSIMOPOIHTA 
         Caption         =   "������� �� �������������� ���������"
         Index           =   37
      End
      Begin VB.Menu Bohu 
         Caption         =   "���������� ��������"
         Index           =   35
         Begin VB.Menu Colors 
            Caption         =   "Colors"
            Index           =   36
         End
         Begin VB.Menu dfsfdf 
            Caption         =   "����������"
            Index           =   37
         End
         Begin VB.Menu loipes 
            Caption         =   "������������ ������ ��������"
            Index           =   38
         End
      End
   End
   Begin VB.Menu ejodos 
      Caption         =   "������"
      Index           =   40
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim m_unloading As Integer, m_menu

Private Sub ������_Click(Index As Integer)
 Screen.MousePointer = 11
' ������.Visible = False
  frmXHM.Show
End Sub

Private Sub �����������_Click(Index As Integer)
     Screen.MousePointer = 11
     
     Menu = 2
     
     PELATES.Show
     'PROCEDURE.Show
End Sub
Private Sub ��������_Click(Index As Integer)
     Screen.MousePointer = 11
     mJOBLIST.Show
End Sub

Private Sub AXRHSIMOPOIHTA_Click(Index As Integer)

'-----------------
Dim mydb, p, x, DT, ORES, counter
Dim AR_S As Integer




Set mydb = Workspaces(0).OpenDatabase("c:\TALOS", False, False, "dBASE IV;")

Set x = mydb.OpenRecordset("XIMITECH")

DT = AXRHSIMOPOIHTES_ORES * 3600
counter = 0
AR_S = 0
x.MoveFirst
Menu = 3
Do While Not x.EOF
         If IsNull(x("addr_prot")) Then GoTo skip1
         If Val(x("addr_prot")) = 0 Then GoTo skip1
         If IsNull(x("last_dosom")) Then GoTo skip1
         If DateDiff("S", CDate(x("Last_Dosom")), Now) < DT Then GoTo skip1
          counter = counter + 1
skip1:
    x.MoveNext
Loop

mydb.Close

If counter = 0 Then
   MsgBox "��� �������� ������������� ��� ������ �� ������������."
   Exit Sub
End If

ans = MsgBox("Y������� " + Str(counter) + " ��� ����������.N� ��������� ��� ����������;", vbYesNo)
If ans = vbNo Then
   Exit Sub
End If


MDIForm1.arxeia(10).Visible = False
MDIForm1.Syntages(20).Visible = False
MDIForm1.Bohuhtika(30).Visible = False
MDIForm1.ejodos(40).Visible = False







Set MyDB1 = Workspaces(0).OpenDatabase("c:\TALOS", False, False, "dBASE IV;")
sql = "UPDATE PROSPAU2 SET epil=0 WHERE STATUS<2;"
MyDB1.Execute sql
MyDB1.Close





Set mydb = Workspaces(0).OpenDatabase("c:\TALOS", False, False, "dBASE IV;")
Set p = mydb.OpenRecordset("PROSPAU2")
Set x = mydb.OpenRecordset("XIMITECH")
counter = 0
AR_S = 0
x.MoveFirst
Menu = 3
Do While Not x.EOF
         If IsNull(x("addr_prot")) Then GoTo skip
         If Val(x("addr_prot")) = 0 Then GoTo skip
         If IsNull(x("last_dosom")) Then GoTo skip
         If DateDiff("S", CDate(x("Last_Dosom")), Now) < DT Then GoTo skip
              p.AddNew
              AR_S = AR_S + 1
              p("ENTOLH") = "000001"
              p("SXESH_MPAN") = 10
              p("AA_PROSPAU") = "01"
               counter = counter + 1
              p("SEIRA") = counter
              p("STATUS") = 1
              p("EPIL") = 99
              p("AYJON") = counter
              p("KOD") = x("KOD")
              p("EK") = 0.01
              p("UserW") = User_ID
              p("hme") = Date
              p("apor") = 1.2
              p("GL") = 0
              p("nero") = 0
              p("gr") = 1500
              p("baros_pani") = 5000
              p.update
skip:
    x.MoveNext
Loop

frmSYNT.Show




End Sub

Private Sub check_kentra_Click(Index As Integer)
    ch_kentr.Show
End Sub

Private Sub colors_Click(Index As Integer)
       Farben.Show
     End Sub
     

Private Sub �������_Click(Index As Integer)
    Screen.MousePointer = 11
Menu = 1
    PELATES.Show
End Sub

Private Sub dfsfdf_Click(Index As Integer)
   PARAMETROI.Show
End Sub

Private Sub ����_Click(Index As Integer)
    Menu = 7 'flag ��� �� ������������ �� ��������� �������� ����� ��
             ' ������
    Screen.MousePointer = 11
    WAIT.Show
End Sub

Private Sub �������������_Click(Index As Integer)
    Screen.MousePointer = 11
     Form1.Show
End Sub



Private Sub ejodos_Click(Index As Integer)
  
  Dim dbsjet As Database
  Dim r As Recordset
  
If system_ready = 0 Then End

  
If m_unloading = 0 Then
     response = MsgBox(mL_CapRes(70), 20, mL_CapRes(71))
     If response = 7 Then ' akyro
        Exit Sub
     End If
End If
  
  m_unloading = 1
 Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\coord_TB.mdb")
      Set r = dbsjet.OpenRecordset("mem3")

        r.MoveFirst
        r.Edit
        r("XRHSTHS") = "NO BUSY"
        r("HOME") = "HOME=OK"
       r.update
       r.Close
       
    Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\watch.mdb")
    Set r = dbsjet.OpenRecordset("moves")
             r.AddNew
               r("movement") = left$("END SYNTAGES", 30)
               r("datetime") = Now
             r.update
             r.Close
              
       

 Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\recipies.mdb")
      Set r = dbsjet.OpenRecordset("userinout")

r.MoveLast

r.Edit
  r("userOut1") = Now
r.update

r.Close


 '  cmd$ = "!out.1-0:!out.2-0:!out.3-1:!out.4-0:!out.5-0:!out.6-0:!out.7-0:!out.8-0:!out.9-0:!out.10-0:!out.13-0:!out.14-0:"
 '  temp% = SendAT6400Block(768, cmd$, 0)
ValveReset
   cmd$ = "!k:!MIDEN:"
   TEMP% = SendAT6400Block(768, cmd$, 0)

  
  
  
  
  
   
End
  
End Sub

Private Sub ferta_Click(Index As Integer)
    Screen.MousePointer = 11
   FERE.Show
End Sub

Private Sub Import_Click()
  Text44.Show
End Sub

Private Sub loipes_Click(Index As Integer)
EPANEKTYP.Show
End Sub

Private Sub make_dial_Click(Index As Integer)
    fereSK.Show
End Sub

Private Sub MDIForm_Load()
   Dim dbsjet As Database
   Dim r As Recordset, MHome, m_err

If App.PrevInstance Then End

system_ready = 0
m_Cmd$ = "    "
On Error GoTo 12
m_err = 0
m_menu = 0
 m_unloading = 0
 ' ----------- m u l t i l i n g u a l -----------------
 MDIForm1.Caption = "" ' mL_CapRes(9) ' "�������� �����"
 arxeia(10).Caption = mL_CapRes(10)   ' "������"
 ������(11).Caption = mL_CapRes(11)  '"������"
 �������(12).Caption = mL_CapRes(12)   '"�������"
 �����������(13).Caption = mL_CapRes(13)  '"�����������"
 pania(14).Caption = mL_CapRes(14) ' "�����"
 Reindex(15).Caption = mL_CapRes(15)  '"������������ �������"
 Syntages(20).Caption = mL_CapRes(20)   '"��������"
 ��������(21).Caption = mL_CapRes(21)  '"��������"
 ����(22).Caption = mL_CapRes(22) ' "����� ��������"
 Bohuhtika(30).Caption = mL_CapRes(30) ' "���������"
' �������������(31).Caption = mL_CapRes(31) '"�������������"
 ferta(32).Caption = mL_CapRes(32) ' "���� ���������"
 phgaine(33).Caption = mL_CapRes(33) ' "������� ���������"
 make_dial(34).Caption = mL_CapRes(34) '"��������� ����������"
 ejodos(40).Caption = mL_CapRes(40)  '"������"
  ' ----------- m u l t i l i n g u a l -----------------
  epil_Menu = 0 ' ����� ��� ����� �������
    
    
    
DBEngine.IniPath = "C:\talos\talos.INI"
  
  'check_up dbase driver
  m_err = "checkfil.dbf,c:\talos"

  
  
  
  
  
  
  
' find_colors
  

m_err = "parametroi,c:\talos\recipies.mdb"
Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\recipies.mdb")
Set r = dbsjet.OpenRecordset("PARAMETROI")
  allazo_xromata = r("allazo_xromata")
  allazo_fonts = r("allazo_fonts")
  prozygismena = r("prozygismena")
   dokimastikh_syntagh = r("dokimastikh_syntagh")
  AXRHSIMOPOIHTES_ORES = r("AXRHSIMOPOIHTES_ORES")
  Default_baros_panioy = r("default_baros_panioy")
  Ar_tel_mpoykal = r("ar_tel_mpoykal")
  Ar_1_mpoykal = r("ar_1_mpoykal")
 
  If IsNull(r("doyleyo_alati")) Then
     DOYLEYO_ALATI = 0
  Else
     DOYLEYO_ALATI = r("doyleyo_alati")
  End If
  UESH_ALATIOY = r("UESH_ALATIOY")
 
 
 
 
 
 r.Close
       
m_err = "users,c:\talos\recipies.mdb"
       
 userinput.Show
' MDIForm1.Caption = kentrikh.Caption + User_ID
       
   Exit Sub
   
  
  
' �� ����� ��� ���� ��� ������� ===========================
            'Response = MsgBox(mL_CapRes(70), 20, mL_CapRes(71))
            'MsgBox mL_CapRes(72), , "72.Talos"  ' "���������� �� Emergency Stop  ' Define message.
            'MsgBox mL_CapRes(73), , "73.Talos"
            'MsgBox mL_CapRes(74), , "74.Talos"
    '      MsgBox mL_CapRes(75), , "75.Talos"
        '   MsgBox mL_CapRes(76), , "Talos"
' �� ����� ��� ���� ��� ������� ===========================




m_err = "IsOSLoaded(768) "
system_ready = IsOSLoaded(768)

m_err = "findvalves,c:\talos\COORD_TB.MDB"

FindValves

If system_ready = 0 Then
    MsgBox "��� ��������� �� ����������� ��� robot"
    Exit Sub
End If
    
m_err = "ISCUTOFF(0)"
    
    
    
    
  
    
   Do While isCutOff(0) = 1
        MsgBox mL_CapRes(72), , "72.Talos"  ' "���������� �� Emergency Stop  ' Define message.
       'MsgBox "���������� �� Emergency Stop  ", , "Talos"  ' Define message.
   Loop
 
 
m_err = "mem3,c:\talos\coord_tb.mdb"
 
   
       Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\coord_tb.mdb")
      Set r = dbsjet.OpenRecordset("mem3")

     r.MoveFirst
      
      If r("XRHSTHS") = "JOYSTICK" Then
         MsgBox mL_CapRes(73), , "73.Talos"
         ' MsgBox "������� !!!  .������ ��������������� �� Joystick ����� ���������� �� ���������� � ���������� Home. �� ����� ������������ ��� ����������.", , "Talos"
         '  MsgBox "������� !!!  .������ ��������������� �� Joystick ����� ���������� �� ���������� � ���������� Home. �� ����� ������������ ��� ����������.", , "Talos"
           End
       End If
      
      If r("HOME") = "HOME=OK" Then
        '
      Else
             MsgBox mL_CapRes(74), , "74.Talos"
            'MsgBox "������� !!!  .������� � �������� ��� ������������." + Chr$(13) + "�� ����� ���������� Home.", , "Talos"
            'MsgBox "������� !!!  .������� � �������� ��� ������������." + Chr$(13) + "�� ����� ���������� Home.", , "Talos"
           End
      End If
      
      
      
      If r("XRHSTHS") = "BUSY" Then
      
           MsgBox mL_CapRes(75), , "75.Talos"
     '     MsgBox "������� !!!  .�� ��������� ��� ���������� � ��� ���� ���������� �����.� ������� ����������.", , "Talos"
     '     MsgBox "������� !!!  .�� ��������� ��� ���������� � ��� ���� ���������� �����.� ������� ����������.", , "Talos"
           End
       End If
      r.Edit
       r("HOME") = "NO HOME"
       r("XRHSTHS") = "BUSY"
       r.update
       r.Close
       
             Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\watch.mdb")
             Set r = dbsjet.OpenRecordset("moves")
             r.AddNew
               r("movement") = left$("BEGIN SYNTAGES", 30)
               r("datetime") = Now
             r.update
             r.Close
       
       
       
       
       
       
m_err = "var,c:\talos\coord_tb.mdb"
       
     Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\coord_tb.mdb")
     Set r = dbsjet.OpenRecordset("var")
            Do Until r.EOF
              V = r("field_name")
              t = r("field_val")
              If IsNull(t) Then
                t = 0
              End If
              cmd$ = "!" + V + "=" + Str$(t) + ":"
              
              TEMP% = SendAT6400Block(768, cmd$, 0)
             r.MoveNext
        Loop
      r.Close
On Error GoTo 0
       
       
       
'      MsgBox "������� ��� ���� �������� ��� HOME  ", , "Talos"
 If error(5, 1) = 1 Then
 
           MsgBox mL_CapRes(76), , "Talos"
   '       MsgBox "������� ����������� ���� ��������� ��������� Home  ", , "Talos"
   '       MsgBox "������� ����������� ���� ��������� ��������� Home  ", , "Talos"
            End
  End If
       
 '  cmd$ = "!out.1-0:!out.2-0:!out.3-1:!out.4-0:!out.5-0:!out.6-0:!out.7-0:!out.8-0:!out.9-0:!out.10-0:!out.13-0:!out.14-0:"
 '  temp% = SendAT6400Block(768, cmd$, 0)
       
  Exit Sub
12:
If Err.Number = 53 Or Err.Number = 3011 Then
     MsgBox "����� ���������� �� ��������� �� ������ " + m_err
     End
     End If
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
If m_unloading = 0 Then
   m_unloading = 1
   ejodos_Click (1)
End If
End Sub


Private Sub Mhded_joblist_Click(Index As Integer)
  Form1.Show
End Sub

Private Sub pania_Click(Index As Integer)
    Screen.MousePointer = 11
Menu = 3
    PELATES.Show

End Sub

Private Sub phgaine_Click(Index As Integer)
     Screen.MousePointer = 11
    PANE.Show
End Sub


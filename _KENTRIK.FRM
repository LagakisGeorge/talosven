VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   ClientHeight    =   3810
   ClientLeft      =   1770
   ClientTop       =   2280
   ClientWidth     =   5880
   Icon            =   "_KENTRIK.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "_KENTRIK.frx":0442
   Begin VB.Menu Arxeia 
      Caption         =   "Αρχεία"
      Index           =   10
      Begin VB.Menu Χημικά 
         Caption         =   "Χημικά"
         Index           =   11
      End
      Begin VB.Menu Πελάτες 
         Caption         =   "Πελάτες"
         Index           =   12
      End
      Begin VB.Menu Διαδικασίες 
         Caption         =   "Διαδικασίες"
         Index           =   13
      End
      Begin VB.Menu pania 
         Caption         =   "Πανιά"
         Index           =   14
      End
   End
   Begin VB.Menu Syntages 
      Caption         =   "Συνταγές"
      Index           =   20
      Begin VB.Menu Εργασίες 
         Caption         =   "Εργασίες"
         Index           =   21
      End
      Begin VB.Menu Πάνε 
         Caption         =   "Λίστα Αναμονής"
         Index           =   22
      End
      Begin VB.Menu import 
         Caption         =   "Import"
         Index           =   23
      End
      Begin VB.Menu import2 
         Caption         =   "Import.2(new version)"
         Index           =   24
      End
   End
   Begin VB.Menu bohuhtika 
      Caption         =   "Βοηθητικά"
      Index           =   30
      Begin VB.Menu ferta 
         Caption         =   "Φέρε Μπουκάλια"
         Index           =   32
      End
      Begin VB.Menu phgaine 
         Caption         =   "Πήγαινε Μπουκάλια"
         Index           =   33
      End
      Begin VB.Menu make_dial 
         Caption         =   "Κατασκευή Διαλυμάτων"
         Index           =   34
      End
      Begin VB.Menu AXRHSIMOPOIHTA 
         Caption         =   "Συνταγή με αχρησιμοποίητα μπουκάλια"
         Index           =   37
      End
      Begin VB.Menu Bohu 
         Caption         =   "Βοηθητικές εργασίες"
         Index           =   35
         Begin VB.Menu paramet 
            Caption         =   "Παράμετροι"
            Index           =   37
         End
         Begin VB.Menu loipes 
            Caption         =   "Επανεκτύπωση παλιών συνταγών"
            Index           =   38
         End
         Begin VB.Menu Reindex 
            Caption         =   "Ανακατασκευή Αρχείων"
            Index           =   15
         End
         Begin VB.Menu Mhden_joblist 
            Caption         =   "Μηδενισμός Αρχείου Εργασιών"
            Index           =   16
         End
      End
   End
   Begin VB.Menu ejodos 
      Caption         =   "Εξοδος"
      Index           =   40
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Dim m_unloading As Integer, m_menu
Dim epitr_syntages

Function Find_ANTICODE(DUM)
' arx = Year(Now) * 12 + Month(Now) * 24 + Day(Now) * 36 + Hour(Now) * 48 + Minute(Now) * 72 + Second(Now) * 96
Dim a, c, b
a = LTrim(str(DUM))

c = 0
For k = 1 To Len(a)
    c = c + Val(Mid$(a, k, 1)) ^ (k)
Next

c = c + Val(a)
Find_ANTICODE = c

End Function

Function Get_ComDLG(value)
'
Dim Mycom, mycomset

  Set Mycom = Workspaces(0).OpenDatabase(GetWindowsDir() + "\SYSTEM", False, False, "dBASE IV;")
  Set mycomset = Mycom.OpenRecordset("COMDLG16.DLL")
  
  Get_ComDLG = mycomset("k")
  mycomset.Close

End Function

Sub Write_Exo()
    Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\watch.mdb")
           Set r = dbsjet.OpenRecordset("commun")
           On Error Resume Next
           r.EDIT
           r("bges_exo") = 1
           r.update
           r.Close
           On Error GoTo 0

End Sub

Private Sub Χημικά_Click(Index As Integer)
 Screen.MousePointer = 11
' Αρχεία.Visible = False
  frmXHM.Show
End Sub

Private Sub Διαδικασίες_Click(Index As Integer)
     Screen.MousePointer = 11
     
     Menu = 2
     
     pelates.Show
     'PROCEDURE.Show
End Sub
Private Sub Εργασίες_Click(Index As Integer)
     Screen.MousePointer = 11
     mJOBLIST.Show
End Sub

Private Sub AXRHSIMOPOIHTA_Click(Index As Integer)

'-----------------
Dim mydb, p, x, DT, ORES, counter
Dim AR_S As Integer




Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)

Set x = mydb.OpenRecordset("XIMITECH")

DT = AXRHSIMOPOIHTES_ORES * 3600
counter = 0
AR_S = 0
x.MoveFirst
Menu = 3
Do While Not x.eof
         If IsNull(x("addr_prot")) Then GoTo skip1
         If Val(x("addr_prot")) = 0 Then GoTo skip1
         If IsNull(x("last_dosom")) Or Not IsDate(x("last_dosom")) Then GoTo skip1
         If DateDiff("S", CDate(x("Last_Dosom")), Now) < DT Then GoTo skip1
          counter = counter + 1
skip1:
    x.MoveNext
Loop

mydb.Close

If counter = 0 Then
   MsgBox mL_Res(52) ' "Δεν υπάρχουν χημικοτεχνικά που πρέπει να σερβιριστούν."
   Exit Sub
End If
' "Yπάρχουν "
ans = MsgBox("53-54" + mL_CapRes(53) + str(counter) + mL_CapRes(54), vbYesNo)
If ans = vbNo Then
   Exit Sub
End If


MDIForm1.Arxeia(10).Visible = False
MDIForm1.Syntages(20).Visible = False
MDIForm1.bohuhtika(30).Visible = False
MDIForm1.ejodos(40).Visible = False







Set MyDB1 = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
SQL = "UPDATE PROSPAU2 SET epil=0 WHERE STATUS<2;"
MyDB1.Execute SQL
MyDB1.Close





Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
Set p = mydb.OpenRecordset("PROSPAU2")
Set x = mydb.OpenRecordset("XIMITECH")
counter = 0
AR_S = 0
x.MoveFirst
Menu = 3
Do While Not x.eof
         If IsNull(x("addr_prot")) Then GoTo skip
         If Val(x("addr_prot")) = 0 Then GoTo skip
         If IsNull(x("last_dosom")) Or Not IsDate(x("last_dosom")) Then GoTo skip
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
     

Private Sub Πελάτες_Click(Index As Integer)
    Screen.MousePointer = 11
Menu = 1
    pelates.Show
End Sub

Private Sub Import_Click(Index As Integer)
  Text44.Show
End Sub





Private Sub import2_Click(Index As Integer)
'   Text2002.Show
End Sub

Private Sub MHDEN_JOBLIST_Click(Index As Integer)
Form1.Show
End Sub

Private Sub paramet_Click(Index As Integer)
   PARAMETROI.Show
End Sub

Private Sub Πάνε_Click(Index As Integer)
    Menu = 7 'flag για να καταλαβαίνει το πρόγραμμα συνταγών ποιός το
             ' κάλεσε
    Screen.MousePointer = 11
    WAIT.Show
End Sub

Private Sub Καλιμπράρισμα_Click(Index As Integer)
    Screen.MousePointer = 11
     Form1.Show
End Sub



Private Sub ejodos_Click(Index As Integer)
  
  Dim dbsjet As Database
  Dim r As Recordset
  
ZAP_MOVES
  
  
If system_ready = 0 Then End
If m_unloading = 0 Then
     response = MsgBox(mL_Res(70), 20, mL_Res(71))
     If response = 7 Then ' akyro
        Exit Sub
     End If
End If
  
  m_unloading = 1
  
  
  Write_Exo
 Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\coord_TB.mdb")
      Set r = dbsjet.OpenRecordset("mem3")

        r.MoveFirst
        r.EDIT
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

r.EDIT
  r("userOut1") = Now
r.update

r.Close


 RobSend ("1out.17-0:1out.18-0:1out.19-0:1out.20-0:1out.21-0:1out.22-0:1out.23-0:1out.24-0:1out.24-0:1out.25-0:1out.26-0:1out.27-0:1OUT.28-0:1OUT.29-0:1OUT.30.0:1OUT.31-0:1OUT.32-0:")
 '  temp% = SendATBlock(768, cmd$, 0)
ValveReset
   cmd$ = "!k:!MIDEN:"
   TEMP% = SendATBlock(768, cmd$, 0)

  RobSend ("!lh 3,3,3,3:")
  RobSend ("!ls 0,0,0,0:")
  
  
  
  
  
   
End
  Screen.MousePointer = 1
End Sub

Private Sub ferta_Click(Index As Integer)
    Screen.MousePointer = 11
   FERE.Show
End Sub

Private Sub loipes_Click(Index As Integer)
   EPANEKTYP.Show
End Sub

Private Sub make_dial_Click(Index As Integer)
    fereSK.Show
End Sub

 Private Sub MDIForm_Load()
'=======================================================
Dim dbsjet As Database
Dim r As Recordset, MHome, m_err, homedone


homedone = 1



On Error GoTo write_mess

If Command = "P" Then
  system_ready = 0
  Load_Parametroi
  MDIForm1.Arxeia(10).Enabled = False
  MDIForm1.Πάνε(22).Enabled = False
  MDIForm1.bohuhtika(30).Enabled = False
  Exit Sub
End If

system_ready = 1

980 Me.top = 0
981 Me.left = 0
982 Me.height = 7200
983 Me.width = 9900
984 vis_talos = 1


FindValves


 Connect_System

If system_ready = 0 And Command <> "P" Then
    MsgBox "Operating system not Loaded"
    End
End If



' c6k.SendFile ("c:\talos\george1.prg")

 MilSec 3000

RobSend ("TDIR")

zyg_Settings = Balance_Settings
  valve_mixed_on = "!" + Cold_Water_On + ":" + "!" + Hot_Water_On + ":"
  valve_mixed_off = "!" + Cold_Water_Off + ":" + "!" + Hot_Water_Off + ":"
'Combo1.AddItem "1%"
'Combo1.AddItem "5%"
'For i = 1 To 12
'Combo1.AddItem Trim(str(i * 10)) + "%"
'Next

MilSec 2000
'RobSend ("run setup")
MilSec 2000


RobSend ("!INFNC2-A:")

'RobSend ("!" + Motor_Output_on + ":")

RobSend ("!" + Motor_Output_on + ":")

RobSend ("DRIVE 1,1,1,1")


'RobSend ("! v 5,5,5,0.5")
'RobSend ("! a 5,5,5,0.5")
'RobSend ("! ad 5,5,5,0.5")


 RobSend ("hombac 1111")
RobSend ("HOMv 1,1,1,0.2")
RobSend ("HOMa 1,1,1,0.2")

'RobSend ("limlvl 110110110110")
'RobSend ("hom ,,0")

'RobSend ("HOM 1,1,,1")

    c6k.fsenabled = True                'enable fast status


 load_taxythtes












' RobSend ("!DRFLVL 1111:!errlvl0:!COMEXR1:!INFNC2-E:!INFEN1:")

#If Win32 Then
   'result% = SetDevice(0, 768)
#End If



2300



aritmos_ximikon = 0
'MsgBox "debug 1"
 'Check_passw

' STHN DIKIA MOY VERION NEA DEIXNEI AYTA POY THELO
'If Dir("D:\LAGEURO\MEM.DBF") <> "" Then
'    import2(24).Visible = True
'Else
'    import2(24).Visible = False
'End If


'to ekana rem giati ειχε λαθη νομιζε οτι δεν εκανα home
For k = 1 To 4
   If Not IsHomeDone(k) Then
   '  MsgBox mL_Res(-1000 + 476) + str(k)
     homedone = 0
   End If
Next
'If homedone = 0 Then End






If DemoVersion = 1 Then MsgBox "DEMO VERSION"

990 If App.PrevInstance Then End
'MsgBox "debug 2"
'On Error Resume Next   '12/99
991 Set mydb = Workspaces(0).OpenDatabase("c:\TALOS\WATCH.mdb")
'MyDB.Execute "SELECT MOVES.* into MOVEPROB from [MOVES];"
'MyDB.Execute "insert into MoveProb  select [MOVES].*  from [MOVES];"
992 mydb.Close
'MsgBox "debug 3"
993 ZAP_MOVES
'On Error GoTo 0'12/99



mUndelete_SYNTAGH = 0  ' δεν έχει κάνει έλεγχο για μισοσβησμένη συνταγή
Lim1 = 17    ' hard limits
Lim2 = 18  ' soft limits
'system_ready = 0
m_Cmd$ = "    "
On Error GoTo 12
m_err = 0
m_menu = 0
 m_unloading = 0
 ' ----------- m u l t i l i n g u a l -----------------
 MDIForm1.Caption = "" ' mL_Res(9) ' "ΚΕΝΤΡΙΚΗ ΦΟΡΜΑ"
 
999  Arxeia(10).Caption = mL_CapRes(10)   ' "Αρχεία"
 Χημικά(11).Caption = mL_CapRes(11)  '"Χημικά"
 Πελάτες(12).Caption = mL_CapRes(12)   '"Πελάτες"
 Διαδικασίες(13).Caption = mL_CapRes(13)  '"Διαδικασίες"
 pania(14).Caption = mL_CapRes(14) ' "Πανιά"
 Reindex(15).Caption = mL_CapRes(15)  '"Ανακατασκευή Αρχείων"
 Syntages(20).Caption = mL_CapRes(20)   '"Συνταγές"
 Εργασίες(21).Caption = mL_CapRes(21)  '"Εργασίες"
 Πάνε(22).Caption = mL_CapRes(22) ' "Λίστα Αναμονής"
 bohuhtika(30).Caption = mL_CapRes(30) ' "Βοηθητικά"
' Καλιμπράρισμα(31).caption = ml_CapRes(31) '"Καλιμπράρισμα"
 ferta(32).Caption = mL_CapRes(32) ' "Φέρε Μπουκάλια"
 phgaine(33).Caption = mL_CapRes(33) ' "Πήγαινε Μπουκάλια"
 make_dial(34).Caption = mL_CapRes(34) '"Κατασκευή Διαλυμάτων"
 ejodos(40).Caption = mL_CapRes(40)  '"Εξοδος"
 
 Mhden_joblist(16).Caption = mL_CapRes(23) '23 " Mηδενισμός αρχείου εργασιών"
 'check_kentra(36).caption = ml_CapRes(24) '24 " Ελεγχος κεντραρισμάτων"
 AXRHSIMOPOIHTA(37).Caption = mL_CapRes(25) '25 " Συνταγή με αχρησιμοποίητα μπουκάλια"
' Colors(36).caption = ml_CapRes(26) '26 " Χρώμα φόρμας"
 paramet(37).Caption = mL_CapRes(27) '27 "Παράμετροι"
 loipes(38).Caption = mL_CapRes(28) '28 "επανεκτύπωση
 Bohu(35).Caption = mL_CapRes(29) '29 "Βοηθητικές εργασίες"
'23 " Mηδενισμός αρχείου εργασιών"
'24 " Ελεγχος κεντραρισμάτων"
'25 " Συνταγή με αχρησιμοποίητα μπουκάλια"
'26 " Χρώμα φόρμας"
'27 "Παράμετροι"
'28 " Mηδενισμός αρχείου εργασιών"
'29 "Βοηθητικές εργασίες"
  
'Me.Caption = "debug 5"
  ' ----------- m u l t i l i n g u a l -----------------
  epil_Menu = 0 ' ΦΛΑΓΚ ΟΤΙ ΕΓΙΝΕ ΕΠΙΛΟΓΗ
    
    
    
DBEngine.IniPath = "C:\talos\talos.INI"
  
  'check_up dbase driver
'  m_err = "checkfil.dbf,c:\talos"
 ' Set MyDB = Workspaces(0).OpenDatabase("c:\TALOS", False, False, "dBASE IV;")
  'Set MySet = MyDB.OpenRecordset("CHECKFIL")
'  MySet.MoveFirst
 ' If left$(MySet("onoma"), 4) = "DELE" Then
  '   MsgBox mL_Res(50) ' "θα γίνει αναδιοργάνωση των αρχείων"
   '    m_err = "reindex.exe,c:\talos"
    ' Reindexing (0)
'  End If
 ' MyDB.Close
  
  
  
  
  
  
  
  
' find_colors
  On Error GoTo STRU_PARAM
ar_real_pothr = 24
MaxTimeWater = 30
m_err = "parametroi,c:\talos\recipies.mdb"


Load_Parametroi
'Me.Caption = "debug 6"
'Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\recipies.mdb")
'Set r = dbsjet.OpenRecordset("PARAMETROI")
'  allazo_xromata = r("allazo_xromata")
 ' allazo_fonts = r("allazo_fonts")
  'prozygismena = r("prozygismena")
'   dokimastikh_syntagh = r("dokimastikh_syntagh")
 ' AXRHSIMOPOIHTES_ORES = r("AXRHSIMOPOIHTES_ORES")
'  Default_baros_panioy = r("default_baros_panioy")
 ' Ar_tel_mpoykal = r("ar_tel_mpoykal")
  'Ar_1_mpoykal = r("ar_1_mpoykal")
'  DUM = r("MPANIO")
 ' If IsNull(r("ar_pothr")) Then
  '   ar_real_pothr = 24
'  Else
 '    ar_real_pothr = r("ar_pothr")
  'End If
'If IsNull(r("maxtimewater")) Then
 '  r.Edit
 '  r("maxtimewater") = MaxTimeWater
 '  r.update
'Else
  ' MaxTimeWater = r("maxtimewater")
'End If
  
 ' DUM = r("ColormatchDir")
  'DUM = r("Colormatchfilename")
  
   
   
   
'If IsNull(r("Fornext_small")) Then
'   r.Edit
 '  r("Fornext_small") = 40000
  ' r("Fornext_big") = 120000
   'r.update
'End If
 '
  ' Fornext_small = r("Fornext_small")
   'Fornext_big = r("Fornext_big")
  
  
'  If IsNull(r("doyleyo_alati")) Then
'     DOYLEYO_ALATI = 0
 'Else
'     DOYLEYO_ALATI = r("doyleyo_alati")
'  End If
'  UESH_ALATIOY = r("UESH_ALATIOY")
'
'  If IsNull(r("vis_talos")) Then
'     vis_talos = 0
' Else
'    vis_talos = r("VIS_TALOS")
'  End If
' DUM = r("METATOPISH_W")
 
 
 
 
' r.MoveNext


'If IsNull(r("maxtimewater")) Then
 '  Avance_Online_Water = 5000
'Else
'   Avance_Online_Water = r("maxtimewater")
'End If
 
  
  
'If IsNull(r("default_baros_panioy")) Then
'   YperXeilish = 300
'   r.Edit
'   r("default_baros_panioy") = 300
'   r.update
'Else
'   YperXeilish = r("default_baros_panioy")
'End If
  
  
  
  
  
  
  
' r.Close
       
       
Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\recipies.mdb")

Set r = dbsjet.OpenRecordset("prospau2")
DUM = r("DATE_DISP")
r.Close
       
Set r = dbsjet.OpenRecordset("jOBLIST")
'DUM = r("XHMIKA")
r.Close
       
       
       
       
       
       
On Error GoTo 12
m_err = "users,c:\talos\recipies.mdb"
       
 userinput.Show

1100 DUM = ADD_FIELD("PANIA", "FYRA", "SINGLE")


1101 If Command = "P" Then
1102      MDIForm1.Arxeia(10).Enabled = False
1103      MDIForm1.Πάνε(22).Enabled = False
1104      MDIForm1.bohuhtika(30).Enabled = False
1105      system_ready = 0
      Exit Sub
1107 End If

'Me.Caption = "debug 8"
1108 m_err = "IsOSLoaded(768) "
'1109 system_ready = IsOSLoaded(768)
'                               RobSend ("!setup:")
MilSec 2000
1110 m_err = "findvalves,c:\talos\COORD_TB.MDB"

1111




'RobSend ("d,,7500:go ,,1:d 9000:go 1:")

'RobSend ("d,,-1000:go ,,1:d 12000:go 1:")









'RobSend ("!" + Gripper_On + ":")

'tmp% = SendAT6400File(1, 768, 0, "c:\talos\talos1.prg", 0)
'RobSend ("!setup:)   '" + Gripper_On + ":")
If system_ready = 0 Then
    MsgBox mL_Res(51) '"δεν φορτώθηκε το λειτουργικό του robot"
    Exit Sub
End If
'  RobSend ("!" + Gripper_Off + ":")

'RobSend ("!d 1000:go 1:!" + Gripper_Off + ":")

1115  load_taxythtes


1116 m_err = "ISCUTOFF(0)"
    
    
    
    
    
    
    
    
  
    
1117   Do While isCutOff(0) = 1
        MsgBox mL_Res(72), , "Talos"  ' "Ελευθέρωσε το Emergency Stop  ' Define message.
       ' "Ελευθέρωσε το Emergency Stop  ", , "Talos"  ' Define message.
1118   Loop
 
 
1119 m_err = "mem3,c:\talos\coord_tb.mdb"
 
   
1120       Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\coord_tb.mdb")
1121      Set r = dbsjet.OpenRecordset("mem3")

1122     r.MoveFirst
      
1123      If r("XRHSTHS") = "JOYSTICK" Then
1124         MsgBox mL_Res(73), , "Talos"
         '  "ΠΡΟΣΟΧΗ !!!  .Επειδή χρησιμοποιήθηκε το Joystick είναι απαραίτητο να εκτελεστεί η διαδικασία Home. Να γίνει επανεκκίνηση του Υπολογιστή.", , "Talos"
         '   "ΠΡΟΣΟΧΗ !!!  .Επειδή χρησιμοποιήθηκε το Joystick είναι απαραίτητο να εκτελεστεί η διαδικασία Home. Να γίνει επανεκκίνηση του Υπολογιστή.", , "Talos"
1125           Write_Exo
1126           End
1127       End If
      
1128      If r("HOME") = "HOME=OK" Then
        '
1129       Else
1130             MsgBox mL_Res(74), , "Talos"
1131             Write_Exo
1132            ' End

1133      End If
      
      
      
1134      If r("XRHSTHS") = "BUSY" Then
      
1135           MsgBox mL_Res(75), , "Talos"
1136           Write_Exo
1137           ' End
1138       End If
      
1139     If vis_talos = 0 Then
1140        r.EDIT
1141        r("HOME") = "NO HOME"
1142       r("XRHSTHS") = "BUSY"
1143        r.update
1144     End If
1145       r.Close
       
1146             Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\watch.mdb")
1147             Set r = dbsjet.OpenRecordset("moves")
1148             r.AddNew
1149               r("movement") = left$("BEGIN SYNTAGES", 30)
1150               r("datetime") = Now
1151             r.update
1152             r.Close
       
       
       
       
       
       
1153 m_err = "var,c:\talos\coord_tb.mdb"
       
1154     Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\coord_tb.mdb")
1155     Set r = dbsjet.OpenRecordset("var")
1156            Do Until r.eof
1157              v = r("field_name")
1158              t = r("field_val")
1159              If IsNull(t) Then
1160                t = 0
1161              End If
1162              cmd$ = "!" + v + "=" + str$(t) + ":"
              
1163              TEMP% = SendATBlock(768, cmd$, 0)
1164             r.MoveNext
1165        Loop
1166      r.Close
On Error GoTo 0
       
       
    On Error Resume Next
       
       
' προγραμμα συνταγων
1167 Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\watch.mdb")
1168           Set r = dbsjet.OpenRecordset("commun")
1169           On Error Resume Next
1170           r.EDIT
1171           r("bges_exo") = 0
1172           r.update
1173           r.Close
           On Error GoTo 0
       
If vis_talos = 1 Then
    'MsgBox "moveNEW"
    'retVal = Shell("C:\talos\movenew.EXE", 3)     ' Run Calculator.
    
#If Win32 Then
'   retVal = Shell("C:\talos\mov6.EXE", 3)
#Else
'   retVal = Shell("C:\talos\movement.EXE", 3)
#End If


End If
       
       
'       "προσοχή δεν κάνω ελέγχους για HOME  ", , "Talos"
1180  If error(5, 1) = 1 Then
 
1181           MsgBox mL_Res(76), , "Talos"
   '        "Διακοπή λειτουργίας λόγω αδυναμίας εκτέλεσης Home  ", , "Talos"
   '        "Διακοπή λειτουργίας λόγω αδυναμίας εκτέλεσης Home  ", , "Talos"
1182            Write_Exo
            End
1183  End If
       
 '  cmd$ = "!out.1-0:!out.2-0:!out.3-1:!out.4-0:!out.5-0:!out.6-0:!out.7-0:!out.8-0:!out.9-0:!out.10-0:!out.13-0:!out.14-0:"
 '  temp% = SendATBlock(768, cmd$, 0)
       
  Exit Sub
  
write_mess:
  MsgBox "error in line " + str(Erl)


 MsgBox mess
  
12:
If Err.Number = 53 Or Err.Number = 3011 Then
      MsgBox "error in line " + str(Erl)
     '  "Εγινε προσπάθεια να διαβαστεί το αρχείο "
     MsgBox mL_Res(55) + m_err
     Write_Exo
     End
     End If
Resume

STRU_PARAM:
'Set MyDB = Workspaces(0).OpenDatabase("c:\TALOS\RECIPIES.MDB", False, False)
' MsgBox "error in line " + str(Erl)
Set mydb = OpenDatabase("c:\TALOS\RECIPIES.MDB")

Set r = mydb.OpenRecordset("PARAMETROI")
ok = 0

'  dum = r("ColormatchDir")
'  dum = r("ColormatchPath")

m1 = 0: m2 = 0: m3 = 0: M4 = 0: M5 = 0: M6 = 0: M7 = 0: M8 = 0: M9 = 0: M10 = 0: M11 = 0: M12 = 0: M13 = 0
For k = 0 To r.Fields.Count - 1
     If UCase(r.Fields(k).Name) = "FORNEXT_SMALL" Then
         m1 = 1
     End If
     If UCase(r.Fields(k).Name) = "FORNEXT_BIG" Then m2 = 1
     If UCase(r.Fields(k).Name) = "UESH_ALATIOY" Then m3 = 1
     If UCase(r.Fields(k).Name) = "VIS_TALOS" Then M4 = 1
     If UCase(r.Fields(k).Name) = "DEFAULT_BAROS_PANIOY" Then M5 = 1
     If UCase(r.Fields(k).Name) = "MPANIO" Then M6 = 1
     If UCase(r.Fields(k).Name) = "DOYLEYO_ALATI" Then M7 = 1
     If UCase(r.Fields(k).Name) = "METATOPISH_W" Then M8 = 1
     If UCase(r.Fields(k).Name) = "COLORMATCHDIR" Then M9 = 1
     If UCase(r.Fields(k).Name) = "COLORMATCHFILENAME" Then M10 = 1
     If UCase(r.Fields(k).Name) = "AR_POTHR" Then M11 = 1
     If UCase(r.Fields(k).Name) = "MPANIO" Then M12 = 1
     If UCase(r.Fields(k).Name) = "MAXTIMEWATER" Then M13 = 1
Next

r.Close

On Error Resume Next
If m1 = 0 Then dbsjet.Execute "ALTER TABLE PARAMETROI ADD  COLUMN FORNEXT_SMALL SINGLE"
If m2 = 0 Then dbsjet.Execute "ALTER TABLE PARAMETROI ADD  COLUMN FORNEXT_BIG SINGLE"
If m3 = 0 Then dbsjet.Execute "ALTER TABLE PARAMETROI ADD  COLUMN UESH_ALATIOY SINGLE"
If M4 = 0 Then dbsjet.Execute "ALTER TABLE PARAMETROI ADD  COLUMN VIS_TALOS SINGLE"
If M5 = 0 Then dbsjet.Execute "ALTER TABLE PARAMETROI ADD  COLUMN DEFAULT_BAROS_PANIOY SINGLE"
If M6 = 0 Then dbsjet.Execute "ALTER TABLE PARAMETROI ADD  COLUMN MPANIO SINGLE"
If M7 = 0 Then dbsjet.Execute "ALTER TABLE PARAMETROI ADD  COLUMN DOYLEYO_ALATI SINGLE"
If M8 = 0 Then dbsjet.Execute "ALTER TABLE PARAMETROI ADD  COLUMN METATOPISH_W SINGLE"
If M9 = 0 Then dbsjet.Execute "ALTER TABLE PARAMETROI ADD  COLUMN COLORMATCHDIR TEXT(50)"
If M10 = 0 Then dbsjet.Execute "ALTER TABLE PARAMETROI ADD COLUMN COLORMATCHFILENAME TEXT(50)"
If M11 = 0 Then dbsjet.Execute "ALTER TABLE PARAMETROI ADD COLUMN AR_POTHR SINGLE"
If M12 = 0 Then dbsjet.Execute "ALTER TABLE PARAMETROI ADD COLUMN MPANIO SINGLE"
If M13 = 0 Then dbsjet.Execute "ALTER TABLE PARAMETROI ADD COLUMN MAXTIMEWATER SINGLE"


DUM = ADD_FIELD("PARAMETROI", "TIME_PREHEAT", "SINGLE")
DUM = ADD_FIELD("PARAMETROI", "WEIGHT_BOTTLE_MIN", "SINGLE")
DUM = ADD_FIELD("PARAMETROI", "WEIGHT_BOTTLE_MAX", "SINGLE")



If Not IsEmpty(Xim1_on) Then
'     DUM = ADD_FIELD("PROSPAU2", "XIM1", "SINGLE")
 '    DUM = ADD_FIELD("PROSPAU2", "XIM2", "SINGLE")
  '   DUM = ADD_FIELD("PROSPAU2", "XIM3", "SINGLE")
   '  DUM = ADD_FIELD("PROSPAU2", "XIM4", "SINGLE")
    ' DUM = ADD_FIELD("PROSPAU2", "XIM5", "SINGLE")
End If

DUM = ADD_FIELD("PROSPAU2", "DATE_DISP", "DATETIME")
DUM = ADD_FIELD("PROSPAU3", "DATE_DISP", "DATETIME")
DUM = ADD_FIELD("DOKPROS2", "DATE_DISP", "DATETIME")

DUM = ADD_FIELD("JOBLIST", "XHMIKA", "SINGLE")
DUM = ADD_FIELD("JOBLIST3", "XHMIKA", "SINGLE")

'r("MPANIO")
Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\recipies.mdb")
Set r = dbsjet.OpenRecordset("PARAMETROI")
Resume

End Sub

Function ADD_FIELD(TABLE, field, TYPOS)
Dim dbsjet, r, m1, k

Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\recipies.mdb")
Set r = dbsjet.OpenRecordset(TABLE)
m1 = 0
For k = 0 To r.Fields.Count - 1
     If UCase(r.Fields(k).Name) = field Then m1 = 1
Next
r.Close
On Error GoTo 0 'Resume Next
m = "ALTER TABLE " + TABLE + " ADD  COLUMN " + field + " " + TYPOS
If m1 = 0 Then dbsjet.Execute m

End Function
Function ADD2_FIELD(DATAB, TABLE, field, TYPOS)
Dim dbsjet, r, m1, k, m

Set dbsjet = DBEngine.Workspaces(0).OpenDatabase(DATAB)
Set r = dbsjet.OpenRecordset(TABLE)
m1 = 0
For k = 0 To r.Fields.Count - 1
     If UCase(r.Fields(k).Name) = field Then m1 = 1
Next
r.Close
On Error GoTo 0 'Resume Next
m = "ALTER TABLE " + TABLE + " ADD  COLUMN " + field + " " + TYPOS
If m1 = 0 Then dbsjet.Execute m

End Function
Private Sub MDIForm_Unload(Cancel As Integer)




If m_unloading = 0 Then
   m_unloading = 1
   ejodos_Click (1)
   Write_Exo
End If
End Sub

Private Sub Check_passw_old()
Dim CODE_NEXT, L
Create_Kodik

'ΑΝΟΙΓΩ ΤΟ ΑΡΧΕΙΟ
Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\RECIPIES.mdb")
             
             Set r = dbsjet.OpenRecordset("RECIP")
             L = 0
        Do While Not r.eof
              L = L + 1
              If Not IsNull(r("k")) Then
                 If Abs(r("k") - 36313631) < 10 Then Exit Sub
              End If
              If IsNull(r("HME")) Then  'ΠΑΩ ΣΤΟ ΤΕΛΕΥΤΑΙΟ ΜΕ ΗΜΕΡΟΜΗΝΙΑ
                 CODE_NEXT = r("K") 'PIANO TON KODIKO TOY EPOMENOY RECORD
                 r.MovePrevious 'GYRIZO PISO GIA NA PIASO THN HMEROMHNIA THS TELEYTAIAS XRHSHS
                 Exit Do
              End If
               r.MoveNext
        Loop
 '       Form1.Caption = L
'        R.Close
        
          r.EDIT
          If IsNull(r("fores")) Then
             r("fores") = 1
          Else
             r("FORES") = r("FORES") + 1
          End If
          
          r.update
          
        
        If Month(Now) <> Month(r("HME")) Or r("fores") > 200 Then
           'ΖΗΤΑΕΙ ΚΩΔΙΚΟ
             w = InputBox("Give The Password " + str(L) + " please.(Communicate with SCLAVOS USA NC  tel 704-423-0171)")
             If CODE_NEXT = Val(w) / 2 Then 'ΚΩΔΙΚΟ ΠΟΥ ΠΛΗΚΤΡΟΛΟΓΗΣΕ Ο ΧΡΗΣΤΗΣ
                  r.MoveNext
                  r.EDIT
                  r("HME") = Now
                  r.update
             ElseIf w = 36313631 Then
                 r.MoveFirst
                 r.EDIT
                 r("k") = 36313631
                 r.update
             Else
                 MsgBox "Sorry , no good Password"
                 End
             End If
         Else
           'CONTINUE
         End If
                      
End Sub

Private Sub Check_passw()

Dim CODE_NEXT, L, fores, arx
epitr_syntages = 3500

2000 Create_Kodik

20010 fores = Get_ComDLG(1)

If fores < epitr_syntages Then Exit Sub


2030 arx = Year(Now) * 124 + Month(Now) * 244 + Day(Now) * 364 + Hour(Now) * 484 + Minute(Now) * 724 + Second(Now) * 964



'ΑΝΟΙΓΩ ΤΟ ΑΡΧΕΙΟ
'ΖΗΤΑΕΙ ΚΩΔΙΚΟ
w = InputBox("Give The Authorization code for number " + str(arx) + " please." + Chr(13) + "(Communicate with SCLAVOS USA NC  tel 704-423-0171)")
If Val(w) = Find_ANTICODE(arx) - Int(arx / 3) Then
    ch_comdlg -100000000

ElseIf Val(w) = Find_ANTICODE(arx) Then
    ch_comdlg -3800
Else
    If fores < epitr_syntages + 300 Then
        MsgBox "You need the Authorization Code. The program soon will stop."
        ch_comdlg 5
        Exit Sub
    Else
        MsgBox "Sorry. You need the Authorization Code."
        End
    End If
    
End If


End Sub

Sub Create_Kodik()

Dim dbsjet, r, errors, mydb, MySet
errors = 0
' GENERATE RECORDS


On Error GoTo createdbf
2500  Set mydb = Workspaces(0).OpenDatabase(GetWindowsDir() + "\SYSTEM", False, False, "dBASE IV;")
2501 Set r = mydb.OpenRecordset("COMDLG16.DLL")
On Error GoTo 0


              For k = 1 To 100
                 r.AddNew
                 If k < 10 Then
                    r("HME") = Now
                 End If
                     r("K") = 5 * k ^ 1 + 6 * k ^ 2 + 7 * k ^ 3 + 8 * k ^ 4
                     
                     If k = 1 Then
                         r("k") = 2500
                     End If
                     r.update
              Next
              
              
              
              r.Close
 
 
ch_comdlg 1
 
 
 
 Exit Sub






'comdlg16.dll

'Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\RECIPIES.mdb")
'On Error GoTo create
'             Set r = dbsjet.OpenRecordset("recip")
'On Error GoTo 0
'If errors = 0 Then Exit Sub
'             For k = 1 To 100
'                r.AddNew
'                If k < 10 Then
'                   r("HME") = Now
'                End If
'                r("K") = 5 * k ^ 1 + 6 * k ^ 2 + 7 * k ^ 3 + 8 * k ^ 4
'                r.update
'             Next
'           r.Close
'           Exit Sub



create:
MsgBox str(Erl)
dbsjet.Execute "CREATE TABLE recip (k long, fores single , hme datetime)"
Set r = dbsjet.OpenRecordset("recip")
errors = 1
Resume Next
Return



createdbf:

mydate = Date
mytime = Time

Date = #12/11/2096#
Time = #12:12:00 PM#

'If Not FileExists(GetWindowsDir() + "\SYSTEM\" + "RECIP.DBF") Then
   mydb.Execute "CREATE TABLE  comdlg16.dll (k long, fores single , hme datetime)"
'End If

Date = mydate
Time = mytime

'Name "recip.dbf" As "COMDLG.DLL"
Set r = mydb.OpenRecordset("comdlg16.dll")
errors = 1
Resume Next
Return





End Sub

Private Sub Mhded_joblist_Click(Index As Integer)
  Form1.Show
End Sub

Private Sub pania_Click(Index As Integer)
    Screen.MousePointer = 11
Menu = 3
    pelates.Show

End Sub

Private Sub phgaine_Click(Index As Integer)
     Screen.MousePointer = 11
    PANE.Show
End Sub

Private Sub Reindex_Click(Index As Integer)
 Screen.MousePointer = 11
 Reindexing (0)
 
 Screen.MousePointer = 0
  MsgBox mL_Res(35)   ' "Η διαδικασία ολοκληρώθηκε"

End Sub

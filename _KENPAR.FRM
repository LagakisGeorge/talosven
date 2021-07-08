VERSION 4.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   ClientHeight    =   3810
   ClientLeft      =   2280
   ClientTop       =   1965
   ClientWidth     =   5880
   Height          =   4500
   Icon            =   "_KENPAR.frx":0000
   Left            =   2220
   LinkTopic       =   "MDIForm1"
   Picture         =   "_KENPAR.frx":0442
   Top             =   1335
   Width           =   6000
   WindowState     =   2  'Maximized
   Begin VB.Menu Arxeia 
      Caption         =   "Αρχεία"
      Enabled         =   0   'False
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
         Enabled         =   0   'False
         Index           =   22
      End
   End
   Begin VB.Menu bohuhtika 
      Caption         =   "Βοηθητικά"
      Enabled         =   0   'False
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
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim m_unloading As Integer, m_menu

Sub Write_Exo()
    Set dbsJet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\watch.mdb")
           Set r = dbsJet.OpenRecordset("commun")
           On Error Resume Next
           r.Edit
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
     
     PELATES.Show
     'PROCEDURE.Show
End Sub
Private Sub Εργασίες_Click(Index As Integer)
     
     Screen.MousePointer = 11
     mJOBLIST.Show
End Sub

Private Sub AXRHSIMOPOIHTA_Click(Index As Integer)

'-----------------
Dim MyDB, p, X, DT, ORES, counter
Dim AR_S As Integer




Set MyDB = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)

Set X = MyDB.OpenRecordset("XIMITECH")

DT = AXRHSIMOPOIHTES_ORES * 3600
counter = 0
AR_S = 0
X.MoveFirst
Menu = 3
Do While Not X.EOF
         If IsNull(X("addr_prot")) Then GoTo skip1
         If Val(X("addr_prot")) = 0 Then GoTo skip1
         If IsNull(X("last_dosom")) Or Not IsDate(X("last_dosom")) Then GoTo skip1
         If DateDiff("S", CDate(X("Last_Dosom")), Now) < DT Then GoTo skip1
          counter = counter + 1
skip1:
    X.MoveNext
Loop

MyDB.Close

If counter = 0 Then
   MsgBox mL_Res(52) ' "Δεν υπάρχουν χημικοτεχνικά που πρέπει να σερβιριστούν."
   Exit Sub
End If
' "Yπάρχουν "
Ans = MsgBox("53-54" + mL_CapRes(53) + Str(counter) + mL_CapRes(54), vbYesNo)
If Ans = vbNo Then
   Exit Sub
End If


MDIForm1.arxeia(10).Visible = False
MDIForm1.Syntages(20).Visible = False
MDIForm1.Bohuhtika(30).Visible = False
MDIForm1.ejodos(40).Visible = False







Set MyDB1 = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
sql = "UPDATE PROSPAU2 SET epil=0 WHERE STATUS<2;"
MyDB1.Execute sql
MyDB1.Close





Set MyDB = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
Set p = MyDB.OpenRecordset("PROSPAU2")
Set X = MyDB.OpenRecordset("XIMITECH")
counter = 0
AR_S = 0
X.MoveFirst
Menu = 3
Do While Not X.EOF
         If IsNull(X("addr_prot")) Then GoTo skip
         If Val(X("addr_prot")) = 0 Then GoTo skip
         If IsNull(X("last_dosom")) Or Not IsDate(X("last_dosom")) Then GoTo skip
         If DateDiff("S", CDate(X("Last_Dosom")), Now) < DT Then GoTo skip
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
              p("KOD") = X("KOD")
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
    X.MoveNext
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
    PELATES.Show
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
  
  Dim dbsJet As Database
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
 Set dbsJet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\coord_TB.mdb")
      Set r = dbsJet.OpenRecordset("mem3")

        r.MoveFirst
        r.Edit
        r("XRHSTHS") = "NO BUSY"
        r("HOME") = "HOME=OK"
       r.update
       r.Close
       
    Set dbsJet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\watch.mdb")
    Set r = dbsJet.OpenRecordset("moves")
             r.AddNew
               r("movement") = left$("END SYNTAGES", 30)
               r("datetime") = Now
             r.update
             r.Close
              
       

 Set dbsJet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\recipies.mdb")
      Set r = dbsJet.OpenRecordset("userinout")

r.MoveLast

r.Edit
  r("userOut1") = Now
r.update

r.Close


 '  cmd$ = "!out.1-0:!out.2-0:!out.3-1:!out.4-0:!out.5-0:!out.6-0:!out.7-0:!out.8-0:!out.9-0:!out.10-0:!out.13-0:!out.14-0:"
 '  temp% = SendATBlock(768, cmd$, 0)
ValveReset
   CMD$ = "!k:!MIDEN:"
   temp% = SendATBlock(768, CMD$, 0)

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
   
   Dim dbsJet As Database
   Dim r As Recordset, MHome, m_err


#If Win32 Then
   Result% = SetDevice(0, 768)
#End If

If DemoVersion = 1 Then MsgBox "DEMO VERSION"

If App.PrevInstance Then End

'On Error Resume Next   '12/99
Set MyDB = Workspaces(0).OpenDatabase("c:\TALOS\WATCH.mdb")
'MyDB.Execute "SELECT MOVES.* into MOVEPROB from [MOVES];"
'MyDB.Execute "insert into MoveProb  select [MOVES].*  from [MOVES];"
MyDB.Close
ZAP_MOVES
'On Error GoTo 0'12/99



mUndelete_SYNTAGH = 0  ' δεν έχει κάνει έλεγχο για μισοσβησμένη συνταγή
Lim1 = 17    ' hard limits
Lim2 = 18  ' soft limits
system_ready = 0
m_Cmd$ = "    "
On Error GoTo 12
m_err = 0
m_menu = 0
 m_unloading = 0
 ' ----------- m u l t i l i n g u a l -----------------
 MDIForm1.Caption = "" ' mL_Res(9) ' "ΚΕΝΤΡΙΚΗ ΦΟΡΜΑ"
 
 arxeia(10).Caption = mL_CapRes(10)   ' "Αρχεία"
 Χημικά(11).Caption = mL_CapRes(11)  '"Χημικά"
 Πελάτες(12).Caption = mL_CapRes(12)   '"Πελάτες"
 Διαδικασίες(13).Caption = mL_CapRes(13)  '"Διαδικασίες"
 pania(14).Caption = mL_CapRes(14) ' "Πανιά"
 Reindex(15).Caption = mL_CapRes(15)  '"Ανακατασκευή Αρχείων"
 Syntages(20).Caption = mL_CapRes(20)   '"Συνταγές"
 Εργασίες(21).Caption = mL_CapRes(21)  '"Εργασίες"
 Πάνε(22).Caption = mL_CapRes(22) ' "Λίστα Αναμονής"
 Bohuhtika(30).Caption = mL_CapRes(30) ' "Βοηθητικά"
' Καλιμπράρισμα(31).caption = ml_CapRes(31) '"Καλιμπράρισμα"
 ferta(32).Caption = mL_CapRes(32) ' "Φέρε Μπουκάλια"
 phgaine(33).Caption = mL_CapRes(33) ' "Πήγαινε Μπουκάλια"
 make_dial(34).Caption = mL_CapRes(34) '"Κατασκευή Διαλυμάτων"
 ejodos(40).Caption = mL_CapRes(40)  '"Εξοδος"
 
 MHDEN_JOBLIST(16).Caption = mL_CapRes(23) '23 " Mηδενισμός αρχείου εργασιών"
 'check_kentra(36).caption = ml_CapRes(24) '24 " Ελεγχος κεντραρισμάτων"
 AXRHSIMOPOIHTA(37).Caption = mL_CapRes(25) '25 " Συνταγή με αχρησιμοποίητα μπουκάλια"
' Colors(36).caption = ml_CapRes(26) '26 " Χρώμα φόρμας"
 paramet(37).Caption = mL_CapRes(27) '27 "Παράμετροι"
 loipes(38).Caption = mL_CapRes(28) '28 "επανεκτύπωση
 bohu(35).Caption = mL_CapRes(29) '29 "Βοηθητικές εργασίες"
'23 " Mηδενισμός αρχείου εργασιών"
'24 " Ελεγχος κεντραρισμάτων"
'25 " Συνταγή με αχρησιμοποίητα μπουκάλια"
'26 " Χρώμα φόρμας"
'27 "Παράμετροι"
'28 " Mηδενισμός αρχείου εργασιών"
'29 "Βοηθητικές εργασίες"
  
  
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
  

m_err = "parametroi,c:\talos\recipies.mdb"
Set dbsJet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\recipies.mdb")
Set r = dbsJet.OpenRecordset("PARAMETROI")
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
  
  If IsNull(r("vis_talos")) Then
     vis_talos = 0
 Else
    vis_talos = r("VIS_TALOS")
  End If
  
  
  
 r.Close
       
m_err = "users,c:\talos\recipies.mdb"
       
 userinput.Show


m_err = "IsOSLoaded(768) "
system_ready = 0 ' IsOSLoaded(768)

m_err = "findvalves,c:\talos\COORD_TB.MDB"

FindValves

'RobSend ("!" + Gripper_On + ":")

'tmp% = SendAT6400File(1, 768, 0, "c:\talos\talos1.prg", 0)
'RobSend ("!setup:)   '" + Gripper_On + ":")
If system_ready = 0 Then
    MsgBox mL_Res(51) '"δεν φορτώθηκε το λειτουργικό του robot"
    Exit Sub
End If
'  RobSend ("!" + Gripper_Off + ":")

'RobSend ("!d 1000:go 1:!" + Gripper_Off + ":")



m_err = "ISCUTOFF(0)"
    
    
    
    
    
    
    
    
  
    
   Do While isCutOff(0) = 1
        MsgBox mL_Res(72), , "Talos"  ' "Ελευθέρωσε το Emergency Stop  ' Define message.
       ' "Ελευθέρωσε το Emergency Stop  ", , "Talos"  ' Define message.
   Loop
 
 
m_err = "mem3,c:\talos\coord_tb.mdb"
 
   
       Set dbsJet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\coord_tb.mdb")
      Set r = dbsJet.OpenRecordset("mem3")

     r.MoveFirst
      
      If r("XRHSTHS") = "JOYSTICK" Then
         MsgBox mL_Res(73), , "Talos"
         '  "ΠΡΟΣΟΧΗ !!!  .Επειδή χρησιμοποιήθηκε το Joystick είναι απαραίτητο να εκτελεστεί η διαδικασία Home. Να γίνει επανεκκίνηση του Υπολογιστή.", , "Talos"
         '   "ΠΡΟΣΟΧΗ !!!  .Επειδή χρησιμοποιήθηκε το Joystick είναι απαραίτητο να εκτελεστεί η διαδικασία Home. Να γίνει επανεκκίνηση του Υπολογιστή.", , "Talos"
           Write_Exo
           End
       End If
      
      If r("HOME") = "HOME=OK" Then
        '
      Else
             MsgBox mL_Res(74), , "Talos"
             Write_Exo
            End
      End If
      
      
      
      If r("XRHSTHS") = "BUSY" Then
      
           MsgBox mL_Res(75), , "Talos"
           Write_Exo
           End
       End If
      
     If vis_talos = 0 Then
        r.Edit
        r("HOME") = "NO HOME"
        r("XRHSTHS") = "BUSY"
        r.update
     End If
       r.Close
       
             Set dbsJet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\watch.mdb")
             Set r = dbsJet.OpenRecordset("moves")
             r.AddNew
               r("movement") = left$("BEGIN SYNTAGES", 30)
               r("datetime") = Now
             r.update
             r.Close
       
       
       
       
       
       
m_err = "var,c:\talos\coord_tb.mdb"
       
     Set dbsJet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\coord_tb.mdb")
     Set r = dbsJet.OpenRecordset("var")
            Do Until r.EOF
              V = r("field_name")
              t = r("field_val")
              If IsNull(t) Then
                t = 0
              End If
              CMD$ = "!" + V + "=" + Str$(t) + ":"
              
              temp% = SendATBlock(768, CMD$, 0)
             r.MoveNext
        Loop
      r.Close
On Error GoTo 0
       
       
       
       
       
' προγραμμα συνταγων
Set dbsJet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\watch.mdb")
           Set r = dbsJet.OpenRecordset("commun")
           On Error Resume Next
           r.Edit
           r("bges_exo") = 0
           r.update
           r.Close
           On Error GoTo 0
       
If vis_talos = 1 Then
 RetVal = Shell("C:\talos\movement.EXE", 3)     ' Run Calculator.
End If
       
       
'       "προσοχή δεν κάνω ελέγχους για HOME  ", , "Talos"
 If error(5, 1) = 1 Then
 
           MsgBox mL_Res(76), , "Talos"
   '        "Διακοπή λειτουργίας λόγω αδυναμίας εκτέλεσης Home  ", , "Talos"
   '        "Διακοπή λειτουργίας λόγω αδυναμίας εκτέλεσης Home  ", , "Talos"
            Write_Exo
            End
  End If
       
 '  cmd$ = "!out.1-0:!out.2-0:!out.3-1:!out.4-0:!out.5-0:!out.6-0:!out.7-0:!out.8-0:!out.9-0:!out.10-0:!out.13-0:!out.14-0:"
 '  temp% = SendATBlock(768, cmd$, 0)
       
  Exit Sub
12:
If Err.Number = 53 Or Err.Number = 3011 Then
     '  "Εγινε προσπάθεια να διαβαστεί το αρχείο "
     MsgBox mL_Res(55) + m_err
     Write_Exo
     End
     End If
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)




If m_unloading = 0 Then
   m_unloading = 1
   ejodos_Click (1)
   Write_Exo
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

Private Sub Reindex_Click(Index As Integer)
 Screen.MousePointer = 11
 Reindexing (0)
 
 Screen.MousePointer = 0
  MsgBox mL_Res(35)   ' "Η διαδικασία ολοκληρώθηκε"

End Sub

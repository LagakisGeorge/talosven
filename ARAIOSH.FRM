VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form araiosh 
   Caption         =   "Form1"
   ClientHeight    =   6744
   ClientLeft      =   -636
   ClientTop       =   528
   ClientWidth     =   8976
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6744
   ScaleWidth      =   8976
   Begin VB.CommandButton Command2 
      Caption         =   "exit"
      Height          =   500
      Left            =   4800
      TabIndex        =   12
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Εκκίνηση παρ.διαλυμάτων"
      Height          =   500
      Left            =   1560
      TabIndex        =   9
      Top             =   3480
      Width           =   1380
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   7200
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "ARAIOSH.frx":0000
      Height          =   1350
      Left            =   210
      OleObjectBlob   =   "ARAIOSH.frx":0014
      TabIndex        =   8
      Top             =   4410
      Width           =   6360
   End
   Begin VB.Label Label10 
      Caption         =   "Παρασκευή διαλυμάτων με αραίωση"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label18 
      Caption         =   "Label18"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      DataField       =   "buk"
      DataSource      =   "Data1"
      Height          =   210
      Left            =   2520
      TabIndex        =   10
      Top             =   4080
      Width           =   1125
   End
   Begin VB.Label Label8 
      Caption         =   "Μεταφορά ποτηριού από ζυγαριά στην θέση του"
      Height          =   300
      Index           =   0
      Left            =   1575
      TabIndex        =   7
      Top             =   2985
      Width           =   4185
   End
   Begin VB.Label Label7 
      Caption         =   "Κεφαλή σε θέση παρκαρίσματος"
      Height          =   300
      Index           =   0
      Left            =   1560
      TabIndex        =   6
      Top             =   2565
      Width           =   4185
   End
   Begin VB.Label Label3 
      Caption         =   "Δοσομέτρηση χρώματος 50γρ"
      Height          =   300
      Index           =   0
      Left            =   1575
      TabIndex        =   5
      Top             =   1440
      Width           =   4185
   End
   Begin VB.Label Label4 
      Caption         =   "Επιστροφή πυκνού διαλύματος στην θέση του"
      Height          =   300
      Index           =   0
      Left            =   1575
      TabIndex        =   4
      Top             =   1800
      Width           =   4185
   End
   Begin VB.Label Label5 
      Caption         =   "Φέρε κεφαλή νερού πάνω από την ζυγαριά"
      Height          =   300
      Index           =   0
      Left            =   1575
      TabIndex        =   3
      Top             =   2160
      Width           =   4185
   End
   Begin VB.Label Label6 
      Caption         =   "Δοσομέτρησε νερό 500γρ"
      Height          =   300
      Index           =   0
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Φέρε το αντίστοιχο πυκνό διάλυμα"
      Height          =   300
      Index           =   0
      Left            =   3840
      TabIndex        =   1
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Μεταφορά από θέση ποτηριού στην ζυγαριά"
      Height          =   300
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   600
      Width           =   4185
   End
End
Attribute VB_Name = "araiosh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim x, APOBARO
Sub Find_Axes()

   x0 = axe(1)
   y0 = axe(2)
   z0 = axe(3)
   w0 = axe(4)
End Sub
Private Sub dosometrhsh(gr)
  Dim cmd$, DUM
   'kafalh 1 pano apo zygo
        cmd$ = "!var12=90:"
        DUM = wait_to_understand(cmd$)
        DUM = Wait_Robot(0)
  
  
 ' Syn_Dos.synarthsh.Caption = "ARAIOSH"
  'synarthsh.Caption = "ARAIOSH"
  Syn_Dos.asked_q = gr '  ZHTOYMENH ΠΟΣΟTHTA ΣE MGRS
  Syn_Dos.Level = x("ypol_prot")      'ΓΡΑΜΑΡΙΑ ΜΠΟΥΚΑΛΙΟΥ
  Syn_Dos.Valve_On = valve_1_on   ' ΑΝΟΙΓΜΑ ΒΑΛΒΙΔΑΣ
  Syn_Dos.Valve_Off = valve_1_off ' ΚΛΕΙΣΙΜΟ ΒΑΛΒΙΔΑΣ
  Syn_Dos.Caption = x("perigrafh") 'περιγραφη χρώματος που δοσομετρείται
  Syn_Dos.Show 1
               
 ' If Val(Label8.Caption) = -5 Or Label8.Caption = "0" Then
  '          m_Counter = m_Counter + 1
               '   If m_Counter < 30 Then GoTo 111
  'End If
            
  Do: Loop Until Label18.Caption <> ""
End Sub

Private Sub Command1_Click()
Dim r, Apait_sk, periek, db, Ap_dial, Apait_sk2, DUM
Dim Z, Z1, COUNTZ, PR_SK, mz, nero, mpoykali, b, dbsjet

Command1.Enabled = False
'Set r = Data1.Recordset 'entoles
'proto mh paraskeyasuen


 Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\skon_tb.mdb")
   Set r = dbsjet.OpenRecordset("entoles")



r.MoveFirst
Do While Not r.eof
  If IsNull(r("nero_act")) Then
     Exit Do
  Else
    If IsNull(r("nero_act")) Then
      Exit Do
    Else
      r.MoveNext
    End If
  End If
Loop

'GoTo 999



'υπολογισμός δοσομέτρησης
'buk = μπουκάλι που έχει τα ποσοστά

Do While Not r.eof
  Apait_sk = r("periek") * r("nerostox") 'mgr
   mpoykali = r("buk")
  b = r.Bookmark
  Set db = OpenDatabase("c:\talos\recipies.mdb")
  Set x = db.OpenRecordset("select *from ximitech where addr_prot=" + str(r("buk")))

  If IsNull(x("real_cons")) Then
     periek = Val(x("morfh"))
  Else
     If x("real_cons") = 0 Then
        periek = Val(x("morfh"))
     Else
        periek = x("real_cons")
     End If
  End If

  Ap_dial = 100 * Apait_sk / periek

  If x("ypol_prot") * 1000 - 10000 < Ap_dial Then
     Ap_dial = x("ypol_prot") * 1000 - 10000
     Apait_sk2 = Ap_dial / 1000 * periek / 100
     r.EDIT
     r("nerostox") = Abs(Int(r("nerostox") * Apait_sk2 / Apait_sk))
     r.update
     Apait_sk = Apait_sk2
  End If
 
 On Error Resume Next
' ΑΝΟΙΓΩ ΣΕΙΡΙΑΚΗ ΠΟΡΤΑ ΚΑΙ ΜΗΔΕΝΙΖΩ ΖΥΓΑΡΙΑ
 MSComm1.commport = Balance_Port
 MSComm1.Settings = Balance_Settings
 MSComm1.InputLen = 0

MSComm1.PortOpen = False
MSComm1.PortOpen = True
On Error GoTo 0
  

MSComm1.Output = "Z " + Chr$(13) + Chr(10)


mz = Zygis0(0)
Do While Abs(mz) > 10
    MSComm1.Output = "Z " + Chr$(13) + Chr(10)
     MilSec 200
     mz = Zygis0(0)
Loop
Label1(0).BackColor = Red
  se_zygo r("pot") 'φερε το μπουκάλι στην ζυγαριά    'kafalh 1 pano apo zygo
  mz = Zygis0(0)
  If mz < 200 Then
      MsgBox mLResnew(906, "Ο ζυγός είναι άδειος.", 1) '  The scale is empty.
      Exit Sub
  End If
  Label1(0).BackColor = Green

Label2(0).BackColor = Red
  mpoyk_kef r("buk") 'fere pothri 1 pano apo zygo
  nero = r("nerostox") - Ap_dial
  If nero < 0 Or nero > 550000 Then
     MsgBox mLResnew(911, "error on water calculation", 1)
     Exit Sub
  End If
  Label6(0).BackColor = Red
  
  DUM = NER_DOSOMETRHSH(nero, Porta_Ner)
  
  DUM = Wait_Robot(0)
  Label2(0).BackColor = Green
  Label6(0).BackColor = Green

'  mpoyk_kef 'debug   gia sbisimo
 ' Timer1.Enabled = False

On Error Resume Next
 ' MSComm1.PortOpen = False
  'MilSec 200
  'MSComm1.PortOpen = False
On Error GoTo 0
  
  
  Label3(0).BackColor = Red

  dosometrhsh (Ap_dial) 'δοσομέτρηση χρώματος
  
  x.EDIT
  x("ypol_prot") = x("ypol_prot") - Val(Label18.Caption) / 1000
  x.update
On Error Resume Next
' ΑΝΟΙΓΩ ΣΕΙΡΙΑΚΗ ΠΟΡΤΑ ΚΑΙ ΜΗΔΕΝΙΖΩ ΖΥΓΑΡΙΑ
 MSComm1.commport = Balance_Port
 MSComm1.Settings = Balance_Settings
 MSComm1.InputLen = 0

MSComm1.PortOpen = False

  dosometrhsh (Ap_dial) 'δοσομέτρηση χρώματος


MSComm1.PortOpen = True
On Error GoTo 0
  
COUNTZ = 0
117
  Z = Zygis0(0)
  MilSec 1000
  Z1 = Zygis0(0)
  COUNTZ = COUNTZ + 1
  If Abs(Z1 - Z) > 10 And COUNTZ < 10 Then GoTo 117
      
 'ΜΕ ΒΓΑΖΕΙ R("BUK")=0
  
  PR_SK = (Label18.Caption / Ap_dial) * Apait_sk
 ' Data1.Recordset.Bookmark = b
  'Data1.Recordset.Edit
  'Data1.Recordset("COLOR_ACT") = PR_SK
  'Data1.Recordset("NERO_ACT") = Z - PR_SK
  'Data1.Recordset("buk") = mpoykali
 ' Data1.Recordset.update
 If PR_SK > 0 Then
   r.EDIT
   r("COLOR_ACT") = PR_SK
   r("nero_ACT") = Z - PR_SK
   r.update
 End If
' dosometrhsh (Ap_dial) 'δοσομέτρηση χρώματος
  
  
  
'  Data1.Recordset.Edit
   
  ' Data1.Recordset.update
   
   
  Label4(0).BackColor = Red
 Label3(0).BackColor = Green
 
 
 'dosometrhsh (Ap_dial) 'δοσομέτρηση χρώματος  debug
 
 
 
  pykno_piso_ r("buk")
  
  Label5(0).BackColor = Green

  Label6(0).BackColor = Red

  kef_park
  Label6(0).BackColor = Green
  
  Label7(0).BackColor = Red

  pot_piso_ r("pot")
  Label7(0).BackColor = Green

  r.MoveNext
  Label1(0).BackColor = Grey
  Label2(0).BackColor = Grey
  Label3(0).BackColor = Grey
  Label4(0).BackColor = Grey
  Label5(0).BackColor = Grey
  Label6(0).BackColor = Grey
  Label7(0).BackColor = Grey
  Label8(0).BackColor = Grey
 
Loop

MsgBox "Η παρασκευή των διαλυμάτων τελείωσε."
End Sub

Private Sub se_zygo(g)
 Dim cmd$, DUM, k
 'fere pothri 1 pano apo zygo
           k = 3000 + g
           cmd$ = "!var12=" + str(k) + ":"
           DUM = wait_to_understand(cmd$)
           DUM = Wait_Robot(0)
End Sub


Private Sub mpoyk_kef(g)
Dim k, cmd$, DUM
'fere mpoykala se kefali
       k = 1000 + g ' Data1.Recordset("buk")
       cmd$ = "!var12=" + str(k) + ":"
       DUM = wait_to_understand(cmd$)
       
End Sub



Private Sub pykno_piso_(g)
Dim k, cmd$, DUM
'epestrece mpoykali 61 sthn uesh toy
       k = 2000 + g 'Data1.Recordset("buk")
       cmd$ = "!var12=" + str(k) + ":"
       DUM = wait_to_understand(cmd$)
       DUM = Wait_Robot(0)
End Sub


Private Sub kefner_sezygo_()
  
         ' πηγαινε το gripper μπροστά και τον 4 άξονα στο νερό
         RobSend ("!d,,," + str(var67_3kef) + ":")
         
         Find_Axes
         w0 = var67_3kef
         RobSend_w ("!go1,,,1:")
    
End Sub

Function Zygis0(x As Single)
  Dim counter, starter, dummy, FromModem$, Y
    ' Flush the input buffer.
       MSComm1.Output = "SI " + Chr$(13) + Chr$(10)
'       MSComm1.Output = "UPD 20 " + Chr$(13) + Chr$(10)
       MSComm1.InBufferCount = 0
    starter = GetCurrentTime()
    counter = 0
    ' Wait for "OK" to come back from the modem.
    Do
       dummy = DoEvents()
          If GetCurrentTime() - starter > 2000 Then
            Zygis0 = -999000999
            Exit Function
      End If
       ' If there is data in the buffer, then read it.
      ' MilSec 10
       If MSComm1.InBufferCount Then
          FromModem$ = FromModem$ + MSComm1.Input
          ' Check for "OK".
          'Me.Caption = FromModem$
          If InStr(FromModem$, Chr$(13)) Then
                FromModem$ = LTrim(Right$(FromModem$, 20))
                If (InStr(FromModem$, ".") Or InStr(6, FromModem$, ".")) And InStr(6, FromModem$, "g") > 10 Then
                       Exit Do
                End If
          End If
          
          
       Else
            counter = counter + 1
            If counter = 500 Then  ' 500
                 counter = 0
                 MSComm1.Output = "SI " + Chr$(13) + Chr(10)
                 MSComm1.InBufferCount = 0
                 FromModem$ = ""
             End If
        End If
    Loop
      If InStr(FromModem$, "UL") > 0 Then
       Zygis0 = -9900000
       Exit Function
   End If
   
    If InStr(FromModem$, "OL") > 0 Then
       Zygis0 = -7900000
       Exit Function
   End If
    
    
 
      If InStr(FromModem$, ".") < 5 Then
         x = InStr(6, FromModem$, ".")
         Y = InStr(x, FromModem$, "g")
      Else
         x = InStr(FromModem$, ".")
         Y = InStr(FromModem$, "g")
      End If
      Zygis0 = Val(Mid$(FromModem$, x - 6, 10)) * 1000
      Me.Caption = Zygis0
     ' Label3(0).Caption = Zygis0
End Function


'
Function NER_DOSOMETRHSH(ByVal gr As Single, ByVal porta As String)

'πιάνω το 95% με την μία
Dim ssdxx As Single, Y As Double, Z As Double, w As Double, fores
Dim x As Single, tmax As Double, tmin As Double, t As Double, texcel As Double, stoxos As Double
Dim xronos, mxronos, mz, APOBARO
Dim Poses_Zyg, starter, DEYTEROLEPTA, posa_loop, anoikth
Dim arx_xron, ll, ii$, zyg0, cmd$, TEMP%, abs_zyg
Dim mgr_sec, a0, t_a0, mS, dddb, vim, r0, dbsjet, var67, k
  DEYTEROLEPTA = 20
  posa_loop = 0
  texcel = 0
 
        
        
 
 
 
 
 
 
 
 
 
 On Error Resume Next


' ΑΝΟΙΓΩ ΣΕΙΡΙΑΚΗ ΠΟΡΤΑ ΚΑΙ ΜΗΔΕΝΙΖΩ ΖΥΓΑΡΙΑ
' MSComm1.CommPort = Balance_Port
' MSComm1.Settings = Balance_Settings
' MSComm1.InputLen = 0

'MSComm1.PortOpen = False
'MSComm1.PortOpen = True
On Error GoTo 0
MSComm1.Output = "Z " + Chr$(13) + Chr(10)


mz = Zygis0(0)
Do While Abs(mz) > 10
    MSComm1.Output = "Z " + Chr$(13) + Chr(10)
     MilSec 200
     mz = Zygis0(0)
Loop

APOBARO = mz
Label9.Caption = APOBARO


 'MSComm1.Output = "Z " + Chr$(13) + Chr(10)

'Qpragm = 799


 
EPIU = gr

'Label7.Caption = EPIU

arx_xron = GetCurrentTime()
If EPIU < 5000 Then     '300
  a = 0
Else

   ll = 0
   ii$ = ""
  Zyg = 0
  zyg0 = Zygis0(0) - APOBARO
  
  
  starter = GetCurrentTime()
  cmd$ = "!" + Cold_Water_On + ":"
  TEMP% = SendATBlock(768, cmd$, 0)
 
  
 
 abs_zyg = 0
 anoikth = 1
 Do While True
          abs_zyg = Zygis0(1500)
          ' δεν απαντά γρήγορα η ζυγαριά
          Do While abs_zyg < -999000000
                cmd$ = "!" + Cold_Water_Off + ":"
                TEMP% = SendATBlock(768, cmd$, 0)
                abs_zyg = Zygis0(3000)
                anoikth = 0
                If GetCurrentTime() - starter > 30000 Then
                         cmd$ = "!" + Cold_Water_Off + ":"
                         TEMP% = SendATBlock(768, cmd$, 0)
                         MilSec 3000
                         NER_DOSOMETRHSH = 0
                         Exit Function
                End If
          Loop
         
          Zyg = abs_zyg - APOBARO
          'Label3(0).Caption = zyg
      Poses_Zyg = Poses_Zyg + 1
      If Val(EPIU) - Zyg < 30000 Then               '  3500       2500=kalo
          Exit Do
       End If
       If Poses_Zyg Mod 500 = 0 Then
           If Zyg - zyg0 < 5 Then
                  Exit Do
           Else
                 zyg0 = Zyg
            End If
       End If
       DoEvents
       If anoikth = 0 Then
          anoikth = 1
          cmd$ = "!" + Cold_Water_On + ":"
          TEMP% = SendATBlock(768, cmd$, 0)
       End If
       
       If GetCurrentTime() - starter > 180000 Then  '3MIN
                         cmd$ = "!" + Cold_Water_Off + ":"
                         TEMP% = SendATBlock(768, cmd$, 0)
                         MilSec 3000
                         NER_DOSOMETRHSH = 0
                         Exit Function
       End If
 Loop
     cmd$ = "!" + Cold_Water_Off + ":"
     TEMP% = SendATBlock(768, cmd$, 0)
     mxronos = GetCurrentTime() - starter

     'Label8.Caption = GetCurrentTime() - starter

      ' τελευταία Ζύγιση
      Zyg = Zygis0(0)
      'Label3(0).Caption = zyg
      a = Zyg
End If


MilSec 500

a = Zygis0(0) - APOBARO

If mxronos = 0 Then
   mgr_sec = 10000
   Else
   mgr_sec = 1000 * a / mxronos
End If

If mgr_sec > 10000 Then mgr_sec = 10000
If mgr_sec < 5000 Then mgr_sec = 5000

'mgr_sec = 0 ' IIf(mxronos = 0, 0, A / mxronos)
starter = GetCurrentTime()
Do While EPIU - a > 1900
     a0 = a
     t_a0 = GetCurrentTime()
     
     If mgr_sec = 0 Then
        a = 0
     Else
         '
         
       If EPIU - a < 4000 Then
             a = metrhs3_ner(0.9 * (EPIU - a) / mgr_sec / 4, 1) ' /2 mono ston griba
       Else
            a = metrhs3_ner((EPIU - a) / mgr_sec, 1) ' /2 mono ston griba
       End If
     End If
     
     mS = GetCurrentTime()
     
   ' τελευταία Ζύγιση
     MilSec 2500
     Zyg = Zygis0(0) - APOBARO   '  Zyg_TRIPLO(40, 40)  '    gis1(0)

     a = Zyg
    ' υπολογίζω τα mgr/sec
     If (mS - t_a0) > 0 And a - a0 > 0 Then
          mgr_sec = 1000 * (a - a0) / (mS - t_a0)
          If mgr_sec > 10000 Then mgr_sec = 10000
          If mgr_sec < 5000 Then mgr_sec = 5000
     End If
    
    
    If GetCurrentTime() - starter > 240000 Then
        NER_DOSOMETRHSH = -999000999
        cmd$ = "!" + Cold_Water_Off + ":"
        TEMP% = SendATBlock(768, cmd$, 0)
        Exit Do
    End If


Loop

If Abs(EPIU - a) > 10000 Then
     NER_DOSOMETRHSH = a '-999000999
Else
    NER_DOSOMETRHSH = a
End If

End Function
Function metrhs3_ner(xr As Double, rixies As Integer) As Single
'**********************************************
Dim dr As Long, cmd$, TEMP%
xr = IIf(xr <= 0, 0, xr)

 cmd$ = "!" + Cold_Water_On + ":"
 TEMP% = SendATBlock(768, cmd$, 0)
     If xr < 500 Then xr = 700
     If xr > 5000 Then xr = 5000
    MilSec Int(xr)
  
  cmd$ = "!" + Cold_Water_Off + ":"
  TEMP% = SendATBlock(768, cmd$, 0)
         
   ' zyg = Zygis0(0)
    metrhs3_ner = Zyg
End Function



Private Sub kef_park()
Dim dbsjet, r, k, VAR61, DUM, L
 Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\coord_tb.mdb")
     Set r = dbsjet.OpenRecordset("var")
     Do Until r.eof
              If r("field_name") = "VAR61" Then
                 k = Val(Mid$(r("field_name"), 4, 2))
                 If k = 61 Then
                    VAR61 = r("field_val")
                    Exit Do
                 End If
              End If
              r.MoveNext
     Loop
     r.Close


 ' πηγαινε ΣΤΗΝ ΘΕΣΗ ΠΑΡΚΑΡΙΣΜΑΤΟΣ
         RobSend ("!d,,," + str(VAR61) + ":")
         Find_Axes
         'Find_Axes
         w0 = VAR61
         RobSend_w ("!go,,,1:")
         DUM = Wait_Robot(0)

End Sub


Private Sub pot_piso_(g)
 'fere pothri 1 piso
 Dim k, cmd$, DUM
 k = 4000 + g 'Data1.Recordset("pot")
           cmd$ = "!var12=" + str(k) + ":"
           DUM = wait_to_understand(cmd$)
           DUM = Wait_Robot(0)
End Sub






Private Sub Command2_Click()
  Unload Me
End Sub



Public Sub kane_klik()
Dim xronos, mxronos, mz, APOBARO, texcel
Dim Poses_Zyg, starter, DEYTEROLEPTA, posa_loop, anoikth
Dim arx_xron, ll, ii$, zyg0, cmd$, TEMP%, abs_zyg
Dim mgr_sec, a0, t_a0, mS, dddb, vim, r0, dbsjet, var67, k
  DEYTEROLEPTA = 20
  posa_loop = 0
  texcel = 0
 
 Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\coord_tb.mdb")
     Set r0 = dbsjet.OpenRecordset("var")
     Do Until r0.eof
              If r0("field_name") = "VAR67" Then
                 k = Val(Mid$(r0("field_name"), 4, 2))
                 If k = 67 Then
                    var67 = r0("field_val")
                    Exit Do
                 End If
              End If
              r0.MoveNext
     Loop
     r0.Close

 
 
 
 
 
 Set dddb = Workspaces(0).OpenDatabase("c:\talos\recipies.MDB")
     Set r0 = dddb.OpenRecordset("parametroi")
      vim = r0("metatopish_W")
      vim = Val(var67) + vim
      r0.Close
        
        
        
        
     'upm4 να γινει 45
   ' Vim = Val(var67) + upm4 * Val(Water_Bypass)      'Μετατόπιση 4ου άξονα 20-30  χιλ δεξιά για να μην στάζει
    cmd$ = "!d,,," + str(vim) + ":!go,,,1:"
    RobSend (cmd$)
      
 MilSec 1000
 
 
 
 
End Sub

Function farb_DOSOMETRHSH(gr As Single, porta As String, ByVal ximiko)

'πιάνω το 95% με την μία
Dim ssdxx As Single, Y As Double, Z As Double, w As Double, fores
Dim x As Single, tmax As Double, tmin As Double, t As Double, texcel As Double, stoxos As Double
Dim mz, a00, Akrivos
 

End Function

VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMSYNT2 
   BackColor       =   &H00808000&
   ClientHeight    =   8490
   ClientLeft      =   855
   ClientTop       =   165
   ClientWidth     =   11040
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8490
   ScaleWidth      =   11040
   WindowState     =   2  'Maximized
   Begin VB.CheckBox ximikon 
      Caption         =   "Χημικό"
      Height          =   255
      Left            =   6240
      TabIndex        =   43
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdΡύθμισηΑυτόματης 
      Caption         =   "Ρύθμιση αυτόματης επιλογής"
      Height          =   240
      Left            =   7320
      TabIndex        =   42
      Top             =   5400
      Width           =   2655
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Βισκόζη"
      Height          =   270
      Left            =   5010
      TabIndex        =   41
      Top             =   5310
      Width           =   1110
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Χωρίς αλάτι"
      Height          =   270
      Left            =   3750
      TabIndex        =   40
      Top             =   5310
      Width           =   1110
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Auto solution selection"
      Height          =   255
      Left            =   7320
      TabIndex        =   39
      Top             =   5115
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Πραγματικές συντ."
      Height          =   255
      Left            =   7320
      TabIndex        =   37
      Top             =   4800
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   3000
      TabIndex        =   33
      Text            =   "0"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Maskedbox2 
      Height          =   285
      Left            =   4680
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox maskedbox1 
      Height          =   285
      Left            =   3840
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   435
      Left            =   3570
      Max             =   -100
      Min             =   100
      TabIndex        =   30
      Top             =   4740
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   435
      Left            =   2730
      TabIndex        =   29
      Top             =   4740
      Width           =   825
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   7
      Left            =   7476
      TabIndex        =   28
      Top             =   528
      Width           =   756
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   6
      Left            =   6648
      TabIndex        =   27
      Top             =   528
      Width           =   756
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   5
      Left            =   5808
      TabIndex        =   26
      Top             =   540
      Width           =   756
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   4
      Left            =   4980
      TabIndex        =   25
      Top             =   540
      Width           =   756
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   3
      Left            =   4164
      TabIndex        =   24
      Top             =   552
      Width           =   756
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   2
      Left            =   3336
      TabIndex        =   23
      Top             =   552
      Width           =   756
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   1
      Left            =   2520
      TabIndex        =   22
      Top             =   552
      Width           =   756
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   320
      Index           =   0
      Left            =   1680
      TabIndex        =   21
      Top             =   540
      Width           =   756
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00000000&
      Caption         =   "Εξοδος"
      Height          =   420
      Left            =   6240
      TabIndex        =   20
      Top             =   4755
      Width           =   930
   End
   Begin VB.Data JOBLIST 
      Caption         =   "JOBLIST"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\recipies.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7365
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "JOBLIST"
      Top             =   6090
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Data PROSXHM 
      Caption         =   "prosxhm"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\recipies.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   405
      Left            =   4485
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"_synmak9-sos.frx":0000
      Top             =   6465
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Nέα Συνταγή"
      Height          =   405
      Left            =   4095
      TabIndex        =   5
      Top             =   4755
      Width           =   930
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\recipies.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   390
      Left            =   4455
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "XIMITECH"
      Top             =   6015
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ενημέρωση Συνταγής"
      Enabled         =   0   'False
      Height          =   420
      Left            =   5130
      TabIndex        =   4
      Top             =   4755
      Width           =   1005
   End
   Begin VB.Data prospau2 
      Caption         =   "prospau2"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\recipies.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   435
      Left            =   2535
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "PROSPAU2"
      Top             =   6465
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\recipies.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2130
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select  DISTINCT KATASKEYAS,perigrafh  from ximitech  WHERE  (SKONH<>1 and addr_prot>0) ORDER BY PERIGRAFH;"
      Top             =   6060
      Visible         =   0   'False
      Width           =   2265
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1275
      Top             =   6345
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "_synmak9-sos.frx":0097
      Height          =   315
      Left            =   1590
      TabIndex        =   1
      Top             =   1185
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      BackColor       =   16777215
      ForeColor       =   0
      ListField       =   "perigrafh"
      BoundColumn     =   "perigrafh"
      Text            =   "DBCombo1"
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   7435
      _Version        =   393216
      Rows            =   30
      Cols            =   20
      FixedCols       =   0
      ScrollBars      =   0
   End
   Begin VB.Label Label13 
      Height          =   240
      Left            =   135
      TabIndex        =   38
      Top             =   5325
      Width           =   3500
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Mπάνιο"
      Height          =   240
      Left            =   150
      TabIndex        =   36
      Top             =   4725
      Width           =   990
   End
   Begin VB.Label Label11 
      Height          =   240
      Left            =   150
      TabIndex        =   35
      Top             =   5025
      Width           =   990
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "ξηρό αλάτι gr/Lit"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1155
      TabIndex        =   34
      Top             =   5205
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFF00&
      Height          =   405
      Left            =   120
      TabIndex        =   32
      Top             =   5640
      Width           =   9855
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Μεταβολή % "
      Height          =   375
      Left            =   1305
      TabIndex        =   31
      Top             =   4845
      Width           =   1380
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Label1"
      DataField       =   "ENTOLH"
      DataSource      =   "JOBLIST"
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   225
      Width           =   600
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Label1"
      DataField       =   "KOD_PEL"
      DataSource      =   "JOBLIST"
      Height          =   255
      Left            =   630
      TabIndex        =   18
      Top             =   225
      Width           =   2520
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Label1"
      DataField       =   "APOXRVSH"
      DataSource      =   "JOBLIST"
      Height          =   255
      Left            =   3375
      TabIndex        =   17
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      Caption         =   "Label1"
      DataField       =   "EIDOS_BAF"
      DataSource      =   "JOBLIST"
      Height          =   255
      Left            =   4455
      TabIndex        =   16
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      Caption         =   "Label1"
      DataField       =   "KOD_ERG"
      DataSource      =   "JOBLIST"
      Height          =   255
      Left            =   5550
      TabIndex        =   15
      Top             =   240
      Width           =   840
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF00&
      Caption         =   "Label1"
      DataField       =   "EIDOS_PANI"
      DataSource      =   "JOBLIST"
      Height          =   255
      Left            =   6615
      TabIndex        =   14
      Top             =   240
      Width           =   1470
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF00&
      Caption         =   "Label1"
      DataField       =   "SXESH_MPAN"
      DataSource      =   "JOBLIST"
      Height          =   255
      Left            =   8160
      TabIndex        =   13
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Εντολή"
      Height          =   240
      Left            =   15
      TabIndex        =   12
      Top             =   0
      Width           =   1005
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Πελάτης"
      Height          =   195
      Left            =   1095
      TabIndex        =   11
      Top             =   0
      Width           =   960
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Απόχρωση"
      Height          =   195
      Left            =   3390
      TabIndex        =   10
      Top             =   15
      Width           =   960
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Βαφή"
      Height          =   195
      Left            =   4470
      TabIndex        =   9
      Top             =   15
      Width           =   960
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Α.Παρ/λιας"
      Height          =   195
      Left            =   5550
      TabIndex        =   8
      Top             =   15
      Width           =   960
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Πανί"
      Height          =   240
      Left            =   6630
      TabIndex        =   7
      Top             =   15
      Width           =   960
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Ημερ.Παραδ."
      Height          =   240
      Left            =   8115
      TabIndex        =   6
      Top             =   15
      Width           =   1005
   End
End
Attribute VB_Name = "FRMSYNT2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim f_barosXimikoy As Long

Dim m_ROW, m_COL
Dim CAN_UPDATE_GRID
Const BLE = &HFF0000
Const ASPRO = &HFFFFFF
Const MAYRO = &H0&
Const kokkino = &HFF&
Const gkri = &HC0C0C0
Const max_Row = 30   ' μέγιστος αριθμός σειρών Grid

Dim old_cb
Dim old_1b
Dim OLD_2B

Dim old_cf
Dim old_1f
Dim OLD_2F
Dim YCOS ' Υψος σειρών του Grid
Dim seira
Dim per_xim(1 To 30) ' περιγραφή του χημικού
Dim KOD_XHM(1 To 30) 'ΚΩΔΙΚΌΣ ΧΗΜΙΚΟΥ
'Dim m_Entolh ' εντολή στην οποία θα κατασκευάσω προσπάθειες
Dim m_aa ' α/α προσπάθειας
Dim M_AA2
Dim proth_fora
Dim Max_Row_Updated

Dim work_focus
Dim Key_Pressed
Const gri = &HC0C0C0
Function Check_Mpanio()

Dim m_nero, mC, k, mAPAIT_POSOT, Aneparkes, Last_Seira
Dim mcount, Xamhla_Mgrs_Rows(1 To 30), Xamhla
Dim dNero, nero_ARAIOY, nero_PYKNOY, mARAIO
Dim mmm_row, mmm_col

'On Error Resume Next



mmm_row = Grid1.Row
mmm_col = Grid1.Col

For k = 1 To max_Row
  Xamhla_Mgrs_Rows(k) = 0
Next


Data1.Recordset.Index = "PER"
mcount = 0
 Dim mNERO As Long
 
    'Ιδια συνταγή
 mNERO = 0
 mAPAIT_POSOT = 0
 mC = 0
 Grid1.Col = 0
 Last_Seira = 0
 Label9.Caption = ""
 Xamhla = 0

                                       
   Label13.BackColor = FRMSYNT2.BackColor
   Label13.Caption = " "




For k = 1 To max_Row
     Grid1.Row = k
     Grid1.Col = 0
      
      If IsNull(Grid1.text) Then
         Exit For
      End If
      
      If Len(LTrim(Trim(Grid1.text))) = 0 Then
          Exit For
      End If
      
              Last_Seira = Last_Seira + 1
                           
              KOD_XHM(Last_Seira) = bres_ximiko(Grid1.text, 1) 'πυκνότερο
              
              If KOD_XHM(Last_Seira) < -900 Then
                  MsgBox "300-301." + mL_Res(300) + Grid1.text + mL_Res(301)
                    '300  "Το χρώμα "
                    '301  " δεν έχει συγκέντρωση. "
                  Exit For
              End If
              
              
              Data1.Recordset.Index = "kod"
              Data1.Recordset.Seek "=", KOD_XHM(Last_Seira)
              
              
               Grid1.Col = 1
               mEK = IIf(IsNull(Grid1.text), 0, Val(Grid1.text))
               Grid1.Col = 2
               mGL = IIf(IsNull(Grid1.text), 0, Val(Grid1.text))
    
               If IsNull(Data1.Recordset("real_cons")) Then
                        Pososto = val2(left$(Data1.Recordset("MORFH"), 3))
               ElseIf Data1.Recordset("real_cons") = 0 Then
                        Pososto = val2(left$(Data1.Recordset("MORFH"), 3))
               Else
                        Pososto = Data1.Recordset("real_cons")
               End If
                  
               
                  
                  
               If Pososto > 0 Then
                         If mEK > 0 Then
                                  
                                  mAPAIT_POSOT = mEK * JOBLIST.Recordset("BAROS_PANI") / Pososto
                             
                                  If mAPAIT_POSOT < 200 Then
                                       Label13.BackColor = kokkino
                                       Label13.Caption = Trim(Data1.Recordset("Perigrafh")) + " " + Format(mAPAIT_POSOT, "#####") + "mgr"
                                  End If
                                  
                                 
                              mNERO = mNERO + mAPAIT_POSOT - mEK * JOBLIST.Recordset("BAROS_PANI") / 100
                         Else
                              mAPAIT_POSOT = mGL * JOBLIST.Recordset("BAROS_PANI") * JOBLIST.Recordset("SXESH_MPAN") / (Pososto * 10)
                              mNERO = mNERO + mAPAIT_POSOT - (mGL) * JOBLIST.Recordset("SXESH_MPAN") / 1000
                         End If
               End If
               
               
               If mAPAIT_POSOT < 500 Then
                   Xamhla = Xamhla + 1
                   Xamhla_Mgrs_Rows(Xamhla) = k 'μαρκάρω τις συνταγές που έχουν λίγα mgrs
               End If
               
               Aneparkes = 0
                   If mAPAIT_POSOT / 1000 > Data1.Recordset("ypol_prot") Then
                      Label9.Caption = Label9.Caption + Data1.Recordset("perigrafh") + "->" + mL_CapRes(871) '"ανεπαρκές υπόλοιπο  "
                      Aneparkes = 1
                   End If
                   
                   
                   If DateDiff("d", Now, Data1.Recordset("LHJ_PROT")) < 0 Then
                      Label9.Caption = Label9.Caption + "  " + IIf(Aneparkes = 1, " &  " + mL_CapRes(872), Data1.Recordset("perigrafh") + "->" + mL_CapRes(872)) 'ληγμένο
                   End If
    
    
Next

If IsNull(JOBLIST.Recordset("xhmika")) Then
   xhmika = 0
Else
   xhmika = JOBLIST.Recordset("xhmika")
End If


Ner_dos = JOBLIST.Recordset("sxesh_mpan") * JOBLIST.Recordset("baros_pani") - mNERO - JOBLIST.Recordset("baros_pani") * JOBLIST.Recordset("apor") - xhmika * 1000


Label11.Caption = Format(IIf(JOBLIST.Recordset("baros_pani") = 0, 0, (xhmika * 1000 + mNERO) / JOBLIST.Recordset("baros_pani")), "##0.#")
If Ner_dos < 0 Then
       yperbash = IIf(JOBLIST.Recordset("baros_pani") = 0, 0, (mNERO + JOBLIST.Recordset("baros_pani") * JOBLIST.Recordset("apor")) / JOBLIST.Recordset("baros_pani"))
       MsgBox mL_Res(302) + " (" + Format(yperbash, "###.##") + ")"
       ' 302 "Υπέρβαση Μπάνιου στην συνταγή "
       Check_Mpanio = -1
Else
     For k = 1 To Xamhla
        'αν υπάρχει αντίστοιχο αραιότερο ,δώσε μου τον κωδικό του ή -1 αν όχι
              Grid1.Row = Xamhla_Mgrs_Rows(k) 'μαρκάρω τις συνταγές που έχουν λίγα mgrs
              Grid1.Col = 0
              mARAIO = bres_ximiko(Grid1.text, 0)
            If KOD_XHM(Grid1.Row) = mARAIO Or mARAIO = -9999 Then 'αραιότερο
                ' δεν υπάρχει αραιότερο
            Else
                'if υπαρχει δώσε μου την διαφορά μπάνιου (αραιου - πυκνού )
                'διαφορά + υπάρχον νερό > μπάνιου=ναι
                       'κρατάω τον παλιό κωδικό
                'οχι
                        ' παίρνω τον κωδικό του αραιότερου
                  
                nero_ARAIOY = Periexon_Nero(mARAIO, Grid1.Row)
                nero_PYKNOY = Periexon_Nero(KOD_XHM(Grid1.Row), Grid1.Row)
                dNero = nero_ARAIOY - nero_PYKNOY
                 If dNero > Ner_dos Then
                      ' δημιουργείται υπέρβαση μπάνιου
                      ' κρατάω το πυκνό
                 Else
                     Ner_dos = Ner_dos - dNero
                     KOD_XHM(Grid1.Row) = mARAIO
                 End If
           End If
    Next k
   Check_Mpanio = 1
   
End If
Grid1.Row = mmm_row
 Grid1.Col = mmm_col
End Function

Function axrhsto_Eyresh_Pyknoteroy(tel_seira)
  Dim mC, mR, k, mT, mCI, mEK, mGL, mMORFH, upd, ff
   mR = Grid1.Row
   mC = Grid1.Col
   upd = 0
   M_R = IIf(Max_Row_Updated = 0, 1, Max_Row_Updated)
   For k = tel_seira To M_R Step -1
       Grid1.Row = k: Grid1.Col = 0
       mT = Grid1.text
       Data1.Recordset.Seek "=", mT
       mCI = Data1.Recordset("KATASKEYAS")
       
       'ff =
       'S = IIf(InStr(ff, "%") = 0, 6, InStr(ff, "%")
       mMORFH = val2(Data1.Recordset("morfh"))
       Grid1.Col = 1
       mEK = Val(Grid1.text)
       Grid1.Col = 2
       mGL = Val(Grid1.text)
       
       'ψάχνω για πυκνότερο
       If mGL + mEK > 0 Then
            Data1.Recordset.MoveFirst
            Do While Not Data1.Recordset.eof
                 If Data1.Recordset("kataskeyas") = mCI Then
                     If Data1.Recordset("skonh") = 0 And val2(Data1.Recordset("morfh")) > mMORFH Then
                         Grid1.Col = 0
                         Grid1.text = Data1.Recordset("perigrafh")
                         upd = 1
                         MsgBox "303-304." + mL_CapRes(303) + mT + Chr(13) + mL_CapRes(304) + Grid1.text
                         '303 "Θα αντικατασταθεί το "
                         '304 " με το "
                      End If
                  End If
                  Data1.Recordset.MoveNext
             Loop
       End If
   Next
   
   If upd = 1 Then
      Eyresh_Pyknoteroy = 1
   Else
      Eyresh_Pyknoteroy = 0
   End If
End Function
Function bres_ximiko(mper, pykn)
'δίνω το όνομα του χημικού και μου δίνει το πυκνότερο/αραιότερο  σε περίπτωση που έχω 2 συγκεντρώσεις
' pykn=1 δίνω το πυκνότερο ,  =0 δίνω το αραιότερο
  Dim mC, mR, k, mT, mCI, mEK, mGL, mMORFH, upd, ff, mTot
       
               Grid1.Col = 1
               mEK = IIf(IsNull(Grid1.text), 0, Val(Grid1.text))
               Grid1.Col = 2
               mGL = IIf(IsNull(Grid1.text), 0, Val(Grid1.text))
        
        mTot = mEK + mGL
       
       
       
       Data1.Recordset.Index = "per"
       Data1.Recordset.Seek "=", mper
       
       mCI = Data1.Recordset("KATASKEYAS")
       bres_ximiko = Data1.Recordset("Kod")
       
       
       If Data1.Recordset("skonh") = 0 Then
           mMORFH = val2(Data1.Recordset("morfh"))
       Else
           mMORFH = 0
       End If
       
       
        
        
            'ψάχνω για πυκνότερο διάλυμα
            Do While Not Data1.Recordset.eof And UCase(Data1.Recordset("perigrafh")) = UCase(mper)
                 
              If Data1.Recordset("kataskeyas") = mCI Then
                     
                     
                  If pykn = 1 Then
                     If Data1.Recordset("addr_prot") > 0 And Data1.Recordset("skonh") = 0 And val2(Data1.Recordset("morfh")) > mMORFH Then
                        bres_ximiko = Data1.Recordset("kod")
                        mMORFH = val2(Data1.Recordset("morfh"))
                       End If
                   Else
                       If Data1.Recordset("addr_prot") > 0 And Data1.Recordset("skonh") = 0 And val2(Data1.Recordset("morfh")) < mMORFH Then
                          bres_ximiko = Data1.Recordset("kod")
                          Exit Function
                       End If
                   End If
                End If
                 
              Data1.Recordset.MoveNext
              If Data1.Recordset.eof Then
                     Exit Do
              End If
          Loop
   
   If mMORFH = 0 Then
      bres_ximiko = -9999
   End If

End Function

Function bres2_ximiko(mper, pykn)
'δίνω το όνομα του χημικού και μου δίνει το πυκνότερο/αραιότερο  σε περίπτωση που έχω 2 συγκεντρώσεις
' pykn=1 δίνω το πυκνότερο ,  =0 δίνω το αραιότερο



' εδω διαβαζει την παραμετρο ελαχιστης ποσότητας για επιλογη πυκνου/αραιου
Dim mydb As Database
Dim work As Workspace
Dim r As Recordset
Set work = Workspaces(0)
Set mydb = work.OpenDatabase("c:\TALOS\RECIPIES.MDB", False, False)
Set r = mydb.OpenRecordset("parametroi")
r.MoveNext
Dim n0 As Long
n0 = IIf(IsNull(r!Time_Preheat), 0, r!Time_Preheat)
r.Close





Dim mC, mR, k, mT, mCI, mEK, mGL, mMORFH, upd, ff, mTot
Dim m1, m2, m3, per1
               Grid1.Col = 1
               mEK = IIf(IsNull(Grid1.text), 0, Val(Grid1.text))
               Grid1.Col = 2
               mGL = IIf(IsNull(Grid1.text), 0, Val(Grid1.text))
        
        mTot = mEK + mGL
       
       
       
       
       Data1.Recordset.Index = "per"
       Data1.Recordset.Seek "=", mper
       mCI = Data1.Recordset("KATASKEYAS")
       bres2_ximiko = Data1.Recordset("Kod")
       
       If Data1.Recordset("timh") > 0 Then
               bres2_ximiko = Data1.Recordset("perigrafh")
             Exit Function
       End If
       
       
Dim SQL As String
Dim e As Recordset

       
Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
SQL = "select ximitech.addr_prot,ximitech.real_cons,ximitech.kataskeyas,ximitech.perigrafh from ximitech where kataskeyas='" + mCI + "' and real_cons>0 and addr_prot>0  order by real_cons;"
Set e = mydb.OpenRecordset(SQL, dbOpenDynaset)


If e.RecordCount = 0 Then
   SQL = "select ximitech.addr_prot,ximitech.real_cons,ximitech.kataskeyas,ximitech.perigrafh from ximitech where kataskeyas='" + mCI + "' and skonh=0 and addr_prot>0  order by real_cons;"
   Set e = mydb.OpenRecordset(SQL, dbOpenDynaset)
   bres2_ximiko = e("perigrafh")
   Exit Function
End If


e.MoveFirst


Dim mesos As Single




m1 = e("real_cons")
per1 = e("perigrafh")

Do While Not e.eof
    e.MoveNext
    
    If IsNull(e("real_cons")) Then e.EDIT: e("real_cons") = val2(e("morfh")): e.update
    
    If e.eof Then Exit Do
    
    If mTot > e("real_cons") Then
       m1 = e("real_cons")
       per1 = e("perigrafh")
    Else
       mesos = (m1 + e("real_cons")) / 2
       If (mTot / m1) > JOBLIST.Recordset("SXESH_MPAN") Then 'LIQUOR RATIO
            m1 = e("real_cons") 'PAIRNO TO PYKNO
            per1 = e("perigrafh")
       ElseIf (mTot / e("REAL_CONS")) * JOBLIST.Recordset("BAROS_PANI") > n0 Then  ' des parapano
            m1 = e("real_cons") 'PAIRNO TO PYKNO
            per1 = e("perigrafh")
       Else
         If mTot > mesos Then
            m1 = e("real_cons") 'PAIRNO TO PYKNO
            per1 = e("perigrafh")
         Else
            ' κρατω τα παλια  ARAIO
         End If
       End If
       Exit Do
    End If
Loop

 bres2_ximiko = per1
' Data1.Recordset.Seek "=", ff
Exit Function


       
       Data1.Recordset.Index = "per"
       Data1.Recordset.Seek "=", mper
       
       mCI = Data1.Recordset("KATASKEYAS")
       bres2_ximiko = Data1.Recordset("Kod")
       
       ff = Data1.Recordset("perigrafh")
       If Data1.Recordset("skonh") = 0 Then
           mMORFH = val2(Data1.Recordset("morfh"))
       Else
           mMORFH = 0
       End If
       
       
       Data1.Recordset.MoveFirst
       Do While Not Data1.Recordset.eof
             If Data1.Recordset("kataskeyas") = mCI Then
                 If Data1.Recordset("Kod") <> bres2_ximiko And Data1.Recordset("addr_prot") > 0 And Data1.Recordset("skonh") = 0 Then
                        ' exo kai allo
                        mesos = (val2(Data1.Recordset("morfh")) + mMORFH) / 2
                        If mTot >= mesos Then
                             If mMORFH > val2(Data1.Recordset("morfh")) Then
                                'take the old one
                                Exit Do
                             Else
                                 bres2_ximiko = Data1.Recordset("Kod") ' the new one
                                 ff = Data1.Recordset("perigrafh")
                                 Exit Do
                             End If
                        Else
                             If mMORFH < val2(Data1.Recordset("morfh")) Then
                                'take the old one
                                Exit Do
                             Else
                                 bres2_ximiko = Data1.Recordset("Kod") ' the new one
                                 ff = Data1.Recordset("perigrafh")
                                 Exit Do
                             End If
                        End If
                        
                        
                 End If
              End If
              Data1.Recordset.MoveNext
        Loop
        
   
      bres2_ximiko = ff
     Data1.Recordset.Seek "=", ff

End Function


Function CdBln(x)
   
   a = InStr(x, ",")
   If a = 0 Then
        CdBln = x
   Else
       CdBln = left(x, a - 1) + "." + Mid(x, a + 1, Len(x) - a)
    End If
 
End Function

Function FIND_LEFT1()
   FIND_LEFT1 = 20 + Grid1.left + Grid1.ColWidth(0)
End Function

Function Periexon_Nero(kod_Xroma, m_Seira As Integer)
Dim arx_R, arx_C, mEK, mGL, mNERO
   arx_R = Grid1.Row
   arx_C = Grid1.Col
   
   Grid1.Row = m_Seira
   mNERO = 0
           
           Data1.Recordset.Index = "kod"
           Data1.Recordset.Seek "=", kod_Xroma
              
              
               Grid1.Col = 1
               mEK = IIf(IsNull(Grid1.text), 0, Val(Grid1.text))
               Grid1.Col = 2
               mGL = IIf(IsNull(Grid1.text), 0, Val(Grid1.text))
    
               If IsNull(Data1.Recordset("real_cons")) Then
                        Pososto = Val(left$(Data1.Recordset("MORFH"), 3))
               ElseIf Data1.Recordset("real_cons") = 0 Then
                        Pososto = Val(left$(Data1.Recordset("MORFH"), 3))
               Else
                        Pososto = Data1.Recordset("real_cons")
               End If
                  
               
                  
                  
               If Pososto > 0 Then
                         If mEK > 0 Then
                              mAPAIT_POSOT = mEK * JOBLIST.Recordset("BAROS_PANI") / Pososto
                              mNERO = mAPAIT_POSOT - mEK * JOBLIST.Recordset("BAROS_PANI") / 100
                         Else
                              mAPAIT_POSOT = mGL * JOBLIST.Recordset("BAROS_PANI") * JOBLIST.Recordset("SXESH_MPAN") / (Pososto * 10)
                              mNERO = mAPAIT_POSOT - (mGL) * JOBLIST.Recordset("SXESH_MPAN") / 1000
                         End If
               End If
Periexon_Nero = mNERO

End Function

Function FIND_LEFT2()
   FIND_LEFT2 = 50 + Grid1.left + Grid1.ColWidth(0) + Grid1.ColWidth(1)
End Function

'
'
Sub show_on_grid()
Dim Ok_Print, M_ALATI, M_COL2, M_ROW2


Ok_Print = 1
' σχεδίαση οθόνης με προηγούμενες συνταγές

On Error Resume Next
Dim k, L As Integer

' σβήνω τις προηγούμενες να μην μένουν απομεινάρια
For k = 1 To Grid1.Rows - 1
   For L = 3 To Grid1.Cols - 1
       Grid1.Row = k
       Grid1.Col = L
       Grid1.text = " "
   Next
Next

Dim mm_aa As String

'το m_aa εχει την τελευταία προσπάθεια + 1
mm_aa = Right$("00" + LTrim(str(Val(m_aa) - 1)), 2)
Grid1.Col = 3

'prospau2.Recordset.Seek "<=", m_Entolh, m_aa

 PROSXHM.Recordset.MoveLast
Do While Not PROSXHM.Recordset.BOF And PROSXHM.Recordset("entolh") = m_entolh
  '   perpathse = 0
     M_ALATI = 0
     Do While Not PROSXHM.Recordset.BOF And PROSXHM.Recordset("aa_prospau") = mm_aa
         If Not IsNull(PROSXHM.Recordset("alati")) Then
             M_ALATI = PROSXHM.Recordset("alati")
         End If


         If Not IsNull(PROSXHM.Recordset("ximikon")) Then
            If PROSXHM.Recordset("ximikon") = 1 Then
               ximikon.value = vbChecked
            Else
               ximikon.value = vbUnchecked
            End If
             
         End If




         If Not IsNull(PROSXHM.Recordset("seira")) Then
              If PROSXHM.Recordset("seira") <= max_Row Then
                 'Grid1.ForeColor = ASPRO '&HFF0000
                 Grid1.Row = PROSXHM.Recordset("ayjon")
                 
                 Grid1.ColAlignment(Grid1.Col) = 0  ' center
                 
                 ' βάζω μεγαλύτερο ύψος στην σειρά με γρ/λιτρο(μόνο στην τελευταία αριστερά κολόνα)
             If Grid1.Col = 3 Then
                 If IsNull(PROSXHM.Recordset("ek")) Then
                 ElseIf PROSXHM.Recordset("ek") > 0 Then
                     Grid1.RowHeight(PROSXHM.Recordset("seira")) = YCOS
                 End If
                 If IsNull(PROSXHM.Recordset("gl")) Then
                 ElseIf PROSXHM.Recordset("gl") > 0 Then
                     Grid1.RowHeight(PROSXHM.Recordset("ayjon")) = YCOS + 15
                 End If
             End If
                If Ok_Print = 1 Then
                   If PROSXHM.Recordset("rek") > 0 And Check1 Then
                      Grid1.text = Format(PROSXHM.Recordset("rek"), "#0.####")
                   Else
                      Grid1.text = Format(PROSXHM.Recordset("ek") + PROSXHM.Recordset("gl"), "#0.####")
                   End If
                End If
                 ' τίτλος προσπάθειας
                 m_ROW = Grid1.Row
                 Grid1.Row = 0
                 '        Grid1.ColAlignment(Grid1.Col) = 2  ' center
                 
                              '  If Ok_Print = 1 Then
               If Ok_Print = 1 Then
                 Grid1.text = "  " + Format(Val(PROSXHM.Recordset("aa_prospau")), "#####")
                End If
                 Grid1.Row = m_ROW
                 
               If Ok_Print = 1 Then
                 Text1(Grid1.Col - 3).text = Format(Val(PROSXHM.Recordset("aa_prospau")), "#####")
                ' Text1(Grid1.Col - 3).Enabled = False
                 Text1(Grid1.Col - 3).BackColor = gri
                 If PROSXHM.Recordset("status") = 2 Then 'ektelestike
                    Text1(Grid1.Col - 3).ForeColor = kokkino
                 ElseIf PROSXHM.Recordset("status") <= 1 Then  ' zygismeno
                    Text1(Grid1.Col - 3).ForeColor = MAYRO '
                 End If
               End If
                 Dim mGrid_col As Integer
                 
                 'γράφω την περιγραφή του χημικοτεχνικού
                 mGrid_col = Grid1.Col
                 Grid1.Col = 0
                 
               '  If Ok_Print = 1 Then
                      Grid1.text = PROSXHM.Recordset("PERIGRAFH")
                ' End If
                 
                 per_xim(Grid1.Row) = PROSXHM.Recordset("PERIGRAFH")
                 
                 
                 Data1.Recordset.Index = "kod"
                 Grid1.Col = mGrid_col
                       
 
 
                  Max_Row_Updated = IIf(PROSXHM.Recordset("ayjon") > Max_Row_Updated, PROSXHM.Recordset("ayjon"), Max_Row_Updated)



                 'Max_Row_Updated = IIf(PROSXHM.Recordset("seira") > Max_Row_Updated, PROSXHM.Recordset("seira"), Max_Row_Updated)
              End If
         End If
   '      perpathse = 1
         PROSXHM.Recordset.MovePrevious
         If PROSXHM.Recordset.BOF Then Exit Do
     Loop
     
      If DOYLEYO_ALATI <> 0 Then
   '   If M_ALATI > 0 And DOYLEYO_ALATI <> 0 Then
         M_COL2 = Grid1.Col
         M_ROW2 = Grid1.Row
         Grid1.Row = 11
         Grid1.Col = 0
         Grid1.text = Label10.Caption
         Grid1.Col = M_COL2
         Grid1.text = M_ALATI
         Grid1.Row = M_ROW2
         Grid1.Col = M_COL2
      End If
      
      
     If PROSXHM.Recordset.BOF Then Exit Do
     mm_aa = Right$("00" + LTrim(str(Val(PROSXHM.Recordset("aa_prospau")))), 2)
     If Grid1.Col + 1 < Grid1.Cols Then
        Grid1.Col = Grid1.Col + 1
     Else
       ' Exit Do
       
       Ok_Print = 0
     ' print_onoma = 1
     End If
     'PROSXHM.Recordset.MovePrevious
'     If perpathse = 0 Then
 '        PROSXHM.Recordset.MovePrevious
 '    End If
Loop

PROSXHM.Recordset.MoveLast



seira = Max_Row_Updated

Dim rf As Integer

rf = Grid1.RowHeight(seira) ' * (seira + 1)

    ' + Grid1.Row το προσθέτω για διόρθωση


DBCombo1.top = rf + Grid1.top + seira + 1
DBCombo1.left = Grid1.left

maskedbox1.top = rf + Grid1.top + seira + 1
maskedbox1.left = FIND_LEFT1()  '10 + Grid1.left + Grid1.ColWidth(0)

Maskedbox2.top = rf + Grid1.top + seira + 1
Maskedbox2.left = FIND_LEFT2() ' = 30 + Grid1.left + Grid1.ColWidth(0) + Grid1.ColWidth(1) + 20

 Grid1.Row = seira
 Grid1.Row = seira
 
 Grid1.Col = 0
 Grid1.Col = 0
  m_ROW = seira + 1
End Sub

'
'
Sub Update_Grid()
Dim mRow2

mRow2 = Grid1.Row
   Grid1.ForeColor = MAYRO 'mayro
   Grid1.Row = m_ROW
   FRMSYNT2.Caption = m_ROW
   Grid1.Col = 0
   
   Grid1.text = DBCombo1
   
   Grid1.Col = 1
   If Val(maskedbox1.text) <= 0 Then
       Grid1.text = "  "
   Else
       seira = IIf(m_ROW < seira, seira, m_ROW)
       Grid1.text = Format(Val(maskedbox1), "#0.###0")
   End If
   
   
   Grid1.Col = 2
   If Val(Maskedbox2) <= 0 Then
       Grid1.text = "  "
   Else
       seira = IIf(m_ROW < seira, seira, m_ROW)
       Grid1.text = Format(Val(Maskedbox2), "#0.##0")
   End If


If Check2 And DBCombo1 <> "" And Val(maskedbox1) > 0 And DBCombo1.Visible Then
   DBCombo1 = bres2_ximiko(DBCombo1, 1)
   Grid1.Col = 0
   Grid1.text = DBCombo1
End If

   per_xim(m_ROW) = DBCombo1
      
   Grid1.Row = mRow2
   
   If DOYLEYO_ALATI = 1 Then
   Dim sum_al As Single
   Dim k As Integer
   
     For k = 1 To 10
        sum_al = sum_al + Val(Replace(Grid1.TextMatrix(k, 1), ",", "."))
     Next
   
   
   Dim mydb As Database
   Dim r As Recordset
   
   Set mydb = Workspaces(0).OpenDatabase("c:\talos\RECIPIES.MDB")
    Set r = mydb.OpenRecordset("ALATI")
    Do While Not r.eof
        If sum_al >= r("APO") And sum_al < r("EOS") Then
            Text3.text = r("TIMH1")
            Exit Do
        End If
        
        r.MoveNext
        
    Loop
    r.Close
    
   
   
   
   
   End If
   
   
   
   
   
   
   
   
   
   
   
      'Label1(m_ROW).Caption = per_xim(m_ROW)
End Sub


Private Sub cmdΡύθμισηΑυτόματης_Click()


Dim mydb As Database
Dim work As Workspace
Dim r As Recordset


 Set work = Workspaces(0)
Set mydb = work.OpenDatabase("c:\TALOS\RECIPIES.MDB", False, False)

'mydb.Execute ""  ' ximikon.text = r("WEIGHT_BOTTLE_MIN")


Set r = mydb.OpenRecordset("parametroi")
r.MoveNext
Dim n0 As Long
n0 = IIf(IsNull(r!Time_Preheat), 0, r!Time_Preheat)
'r.Close

n0 = InputBox("ελαχιστο βάρος δοσομέτρησης σε mg για να αποφευγει το πυκνό(500-1000) ", "", n0)
If n0 < 100 Or n0 > 5000 Then
    n0 = 500
End If





r.EDIT
  r!Time_Preheat = n0
r.update


'    Text1.text = r("MPANIO")
 '   Text2.text = r("default_baros_panioy")
     r.Close
End Sub

Private Sub Command1_Click()
Text3.text = "0"
   
'Dim r
'r = 0
'For r = 1 To 18
 '  Grid1.Row = r
 '  Grid1.Col = 1
 '  Grid1.Text = Str(r)
'Next
   
   
 '  Grid1.Row = 1
  ' Grid1.Col = 1
   'Grid1.Text = "dfdsf"

'prospau2.Enabled = False
'Set MyDB = Workspaces(0).OpenDatabase("c:\TALOS", False, False, "dBASE IV;")


'SQL = "CREATE INDEX entolh ON prospau2 (ENTOLH,AA_PROSPAU,SEIRA);"
'MyDB.Execute SQL
'prospau2.Enabled = True
   
End Sub


      
Function val2(n)
       s = IIf(InStr(n, "%") = 0, 6, InStr(n, "%"))
      If IsNull(s) Then
         val2 = 0
    Else
       val2 = Val(left(n, s - 1))
    End If
End Function


Private Sub Check1_Click()
Dim r, c, m
    
    
    r = Grid1.Row
    c = Grid1.Col
    Grid1.Row = 0
    Grid1.Col = 3
    'αν δεν εχει παλιες συνταγές τότε χάνει αυτήν που έγραφες
    If Val(Grid1.text) > 0 And Command3.Enabled = True Then
       show_on_grid
       proth_fora = 1
       Form_Paint
    End If
    Grid1.Row = r
    Grid1.Col = c
End Sub

Private Sub Command2_Click()
'----------------------------- ΕΝΗΜΕΡΩΣΗ ΣΥΝΤΑΓΗΣ ------------------------
Dim update, mC, mGL, mEK, m_ok

  If maskedbox1.Enabled = False Then
      MsgBox mL_Res(305) ' "Εχεις ήδη ενημερώσει την συνταγή.Ζήτησε <Νέα Συνταγή>"
      ' 305 "Εχεις ήδη ενημερώσει την συνταγή.Ζήτησε <Νέα Συνταγή>"
      
      Exit Sub
  End If
 On Error Resume Next
  If Check_Mpanio() = -1 Then Exit Sub
  On Error GoTo 0
  
  
  
     
   m_ok = 0
  For k = 1 To max_Row - 1
               Grid1.Row = k - 1
               Grid1.Col = 1
               mEK = IIf(IsNull(Grid1.text), 0, Val(Grid1.text))
               Grid1.Col = 2
               mGL = IIf(IsNull(Grid1.text), 0, Val(Grid1.text))
             If mEK + mGL > 0 Then
                   Grid1.Col = 0
                   If IsNull(Grid1.text) Then
                       MsgBox mL_Res(306) + str(k - 1)
                        ' 306 "Δεν δηλώσατε χρώμα στην σειρά "
                       Exit Sub
                   ElseIf Len(Trim(Grid1.text)) = 0 Then
                       MsgBox mL_Res(307) + str(k - 1)
                       '307  "Δεν δηλώσατε χρώμα στην σειρά "
                     Exit Sub
                   End If
                   m_ok = 1
             End If
  Next
  If m_ok = 0 Then
     MsgBox mL_Res(308) ' "Ελλειπής συνταγή"
     Exit Sub
  End If
         
         
If M_AA2 > 0 Then
  Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
  SQL = "delete *from prospau2   WHERE entolh='" + m_entolh + "' and aa_prospau='" + Right$("00" + LTrim(str(M_AA2)), 2) + "';"
  mydb.Execute SQL
End If
         
         
         
         
         
     
  Command2.Enabled = False
  Command3.Enabled = True
  update = False
  
     
  mC = 0
  Grid1.Col = 0
  
  
  
  
  
  Data1.Recordset.Index = "kod"
  
  For k = 1 To max_Row
     Grid1.Row = k
     Grid1.Col = 0
     If Not IsNull(Grid1.text) Then
          If Len(LTrim(Trim(Grid1.text))) = 0 Then
              Exit For
          Else
            update = True
              Data1.Recordset.Seek "=", KOD_XHM(k)
              
              Grid1.Col = 1
              
               mEK = IIf(IsNull(Grid1.text), 0, Val(Grid1.text))
               Grid1.Col = 2
               mGL = IIf(IsNull(Grid1.text), 0, Val(Grid1.text))
             If mEK + mGL > 0 Then
              mC = mC + 1
              prospau2.Recordset.AddNew
              prospau2.Recordset("ENTOLH") = m_entolh
              prospau2.Recordset("SXESH_MPAN") = mJOBLIST.JOBLIST.Recordset("SXESH_MPAN")
             If M_AA2 = 0 Then
                 prospau2.Recordset("AA_PROSPAU") = m_aa
             Else
                 prospau2.Recordset("AA_PROSPAU") = Right$("00" + LTrim(str(M_AA2)), 2)
             End If
              prospau2.Recordset("SEIRA") = mC
                    prospau2.Recordset("STATUS") = 0
              prospau2.Recordset("AYJON") = Grid1.Row
              prospau2.Recordset("KOD") = Data1.Recordset("KOD")
              prospau2.Recordset("PERIGR") = Data1.Recordset("PERIGRAFH")
              prospau2.Recordset("apoxrvsh") = mJOBLIST.JOBLIST.Recordset("apoxrvsh")
              prospau2.Recordset("kod_pel") = mJOBLIST.JOBLIST.Recordset("kod_pel")
              prospau2.Recordset("hme_parad") = mJOBLIST.JOBLIST.Recordset("hme_parad")
              prospau2.Recordset("baros_pani") = mJOBLIST.JOBLIST.Recordset("baros_pani")
              
              
              If ximikon.value = vbChecked Then
                 prospau2.Recordset("ximikon") = f_barosXimikoy
              Else
                 prospau2.Recordset("ximikon") = 0
              End If
              
              
              prospau2.Recordset("EK") = mEK
              prospau2.Recordset("UserW") = User_ID
              Grid1.Col = 2
              prospau2.Recordset("hme") = Date
              If IsNull(JOBLIST.Recordset("apor")) Then
                            prospau2.Recordset("apor") = 0
              Else
                            prospau2.Recordset("apor") = JOBLIST.Recordset("apor")
              End If
              prospau2.Recordset("GL") = mGL
              prospau2.Recordset("ALATI") = Text3.text
              prospau2.Recordset.update
             End If
          End If
       End If
  Next
  
  If update Then
     If M_AA2 = 0 Then
         m_aa = Right$("00" + LTrim(str(Val(m_aa) + 1)), 2)
     End If
     PROSXHM.Refresh
  End If
   show_on_grid
   
    
    
    rf = Grid1.RowPos(m_ROW) 'Grid1.SelStartRow)
    ' + Grid1.Row το προσθέτω για διόρθωση
DBCombo1.top = rf + Grid1.top + Grid1.Row
DBCombo1.left = Grid1.left

maskedbox1.top = rf + Grid1.top + Grid1.Row
maskedbox1.left = FIND_LEFT1() ' Grid1.left + Grid1.ColWidth(0)

Maskedbox2.top = rf + Grid1.top + Grid1.Row
Maskedbox2.left = FIND_LEFT2() ' = Grid1.left + Grid1.ColWidth(0) + Grid1.ColWidth(1) + 20
    
maskedbox1.text = 0
Maskedbox2.text = 0
   
   
   
   
   DBCombo1.Enabled = False
   maskedbox1.Enabled = False
   Maskedbox2.Enabled = False
    Check1.Enabled = True
   Text2.text = " "
   
M_AA2 = 0
   
End Sub

Private Sub Command3_Click()
' **********************   ΝΕΑ ΣΥΝΤΑΓΗ   ***************************
Dim m As String
m = ""

 work_focus = True
   Check1.Enabled = False
   Command3.Enabled = False
   Command2.Enabled = True
   DBCombo1.Enabled = True
   maskedbox1.Enabled = True
   Maskedbox2.Enabled = True
   
   'DBCombo1.Visible = True
   'MaskEdBox1.Visible = True
   'MaskEdBox2.Visible = True

  Command5.Enabled = True
Dim k As Integer

For k = 1 To Max_Row_Updated
    Grid1.Row = k
    Grid1.Col = 1
    Grid1.text = ""
    
    
    Grid1.Col = 2
    Grid1.text = m
Next




For k = 1 To Max_Row_Updated
    Grid1.Row = k
    Grid1.Col = 3
    m = Grid1.text
    If Grid1.RowHeight(k) = YCOS Then
       Grid1.Col = 1
    Else
       Grid1.Col = 2
    End If
    Grid1.text = m
Next
If DBCombo1.Visible And DBCombo1.Enabled Then DBCombo1.SetFocus


'Text3.Text = 0   ' ALATI
End Sub

Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error Resume Next
  
   If KeyCode = 13 Then SendKeys "{TAB}"

End Sub


Private Sub Command4_Click()
Unload Me
mJOBLIST.Show
End Sub

Private Sub Command5_Click()
   Text3.text = Val(Replace(Text3.text, ",", ".")) * 0.8
   Command5.Enabled = False
   
End Sub

Private Sub DBCombo1_Click(Area As Integer)
      'Update_Grid

End Sub

Private Sub DBCombo1_GotFocus()
  ' DBCombo1.BackColor = BLE
  ' DBCombo1.ForeColor = ASPRO
End Sub


Private Sub DBCombo1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub DBCombo1_LostFocus()
   DBCombo1.BackColor = old_cb
   DBCombo1.ForeColor = old_cf
   Dim k As Integer
   
For k = 1 To seira
   If k = m_ROW Then
   Else
      If DBCombo1 = per_xim(k) Then
          MsgBox mL_Res(309) '"Δεν επιτρέπεται να χρησιμοποιείς δύο φορές το ίδιο μπουκάλι" ' περιγραφή του χημικού
          DBCombo1 = " "
      End If
   End If
Next

 Update_Grid
' If MaskEdBox1.Visible Then
 '   If MaskEdBox1.Enabled = True Then
  '       MaskEdBox1.SetFocus
   ' End If
 'End If
End Sub



Private Sub Form_Click()
 
 
 
 'Dim A
 'A = 1
End Sub


Private Sub Form_DblClick()
 Dim a
      a = save_colors(Me, Me.BackColor)
End Sub


Private Sub Form_Load()
Dim L
M_AA2 = 0
work_focus = False
MDIForm1.Arxeia(10).Visible = False
MDIForm1.Syntages(20).Visible = False
MDIForm1.bohuhtika(30).Visible = False
MDIForm1.ejodos(40).Visible = False




Label12.Caption = mL_CapRes(860) '  "μπανιο"
Check1.Caption = mL_CapRes(861) 'τελικη συνταγη
Label16.Caption = mL_CapRes(280) '  "Εντολή"
Label17.Caption = mL_CapRes(281) '  "Πελάτης"
Label18.Caption = mL_CapRes(282) '  "Απόχρωση"
Label19.Caption = mL_CapRes(283) '  "Βαφή"
Label20.Caption = mL_CapRes(284) '  "Αρ.Παρα/λίας"
Label21.Caption = mL_CapRes(285) '  "Πανί"
Label22.Caption = mL_CapRes(860) '  "μπανιο
Check2.Caption = mLResnew(910, "Auto solution selection*", 0)

Label8.Caption = mL_CapRes(295) 'Μεταβολή %
Command3.Caption = mL_CapRes(296) 'Nέα Συνταγή
Command2.Caption = mL_CapRes(297) 'Ενημέρωση Συνταγής
Command4.Caption = mL_CapRes(298) 'Εξοδος
Label10.Caption = mL_CapRes(299) 'Ξηρό αλάτι gr/Lit
If DOYLEYO_ALATI = 0 Then
    Text3.Visible = False
    Label10.Visible = False
End If

a = find_colors(Me)





'max_Row = 30
Max_Row_Updated = 0
proth_fora = 1
'       m_Entolh = "000025"
seira = 1
old_cb = DBCombo1.BackColor
old_1b = maskedbox1.BackColor
OLD_2B = Maskedbox2.BackColor

old_cf = DBCombo1.ForeColor
old_1f = maskedbox1.ForeColor
OLD_2F = Maskedbox2.ForeColor
m_ROW = 1
m_COL = 0
'Grid1.ColAlignment(2) = vbCenter '    ColAlignment(2) = 2  ' center
Grid1.ColAlignment(1) = 2   ' center
    
CAN_UPDATE_GRID = 1  'FLAG ΑΝ ΘΑ ΜΕΤΑΚΙΝΕΙΤΑΙ ΤΟ COMBOBOX1

Grid1.width = Me.width

 'πλάτος στήλης 1 =  MaskEdBox1.Width
   Grid1.Row = 1
   Grid1.Col = 1
   Grid1.ColWidth(1) = maskedbox1.width
   
 'πλάτος στήλης 2 =  MaskEdBox2.Width
   Grid1.Row = 1
   Grid1.Col = 2
   Grid1.ColWidth(2) = Maskedbox2.width
   
'πλάτος 0 στήλης = πλάτος  DBCombo1.Width
   Grid1.Row = 1
   Grid1.Col = 0
   Grid1.ColWidth(0) = DBCombo1.width
   
   
   
   Dim k As Integer
   
   YCOS = 315  ' 315
For k = 0 To Grid1.Rows - 1
   Grid1.RowHeight(k) = YCOS
Next
DBCombo1.height = YCOS
maskedbox1.height = YCOS
Maskedbox2.height = YCOS
   

Grid1.Row = 0
Grid1.Col = 1
Grid1.text = "   % "
Grid1.Col = 2
Grid1.text = "   g/l "

Dim r As Integer

For k = 0 To 2
  For r = 1 To Grid1.Rows - 1
     Grid1.Row = r
     Grid1.Col = k
     'Grid1.BackColor = BLE
   Next
Next
   
   Dim rf As Long
   
rf = Grid1.RowHeight(0)
   
DBCombo1.top = rf + Grid1.top
DBCombo1.left = Grid1.left

maskedbox1.top = rf + Grid1.top
maskedbox1.left = FIND_LEFT1() ' 20 + Grid1.left + Grid1.ColWidth(0)

Maskedbox2.top = rf + Grid1.top
Maskedbox2.left = FIND_LEFT2() ' = 40 + Grid1.left + Grid1.ColWidth(0) + Grid1.ColWidth(1)



   
   
 Grid1.Row = 1 'SelStartRow = 1
 Grid1.Row = 1 'SelEndRow = 1
 
 Grid1.Col = 0 'SelStartCol = 0
 Grid1.Col = 0 '  SelEndCol = 0
 
 
 'ΥΠΟΛΟΓΊΖΩ ΤΙΣ ΘΕΣΕΙΣ ΤΩΝ ΤΕΧΤ
    
L = Grid1.left + Grid1.ColWidth(0) + Grid1.ColWidth(1) + Grid1.ColWidth(2)
    Text1(0).left = 10 + L  '50
    Text1(0).width = Grid1.ColWidth(4) - 10
    Text1(0).top = Grid1.top + 10
    'Text1(0).BackColor = 100
    
 
 For k = 1 To 7
    Text1(k).width = Grid1.ColWidth(2 + k) - 50
    L = L + Grid1.ColWidth(2 + k)
    'Text1(k).left = L + 50 + 20 * k
    Text1(k).left = L + 0 + 20 * k
    
    
    Text1(k).top = Grid1.top + 10
 Next
 
 
 
'mydb.Execute ""  ' ximikon.text = r("WEIGHT_BOTTLE_MIN")


Dim mydb As Database
Dim work As Workspace

Set work = Workspaces(0)
Set mydb = work.OpenDatabase("c:\TALOS\RECIPIES.MDB", False, False)




Dim r7 As Recordset

Set r7 = mydb.OpenRecordset("parametroi")
r7.MoveNext
f_barosXimikoy = IIf(IsNull(r7("WEIGHT_BOTTLE_MIN")), 0, r7("WEIGHT_BOTTLE_MIN"))
r7.Close
 
 
 
 
 
 
End Sub




Private Sub Form_Paint()
Dim mydb, XX
If proth_fora = 1 Then
   
    
   
   
   proth_fora = 0
   JOBLIST.Recordset.Index = "entolh"
   JOBLIST.Recordset.Seek "=", m_entolh
   
   If prospau2.Recordset.RecordCount = 0 Then
      prospau2.Recordset.AddNew
      prospau2.Recordset.update
   End If

'----------------------------------------
   Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
    Set XX = mydb.OpenRecordset("ximitech", dbOpenTable)
    XX.Index = "kod"
   
    PROSXHM.RecordSource = "select prospau2.* from prospau2 where entolh='" + m_entolh + "' order by aa_prospau,seira"
    PROSXHM.Refresh
   
    If Not PROSXHM.Recordset.eof Then PROSXHM.Recordset.MoveFirst
    Do While Not PROSXHM.Recordset.eof
       XX.Seek "=", PROSXHM.Recordset("kod")
       If XX.NoMatch Then
          Command3.Enabled = False
          MsgBox "890-891." + mL_CapRes(890) + PROSXHM.Recordset("kod") + " " + PROSXHM.Recordset("perigr") + mL_CapRes(891) ''"Εχει διαγραφεί από το αρχείο χημικών ο κωδικός "     ".Αδύνατη η κατασκευή συνταγών."
          
          Exit Sub
       End If
       PROSXHM.Recordset.MoveNext
    Loop
    XX.Close
   PROSXHM.RecordSource = "select prospau2.*,ximitech.perigrafh from prospau2 inner join ximitech on prospau2.kod=ximitech.kod where entolh='" + m_entolh + "' order by aa_prospau,seira"
    PROSXHM.Refresh
'----------------------------------------
   
   
   
   
   
   ' SQL = "CREATE INDEX entolh ON prospau2 (ENTOLH,AA_PROSPAU,SEIRA);"
   prospau2.Recordset.Index = "entolh"
   prospau2.Recordset.Seek "<=", m_entolh, "99"
    
    On Error GoTo ff
        a = prospau2.Recordset("entolh")
    On Error GoTo 0
   ' υπάρχει ήδη προσπάθεια
   If prospau2.Recordset("entolh") = m_entolh Then
      m_aa = Right$("00" + LTrim(str(Val(prospau2.Recordset("aa_prospau")) + 1)), 2)
   Else
      m_aa = "01"
   End If
   show_on_grid

Dim rf As Long

'δείχνω τα combo & τα maskedbox στην τελευταία σειρά +1
rf = Grid1.RowPos(m_ROW) 'Grid1.SelStartRow)
DBCombo1.top = rf + Grid1.top + Grid1.Row
DBCombo1.left = Grid1.left

maskedbox1.top = rf + Grid1.top + Grid1.Row
maskedbox1.left = FIND_LEFT1() ' Grid1.left + Grid1.ColWidth(0)

Maskedbox2.top = rf + Grid1.top + Grid1.Row
Maskedbox2.left = FIND_LEFT2() ' = Grid1.left + Grid1.ColWidth(0) + Grid1.ColWidth(1) + 20

maskedbox1.Enabled = False
Maskedbox2.Enabled = False
DBCombo1.Enabled = False
Grid1.Row = 1



If DOYLEYO_ALATI = 0 Then
Else
    Text3.left = maskedbox1.left
    Text3.width = (Maskedbox2.left + Maskedbox2.width) - maskedbox1.left
End If


 Command3.SetFocus



End If


Exit Sub

ff:
a = Err.Number

   'If prospau2.Recordset.RecordCount = 0 Then
      prospau2.Recordset.AddNew
      prospau2.Recordset.update
      prospau2.Recordset.MoveFirst
   'End If
   Resume Next




End Sub



Private Sub Form_Resize()
Grid1.width = Me.width
End Sub

Private Sub Grid1_Click()
Dim mRow2, dum
11
If work_focus = False Then
   MsgBox mL_Res(310) '"Ζήτησε Νέα συνταγή"
   Exit Sub
End If
 
 
 
  If maskedbox1.Enabled = False Then
      MsgBox mL_Res(311) ' "Εχεις ήδη ενημερώσει την συνταγή.Ζήτησε <Νέα Συνταγή>"
      Exit Sub
  End If

 
 
 If CAN_UPDATE_GRID = 0 Then
      Exit Sub
 End If
   
 ' αν πάει να πηδήξει σειρά να μην προχωρά
 If Grid1.Row > seira + 1 Then
    Exit Sub
 End If
 
 ' αν η σειρά έχει ήδη χημικό και δεν μπορεί να αλλάξει
 ' να απενεργοποιείται το Combo1
 If Grid1.Row <= Max_Row_Updated Then
    'DBCombo1.Enabled = False
    DBCombo1.Visible = False
 Else
    'DBCombo1.Enabled = True
    DBCombo1.Visible = True
 End If
 
    
' η στήλη και η σειρά στην οποία έκανα κλικ
    m_ROW = Grid1.Row
    m_COL = Grid1.Col
    
Dim rf As Integer

    
    rf = Grid1.RowPos(m_ROW) 'Grid1.SelStartRow)
    ' + Grid1.Row το προσθέτω για διόρθωση
DBCombo1.top = rf + Grid1.top + Grid1.Row
DBCombo1.left = Grid1.left

maskedbox1.top = rf + Grid1.top + Grid1.Row
maskedbox1.left = FIND_LEFT1() ' Grid1.left + Grid1.ColWidth(0) + 20

Maskedbox2.top = rf + Grid1.top + Grid1.Row
Maskedbox2.left = FIND_LEFT2() ' = 30 + Grid1.left + Grid1.ColWidth(0) + Grid1.ColWidth(1) + 20
    
    Dim M_ROW2 As Integer
    
 M_ROW2 = Grid1.Row 'κρατάω την τιμή για να δώ αν θα αλλάξει παρακάτω
    
    Grid1.Row = m_ROW
    Grid1.Col = 0
    DBCombo1 = Grid1.text

Grid1.Col = 1
maskedbox1.text = Val(Grid1.text)
Grid1.Col = 2
Maskedbox2.text = Grid1.text
CAN_UPDATE_GRID = 1

If M_ROW2 <> m_ROW Then
'   GoTo 11
   maskedbox1.SetFocus  ' 31-1
   
End If
dum = Check_Mpanio()  '
End Sub





Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
' Grid1.SelEndRow = Grid1.SelStartRow
' Grid1.SelEndCol = Grid1.SelStartCol
 
  maskedbox1.Visible = True
 Maskedbox2.Visible = True

 
End Sub

Private Sub MaskEdBox1_Change()
   'Update_Grid
  

End Sub

Private Sub MaskEdBox1_GotFocus()




If work_focus = False Then
   Exit Sub
End If

   Maskedbox2.Enabled = True
   
   maskedbox1.BackColor = BLE
   maskedbox1.ForeColor = ASPRO
   'Key_Pressed = 0

End Sub


Private Sub MaskEdBox1_KeyDown(KeyCode As Integer, Shift As Integer)




On Error Resume Next
     If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub MaskEdBox1_LostFocus()
Dim a As Integer, m, mRow2
If work_focus = False Then
   Exit Sub
End If

On Error Resume Next

mRow2 = Grid1.Row
   
maskedbox1.text = CdBln(maskedbox1.text)
   
   maskedbox1.BackColor = old_1b
   maskedbox1.ForeColor = old_1f
   If Val(Maskedbox2.text) > 0 Then
      maskedbox1.text = 0
   End If
   
'MaskEdBox2.SetFocus

   
   
   
   
   
   a = Grid1.Col
   Grid1.Col = 3
   m = Val(maskedbox1.text) 'Grid1.Text
   If Val(m) > 0 Then
      Text2.text = Format(100 * (Val(maskedbox1.text) + Val(Maskedbox2.text) - Val(m)) / Val(m), "####.0#")
   End If
   Grid1.Col = a
   
   
   Update_Grid
   
   If Val(m) > 0 Then
     Maskedbox2.Enabled = False
    ' MaskEdBox1.SetFocus
     Grid1_Click
     If mRow2 <> Grid1.Row Then
        maskedbox1.SetFocus
     End If
   End If
  ' Check_Mpanio
   
End Sub


Private Sub MaskEdBox2_Change()
   
 'Update_Grid
End Sub

Private Sub MaskEdBox2_GotFocus()

If work_focus = False Then
   Exit Sub
End If
   Maskedbox2.BackColor = BLE
   Maskedbox2.ForeColor = ASPRO
   Key_Pressed = 0  'φλαγκ που βλέπει αν στο διάστημα got focus - lost focus κάποιος πάτησε πλήκτρο
End Sub


Private Sub MaskEdBox2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{TAB}"

End Sub

Private Sub MaskEdBox2_LostFocus()
Dim a As Integer, m, mRow2

If work_focus = False Then
   Exit Sub
End If
   
   
On Error Resume Next
   
   mRow2 = Grid1.Row
   
   Maskedbox2.text = CdBln(Maskedbox2.text)
   
   Maskedbox2.BackColor = old_1b
   Maskedbox2.ForeColor = old_1f
   
   If Val(maskedbox1.text) > 0 Then
      Maskedbox2.text = 0
   End If
    
    
   a = Grid1.Col
   Grid1.Col = 3
   m = Grid1.text
   If Val(m) > 0 Then
      Text2.text = Format(100 * (Val(maskedbox1.text) + Val(Maskedbox2.text) - Val(m)) / Val(m), "####.0#")
   End If
   Grid1.Col = a
    
    
    Update_Grid
    'Grid1.Row = mRow2
     Grid1.Row = Grid1.Row + 1
         maskedbox1.SetFocus
     Grid1_Click
     
    dum = Check_Mpanio()

     If DBCombo1.Visible Then
        DBCombo1.SetFocus
     End If
     

End Sub

Private Sub Text1_DblClick(Index As Integer)
' **********************  DIORUVSE ΣΥΝΤΑΓΗ   ***************************
' diortonei mono ayta poy den ektelesthkan
If Not Text1(Index).ForeColor = MAYRO Then
   Exit Sub
End If






Text1(Index).BackColor = Red
M_AA2 = m_aa - 1 - Index
 
 
 work_focus = True
   Check1.Enabled = False
   Command3.Enabled = False
   Command2.Enabled = True
   DBCombo1.Enabled = True
   maskedbox1.Enabled = True
   Maskedbox2.Enabled = True
   
   'DBCombo1.Visible = True
   'MaskEdBox1.Visible = True
   'MaskEdBox2.Visible = True


For k = 1 To Max_Row_Updated
    Grid1.Row = k
    Grid1.Col = 3 + Index
    m = Grid1.text
    If Grid1.RowHeight(k) = YCOS Then
       Grid1.Col = 1
    Else
       Grid1.Col = 2
    End If
    Grid1.text = m
Next
'Text3.Text = 0   ' ALATI

      





End Sub


Private Sub Text2_LostFocus()
Dim mcol, mRow2 As Integer
  mRow2 = Grid1.Row
  mcol = Grid1.Col
  If grid1_row >= Max_Row_Updated Then
     Exit Sub
  End If
  If Val(maskedbox1.text) > 0 Then
     Grid1.Col = 3
     maskedbox1.text = Val(Grid1.text) * (100 + Val(Text2.text)) / 100
     Grid1.Col = 1
     Grid1.text = maskedbox1.text
     
  End If
If Val(Maskedbox2.text) > 0 Then
     Grid1.Col = 3
     Maskedbox2.text = Val(Grid1.text) * (100 + Val(Text2.text)) / 100
  End If
  Grid1.Col = mcol
End Sub


Private Sub Text3_LostFocus()
   If Not IsNumeric(Text3.text) Then
       Text3.text = 0
  End If
End Sub

Private Sub VScroll1_Change()
   Text2.text = VScroll1.value
   Text2_LostFocus
End Sub


Private Sub VScroll1_GotFocus()
   
   VScroll1.value = Val(Text2.text)
   VScroll1.SmallChange = 1
   
End Sub



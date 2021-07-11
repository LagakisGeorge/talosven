VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form WAIT 
   BackColor       =   &H00FFFF00&
   ClientHeight    =   6975
   ClientLeft      =   1335
   ClientTop       =   1320
   ClientWidth     =   9660
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6975
   ScaleWidth      =   9660
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "start weighing"
      Height          =   375
      Left            =   7425
      TabIndex        =   20
      Top             =   375
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Update selections"
      Height          =   330
      Left            =   7350
      TabIndex        =   18
      Top             =   1185
      Value           =   1  'Checked
      Width           =   2280
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   7815
      Top             =   5235
   End
   Begin VB.Data apor_syntages 
      Caption         =   "apor_syntages"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\RECIPIES.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2235
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "apor_syntages"
      Top             =   6645
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.CommandButton C_MHDEN 
      Caption         =   "ΜΗΔΕΝΙΣΜΟΣ ΛΙΣΤΑΣ"
      Height          =   435
      Left            =   7305
      TabIndex        =   12
      Top             =   3450
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "TARE"
      Height          =   210
      Left            =   7440
      TabIndex        =   11
      Top             =   450
      Width           =   1080
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Α/Α επιλογής"
      Height          =   270
      Left            =   7440
      TabIndex        =   7
      Top             =   2415
      Width           =   1935
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Αρ.Εργασίας"
      Height          =   270
      Left            =   7440
      TabIndex        =   5
      Top             =   2115
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Πελάτη/Απόχρωση"
      Height          =   285
      Left            =   7440
      TabIndex        =   4
      Top             =   1860
      Width           =   1950
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ΔΙΑΓΡΑΦΗ ΠΡΟΣΠΑΘΕΙΑΣ"
      Height          =   465
      Left            =   7290
      TabIndex        =   3
      Top             =   2895
      Width           =   2430
   End
   Begin VB.Data ximitech 
      Caption         =   "XIMITECH"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\recipies.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "XIMITECH"
      Top             =   6705
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ΕΞΟΔΟΣ"
      Height          =   450
      Left            =   7305
      TabIndex        =   1
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ΕΚΤΕΛΕΣΗ  ΣΥΝΤΑΓΩΝ"
      Height          =   435
      Left            =   7305
      TabIndex        =   0
      Top             =   3960
      Width           =   2430
   End
   Begin VB.Data SQLPROSPAU 
      Caption         =   "SQLPROSPAU"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\recipies.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3900
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "SELECT PROSPAU2.* FROM PROSPAU2  where PROSPAU2.SEIRA=1 AND PROSPAU2.STATUS<2 ORDER BY PROSPAU2.EPIL;"
      Top             =   6270
      Visible         =   0   'False
      Width           =   1515
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   8025
      Top             =   6450
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSDBGrid.DBGrid DBGrid3 
      Bindings        =   "_WAIT.frx":0000
      Height          =   2010
      Left            =   75
      OleObjectBlob   =   "_WAIT.frx":001C
      TabIndex        =   17
      Top             =   5085
      Visible         =   0   'False
      Width           =   7560
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   465
      Top             =   6855
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "_WAIT.frx":0BC7
      Height          =   4755
      Left            =   105
      OleObjectBlob   =   "_WAIT.frx":0BE0
      TabIndex        =   6
      Top             =   120
      Width           =   7170
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   105
      TabIndex        =   19
      Top             =   4890
      Width           =   2000
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Βάρος Πανιού"
      Height          =   285
      Left            =   7290
      TabIndex        =   16
      Top             =   855
      Width           =   1245
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tare"
      Height          =   285
      Left            =   7305
      TabIndex        =   15
      Top             =   435
      Width           =   1185
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "00000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8580
      TabIndex        =   14
      Top             =   810
      Width           =   1080
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "00000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8565
      TabIndex        =   13
      Top             =   450
      Width           =   1080
   End
   Begin VB.Shape Shape2 
      Height          =   1245
      Left            =   7260
      Top             =   1590
      Width           =   2460
   End
   Begin VB.Shape Shape1 
      Height          =   1440
      Left            =   7275
      Top             =   120
      Width           =   2460
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "00000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8565
      TabIndex        =   10
      Top             =   180
      Width           =   1080
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ενδειξη ζυγού"
      Height          =   240
      Left            =   7305
      TabIndex        =   9
      Top             =   180
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ταξινόμηση κατά"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7470
      TabIndex        =   8
      Top             =   1620
      Width           =   2055
   End
   Begin VB.Label text1 
      BackStyle       =   0  'Transparent
      Height          =   465
      Left            =   0
      TabIndex        =   2
      Top             =   5670
      Width           =   8985
   End
End
Attribute VB_Name = "WAIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Dim ayjon
Dim err_msg(1 To 10)
Dim mTara
Dim proth_fora
Dim yparxei_zygos
Dim zygizo
Dim dokim(1 To 200)  '\xhmika που θα χρησιμοποιηθούν στην δοκιμαστική συνταγή
Dim neo_x ' για το παραπάνω counter
Dim DT
Dim old_C
Const Keramidi = &H80C0FF
Const Prasino = &H80FF&
Const Portokali = &H80FF&
Dim Diplo_Pathma
Dim mtime   '= GetCurrentTime()
Dim EXO_SYNTAGES_GIA_SERVIRISMA

Function Zyg_Hi_Triplo(ByVal Kauysterhsh As Long, akrivia) As Single
Dim Z1, z2, z3, mo, counter
Command1.Enabled = False
counter = 1
 Do While True
   counter = counter + 1
   Z1 = Zygis0(0)
     MilSec Kauysterhsh
   z2 = Zygis0(0)
      MilSec Kauysterhsh
   z3 = Zygis0(0)
   mo = (Z1 + z2 + z3) / 3
   If Abs(Z1 - mo) < akrivia And Abs(z2 - mo) < akrivia And Abs(z3 - mo) < akrivia Then
         Exit Do
   End If
   If counter > 50 Then
      MsgBox mL_Res(483) '"Η ζυγαριά δεν σταθεροποιείται.Η τιμή δεν είναι ακριβής"
      Exit Do
   End If
 Loop
Zyg_Hi_Triplo = Int(mo + 0.5)
Command1.Enabled = True

End Function

Function Zygis0(x As Single)
  Dim counter, mstarter
  
  
   If Balance_Type = "ADAM" Then
       Zygis0 = Zygis0_ADAM(MSComm1, x)
       Exit Function
    End If
    
  
  
  
    ' Flush the input buffer.
       MSComm1.Output = "SI " + Chr$(13) + Chr$(10)
      ' MSComm1.Output = "UPD 20 " + Chr$(13) + Chr$(10)
       MSComm1.InBufferCount = 0
    mstarter = GetCurrentTime()
    counter = 0
    ' Wait for "OK" to come back from the modem.
    Do
       dummy = DoEvents()
          If GetCurrentTime() - mstarter > 5000 Then
            Zygis0 = -999000999
            Exit Function
      End If
       ' If there is data in the buffer, then read it.
      ' MilSec 10
       If MSComm1.InBufferCount Then
          FromModem$ = FromModem$ + MSComm1.Input
          ' Check for "OK".
          Me.Caption = FromModem$
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
                 On Error Resume Next
                    MSComm1.PortOpen = True
                 On Error Resume Next
                    
                 MSComm1.Output = "SIR " + Chr$(13) + Chr(10)
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

End Function


Function add_apor(e, p, aitia)
  On Error Resume Next
        apor_syntages.Recordset.AddNew
       
        apor_syntages.Recordset("entolh") = e
        apor_syntages.Recordset("aa_prospau") = p
        apor_syntages.Recordset("aitia") = left(aitia, 70)
        
        apor_syntages.Recordset.update
     add_apor = 0
End Function
Sub undelete_SYNTAGH()


On Error Resume Next
' επαναφέρω την μισοτελειωμένη συνταγή
Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
SQL = "select prospau2.* from prospau2 where prospau2.status<2 order by entolh,aa_prospau,seira;"
Set e = mydb.OpenRecordset(SQL, dbOpenDynaset)
mishE = "0"
mishA = 0
    
e.MoveFirst
m_ss = e("entolh")
m_aa = e("aa_prospau")
ok = 0
Do While Not e.eof
     If m_ss = e("entolh") And m_aa = e("aa_prospau") And e("seira") = 1 Then
          ok = 1
    End If
    e.MoveNext
  If e.eof Then
       If ok = 0 Then
                mishE = m_ss
                mishA = m_aa
        End If
       
  Else
      If m_ss = e("entolh") And m_aa = e("aa_prospau") Then
         ' παμε παρακάτω
      Else
          If ok = 1 Then
              ' no problem
                ok = 0
                m_ss = e("entolh")
                m_aa = e("aa_prospau")
           Else
                mishE = m_ss
                mishA = m_aa  'e("aa_prospau")
                Exit Do
            End If  ' ok=0
       End If   ' ---
    End If ' eof
Loop
    
    

'MyDB.Close




           
      
      
      
'          e.MoveFirst
'        mishE = e("entolh")
'          mishA = e("aa_prospau")
'          If Not IsNull(mishE) Then
 '            If Val(mishE) > 0 Then
  '             sql = "UPDATE PROSPAU2 SET STATUS=1  WHERE entolh='" + mishE + "' and aa_prospau='" + mishA + "';"
   '            MyDB.Execute sql
    '         End If
     '     End If
          



      
'          e.MoveFirst
'        mishE = e("entolh")
'          mishA = e("aa_prospau")
          If Not IsNull(mishE) Then
             If Val(mishE) > 0 Then
               SQL = "UPDATE PROSPAU2 SET STATUS=1,regr=0,RENERO=0  WHERE entolh='" + mishE + "' and aa_prospau='" + mishA + "';"
           '    On Error GoTo 0
               mydb.Execute SQL
               
               SQL = "select prospau2.* from prospau2 where  entolh='" + mishE + "' and aa_prospau='" + mishA + "';"
              Set e2 = mydb.OpenRecordset(SQL, dbOpenDynaset)
               e2.MoveFirst
               ' βλέπω αν όντως υπάρχει η seira=1 αλλοιώς το σβήνω
               okk = 0
               Do While Not e2.eof
                    If e2("seira") = 1 Then okk = 1
                    e2.MoveNext
                Loop
                If okk = 0 Then
                     SQL = "delete *from prospau2   WHERE entolh='" + mishE + "' and aa_prospau='" + mishA + "';"
                     mydb.Execute SQL
                End If
               
               
             End If
          End If



















mydb.Close
 SQLPROSPAU.Refresh

End Sub

Private Sub AXRHSIMOPOIHTA()
'-----------------
Dim mydb, p, x, counter, SQL, m_Mpoykal

Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
'Set p = MyDB.OpenRecordset("PROSPAU2")
Set x = mydb.OpenRecordset("XIMITECH")

 m_Mpoykal = 0


SQL = "delete *from prospau2   WHERE entolh='000000' and aa_prospau='01';"
mydb.Execute SQL


SQL = "SELECT *FROM PROSPAU2  where  STATUS<2 and epil>0 ;"
Set p = mydb.OpenRecordset(SQL, dbOpenDynaset)



counter = 0

p.MoveFirst
Do While Not p.eof
   If p("ENTOLH") = "000000" And p("epil") = 1 And p("baros_pani") = 5001 Then
      Exit Sub
   End If
   p.MoveNext
Loop


x.MoveFirst
Do While Not x.eof
         If IsNull(x("addr_prot")) Then GoTo skip
         If x("addr_prot") = 0 Then GoTo skip
         If IsNull(x("last_dosom")) Or Not IsDate(x("last_dosom")) Then
              ' nothing   GoTo skip
         Else
              If DateDiff("S", CDate(x("Last_Dosom")), Now) < DT Then GoTo skip
         End If
         
         found = 0
         
         For k = 1 To neo_x
            If dokim(k) = x("kod") Then found = 1 'ειναι στον κατάλογο
         Next
         
         If found = 0 Then GoTo skip
              m_Mpoykal = m_Mpoykal + 1
              
              p.AddNew
              p("ENTOLH") = "000000"
              p("SXESH_MPAN") = 10
              p("AA_PROSPAU") = "01"
               counter = counter + 1
              p("SEIRA") = counter
              p("STATUS") = 1
                p("EPIL") = 1
              p("AYJON") = counter
              p("KOD") = x("KOD")
              p("EK") = 0.2
              p("UserW") = User_ID
              p("hme") = Date
              p("apor") = 1.2
              p("GL") = 0
              p("nero") = 5000
              p("gr") = 2000
              p("baros_pani") = 5001
              p("perigr") = x("perigrafh")
              p.update
              
              If m_Mpoykal >= 30 Then
                  Exit Do
              End If
              
skip:
    x.MoveNext
Loop

End Sub


Function Mhdenismos_Zygoy(DUM)
Dim Can_Zero
Screen.MousePointer = 11
Can_Zero = 1
' ΑΝΟΙΓΩ ΣΕΙΡΙΑΚΗ ΠΟΡΤΑ ΚΑΙ ΜΗΔΕΝΙΖΩ ΖΥΓΑΡΙΑ
On Error Resume Next
       MSComm1.commport = Balance_Port
       MSComm1.Settings = Balance_Settings
       MSComm1.InputLen = 0
       MSComm1.PortOpen = False
       MSComm1.PortOpen = True
       
       
       
        G_Balance_Digits MSComm1, 3
     '  MSComm1.Output = Balance_3_Digits + Chr$(13)
      MilSec 500
      G_REZERO MSComm1
            '   MSComm1.Output = "Z " + Chr(13) + Chr(10)
         MilSec 500
      
      ' MSComm1.Output = Balance_3_Digits + Chr$(13)
         
       mz = Zygis0(0)
       counter_num = 0
         
       Do While Abs(mz) > 2
            counter_num = counter_num + 1
            If counter_num >= 2 Then
                  MsgBox mL_Res(456), , "Talos" ' 456 "Παρακαλώ ανοίξτε τη Ζυγαριά !"
                  MilSec 2000
                 ' MSComm1.Output = "Z " + Chr$(13) + Chr(10)
                  G_REZERO MSComm1
                 
            End If
            MilSec 100
            
            G_REZERO MSComm1
            ' MSComm1.Output = "Z " + Chr$(13) + Chr(10)
            
            mz = Zygis0(0)
            If counter_num >= 3 And Abs(mz) > 6 Then
                Can_Zero = 0
                Exit Do
            End If
       Loop

On Error GoTo 0

If Can_Zero = 1 Then
   Mhdenismos_Zygoy = 1
Else
    If mz < 900000000 Then
        Mhdenismos_Zygoy = -1
        MsgBox mL_Res(457) '"Η ζυγαριά παραμένει κλειστή"
    Else
        Mhdenismos_Zygoy = 0
    End If
End If
Screen.MousePointer = 1
End Function
         
         
Private Sub Check1_Click()
' If Check1.Value = 1 Then
'    SQLPROSPAU.RecordSource = "SELECT * FROM PROSPAU WHERE STATUS in ('0','1',' ') ORDER BY aa;"
' Else
'    SQLPROSPAU.RecordSource = "SELECT * FROM PROSPAU WHERE STATUS in ('0','1',' ') ORDER BY entolh;"
' End If
    
'   SQLPROSPAU.Recordset.MoveLast
'   SQLPROSPAU.Refresh
'       DBGrid1.Visible = False
'       DBGrid1.Visible = True
       
If Check1 Then
   SQLPROSPAU.Refresh
End If

End Sub
Function check_idia(x, a_Err, MaxEpil)
'------------------------ ελέγχω αν υπάρχει το ίδιο χρώμα 2 φορές-----------
Dim synt, k, mSeir As Integer, L

For synt = 1 To MaxEpil
  If x(synt, 1, 1) = -999 Then Exit For
     
     ' για την ίδια συνταγή ελέγχω μήπως έχει 2 ίδια χρώματα
     For mSeir = 1 To 30
           If x(synt, mSeir, 1) = -999 Then Exit For

           For L = mSeir + 1 To 30
                  If x(synt, L, 1) = -999 Then Exit For
                     Text1.Caption = x(synt, mSeir, 4) + " " + x(synt, mSeir, 5) + "-------" + x(synt, mSeir, 1) + "....." + x(synt, L, 1)
                  If x(synt, mSeir, 1) = x(synt, L, 1) Then
                     a_Err(synt) = 1
                     MsgBox "458-460." + mL_CapRes(458) + x(synt, mSeir, 1) + mL_CapRes(459) + x(synt, mSeir, 4) + mL_CapRes(460) + x(synt, mSeir, 5)
                     '458   "Υπάρχει 2 φορές το χρώμα με κωδικό "
                     '459   " στην συνταγή "
                     '460   " προσπάθεια "
                     
                     DUM = add_apor(x(synt, mSeir, 4), x(synt, mSeir, 5), err_msg(1))
                     Exit For
                  End If
           Next
     Next
Next



'ελεγχοι
 
 '  1. αν υπάρχουν 2 ίδια χημικά στην ίδια συνταγή
 '  2. αν οι ποσότητες των χρωμάτων αρκούν για την εκτέλεση των συνταγών
 '  3. αν έχω υπέρβαση μπάνιου
 '  4. αν έχω υπερχείλιση ποτηριού (μεγάλο πανί & μεγάλο μπάνιο)
 '  5. αν το μπουκάλι είναι πάνω στο τραπέζι
 '  6. δεν υπάρχει κάν το χρώμα στο αρχείο
 '  7. δεν έχει συγκέντρωση το χρώμα
 '  8. δεν αποθηκευτηκε ο κωδικός του χρώματος στην συνταγή

End Function
Function Update_Gr(x, a_Err, MaxEpil)
' ελεγχος μπάνιου,επάρκειας χημικών,αχρησιμοποίητες μπουκάλες
Dim synt, k, mSeir As Integer, L, e, Pososto, mAPAIT_POSOT, mNERO, mXR(1 To Ar_Mpoykal)
Dim Ner_dos, XIM, Job, xhmika
Dim db As Database, r As Recordset, SQL As String
Dim Eleiph(1 To 100), mEleiph
Dim ryum_PH, M33, M34 ' φλαγκ και bookmarks για PH
Dim dieyt(1 To 300), c_dieyt
Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
SQL = "select prospau2.* from prospau2 where epil>0 and epil<100 and status<2  order by epil,seira;"
Set e = mydb.OpenRecordset(SQL, dbOpenDynaset)


Set Job = mydb.OpenRecordset("joblist")
    Job.Index = "entolh"



' πίνακας όπου αποθηκεύονται οι ποσότητες που θα χρειαστούν για την δοσομέτρηση
' των συνταγών σε κάθε χρώμα π.χ.  mXR(100) η ποσότητα που θα τραβήξουμε από το μπουκάλι 100
For k = 1 To Ar_Mpoykal
   mXR(k) = 0
Next


Set XIM = XIMITECH.Recordset
XIM.Index = "KOD"

'dyo ximika me idies dieyuynseis
c_dieyt = 0
XIM.MoveFirst
Do While Not XIM.eof
   If XIM("addr_prot") > 0 Then
    
     If XIM("addr_prot") = 206 Then
         debugging = 1
      End If
        
       c_dieyt = c_dieyt + 1
       For k = 1 To c_dieyt - 1
              If XIM("addr_prot") = dieyt(k) Then
                    MsgBox mLResnew(909, "* the same address " + str(XIM("addr_prot")) + " two times.remove both of them with permanent removal", 1)
                    Exit Function
              End If
       Next
       dieyt(c_dieyt) = XIM("addr_prot")

   End If
   XIM.MoveNext
Loop






e.MoveLast
Do While Not e.BOF
    'Ιδια συνταγή
    arx = e("entolh") + e("aa_prospau")
    mNERO = 0
    mAPAIT_POSOT = 0
    ryum_PH = 0 '  1=εχει ρυθμιστικό PH
    Do While Not e.BOF And arx = e("entolh") + e("aa_prospau")
           XIM.Seek "=", e("kod")
           If XIM.NoMatch Then  '"Δεν υπάρχει στο αρχείο ο κωδικός "  ".Δεν θα εκτελεστεί η συνταγή "
             MsgBox "474,469." + mL_CapRes(474) + e("kod") + mL_CapRes(469) + e("entolh") + e("aa_prospau")
             a_Err(e("epil")) = "4"
             DUM = add_apor(e("entolh"), e("aa_prospau"), mL_CapRes(474) + e("kod")) '"Δεν υπάρχει στο αρχείο ο κωδικός "
           End If
           
           If Val(a_Err(e("epil"))) = 0 And (Not XIM.NoMatch) Then
                  If IsNull(XIM("real_cons")) Then
                        mj = InStr(1, XIM("MORFH"), "%")
                        Pososto = Val(left$(XIM("MORFH"), mj - 1))
                        If Pososto = 99 Then
                             M33 = e.Bookmark
                             ryum_PH = 1
                        End If
                  ElseIf XIM("real_cons") = 0 Then
                        mj = InStr(1, XIM("MORFH"), "%")
                        Pososto = Val(left$(XIM("MORFH"), mj - 1))
                        If Pososto = 99 Then
                             M33 = e.Bookmark
                             ryum_PH = 1
                        End If
                  Else
                        Pososto = XIM("real_cons")
                  End If
                  If Pososto > 0 Then
                         If e("EK") > 0 Then
                              mAPAIT_POSOT = Int(e("EK") * e("BAROS_PANI") / Pososto)
                              mNERO = Int(mNERO + mAPAIT_POSOT - (e("EK") * e("BAROS_PANI") / 100))
                         Else
                              mAPAIT_POSOT = e("GL") * e("BAROS_PANI") * e("SXESH_MPAN") / (Pososto * 10)
                              mNERO = mNERO + mAPAIT_POSOT - (e("GL") * e("SXESH_MPAN") / 1000)
                         End If
                   Else
                              mAPAIT_POSOT = 0
                              fault = 1
                              Beep
                              Text1.Caption = Text1.Caption + "δεν έχω συγκέντρωση για το " + XIM("perigrafh")
                   End If
                   mXR(XIM("addr_prot")) = mXR(XIM("addr_prot")) + mAPAIT_POSOT
                   
                   
                   
                   Job.Seek "=", e("entolh")
                   
                   If Job.NoMatch Then
                       xhmika = 0
                   Else
                       If IsNull(Job("xhmika")) Then
                          xhmika = 0
                       Else
                           xhmika = Job("xhmika")
                       End If
                   End If
                   
                   ' αποθηκεύω τα απαιτούμενα μγρ για κάθε συστατικό
                   e.EDIT
                   e("gr") = Int(0.5 + mAPAIT_POSOT)
                   If Pososto > 0 Then
                       e("perigr") = left(XIM("perigrafh"), 30)  '+ "*" + Format(Pososto, "0.00")
                   End If
                   If e("seira") = 1 Then  ' αποθηκεύω το νερό στην 1η σειρά
                       e("ALATIGR") = e("ALATI") * (e("sxesh_mpan") * e("baros_pani") / 1000)    ' MGR αλατιού
                      
                      
Ner_dos = Int(e("sxesh_mpan") * e("baros_pani") - xhmika * 1000 - (mNERO + e("baros_pani") * e("apor")))
'Ner_dos = Int((100 - xhmika) / 100 * e("sxesh_mpan") * e("baros_pani") - (mNERO + e("baros_pani") * e("apor")))
                      
                      If ryum_PH = 1 Then 'αν έχει ρυθμιστικό PH ρίξε ένα γραμμάριο
                         e("nero") = 1000
                      Else
                         e("nero") = Ner_dos
                      End If
                      
                      If Ner_dos < 0 Then  ' ελεγχος ΜΠΑΝΙΟΥ
                          a_Err(e("epil")) = 3
                          MsgBox "461-462." + mL_CapRes(461) + e("entolh") + mL_CapRes(462) + e("aa_prospau")
                          '461 "Υπέρβαση Μπάνιου στην συνταγή "
                          '462 " προσπάθεια "
                          
                          DUM = add_apor(e("entolh"), e("aa_prospau"), err_msg(3))
                      End If
                   End If
                   e.update
                                      
                   If ryum_PH = 1 And e("seira") = 1 Then 'αν έχει ρυθμιστικό βάλε ποσότητα νερού σε ρυθμιστικό
                          M34 = e.Bookmark 'κρατώ την θέση για να την ξαναβρώ
                          e.Bookmark = M33
                          e.EDIT 'βάζω στο ρυθμιστικό την ποσότητα του νερού-1000 που θα τα ρίξει η βαλβίδα του νερού
                          e("gr") = Ner_dos - 1000
                          e.update
                          e.Bookmark = M34 'ξαναπηγαίνει στην θέση του
                   End If

           End If
           e.MovePrevious
           If e.BOF Then  ' ΕΧΕΙ ΑΛΛΑΞΕΙ ΣΥΝΤΑΓΗ.πριν αλλάξω συνταγή υπολογίζω το ρυθμ.PH
              Exit Do
           End If
    Loop  ' ΊΔΙΑ ΣΥΝΤΑΓΗ
    If e.BOF Then Exit Do
Loop


'--------------------
XIM.MoveFirst
mEleiph = 0
neo_x = 0
Do While Not XIM.eof
  
  If IsNull(XIM("addr_prot")) Then GoTo skip_Update_GR
  If XIM("addr_prot") = 0 Then GoTo skip_Update_GR
  If mXR(XIM("addr_prot")) = 0 Then GoTo skip_Update_GR ' δεν χρησιμοποιειται στην συνταγή
     
     If mXR(XIM("addr_prot")) / 1000 > (XIM("ypol_prot") - 50) And mXR(XIM("addr_prot")) > 0 Then
        MsgBox mL_Res(463) + XIM("perigrafh")
        '463 " Δεν επαρκεί ή ποσότητα τού "
        DUM = add_apor(e("entolh"), e("aa_prospau"), mL_Res(463) + XIM("perigrafh"))
        mEleiph = mEleiph + 1
        Eleiph(mEleiph) = XIM("kod")
     Else
         If XIM("lhj_prot") < Date Then
             MsgBox str(XIM("addr_prot")) + " " + Trim(XIM("perigrafh")) + "  " + mL_Res(129) + " : " + Format(XIM("lhj_prot"), "dd/mm/yyyy")
         End If
         If IsNull(XIM("last_dosom")) Or Not IsDate(XIM("last_dosom")) Then
            neo_x = neo_x + 1
            dokim(neo_x) = XIM("kod") ' κρατάει τους κωδικούς για την δοκιμαστική συνταγή
         Else
         
            If DateDiff("S", CDate(XIM("Last_Dosom")), Now) > DT Then
                neo_x = neo_x + 1
                dokim(neo_x) = XIM("kod") ' κρατάει τους κωδικούς για την δοκιμαστική συνταγή
            End If
         
         End If
     End If


skip_Update_GR:
  XIM.MoveNext

Loop


e.MoveFirst
Do While Not e.eof
   'αν δεν έχει ήδη μαρκαριστεί η συνταγή σαν αποριφθείσα
   If Val(a_Err(e("epil"))) = 0 Then
       For k = 1 To mEleiph
          If e("kod") = Eleiph(k) Then
             a_Err(e("epil")) = "2"
             DUM = add_apor(e("entolh"), e("aa_prospau"), err_msg(2))

          End If
       Next
   End If
   e.MoveNext
Loop


' μαρκάρισμα συνταγών για εκτέλεση status=1
e.MoveFirst
Do While Not e.eof
    If Val(a_Err(e("epil"))) = 0 And e("status") < 2 Then
        e.EDIT
        e("status") = 1
        e.update
        If e("seira") = 1 Then
            EXO_SYNTAGES_GIA_SERVIRISMA = EXO_SYNTAGES_GIA_SERVIRISMA + 1
        End If
    Else
       e.EDIT
       e("status") = 0
       e.update
    End If
   e.MoveNext
Loop



e.Close

End Function

Private Sub C_MHDEN_Click()
Dim a, b, response
  
On Error Resume Next
  
  a = SQLPROSPAU.Recordset.RecordCount
  
  
         
         response = MsgBox(mL_Res(464), vbYesNo) ' +  + SQLPROSPAU.Recordset("aa_prospau")
         '464 "Να διαγραφεί όλη η λίστα με όλες τις συνταγές; "
         If response = vbYes Then
         
            Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
            sql_del = "DELETE PROSPAU2.* FROM PROSPAU2 WHERE prospau2.status<2;"
            mydb.Execute sql_del
            mydb.Close
         End If
         
         SQLPROSPAU.Refresh
         DBGrid1.Refresh

On Error GoTo 0





End Sub

Private Sub Command1_Click()
'exodos
Dim flag, mcount
Screen.MousePointer = 1
flag = 0
mcount = 0
If prozygismena = 0 And Command4.Visible = False Then ' Caption <> mL_CapRes(417)
      Timer1_Timer
      
      Do While Val(Label3.Caption) > 200
          flag = 1
          If mcount > 2 Then Exit Do
          MsgBox mL_Res(465), , "Talos"
          '465 "Απομακρύνετε τον κύλινδρο από την ζυγαριά"
          DoEvents
          mcount = mcount + 1
          Timer1_Timer
      Loop
      
         Label3.Caption = mz
         Label5.Caption = " "
      
      
      Timer1.Enabled = False
      C_MHDEN.Enabled = True
      Command5.Enabled = True
      Command3.Enabled = True
      Command1.Caption = mL_CapRes(417)  ' "ΕΞΟΔΟΣ"
     If flag = 1 Then
        Exit Sub
     End If
End If
   sel_menu = 0
If prozygismena = 0 Then
  Timer1.Enabled = False
  
  MilSec 100
  On Error Resume Next
    
    MSComm1.PortOpen = False
  On Error GoTo 0
 End If
 
  
MDIForm1.Arxeia(10).Visible = True
MDIForm1.Syntages(20).Visible = True
MDIForm1.bohuhtika(30).Visible = True
MDIForm1.ejodos(40).Visible = True
  
  
  Unload Me

End Sub


Private Sub old_Command3_Click()
'***************************** ΕΚΤΕΛΕΣΗ ΣΥΝΤΑΓΩΝ *************************************

     
End Sub


Private Sub Command2_Click()
'tare
Dim mz, mz2
  If prozygismena <> 0 Then
      Exit Sub
  End If
  Screen.MousePointer = 1
'  Command2.Enabled = False
  Command1.Enabled = False
  
  If yparxei_zygos = 0 Then
          If Mhdenismos_Zygoy(0) <> 1 Then
              MsgBox mL_Res(466) ' "Ζύγισμα αδύνατο"
              Command2.Enabled = True
              Command1.Enabled = True

              Exit Sub
          End If
          yparxei_zygos = 1
          Command2.Caption = "TARE"
  End If
         


1432
        
        MilSec 500
         mz = Zygis0(0)
        
        MilSec 500
         mz2 = Zygis0(0)
       
    If Abs(mz2 - mz) > 2 Then GoTo 1432
    
    
    
       
         If mz < 20 Then
             MsgBox mL_Res(467), , "Talos"
             ' 467 "Παρακαλώ τοποθετήστε τον κύλινδρο στην Ζυγαριά !"
         Else
             mTara = mz
             Command2.Enabled = False
             Label3.Caption = mz
             Label4.Caption = mTara
    
             'Timer1.Enabled = True
         End If
         C_MHDEN.Enabled = False
         Command5.Enabled = False
         Command3.Enabled = False
        ' Command2.Enabled = True
         Command1.Enabled = True
         Command1.Caption = mL_CapRes(420) '"Τέλος ζυγίσεων"
End Sub

Private Sub Command3_Click()
Command3.Enabled = False
 
 'ελεγχοι
'epil= select prospau2.*,ximitech.morfh,ximitech.perigrafh,ximitech.kod  from prospau2 inner join ximitech on prospau2.kod=ximitech.kod  where prospau2.epil>0 and prospau2.epil<100 order by prospau2.epil,prospau2.seira;
 '**** 1. αν υπάρχουν 2 ίδια χημικά στην ίδια συνταγή
 '**** 2. αν οι ποσότητες των χρωμάτων αρκούν για την εκτέλεση των συνταγών
 '**** 3. αν έχω υπέρβαση μπάνιου
 '**** 4. αν έχω υπερχείλιση ποτηριού (μεγάλο πανί & μεγάλο μπάνιο)
 '***  5. αν το μπουκάλι είναι πάνω στο τραπέζι
 '***  6. δεν υπάρχει κάν το χρώμα στο αρχείο
 '***  7. δεν έχει συγκέντρωση το χρώμα
 '***  8. δεν αποθηκευτηκε ο κωδικός του χρώματος στην συνταγή
 'ΥΠΟΛΟΓΙΣΜΟΙ
 '***  1. Ποσότητα δοσομέτρησης
 '***  2. Ποσότητα νερού
Dim mpanio(1 To 100), pani(1 To 100), apor(1 To 100)
Dim e, x(1 To 100, 1 To 30, 1 To 5), MaxEpil, a_Err(1 To 100), synt
Dim mydb, DUM, ans, m_error, flag_last_syntagh
Dim mEpilogh, mBaros_pp
Dim Mycom, mycomset


If prozygismena = 0 Then
  Timer1.Enabled = False
  MilSec 100
  On Error Resume Next
      MSComm1.PortOpen = False
  On Error GoTo 0
End If



Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
SQL = "SELECT *FROM PROSPAU2  where  STATUS<2  ORDER BY entolh,aa_prospau,seira;"
Set all_prospau = mydb.OpenRecordset(SQL, dbOpenDynaset)

Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
SQL = "select prospau2.*,ximitech.*  from prospau2 inner join ximitech on prospau2.kod=ximitech.kod  where prospau2.epil>0 and prospau2.epil<100  and status<2 order by prospau2.epil,prospau2.seira;"
Set EPIL = mydb.OpenRecordset(SQL, dbOpenDynaset)




'all_PROSPAU. = SELECT *FROM PROSPAU2  where  STATUS<2  ORDER BY entolh,aa_prospau,seira;
'epil.Enabled = select prospau2.*,ximitech.*  from prospau2 inner join ximitech on prospau2.kod=ximitech.kod  where prospau2.epil>0 and prospau2.epil<100  and status<2 order by prospau2.epil,prospau2.seira;

mEpilogh = 0
mBaros_pp = 0
If all_prospau.RecordCount = 0 Then Exit Sub

Screen.MousePointer = 11

all_prospau.MoveFirst
Do While Not all_prospau.eof
   If all_prospau("seira") = 1 Then
      mEpilogh = all_prospau("epil")
      mBaros_pp = all_prospau("baros_pani")
   Else
      all_prospau.EDIT
      all_prospau("epil") = mEpilogh
      all_prospau("baros_pani") = mBaros_pp
      all_prospau.update
   End If
   all_prospau.MoveNext
Loop










m_error = " "
err_msg(1) = mL_CapRes(862) '" υπάρχουν 2 ίδια χημικά στην ίδια συνταγή"
err_msg(2) = mL_CapRes(863) '" οι ποσότητες των χρωμάτων δεν αρκούν"
err_msg(3) = mL_CapRes(864) '"  έχω υπέρβαση μπάνιου"
err_msg(4) = mL_CapRes(865) '"έχω υπερχείλιση ποτηριού "
err_msg(5) = mL_CapRes(866) ' "το μπουκάλι λείπει από το τραπέζι"
err_msg(6) = mL_CapRes(867) '"δεν υπάρχει κάν το χρώμα στο αρχείο"
err_msg(7) = mL_CapRes(868) '" δεν έχει συγκέντρωση το χρώμα"
err_msg(8) = mL_CapRes(869) '"δεν αποθηκευτηκε ο κωδικός του χρώματος στην συνταγή"

'Set MyDB1 = Workspaces(0).OpenDatabase("c:\TALOS", False, False, "dBASE IV;")
'sql = "UPDATE PROSPAU2 SET STATUS=0 WHERE STATUS=1;"
'MyDB1.Execute sql
'MyDB1.Close



Set mydb = Workspaces(0).OpenDatabase("c:\TALOS\recipies.mdb")
SQL = "delete apor_syntages.* from apor_syntages"
mydb.Execute SQL
mydb.Close



Set apor_SYNT = apor_syntages.Recordset
'Set e = epil.Recordset

'μηδενισμό πίνακα
For k = 1 To 100: For L = 1 To 30: For m = 1 To 5: x(k, L, m) = -999: Next: Next: Next
For k = 1 To 100: a_Err(k) = "": Next
Screen.MousePointer = 1


'epil.Refresh
If EPIL.RecordCount = 0 Then
   Screen.MousePointer = 1
   MsgBox mL_Res(468) '"Δεν έχουν επιλεγεί συνταγές"
   Exit Sub
End If


'φορτώνω τις συνταγές σε πίνακα
flag_last_syntagh = "000000"
EPIL.MoveFirst
MaxEpil = 0
Do While Not EPIL.eof
  synt = EPIL("epil")
  m_error = " "
  ' υπερχείλιση
 If flag_last_syntagh <> EPIL("entolh") + EPIL("aa_prospau") Then
    If IsNull(EPIL("baros_pani") * EPIL("SXESH_MPAN")) Then
      a_Err(synt) = 3    ' ??????
    Else
        If EPIL("baros_pani") * EPIL("SXESH_MPAN") / 1000 > YperXeilish Then
          a_Err(synt) = 4
          MsgBox "469-471." + mL_CapRes(469) + EPIL("entolh") + mL_CapRes(470) + EPIL("aa_prospau") + mL_CapRes(471)
          '469 "Δεν θα εκτελεστεί η συνταγή "
          '470 " προσπάθεια "   nattaporn@penn-asia.com
          '471 " γιατί θα υπερχειλίσει το ποτήρι"
          
          DUM = add_apor(EPIL("entolh"), EPIL("aa_prospau"), err_msg(4))
          flag_last_syntagh = EPIL("entolh") + EPIL("aa_prospau")
        End If
    End If
 End If
  
  
  '
  If IsNull(EPIL("prospau2.KOD")) Or EPIL("epil") = 0 Or EPIL("seira") = 0 Then
       a_Err(synt) = a_Err(synt) + "8"
       MsgBox "472-473." + mL_CapRes(472) + EPIL("entolh") + mL_CapRes(473) + EPIL("aa_prospau")
       ' 472  "Δεν καταχωρήθηκε το χρώμα στην συνταγή "
       ' 473  " προσπάθεια "
       
       DUM = add_apor(EPIL("entolh"), EPIL("aa_prospau"), err_msg(8))
  Else
        x(EPIL("epil"), EPIL("seira"), 1) = EPIL("prospau2.KOD")
       
        If EPIL("XIMITECH.KOD") <> EPIL("prospau2.KOD") Then
           a_Err(synt) = a_Err(synt) + "6"
           MsgBox "474-476." + mL_CapRes(474) + EPIL("prospau2.kod") + mL_CapRes(475) + EPIL("entolh") + mL_CapRes(476) + EPIL("aa_prospau")
           
           ' 474 "Δεν υπάρχει στο αρχείο το χρώμα "
           ' 475 " στην συνταγή "
           ' 476 " προσπάθεια "
           
           DUM = add_apor(EPIL("entolh"), EPIL("aa_prospau"), err_msg(6) + " code:" + EPIL("prospau2.kod"))
        Else
           If IsNull(EPIL("morfh")) Then
              a_Err(synt) = a_Err(synt) + "7"
              MsgBox "477,475." + mL_CapRes(477) + EPIL("prospau2.kod") + "-" + EPIL("perigrafh") + mL_CapRes(475) + EPIL("entolh") + " - " + EPIL("aa_prospau")
              '477 "Δεν έχει συγκέντρωση το χρώμα "
              DUM = add_apor(EPIL("entolh"), EPIL("aa_prospau"), err_msg(6) + EPIL("perigrafh"))
           End If
           
           If IsNull(EPIL("addr_prot")) Then
                   m_error = "5"
           Else
                 If EPIL("addr_prot") = 0 Then
                     m_error = "5"
                 End If
           End If
           
           If m_error = "5" Then
              a_Err(synt) = a_Err(synt) + "5"
              MsgBox "478,475." + mL_CapRes(478) + EPIL("prospau2.kod") + "-" + EPIL("perigrafh") + mL_CapRes(475) + EPIL("entolh") + "-" + EPIL("aa_prospau")
              '478 "Δεν υπάρχει στο τραπέζι το χρώμα "
              DUM = add_apor(EPIL("entolh"), EPIL("aa_prospau"), EPIL("perigrafh") + ":" + err_msg(5))
           End If
        
        End If
  End If
  
  
  x(EPIL("epil"), EPIL("seira"), 2) = IIf(IsNull(EPIL("EK")), 0, EPIL("EK"))
  x(EPIL("epil"), EPIL("seira"), 3) = IIf(IsNull(EPIL("GL")), 0, EPIL("GL"))
  x(EPIL("epil"), EPIL("seira"), 4) = IIf(IsNull(EPIL("entolh")), 0, EPIL("ENTOLH"))
  x(EPIL("epil"), EPIL("seira"), 5) = IIf(IsNull(EPIL("AA_PROSPAU")), 0, EPIL("AA_PROSPAU"))
  MaxEpil = IIf(MaxEpil > EPIL("epil"), MaxEpil, EPIL("epil"))
  
  EPIL.MoveNext
Loop

'ελέγχω για 2 ίδια χρώματα στην ίδια συνταγή
a = check_idia(x, a_Err, MaxEpil)

' ενημερώνω τον πίνακα με τις απαιτούμενες ποσότητες
a = Update_Gr(x, a_Err, MaxEpil)

apor_syntages.Refresh
DBGrid3.Refresh
DBGrid3.Visible = True

If EXO_SYNTAGES_GIA_SERVIRISMA = 0 Then
   Screen.MousePointer = 1
   MsgBox mL_Res(479) '479 "Δεν υπάρχουν συνταγές έτοιμες για σερβίρισμα"
   Exit Sub
End If


 ch_comdlg EXO_SYNTAGES_GIA_SERVIRISMA
 
' Set Mycom = Workspaces(0).OpenDatabase(GetWindowsDir() + "\SYSTEM", False, False, "dBASE IV;")
 ' Set mycomset = Mycom.OpenRecordset("COMDLG16.DLL")
  '  mycomset.Edit
'    mycomset("k") = mycomset("k") + EXO_SYNTAGES_GIA_SERVIRISMA
 '   mycomset.update
'  mycomset.Close






If neo_x > 0 And dokimastikh_syntagh = -1 Then
   ans = MsgBox("480-482." + mL_CapRes(480) + str(neo_x) + mL_CapRes(481) + Chr(13) + mL_CapRes(482), vbYesNo)
   ' 480 "Υπάρχουν "
   ' 481 " μπουκάλια αχρησιμοποίητα για .. ώρες."
   ' 482 "Να γίνει δοκιμαστική συνταγή;"
    
    
    
    If ans = vbYes Then
       AXRHSIMOPOIHTA
       'frmSYNT.Show
       'Exit Sub
    End If
End If
ans = vbYes ' MsgBox(mL_Res(483), vbYesNo)
    ' ("Προχωρώ στην εκτέλεση των συνταγών"
    If ans = vbYes Then
      If Val(Label5.Caption) > 10 Then
         Screen.MousePointer = 1
         MsgBox mL_Res(484) ' "Απομάκρυνε τό βάρος από την ζυγαριά.Μετά την απομάκρυνση του βάρους πάτησε ΟΚ. "
         MilSec 2000
         If Val(Label5.Caption) > 10 Then
             Screen.MousePointer = 1
             MsgBox "485,487." + mL_CapRes(485) + Chr(13) + mL_CapRes(487)
             '485 "Απομάκρυνε τό βάρος από την ζυγαριά.
             '487 "Μετά την απομάκρυνση του βάρους ξαναζητήστε ΕΚΤΕΛΕΣΗ ΣΥΝΤΑΓΩΝ. "
             Exit Sub
         End If
      End If
      
      
'If Balance_Type = "ADAM" Then
'      frmSYNTadam.Show  ' syntages
'Else
      frmSYNT.Show  ' syntages
        mUndelete_SYNTAGH = 0
'End If
      
      
      Screen.MousePointer = 1
            Unload Me

     End If
'MsgBox "Απομάκρυνε τό βάρος από την ζυγαριά.Μετά την απομάκρυνση του βάρους πάτησε ΟΚ. "

'Command3.Enabled = True


End Sub


Private Sub Command4_Click()
'start weighing
If prozygismena = 0 Then
      Command4.Visible = False
      If Mhdenismos_Zygoy(0) <> 1 Then
         yparxei_zygos = 0
         Label3.Enabled = False
         Label4.Enabled = False
         Label5.Enabled = False
         Command2.Caption = mL_CapRes(870) '"Μηδενισμός"
         
         
      Else
         yparxei_zygos = 1
      End If
      MsgBox mL_Res(467), , "Talos"
             ' 467 "Παρακαλώ τοποθετήστε τον κύλινδρο στην Ζυγαριά !"
      
      
      G_Balance_Digits MSComm1, 2
      Timer1_Timer
      Label5.Caption = ""

   End If
End Sub

Private Sub Command5_Click()
  Dim a, b, response
  Screen.MousePointer = 1
On Error Resume Next
  
  a = SQLPROSPAU.Recordset.RecordCount
  
  If IsNull(a) Or a = 0 Then
        Exit Sub
   Else
        If SQLPROSPAU.Recordset.AbsolutePosition <> -1 Then
         b = SQLPROSPAU.Recordset("ENTOLH")
         c = SQLPROSPAU.Recordset("AA_PROSPAU")
         
         response = MsgBox(mL_Res(488) + b + "-" + c, vbYesNo) ' +  + SQLPROSPAU.Recordset("aa_prospau")
         ' 485 "Να διαγραφεί η εντολή "
         If response = vbYes Then
         
            Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
            sql_del = "DELETE DISTINCTROW PROSPAU2.* FROM PROSPAU2 WHERE ENTOLH='" + b + "' AND AA_PROSPAU='" + c + "';"
            mydb.Execute sql_del
            mydb.Close
         End If
         SQLPROSPAU.Recordset.MoveLast
         SQLPROSPAU.Refresh
         DBGrid1.Refresh
    End If
  End If
On Error GoTo 0
End Sub





Private Sub DBGrid1_Click()
'If Diplo_Pathma = 0 Then
  Diplo_Pathma = 1
  Screen.MousePointer = 11
'End If
'Else
 ' Screen.MousePointer = 0
  'shmadema
  'Diplo_Pathma = 0
'End If
 
End Sub

Private Sub shmadema()
'====================================================================
Dim e, p, a As Variant, DIAGRAFEIS As Integer
Dim M_UPDATE, s, baros, baros1, baros2, mcount
Dim Mydbase, MBAROS

 
 
Set s = SQLPROSPAU.Recordset

If prozygismena = 0 And (IsNull(s("EPIL")) Or s("EPIL") = 0) Then
   If Val(Label3.Caption) < 100000 Then
      MsgBox mL_Res(486)
      ' "Δεν έχεις βάλει τον κύλινδρο "
      
      Exit Sub
   End If
End If

Command1.Enabled = False
'Set p = PROSPAU.Recordset
'PROSPAU=select *from prospau2  where seira=1  and status<2




'Set s = SQLPROSPAU.Recordset
    On Error Resume Next
 M_UPDATE = 0
If s.RecordCount > 0 Then
     If IsNull(s("EPIL")) Or s("EPIL") = 0 Then
        If ayjon >= ar_real_pothr Then
            MsgBox "Too much recipies.You exceed the number of pots"
            Exit Sub
         Else
                ayjon = ayjon + 1  'σημάδεψε συνταγή
          End If
          M_UPDATE = 1
          Label8.Caption = s("entolh") + "-" + s("aa_prospau") + "=" + str(ayjon)
          
     Else
          DIAGRAFEIS = s("EPIL") ' ξεσημάδεψε συνταγή
          M_UPDATE = 2
     End If
     e = s("entolh")
     a = s("aa_prospau")
     
     If prozygismena = -1 Then
         Set YourDB = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
         Set your = YourDB.OpenRecordset("joblist")
         your.Index = "entolh"
         your.Seek "=", e  ', A
         
         If your.eof Then
            baros = Default_baros_panioy
         Else
           If IsNull(your("baros_pani")) Then
                       baros = Default_baros_panioy
           Else
              baros = your("baros_pani")
           End If
         End If
         
     Else
       If M_UPDATE <> 2 Then
             Label5.BackColor = Portokali


mcount = 0
9987
             mcount = mcount + 1
             baros1 = Zygis0(0) - mTara
             Timer1_Timer
             MilSec 1000
             baros2 = Zygis0(0) - mTara
             If mcount > 5 Then baros1 = 0
             If Abs(baros2 - baros1) > 5 Then GoTo 9987
             baros = baros1
             ' enhmerono ta labels
             Label3.Caption = baros + mTara
             Label5.Caption = baros
             
             
             
        End If
          If baros < 0 Then
             baros = 0
          End If
         Label5.BackColor = old_C
     End If
    
     'PROSPAU.Recordset.MoveFirst
Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)

'' radaeli
Set Job = mydb.OpenRecordset("joblist")
    Job.Index = "entolh"
    Job.Seek "=", e
   
   pani = "φουφουτος"
If Job.NoMatch Then
Else
   If IsNull(Job("EIDOS_PANI")) Then
   Else
      pani = Job("EIDOS_PANI")
   End If
End If
Job.Close
Set pania = mydb.OpenRecordset("PANIA")

pania.Index = "PANI"
pania.Seek "=", pani
FYRA = 0
If pania.NoMatch Then
Else
   If IsNull(pania("FYRA")) Then
   
   Else
      FYRA = pania("FYRA")
   End If
End If
MBAROS = baros * (100 + FYRA) / 100

' Set rr = OpenRecordset("pania")

SQL = "update prospau2 set baros_pani=" + LTrim(str(MBAROS)) + ",epil=" + LTrim(str(ayjon)) + " where entolh='" + e + "' and aa_prospau='" + a + "';"

On Error GoTo 0
mydb.Execute SQL

     
     
End If
 
 ' μόλις διαγράψω το χ, τα μεγαλύτερα του χ γίνονται ν=ν-1
 
 If M_UPDATE = 2 Then
   
     Set mydb = Workspaces(0).OpenDatabase("C:\TALOS\RECIPIES.MDB", False, False)
     SQL = "update prospau2 set epil=epil-1  where status<2 and epil>" + LTrim(str(DIAGRAFEIS)) + ";"
     mydb.Execute SQL
   
   
     SQL = "update prospau2 set epil=0,baros_pani=0  where entolh='" + e + "' and aa_prospau='" + a + "';"
     mydb.Execute SQL
    ayjon = ayjon - 1
 End If
'
 
 '
 If Check1 Then
     n = SQLPROSPAU.Recordset.AbsolutePosition
'   SQLPROSPAU.Recordset.Bookmark = SQLPROSPAU.Recordset.Bookmark
   SQLPROSPAU.Refresh
   SQLPROSPAU.Recordset.Move n
End If
 DBGrid1.Refresh
 On Error GoTo 0

Command1.Enabled = True

End Sub

Private Sub DBGrid1_DblClick()
'shmadema
End Sub

Private Sub DBGrid1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'  WAIT.Caption = "We select a recipy by double click"
'  "Με διπλό κλίκ επιλέγουμε ή ακυρώνουμε την συνταγή"
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Diplo_Pathma = 1 Then
   'Screen.MousePointer = 0
 'If SQLPROSPAU.RecordSource = "SELECT PROSPAU2.*,JOBLIST.* FROM PROSPAU2 INNER JOIN JOBLIST ON PROSPAU2.ENTOLH=JOBLIST.ENTOLH where PROSPAU2.SEIRA=1 AND PROSPAU2.STATUS<2 ORDER BY JOBLIST.kod_pel,JOBLIST.APOXRVSH;" Then
 'Else
 '   SQLPROSPAU.RecordSource = "SELECT PROSPAU2.*,JOBLIST.* FROM PROSPAU2 INNER JOIN JOBLIST ON PROSPAU2.ENTOLH=JOBLIST.ENTOLH where PROSPAU2.SEIRA=1 AND PROSPAU2.STATUS<2 ORDER BY JOBLIST.kod_pel,JOBLIST.APOXRVSH;"
 'End If
    shmadema
   
    
    Diplo_Pathma = 0
    Screen.MousePointer = 0
End If

End Sub

Private Sub Form_DblClick()
         Dim a
     ' CommonDialog1.ShowColor
     
      'A = CommonDialog1.Color
      
      
      
      a = save_colors(Me, Me.BackColor)
'      A = 1
      
End Sub

Private Sub Form_Load()
  mtime = GetCurrentTime()
  yparxei_zygos = 1
  zygizo = 0
 old_C = Label5.BackColor
neo_x = 0
'Me.BackColor =
a = find_colors(Me)
MDIForm1.Arxeia(10).Visible = False
MDIForm1.Syntages(20).Visible = False
MDIForm1.bohuhtika(30).Visible = False
MDIForm1.ejodos(40).Visible = False
mTara = 0
EXO_SYNTAGES_GIA_SERVIRISMA = 0
proth_fora = 0
Diplo_Pathma = 0

DT = AXRHSIMOPOIHTES_ORES * 3600

      
Command5.Caption = mL_CapRes(414) 'diagrafi
Command3.Caption = mL_CapRes(416) 'ektelesi sintagon
Command1.Caption = mL_CapRes(417) ' ejodos
Command2.Caption = mL_CapRes(418) ' TARE
C_MHDEN.Caption = mL_CapRes(419) '  ";;ΜΗΔΕΝΙΣΜΟΣ ΛΙΣΤΑΣ"


DBGrid1.Columns(0).Caption = mL_CapRes(430) '   "ar.Eργασία"
DBGrid1.Columns(1).Caption = mL_CapRes(431) '  prosp
DBGrid1.Columns(2).Caption = mL_CapRes(432) '   "epilogi
DBGrid1.Columns(3).Caption = mL_CapRes(433) '"Σχέση μπάνιου" '
DBGrid1.Columns(4).Caption = mL_CapRes(434)  '  απόχρωση
DBGrid1.Columns(5).Caption = mL_CapRes(435) '  Πελάτης
DBGrid1.Columns(6).Caption = mL_CapRes(436) '  Βάρος πανιου
DBGrid1.Columns(7).Caption = mL_CapRes(437) ' Ημερ.Παράδοσης
DBGrid1.Columns(8).Caption = mL_CapRes(438) ' User
Label2.Caption = mL_CapRes(405) ' ";;Ενδειξη ζυγού"
Label7.Caption = mL_CapRes(406) 'Βάρος Πανιού
Label1.Caption = mL_CapRes(407) 'Ταξινόμηση κατά"
Option1.Caption = mL_CapRes(408) ' "Πελάτη/Απόχρωση"
Option2.Caption = mL_CapRes(409) ' "Αρ.Εργασίας"
Option3.Caption = mL_CapRes(410) ' "Α/Α επιλογής"

DBGrid3.Caption = mL_CapRes(440) '"ΑΠΟΡΙΦΘΕΙΣΕΣ ΣΥΝΤΑΓΕΣ"
DBGrid3.Columns(0).Caption = mL_CapRes(441) ' "Εντολή"
DBGrid3.Columns(1).Caption = mL_CapRes(442) ' "α/α προσπ"
DBGrid3.Columns(2).Caption = mL_CapRes(443) ' "Aιτία απόρριψης


Option1.BackColor = WAIT.BackColor
Option2.BackColor = WAIT.BackColor
Option3.BackColor = WAIT.BackColor
Check1.BackColor = WAIT.BackColor
      
       DBEngine.IniPath = "C:\talos\talos.INI"
End Sub

Private Sub Form_Paint()

If proth_fora = 1 Then
   'Label4.Caption = -(mtime - GetCurrentTime()) / 1000
   Exit Sub
End If
Command1.Enabled = False
Command3.Enabled = False
Command5.Enabled = False
C_MHDEN.Enabled = False

proth_fora = 1
          
          Screen.MousePointer = 1
          ayjon = 0

If mUndelete_SYNTAGH = 0 Then ' δεν έχει κάνει έλεγχο για μισοσβησμένη συνταγή
        undelete_SYNTAGH
        mUndelete_SYNTAGH = 1
End If


 If SQLPROSPAU.Recordset.RecordCount = 0 Then
        ayjon = 0
Else
        SQLPROSPAU.Recordset.MoveLast
       If IsNull(SQLPROSPAU.Recordset("epil")) Then
         ayjon = 0
       Else
        ayjon = SQLPROSPAU.Recordset("epil")
       End If
End If

 SQLPROSPAU.RecordSource = "SELECT prospau2.* FROM PROSPAU2   where SEIRA=1 AND prospau2.STATUS<2  ORDER BY prospau2.entolh;"
 'SQLPROSPAU.Refresh


      
      
      
      
      
      
   'If prozygismena = 0 Then
   '   If Mhdenismos_Zygoy(0) <> 1 Then
   '      yparxei_zygos = 0
   '      Label3.Enabled = False
   '      Label4.Enabled = False
   '      Label5.Enabled = False
   '      Command2.Caption = mL_CapRes(870) '"Μηδενισμός"
   '   Else
   '      yparxei_zygos = 1
   '   End If
   'End If
On Error GoTo 0
Command1.Enabled = True
Command3.Enabled = True
Command5.Enabled = True
C_MHDEN.Enabled = True
Label4.Caption = -(mtime - GetCurrentTime()) / 1000

End Sub

Private Sub Form_Unload(Cancel As Integer)

'MDIForm1.arxeia.Visible = True
'MDIForm1.Syntages.Visible = True
'MDIForm1.Bohuhtika.Visible = True
'MDIForm1.ejodos.Visible = True






End Sub

Private Sub Option1_Click()
  SQLPROSPAU.RecordSource = "SELECT PROSPAU2.* FROM PROSPAU2  where PROSPAU2.SEIRA=1 AND PROSPAU2.STATUS<2 ORDER BY kod_pel,APOXRVSH;"
  SQLPROSPAU.Refresh
End Sub


Private Sub Option2_Click()
  SQLPROSPAU.RecordSource = "SELECT PROSPAU2.* FROM PROSPAU2  where SEIRA=1 AND PROSPAU2.STATUS<2 ORDER BY ENTOLH,AA_PROSPAU;"
  SQLPROSPAU.Refresh
End Sub


Private Sub Option3_Click()
  SQLPROSPAU.RecordSource = "SELECT PROSPAU2.* FROM PROSPAU2  where PROSPAU2.SEIRA=1 AND PROSPAU2.STATUS<2 ORDER BY EPIL;"
  SQLPROSPAU.Refresh
  SQLPROSPAU.Recordset.MoveLast
End Sub


Private Sub Timer1_Timer()
         MilSec 500
         mz = Zygis0(0)
         Label3.Caption = mz
         Label5.Caption = mz - mTara
End Sub

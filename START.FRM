VERSION 4.00
Begin VB.Form Start 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Talos : Διαδικασία Εναρξης"
   ClientHeight    =   6180
   ClientLeft      =   1305
   ClientTop       =   345
   ClientWidth     =   8355
   ControlBox      =   0   'False
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   162
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0080FFFF&
   Height          =   6585
   HelpContextID   =   1023
   Icon            =   "START.frx":0000
   Left            =   1245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "START.frx":0442
   ScaleHeight     =   6180
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   Top             =   0
   Width           =   8475
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   2715
      Top             =   450
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Fonts"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      HelpContextID   =   1023
      Left            =   4380
      TabIndex        =   16
      Top             =   1455
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Language"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      HelpContextID   =   1023
      Left            =   2535
      TabIndex        =   15
      Top             =   1470
      Width           =   1320
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "START.frx":AA04
      Left            =   330
      List            =   "START.frx":AA06
      TabIndex        =   14
      Top             =   1665
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2220
      Top             =   450
   End
   Begin VB.CommandButton Command1 
      Caption         =   "JOYSTICK"
      Enabled         =   0   'False
      Height          =   500
      HelpContextID   =   1023
      Left            =   3510
      TabIndex        =   13
      Top             =   465
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "EXIT"
      Enabled         =   0   'False
      Height          =   500
      HelpContextID   =   1023
      Left            =   6705
      TabIndex        =   4
      Top             =   465
      Width           =   1230
   End
   Begin VB.CommandButton Command2 
      Caption         =   "HOME"
      Enabled         =   0   'False
      Height          =   500
      HelpContextID   =   1023
      Left            =   390
      TabIndex        =   0
      Top             =   465
      Width           =   1215
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   7455
      TabIndex        =   24
      Top             =   1515
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Press F1 for Help"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   3375
      TabIndex        =   23
      Top             =   105
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   7020
      TabIndex        =   22
      Top             =   1515
      Visible         =   0   'False
      Width           =   405
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5010
      Top             =   570
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      CDTimeout       =   0
      CommPort        =   1
      CTSTimeout      =   0
      DSRTimeout      =   0
      DTREnable       =   -1  'True
      Handshaking     =   0
      InBufferSize    =   1024
      InputLen        =   0
      Interval        =   1000
      NullDiscard     =   0   'False
      OutBufferSize   =   512
      ParityReplace   =   "?"
      RThreshold      =   0
      RTSEnable       =   0   'False
      Settings        =   "9600,n,8,1"
      SThreshold      =   0
   End
   Begin VB.Label clue6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   4950
      TabIndex        =   21
      Top             =   105
      Width           =   3000
   End
   Begin VB.Label msg1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   6075
      TabIndex        =   20
      Top             =   1515
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   6555
      TabIndex        =   19
      Top             =   1515
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   3555
      TabIndex        =   18
      Top             =   4470
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   405
      TabIndex        =   17
      Top             =   105
      Width           =   2940
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   " 5. Μετά το πέρας της διαδικασίας το Terminal πρέπει να παραμείνει ανοικτό σε σμίκρυνση."
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   90
      TabIndex        =   12
      Top             =   3840
      Width           =   8295
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait ..."
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   165
      TabIndex        =   11
      Top             =   4500
      Width           =   3345
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   " 4. Μετά την Εκτέλεση του Setup , να ακούγεται ο ρυθμικός ήχος της παλμογεννήτριας."
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   90
      TabIndex        =   10
      Top             =   3510
      Width           =   8295
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   " ! Αν δεν τηρούνται οι παραπάνω προϋποθέσεις ακυρώστε και χρησιμοποιήστε το Joystick."
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   90
      TabIndex        =   9
      Top             =   4185
      Width           =   8295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   165
      TabIndex        =   8
      Top             =   2115
      Width           =   7980
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   " 3. Δεν πρέπει να υπάρχει τίποτα στο τάσι της Ζυγαριάς."
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   90
      TabIndex        =   7
      Top             =   3195
      Width           =   8295
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   " 2. Δεν πρέπει να υπάρχουν μπουκάλια στις κεφαλές δοσομέτρησης ."
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   90
      TabIndex        =   6
      Top             =   2880
      Width           =   8295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   " 1. H Αρπάγη πρέπει να βρίσκεται σε ελεύθερο χώρο και ΔΕΝ ΠΡΕΠΕΙ ΝΑ ΚΡΑΤΑΕΙ ΤΙΠΟΤΑ!"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   90
      TabIndex        =   5
      Top             =   2565
      Width           =   8295
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   225
      TabIndex        =   3
      Top             =   5715
      Width           =   7455
   End
   Begin VB.Label label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "  "
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   225
      TabIndex        =   2
      Top             =   4890
      Width           =   3765
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   240
      TabIndex        =   1
      Top             =   5310
      Width           =   4260
   End
End
Attribute VB_Name = "Start"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub daily_report_show()
Dim db As Database, r  As Recordset, t_quer As Recordset, t_dat   As Recordset, t    As Recordset
On Error Resume Next
 Set db = OpenDatabase("c:\talos\skon_tb.mdb")
SQL = "DELETE DISTINCTROW Syntages_datum.* FROM Syntages_datum;"
db.Execute SQL
SQL = "DELETE DISTINCTROW Syn_Comments.* FROM Syn_Comments;"
db.Execute SQL

Set r = db.OpenRecordset("Syntages")
Set t_dat = db.OpenRecordset("Syntages_datum")
r.MoveLast
Datum = DateValue(r("Datum"))
Do Until r.BOF
If Datum = DateValue(r("Datum")) Then GoSub Upd_Rec Else Exit Do
r.MovePrevious
Loop


Set t = db.OpenRecordset("Syn_Comments")
Set r = db.OpenRecordset("Q_Synt_Stat_Datum")
Do Until r.EOF
  I = I + 1
    com1 = "": com2 = "": com3 = "": COM4 = ""
  If r("AvgOfTryals") > 4 Then com2 = "Απροσάρμοστη παροχή βαλβίδας. "
  If r("AvgOfLevel") < 100 Then com1 = "Χαμηλή στάθμη φιάλης. "
  If r("AvgOfLasted") > 40 Then com3 = "Χαμηλή παροχή, αργή  Ζύγιση. "
  If com1 + com2 + com3 = "" Then COM4 = "Ελέγξτε βαλβίδα και φιάλη."
  
  com = com1 + com2 + com3 + COM4
  If I < 11 Then GoSub add_comment Else Exit Do
  r.MoveNext
Loop
If Err > 0 Then r.Close:  t.Close: Exit Sub
  r.Close
 t.MoveLast
 rec_num = t.RecordCount - 1
 If rec_num < 1 Then t.Close: Exit Sub
Msg = "Οι παρακάτω φιάλες εμφανίζουν πρόβλημα δοσομέτρησης !" + Chr(13) + Chr(13)
t.MoveFirst
Do Until t.EOF
Msg = Msg + t("Description") + " : " + t("Comments") + Chr(13)
t.MoveNext
Loop

ans = MsgBox(Msg + Chr(13) + "Να γίνει εκτύπωση  ;", 276, "Talos : Bottle Daily Report")
If ans = 7 Then t.Close: Exit Sub
t.MoveFirst

str_line$ = "Talos : Προβλήματα Φιαλών βάσει δοσομετρήσεων της  " + Str(DateValue(Datum))
Printer.Print str_line$
str_line$ = "       "
Printer.Print str_line$
str_line$ = "  Φιάλη                       Πρόβλημα          "
Printer.Print str_line$
str_line$ = "============================================================================="
Printer.Print str_line$
Do Until t.EOF
str_line$ = t("Description") + " : " + t("Comments")
Printer.Print str_line$
str_line$ = "  ---------------------------------------------------------------------------------------------"
Printer.Print str_line$
t.MoveNext
Loop
Printer.EndDoc
Exit Sub

Upd_Rec:
t_dat.AddNew
 t_dat("Datum") = r("Datum")
 t_dat("Description") = r("Description")
 t_dat("Asked_Q") = r("Asked_Q")
 t_dat("Actual_Q") = r("Actual_Q")
 t_dat("Level") = r("Level")
 t_dat("Tryals") = r("tryals")
 t_dat("Lasted") = r("Lasted")
 t_dat.Update
 Return

add_comment:
t.AddNew
t("Description") = r("Description")
t("Comments") = com
t.Update
Return
End Sub
Sub Send_Programs()
Label1 = "": Label2 = "": label3 = "" ': Label10 = ""
DoEvents
 Label14.Visible = True
Label10 = "1 .  " + LoadLabel(474)
cmd$ = Program1
tmp% = SendFile(1, Device_Address%, 0, cmd$, 0)
GoSub show_Lab
Label10 = "2 .  " + LoadLabel(474)
cmd$ = Program2
tmp% = SendFile(1, Device_Address%, 0, cmd$, 0)
GoSub show_Lab

Label10 = "3 .  " + LoadLabel(474)
cmd$ = Program3
tmp% = SendFile(1, Device_Address%, 0, cmd$, 0)
GoSub show_Lab

If Program4 <> "" Then
Label10 = "4 .  " + LoadLabel(474)
cmd$ = Program4
tmp% = SendFile(1, Device_Address%, 0, cmd$, 0)
GoSub show_Lab
End If
 Label14.Visible = False
Label10 = "  " + LoadLabel(473)
cmd$ = "setup:"
RobSend (cmd$)
Label13 = "1"
MilSec 1000
cmd$ = "!ERRLVL0:"
RobSend (cmd$)
       MilSec (1000)
       VarNew
       MilSec (1000)
       cmd$ = "!" + Ready_Bit + ":"
      RobSend (cmd$)
 Label2 = " " + LoadLabel(464)
 MilSec (100)
 ValveReset
Exit Sub

show_Lab:
Num = 3000
 Label14.Width = 1: Label14.Caption = "0%"
For I = 1 To 20
    MilSec Num / 20
    Label14.Caption = Format(I / 20, "##%")
    Label14.Width = 5000 * I / 20
Next I
Return

End Sub
Function porta() As String
    If IsOSLoaded(Device_Address%) = 0 Then porta = "κ": Exit Function
    Result% = Request_Status(Device_Address%)
    Set_Pointer Device_Address%, INPUT_STATUS
    Read_Status Device_Address%, status_high%, status_low%, status&
     If (status_low% And 2 ^ 1) = 0 Then
          porta = "κ"
      Else
          porta = "α"
     End If
End Function
Function Pause() As Integer
    Result% = Request_Status(Device_Address%)
    Set_Pointer Device_Address%, INPUT_STATUS
    Read_Status Device_Address%, status_high%, status_low%, status&
     If (status_low% And 2 ^ 3) = 0 Then
          Pause = 1
      Else
          Pause = 0
     End If
End Function

Private Sub Combo1_Click()
Dim MyDB As Database, r As Recordset, t As Recordset
  Screen.MousePointer = 11
  Set MyDB = OpenDatabase("c:\talos\coord_tb.mdb")
  Set r = MyDB.OpenRecordset("labels")
  Set t = MyDB.OpenRecordset("Language")
     t.MoveFirst
 Do Until r.EOF
     r.Edit
     GoSub find_Field
     r.Update
     r.MoveNext
     t.MoveNext
     Loop
     t.Close
     
     
 Set MyDB = OpenDatabase("c:\talos\coord_tb.mdb")
   Set t = MyDB.OpenRecordset("Language")
Do While Not t.EOF
    t.Edit
    t("Label") = t(Combo1)
    t.Update
    t.MoveNext
    't.Update
Loop
t.Close
     
     
     
     
     
     
     Set r = MyDB.OpenRecordset("Βαλβιδες")
    Do Until r.EOF
      If r("valve_kod") = 15 Then
      r.Edit
        r("Valve_off") = Combo1
        r.Update
      Exit Do
    End If
    r.MoveNext
    Loop
    r.Close
   Command5.Enabled = True
   Command4.Enabled = True
   Command4.SetFocus
   Combo1.Visible = False
   Me.Caption = "Talos : " + LoadLabel(491)
Label7 = LoadLabel(449)
Label4 = LoadLabel(450)
Label5 = LoadLabel(451)
Label6 = LoadLabel(452)
Label9 = LoadLabel(453)
Label11 = LoadLabel(454)
Label8 = LoadLabel(455)
 Label2 = " " + LoadLabel(464)
 Label10 = "  " + LoadLabel(473)
Command4.Enabled = True
Screen.MousePointer = 1






Exit Sub

find_Field:
     Select Case Combo1
     Case "Greek"
      If IsNull(t("Greek")) = 0 Then r("Label") = t("Greek")
      Case "English"
       If IsNull(t("English")) = 0 Then r("Label") = t("English")
      Case "Turkey"
        If IsNull(t("Turkey")) = 0 Then r("Label") = t("Turkey")
      Case "Indian"
       If IsNull(t("Indian")) = 0 Then r("Label") = t("Indian")
      Case "Italian"
       If IsNull(t("Italian")) = 0 Then r("Label") = t("Italian")
       Case "French"
       If IsNull(t("French")) = 0 Then r("Label") = t("French")
       Case "German"
       If IsNull(t("German")) = 0 Then r("Label") = t("German")
       Case "Arabian"
       If IsNull(t("Arabian")) = 0 Then r("Label") = t("Arabian")
       Case "Chinese"
       If IsNull(t("Chinese")) = 0 Then r("Label") = t("Chinese")
     End Select
       Return

End Sub
Private Sub Butt_Enable(YesNo As Boolean)
Command1.Enabled = YesNo
Command4.Enabled = YesNo
Command5.Enabled = YesNo
Command3.Enabled = YesNo
Command2.Enabled = YesNo
If YesNo = True Then Screen.MousePointer = 1 Else Screen.MousePointer = 11
End Sub
Private Sub Command1_Click()
On Error GoTo err_dec
     Screen.MousePointer = 11
     If Label13 = "" Then Send_Programs
     If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    Butt_Enable (False)
retval = Shell("C:\talos\Joystick.EXE", 1)
    Me.WindowState = 1
   
 If IsEmpty(retval) Then MsgBox LoadLabel(461), , "Talos"
Screen.MousePointer = 1
While GetModuleUsage%(retval) > 0
        DoEvents
        If Me.WindowState <> 1 Then Me.WindowState = 1
Wend
If MSComm1.PortOpen = False Then MSComm1.PortOpen = True
Butt_Enable (True)
  Me.WindowState = 2
Exit Sub
err_dec:
MsgBox "Error : " + Err.Description, , "Talos"
Resume Next
End Sub
Private Sub Command2_Click()
Dim dbsJet As Database
Dim r As Table, Msg As String, counter As Long

On Error GoTo err_check
Butt_Enable (False)
Timer1.Enabled = True
Me.WindowState = 2
If system_ready = 0 Then
ans = MsgBox(LoadLabel(447) + Chr(13) + LoadLabel(487) + Chr(13) + LoadLabel(448), 19, "Talos : Press F1 for Help", "c:\talos\talos.hlp", "1003")
  If ans = 2 Then Butt_Enable (True):   Exit Sub

  If ans = 7 Then End
Label13 = ""
End If
If Label13 = "" Then
Butt_Enable (False)
Send_Programs
End If
   Label1 = " " + LoadLabel(471)
   If Check_Motor(13, 0) = 1 Or Check_Motor(14, 0) = 1 Then
          ans = MsgBox(LoadLabel(470), 19, "Talos : Press F1 for Help", "c:\talos\talos.hlp", "1003")
           If ans = 7 Then End
           If ans = 2 Then
                Butt_Enable (True)
                Exit Sub
            End If
          Label1 = " " + LoadLabel(468)
    Else
          Label1 = " " + LoadLabel(469)
    End If
                         'Ελεγχος Πίεσης Αέρα
  label3 = "Air Pressure  Checking ..."
  If FindInput(Air_Input) = 1 Then
          label3 = "Air Pressure very Low "
        ans = MsgBox(LoadLabel(125) + Chr(13) + LoadLabel(194), 275, "Talos")
         If ans = 2 Then Butt_Enable (True): Exit Sub
         If ans = 7 Then End
Else
   label3 = "Air Pressure OK"
End If
MilSec 1000
Msg = LoadLabel(465) + Chr$(13) + LoadLabel(466) + Chr$(13) + LoadLabel(467)
Style = vbYesNo + vbCritical + vbDefaultButton2
Title = LoadLabel(449)
response = MsgBox(Msg, Style, Title)
   If response = 7 Then Butt_Enable (True): Command1_Click: Exit Sub
    label3 = "Gripper Checking ..."
    label3.Refresh
    label3 = Check_Gripper()
    Label10 = " " + LoadLabel(472)
7
 If label3 = "Gripper : Problem Detected" Then GoSub Send_Message
Do
     cmd$ = "!out." + Ready_Bit + "-1:"
      RobSend (cmd$)
      MilSec (200)
Loop Until FindOutput(Ready_Bit) = 1
     cmd$ = "!k:"
    RobSend (cmd$)
    If Label15 = " " Then GoTo 7
Command1.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command3.Enabled = False
     rethom = Shell("C:\talos\chek_hom.EXE", 1)
     
   While GetModuleUsage%(rethom) > 0: DoEvents: Wend
   
       If Check_Motor(5, 1) = 1 Then MsgBox LoadLabel(463), , "Talos : Press F1 for Help", "c:\talos\talos.hlp", "1003": End: Exit Sub
                 
         Msg = LoadLabel(475)
         response = MsgBox(Msg, Style, "  HOME  CHECK POINT", Help, Ctxt)
         Butt_Enable (True)
         
     If response = 7 Then Exit Sub
    
         MsgBox LoadLabel(462), , "Talos"
        Label17 = "1"
If Label17 = "0" Then
    ans = MsgBox(LoadLabel(489), 276, "Talos")
     If ans = 7 Then Exit Sub
End If
 If clue6 = " Scale is not responding " Then
            ans = MsgBox(LoadLabel(319) + Chr(13) + LoadLabel(422), 20, "Talos")
            If ans = 7 Then Exit Sub
  End If
  
     cmd$ = "!k:"
   RobSend (cmd$)
   MilSec 200
  cmd$ = "!" + Motor_Output_off + ":"
  RobSend (cmd$)
  
  If Recipy_Days > 0 Then daily_report_show
  
   Set MyDB = OpenDatabase("c:\talos\coord_tb.mdb")
    Set r = MyDB.OpenRecordset("mem3", dbOpenTable)
    r.Edit
       r("HOME") = "HOME=OK"
       r("XRHSTHS") = "NO BUSY"
       r.Update
     r.Close
     
     Msg = "": counter = 0
      Set MyDB = OpenDatabase("c:\TALOS\recipies.mdb")
      Set Col = MyDB.OpenRecordset("XIMITECH")
            Do Until Col.EOF
                If Col("ypol_prot") < 80 Then
                        counter = counter + 1
                        Msg = Msg + Col("PERIGRAFH") + Chr(13)
               End If
               If counter > 15 Then
                        Msg = Msg + "  and more ...." + Chr(13)
                        Exit Do
            End If
          Col.MoveNext
          Loop
          
          If Msg <> "" Then MsgBox "Τα μπουκάλια :" + Chr(13) + Msg + "έχουν πολύ λίγο περιεχόμενο και κοντεύουν να αδειάσουν . Ανανεώστε τα ...", , "Talos"

 End
Exit Sub
Send_Message:
        ans = MsgBox(msg1 + "  " + LoadLabel(482) + Chr(13) + "   " + LoadLabel(483) + Chr(13) + "   " + LoadLabel(484) + Chr(13) + "   " + LoadLabel(486) + "   " + Chr(13) + LoadLabel(485) + Chr(13) + LoadLabel(448), 19, "Talos")
         If ans = 7 Then End
         If ans = 2 Then Screen.MousePointer = 1: Command2.Enabled = True: Exit Sub
    Return
err_check:
Err_Ck (Err)
      Resume Next
End Sub

Private Sub Command3_Click()
If Label17 = "0" Then
    ans = MsgBox(LoadLabel(489), 276, "Talos")
     If ans = 7 Then Exit Sub
End If
 If clue6 = " Scale is not responding " Then
            ans = MsgBox(LoadLabel(319) + Chr(13) + LoadLabel(422), 20, "Talos")
            If ans = 7 Then Exit Sub
  End If
    
    cmd$ = "!k:"
   RobSend (cmd$)
   MilSec 200
  cmd$ = "!" + Motor_Output_off + ":"
  RobSend (cmd$)

 End
End Sub


Private Sub Command4_Click()
  Dim MyDB As Database, r As Recordset
   Combo1.Visible = True
   Combo1.Clear
    Set MyDB = OpenDatabase("c:\talos\coord_tb.mdb")
    Set r = MyDB.OpenRecordset("Language")
   For I = 1 To r.Fields.Count - 1
    Combo1.AddItem r.Fields(I).Name
   Next I
   r.Close
     Set r = MyDB.OpenRecordset("Βαλβιδες")
    Do Until r.EOF
      If r("valve_kod") = 15 Then
       Combo1 = r("Valve_off")
       Exit Do
    End If
    r.MoveNext
    Loop
    r.Close
  Combo1.SetFocus
   Command4.Enabled = False
   Command5.Enabled = False
End Sub

Private Sub Command5_Click()

On Error Resume Next
  Dim MyDB As Database, r As Recordset
 CommonDialog1.Flags = cdlCFBoth
 CommonDialog1.ShowFont
 

    Set MyDB = OpenDatabase("c:\talos\coord_tb.mdb")
    Set r = MyDB.OpenRecordset("Βαλβιδες")
    Do Until r.EOF
      If r("valve_kod") = 15 Then
      r.Edit
        r("Valve_on") = CommonDialog1.FontName
        r.Update
      Exit Do
    End If
    r.MoveNext
    Loop
    r.Close
    
        For I = 0 To Start.Controls.Count - 1
        If Start.Controls(I).Name <> "Command4" And Start.Controls(I).Name <> "Command5" Then
             Start.Controls(I).Font = CommonDialog1.FontName
        End If
         Next

End Sub

Function zygisi4(Zygis_Kind As Variant)
On Error GoTo er_det
   Dim counter
 
   counter = 0: tot_counter = 0
 If system_ready = 0 Then zygisi4 = "System not Working": Exit Function
  mStart = GetCurrentTime()
  
   MSComm1.Output = "UPD 20" + Chr$(13) + Chr(10)
   MSComm1.Output = "SIR" + Chr$(13) + Chr(10)
  
   If Zygis_Kind = "OK" Then Label1 = "Scaling ..."
    Do
    If GetCurrentTime() - mStart > 10001 Then zygisi4 = " ": Exit Function
     MSComm1.InBufferCount = 0
      FromModem$ = ""
   'On Error GoTo 0
  
       MilSec (20)
      dummy = DoEvents()
       If MSComm1.InBufferCount Then
           buf = MSComm1.InBufferCount
          FromModem$ = FromModem$ + MSComm1.Input
           If InStr(FromModem$, "OL") > 0 Then
                'MsgBox "Scale Overload ...", , "Talos"
                zygisi4 = "OL"
                Exit Do
            End If
           If InStr(FromModem$, "UL") > 0 Then
                'MsgBox "Scale Overload ...", , "Talos"
                zygisi4 = "UL"
                Exit Do
            End If
            If GetCurrentTime() - mStart > 5000 And Zygis_Kind = "OK" Then Zygis_Kind = 0: tot_counter = 0
         ' asq = FromModem$
          ' asw = MSComm1.Input
          If Zygis_Kind = "OK" Then
          If InStr(FromModem$, Chr$(13)) And (InStr(FromModem$, ".") > 5 Or InStr(6, FromModem$, ".") > 5) Then
               ' If InStr(FromModem$, "OK") > 0 Then
                    Exit Do
               ' End If
         End If
        Else
          If InStr(FromModem$, Chr$(13)) Then
                If (InStr(FromModem$, ".") Or InStr(6, FromModem$, ".")) And InStr(6, FromModem$, "g") > 10 Then
                       Exit Do
                End If
          End If
       End If
       counter = counter + 1
       If counter >= 15 Then tot_counter = tot_counter + 1
            'Label10 = counter
          If tot_counter >= 20 And Zygis_Kind <> "OK" Then
          counter = 0: tot_counter = tot_counter + 1
               zygisi4 = " "
               Exit Function
           End If
          If counter >= 50 Then  ' 500
             counter = 0: tot_counter = tot_counter + 1
               MSComm1.InBufferCount = 0
              MSComm1.Output = Balance_Asking + Chr$(13) + Chr(10)
              End If
     End If
     'If Zygis_Kind = "OK" Then
            ' Me.Zyg_Show.Width = (GetCurrentTime() - mStart) / 1000 * 650
             ' Me.Zyg_Show.Caption = Int(((GetCurrentTime() - mStart) / 1000) + 0.5) & " Sec"
             'Me.Zyg_Show.Refresh
     ' End If
    Loop
    
      If InStr(FromModem$, ".") < 5 Then
         X = InStr(6, FromModem$, "  ")
         Y = InStr(X, FromModem$, "g")
      Else
         X = InStr(FromModem$, "  ")
         Y = InStr(FromModem$, "g")
      End If

         zygisi4 = Mid$(FromModem$, X, Y - X)
    
    
    
    
    
    '    zygisi4 = left$(FromModem$, 9)
er_ex:
Exit Function
er_det:
 zygisi4 = " "
Resume er_ex
End Function

Function zyg_decimals()
On Error Resume Next
asw = zygisi4(0)
    poin = InStr(asw, ".")
    For I = 1 To 20
       If IsNumeric(Mid(asw, poin + I, 1)) Then dec = dec + 1 Else Exit For
     Next I
 zyg_decimals = dec
 
Exit Function
End Function

Private Sub anix_zygaria()
On Error Resume Next
    If MSComm1.PortOpen = False Then MSComm1.CommPort = Balance_Port
    If MSComm1.PortOpen = False Then MSComm1.PortOpen = True
    MSComm1.Settings = Balance_Settings
    MSComm1.InputLen = 0
End Sub
Function SendFile(index_1, Device_Address%, index_2, cmd$, index_3) As Boolean
On Error Resume Next
tmp% = SendAT6400File(index_1, Device_Address%, 0, cmd$, index_2)
SendFile = True
End Function



Function Check_Gripper()
Dim asw, sfalma, counter
 sfalma = 0: counter = 0
check1:
  counter = counter + 1
 Gripper = FindInput(Gripper_input)
 asw = Gripper
If Gripper = 0 Then
        cmd$ = "!" + LTrim(Gripper_Off) + ":":    msg1 = LoadLabel(480) + Chr(13)
Else
        cmd$ = "!" + LTrim(Gripper_On) + ":": msg1 = LoadLabel(481) + Chr(13)
End If
         RobSend (cmd$)
         Start_Time = GetCurrentTime()
  Do
        Gripper = FindInput(Gripper_input)
         Cur_Time = GetCurrentTime()
         If Cur_Time - Start_Time >= 4000 Then sfalma = 1: Exit Do   '    Στα 4 sec ο Gripper πρέπει να ανοίξει ή να κλείσει
  Loop Until Gripper <> asw
  
   If sfalma = 1 Then
                Check_Gripper = "Gripper : Problem Detected"
         Exit Function
End If
If counter < 2 Then MilSec 2000: GoTo check1
'Επανάληψη
If Gripper = 1 Then
        MilSec 800
        cmd$ = "!" + LTrim(Gripper_On) + ":"
        RobSend (cmd$)       'Ασε τον ανοικτό
End If
   If sfalma = 1 Then Check_Gripper = "Gripper : Problem Detected" Else Check_Gripper = "Gripper OK"
Exit Function

End Function

Private Sub Form_Activate()
If Label1 = " Ready to HOME ..." Then Exit Sub
anix_zygaria
Screen.MousePointer = 11
Send_Programs
Preheat.Show 1
Me.HelpContextID = 1003
Label16.Visible = True
Command2.Enabled = True: Command1.Enabled = True
Command4.Enabled = True: Command5.Enabled = True: Command3.Enabled = True
Command2.SetFocus
Label1 = " Ready to HOME ..."
Screen.MousePointer = 1
End Sub

Private Sub Form_Click()
If Combo1.Visible = True Then
   Combo1.Visible = False
   Command5.Enabled = True
   Command4.Enabled = True
End If
End Sub

Private Sub Form_Load()
On Error GoTo err_find
   Dim dbsJet As Database
   Dim r As Table, t As Recordset
If App.PrevInstance Then End
system_ready = IsOSLoaded(Device_Address%)
If system_ready = 1 Then Label12 = "Operating System Loaded" Else Label12 = "Operating System Not Loaded"
If system_ready = 0 Then
ans = MsgBox(LoadLabel(447) + Chr(13) + LoadLabel(487) + Chr(13) + LoadLabel(448), 20, "Talos")
  If ans = 7 Then End
End If

 Check_User
 
FindValves
On Error Resume Next
 asw = FindFonts()
  For I = 0 To Me.Controls.Count - 1
    Me.Controls(I).Font = asw
 Next
Me.Caption = "Talos : " + LoadLabel(491)
Label7 = LoadLabel(449)
Label4 = LoadLabel(450)
Label5 = LoadLabel(451)
Label6 = LoadLabel(452)
Label9 = LoadLabel(453)
Label11 = LoadLabel(454)
Label8 = LoadLabel(455)
Screen.MousePointer = 1
Label13 = ""

If Recipy_Days > 0 Then
      Set dbsJet = OpenDatabase("c:\Talos\skon_tb.mdb")
      Set t = dbsJet.OpenRecordset("Syntages")
      Do Until t.EOF
            If DateValue(t("Datum")) < DateValue(Now()) - Recipy_Days Then t.Delete
            t.MoveNext
      Loop
      t.Close
End If
Exit Sub
err_find:
Err_Ck (Err)
Resume Next
End Sub


Private Sub Timer1_Timer()
Dim prob_found As Integer
prob_found = 0
Style = 20 + vbDefaultButton2
If Check_CutOff() <> 0 Then
    prob_found = 1
     Label15 = " "
      ans = MsgBox(LoadLabel(456) + Chr(13) + LoadLabel(457) + Chr(13) + LoadLabel(458) + Chr(13) + LoadLabel(448), Style, "Talos")
     If ans = 7 Then End
     Exit Sub
Else
     prob_found = 0
End If
'If FindInput(Pause_input) = 1 Then
 '    prob_found = 1
  '   Label15 = " "
   '   ans = MsgBox(LoadLabel(459) + Chr(13) + LoadLabel(448), Style, "Talos")
    ' If ans = 7 Then End
     'Exit Sub
'Else
     prob_found = 0
'End If
If FindInput(Door_input) = 1 Then
     prob_found = 1
     Label15 = " "
        ans = MsgBox(LoadLabel(460) + Chr(13) + LoadLabel(448), Style, "Talos")
       If ans = 7 Then End
       Exit Sub
Else
    prob_found = 0
End If
If prob_found = 1 Then Label15 = " " Else Label15 = ""
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
 zyg_dec = "3"   'zyg_decimals()
If zyg_dec = "2" Then
    zyg_display = Format(zygisi4(0), "##0.00")
     zyg_display.BackColor = Green
     clue6 = ""
ElseIf zyg_dec = "3" Then
    zyg_display = Format(zygisi4(0), "##0.000")
     zyg_display.BackColor = Green
     clue6 = ""
Else
        zyg_display = "-"
        zyg_display.BackColor = Red
        clue6 = " Scale is not responding "
End If
zyg_display.Refresh
 Exit Sub


End Sub



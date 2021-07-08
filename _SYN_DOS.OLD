VERSION 4.00
Begin VB.Form Syn_Dos 
   BackColor       =   &H00C0C000&
   Caption         =   "Talos : Δοσομέτρηση Συνταγών"
   ClientHeight    =   1410
   ClientLeft      =   5295
   ClientTop       =   1395
   ClientWidth     =   4245
   ControlBox      =   0   'False
   Height          =   1815
   Left            =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4245
   Top             =   1050
   Width           =   4365
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Cancel"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3495
      TabIndex        =   24
      Top             =   1110
      Width           =   765
   End
   Begin VB.TextBox Valve_On 
      Height          =   285
      Left            =   3960
      TabIndex        =   23
      Top             =   330
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox Valve_Off 
      Height          =   285
      Left            =   3705
      TabIndex        =   22
      Top             =   315
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox asked_q 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   30
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "1000"
      Top             =   255
      Width           =   690
   End
   Begin VB.TextBox real_q_dis 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   735
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   255
      Width           =   690
   End
   Begin VB.TextBox dif 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   255
      Width           =   615
   End
   Begin VB.TextBox Level 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   915
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   825
      Width           =   510
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   30
      TabIndex        =   1
      Text            =   "Bottle Level"
      Top             =   825
      Width           =   900
   End
   Begin VB.TextBox Target_Value 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   3300
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   255
      Width           =   945
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Target Safety"
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2880
      TabIndex        =   28
      Top             =   825
      Width           =   1035
   End
   Begin VB.Label Trgt_Safety 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   3900
      TabIndex        =   27
      Top             =   825
      Width           =   345
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Target Approach"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2415
      TabIndex        =   26
      Top             =   555
      Width           =   1305
   End
   Begin VB.Label Target_Approach_dis 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3705
      TabIndex        =   25
      Top             =   555
      Width           =   540
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   870
      Top             =   825
      _version        =   65536
      _extentx        =   847
      _extenty        =   847
      _stockprops     =   0
      cdtimeout       =   0
      commport        =   1
      ctstimeout      =   0
      dsrtimeout      =   0
      dtrenable       =   -1  'True
      handshaking     =   0
      inbuffersize    =   1024
      inputlen        =   0
      interval        =   1000
      nulldiscard     =   0   'False
      outbuffersize   =   512
      parityreplace   =   "?"
      rthreshold      =   0
      rtsenable       =   0   'False
      settings        =   "9600,n,8,1"
      sthreshold      =   0
   End
   Begin VB.Label Last 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2040
      TabIndex        =   21
      Top             =   255
      Width           =   600
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Target"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   30
      TabIndex        =   20
      Top             =   30
      Width           =   690
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Actual"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   735
      TabIndex        =   19
      Top             =   30
      Width           =   690
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Differ."
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   1440
      TabIndex        =   18
      Top             =   30
      Width           =   615
   End
   Begin VB.Label Label43 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lasted"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   2040
      TabIndex        =   17
      Top             =   30
      Width           =   600
   End
   Begin VB.Label cur_timer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2400
      TabIndex        =   16
      Top             =   825
      Width           =   465
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tare"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   2625
      TabIndex        =   15
      Top             =   30
      Width           =   690
   End
   Begin VB.Label Tara_dis 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2625
      TabIndex        =   14
      Top             =   255
      Width           =   690
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Target Value"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   3300
      TabIndex        =   13
      Top             =   30
      Width           =   945
   End
   Begin VB.Label rix 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1935
      TabIndex        =   12
      Top             =   555
      Width           =   465
   End
   Begin VB.Label Zyg_Show 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C0C0FF&
      Height          =   250
      Left            =   825
      TabIndex        =   11
      Top             =   1140
      Width           =   30
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Action "
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   250
      Left            =   30
      TabIndex        =   10
      Top             =   1140
      Width           =   780
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Time Target "
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Top             =   825
      Width           =   990
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bottle Rate "
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   30
      TabIndex        =   8
      Top             =   555
      Width           =   900
   End
   Begin VB.Label Rate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   915
      TabIndex        =   7
      Top             =   555
      Width           =   510
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tryals "
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   555
      Width           =   510
   End
End
Attribute VB_Name = "Syn_Dos"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim cancel_pressed As Boolean
Dim SafetyFactor(10)
Dim Rate_Increasing As Double, Rate_Increasing_Point As Double
Dim Low_Rate_Safety As Double, Global_Buttom_Step As Double
Dim A_Rate As Double, B_Rate As Double, Target_Safety_Break_Point As Integer
Dim Online_Shoot As Integer
Sub FindSafetyFactors()
Dim Db As Database, r As Recordset
Set Db = OpenDatabase("C:\Talos\skon_tb.mdb")
Set r = Db.OpenRecordset("Safety_Factors")
Do Until r.EOF
I = I + 1
If I < 7 Then
        SafetyFactor(I) = r("Safety_Factor")
Else
        If I = 7 Then Rate_Increasing = r("Safety_Factor")
        If I = 8 Then Rate_Increasing_Point = r("Safety_Factor")
        If I = 9 Then Low_Rate_Safety = r("Safety_Factor")
        If I = 10 Then Global_Buttom_Step = r("Safety_Factor")
        If I = 11 Then A_Rate = r("Safety_Factor")
        If I = 12 Then B_Rate = r("Safety_Factor")
        If I = 13 Then Target_Safety_Break_Point = r("Safety_Factor")
        If I = 14 Then Online_Shoot = r("Safety_Factor")
End If
r.MoveNext
Loop
r.Close
End Sub

Public Sub FindValves()
Dim MyDB As Database, r As Recordset
On Error GoTo er_c:
Set MyDB = OpenDatabase("c:\talos\coord_tb.mdb")
Set r = MyDB.OpenRecordset("Βαλβιδες")
Beaker_Bottle = ""
Do Until r.EOF
  If r("valve_kod") = 1 Then
    Gripper_On = r("valve_on")
     Gripper_Off = r("valve_off")
   End If
   If r("valve_kod") = 2 Then
      Ready_Bit = Val(Mid(r("valve_on"), 5, 2))
   End If
    If r("valve_kod") = 3 Then
   Braxionas = r("valve_on")
   End If
  If r("valve_kod") = 4 Then
    skon_motor_on = r("valve_on")
     skon_motor_off = r("valve_off")
   End If
   If r("valve_kod") = 5 Then
     Cold_Water_On = r("valve_on")
     Cold_Water_Off = r("valve_off")
   End If
   If r("valve_kod") = 6 Then
     Hot_Water_On = r("valve_on")
     Hot_Water_Off = r("valve_off")
   End If
   If r("valve_kod") = 7 Then
     mixer_on = r("valve_on")
     mixer_off = r("valve_off")
   End If
   If r("valve_kod") = 8 Then
     Balance_Port = r("valve_on")
  End If
  If r("valve_kod") = 9 Then
     valve_1_on = r("valve_on")
     valve_1_off = r("valve_off")
   End If
      If r("valve_kod") = 10 Then
     valve_2_on = r("valve_on")
     valve_2_off = r("valve_off")
   End If
   If r("valve_kod") = 11 Then
     Alarm1_on = r("valve_on")
     Alarm1_off = r("valve_off")
  End If
   If r("valve_kod") = 12 Then
     Pause_input = r("valve_on")
     Air_Input = r("valve_off")
  End If
    If r("valve_kod") = 13 Then
     Door_input = r("valve_on")
  End If
  If r("valve_kod") = 14 Then
    Gripper_input = r("valve_on")
  End If
   If r("valve_kod") = 15 Then
    Talos_Fonts = r("valve_on")
    Talos_Language = r("valve_off")
  End If
  If r("valve_kod") = 16 Then
    Motor_Input = r("valve_on")
    Motor_Preheating = r("valve_off")
  End If
   If r("valve_kod") = 17 Then
    Motor_output_on = r("valve_on")
    Motor_output_off = r("valve_off")
  End If
    If r("valve_kod") = 18 Then
    Alarm2_on = r("valve_on")
    Alarm2_off = r("valve_off")
  End If
    If r("valve_kod") = 19 Then
    Alarm3_On = r("valve_on")
    Alarm3_Off = r("valve_off")
  End If
   If r("valve_kod") = 20 Then
    Central_Water_On = r("valve_on")
    Central_Water_Off = r("valve_off")
  End If
  If r("valve_kod") = 21 Then
    Balance_Settings = r("valve_on")
    Balance_Asking = r("valve_off")
 End If
  If r("valve_kod") = 22 Then
    Balance_2_Digits = r("valve_on")
    Balance_3_Digits = r("valve_off")
 End If
    If r("valve_kod") = 23 Then
    Balance_Zero = r("valve_on")
    Balance_Clear = r("valve_off")
 End If
    If r("valve_kod") = 24 Then
    Water_Bypass = r("valve_on")
  End If
  If r("valve_kod") = 25 Then
    Balance_Type = r("valve_on")
    Balance_Mode = r("valve_off")
 End If
  If r("valve_kod") = 26 Then
    Program1 = r("valve_on")
    Program2 = r("valve_off")
 End If
   If r("valve_kod") = 27 Then
    Program3 = r("valve_on")
  If IsNull(r("valve_off")) Then Program4 = "" Else Program4 = r("valve_off")
 End If
If r("valve_kod") = 28 Then
    Beaker_Bottle = r("valve_on")
End If
r.MoveNext
Loop
r.Close

Exit Sub
er_c:
If Err = 13 Then Resume Next
If Err = 94 Then Resume Next
'Err_Ck (Err)
Resume Next
End Sub

Private Sub Command1_Click()
 Command1.Enabled = False
cancel_pressed = True
End Sub

Private Sub Form_Activate()
DoEvents
cancel_pressed = False
If Me.Caption = "Standard Time" Then
Me.real_q_dis = StandardTime(Asked_Q, Level)
Else
    Me.real_q_dis = Servirisma(Asked_Q, Level)
End If
Form2.real_q_dis = Me.real_q_dis
asw = Val(Form2.Level)
Form2.Level = asw - Val(Me.real_q_dis) / 1000
Form2.Target_Safety_Dis = Trgt_Safety
Unload Me
End Sub

Private Sub anix_zygaria()
On Error GoTo err_fnd
  MSComm1.CommPort = Balance_Port
    MSComm1.Settings = Balance_Settings
    If MSComm1.PortOpen = False Then MSComm1.PortOpen = True
    MSComm1.Settings = Balance_Settings
    MSComm1.InputLen = 0
Exit Sub
err_fnd:
If Err = 8013 Then
    zyg_err = " "
Else
MsgBox Err.Description, , "Talos "
End If
Resume Next
End Sub
Private Sub Form_Load()
FindValves
FindSafetyFactors
system_ready = IsOSLoaded(768)
anix_zygaria

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
  
   Form2.rix.Caption = rix
    Form2.last = Me.last
   MSComm1.PortOpen = False
End Sub

Sub Valve_Off_Click()
cmd$ = "!" + Valve_off + ":"
robsend (cmd$)

End Sub

Private Sub Valve_on_Click()
cmd$ = "!" + Valve_on + ":"
robsend (cmd$)

End Sub

Function check_zyg(arg) As String
check_zyg = arg
If InStr(arg, "OL") > 0 Then
    check_zyg = "Error - 3"
    Exit Function
End If
If arg = "System not Working" Then
    check_zyg = "Error - 4"
    Exit Function
End If
If arg = " " Then
    check_zyg = "Error - 5"
    Exit Function
End If
If InStr(arg, "UL") > 0 Then
    check_zyg = "Error - 6"
    Exit Function
End If
End Function
Private Function StandardTime(Asked_Q, Level) As Single
  start_dis = GetCurrentTime()
If Level = "" Then
   Screen.MousePointer = 1
    StandardTime = 0
    Exit Function
End If
If Level < 50 Then
   Screen.MousePointer = 1
    StandardTime = -10
    Exit Function
End If
 tot_cycles = 0: tryal = 0
Refresh

   Screen.MousePointer = 11
   MSComm1.Output = Balance_3_Digits + Chr$(13)
8
If Val(Asked_Q) < 100 Then
       arw = check_zyg(zygisi4("OK")): If InStr(arw, "Error") > 0 Then GoTo stand_error_exit
      tara = Val(arw) * 1000
Else
       arw = check_zyg(zygisi4(0)): If InStr(arw, "Error") > 0 Then GoTo stand_error_exit
      tara = Val(arw) * 1000
End If

Me.tara_dis = tara: Target_Value = tara + Val(Asked_Q): count_val = 0: Buttom_Step_Reach = False
rix = 1: Me.dif = 0: Me.last = 0: Me.cur_timer = "": Me.Rate = ""
Real_Q = 0: real_q_dis = 0: Safety_Factor = 1: Final_target = 2: Approach_Target = False
 Me.Zyg_Show.Width = 0
first_Time_Wait = 2000
Max_Time_Wait = Val(Asked_Q) * 10

 Valve_on_Click
  start = Val(GetCurrentTime())
  MilSec Val(Asked_Q)
 Valve_Off_Click
 last_time = Val(GetCurrentTime())
Dosing_Time = last_time - start
If cancel_pressed = True Then arw = "Error - 9": GoTo stand_error_exit
     Label1 = "Waiting ...": Me.Label1.Refresh
     Me.Zyg_Show = first_Time_Wait / 1000 & " Sec"
      Me.Zyg_Show.Width = 450
       MilSec Val(first_Time_Wait)
        arw = check_zyg(zygisi4("OK")): If InStr(arw, "Error") > 0 Then GoTo stand_error_exit
        asw = Val(arw) * 1000 - Val(tara)
        real_q_dis = Int(asw + 0.5)
       If real_q_dis < 0 Then real_q_dis = 0
    StandardTime = Val(real_q_dis)
If cancel_pressed = True Then arw = "Error - 9": GoTo stand_error_exit
Exit Function
stand_error_exit:
  StandardTime = -7
     Screen.MousePointer = 1
End Function
Private Function Servirisma(Asked_Q As String, Level As String) As Single
Dim Db As Database, r As Recordset
Dim asw As String, Real_Q, tara As String, Asked_Stage As Integer, Current_Dosing As String
Dim Buttom_Step_Reach As Boolean, Dosing_String As String, Real_Dosing_Time As String
Dim Bottle_Number As Integer, Bottle_Found As Boolean, Buttom_repare As Integer
Dim gram As String, Tot_Gram As String, ader As Integer
Dim Target_Safety As Integer
On Error GoTo serv_error_exit

'=====  Σφάλματα που επιστρέφονται

'Servirisma = 0    ==> Ελλειπή Arguments
'Servirisma = -1   ==> Μηδέν ή αρνητικό Asked_Q
'Servirisma = -2   ==> η βαλβίδα δεν τρέχει
'Servirisma = -3   ==> η ζυγαρια υπερφορτώθηκε
'Servirisma = -4   ==> δεν έγινε διαδικασία Εναρξης
'Servirisma = -5   ==> Η ζυγαριά δεν ανταποκρίνεται
'Servirisma = -6   ==> Η ζυγαριά εκτός περιοχής
'Servirisma = -7   ==> Η ζύγιση έχει αποτύχει
'Servirisma = -8   ==> Παρουσιάσθηκε απροσδόκητο λάθος στον αλγόρυθμο ζύγισης
'Servirisma = -9   ==> Η ζύγιση ακυρώθηκε
'Servirisma = -10   ==> Η στάθμη του μπουκαλιού είναι πολύ χαμηλή <50 gr
'Servirisma = -11   ==> Υπερβολική  ζητούμενη ποσότητα  (max. 100 gr)
  start_dis = GetCurrentTime()
If Level = "" Then
   Screen.MousePointer = 1
    Servirisma = 0
    Exit Function
End If
If Level < 50 Then
   Screen.MousePointer = 1
    Servirisma = -10
    Exit Function
End If
 tot_cycles = 0: tryal = 0
Refresh

   Screen.MousePointer = 11
   MSComm1.Output = Balance_3_Digits + Chr$(13)
8
If Val(Asked_Q) < 100 Then
       arw = check_zyg(zygisi4("OK")): If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
      tara = Val(arw) * 1000
Else
       arw = check_zyg(zygisi4(0)): If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
      tara = Val(arw) * 1000
End If

'A_Rate = 0.000567: B_Rate = 0.26

Me.tara_dis = tara: Target_Value = tara + Val(Asked_Q): count_val = 0: Buttom_Step_Reach = False
rix = 1: Me.dif = 0: Me.last = 0: Me.cur_timer = "": Me.Rate = ""
Real_Q = 0: real_q_dis = 0: Safety_Factor = 1: Final_target = 2: Approach_Target = False
 Me.Zyg_Show.Width = 0: Ocasion = 0
first_Time_Wait = 2000: Dosing_String = " ": Real_Dosing_Time = " "
Max_Time_Wait = Val(Asked_Q) * 20

GoSub Find_Bottle_Number
Dosing_Rate = A_Rate * Val(Level) + B_Rate
GoSub Find_Bottle_Buttom_Step
 Rate = Int(Dosing_Rate * 1000 + 0.5) / 1000
 GoSub Find_Final_target
  
  Target_Approach_dis = Final_target
  If Final_target = 0 Then
    Screen.MousePointer = 1
    Servirisma = -11
    Exit Function
End If

'============================   Αρχή Συνάρτησης   ======================
     fact = Val(Me.Level) + Val(Target_Safety)

'==================                Μεγάλες Ποσότητες (πάνω από 400 mgr)
If Val(Asked_Q) > fact Then
Ocasion = 1
 Label1 = "Dosing ...": Label1.Refresh
Valve_on_Click
  start = Val(GetCurrentTime())
Do
       arw = check_zyg(zygisi4(0)): If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
      asw = Val(arw) * 1000 - Val(tara)
     Me.real_q_dis = Int(asw + 0.5)
     Me.dif = Val(Asked_Q) - Val(real_q_dis)
     If cancel_pressed = True Then
              Servirisma = -9
                Valve_Off_Click
                 Screen.MousePointer = 1
                 Exit Function
      End If
     If GetCurrentTime() > start + Max_Time_Wait Then
                  Servirisma = -2
                 Valve_Off_Click
                 Screen.MousePointer = 1
                 Exit Function
      End If
     If (Val(asw) + fact) / Val(Asked_Q) * 2500 > 0 Then Zyg_Show.Width = Int((Val(asw) + fact) / Val(Asked_Q) * 2500)
      Zyg_Show = Int((Val(asw)) / Val(Asked_Q) * 100) & " %"
Loop Until Val(asw) + fact >= Val(Asked_Q) + Final_target
asw1 = asw
1
 Valve_Off_Click
 last_time = Val(GetCurrentTime())
Dosing_Time = last_time - start: Real_Dosing_Time = Real_Dosing_Time + Str(Dosing_Time) + " "
2
      If cancel_pressed = True Then arw = "Error - 9": GoTo serv_error_exit
      Label1 = "Waiting ...": Me.Label1.Refresh
      Me.Zyg_Show = first_Time_Wait / 1000 & " Sec"
      Me.Zyg_Show.Width = 450
      MilSec Val(first_Time_Wait)
      arw = check_zyg(zygisi4("OK")): If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
      count_val = count_val + 1
        If count_val > 14 Then
            Servirisma = -7
                Valve_Off_Click
                 Screen.MousePointer = 1
                 Exit Function
        End If
        asw = Val(arw) * 1000 - Val(tara)
        real_q_dis = Int(asw + 0.5)
         Me.dif = Val(Asked_Q) - Val(real_q_dis)
         'Αλλαγή του Trgt_Safety
             If Val(rix) = 1 Then
                If Me.dif - Final_target < 0 Then
                      Target_Safety = Target_Safety + 25
                      Trgt_Safety = Target_Safety
                 ElseIf Me.dif > Target_Safety_Break_Point Then
                     Target_Safety = Target_Safety - 25
                     Trgt_Safety = Target_Safety
                End If
            End If
If Val(asw) < Val(asw1) Then
      prosp = prosp + 1
      asw1 = asw
       If prosp > 7 Then
      Servirisma = -7
      Exit Function
End If
      GoTo 2
End If
      If rix = 1 Then Real_Q = 0
          last_q = asw - Val(Real_Q)
         If last_q < 0 Then
             arw = check_zyg(zygisi4("OK")): If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
              asw = Val(arw) * 1000 - Val(tara)
               Real_Q = Int(asw + 0.5)
               Me.real_q_dis = Real_Q
               Me.dif = Val(Asked_Q) - Val(Real_Q)
               Dosing_Rate = 0.55
              GoTo 3
        End If
         Real_Q = Int(asw + 0.5)
         Me.real_q_dis = Real_Q

Me.dif = Val(Asked_Q) - Val(Real_Q)
If Val(Me.dif) > Online_Shoot Then Me.dif = Online_Shoot
Me.last = Val(GetCurrentTime()) - start
If Val(Me.dif) > 50 And last_q / Dosing_Time > Dosing_Rate * 0.9 Then Dosing_Rate = last_q / Dosing_Time ' Mgr/Msec
Rate = Int(Dosing_Rate * 1000 + 0.5) / 1000
If Val(Me.dif) < 50 And Val(Me.dif) > Final_target Then
     Approach_Target = True
End If
If Dosing_Rate <= 0 Then
      sec_tim = 20: GoTo 2
Else
      sec_tim = (Me.dif / Dosing_Rate) * Safety_Factor
End If
If Approach_Target = True Then
   If Val(Me.dif) > Final_target Then
     If Buttom_Step_Reach = True And last_q <= 1 Then Buttom_step = Buttom_step + 2
    End If
End If
  If Val(Me.dif) > Final_target Then GoSub Find_Safety_Factor
    If sec_tim > 10000 Then GoTo 2
    If sec_tim <= 30 And Val(rix) > 1 Then sec_tim = Buttom_step + Val(Me.dif) / 2 - 1: Buttom_Step_Reach = True: Me.cur_timer = Int(sec_tim + 0.5)
Me.cur_timer = Int(sec_tim + 0.5)
3
 If cancel_pressed = True Then arw = "Error - 7":  GoTo serv_error_exit
If Val(Me.dif) > Final_target Then
    If sec_tim <= 30 Then first_Time_Wait = 2500
    If sec_tim >= Max_Time_Wait Then GoTo 2
    If sec_tim = Buttom_step Then
        sec_tim = Buttom_step + Val(Me.dif) / 2 - 1
        Me.cur_timer = Int(sec_tim + 0.5)
    End If
    count_val = 0
     Dosing_String = Dosing_String + Str(Real_Q) + " "
     Label1 = "Dosing ...": Label1.Refresh
    Valve_on_Click
    start = Val(GetCurrentTime())
    MilSec Val(Int(sec_tim + 0.5))
    rix = Val(rix) + 1
    If rix > 14 Then arw = "Error - 7":   GoTo serv_error_exit
    GoTo 1
End If

'============      Ποσότητες από 1 εως 400  mgr   ==============

ElseIf Val(Asked_Q) > 1 Then
Ocasion = 2
        first_Time_Wait = 1000
        If Val(Asked_Q) < 20 Then first_Time_Wait = 3000
        sec_tim = Val(Asked_Q) / Dosing_Rate * Low_Rate_Safety
       If sec_tim < 30 Then GoSub Find_Safety_Factor
        If sec_tim < Buttom_step Then sec_tim = Buttom_step
        cur_timer = Int(sec_tim + 0.5)
        Label1 = "Dosing ..."
         Valve_on_Click
         start = Val(GetCurrentTime())
        MilSec Val(sec_tim)
4
       If cancel_pressed = True Then arw = "Error - 9":  GoTo serv_error_exit
        last_time = Val(GetCurrentTime())
        Valve_Off_Click
        Dosing_Time = last_time - start:  Real_Dosing_Time = Real_Dosing_Time + Str(Dosing_Time) + " "
        Label1 = "Waiting ...": Me.Zyg_Show = first_Time_Wait / 1000 & " Sec"
         Me.Zyg_Show.Width = 450
         MilSec Val(first_Time_Wait)
5
       If cancel_pressed = True Then arw = "Error - 9":  GoTo serv_error_exit
        Label1.Visible = True
        arw = check_zyg(zygisi4("OK")): If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
        asw = Val(arw) * 1000 - Val(tara)
        last_q = asw - Val(Real_Q)
      If last_q <= 0 Then
            tryal = tryal + 1
            If tryal > 10 Then arw = "Error - 2":  GoTo serv_error_exit
       Else
            tryal = 0
      End If
        Real_Q = Int(asw + 0.5)
         Me.real_q_dis = Real_Q
        Me.dif = Val(Asked_Q) - Val(Real_Q)
        If Val(Me.dif) > Online_Shoot Then Me.dif = Online_Shoot
         If Val(Me.dif) < 50 Then Bottle_Found = True
        Me.last = Val(GetCurrentTime()) - start
       If last_q <= 0 Then
            tara = Val(tara_dis) + last_q: tara_dis = tara
            Target_Value = tara + Val(Asked_Q)
            Buttom_step = Buttom_step + 2
            GoTo 5
    End If
      
        If Val(Me.dif) > Final_target Then GoSub Find_Safety_Factor
        ' If Val(Me.dif) > Val(Asked_Q) / 2 And Val(Asked_Q) > 100 Then sec_tim = sec_tim / 2
       Me.cur_timer = Int(sec_tim + 0.5)
    If Val(Me.dif) > Final_target Then
           If sec_tim > Max_Time_Wait Then
              tryal = tryal + 1
                If tryal > 10 Then
                        arw = "Error - 2":   GoTo serv_error_exit
                 Else
                      GoTo 5
                End If
            End If
            Dosing_String = Dosing_String + Str(Real_Q) + " "
            Me.Label1 = "Dosing ...": Me.Zyg_Show.Width = 0
            Valve_on_Click
           start = Val(GetCurrentTime())
            MilSec Val(sec_tim)
            rix = Val(rix) + 1
            If rix > 15 Then arw = "Error - 7":   GoTo serv_error_exit
            GoTo 4
    End If

ElseIf Val(Asked_Q) <= 0 Then
    Servirisma = -1
    Screen.MousePointer = 1
     Exit Function
End If

'============      Εξοδος Συνάρτησης ==============================================

Servirisma = Val(Real_Q)
last = GetCurrentTime() - start_dis
 Dosing_String = Dosing_String + Str(Real_Q) + " "
9
'On Error Resume Next
Set Db = OpenDatabase("C:\Talos\skon_tb.mdb")
Set r = Db.OpenRecordset("Syntages")
r.AddNew
If Me.Caption <> "" Then r("Description") = Left(Me.Caption, 50) Else r("Description") = "Botlle Unknown"
r("Level") = Int(Val(Me.Level))
r("Tryals") = Val(rix)
r("Asked_Q") = Val(Asked_Q)
If Abs(Val(Me.real_q_dis) - Val(Asked_Q)) <= Final_target Then target_fin = Val(Asked_Q)
If Val(Me.real_q_dis) < Val(Asked_Q) - Final_target Then target_fin = Val(Me.real_q_dis) + Final_target
If Val(Me.real_q_dis) > Val(Asked_Q) + Final_target Then target_fin = Val(Me.real_q_dis) - Final_target
r("Actual_q") = target_fin
r("Lasted") = Int(Val((Me.last)) / 1000 + 0.5)
r("Level_Stage") = Int((Val(Level) + 50) / 100)
r("Asked_Stage") = Asked_Stage
r("Dosing_String") = Dosing_String + Real_Dosing_Time
r.Update
r.Close

   '======================= Ενημέρωση Αρχείου Buttom Step   =======================
   If Bottle_Number > 0 Then
  Set r = Db.OpenRecordset("Bottle_Buttom_Steps", dbOpenDynaset)
   If Val(Asked_Q) > 400 Then ader = 3 Else ader = 6
  If Bottle_Found = True Then
    cret = "[Bottle_Number]=" & Bottle_Number
    r.FindFirst cret
    r.Edit
        If Buttom_Step_Reach = True Then
            dior = (Val(Asked_Q) - target_fin) / 2
            ade = (Buttom_step + Buttom_repare * 2) / 3
            asd = Int(ade + Val(rix) - ader + dior)
             If asd > 40 Then asd = 40
             If asd < 8 Then asd = 8
        End If
       r("Buttom_Step") = asd
       r("Target_Safety") = Target_Safety
       Select Case Val(Level)
             Case Is > 450: r("Rate_500") = Dosing_Rate + A_Rate * (500 - Val(Level))
             Case Is > 350: r("Rate_400") = Dosing_Rate + A_Rate * (400 - Val(Level))
             Case Is > 250: r("Rate_300") = Dosing_Rate + A_Rate * (300 - Val(Level))
             Case Is > 150: r("Rate_200") = Dosing_Rate + A_Rate * (200 - Val(Level))
             Case Else: r("Rate_100") = Dosing_Rate + A_Rate * (100 - Val(Level))
       End Select

    r.Update
   Else
    r.AddNew
         r("Target_Safety") = Target_Safety
        If Buttom_Step_Reach = True Then
              dior = (Val(Asked_Q) - target_fin) / 2
              ade = (Buttom_step + Buttom_repare * 2) / 3
              asd = Int(ade + Val(rix) - ader + dior)
              If asd > 40 Then asd = 40
              If asd < 5 Then asd = 5
              r("Buttom_Step") = asd
        End If
        r("Bottle_Number") = Bottle_Number
   r.Update
   End If
   r.Close
 End If
 Screen.MousePointer = 1
Exit Function

 '===================  Τέλος Συνάρτησης  ============================
serv_error_exit:
If Err > 0 Then
      Resume Next
      arw = "- 8"
End If
 Valve_Off_Click
If InStr(arw, "- 2") Then
    Servirisma = -2
    Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 3") > 0 Then
    Servirisma = -3
    Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 8") > 0 Then
    Servirisma = -8
    Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 6") > 0 Then
    Servirisma = -6
    Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 7") > 0 Then
    Servirisma = -7
    Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 4") Then
    Servirisma = -4
    Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 5") Then
    Servirisma = -5
     Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 9") Then
    Servirisma = -9
     Screen.MousePointer = 1
    Exit Function
End If
GoTo 9


Find_Safety_Factor:
If last_q > 12 Then
cur_Dosing_Rate = Dosing_Rate
Select Case sec_tim
Case Is > 700
first_Time_Wait = 1000
   Safety_Factor = SafetyFactor(6)
Case Is > 300
first_Time_Wait = 1500
Safety_Factor = SafetyFactor(5)
Case Is > 100
first_Time_Wait = 2000
Safety_Factor = SafetyFactor(4)
Case Is > 50
first_Time_Wait = 2000
Safety_Factor = SafetyFactor(3)
Case Is > 25
first_Time_Wait = 2500
Safety_Factor = SafetyFactor(2)
Case Else
first_Time_Wait = 3000
Safety_Factor = SafetyFactor(1)
End Select
Dosing_Rate = (last_q / Dosing_Time + cur_Dosing_Rate * 9) / 10
If Dosing_Rate <= 0 Then Dosing_Rate = 0.5
End If
 Rate = Int(Dosing_Rate * 1000 + 0.5) / 1000
sec_tim = (Me.dif / Dosing_Rate) * Safety_Factor
If sec_tim <= Buttom_step Then sec_tim = Buttom_step: Buttom_Step_Reach = True
If sec_tim <= Buttom_step Then sec_tim = Buttom_step: Buttom_Step_Reach = True
If sec_tim < 40 Then
        New_Rate = Dosing_Rate * (1 + Rate_Increasing * (Rate_Increasing_Point - sec_tim) / 30)
        sec_tim = Me.dif / New_Rate
    End If
If sec_tim <= Buttom_step Then sec_tim = Buttom_step
Me.cur_timer = Int(sec_tim + 0.5)
Return


Find_Final_target:
Select Case Val(Asked_Q)
Case Is < 50
Final_target = 2
Asked_Stage = 1
Case Is < 200
Final_target = 2
Asked_Stage = 2
Case Is < 1000
Final_target = 3
Asked_Stage = 3
Case Is < 2000
Final_target = 3
Asked_Stage = 4
Case Is < 3000
Final_target = 15
Asked_Stage = 5
Case Is < 4000
Final_target = 20
Asked_Stage = 6
Case Is < 10000
Final_target = Val(Asked_Q) * 0.005
Asked_Stage = 7
Case Is < 25000
Final_target = Val(Asked_Q) * 0.005
Asked_Stage = 8
Case Is < 100001
Final_target = Val(Asked_Q) * 0.005
Asked_Stage = 9
Case Else
Final_target = 0
Asked_Stage = 0
End Select
Return

Find_Bottle_Buttom_Step:
Set Db = OpenDatabase("C:\Talos\skon_tb.mdb")
Set r = Db.OpenRecordset("Bottle_Buttom_Steps", dbOpenDynaset)
    cret = "[Bottle_Number]=" & Bottle_Number
    r.FindFirst cret
    If r.NoMatch Then
        Buttom_step = Global_Buttom_Step: Buttom_repare = Global_Buttom_Step
        Bottle_Found = False
        Target_Safety = 225
        Trgt_Safety = Target_Safety
    Else
         Bottle_Found = True
         Buttom_step = r("Buttom_Step"): Buttom_repare = Buttom_step
         If Not IsNull(r("Target_safety")) Then Target_Safety = r("Target_safety") Else Target_Safety = 225
        Trgt_Safety = Target_Safety
        Select Case Level
             Case Is > 450
             If Not IsNull(r("Rate_500")) Then Dosing_Rate = r("Rate_500") - A_Rate * (500 - Val(Level))
             Case Is > 350
             If Not IsNull(r("Rate_400")) Then Dosing_Rate = r("Rate_400") - A_Rate * (400 - Val(Level))
             Case Is > 250
             If Not IsNull(r("Rate_300")) Then Dosing_Rate = r("Rate_300") - A_Rate * (300 - Val(Level))
             Case Is > 150
             If Not IsNull(r("Rate_200")) Then Dosing_Rate = r("Rate_200") - A_Rate * (200 - Val(Level))
             Case Else
             If Not IsNull(r("Rate_100")) Then Dosing_Rate = r("Rate_100") - A_Rate * (100 - Val(Level))
       End Select
    End If
 r.Close

Return

Find_Bottle_Number:
gram = "": Tot_Gram = ""
If Me.Caption = "" Then Bottle_Number = 0: Return
For I = 1 To Len(Me.Caption)
    gram = Mid(Me.Caption, I, 1)
    If gram = ":" Then Exit For
    If IsNumeric(gram) Then Tot_Gram = Tot_Gram + gram
Next I
Bottle_Number = Val(Tot_Gram)
Return

End Function
Function zygisi4(Zygis_Kind As Variant)
On Error GoTo er_det
   Dim counter
 
   counter = 0: tot_counter = 0: Label10 = 0
 If system_ready = 0 Then zygisi4 = "System not Working": Exit Function
  start = GetCurrentTime()
   If Zygis_Kind = "OK" Then Label1 = "Scaling ..."
    Do
    If GetCurrentTime() - start > 10001 Then zygisi4 = " ": Exit Function
     MSComm1.InBufferCount = 0
      FromModem$ = ""
   MSComm1.Output = Balance_Asking + Chr$(13)
       MilSec (40)
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
            If GetCurrentTime() - start > 5000 And Zygis_Kind = "OK" Then Zygis_Kind = 0: tot_counter = 0
         ' asq = FromModem$
          ' asw = MSComm1.Input
          If Zygis_Kind = "OK" Then
           If Mid$(FromModem$, 5, 1) = "." Or Mid$(FromModem$, 6, 1) = "." Or Mid$(FromModem$, 7, 1) = "." Then
                If InStr(FromModem$, "OK") > 0 Then
                    Exit Do
                End If
         End If
        Else
          If InStr(FromModem$, Chr$(13)) Then
                If Mid$(FromModem$, 5, 1) = "." Or Mid$(FromModem$, 6, 1) = "." Or Mid$(FromModem$, 7, 1) = "." Then
                       Exit Do
                End If
          End If
       End If
       counter = counter + 1
       If counter >= 15 Then tot_counter = tot_counter + 1
            Label10 = counter
          If tot_counter >= 20 And Zygis_Kind <> "OK" Then
          counter = 0: tot_counter = tot_counter + 1
               zygisi4 = " "
               Exit Function
           End If
          If counter >= 50 Then  ' 500
             counter = 0: tot_counter = tot_counter + 1
               MSComm1.InBufferCount = 0
              MSComm1.Output = Balance_Asking + Chr$(13)
              End If
     End If
     If Zygis_Kind = "OK" Then
             Me.Zyg_Show.Width = (GetCurrentTime() - start) / 1000 * 650
              Me.Zyg_Show.Caption = Int(((GetCurrentTime() - start) / 1000) + 0.5) & " Sec"
             Me.Zyg_Show.Refresh
      End If
    Loop
         zygisi4 = Left$(FromModem$, 9)
er_ex:
Exit Function
er_det:
 zygisi4 = " "
Resume er_ex
End Function

Sub robsend(cmd$)
   temp% = SendAT6400Block(768, cmd$, 0)
End Sub

Sub MilSec(wait As Long)
     start = Val(GetCurrentTime())
     Do
        c_tim = Val(GetCurrentTime())
        DoEvents
     Loop Until c_tim >= start + wait
End Sub


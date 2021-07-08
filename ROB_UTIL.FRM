VERSION 4.00
Begin VB.Form Form2 
   BackColor       =   &H0080FFFF&
   Caption         =   "Talos : Dosometric Utility"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   1365
   ClientWidth     =   9450
   DrawMode        =   7  'Invert
   ForeColor       =   &H000000FF&
   Height          =   7005
   Icon            =   "ROB_UTIL.frx":0000
   Left            =   0
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   9450
   Top             =   1020
   Width           =   9570
   Begin VB.TextBox Target_Safety_Dis 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
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
      Height          =   300
      Left            =   1905
      TabIndex        =   19
      Top             =   225
      Width           =   1005
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "Standart Time     "
      ForeColor       =   &H00FFFF80&
      Height          =   240
      Left            =   7920
      TabIndex        =   18
      Top             =   15
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   285
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   240
      Width           =   1485
   End
   Begin VB.TextBox Bottle_Description 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   6405
      TabIndex        =   16
      Top             =   240
      Width           =   1500
   End
   Begin VB.TextBox Level 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   4320
      TabIndex        =   7
      Text            =   "350"
      Top             =   240
      Width           =   990
   End
   Begin VB.TextBox real_q_dis 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   2940
      TabIndex        =   6
      Top             =   240
      Width           =   690
   End
   Begin VB.TextBox asked_q 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   3645
      TabIndex        =   5
      Text            =   "600"
      Top             =   240
      Width           =   690
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete Old Values"
      Height          =   540
      Left            =   930
      TabIndex        =   4
      Top             =   0
      Width           =   960
   End
   Begin VB.TextBox valve_1_off 
      Height          =   285
      Left            =   450
      TabIndex        =   3
      Top             =   -120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox valve_1_on 
      Height          =   285
      Left            =   225
      TabIndex        =   2
      Top             =   -90
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.ListBox List1 
      Columns         =   6
      Height          =   6105
      Left            =   0
      TabIndex        =   1
      Top             =   555
      Width           =   9555
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Start Dosing"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   885
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Target Safety"
      BeginProperty Font 
         name            =   "MS Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   1905
      TabIndex        =   20
      Top             =   15
      Width           =   1005
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bottle Description"
      BeginProperty Font 
         name            =   "MS Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   6405
      TabIndex        =   15
      Top             =   15
      Width           =   1500
   End
   Begin VB.Label rix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         name            =   "MS Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   5940
      TabIndex        =   14
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tryals"
      BeginProperty Font 
         name            =   "MS Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   5940
      TabIndex        =   13
      Top             =   15
      Width           =   495
   End
   Begin VB.Label Label43 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lasted"
      BeginProperty Font 
         name            =   "MS Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   5325
      TabIndex        =   12
      Top             =   15
      Width           =   600
   End
   Begin VB.Label Last 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         name            =   "MS Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   5325
      TabIndex        =   11
      Top             =   240
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bottle Content"
      BeginProperty Font 
         name            =   "MS Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   4320
      TabIndex        =   10
      Top             =   15
      Width           =   990
   End
   Begin VB.Label Label41 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Actual"
      BeginProperty Font 
         name            =   "MS Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   2940
      TabIndex        =   9
      Top             =   15
      Width           =   690
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Target"
      BeginProperty Font 
         name            =   "MS Serif"
         charset         =   1
         weight          =   400
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   3645
      TabIndex        =   8
      Top             =   15
      Width           =   690
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   8790
      Top             =   660
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
End
Attribute VB_Name = "Form2"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Dim system_ready As Integer
Dim cmd As String
Function zyg_decimals()
 Dim counter
    
 If system_ready = 0 Then zyg_decimals = 0: Exit Function
    ' Flush the input buffer.
       MSComm1.Output = "SEND" + Chr$(13)
       MSComm1.InBufferCount = 0
    
    counter = 0: tot_counter = 0
    ' Wait for "OK" to come back from the modem.
    Do
       dummy = DoEvents()
       
       ' If there is data in the buffer, then read it.
       If MSComm1.InBufferCount Then
          FromModem$ = FromModem$ + MSComm1.Input
          ' Check for "OK".
          If InStr(FromModem$, Chr$(13)) Then
                If Mid$(FromModem$, 5, 1) = "." Or Mid$(FromModem$, 6, 1) = "." Or Mid$(FromModem$, 7, 1) = "." Then
                    Exit Do
                End If
          End If
       Else
            counter = counter + 1
            If tot_counter = 10 Then
               zyg_decimals = 0
             Exit Function
            End If
            If counter >= 50 Then  ' 500
                counter = 0: tot_counter = tot_counter + 1
                MSComm1.Output = "SEND" + Chr$(13)
                MSComm1.InBufferCount = 0
                FromModem$ = ""
             End If
        End If
    Loop
     zyg_decimals = 0: dec = 0
    poin = InStr(FromModem$, ".")
    For I = 1 To 20
       If IsNumeric(Mid(FromModem$, poin + I, 1)) Then dec = dec + 1 Else Exit For
     Next I
       
 
          zyg_decimals = dec
 
       
End Function
Public Function CheckForNull(ByVal arg) As Integer
On Error GoTo er_chek_val
CheckForNull = 1
 If IsNull(arg) Then CheckForNull = 0
 If IsNumeric(arg) Then
     If Val(arg * 100000000) = 0 Then CheckForNull = 0
End If
If arg = "" Then CheckForNull = 0
 Exit Function
er_chek_val: Stop
End Function

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


Sub MilSec(wait As Long)
     start = Val(GetCurrentTime())
     Do
        c_tim = Val(GetCurrentTime())
        DoEvents
     Loop Until c_tim >= start + wait
End Sub

Private Sub Change_Digits_Click()
If zyg_dec = "3" Then
   MSComm1.Output = "HI RANGE" + Chr$(13)
   MilSec 300
Else
   MSComm1.Output = "LO RANGE" + Chr$(13)
End If
End Sub

Sub robsend(cmd$)
   temp% = SendAT6400Block(768, cmd$, 0)
End Sub

Private Sub Check1_Click()
If Check1 = 1 Then
    Me.Asked_Q = 10
    Label40 = "Msec"
Else
    Me.Asked_Q = 600
     Label40 = "Target"
End If
End Sub

Private Sub Command1_Click()
List1.Clear

End Sub


Private Sub Check5_Click()
 If Check5 = 1 Then valve_1_on_Click Else valve_1_off_Click
End Sub

Private Sub Command2_Click()
Command2.Enabled = False
Check4.Enabled = False
start = GetCurrentTime()
For I = 1 To 30
cur_tim = GetCurrentTime() - start
flag_zyg = I
Next I
cur_tim = GetCurrentTime() - start
Check4.Enabled = True
Command2.Enabled = True
End Sub

Private Sub Command5_Click()
If Command5.Caption = "Stop Dosing" Then
       Command5.Caption = "Start Dosing"
    Else
      Load Syn_Dos
      Command5.Caption = "Stop Dosing"
      max_dos = 50
      For I = 1 To Val(max_dos)
      Text1 = "   Working ..."
           If Check1 = 1 Then
           Syn_Dos.Caption = "Standard Time"
           ElseIf Me.Bottle_Description = "" Then
           Syn_Dos.Caption = "Talos : Dosing Utility"
           Else: Syn_Dos.Caption = Me.Bottle_Description
           End If
             Syn_Dos.Asked_Q = Me.Asked_Q
             Syn_Dos.Level = Me.Level
             Syn_Dos.Valve_on = Me.valve_1_on
             Syn_Dos.Valve_off = Me.valve_1_off
             Syn_Dos.Show 1
              
              If Val(Me.real_q_dis) < 0 Or (Val(Me.real_q_dis) = 0 And Check1 = 0) Then
             
              DoEvents
                  GoSub send_mesage
                  Level.SetFocus
                   If Val(Me.real_q_dis) <= 0 Then
                        Command5_Click
                        Exit Sub
                 End If
             End If
        List1.AddItem Str(I) + "," + Str(Asked_Q) + ", " + Str(real_q_dis) + " ," + Str(Int((Val(last) / 1000) + 0.5)) + " ," + Str(Val(rix))
     If Command5.Caption = "Start Dosing" Then Exit For
      Next I
     If I > 15 Then Command5.Caption = "Start Dosing"
 Text1 = "   Waiting ..."
 End If
 
Exit Sub
send_mesage:
Select Case real_q_dis
Case 0
msg = "Ελλειπή Arguments" + " !"

Case -1
msg = "Μηδέν ή αρνητικό Asked_Q" + " !"
Case -2
msg = " Η βαλβίδα δεν τρέχει " + " !"
Case -3
msg = " Η ζυγαρια υπερφορτώθηκε " + " !"
Case -4
msg = "Δεν έγινε διαδικασία Εναρξης " + " !"
Case -5
msg = "Η ζυγαριά δεν ανταποκρίνεται " + " !"
Case -6
msg = "Η ζυγαριά ευρίσκεται εκτός περιοχής " + " !"
Case -7
msg = "Η ζύγιση έχει καταργηθεί " + " !"
Case -8
 msg = "Παρουσιάσθηκε απροσδόκητο λάθος στον αλγόρυθμο ζύγισης"
 Case -9
 msg = "Η ζύγιση έχει καταργηθεί " + " !"
  Case -10
 msg = "Υπερβολικά χαμηλή στάθμη φιάλης " + " !"
 Case ""
 msg = "Παρουσιάσθηκε απροσδόκητο λάθος στον αλγόρυθμο ζύγισης"
End Select
MsgBox msg, , "Talos : Dosing Function"
Return
End Sub


Private Sub Form_Activate()
Me.WindowState = 2
End Sub

Private Sub Form_Load()
system_ready = IsOSLoaded(768)
      FindValves
     
      valve_mixed_on = "!" + valve_cold_on + ":" + "!" + valve_hot_on + ":"
      valve_mixed_off = "!" + valve_cold_off + ":" + "!" + valve_hot_off + ":"
         
If CheckForNull(Ready_Bit) = 0 Then MsgBox "Δεν είναι  δηλωμένο το Ok Bit. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός = 2 ). Η διαδικασία τερματίζεται...", , "Talos : Trouble Shooting Utility.": End
If CheckForNull(skon_motor_on) = 0 Then MsgBox "Δεν είναι  δηλωμένο το output του κινητήρα της σκόνης. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός = 4 ). Η διαδικασία τερματίζεται...", , "Talos : Trouble Shooting Utility.": End
If CheckForNull(skon_motor_off) = 0 Then MsgBox "Δεν είναι  δηλωμένο το output του κινητήρα της σκόνης. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός = 4 ). Η διαδικασία τερματίζεται...", , "Talos : Trouble Shooting Utility.": End
If CheckForNull(Cold_Water_On) = 0 Then MsgBox "Δεν είναι  δηλωμένο το output της βαλβίδας του κρύου νερού. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός = 5 ). Η διαδικασία τερματίζεται...", , "Talos : Trouble Shooting Utility.": End
If CheckForNull(Cold_Water_Off) = 0 Then MsgBox "Δεν είναι  δηλωμένο το output της βαλβίδας του κρύου νερού. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός = 5 ). Η διαδικασία τερματίζεται...", , "Talos : Trouble Shooting Utility.": End
If CheckForNull(Hot_Water_On) = 0 Then MsgBox "Δεν είναι  δηλωμένο το output της βαλβίδας του ζεστού νερού. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός = 6 ). Η διαδικασία τερματίζεται...", , "Talos : Trouble Shooting Utility.": End
If CheckForNull(Hot_Water_Off) = 0 Then MsgBox "Δεν είναι  δηλωμένο το output της βαλβίδας του ζεστού νερού. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός = 6 ). Η διαδικασία τερματίζεται...", , "Talos : Trouble Shooting Utility.": End
If CheckForNull(mixer_on) = 0 Then MsgBox "Δεν είναι  δηλωμένο το output του κινητήρα της ανάδευσης. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός = 7 ). Η διαδικασία τερματίζεται...", , "Talos : Trouble Shooting Utility.": End
If CheckForNull(mixer_off) = 0 Then MsgBox "Δεν είναι  δηλωμένο το output του κινητήρα της ανάδευσης. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός = 7 ). Η διαδικασία τερματίζεται...", , "Talos : Trouble Shooting Utility.": End
If CheckForNull(Alarm1_on) = 0 Then MsgBox "Δεν είναι  δηλωμένο το outpout του Βομβητή. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός = 11 ). Η διαδικασία τερματίζεται...", , "Talos : Trouble Shooting Utility.": End
If CheckForNull(Alarm1_off) = 0 Then MsgBox "Δεν είναι  δηλωμένο το outpout του Βομβητή. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός = 11 ). Η διαδικασία τερματίζεται...", , "Talos : Trouble Shooting Utility.": End
If CheckForNull(Balance_Port) = 0 Then MsgBox "Δεν είναι  δηλωμένη η πόρτα της ζυγαριάς. ( Coordin --> Βαλβίδες Συστήματος ,  Κωδικός = 8 ).Η διαδικασία τερματίζεται...", , "Talos : Trouble Shooting Utility.": End

'anix_zygaria
'zyg_dec = zyg_decimals()
Exit Sub
er_form:
If Err = 94 Then
   MsgBox "Λείπουν στοιχεία από το αρχείο coord_tb.mdb !", , "Talos"
   Resume Next
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub



Private Sub valve_1_off_Click()
cmd$ = "!" + valve_1_off + ":"
robsend (cmd$)

End Sub

Private Sub valve_1_on_Click()
cmd$ = "!" + valve_1_on + ":"
robsend (cmd$)

End Sub




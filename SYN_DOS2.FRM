VERSION 4.00
Begin VB.Form Syn_Dos 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Talos : Δοσομέτρηση Συνταγών"
   ClientHeight    =   1410
   ClientLeft      =   2385
   ClientTop       =   2460
   ClientWidth     =   4275
   ControlBox      =   0   'False
   Height          =   1815
   Left            =   2325
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4275
   Top             =   2115
   Width           =   4395
   Begin VB.TextBox asked_q 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H0000FFFF&
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
      BackColor       =   &H000040C0&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   735
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   255
      Width           =   690
   End
   Begin VB.TextBox dif 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   1455
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   255
      Width           =   690
   End
   Begin VB.TextBox Level 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   1155
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   825
      Width           =   525
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   30
      TabIndex        =   1
      Text            =   "Bottle Content"
      Top             =   825
      Width           =   1095
   End
   Begin VB.TextBox Target_Value 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   3300
      TabIndex        =   0
      Top             =   255
      Width           =   945
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
      Left            =   2175
      TabIndex        =   21
      Top             =   255
      Width           =   600
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
      Left            =   30
      TabIndex        =   20
      Top             =   30
      Width           =   690
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
      Left            =   735
      TabIndex        =   19
      Top             =   30
      Width           =   690
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Differ."
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
      Left            =   1455
      TabIndex        =   18
      Top             =   30
      Width           =   690
   End
   Begin VB.Label Label43 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Milisec"
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
      Left            =   2175
      TabIndex        =   17
      Top             =   30
      Width           =   600
   End
   Begin VB.Label cur_timer 
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
      Height          =   225
      Left            =   2340
      TabIndex        =   16
      Top             =   570
      Width           =   435
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tare"
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
      Left            =   2790
      TabIndex        =   15
      Top             =   30
      Width           =   495
   End
   Begin VB.Label Tara_dis 
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
      Left            =   2790
      TabIndex        =   14
      Top             =   255
      Width           =   495
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Target Value"
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
      Left            =   3300
      TabIndex        =   13
      Top             =   30
      Width           =   945
   End
   Begin VB.Label rix 
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
      Height          =   240
      Left            =   3795
      TabIndex        =   12
      Top             =   570
      Width           =   450
   End
   Begin VB.Label Zyg_Show 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
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
      Height          =   225
      Left            =   825
      TabIndex        =   11
      Top             =   1140
      Width           =   30
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Action "
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
      Height          =   225
      Left            =   30
      TabIndex        =   10
      Top             =   1140
      Width           =   780
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Time Target "
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
      Height          =   225
      Left            =   1455
      TabIndex        =   9
      Top             =   570
      Width           =   855
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bottle Rate "
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
      Height          =   225
      Left            =   30
      TabIndex        =   8
      Top             =   570
      Width           =   870
   End
   Begin VB.Label Rate 
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
      Height          =   225
      Left            =   915
      TabIndex        =   7
      Top             =   570
      Width           =   495
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tryals "
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
      Height          =   225
      Left            =   2790
      TabIndex        =   6
      Top             =   570
      Width           =   975
   End
End
Attribute VB_Name = "Syn_Dos"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
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
Private Function Servirisma(Asked_Q As String, Level As String) As Single
'On Error Resume Next
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

Dim asw As String, Real_Q, tara As String
'On Error GoTo serv_error_exit
  start_dis = GetCurrentTime()
If Level = "" Then
     MsgBox "Συμπλήρωσε το περιεχόμενο του μπουκαλιού (σε Gr ) πχ 400", , "Γιάννης"
    Servirisma = 0
    Exit Function
End If
 tot_cycles = 0
Refresh
Screen.MousePointer = 11
zyg_timer.Enabled = False

If Val(Asked_Q) < 30 Then Check1 = 0
 arw = check_zyg(zygisi4(0)): If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
asw = Val(arw) * 1000


If Check1 = 1 Then
6
      rezero_Click
     arw = check_zyg(zygisi4(0)): If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
      asw = Val(arw) * 1000
       If Val(asw) <> 0 Then GoTo 6
End If

8
If Val(Asked_Q) < 100 Then
       arw = check_zyg(zygisi4("OK")): If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
      tara = Val(arw) * 1000
Else
       arw = check_zyg(zygisi4(0)): If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
      tara = Val(arw) * 1000
End If
Me.tara_dis = tara: Target_Value = tara + Val(Asked_Q)
rix = 1: Me.dif = 0: Me.last = 0: Me.cur_timer = "": first_Time_Wait = 2000: Me.Rate = ""
Real_Q = 0: real_q_dis = 0: Safety_Factor = 1: Final_target = 2: Approach_Target = False
Target_Safety = 150: Me.Zyg_Show.Width = 0
 start = Val(GetCurrentTime())
  GoSub Find_Rate
'================   Αρχή Συνάρτησης   ======================

'==========                     Μεγάλες Ποσότητες (πάνω από 400 mgr)
If Val(Asked_Q) > 400 Then
 Label1 = "Dosing ...": Label1.Refresh
 valve_1_on_Click
Do
       arw = check_zyg(zygisi4(0)): If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
      asw = Val(arw) * 1000 - Val(tara)
     fact = 0.4 * Val(Me.Level) + Val(Target_Safety)
     Me.real_q_dis = asw
     If GetCurrentTime() > start + 4000 Then
          If asw <= 0 Then
                 Servirisma = -2
                 valve_1_off_Click
                 Screen.MousePointer = 1
                 Exit Function
          End If
      End If
      Zyg_Show.Width = (Val(asw) + fact) / Val(Asked_Q) * 2500
      Zyg_Show = (Val(asw)) / Val(Asked_Q) * 100 & " %"
Loop Until Val(asw) + fact >= Val(Asked_Q)
asw1 = asw
1
last_time = Val(GetCurrentTime())
 valve_1_off_Click
Dosing_Time = last_time - start
2
     Label1 = "Waiting ...": Me.Label1.Refresh
     Me.Zyg_Show = first_Time_Wait / 1000 & " Sec"
      Me.Zyg_Show.Width = 450
       MilSec Val(first_Time_Wait)
        arw = check_zyg(zygisi4("OK")): If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
        asw = Val(arw) * 1000 - Val(tara)
If Val(asw) < Val(asw1) Then
      prosp = prosp + 1
      If prosp > 4 Then
      Servirisma = -7
      Exit Function
End If
      GoTo 2
End If
          last_q = asw - Val(Real_Q)
         If last_q < 0 Then
             arw = check_zyg(zygisi4("OK")): If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
              asw = Val(arw) * 1000 - Val(tara)
               Real_Q = asw
               Me.real_q_dis = Real_Q
               Me.dif = Val(Asked_Q) - Val(Real_Q)
               dosing_Rate = 0.4
              GoTo 3
        End If
         Real_Q = asw
         Me.real_q_dis = Real_Q

Me.dif = Val(Asked_Q) - Val(Real_Q)
Me.last = Val(GetCurrentTime()) - start
If Val(Me.dif) > 50 Then
      dosing_Rate = last_q / Dosing_Time ' Mgr/Msec
End If
Rate = Int(dosing_Rate * 1000 + 0.5) / 1000
If Val(Me.dif) < 50 And Val(Me.dif) > Final_target Then
     Approach_Target = True
End If
If dosing_Rate <= 0 Then
      sec_tim = 20: GoTo 2
Else
      sec_tim = (Me.dif / dosing_Rate) * Safety_Factor
End If
If Approach_Target = True Then
            GoSub Find_Dosing_Rate
End If
If sec_tim > 1000 Then GoTo 2
If sec_tim < Butom_Step Then
      sec_tim = Butom_Step
End If
Me.cur_timer = Int(sec_tim + 0.5)
3
If Val(Me.dif) > Final_target Then
If sec_tim <= 30 Then first_Time_Wait = 2500
 Label1 = "Dosing ...": Label1.Refresh
 valve_1_on_Click
 start = Val(GetCurrentTime())
MilSec Val(Int(sec_tim + 0.5))
rix = Val(rix) + 1
GoTo 1
End If

'============      Ποσότητες από 10 εως 400  mgr   ==============

ElseIf Val(Asked_Q) > 1 Then
  first_Time_Wait = 1000
  If Val(Asked_Q) < 20 Then first_Time_Wait = 2000
    If Val(Asked_Q) < 50 Then Final_target = 1
     GoSub Find_Rate
         Rate = Int(dosing_Rate * 1000 + 0.5) / 1000
        target1 = Asked_Q * 3 / 4
        sec_tim = target1 / dosing_Rate
        If sec_tim < Butom_Step Then sec_tim = Butom_Step
        cur_timer = Int(sec_tim + 0.5)
        Label1 = "Dosing ..."
         valve_1_on_Click
         start = Val(GetCurrentTime())
        MilSec Val(sec_tim)
4
        last_time = Val(GetCurrentTime())
        valve_1_off_Click
        Dosing_Time = last_time - start
        Label1 = "Waiting ...": Me.Zyg_Show = first_Time_Wait / 1000 & " Sec"
         Me.Zyg_Show.Width = 450
         MilSec Val(first_Time_Wait)
5
        Label1.Visible = True
        arw = check_zyg(zygisi4("OK")): If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
        asw = Val(arw) * 1000 - Val(tara)
        If asw < 0 Then GoTo 8
         last_q = asw - Val(Real_Q)
         Real_Q = asw
           Me.real_q_dis = Real_Q
            Me.dif = Val(Asked_Q) - Val(Real_Q)
        If last_q >= 2 Then dosing_Rate = last_q / Dosing_Time ' Mgr/Msec
    
        Me.last = Val(GetCurrentTime()) - start
        Rate = Int(dosing_Rate * 1000 + 0.5) / 1000
       Rate.Refresh
       If dosing_Rate <= 0 Then GoTo 5 'Else sec_tim = (Me.dif / dosing_Rate) * Safety_Factor
       GoSub Find_Dosing_Rate
      Me.cur_timer = Int(sec_tim + 0.5)
        If Val(Me.dif) > Final_target Then
          Me.Label1 = "Dosing ...": Me.Zyg_Show.Width = 0
                valve_1_on_Click
                start = Val(GetCurrentTime())
            MilSec Val(sec_tim)
            rix = Val(rix) + 1
            GoTo 4
        End If

ElseIf Val(Asked_Q) <= 0 Then
Servirisma = -1
 Screen.MousePointer = 1
 Exit Function
End If
'============      Εξοδος Συνάρτησης
Servirisma = Val(Real_Q)
last = GetCurrentTime() - start_dis
9
 zyg_timer.Enabled = True
 Screen.MousePointer = 1
Exit Function
serv_error_exit:
 valve_1_off_Click
If InStr(arw, "- 3") > 0 Then
    Servirisma = -3
    Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 6") > 0 Then
    Servirisma = -6
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

Resume 9

Find_Rate:
Select Case Level
Case Is > 600
dosing_Rate = 0.5
Butom_Step = 12
Case Is > 500
dosing_Rate = 0.45
Butom_Step = 12
Case Is > 400
dosing_Rate = 0.4
Butom_Step = 13
Case Is > 300
dosing_Rate = 0.35
Butom_Step = 13
Case Is > 200
dosing_Rate = 0.3
Butom_Step = 14
Case Is > 100
dosing_Rate = 0.25
Butom_Step = 14
Case Else
dosing_Rate = 0.2
Butom_Step = 14
End Select
Return

Find_Dosing_Rate:
Select Case sec_tim
Case Is > 700
first_Time_Wait = 1000
dosing_Rate = last_q / Dosing_Time  ' Mgr/Msec
Safety_Factor = 1
Case Is > 300
first_Time_Wait = 1500
dosing_Rate = last_q / Dosing_Time  ' Mgr/Msec
Safety_Factor = 0.97
Case Is > 100
first_Time_Wait = 2000
dosing_Rate = last_q / Dosing_Time  ' Mgr/Msec
Safety_Factor = 0.93
Case Is > 50
first_Time_Wait = 2500
Safety_Factor = 0.85
Case Is > 25
first_Time_Wait = 2500
Safety_Factor = 0.78
Case Else
first_Time_Wait = 3000
Safety_Factor = 0.72
End Select
If dosing_Rate <= 0 Then dosing_Rate = 0.5
sec_tim = (Me.dif / dosing_Rate) * Safety_Factor
If sec_tim < Butom_Step Then sec_tim = Butom_Step
Return
End Function
Function zygisi4(Zygis_Kind As Variant)
On Error GoTo er_ex
   Dim counter
 
   counter = 0: tot_counter = 0: Label10 = 0
 If system_ready = 0 Then zygisi4 = "System not Working": Exit Function
  start = GetCurrentTime()
   If Zygis_Kind = "OK" Then Label1 = "Scaling ..."
    Do
    If GetCurrentTime() - start > 15001 Then zygisi4 = " ": Exit Function
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
             Me.Zyg_Show.Width = (GetCurrentTime() - start) / 1000 * 300
              Me.Zyg_Show.Caption = Int(((GetCurrentTime() - start) / 1000) + 0.5) & " Sec"
             Me.Zyg_Show.Refresh
      End If
    Loop
         zygisi4 = Left$(FromModem$, 9)
er_ex:
Exit Function
er_det:
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

Private Sub Form_Load()

End Sub

Private Sub tara_dis_Click()

End Sub

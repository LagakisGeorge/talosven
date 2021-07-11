VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Syn_Dos 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1410
   ClientLeft      =   4425
   ClientTop       =   2925
   ClientWidth     =   4260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1410
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   1125
      Width           =   1980
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\KATOFLIA.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "katofli"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   195
      Left            =   1275
      TabIndex        =   27
      Top             =   1365
      Width           =   75
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3510
      TabIndex        =   24
      Top             =   1110
      Width           =   750
   End
   Begin VB.TextBox Valve_On 
      Height          =   285
      Left            =   1605
      TabIndex        =   23
      Top             =   1155
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox Valve_Off 
      Height          =   285
      Left            =   1020
      TabIndex        =   22
      Top             =   1125
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.TextBox asked_q 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   30
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "1000"
      Top             =   255
      Width           =   1125
   End
   Begin VB.TextBox real_q_dis 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   255
      Width           =   1020
   End
   Begin VB.TextBox dif 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2205
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   255
      Width           =   500
   End
   Begin VB.TextBox Level 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   975
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   825
      Width           =   450
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
      Width           =   960
   End
   Begin VB.TextBox Target_Value 
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
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3255
      Top             =   825
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label synarthsh 
      Caption         =   "Label2"
      Height          =   135
      Left            =   2160
      TabIndex        =   30
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Trgt_Safety 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   3900
      TabIndex        =   29
      Top             =   825
      Width           =   345
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Max Time"
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2730
      TabIndex        =   28
      Top             =   825
      Width           =   1185
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Target Approach"
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2535
      TabIndex        =   26
      Top             =   555
      Width           =   1455
   End
   Begin VB.Label Target_Approach_dis 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   3975
      TabIndex        =   25
      Top             =   555
      Width           =   270
   End
   Begin VB.Label Last 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2715
      TabIndex        =   21
      Top             =   255
      Width           =   555
   End
   Begin VB.Label Label40 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Target"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   30
      TabIndex        =   20
      Top             =   30
      Width           =   1125
   End
   Begin VB.Label Label41 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Actual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   1170
      TabIndex        =   19
      Top             =   30
      Width           =   1020
   End
   Begin VB.Label Label42 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Differ."
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   2205
      TabIndex        =   18
      Top             =   30
      Width           =   495
   End
   Begin VB.Label Label43 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Milisec"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   2715
      TabIndex        =   17
      Top             =   30
      Width           =   555
   End
   Begin VB.Label cur_timer 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2280
      TabIndex        =   16
      Top             =   825
      Width           =   435
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tare"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   2985
      TabIndex        =   15
      Top             =   1065
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Tara_dis 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2985
      TabIndex        =   14
      Top             =   1290
      Visible         =   0   'False
      Width           =   765
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
      Height          =   285
      Left            =   2100
      TabIndex        =   12
      Top             =   555
      Width           =   420
   End
   Begin VB.Label Zyg_Show 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C0C0FF&
      Height          =   270
      Left            =   825
      TabIndex        =   11
      Top             =   1125
      Width           =   30
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Action "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   30
      TabIndex        =   10
      Top             =   1125
      Width           =   780
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Time"
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Top             =   825
      Width           =   870
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bottle Rate "
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   30
      TabIndex        =   8
      Top             =   555
      Width           =   960
   End
   Begin VB.Label Rate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   975
      TabIndex        =   7
      Top             =   555
      Width           =   450
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tryals "
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   555
      Width           =   675
   End
End
Attribute VB_Name = "Syn_Dos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cancel_pressed As Boolean
Dim SafetyFactor(10)
Dim Rate_Increasing As Double, Rate_Increasing_Point As Double
Dim Low_Rate_Safety As Double, Global_Buttom_Step As Double
Dim A_Rate As Double, B_Rate As Double, Target_Safety_Break_Point As Integer
Dim Online_Shoot As Integer
Dim f_start

Private Function anaz_parox(mpoykali, katofli)
Set r = Data1.Recordset
r.Index = "katofli"
r.Seek ">=", mpoykali, Int(Int(Val(katofli)) / 10) * 10

If r.NoMatch Then
Else
   
   
   s1 = r("staumh")
   p1 = r("paroxh")
   
   r.MoveNext
   
   s2 = r("staumh")
   p2 = r("paroxh")
   
   
   
On Error GoTo DENYPARXEI

DX = (p2 - p1) / (s2 - s1)

   
   
   
   

End If
anaz_parox = (p1 + (Val(katofli) - s1) * DX) / 1000 '
Exit Function
DENYPARXEI:
anaz_parox = -1
End Function

Private Function katax_katofli(mpoykali, katofli, logos)
Dim m
katax_katofli = 1
'Exit Function



On Error Resume Next

Set r = Data1.Recordset
r.Index = "katofli"
r.Seek "<=", mpoykali, katofli

m = r("paroxh") * logos
If m > 700 Then m = 700
If m < 200 Then m = 200



If Command = "I" Then  ' Installation mode
   r.MoveFirst
   Do While Not r.eof
      r.EDIT
      r("paroxh") = r("paroxh") * logos
      r("hme") = Now
      r.update
      r.MoveNext
   Loop
   katax_katofli = 1
   Exit Function

End If



   


'If r.NoMatch Then
'   r.AddNew
'   r("mpoykali") = mpoykali
'   r("staumh") = Int(katofli)
'   r("paroxh") = Val(paroxh) * 1000
'   r("hme") = Now
'   r.Update
'Else
   r.EDIT
   r("paroxh") = m  ' r("paroxh") * logos
   r("hme") = Now
   r.update

   r.MoveNext

If r("mpoykali") = mpoykali Then
   r.EDIT
   r("paroxh") = m '  r("paroxh") * logos
   r("hme") = Now
   r.update

End If


'End If
katax_katofli = 1
End Function







Sub FindSafetyFactors()
Dim db As Database, r As Recordset
Set db = OpenDatabase("C:\Talos\skon_tb.mdb")
Set r = db.OpenRecordset("Safety_Factors")
Do Until r.eof
i = i + 1
If i < 7 Then
        SafetyFactor(i) = r("Safety_Factor")
Else
        If i = 7 Then Rate_Increasing = r("Safety_Factor")
        If i = 8 Then Rate_Increasing_Point = r("Safety_Factor")
        If i = 9 Then Low_Rate_Safety = r("Safety_Factor")
        If i = 10 Then Global_Buttom_Step = r("Safety_Factor")
        If i = 11 Then A_Rate = r("Safety_Factor")
        If i = 12 Then B_Rate = r("Safety_Factor")
        If i = 13 Then Target_Safety_Break_Point = r("Safety_Factor")
        If i = 14 Then Online_Shoot = r("Safety_Factor")
End If
r.MoveNext
Loop
r.Close
End Sub


Private Sub Command1_Click()
 Command1.Enabled = False
cancel_pressed = True
End Sub

Private Sub Form_Activate()
DoEvents
cancel_pressed = False
anix_zygaria
Me.real_q_dis = SERVIRISMA(asked_q, Level)
Unload Me
End Sub

Private Sub anix_zygaria()

'On Error Resume Next
' ΑΝΟΙΓΩ ΣΕΙΡΙΑΚΗ ΠΟΡΤΑ ΚΑΙ ΜΗΔΕΝΙΖΩ ΖΥΓΑΡΙΑ
' MSComm1.CommPort = Balance_Port
' MSComm1.Settings = Balance_Settings
' MSComm1.InputLen = 0
' MSComm1.PortOpen = False
' MSComm1.PortOpen = True
' Exit Sub

Text1.text = "Scales start Communication"
On Error GoTo err_fnd
  
  MSComm1.commport = Balance_Port
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
' system_ready = IsOSLoaded(768)
'anix_zygaria

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
 If synarthsh.Caption = "ARAIOSH" Then
   araiosh.Label18.Caption = Me.real_q_dis
 ElseIf synarthsh.Caption = "SKONES" Then
    skones.Label8.Caption = Me.real_q_dis
 Else
   frmSYNT.Label8.Caption = Me.real_q_dis
 End If
 
   MSComm1.PortOpen = False
End Sub

Sub Valve_Off_Click()
cmd$ = "!" + Valve_Off + ":"
RobSend (cmd$)

End Sub

Private Sub Valve_on_Click()
cmd$ = "!" + Valve_On + ":"
RobSend (cmd$)

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
Private Function OLDServirisma(asked_q As String, Level As String) As Single
Dim db As Database, r As Recordset
Dim asw As String, Real_Q, tara As String, Asked_Stage As Integer, Current_Dosing As String
Dim Buttom_Step_Reach As Boolean, Dosing_String As String, Real_Dosing_Time As String
Dim Bottle_Number As Integer, Bottle_Found As Boolean, Buttom_repare As Integer
Dim gram As String, Tot_Gram As String, ader As Integer
Dim Target_Safety As Integer, Counters
Dim Ejatmish, mColor1
Dim m_time, m_start, m_parox
Ejatmish = 1
max_rate = 0
On Error GoTo serv_error_exit
'Asked_Q = 1900
'=====  Σφάλματα που επιστρέφονται

'oldservirisma = 0    ==> Ελλειπή Arguments
'oldservirisma = -1   ==> Μηδέν ή αρνητικό Asked_Q
'oldservirisma = -2   ==> η βαλβίδα δεν τρέχει
'oldservirisma = -3   ==> η ζυγαρια υπερφορτώθηκε
'oldservirisma = -4   ==> δεν έγινε διαδικασία Εναρξης
'oldservirisma = -5   ==> Η ζυγαριά δεν ανταποκρίνεται
'oldservirisma = -6   ==> Η ζυγαριά εκτός περιοχής
'oldservirisma = -7   ==> Η ζύγιση έχει αποτύχει
'oldservirisma = -8   ==> Παρουσιάσθηκε απροσδόκητο λάθος στον αλγόρυθμο ζύγισης
'oldservirisma = -9   ==> Η ζύγιση ακυρώθηκε
'oldservirisma = -10   ==> Η στάθμη του μπουκαλιού είναι πολύ χαμηλή <50 gr
'oldservirisma = -11   ==> Υπερβολική  ζητούμενη ποσότητα  (max. 100 gr)
start_dis = GetCurrentTime()



If Level = "" Then
    Level = 300
    'Screen.MousePointer = 1
    'oldservirisma = 0
    'Exit Function
End If

If Level < 50 Then
    Screen.MousePointer = 1
    OLDServirisma = -10
    Exit Function
End If
tot_cycles = 0: tryal = 0
Refresh

   Screen.MousePointer = 11
   'MSComm1.Output = Balance_3_Digits + Chr$(13)
    G_Balance_Digits MSComm1, 3
    
Counters = 0
8
If Val(asked_q) < 100 Then
       arw = diplo_zygi(1000, 0.002) ': If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
       'arw = check_zyg(zygisi4("OK")): If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
       tara = Val(arw) * 1000
'       GoTo 8  '22-10
Else
       arw = diplo_zygi(500, 0.004) ': If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
        ' arw = check_zyg(zygisi4(0))
       'If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
       tara = Val(arw) * 1000
 '      GoTo 8  '22-10
End If

Counters = Counters + 1
If tara < 200 Then
   If Counters >= 2 Then GoTo 1961   ' serv_error_exit
   MilSec 3000
   GoTo 8
End If
1961


'A_Rate = 0.000567: B_Rate = 0.26

Me.Tara_dis = tara: Target_Value = tara + Val(asked_q): count_val = 0
Buttom_Step_Reach = False
rix = 1: Me.dif = 0: Me.Last = 0: Me.cur_timer = "": Me.Rate = ""
Real_Q = 0: real_q_dis = 0: Safety_Factor = 1: FINAL_TARGET = 2: Approach_Target = False
Me.Zyg_Show.width = 0: Ocasion = 0
first_Time_Wait = 2000: Dosing_String = " ": Real_Dosing_Time = " "
max_time_wait = Val(asked_q) * 40

GoSub Find_Bottle_Number
   Dosing_Rate = A_Rate * Val(Level) + B_Rate
   max_rate = Dosing_Rate
GoSub Find_Bottle_Buttom_Step
Rate = Int(Dosing_Rate * 1000 + 0.5) / 1000
GoSub Find_Final_target
  
  Target_Approach_dis = FINAL_TARGET
  If FINAL_TARGET = 0 Then
    Screen.MousePointer = 1
    OLDServirisma = -11
    Exit Function
  End If

'============================   Αρχή Συνάρτησης   ======================
     fact = 1000 ' Val(Me.Level) + Val(Target_Safety)   2-11-2000

'==================                Μεγάλες Ποσότητες (πάνω από 400 mgr)
If Val(asked_q) >= fact Then
Ocasion = 1
 Label1 = "Dosing ...": Label1.Refresh
Valve_on_Click
  start = Val(GetCurrentTime())
 
 Select Case Val(asked_q)
      Case Is < 1500
          m_syntelesths = 0.5
      
      Case Is < 3000
          m_syntelesths = 0.5  '0.9   22/10
          
      Case Is < 5000
         m_syntelesths = 0.5
      Case Is >= 5000
           m_syntelesths = 0.8
      End Select
'  m_syntelesths = 0.9 ' ακυρωνει το παραπανω select




mColor1 = Rate.BackColor
 Rate.BackColor = &HFFFF00  'OYRANI
MilSec 1500  '<1500 //δισπ

Do
 
      arw = check_zyg(zygisi4(0))
      If InStr(arw, "Error") > 0 Then
         '=============
          Valve_Off_Click
         '=============
         JANA = 0
         Do While JANA < 5
             MilSec 500
            arw = check_zyg(zygisi4(0))
             If InStr(arw, "Error") > 0 Then
                  JANA = JANA + 1
             Else
                  Valve_on_Click
                  Exit Do
             End If
             
         Loop
         If JANA >= 5 Then GoTo serv_error_exit
      End If
      
      asw = Val(arw) * 1000 - Val(tara) 'καθαρό βάρος
      Me.real_q_dis = Int(asw + 0.5)
      Me.dif = Val(asked_q) - Val(real_q_dis) 'διαφορά από στόχο
     If cancel_pressed = True Then
              OLDServirisma = -9
                Valve_Off_Click
                 Screen.MousePointer = 1
                 Exit Function
      End If
     If GetCurrentTime() > start + max_time_wait Then
                  OLDServirisma = -2
                 Valve_Off_Click
                 Screen.MousePointer = 1
                 Exit Function
      End If
       ' δείχνει την μπάρα που γεμίζει
      If (Val(asw) + fact) / Val(asked_q) * 2500 > 0 Then Zyg_Show.width = Int((Val(asw) + fact) / Val(asked_q) * 2500)
      Zyg_Show = Int((Val(asw)) / Val(asked_q) * 100) & " %"
Loop Until Val(asked_q) - Val(asw) < 1500 'Val(asw) + FACT >= (Val(Asked_Q) + FINAL_TARGET) * m_syntelesths '  * 0.9  ' 23-4-2000


      Rate.BackColor = mColor1





' fact = Val(Me.Level) + Val(Target_Safety) περιθώριο που πρέπει να έχω στην online ζύγιση
' Final_target : επιτρεπόμενη απόκλιση απο τον στόχο 2μγρ μέχρι τα 1000μγρ
'l_count = 1
asw1 = asw
1
      'l_count = l_count + 1
      
      '=============
      Valve_Off_Click
      '=============
      Dosing_String = Dosing_String + left(str(Dosing_Rate), 4)
      last_time = Val(GetCurrentTime())  ' 1o kl k loop
      Dosing_Time = last_time - start
      Real_Dosing_Time = Real_Dosing_Time + str(Dosing_Time) + " "
2
      If cancel_pressed = True Then arw = "Error - 9": GoTo serv_error_exit
      Label1 = "Waiting ...": Me.Label1.Refresh
      
      ' μπάρα με γέμισμα
      Me.Zyg_Show = first_Time_Wait / 1000 & " Sec"
      Me.Zyg_Show.width = 450
      
      MilSec Val(first_Time_Wait)
      
      
      
      
      
      
      
      
      
      
      
      
      ' MilSec 1000   ' 22-10
'1960
 '     arw = check_zyg(zygisi4("OK")): If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
  '    MilSec 1000
   '   arw2 = check_zyg(zygisi4("OK")): If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
    '  If Abs(Val(arw2) - Val(arw)) > 0.01 Then GoTo 1960
      
      arw = diplo_zygi(1500, 0.005): If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
      
      count_val = count_val + 1
      
      If count_val > 48 Then   '14   6-3-2000
                 OLDServirisma = -7
                 Valve_Off_Click
                 Screen.MousePointer = 1
                 Exit Function
      End If
        
      asw = Val(arw) * 1000 - Val(tara) 'καθαρό βάρος
      real_q_dis = Int(asw + 0.5)
      Me.dif = Val(asked_q) - Val(real_q_dis)
      
      
      If Val(asked_q) >= 5000 And Dosing_Time > 0 And asw > 0 Then
          m_parox = asw / Dosing_Time
          If m_parox > 400 Then
             Label16.BackColor = &HC0C000  'skoyro ble
             
             If Val(asked_q) < 5000 Then
                 m_time = (Me.dif - 90) / m_parox
             Else
                 m_time = Me.dif / m_parox
             End If
             
             m_start = GetCurrentTime()
             '=============
             Valve_on_Click
             '=============
             Do While GetCurrentTime() < m_start + m_time
             Loop
             '=============
             Valve_Off_Click      'xrhstakos
             '=============
             arw = diplo_zygi(1500, 0.005): If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
             asw = Val(arw) * 1000 - Val(tara) 'καθαρό βάρος
             real_q_dis = Int(asw + 0.5)
             Me.dif = Val(asked_q) - Val(real_q_dis)
          End If
      End If
      
      'Αλλαγή του Trgt_Safety
           If Val(rix) = 1 Then
                If Me.dif - FINAL_TARGET < 0 Then
                      Target_Safety = Target_Safety + 25
                      Trgt_Safety = Target_Safety
                 ElseIf Me.dif > Target_Safety_Break_Point Then
                      Target_Safety = Target_Safety - 25
                      Trgt_Safety = Target_Safety
                End If
            End If

' δεν έρριξε τίποτα , συνεχίζω ανοιχτός
If Val(asw) < Val(asw1) Then  'asw1:προηγούμενη ζύγιση  asw:τωρινή
       prosp = prosp + 1
       asw1 = asw
       If prosp > 34 Then    '7    6-3-2000
          OLDServirisma = -7
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
               Me.dif = Val(asked_q) - Val(Real_Q)
               Dosing_Rate = 0.55
               
              GoTo 3
         End If
         Real_Q = Int(asw + 0.5)
         Me.real_q_dis = Real_Q

Me.dif = Val(asked_q) - Val(Real_Q)
' 4-5-2001  If Val(Me.dif) > Online_Shoot Then Me.dif = Online_Shoot
Me.Last = Val(GetCurrentTime()) - start
If Val(Me.dif) > 50 And last_q / Dosing_Time > Dosing_Rate Then   '*0.9 22/10
   Dosing_Rate = (last_q / Dosing_Time)  ' Mgr/Msec   22-10
   If Dosing_Rate > max_rate Then max_rate = Dosing_Rate
End If
'22-10 End If


Rate = Int(Dosing_Rate * 1000 + 0.5) / 1000

If Val(Me.dif) < 50 And Val(Me.dif) > FINAL_TARGET Then
     Approach_Target = True  'εχω φτάσει κοντά στον στόχο
End If

If Dosing_Rate <= 0 Then
      sec_tim = 20: GoTo 2  ' συνέχισε ανοιχτός
Else
   
   If Val(asked_q) < 2000 Then  ' 2001/10
      sec_tim = ((Me.dif) / Dosing_Rate) * Safety_Factor * 0.2    ' 2-11-2000
   Else ' 2-11-2000
      sec_tim = ((Me.dif) / Dosing_Rate) * Safety_Factor ' 26-10
   End If ' 2-11-2000
      
   If Me.dif < 150 Then sec_tim = sec_tim / 3  'lagakis 6-3
End If

If Approach_Target = True Then
    If Val(Me.dif) > FINAL_TARGET Then
      If Buttom_Step_Reach = True And last_q <= 1 Then
         Buttom_step = Buttom_step + 2  ' αν δεν ρίχνει ανέβασε το κατώφλι
      End If
    End If
End If
    If Val(Me.dif) > FINAL_TARGET Then GoSub Find_Safety_Factor
    If sec_tim > 10000 Then GoTo 2
    
    If sec_tim <= 30 And Val(rix) > 1 Then
        sec_tim = Buttom_step + Val(Me.dif) / 2 - 1
        Buttom_Step_Reach = True ' μπήκαμε στα μικρά
        Me.cur_timer = Int(sec_tim + 0.5)
    End If
Me.cur_timer = Int(sec_tim + 0.5)
3

If cancel_pressed = True Then arw = "Error - 7":  GoTo serv_error_exit

If Val(Me.dif) > FINAL_TARGET Then
    If sec_tim <= 30 Then first_Time_Wait = 2500
    If sec_tim >= max_time_wait Then GoTo 2
    
    ' χρόνος μικρότερος από το κατώφλι άρα ακολουθώ το κατώφλι
    If sec_tim <= Buttom_step Then
         sec_tim = Buttom_step + Val(Me.dif) / 2 - 1
         Me.cur_timer = Int(sec_tim + 0.5)
    End If
    
    
    If Me.dif < 30 And asked_q < 3000 Then
          'for-next
    
           For ll = 1 To 90
                   Valve_on_Click
                        'If Me.dif > 10 Then
                                mFornext_big = Fornext_big * 1.9
                                For k = 1 To mFornext_big: Next  ' 5000
                        'Else
                         '       For k = 1 To Fornext_small: Next  ' 100
                       ' End If
                   Valve_Off_Click
                   
                   MilSec 1000
    
                   asw = 1000 * diplo_zygi(500, 0.002)
                
                   last_q = Val(asw) - Val(tara) - Real_Q
                   
                   Real_Q = Val(asw) - Val(tara)
                   
                   If cancel_pressed = True Then arw = "Error - 9":  GoTo serv_error_exit

                   
                   MilSec 500
                   Me.real_q_dis = Real_Q
                   Me.dif = Val(asked_q) - Val(Real_Q)
                   If Val(Me.dif) < FINAL_TARGET Then GoTo 119
                   
                 Next
    
    
    
    
    
    
    
    
    
    
    
    End If
    
    
    count_val = 0
    Dosing_String = Dosing_String + str(Real_Q) + " "
    Label1 = "Dosing ...": Label1.Refresh
    
    
    '============
     Valve_on_Click
    '============
    start = Val(GetCurrentTime())
    MilSec Val(Int(sec_tim + 0.5))
    rix = Val(rix) + 1  ' ριξιές
    If rix > 30 Then arw = "Error - 7":   GoTo serv_error_exit
    GoTo 1




End If











ElseIf Val(asked_q) > 1 And Val(asked_q) < 2 Then

     start = Val(GetCurrentTime())

     Me.dif = Val(asked_q)
     tara = 1000 * zygisi4(0)


     Do While Val(Me.dif) > FINAL_TARGET
                  Valve_on_Click
                        ' SELECT CASE
                        If Me.dif > 25 Then
                             If Me.dif > 80 Then
                                MilSec 4   '1
                             ElseIf Me.dif > 40 Then
                                MilSec 2
                             Else
                                MilSec 1
                                'For k = 1 To 120000: Next  ' MilSec 10
                             End If
                        Else
                                 If Me.dif > 10 Then
                                   For k = 1 To 5000: Next  ' 5000
                                 Else
                                   For k = 1 To 100: Next  ' 100
                                 End If
                        
                        
                                 ' MilSec 200 / ForNexts_Milsec
                                 Me.Caption = "...."
                                'For k = 1 To 200: Next  ' MilSec 10
                        End If
                   Valve_Off_Click
                   MilSec 500
                   
                   
                   Real_Q = 1000 * zygisi4("OK") - Val(tara)
                   Me.real_q_dis = Real_Q
                   Me.dif = Val(asked_q) - Val(Real_Q)
                   If Val(Me.dif) < FINAL_TARGET Then GoTo 119
                   If GetCurrentTime() - start > 500000 Then GoTo 119
      Loop

'============      Ποσότητες από 1 εως 400  mgr   ==============

ElseIf Val(asked_q) >= 0 And Val(asked_q) < fact Then
        
        Ocasion = 2
        first_Time_Wait = 1000
        If Val(asked_q) < 30 Then first_Time_Wait = 3000
        sec_tim = Val(asked_q) / Dosing_Rate * Low_Rate_Safety * 0.5
        If sec_tim < 30 Then GoSub Find_Safety_Factor
        If sec_tim < Buttom_step Then sec_tim = Buttom_step
        cur_timer = Int(sec_tim + 0.5)
        Label1 = "Dosing ..."
        '=============
        Valve_on_Click
        '=============
        start = Val(GetCurrentTime())
        MilSec Val(sec_tim)
4
        If cancel_pressed = True Then arw = "Error - 9":  GoTo serv_error_exit
        
        last_time = Val(GetCurrentTime())
        '==============
        Valve_Off_Click
        '==============
        Dosing_Time = last_time - start:  Real_Dosing_Time = Real_Dosing_Time + str(Dosing_Time) + " "
        Label1 = "Waiting ...": Me.Zyg_Show = first_Time_Wait / 1000 & " Sec"
        Me.Zyg_Show.width = 450
        MilSec Val(first_Time_Wait)
5
       If cancel_pressed = True Then arw = "Error - 9":  GoTo serv_error_exit
          Label1.Visible = True
          MilSec 500
          'ARW = check_zyg(zygisi4("OK"))
          
';
 '     arw = check_zyg(zygisi4("OK")): If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
  '    MilSec 1000
   '   arw2 = check_zyg(zygisi4("OK")): If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
    '  If Abs(Val(arw2) - Val(arw)) > 0.002 Then GoTo
            
1964
          arw = diplo_zygi(2000, 0.005): If InStr(arw, "Error") > 0 Then GoTo serv_error_exit '1000 htan stis 1-11-2000
          '  If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
          asw = Val(arw) * 1000 - Val(tara)
          last_q = asw - Val(Real_Q)
      If last_q <= 0 Then
            ' 3-11-2000
            If last_q < -1000 Then GoTo 1964
            If Ejatmish = 1 Then   ' 1 labainei ypoch toy thn ejatmish
                tara = Val(tara) + last_q
            End If
            tryal = tryal + 1
            If tryal > 10 Then arw = "Error - 2":  GoTo serv_error_exit
       Else
            tryal = 0
      End If
         Real_Q = Int(asw + 0.5)
         Me.real_q_dis = Real_Q
         Me.dif = Val(asked_q) - Val(Real_Q)
         If Val(Me.dif) > Online_Shoot Then Me.dif = Online_Shoot
         If Val(Me.dif) < 50 Then Buttom_Step_Reach = True
         Me.Last = Val(GetCurrentTime()) - start
       
       If last_q <= 0 Then
            tara = Val(Tara_dis) + last_q: Tara_dis = tara
            Target_Value = tara + Val(asked_q)
            Buttom_step = Buttom_step + 2
            GoTo 5
       End If
      
        If Val(Me.dif) > FINAL_TARGET Then GoSub Find_Safety_Factor
           Me.cur_timer = Int(sec_tim + 0.5)
        
        If Val(Me.dif) > FINAL_TARGET Then
           If sec_tim > max_time_wait Then
                tryal = tryal + 1
                If tryal > 10 Then
                        arw = "Error - 2":   GoTo serv_error_exit
                 Else
                      GoTo 5
                End If
            End If
            Dosing_String = Dosing_String + str(Real_Q) + " "
            Me.Label1 = "Dosing ...": Me.Zyg_Show.width = 0
            
            
            If last_q > 2 Or Val(asked_q) < 90 Then
               If Me.dif < 40 Or Val(asked_q) < 90 Then
                  Me.Label1 = "Dosing ---..."
                  For ll = 1 To 90
                   Valve_on_Click
                        If Me.dif > 10 Then
                                
                                For k = 1 To Fornext_big: Next  ' 5000
                        Else
                                For k = 1 To Fornext_small: Next  ' 100
                        End If
                   Valve_Off_Click
              MilSec 1000
1962
                
                   asw = 1000 * diplo_zygi(500, 0.002)
                
                   last_q = Val(asw) - Val(tara) - Real_Q
                   If last_q < 0 Then
                       If last_q < -500 Then GoTo 1962
                       If Ejatmish = 1 Then   ' 1 labainei ypoch toy thn ejatmish
                          tara = Val(tara) + last_q
                       End If
                   End If
                   Real_Q = Val(asw) - Val(tara)
                   
                   If cancel_pressed = True Then arw = "Error - 9":  GoTo serv_error_exit

                   
                   MilSec 500
                   Me.real_q_dis = Real_Q
                   Me.dif = Val(asked_q) - Val(Real_Q)
                   If Val(Me.dif) < FINAL_TARGET Then GoTo 119
                 Next
               Else
                   'SEC_TIM = SEC_TIM * 0.1
               End If
            End If
            '=============
             Valve_on_Click
            '=============
            start = Val(GetCurrentTime())
            MilSec Val(sec_tim)
            rix = Val(rix) + 1
            If rix > 30 Then arw = "Error - 7":   GoTo serv_error_exit
            GoTo 4
    End If

ElseIf Val(asked_q) <= 0 Then
    OLDServirisma = -1
    Screen.MousePointer = 1
     Exit Function
End If





'============      Εξοδος Συνάρτησης ==============================================
119
Last = GetCurrentTime() - start_dis
 Dosing_String = Dosing_String + str(Real_Q) + " "
9
'On Error Resume Next
Set db = OpenDatabase("C:\Talos\skon_tb.mdb")
Set r = db.OpenRecordset("Syntages")
r.AddNew
If Me.Caption <> "" Then r("Description") = left(Me.Caption, 50) Else r("Description") = "Botlle Unknown"
r("Level") = Int(Val(Me.Level))
r("Tryals") = Val(rix)
r("Asked_Q") = Val(asked_q)
If Abs(Val(Me.real_q_dis) - Val(asked_q)) <= FINAL_TARGET Then target_fin = Val(asked_q)
If Val(Me.real_q_dis) < Val(asked_q) - FINAL_TARGET Then target_fin = Val(Me.real_q_dis) + FINAL_TARGET
If Val(Me.real_q_dis) > Val(asked_q) + FINAL_TARGET Then target_fin = Val(Me.real_q_dis) - FINAL_TARGET
r("Actual_q") = target_fin
r("Lasted") = Int(Val((Me.Last)) / 1000 + 0.5)
r("Level_Stage") = Int((Val(Level) + 50) / 100)
r("Asked_Stage") = Asked_Stage
r("Dosing_String") = Dosing_String + Real_Dosing_Time
r.update
r.Close

   '======================= Ενημέρωση Αρχείου Buttom Step   =======================
   If Bottle_Number > 0 Then
  Set r = db.OpenRecordset("Bottle_Buttom_Steps", dbOpenDynaset)
   If Val(asked_q) > 400 Then ader = 3 Else ader = 6
  If Bottle_Found = True Then
    cret = "[Bottle_Number]=" & Bottle_Number
    r.FindFirst cret
    r.EDIT
        If Buttom_Step_Reach = True Then
            dior = (Val(asked_q) - target_fin) / 2
            ade = (Buttom_step + Buttom_repare * 2) / 3
            asd = Int(ade + Val(rix) - ader + dior)
             If asd > 40 Then asd = 40
             If asd < Global_Buttom_Step Then asd = Global_Buttom_Step
        End If
       If asd > 0 Then r("Buttom_Step") = asd Else r("Buttom_Step") = Global_Buttom_Step
       r("Target_Safety") = Target_Safety
       Select Case Val(Level)
             Case Is > 450: r("Rate_500") = Dosing_Rate + A_Rate * (500 - Val(Level))
             Case Is > 350: r("Rate_400") = Dosing_Rate + A_Rate * (400 - Val(Level))
             Case Is > 250: r("Rate_300") = Dosing_Rate + A_Rate * (300 - Val(Level))
             Case Is > 150: r("Rate_200") = Dosing_Rate + A_Rate * (200 - Val(Level))
             Case Else: r("Rate_100") = Dosing_Rate + A_Rate * (100 - Val(Level))
       End Select

    r.update
   Else
    r.AddNew
         r("Target_Safety") = Target_Safety
        If Buttom_Step_Reach = True Then
              dior = (Val(asked_q) - target_fin) / 2
              ade = (Buttom_step + Buttom_repare * 2) / 3
              asd = Int(ade + Val(rix) - ader + dior)
              If asd > 40 Then asd = 40
              If asd < Global_Buttom_Step Then asd = Global_Buttom_Step
        End If
        If asd > 0 Then r("Buttom_Step") = asd Else r("Buttom_Step") = Global_Buttom_Step
       If Bottle_Number > 0 Then r("Bottle_Number") = Bottle_Number
   r.update
   End If
   r.Close
 End If
 Screen.MousePointer = 1
Exit Function

 '===================  Τέλος Συνάρτησης  ============================
serv_error_exit:
'oldservirisma = 0    ==> Ελλειπή Arguments
'oldservirisma = -1   ==> Μηδέν ή αρνητικό Asked_Q
'oldservirisma = -2   ==> η βαλβίδα δεν τρέχει
'oldservirisma = -3   ==> η ζυγαρια υπερφορτώθηκε
'oldservirisma = -4   ==> δεν έγινε διαδικασία Εναρξης
'oldservirisma = -5   ==> Η ζυγαριά δεν ανταποκρίνεται
'oldservirisma = -6   ==> Η ζυγαριά εκτός περιοχής
'oldservirisma = -7   ==> Η ζύγιση έχει αποτύχει
'oldservirisma = -8   ==> Παρουσιάσθηκε απροσδόκητο λάθος στον αλγόρυθμο ζύγισης
'oldservirisma = -9   ==> Η ζύγιση ακυρώθηκε
'oldservirisma = -10   ==> Η στάθμη του μπουκαλιού είναι πολύ χαμηλή <50 gr
'oldservirisma = -11   ==> Υπερβολική  ζητούμενη ποσότητα  (max. 100 gr)

'anix_zygaria
If Err > 0 Then
      Resume Next
      arw = "- 8"
End If
 Valve_Off_Click
If InStr(arw, "- 2") Then
    OLDServirisma = -2
    Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 3") > 0 Then
    OLDServirisma = -3
    Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 8") > 0 Then
    OLDServirisma = -8
    Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 6") > 0 Then
    OLDServirisma = -6
    Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 7") > 0 Then
    OLDServirisma = -7
    Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 4") Then
    OLDServirisma = -4
    Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 5") Then
    OLDServirisma = -5
     Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 9") Then
    OLDServirisma = -9
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

If Val(asked_q) < 600 Then
  sec_tim = ((Me.dif - 50) / Dosing_Rate) * Safety_Factor ' * 0.7  26-10
Else
  sec_tim = ((Me.dif) / Dosing_Rate) * Safety_Factor ' * 0.7  26-10
End If
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
Select Case Val(asked_q)
Case Is < 50
FINAL_TARGET = 2
Asked_Stage = 1

Case Is < 200
FINAL_TARGET = 2
Asked_Stage = 2

Case Is < 1000
FINAL_TARGET = 3
Asked_Stage = 3
Case Is < 2000
FINAL_TARGET = Val(asked_q) * 0.2 / 100  '3   22-10
Asked_Stage = 4
Case Is < 3000
FINAL_TARGET = Val(asked_q) * 0.2 / 100  '   22-10 15
Asked_Stage = 5
Case Is < 4000
FINAL_TARGET = Val(asked_q) * 0.2 / 100  '   22-10 20
Asked_Stage = 6
Case Is < 10000
FINAL_TARGET = Val(asked_q) * 0.002
Asked_Stage = 7
Case Is < 25000
FINAL_TARGET = Val(asked_q) * 0.002
Asked_Stage = 8
Case Is < 100001
FINAL_TARGET = Val(asked_q) * 0.002
Asked_Stage = 9
Case Else
FINAL_TARGET = 0
Asked_Stage = 0
End Select
Return

Find_Bottle_Buttom_Step:
Set db = OpenDatabase("C:\Talos\skon_tb.mdb")
Set r = db.OpenRecordset("Bottle_Buttom_Steps", dbOpenDynaset)
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
Bottle_Number = Val(Me.Caption)
Return

End Function

Function zygisi4(Zygis_Kind As Variant)

On Error GoTo er_det
   Dim counter, m_start
 
   If Balance_Type = "ADAM" Then
      m_start = Format(Zygis0_ADAM(MSComm1, 0) / 1000, "#####.000")
      zygisi4 = m_start
      Exit Function
   End If
 
 
' On Error GoTo 0
   counter = 0: tot_counter = 0: 'Label10 = 0
 If system_ready = 0 Then zygisi4 = "System not Working": Exit Function
  m_start = GetCurrentTime()
  
MSComm1.Output = "MW 3 2 2 0 1 0 0 0 0 2 1. " + Chr(34) + "[C]" + Chr(34) + " 0" + Chr(10) + Chr(13)
MilSec 30
  
  ' MSComm1.Output = "UPD 10" + Chr$(13) + Chr(10)
   MSComm1.Output = "SIR " + Chr$(13) + Chr(10)
   
If Zygis_Kind = "OK" Then
      Label1 = "Scaling ..."
'      MSComm1.Output = "SIRU" + Chr$(13) + Chr(10)
Else
 '  MSComm1.Output = "SIR" + Chr$(13) + Chr(10)
End If

    Do
    If GetCurrentTime() - m_start > 2000 Then  '22-10
        zygisi4 = " ": Exit Function
    End If
     MSComm1.InBufferCount = 0
      FromModem$ = ""
   'On Error GoTo 0
  
       MilSec (40)
      dummy = DoEvents()
       If MSComm1.InBufferCount Then
           BUF = MSComm1.InBufferCount
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
                     
          
          If InStr(FromModem$, Chr$(13)) Then
                FromModem$ = LTrim(Right$(FromModem$, 20))
                If (InStr(FromModem$, ".") Or InStr(6, FromModem$, ".")) And InStr(6, FromModem$, "g") > 10 Then
                       Exit Do
                End If
          End If
           
           'If InStr(FromModem$, Chr$(13)) And (InStr(FromModem$, ".") > 5 Or InStr(6, FromModem$, ".") > 5) And InStr(6, FromModem$, "g") > 10 Then
           '         Exit Do
           'End If
       Else
          If InStr(FromModem$, Chr$(13)) Then
            FromModem$ = LTrim(Right$(FromModem$, 20))
                If (InStr(FromModem$, ".") Or InStr(6, FromModem$, ".")) And InStr(6, FromModem$, "g") > 10 Then
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
              MSComm1.Output = Balance_Asking + Chr$(13) + Chr(10)
              End If
     End If
     If Zygis_Kind = "OK" Then
             Me.Zyg_Show.width = (GetCurrentTime() - start) / 1000 * 650
              Me.Zyg_Show.Caption = Int(((GetCurrentTime() - start) / 1000) + 0.5) & " Sec"
             Me.Zyg_Show.Refresh
      End If
    Loop
    
      If InStr(FromModem$, ".") < 5 Then
         x = InStr(6, FromModem$, ".")
         Y = InStr(x, FromModem$, "g")
      Else
         x = InStr(FromModem$, ".")
         Y = InStr(FromModem$, "g")
      End If

      'Zygis0_time = Val(Mid$(FromModem$, X - 6, 10)) * 1000 'If Val(Label2) < 1 Then
          
    
    
    
    
'      If InStr(FromModem$, ".") < 5 Then
'         X = InStr(6, FromModem$, "  ")
'         Y = InStr(X, FromModem$, "g")
'      Else
'         X = InStr(FromModem$, "  ")
'         Y = InStr(FromModem$, "g")
'      End If

         zygisi4 = Mid$(FromModem$, x - 6, 10)
         'mx = 0
         'For k = 1 To 10
          '  If Not IsNumeric(Mid$(FromModem$, X - 6, 1)) Then mx = mx + 1
         'Next
         'If mx > 4 Then
         '    mx = 0
         'End If
    
    
    
    '    zygisi4 = left$(FromModem$, 9)
er_ex:
Exit Function
er_det:
 zygisi4 = " "
Resume er_ex
End Function

Sub RobSend(cmd$)
   TEMP% = SendATBlock(768, cmd$, 0)
End Sub

Function diplo_zygi(mS, akriv)
Dim mC, arw, arw2
' ms miliseconds katisterisi , akriv akrivia  0.001 = 1 mgr

mC = 0
23460
      mC = mC + 1
      If mC > 30 Then
         diplo_zygi = "Error"
         Exit Function
      End If
      arw = Val(check_zyg(zygisi4(0)))
      
      If mS = 1000 Then
         MilSec 1000
      ElseIf mS = 500 Then
         MilSec 500
      ElseIf mS = 1500 Then
         MilSec 1500
      ElseIf mS = 2000 Then
         MilSec 2000
      Else
          MilSec 700
      End If
      
      arw2 = Val(check_zyg(zygisi4(0)))
      
      cur_timer.Caption = Int((GetCurrentTime() - f_start) / 1000)
      
      'If cur_timer.BackColor = vbMagenta Then cur_timer.BackColor = vbYellow Else cur_timer.BackColor = vbMagenta
      
      If Abs(Val(arw2) - Val(arw)) > akriv Then GoTo 23460
      
      diplo_zygi = arw2
      
      
      
End Function
Sub MilSec(WAIT As Long)
     start = Val(GetCurrentTime())
     Do
        c_tim = Val(GetCurrentTime())
        DoEvents
     Loop Until c_tim >= start + WAIT
End Sub


'
Private Function SERVIRISMA(asked_q As String, Level As String)
'17-10-02 προστέθηκε     If elax_xronos > 5000 Then elax_xronos = 5000  ' 17-10-2002 για να μην βγαζει λαθος χρόνους
'
'
'
'
'
'
'
'

Dim db As Database, r As Recordset
Dim asw As String, Real_Q, tara As String, Asked_Stage As Integer, Current_Dosing As String
Dim Buttom_Step_Reach As Boolean, Dosing_String As String, Real_Dosing_Time As String
Dim Bottle_Number As Integer, Bottle_Found As Boolean, Buttom_repare As Integer
Dim gram As String, Tot_Gram As String, ader As Integer
Dim Target_Safety As Integer, Counters
Dim Ejatmish, mColor1, AVANCE_ONLINE_SYNTAGHS
Dim m_time, m_start, m_parox
Dim Start2, rate1, STR_DOS
Dim mpoykali, metrhma_1_balbidas, start, zht
Dim EINAI_Ximiko As Integer
Dim m_Asw2, dz, acumMul
Dim abGaltos
abGaltos = 0

On Error Resume Next
' //////////////////////////////////////
max_time_wait = Val(asked_q) / 1000 * 5   ' 10 SEC/GR

mpoykali = Val(Syn_Dos.Caption)
If InStr(Syn_Dos.Caption, "&&") Then EINAI_Ximiko = 1 Else EINAI_Ximiko = 0
metrhma_1_balbidas = 0
STR_DOS = ""

 If Balance_Type = "ADAM" Then
    AVANCE_ONLINE_SYNTAGHS = 1200
 Else
    AVANCE_ONLINE_SYNTAGHS = 1600
 End If



' ELEGXOS LEVEL
If Level = "" Then Level = 300
If Level < 50 Then Screen.MousePointer = 1: SERVIRISMA = -10: Exit Function
G_Balance_Digits MSComm1, 3   ' 3 DIGITS
Counters = 0
8
If Val(asked_q) < 100 Then
       arw = diplo_zygi(1000, 0.002) ': If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
       tara = Int(Val(arw) * 1000 + 0.5)
Else
       arw = diplo_zygi(500, 0.004) ': If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
       tara = Int(Val(arw) * 1000 + 0.5)
End If
Text1.text = "Tare " + str(Counters)
Counters = Counters + 1
If tara < 1 Then 'debug htan 200  22-10-2001              '  ==> Η ζυγαριά δεν ανταποκρίνεται
   If Counters >= 2 Then SERVIRISMA = -5: GoTo serv_error_exit
   MilSec 3000
   GoTo 8
End If

start = GetCurrentTime()
f_start = start

Text1.text = "Dispensing..."
'=====================================================================
'=====================================================================
'=====================================================================
'=====================================================================
If asked_q >= 2000 Then
'=====================================================================
'=====================================================================
'=====================================================================
'=====================================================================
  max_time_wait = Val(asked_q) * 10   ' 10 SEC/GR
  If max_time_wait < 90000 Then
     max_time_wait = 90000
  End If
  Label1 = "Dosing ..."
  
asw = 0
If EINAI_Ximiko Then AVANCE_ONLINE_SYNTAGHS = AVANCE_ONLINE_SYNTAGHS / 2


Trgt_Safety.Caption = Int(max_time_wait / 1000)


'on-line
'=============
Valve_on_Click
'=============
Do
      asw = Int(1000 * Val(check_zyg(zygisi4(0)))) - tara 'καθαρό βάρος
      '      asw = 1000 * Val(zygisi4(0)) - tara 'καθαρό βάρος
      
      If asw < -100 Then
         '=============
          Valve_Off_Click
         '=============
          Do While asw < -100
             asw = 1000 * Val(check_zyg(zygisi4(0))) - tara 'καθαρό βάρος
             MilSec 1000
             If GetCurrentTime() > start + max_time_wait Then
                SERVIRISMA = -2: Screen.MousePointer = 1: Exit Function
             End If
             cur_timer = Int((GetCurrentTime() - start) / 1000)
             ' If cur_timer.BackColor = vbMagenta Then cur_timer.BackColor = vbYellow Else cur_timer.BackColor = vbMagenta
             
             DoEvents
          Loop
          '=============
          Valve_on_Click
          '=============
      End If
      
      cur_timer = Int((GetCurrentTime() - start) / 1000)
      'If cur_timer.BackColor = vbMagenta Then cur_timer.BackColor = vbYellow Else cur_timer.BackColor = vbMagenta
       If FindInput(OverFlowInput) = 0 Then
               MilSec 200
               If FindInput(OverFlowInput) = 0 Then
                  Valve_Off_Click
                  Stamathma_Talos "OVERFLOW COLOR"
                  RobSend ("!" + Alarm1_off + ":")
                  RobSend ("!" + Alarm1_on + ":")
                  MsgBox "WATER ON SCALE"
                  End
               End If
       End If
      
      Me.real_q_dis = Format(Int(asw + 0.5), "######")
      Me.dif = Val(asked_q) - Val(real_q_dis) 'διαφορά από στόχο
      
      If cancel_pressed = True Then SERVIRISMA = -9: Valve_Off_Click: Screen.MousePointer = 1: Exit Function
      If GetCurrentTime() > start + max_time_wait Then SERVIRISMA = -2: Valve_Off_Click: Screen.MousePointer = 1: Exit Function
      
       ' δείχνει την μπάρα που γεμίζει
     ' Zyg_Show.width = Int((Val(asw)) / Val(Asked_Q) * 2500)
     ' Zyg_Show = Int((Val(asw)) / Val(Asked_Q) * 100) & " %"
Loop Until Val(asked_q) - Val(asw) < AVANCE_ONLINE_SYNTAGHS
'=============
Valve_Off_Click
'=============
   
   
   EL_TIME = GetCurrentTime() - start
   
If Balance_Type = "ADAM" Then
   MilSec 5000  '2000

Else
   MilSec 2000  'cortesi=3000  16/1/2003         2000
End If
   asw = 1000 * diplo_zygi(2000, 0.001) - tara  '1000,1
   real_q_dis = Format(Int(asw + 0.5), "######")
   Rate = Val(asw) / EL_TIME
   STR_DOS = STR_DOS + str(Int(asw)) + "-"
   
If metrhma_1_balbidas = 1 Then
   DUM = katax_katofli(mpoykali, Level, Val(Rate))
Else
   
   
   prox = anaz_parox(mpoykali, Level)
   
   If prox > 0 Then logos = Val(Rate) / prox Else logos = 1
   
   If prox > 0 And Abs(Val(Rate) - prox) / prox * 100 > 6 Then
       If Val(Rate) < prox Then
            Rate = 0.985 * prox
            logos = 0.985
            Target_Approach_dis = "Y-"
       Else
            Rate = 1.04 * prox
            logos = 1.04
            Target_Approach_dis = "Y+"
       End If
   End If
   
   DUM = katax_katofli(mpoykali, Level, logos)
   
End If

                       
        rate1 = Val(Rate)   'pliroforiako

   
        

  
  
  If asw + 5 >= Val(asked_q) Then GoTo 101
  ELAX_XRONOS = (Val(asked_q) - asw - 0) / Val(Rate) ' XRONOS POY MPORO NA KANO DISPENSE XORIS FOBO


m_Asw2 = 0: acumMul = 1

10
'1h ypologizomenh
 



If GetCurrentTime() > start + max_time_wait Then SERVIRISMA = -2: Valve_Off_Click: Screen.MousePointer = 1: Exit Function

 cur_timer = Int((GetCurrentTime() - start) / 1000)
 'If cur_timer.BackColor = vbMagenta Then cur_timer.BackColor = vbYellow Else cur_timer.BackColor = vbMagenta


If ELAX_XRONOS > 10000 Then ELAX_XRONOS = 10000

'---------------
 Valve_on_Click
'--------------
Start2 = GetCurrentTime
Do While GetCurrentTime() - Start2 < ELAX_XRONOS
Loop
'=============
 Valve_Off_Click
'=============
   If FindInput(OverFlowInput) = 0 Then
                  Valve_Off_Click
                  Stamathma_Talos "OVERFLOW COLOR"
                  RobSend ("!" + Alarm1_off + ":")
                  RobSend ("!" + Alarm1_on + ":")
                  MsgBox "WATER ON SCALE"
                  End
   End If
   
   t2 = GetCurrentTime()
   tim2 = t2 - Start2
   
   'While GetCurrentTime() - t2 < 2500: DoEvents: Wend '
   
   MilSec 1000  'cortesi=4000   normal=3000
   
   asw2 = 1000 * diplo_zygi(1000, 0.001) - tara
   
   dz = asw2 - m_Asw2
   m_Asw2 = asw2
   real_q_dis = asw2
   
   DUM = 0 'stop debug
        
'   STR_DOS = STR_DOS + str(Int(asw2)) + "-" + LTrim(str(Int(Rate))) + "//"
        
'======================================================================
  zht = (asked_q - asw2 - 25)  '  17-9-2002 zht = (Asked_Q - asw2 - 25)

    
  zht = IIf(zht <= 0, Val(asked_q) - asw2 - 2, zht)
   ELAX_XRONOS = zht / Val(Rate) ' XRONOS POY MPORO NA KANO DISPENSE XORIS FOBO
        
   If ELAX_XRONOS > 5000 Then ELAX_XRONOS = 5000  ' 17-10-2002 για να μην βγαζει λαθος χρόνους
        
        
'If dz <= 0 Then
'    acumMul = acumMul * 1.1   'htan  1.1 kai to ekana 1.5 25/5/2002
'    elax_xronos = elax_xronos * acumMul
'ElseIf dz >= 1 And dz <= 3 Then
'    acumMul = acumMul * 1.05   'htan  1.1 kai to ekana 1.5 25/5/2002
'    elax_xronos = elax_xronos * acumMul
'ElseIf dz >= 4 And dz <= 5 Then
'    acumMul = acumMul * 1.02   'htan  1.1 kai to ekana 1.5 25/5/2002
'    elax_xronos = elax_xronos * acumMul
'Else
    acumMul = 1
'End If

  Target_Value.text = ELAX_XRONOS
  
 If cancel_pressed = True Then SERVIRISMA = -9: Valve_Off_Click: Screen.MousePointer = 1: Exit Function
  
  
  If asw2 + 5 >= Val(asked_q) Then GoTo 101
  
  
'  If asw2 + 15 >= Asked_Q Then
'                  Target_Approach_dis = "For"
'
'                   Ejatmish = Fornext_big
'                  For ll = 1 To 90
'                        Valve_on_Click
'                              For k = 1 To Ejatmish: Next
'                        Valve_Off_Click
'                        MilSec 1000
'                        asw2 = 1000 * diplo_zygi(1000, 1) - tara
'                        If asw2 <= real_q_dis Then Ejatmish = Ejatmish * 1.2  ' was 1.5 at 4/7/2002
'                        Target_Value.text = Ejatmish
'                        real_q_dis = asw2
'                        Target_Approach_dis = "For"
'                        ' STR_DOS = STR_DOS + str(Int(asw)) + "*"
'                        If asw2 + 5 >= Val(Asked_Q) Then
'                           Exit For
'                        End If
'                  Next
'   End If







  If asw2 + 5 >= asked_q Then GoTo 101 Else GoTo 10
        
        
        
'=====================================================================
'=====================================================================
'=====================================================================
ElseIf asked_q < 2000 Then
'=====================================================================
'=====================================================================
'=====================================================================
'=====================================================================
        
  max_time_wait = 300000   ' 5 min
  Label1 = "Dosing ..."
  Trgt_Safety.Caption = Int(max_time_wait / 1000)
asw = 0

If asked_q < 1000 Then
   abGaltos = 1
   GoTo 21
End If
'on-line
Target_Approach_dis = "ON"
'=============
Valve_on_Click
'=============
Do
      asw = Int(1000 * Val(check_zyg(zygisi4(0)))) - tara  'καθαρό βάρος
      If asw < -500 Then
         '=============
          Valve_Off_Click
         '=============
          Do While asw < -500
             asw = 1000 * Val(check_zyg(zygisi4(0))) - tara 'καθαρό βάρος
             MilSec 1000
             If GetCurrentTime() > start + max_time_wait Then
                SERVIRISMA = -2: Screen.MousePointer = 1: Exit Function
             End If
             DoEvents
             cur_timer = Int((GetCurrentTime() - start) / 1000)
             ' If cur_timer.BackColor = vbMagenta Then cur_timer.BackColor = vbYellow Else cur_timer.BackColor = vbMagenta
            ' Exit Do
             
          Loop
          '=============
          Valve_on_Click
          '=============
      End If
      
      If FindInput(OverFlowInput) = 0 Then
                  Valve_Off_Click
                  Stamathma_Talos "OVERFLOW COLOR"
                  RobSend ("!" + Alarm1_off + ":")
                  RobSend ("!" + Alarm1_on + ":")
                  MsgBox "WATER ON SCALE"
                  End
       End If
      
      
       cur_timer = Int((GetCurrentTime() - start) / 1000)
       ' If cur_timer.BackColor = vbMagenta Then cur_timer.BackColor = vbYellow Else cur_timer.BackColor = vbMagenta
      
      
      
      
      Me.real_q_dis = Format(Int(asw), "######")
      Me.dif = Val(asked_q) - Val(real_q_dis) 'διαφορά από στόχο
      If Me.dif < 0 Then
          cur_timer = Int((GetCurrentTime() - start) / 1000)
      End If
      If cancel_pressed = True Then SERVIRISMA = -9: Valve_Off_Click: Screen.MousePointer = 1: Exit Function
      If GetCurrentTime() > start + max_time_wait Then SERVIRISMA = -2: Valve_Off_Click: Screen.MousePointer = 1: Exit Function
       ' δείχνει την μπάρα που γεμίζει
     ' Zyg_Show.width = Int((Val(asw)) / Val(Asked_Q) * 2500)
     ' Zyg_Show = Int((Val(asw)) / Val(Asked_Q) * 100) & " %"
Loop Until Val(asked_q) - Val(asw) < 1600   ' cortesi=1200 normal=1000
'=============
Valve_Off_Click
'=============
MilSec 2000
   
   EL_TIME = GetCurrentTime() - start
   asw = Int(1000 * diplo_zygi(1000, 0.002) - tara)
   real_q_dis = Format(Int(asw + 0.5), "######")
   
   
   
21
   
 cur_timer = Int((GetCurrentTime() - start) / 1000)
 ' If cur_timer.BackColor = vbMagenta Then cur_timer.BackColor = vbYellow Else cur_timer.BackColor = vbMagenta
   
If asw > 700 Then
   Rate = Val(asw) / EL_TIME
   prox = anaz_parox(mpoykali, Level)
   
   If prox > 0 Then logos = Val(Rate) / prox Else logos = 1
   
   If prox > 0 And Abs(Val(Rate) - prox) / prox * 100 > 2 Then
       'If Rate < prox Then Rate = 0.99 * prox Else Rate = 1.01 * prox
       
       If Rate < prox Then
            Rate = 0.985 * prox
            logos = 0.985
       Else
            Rate = 1.02 * prox
            logos = 1.02
       End If
       
   End If
   
   
   DUM = katax_katofli(mpoykali, Level, logos)
                       rate1 = Val(Rate)   'pliroforiako

Else
                       
   Rate = anaz_parox(mpoykali, Level)

End If

 If cancel_pressed = True Then SERVIRISMA = -9: Valve_Off_Click: Screen.MousePointer = 1: Exit Function
  
  
  If asw + 2 >= asked_q Then GoTo 101
  If EINAI_Ximiko = 1 And asw + 20 >= asked_q Then GoTo 101
  
  

  zht = (asked_q - asw - IIf(asked_q / 10 > 15, asked_q / 10, 15))
  
If abGaltos = 1 Then
       abGaltos = 0
       zht = zht * 0.9  '27.5.200  me stoxo 975 erixne synexeia 1010
End If

  
  
  ELAX_XRONOS = zht / Val(Rate) ' XRONOS POY MPORO NA KANO DISPENSE XORIS FOBO

   If EINAI_Ximiko = 1 Then ELAX_XRONOS = (asked_q - asw + 20) / Val(Rate)



11

  m_Asw2 = 0: acumMul = 1


If GetCurrentTime() > start + max_time_wait Then SERVIRISMA = -2: Valve_Off_Click: Screen.MousePointer = 1: Exit Function

 cur_timer = Int((GetCurrentTime() - start) / 1000)
 ' If cur_timer.BackColor = vbMagenta Then cur_timer.BackColor = vbYellow Else cur_timer.BackColor = vbMagenta


If ELAX_XRONOS > 10000 Then ELAX_XRONOS = 10000


'1h ypologizomenh
'---------------
 Valve_on_Click
'--------------
 
Start2 = GetCurrentTime
Do While GetCurrentTime() - Start2 < ELAX_XRONOS
Loop
'=============
 Valve_Off_Click
'=============
        If FindInput(OverFlowInput) = 0 Then
                  Valve_Off_Click
                  Stamathma_Talos "OVERFLOW COLOR"
                  RobSend ("!" + Alarm1_off + ":")
                  RobSend ("!" + Alarm1_on + ":")
                  MsgBox "WATER ON SCALE"
                  End
        End If
   
   
   t2 = GetCurrentTime()
   tim2 = t2 - Start2
  
   MilSec 1000  'cortesi=4000  normal=3000    17-9-2002 HTAN 3000
   asw2 = 1000 * diplo_zygi(1000, 0.002) - tara
      
   dz = asw2 - m_Asw2
   
   m_Asw2 = asw2
   DUM = 0 'stop debug
   logos = 1
        
  If asw2 - Val(real_q_dis) > 100 Then
    If (asw2 - Val(real_q_dis)) > zht Then
            Rate = 1.02 * Val(Rate)
            logos = 1.02
            Target_Approach_dis = "Y+"

    ElseIf (asw2 - Val(real_q_dis)) < zht Then
            Rate = 0.985 * Val(Rate)
            logos = 0.985
            Target_Approach_dis = "Y-"

    End If
     DUM = katax_katofli(mpoykali, Level, logos)
  End If
  
    real_q_dis = asw2
        
        
        
        
        
        
        
        
'======================================================================
   zht = (asked_q - asw2 - 15)  ' feli  25   19/9/2002
   ELAX_XRONOS = zht / Val(Rate) ' XRONOS POY MPORO NA KANO DISPENSE XORIS FOBO
        
    STR_DOS = STR_DOS + str(Int(asw2)) + "-"
    
    If EINAI_Ximiko = 1 Then ELAX_XRONOS = (asked_q - asw2 + 20) / Val(Rate)
    If EINAI_Ximiko = 1 And asw + 20 >= asked_q Then GoTo 101
        
        
'If dz <= 0 Then
'    acumMul = acumMul * 1.1   'htan  1.1 kai to ekana 1.5 25/5/2002
'    elax_xronos = elax_xronos * acumMul
'ElseIf dz >= 1 And dz <= 3 Then
'    acumMul = acumMul * 1.05   'htan  1.1 kai to ekana 1.5 25/5/2002
'    elax_xronos = elax_xronos * acumMul
'ElseIf dz >= 4 And dz <= 5 Then
'    acumMul = acumMul * 1.02   'htan  1.1 kai to ekana 1.5 25/5/2002
'    elax_xronos = elax_xronos * acumMul
'Else
    acumMul = 1
'End If
  
  
  Target_Value.text = ELAX_XRONOS
       
      If cancel_pressed = True Then SERVIRISMA = -9: Valve_Off_Click: Screen.MousePointer = 1: Exit Function
        
        
  If asw2 + 25 <= Val(asked_q) Then GoTo 11
        
            
            
    ' ΧΡΕΙ’ΖΟΜΑΙ ΛΙΓΟΤΕΡΟ ΑΠΟ 20 ΜΓΡ
               If asw2 + 3 < Val(asked_q) Then
                  
                   Ejatmish = Fornext_small
                  For ll = 1 To 90
                        Valve_on_Click
                              For k = 1 To Ejatmish: Next
                        Valve_Off_Click
                        If asked_q <= 100 Then
                           MilSec 2000
                        Else
                           If asw + 5 >= Val(asked_q) Then
                               MilSec 2000
                           Else
                               MilSec 1000
                           End If
                        End If
                        asw = Int(1000 * diplo_zygi(1000, 0.001) + 0.5) - tara
                        If asw <= Val(real_q_dis) Then
                           Ejatmish = Ejatmish * 1.1  ' was 1.5 at 4/7/2002
                        End If
                        Target_Value.text = Ejatmish
                        
                        real_q_dis = Format(Int(asw + 0.5), "######")
                        Target_Approach_dis = "For"
                        STR_DOS = STR_DOS + str(Int(asw)) + "*"
                        If asw + 2 >= Val(asked_q) Then
                           Exit For
                        End If
                        If cancel_pressed = True Then SERVIRISMA = -9: Valve_Off_Click: Screen.MousePointer = 1: Exit Function
                        cur_timer = Int((GetCurrentTime() - start) / 1000)
                        ' If cur_timer.BackColor = vbMagenta Then cur_timer.BackColor = vbYellow Else cur_timer.BackColor = vbMagenta

                  Next
               End If
End If

101







arw = diplo_zygi(2000, 0.002)
Real_Q = Int(1000 * Val(arw) + 0.5) - Val(tara)
SERVIRISMA = Val(Real_Q)
STR_DOS = STR_DOS + "*" + LTrim(str(Real_Q))




On Error GoTo 0

If Abs(asked_q - SERVIRISMA) > 5 Then
 
  Set db = OpenDatabase("C:\Talos\katoflia.mdb")
  Set r = db.OpenRecordset("dosom")
  r.AddNew
  r("string") = left(STR_DOS, 150)
  r("mpoykali") = mpoykali
  r("hme") = Now
  r("apotelesma") = asked_q - SERVIRISMA
  r.update
  
End If




'SERVIRISMA = str(Int(Real_Q))   + " rate:" + str(Int(1000 * rate1))



Exit Function

 '===================  Τέλος Συνάρτησης  ============================
serv_error_exit:
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

'anix_zygaria
'   arw = check_zyg(zygisi4(0))

If Err > 0 Then
      Resume Next
      arw = "- 8"
End If
 Valve_Off_Click
If InStr(arw, "- 2") Then
    SERVIRISMA = -2
    Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 3") > 0 Then
    SERVIRISMA = -3
    Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 8") > 0 Then
    SERVIRISMA = -8
    Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 6") > 0 Then
    SERVIRISMA = -6
    Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 7") > 0 Then
    SERVIRISMA = -7
    Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 4") Then
    SERVIRISMA = -4
    Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 5") Then
    SERVIRISMA = -5
     Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 9") Then
    SERVIRISMA = -9
     Screen.MousePointer = 1
    Exit Function
End If



End Function

'
Function Zyg_TRIPLO(Kauysterhsh As Single, akrivia) As Single
Dim Z1, z2, z3, mo, metr
metr = 0
 Do While True
 
 
   If metr = 1 Then
       'For k = 1 To kauysterhsh: DoEvents: Next
       MilSec Kauysterhsh / 13.7
   End If
 
 '  z1 = Zyg_KAUARO(0)
   If Z1 > TREXEI And z2 > TREXEI Then
        mo = TREXEI
         Exit Do
     End If
   
   If metr = 1 Then
        mo = (Z1 + z2 + z3) / 3
        If Abs(Z1 - mo) < akrivia And Abs(z2 - mo) < akrivia And Abs(z3 - mo) < akrivia Then
           Exit Do
        End If
   End If
     
     
     MilSec Kauysterhsh / 13.7
     'For k = 1 To kauysterhsh: DoEvents: Next
     'z2 = Zyg_KAUARO(0)
     If z2 > TREXEI And z3 > TREXEI Then
         mo = TREXEI
         Exit Do
     End If
   
   If metr = 1 Then
      mo = (Z1 + z2 + z3) / 3
      If Abs(Z1 - mo) < akrivia And Abs(z2 - mo) < akrivia And Abs(z3 - mo) < akrivia Then
          Exit Do
      End If
   End If
     
   'For k = 1 To kauysterhsh: DoEvents: Next
   MilSec Kauysterhsh / 13.7
  ' z3 = Zyg_KAUARO(0)
   If z3 > TREXEI And Z1 > TREXEI Then
         mo = TREXEI
         Exit Do
     End If
   mo = (Z1 + z2 + z3) / 3
   If Abs(Z1 - mo) < akrivia And Abs(z2 - mo) < akrivia And Abs(z3 - mo) < akrivia Then
         Exit Do
   End If
   If mo - EPIU > 17000 Then ' υπερχείλιση
       Exit Do
   End If
   metr = 1
 Loop
Zyg_TRIPLO = Int(mo + 0.5)

If Zyg_TRIPLO < 0 Then
     APOBARO = APOBARO + Zyg_TRIPLO
     Zyg_TRIPLO = 0
End If
   
   If mo - EPIU > 17000 Then ' υπερχείλιση
       Zyg_TRIPLO = TREXEI
   End If
End Function

Private Function salvadorSERVIRISMA(asked_q As String, Level As String)
'17-10-02 προστέθηκε     If elax_xronos > 5000 Then elax_xronos = 5000  ' 17-10-2002 για να μην βγαζει λαθος χρόνους

Dim db As Database, r As Recordset
Dim asw As String, Real_Q, tara As String, Asked_Stage As Integer, Current_Dosing As String
Dim Buttom_Step_Reach As Boolean, Dosing_String As String, Real_Dosing_Time As String
Dim Bottle_Number As Integer, Bottle_Found As Boolean, Buttom_repare As Integer
Dim gram As String, Tot_Gram As String, ader As Integer
Dim Target_Safety As Integer, Counters
Dim Ejatmish, mColor1, AVANCE_ONLINE_SYNTAGHS
Dim m_time, m_start, m_parox
Dim Start2, rate1, STR_DOS
Dim mpoykali, metrhma_1_balbidas, start, zht
Dim EINAI_Ximiko As Integer
Dim m_Asw2, dz, acumMul
Dim abGaltos, asw22
abGaltos = 0

On Error Resume Next
' //////////////////////////////////////
max_time_wait = Val(asked_q) / 1000 * 5   ' 10 SEC/GR

mpoykali = Val(Syn_Dos.Caption)
If InStr(Syn_Dos.Caption, "&&") Then EINAI_Ximiko = 1 Else EINAI_Ximiko = 0
metrhma_1_balbidas = 0
STR_DOS = ""

 If Balance_Type = "ADAM" Then
    AVANCE_ONLINE_SYNTAGHS = 1200
 Else
    AVANCE_ONLINE_SYNTAGHS = 1200
 End If



' ELEGXOS LEVEL
If Level = "" Then Level = 300
If Level < 50 Then Screen.MousePointer = 1: salvadorSERVIRISMA = -10: Exit Function
G_Balance_Digits MSComm1, 3   ' 3 DIGITS
Counters = 0
8
If Val(asked_q) < 100 Then
       arw = diplo_zygi(1000, 0.002) ': If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
       tara = Val(arw) * 1000
Else
       arw = diplo_zygi(500, 0.004) ': If InStr(arw, "Error") > 0 Then GoTo serv_error_exit
       tara = Val(arw) * 1000
End If
Counters = Counters + 1
If tara < 1 Then 'debug htan 200  22-10-2001              '  ==> Η ζυγαριά δεν ανταποκρίνεται
   If Counters >= 2 Then salvadorSERVIRISMA = -5: GoTo serv_error_exit
   MilSec 3000
   GoTo 8
End If

start = GetCurrentTime()



'=====================================================================
'=====================================================================
'=====================================================================
'=====================================================================
If asked_q >= 200000 Then
'=====================================================================
'=====================================================================
'=====================================================================
'=====================================================================
  max_time_wait = Val(asked_q) * 10   ' 10 SEC/GR
  If max_time_wait < 90000 Then
     max_time_wait = 90000
  End If
  Label1 = "Dosing ..."
  
asw = 0
If EINAI_Ximiko Then AVANCE_ONLINE_SYNTAGHS = AVANCE_ONLINE_SYNTAGHS / 2

'on-line
'=============
Valve_on_Click
'=============
Do
      asw = 1000 * Val(check_zyg(zygisi4(0))) - tara 'καθαρό βάρος
      '      asw = 1000 * Val(zygisi4(0)) - tara 'καθαρό βάρος
      
      If asw < -100 Then
         '=============
          Valve_Off_Click
         '=============
          Do While asw
             asw = 1000 * Val(check_zyg(zygisi4(0))) - tara 'καθαρό βάρος
             MilSec 1000
             If GetCurrentTime() > start + max_time_wait Then
                salvadorSERVIRISMA = -2: Screen.MousePointer = 1: Exit Function
             End If
          Loop
          '=============
          Valve_on_Click
          '=============
      End If
      
      Me.real_q_dis = Int(asw + 0.5)
      Me.dif = Val(asked_q) - Val(real_q_dis) 'διαφορά από στόχο
      
      If cancel_pressed = True Then salvadorSERVIRISMA = -9: Valve_Off_Click: Screen.MousePointer = 1: Exit Function
      If GetCurrentTime() > start + max_time_wait Then salvadorSERVIRISMA = -2: Valve_Off_Click: Screen.MousePointer = 1: Exit Function
      
       ' δείχνει την μπάρα που γεμίζει
     ' Zyg_Show.width = Int((Val(asw)) / Val(Asked_Q) * 2500)
     ' Zyg_Show = Int((Val(asw)) / Val(Asked_Q) * 100) & " %"
Loop Until Val(asked_q) - Val(asw) < AVANCE_ONLINE_SYNTAGHS
'=============
Valve_Off_Click
'=============
   
   
   EL_TIME = GetCurrentTime() - start
   
If Balance_Type = "ADAM" Then
   MilSec 2000  '2000

Else
   MilSec 2000  '2000
End If
   asw = 1000 * diplo_zygi(2000, 1) - tara  '1000,1
   real_q_dis = asw
   Rate = Val(asw) / EL_TIME
   STR_DOS = STR_DOS + str(Int(asw)) + "-"
   
If metrhma_1_balbidas = 1 Then
   DUM = katax_katofli(mpoykali, Level, Rate)
Else
   
   
   prox = anaz_parox(mpoykali, Level)
   
   If prox > 0 Then logos = Rate / prox Else logos = 1
   
   If prox > 0 And Abs(Rate - prox) / prox * 100 > 6 Then
       If Rate < prox Then
            Rate = 0.985 * prox
            logos = 0.985
            Target_Approach_dis = "Y-"
       Else
            Rate = 1.04 * prox
            logos = 1.04
            Target_Approach_dis = "Y+"
       End If
   End If
   
   DUM = katax_katofli(mpoykali, Level, logos)
   
End If

                       
        rate1 = Rate   'pliroforiako

   
        

  
  
  If asw + 5 >= asked_q Then GoTo 101
  ELAX_XRONOS = (asked_q - asw - 0) / Rate ' XRONOS POY MPORO NA KANO DISPENSE XORIS FOBO


m_Asw2 = 0: acumMul = 1

10
'1h ypologizomenh
 



If GetCurrentTime() > start + max_time_wait Then salvadorSERVIRISMA = -2: Valve_Off_Click: Screen.MousePointer = 1: Exit Function


'---------------
 Valve_on_Click
'--------------
Start2 = GetCurrentTime
Do While GetCurrentTime() - Start2 < ELAX_XRONOS
Loop
'=============
 Valve_Off_Click
'=============
   
   t2 = GetCurrentTime()
   tim2 = t2 - Start2
   
   'While GetCurrentTime() - t2 < 2500: DoEvents: Wend '
   
   
          ' mono sto salvador 4000  kanonika =3000
   MilSec 4000
   
          ' mono sto salvador 2000  kanonika =1000
   asw2 = 1000 * diplo_zygi(3000, 1) - tara
   
   dz = asw2 - m_Asw2
   m_Asw2 = asw2
   real_q_dis = asw2
   
   DUM = 0 'stop debug
        
'   STR_DOS = STR_DOS + str(Int(asw2)) + "-" + LTrim(str(Int(Rate))) + "//"
        
'======================================================================
  zht = (asked_q - asw2 - 25)  '  17-9-2002 zht = (Asked_Q - asw2 - 25)

    
  zht = IIf(zht <= 0, asked_q - asw2 - 2, zht)
   ELAX_XRONOS = zht / Rate ' XRONOS POY MPORO NA KANO DISPENSE XORIS FOBO
        
   If ELAX_XRONOS > 5000 Then ELAX_XRONOS = 5000  ' 17-10-2002 για να μην βγαζει λαθος χρόνους
        
        
'If dz <= 0 Then
'    acumMul = acumMul * 1.1   'htan  1.1 kai to ekana 1.5 25/5/2002
'    elax_xronos = elax_xronos * acumMul
'ElseIf dz >= 1 And dz <= 3 Then
'    acumMul = acumMul * 1.05   'htan  1.1 kai to ekana 1.5 25/5/2002
'    elax_xronos = elax_xronos * acumMul
'ElseIf dz >= 4 And dz <= 5 Then
'    acumMul = acumMul * 1.02   'htan  1.1 kai to ekana 1.5 25/5/2002
'    elax_xronos = elax_xronos * acumMul
'Else
    acumMul = 1
'End If

  Target_Value.text = ELAX_XRONOS
  
 If cancel_pressed = True Then salvadorSERVIRISMA = -9: Valve_Off_Click: Screen.MousePointer = 1: Exit Function
  
  
  If asw2 + 5 >= asked_q Then GoTo 101
  
  
'  If asw2 + 15 >= Asked_Q Then
'                  Target_Approach_dis = "For"
'
'                   Ejatmish = Fornext_big
'                  For ll = 1 To 90
'                        Valve_on_Click
'                              For k = 1 To Ejatmish: Next
'                        Valve_Off_Click
'                        MilSec 1000
'                        asw2 = 1000 * diplo_zygi(1000, 1) - tara
'                        If asw2 <= real_q_dis Then Ejatmish = Ejatmish * 1.2  ' was 1.5 at 4/7/2002
'                        Target_Value.text = Ejatmish
'                        real_q_dis = asw2
'                        Target_Approach_dis = "For"
'                        ' STR_DOS = STR_DOS + str(Int(asw)) + "*"
'                        If asw2 + 5 >= Val(Asked_Q) Then
'                           Exit For
'                        End If
'                  Next
'   End If







  If asw2 + 5 >= asked_q Then GoTo 101 Else GoTo 10
        
        
        
'=====================================================================
'=====================================================================
'=====================================================================
ElseIf asked_q < 200000 Then
'=====================================================================
'=====================================================================
'=====================================================================
'=====================================================================
    STR_DOS = "Stox=" + LTrim(str(asked_q))
    
  max_time_wait = 300000   ' 5 min
  Label1 = "Dosing ..."
  
asw = 0

If asked_q < 100000 Then   'salbador 100000 kanonika=1000
   abGaltos = 1
   GoTo 21
End If
'on-line
Target_Approach_dis = "ON"
'=============
Valve_on_Click
'=============
Do
      asw = 1000 * Val(check_zyg(zygisi4(0))) - tara 'καθαρό βάρος
      If asw < -100 Then
         '=============
          Valve_Off_Click
         '=============
          Do While asw
             asw = 1000 * Val(check_zyg(zygisi4(0))) - tara 'καθαρό βάρος
             MilSec 1000
             If GetCurrentTime() > start + max_time_wait Then
                salvadorSERVIRISMA = -2: Screen.MousePointer = 1: Exit Function
             End If
          Loop
          '=============
          Valve_on_Click
          '=============
      End If
      
      Me.real_q_dis = Int(asw + 0.5)
      Me.dif = Val(asked_q) - Val(real_q_dis) 'διαφορά από στόχο
      
      If cancel_pressed = True Then salvadorSERVIRISMA = -9: Valve_Off_Click: Screen.MousePointer = 1: Exit Function
      If GetCurrentTime() > start + max_time_wait Then salvadorSERVIRISMA = -2: Valve_Off_Click: Screen.MousePointer = 1: Exit Function
       ' δείχνει την μπάρα που γεμίζει
     ' Zyg_Show.width = Int((Val(asw)) / Val(Asked_Q) * 2500)
     ' Zyg_Show = Int((Val(asw)) / Val(Asked_Q) * 100) & " %"
Loop Until Val(asked_q) - Val(asw) < 1000
'=============
Valve_Off_Click
'=============
MilSec 1000
   
   EL_TIME = GetCurrentTime() - start
   asw = 1000 * diplo_zygi(1000, 2) - tara
   real_q_dis = asw
   
   
   
21
   
   
If asw > 700 Then
   Rate = Val(asw) / EL_TIME
   prox = anaz_parox(mpoykali, Level)
   
   If prox > 0 Then logos = Rate / prox Else logos = 1
   
   If prox > 0 And Abs(Rate - prox) / prox * 100 > 2 Then
       'If Rate < prox Then Rate = 0.99 * prox Else Rate = 1.01 * prox
       
       If Rate < prox Then
            Rate = 0.985 * prox
            logos = 0.985
       Else
            Rate = 1.02 * prox
            logos = 1.02
       End If
       
   End If
   
   
   DUM = katax_katofli(mpoykali, Level, logos)
                       rate1 = Rate   'pliroforiako

Else
                       
   Rate = anaz_parox(mpoykali, Level)

End If

 If cancel_pressed = True Then salvadorSERVIRISMA = -9: Valve_Off_Click: Screen.MousePointer = 1: Exit Function
  
  
  If asw + 2 >= asked_q Then GoTo 101
  If EINAI_Ximiko = 1 And asw + 20 >= asked_q Then GoTo 101
  
  

  zht = (asked_q - asw - IIf(asked_q / 10 > 15, asked_q / 10, 15))
  
If abGaltos = 1 Then
       abGaltos = 0
       zht = zht * 0.9  '27.5.200  me stoxo 975 erixne synexeia 1010
End If

  
  
  ELAX_XRONOS = zht / Rate ' XRONOS POY MPORO NA KANO DISPENSE XORIS FOBO

   If EINAI_Ximiko = 1 Then ELAX_XRONOS = (asked_q - asw + 20) / Rate



11

  m_Asw2 = 0: acumMul = 1


If GetCurrentTime() > start + max_time_wait Then salvadorSERVIRISMA = -2: Valve_Off_Click: Screen.MousePointer = 1: Exit Function


'1h ypologizomenh
'---------------
 Valve_on_Click
'--------------
Start2 = GetCurrentTime
Do While GetCurrentTime() - Start2 < ELAX_XRONOS
Loop
'=============
 Valve_Off_Click
'=============
   
   
   
   t2 = GetCurrentTime()
   tim2 = t2 - Start2
  
  
  
     
  ' salvador 4000   kanonoika =  >    3000
  MilSec 6000  ' 17-9-2002 HTAN 3000
   
  ' salvador 2000   kanonoika =  >    1000
 
 
  
  asw2 = 1000 * diplo_zygi(1000, 0.02) - tara
     STR_DOS = STR_DOS + "1/3." + str(Int(asw2)) + "-"
  
  asw22 = 1000 * diplo_zygi(1000, 0.02) - tara
     STR_DOS = STR_DOS + "2/3." + str(Int(asw22)) + "-"

  asw23 = 1000 * diplo_zygi(1000, 0.02) - tara
     STR_DOS = STR_DOS + "3/3." + str(Int(asw23)) + "-"

  asw2 = (asw2 + asw22 + asw23) / 3
  
  dz = asw2 - m_Asw2
   
  m_Asw2 = asw2
  DUM = 0 'stop debug
  logos = 1
        
  If asw2 - Val(real_q_dis) > 100 Then
    If (asw2 - Val(real_q_dis)) > zht Then
            Rate = 1.02 * Rate
            logos = 1.02
            Target_Approach_dis = "Y+"
    ElseIf (asw2 - Val(real_q_dis)) < zht Then
            Rate = 0.985 * Rate
            logos = 0.985
            Target_Approach_dis = "Y-"
    End If
    DUM = katax_katofli(mpoykali, Level, logos)
  End If
  real_q_dis = asw2

'======================================================================
   zht = (asked_q - asw2 - 15)  ' feli  25   19/9/2002
   ELAX_XRONOS = zht / Rate ' XRONOS POY MPORO NA KANO DISPENSE XORIS FOBO
        
    STR_DOS = STR_DOS + "MO=" + str(Int(asw2)) + "-"
    
    
    STR_DOS = STR_DOS + "shm=" + str(zht)
    
    If EINAI_Ximiko = 1 Then ELAX_XRONOS = (asked_q - asw2 + 20) / Rate
    If EINAI_Ximiko = 1 And asw + 20 >= asked_q Then GoTo 101
        
        
'If dz <= 0 Then
'    acumMul = acumMul * 1.1   'htan  1.1 kai to ekana 1.5 25/5/2002
'    elax_xronos = elax_xronos * acumMul
'ElseIf dz >= 1 And dz <= 3 Then
'    acumMul = acumMul * 1.05   'htan  1.1 kai to ekana 1.5 25/5/2002
'    elax_xronos = elax_xronos * acumMul
'ElseIf dz >= 4 And dz <= 5 Then
'    acumMul = acumMul * 1.02   'htan  1.1 kai to ekana 1.5 25/5/2002  talos1000 36313631  00503-277-0066
'    elax_xronos = elax_xronos * acumMul
'Else
    acumMul = 1
'End If
      Target_Value.text = ELAX_XRONOS
      If cancel_pressed = True Then salvadorSERVIRISMA = -9: Valve_Off_Click: Screen.MousePointer = 1: Exit Function
  
  If asw2 + 20 <= Val(asked_q) Then GoTo 11
    ' ΧΡΕΙ’ΖΟΜΑΙ ΛΙΓΟΤΕΡΟ ΑΠΟ 20 ΜΓΡ
                  
                   Ejatmish = Fornext_small
                  For ll = 1 To 90
                        Valve_on_Click
                              For k = 1 To Ejatmish: Next
                        Valve_Off_Click
                        If asked_q <= 100 Then
                           MilSec 2000
                        Else
                           If asw + 5 >= Val(asked_q) Then
                               MilSec 2000
                           Else
                               MilSec 1000
                           End If
                        End If
                        asw = 1000 * diplo_zygi(1000, 1) - tara
                        If asw <= real_q_dis Then Ejatmish = Ejatmish * 1.2  ' was 1.5 at 4/7/2002
                        Target_Value.text = Ejatmish
                        
                        real_q_dis = asw
                        Target_Approach_dis = "For"
                        STR_DOS = STR_DOS + "F" + str(Int(asw))
                        If asw + 2 >= Val(asked_q) Then
                           Exit For
                        End If
                        If cancel_pressed = True Then salvadorSERVIRISMA = -9: Valve_Off_Click: Screen.MousePointer = 1: Exit Function
                  Next

End If

101







arw = diplo_zygi(2000, 0.002)
Real_Q = 1000 * Val(arw) - Val(tara)
salvadorSERVIRISMA = Val(Real_Q)
STR_DOS = STR_DOS + "PR" + LTrim(str(Real_Q))




On Error GoTo 0

If Abs(asked_q - salvadorSERVIRISMA) > 5 Then
 
  Set db = OpenDatabase("C:\Talos\katoflia.mdb")
  Set r = db.OpenRecordset("dosom")
  r.AddNew
  r("string") = left(STR_DOS, 150)
  r("mpoykali") = mpoykali
  r("hme") = Now
  r("apotelesma") = asked_q - salvadorSERVIRISMA
  r.update
  
End If




'salvadorSERVIRISMA = str(Int(Real_Q))   + " rate:" + str(Int(1000 * rate1))



Exit Function

 '===================  Τέλος Συνάρτησης  ============================
serv_error_exit:
'salvadorSERVIRISMA = 0    ==> Ελλειπή Arguments
'salvadorSERVIRISMA = -1   ==> Μηδέν ή αρνητικό Asked_Q
'salvadorSERVIRISMA = -2   ==> η βαλβίδα δεν τρέχει
'salvadorSERVIRISMA = -3   ==> η ζυγαρια υπερφορτώθηκε
'salvadorSERVIRISMA = -4   ==> δεν έγινε διαδικασία Εναρξης
'salvadorSERVIRISMA = -5   ==> Η ζυγαριά δεν ανταποκρίνεται
'salvadorSERVIRISMA = -6   ==> Η ζυγαριά εκτός περιοχής
'salvadorSERVIRISMA = -7   ==> Η ζύγιση έχει αποτύχει
'salvadorSERVIRISMA = -8   ==> Παρουσιάσθηκε απροσδόκητο λάθος στον αλγόρυθμο ζύγισης
'salvadorSERVIRISMA = -9   ==> Η ζύγιση ακυρώθηκε
'salvadorSERVIRISMA = -10   ==> Η στάθμη του μπουκαλιού είναι πολύ χαμηλή <50 gr
'salvadorSERVIRISMA = -11   ==> Υπερβολική  ζητούμενη ποσότητα  (max. 100 gr)

'anix_zygaria
'   arw = check_zyg(zygisi4(0))

If Err > 0 Then
      Resume Next
      arw = "- 8"
End If
 Valve_Off_Click
If InStr(arw, "- 2") Then
    salvadorSERVIRISMA = -2
    Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 3") > 0 Then
    salvadorSERVIRISMA = -3
    Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 8") > 0 Then
    salvadorSERVIRISMA = -8
    Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 6") > 0 Then
    salvadorSERVIRISMA = -6
    Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 7") > 0 Then
    salvadorSERVIRISMA = -7
    Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 4") Then
    salvadorSERVIRISMA = -4
    Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 5") Then
    salvadorSERVIRISMA = -5
     Screen.MousePointer = 1
    Exit Function
End If
If InStr(arw, "- 9") Then
    salvadorSERVIRISMA = -9
     Screen.MousePointer = 1
    Exit Function
End If



End Function

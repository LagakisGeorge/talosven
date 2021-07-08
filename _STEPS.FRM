VERSION 4.00
Begin VB.Form steps 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Talos : Robot Step Moving"
   ClientHeight    =   5925
   ClientLeft      =   2460
   ClientTop       =   2655
   ClientWidth     =   5400
   ControlBox      =   0   'False
   Height          =   6330
   Left            =   2400
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "STEPS.frx":0000
   ScaleHeight     =   5925
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   Top             =   2310
   Width           =   5520
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1350
      Left            =   3000
      TabIndex        =   32
      Top             =   1995
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   4485
      Top             =   4980
   End
   Begin VB.CommandButton Button22 
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4005
      TabIndex        =   26
      Top             =   4275
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "STOP"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4005
      TabIndex        =   0
      Top             =   3525
      Width           =   1185
   End
   Begin VB.TextBox step 
      Alignment       =   2  'Center
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
      Left            =   3675
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1215
      Width           =   1455
   End
   Begin VB.PictureBox term 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H0080FF80&
      Height          =   510
      Index           =   3
      Left            =   3075
      MousePointer    =   10  'Up Arrow
      Picture         =   "STEPS.frx":757B2
      ScaleHeight     =   480
      ScaleWidth      =   495
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3420
      Width           =   525
   End
   Begin VB.PictureBox term 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H0080FF80&
      Height          =   510
      Index           =   5
      Left            =   360
      MousePointer    =   10  'Up Arrow
      Picture         =   "STEPS.frx":75BF4
      ScaleHeight     =   480
      ScaleWidth      =   525
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2160
      Width           =   555
   End
   Begin VB.PictureBox term 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H0080FF80&
      Height          =   510
      Index           =   2
      Left            =   1155
      MousePointer    =   10  'Up Arrow
      Picture         =   "STEPS.frx":76036
      ScaleHeight     =   480
      ScaleWidth      =   495
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3435
      Width           =   525
   End
   Begin VB.PictureBox term 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H0080FF80&
      Height          =   510
      Index           =   4
      Left            =   360
      MousePointer    =   10  'Up Arrow
      Picture         =   "STEPS.frx":76478
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4995
      Width           =   510
   End
   Begin VB.PictureBox term 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H0080FF80&
      Height          =   510
      Index           =   0
      Left            =   2085
      MousePointer    =   10  'Up Arrow
      Picture         =   "STEPS.frx":768BA
      ScaleHeight     =   480
      ScaleWidth      =   525
      TabIndex        =   4
      Top             =   2460
      Width           =   555
   End
   Begin VB.PictureBox term 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H0080FF80&
      Height          =   510
      Index           =   1
      Left            =   2115
      MousePointer    =   10  'Up Arrow
      Picture         =   "STEPS.frx":76CFC
      ScaleHeight     =   480
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   4455
      Width           =   525
   End
   Begin VB.PictureBox term 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H0080FF80&
      Height          =   510
      Index           =   6
      Left            =   795
      MousePointer    =   10  'Up Arrow
      Picture         =   "STEPS.frx":7713E
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2805
      Width           =   510
   End
   Begin VB.PictureBox term 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H0080FF80&
      Height          =   510
      Index           =   7
      Left            =   780
      MousePointer    =   10  'Up Arrow
      Picture         =   "STEPS.frx":77580
      ScaleHeight     =   480
      ScaleWidth      =   495
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4335
      Width           =   525
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FF0000&
      Height          =   1620
      Left            =   3675
      TabIndex        =   31
      Top             =   1560
      Width           =   1440
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Press F1 for Help"
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
      Left            =   255
      TabIndex        =   30
      Top             =   90
      Width           =   1860
   End
   Begin VB.Label Grip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Gripper"
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
      Height          =   510
      Left            =   1950
      MousePointer    =   10  'Up Arrow
      TabIndex        =   29
      Top             =   3435
      Width           =   825
   End
   Begin VB.Label grip_off 
      Height          =   240
      Left            =   3930
      TabIndex        =   28
      Top             =   5145
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label grip_on 
      Height          =   240
      Left            =   3915
      TabIndex        =   27
      Top             =   4875
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label Text27 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   3675
      TabIndex        =   24
      Top             =   675
      Width           =   1455
   End
   Begin VB.Label Text29 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2820
      TabIndex        =   23
      Top             =   630
      Width           =   825
   End
   Begin VB.Label st_wp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2820
      TabIndex        =   22
      Top             =   1710
      Width           =   825
   End
   Begin VB.Label st_zp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2820
      TabIndex        =   21
      Top             =   1440
      Width           =   825
   End
   Begin VB.Label st_yp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2820
      TabIndex        =   20
      Top             =   1170
      Width           =   825
   End
   Begin VB.Label st_xp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2820
      TabIndex        =   19
      Top             =   900
      Width           =   825
   End
   Begin VB.Label st_w 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2010
      TabIndex        =   18
      Top             =   1710
      Width           =   825
   End
   Begin VB.Label st_z 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2010
      TabIndex        =   17
      Top             =   1440
      Width           =   825
   End
   Begin VB.Label st_y 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2010
      TabIndex        =   16
      Top             =   1170
      Width           =   825
   End
   Begin VB.Label st_x 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2010
      TabIndex        =   15
      Top             =   900
      Width           =   825
   End
   Begin VB.Label Text15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   255
      TabIndex        =   14
      Top             =   1710
      Width           =   1770
   End
   Begin VB.Label Text14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   255
      TabIndex        =   13
      Top             =   1440
      Width           =   1770
   End
   Begin VB.Label Text13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   255
      TabIndex        =   12
      Top             =   1170
      Width           =   1770
   End
   Begin VB.Label Text12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   255
      TabIndex        =   11
      Top             =   900
      Width           =   1770
   End
   Begin VB.Label Text21 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   700
         size            =   8.25
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   255
      TabIndex        =   10
      Top             =   630
      Width           =   2580
   End
   Begin VB.Label Text40 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   255
      TabIndex        =   9
      Top             =   360
      Width           =   4875
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   3615
      Left            =   270
      Top             =   2070
      Width           =   3390
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderStyle     =   3  'Dot
      X1              =   2355
      X2              =   2355
      Y1              =   2970
      Y2              =   4515
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FF00&
      BorderStyle     =   3  'Dot
      X1              =   1620
      X2              =   3150
      Y1              =   3675
      Y2              =   3675
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FF00&
      BorderStyle     =   3  'Dot
      X1              =   615
      X2              =   615
      Y1              =   2730
      Y2              =   4950
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0000FF00&
      BorderStyle     =   3  'Dot
      X1              =   1020
      X2              =   1020
      Y1              =   4320
      Y2              =   3315
   End
End
Attribute VB_Name = "steps"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Button22_Click()
For I = 0 To 7
   If term(I).BackColor = Red Then
     MsgBox LoadLabel(428), , "Talos"
     Exit Sub
  End If
Next I
cmd$ = "!ma1111:"
RobSend (cmd$)

Unload Me
End Sub

Private Sub Command1_Click()
cmd$ = "!k:"
RobSend (cmd$)
End Sub

Private Sub Command2_Click()
  Grip_Click
End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
     Case 104
     cmd$ = "!d" + Str(-st_xp) + ":"
      RobSend (cmd$)
      cmd$ = "!go1:"
      RobSend (cmd$)
Case 98
     cmd$ = "!d" + Str(st_xp) + ":"
       RobSend (cmd$)
       cmd$ = "!go1:"
       RobSend (cmd$)
Case 100
    cmd$ = "!d," + Str(-st_yp) + ":"
      RobSend (cmd$)
      cmd$ = "!go,1:"
      RobSend (cmd$)
Case 102
    cmd$ = "!d," + Str(st_yp) + ":"
      RobSend (cmd$)
      cmd$ = "!go,1:"
      RobSend (cmd$)
 End Select
End Sub

Private Sub Form_Activate()
cmd$ = "!k:"
RobSend (cmd$)
MilSec 50
cmd$ = "!ma0000:"
RobSend (cmd$)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 10 Then
         dddd = 9
     End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 10 Then
         dddd = 9
     End If
End Sub

Private Sub Form_Load()
  Dim MyDB As Database, r As Recordset
    Set MyDB = OpenDatabase("c:\talos\coord_tb.mdb")
    Set r = MyDB.OpenRecordset("Ταχυτητες")
          
          Do Until r.EOF
           If r("ΑΑ_Αξονα") = 1 Then st_x = r("Μοναδες_μμ")
           If r("ΑΑ_Αξονα") = 2 Then st_y = r("Μοναδες_μμ")
           If r("ΑΑ_Αξονα") = 3 Then st_z = r("Μοναδες_μμ")
           If r("ΑΑ_Αξονα") = 4 Then st_w = r("Μοναδες_μμ")
           r.MoveNext
          Loop
          r.Close
    Set r = MyDB.OpenRecordset("Βαλβιδες")
         Do Until r.EOF
           If r("valve_Kod") = 1 Then
            grip_on = r("valve_on")
            grip_off = r("valve_off")
            Exit Do
           End If
           r.MoveNext
          Loop
          r.Close

    step = ".25"
     st_xp = Int(st_x * Val(step))
     st_yp = Int(st_y * Val(step))
     st_zp = Int(st_z * Val(step))
     st_wp = Int(st_w * Val(step))

    Label1 = LoadLabel(296)
     Text40.Caption = LoadLabel(298)
     Text21.Caption = LoadLabel(299)
     text12.Caption = LoadLabel(300)
     Text13.Caption = LoadLabel(301)
     Text14.Caption = LoadLabel(302)
     Text15.Caption = LoadLabel(303)
     Text27.Caption = LoadLabel(304)
     Button22.Caption = LoadLabel(305)
     Text29.Caption = LoadLabel(306)
     On Error Resume Next
asw = FindFonts()
  For I = 0 To Me.Controls.Count - 1
    Me.Controls(I).Font = asw
 Next
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   Grip_Click
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
cmd$ = "!ma1111:"
RobSend (cmd$)

End Sub

Private Sub Grip_Click()
 On Error Resume Next
  If Gripper = 1 Then  '0 = ανοικτός
       cmd$ = "!" + CStr(grip_on) + ":"
       RobSend (cmd$)
  Else
       cmd$ = "!" + CStr(grip_off) + ":"
       RobSend (cmd$)
  End If
End Sub

Private Sub Grip_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Grip.BackColor = Orange
End Sub

Private Sub Grip_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Grip.BackColor = Green
End Sub

Private Sub step_LostFocus()
If Val(step) > 2 Then
    MsgBox LoadLabel(297), , "Talos"
   step = Val(st_xp) / Val(st_x)
   Exit Sub
End If
 If Val(step) = 0 Then
     If InStr(step, ",") Or InStr(step, ".") Then
        MsgBox LoadLabel(431) + LoadLabel(432), , "Talos"
        Exit Sub
     Else
       MsgBox LoadLabel(432), , "Talos"
       Exit Sub
    End If
End If
     st_xp = Int(st_x * Val(step))
     st_yp = Int(st_y * Val(step))
     st_zp = Int(st_z * Val(step))
     st_wp = Int(st_w * Val(step))
End Sub

Private Sub term_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

 Select Case Index
 Case 0
     cmd$ = "!d" + Str(-st_xp) + ":"
      RobSend (cmd$)
      cmd$ = "!go1:"
      RobSend (cmd$)
Case 1
     cmd$ = "!d" + Str(st_xp) + ":"
       RobSend (cmd$)
       cmd$ = "!go1:"
       RobSend (cmd$)
Case 2
    cmd$ = "!d," + Str(-st_yp) + ":"
      RobSend (cmd$)
      cmd$ = "!go,1:"
      RobSend (cmd$)
Case 3
    cmd$ = "!d," + Str(st_yp) + ":"
      RobSend (cmd$)
      cmd$ = "!go,1:"
      RobSend (cmd$)
Case 4
    cmd$ = "!d,," + Str(-st_zp) + ":"
      RobSend (cmd$)
      cmd$ = "!go,,1:"
      RobSend (cmd$)
Case 5
     cmd$ = "!d,," + Str(st_zp) + ":"
      RobSend (cmd$)
      cmd$ = "!go,,1:"
      RobSend (cmd$)
Case 6
    cmd$ = "!d,,," + Str(-st_wp) + ":"
    RobSend (cmd$)
    cmd$ = "!go,,,1:"
    RobSend (cmd$)
Case 7
    cmd$ = "!d,,," + Str(st_wp) + ":"
    RobSend (cmd$)
    cmd$ = "!go,,,1:"
    RobSend (cmd$)
End Select
 If term(Index).BackColor = Green Then term(Index).BackColor = Orange

End Sub

Private Sub term_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not term(Index).BackColor = Red Then term(Index).BackColor = Green
End Sub

Private Sub Timer1_Timer()

On Error GoTo er_tim
Gripper = FindInput(Gripper_input)
If Gripper = 0 Then Grip.Caption = "Gripper Opened" Else Grip.Caption = "Gripper Closed"
     Begin_Ccw_Mask = CCW1_MASK
    Begin_Cw_Mask = CW1_MASK
    Begin_Home_Mask = HOME1_MASK
    Set_Pointer Device_Address%, LIMIT_STATUS
    Read_Status Device_Address%, status_high%, status_low%, status&
    For I = 0 To 3
        If (status_high% And Begin_Ccw_Mask) > 0 Then
            '***** Limit switch is closed *******
            If (Last_Ccw_State(I) = LIMIT_OPEN) Then
                Last_Ccw_State(I) = LIMIT_CLOSED
                term(I * 2).BackColor = Green
                term(I * 2).Refresh
            End If
        Else
            '***** Limit switch is open, limit is active *******
            If (Last_Ccw_State(I) = LIMIT_CLOSED) Then
                Last_Ccw_State(I) = LIMIT_OPEN
               term(I * 2).BackColor = Red
               term(I * 2).Refresh
            End If
        End If
        If (status_high% And Begin_Cw_Mask) > 0 Then
            '***** Limit switch is closed *******
            If (Last_Cw_State(I) = LIMIT_OPEN) Then
                Last_Cw_State(I) = LIMIT_CLOSED
                term(I * 2 + 1).BackColor = Green
                term(I * 2 + 1).Refresh
            End If
        Else
            '***** Limit switch is open, limit is active *******
            If (Last_Cw_State(I) = LIMIT_CLOSED) Then
                Last_Cw_State(I) = LIMIT_OPEN
                term(I * 2 + 1).BackColor = Red
                term(I * 2 + 1).Refresh
            End If
        End If
        Begin_Ccw_Mask = Begin_Ccw_Mask * &H4
        Begin_Cw_Mask = Begin_Cw_Mask * &H4
        Begin_Home_Mask = Begin_Home_Mask * &H2
    Next I

    ' ******* Update Pulse Cutoff Status ****************
    Set_Pointer Device_Address%, INO_STATUS
    Read_Status Device_Address%, status_high%, status_low%, status&
    If (status_high% And PCUT_MASK) > 0 Then
        If (Last_Pcut_State = 0) Then
            Last_Pcut_State = 1
            'LBL_Message_Bar.Caption = "So Far - So Good"
        End If
    Else
        If (Last_Pcut_State = 1) Then
            Last_Pcut_State = 0
           ' LBL_Message_Bar.Caption = "Pulse Cutoff Input (P-CUT) Active"
        End If
    End If
 Exit Sub

er_tim:

 If Err = 2467 Then
  MsgBox LoadLabel(43), , "Talos"
 End
 End If
Resume Next

End Sub


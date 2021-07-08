VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   405
   ClientLeft      =   3870
   ClientTop       =   3180
   ClientWidth     =   1695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Height          =   810
   Left            =   3810
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   405
   ScaleWidth      =   1695
   Top             =   2835
   Width           =   1815
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   7500
      Left            =   -135
      Top             =   -30
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\TALOS\KATOFLIA.MDB"
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1050
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "mixer"
      Top             =   3600
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Label Label1 
      Caption         =   "Mixer"
      BeginProperty Font 
         name            =   "MS Sans Serif"
         charset         =   1
         weight          =   400
         size            =   24
         underline       =   0   'False
         italic          =   0   'False
         strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   285
      TabIndex        =   0
      Top             =   -90
      Width           =   1230
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
Dim mix_On, mix_Off As String
Dim mixtim As Long
Dim starter1 As Long
Dim CMD$
Private Sub Mixing_Click()
Dim TIM_COUNT
      secwait (1)
      CMD$ = "!" + mixer_ON + ":"
      RobSend (CMD$)
      TIM_COUNT = 0
      mixing_make (mixtim)
      Wait_Robot (0)
      CMD$ = "!" + mixer_OFF + ":"
      RobSend (CMD$)
End Sub

Function secwait(tim As Long) 'Περιμενε σε sec
Dim ALTIM, REFTIM, REFTIM1, CURTIM1, CURTIM2, TIME1, TIMDIF
If system_ready = 0 Then secwait = tim: Exit Function
ALTIM = 0
REFTIM = Second(Time): REFTIM1 = Time
Do
CURTIM1 = Second(Time)
    Do
     DoEvents
     CURTIM2 = Second(Time)
     TIME1 = Time
   Loop Until CURTIM2 <> CURTIM1
ALTIM = ALTIM + 1
TIMDIF = ALTIM - REFTIM
Loop Until ALTIM >= tim
End Function
Public Sub old_mixing_make(mix_tim)
Dim TIM_COUNT, mixtim_fin
      
   RobSend ("!" + mix_On + ":"): secwait (2)
   RobSend ("!" + mix_Off + ":"): secwait (1)
   
   RobSend ("!" + mix_On + ":")
  Do
     TIM_COUNT = TIM_COUNT + 1
     If TIM_COUNT / 5 = Int(TIM_COUNT / 5) Then GoSub mix_stop
      If Val(mixtim_fin) <> mix_tim And TIM_COUNT > 50 Then
                 MilSec 500  ' if arxiko perisoterh ora stiring gia na piasei megales taxythtes
                 secwait (1)
      Else
               MilSec 500
                 secwait (1)
      End If
     Loop Until TIM_COUNT >= Val(mix_tim)
Exit Sub

mix_stop:
   RobSend ("!" + mix_Off + ":"): secwait (0)
   RobSend ("!" + mix_On + ":")
   Return
End Sub

Sub mixing_make(mix_tim)
Dim TIM_COUNT
TIM_COUNT = 0
      
   RobSend ("!" + mix_On + ":"): secwait (2)
   RobSend ("!" + mix_Off + ":"): secwait (1)
   
   RobSend ("!" + mix_On + ":")
   
   Timer1.Enabled = True
   
   
   Return
End Sub

Private Sub Form_Activate()







Me.WindowState = 1
   mixtim = Data1.Recordset("time")
   FindValves
   If Data1.Recordset("n1") = 1 Then
      mix_On = mixer_ON
      mix_Off = mixer_OFF
   Else
      mix_On = mixer2_ON
      mix_Off = mixer2_OFF
   
   End If
   
   system_ready = IsOSLoaded(768)
   
         '   Exit Sub
 starter1 = GetCurrentTime()
      
      CMD$ = "!" + mix_On + ":"
      RobSend (CMD$)
      
      
   secwait (2)
   RobSend ("!" + mix_Off + ":"): secwait (1)
   
   RobSend ("!" + mix_On + ":")
   
   
   Timer1.Enabled = True
         
      
      
 '     tim_count = 0
'      mixing_make (mixtim)
 '     Wait_Robot (0)
   '   cmd$ = "!" + mix_Off + ":"
  '    RobSend (cmd$)

     ' End
End Sub

Private Sub Timer1_Timer()
   Dim K
   RobSend ("!" + mix_Off + ":")
   For K = 1 To 1000: Next '
   MilSec 1000
   If GetCurrentTime() - starter1 > mixtim * 1000 Then
       End
   End If
   
   
   
   
   
   RobSend ("!" + mix_On + ":")




End Sub



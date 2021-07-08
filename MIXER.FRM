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
Dim mix_On, mix_Off As String
Private Sub Mixing_Click()
      secwait (1)
      cmd$ = "!" + mixer_ON + ":"
      RobSend (cmd$)
      tim_count = 0
      mixing_make (mixtim)
      Wait_Robot (0)
      cmd$ = "!" + mixer_OFF + ":"
      RobSend (cmd$)
End Sub

Function secwait(tim As Long) 'Περιμενε σε sec
If system_ready = 0 Then secwait = tim: Exit Function
altim = 0
reftim = Second(Time): reftim1 = Time
Do
curtim1 = Second(Time)
    Do
     DoEvents
     curtim2 = Second(Time)
     time1 = Time
   Loop Until curtim2 <> curtim1
altim = altim + 1
timdif = altim - reftim
Loop Until altim >= tim
End Function
Public Sub mixing_make(mix_tim)
tim_count = 0
      
   RobSend ("!" + mix_On + ":"): secwait (2)
   RobSend ("!" + mix_Off + ":"): secwait (1)
   RobSend ("!" + mix_On + ":")
  Do
     tim_count = tim_count + 1
     If tim_count / 5 = Int(tim_count / 5) Then GoSub mix_stop
      If Val(mixtim_fin) <> mix_tim And tim_count > 50 Then
                 MilSec 500  ' if arxiko perisoterh ora stiring gia na piasei megales taxythtes
                 secwait (1)
      Else
               MilSec 500
                 secwait (1)
      End If
     Loop Until tim_count >= Val(mix_tim)
Exit Sub

mix_stop:
   RobSend ("!" + mix_Off + ":"): secwait (0)
   RobSend ("!" + mix_On + ":")
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
      secwait (1)
      cmd$ = "!" + mix_On + ":"
      RobSend (cmd$)
      
      tim_count = 0
      mixing_make (mixtim)
      Wait_Robot (0)
      cmd$ = "!" + mix_Off + ":"
      RobSend (cmd$)

      End
End Sub

Private Sub Timer1_Timer()

End Sub



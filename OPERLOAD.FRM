VERSION 4.00
Begin VB.Form operload 
   BackColor       =   &H00FF0000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Talos : Loading Robot Operating System ..."
   ClientHeight    =   705
   ClientLeft      =   2400
   ClientTop       =   2790
   ClientWidth     =   5175
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Height          =   1110
   Icon            =   "OPERLOAD.frx":0000
   Left            =   2340
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   705
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   Top             =   2445
   Width           =   5295
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   90
      TabIndex        =   0
      Top             =   210
      Visible         =   0   'False
      Width           =   15
   End
End
Attribute VB_Name = "operload"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub MilSec(wait As Long)
Cls
DrawWidth = 1
     Start_Time = GetCurrentTime()
     Do
        XPos = Rnd * ScaleWidth
        YPos = Rnd * ScaleHeight
        PSet (XPos, YPos), QBColor(Rnd * 15)
        Cur_Time = GetCurrentTime()
        DoEvents
     Loop Until Cur_Time >= Start_Time + wait
End Sub
Private Sub Form_GotFocus()
Dim Num
If App.PrevInstance Then End
' asw --> integer
'Declare Function GetModuleHandle Lib "Kernel" (ByVal IpModuleName As String) As Integer
If GetModuleHandle("Coordin.exe") Then GoTo msg_ex
If GetModuleHandle("Joystick.exe") Then GoTo msg_ex
If GetModuleHandle("Chek_Hom.exe") Then GoTo msg_ex
If GetModuleHandle("Em6000.exe") Then GoTo msg_ex
If GetModuleHandle("Ed6000.exe") Then GoTo msg_ex
If GetModuleHandle("Start.exe") Then GoTo msg_ex
If GetModuleHandle("Syntages.exe") Then GoTo msg_ex
If GetModuleHandle("Syntage0.exe") Then GoTo msg_ex
If GetModuleHandle("Skones.exe") Then GoTo msg_ex



 Num = 25000
Label1.Refresh
Screen.MousePointer = 11
 os_lo = IsOSLoaded(768)
 hlp = "C:Talos\Talos.hlp"
    If os_lo = 1 Then
    ans = MsgBox(LoadLabel(492) + Chr(13) + LoadLabel(448), 276, _
    "Talos : W A R N I N G : Press F1 for Help ", hlp, 1000)
         If ans = 7 Then End
   End If

  Label1.Visible = True

 For i = 1 To 7
    MilSec 100
    Label1.Caption = Format(i / 20, "##%")
    Label1.Width = 5000 * i / 20
Next i
Label1.Caption = "30% Please wait..."
Label1.Refresh
DoEvents
  OK_To_Run = OSLoad(768, "c:\Talos\at6400.ops /Aux1", 0)
For i = 8 To 20
    MilSec Num / 20
    Label1.Caption = Format(i / 20, "##%")
    Label1.Width = 5000 * i / 20
    Has_Data% = RecvAT6400Block(768, response$, 0)
  If Has_Data% = 0 Then Exit For
Next i
    Label1.Caption = "100%"
    Label1.Width = 5000
    
   If Not GetModuleHandle("Em6000.exe") Then Terminal = Shell("C:\talos\EM6000.EXE", 6)
   
   star = Shell("C:\talos\start.EXE", 1)
      Screen.MousePointer = 1
 End
msg_ex:
 hlp = "C:Talos\Talos.hlp"
 MsgBox LoadLabel(493) + Chr(13) + LoadLabel(494), , "Talos : Can not  start Robot ( Press F1 for Help )", hlp, 1023
End
End Sub


VERSION 4.00
Begin VB.Form Home 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Διαδικασία ΗΟΜΕ : Εξακρίβωση Θέσεων"
   ClientHeight    =   1515
   ClientLeft      =   2070
   ClientTop       =   1800
   ClientWidth     =   6150
   ControlBox      =   0   'False
   FontTransparent =   0   'False
   ForeColor       =   &H00C0C000&
   Height          =   1920
   Icon            =   "CHEK_SOF.frx":0000
   Left            =   2010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "CHEK_SOF.frx":0442
   ScaleHeight     =   1515
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   Top             =   1455
   Width           =   6270
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5625
      Top             =   1005
   End
   Begin VB.Label Preheat 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   300
      Left            =   75
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   45
      Width           =   75
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   225
      Left            =   5085
      TabIndex        =   8
      Top             =   645
      Width           =   510
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Z"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   225
      Left            =   3600
      TabIndex        =   7
      Top             =   645
      Width           =   510
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   225
      Left            =   2115
      TabIndex        =   6
      Top             =   645
      Width           =   510
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   225
      Left            =   645
      TabIndex        =   5
      Top             =   645
      Width           =   510
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Index           =   3
      Left            =   5250
      TabIndex        =   4
      Top             =   345
      Width           =   60
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Index           =   2
      Left            =   3795
      TabIndex        =   3
      Top             =   345
      Width           =   60
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Index           =   1
      Left            =   2310
      TabIndex        =   2
      Top             =   345
      Width           =   60
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "System"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Index           =   0
      Left            =   825
      TabIndex        =   1
      Top             =   345
      Width           =   60
   End
   Begin VB.Label pinak 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   1020
      Width           =   60
   End
End
Attribute VB_Name = "Home"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Function ax_position(ByVal Axis As Integer) As Long
On Error GoTo err_2
 Result% = Request_Status(768)
     Set_Pointer 768, AXIS1_MOTOR
      For I = 1 To 4
      Read_Status 768, status_high%, status_low%, status&
       If Axis = I Then
         ax_position = status&
        Exit Function
       End If
     Next I
EX_F2:
ax_position = -1
Exit Function

err_2:
   If Err = 53 Then system_ready = 0
    Resume EX_F2
End Function


Private Sub Command1_Click()
If Command1.caprion = "Exit" Then End
cmd$ = "!k:"
RobSend (cmd$)
  cmd$ = "!out." + Ready_Bit + "-0:"
  RobSend (cmd$)
pinak = " Stopped ..."
Command1.Caption = "Exit"
End Sub

Private Sub Form_Load()
Dim Stop_Home As Integer
Dim lsneg, lspos


If App.PrevInstance Then Exit Sub
    
Result% = SetDevice(0, 768)



    Static units(4), hom_center(4), posit(4), low_veloc(4), tzog(4)
    
    On Error GoTo er_ch
    Dim tzg As Long, pos As Long, home_raung As Integer, home_raung_counter As Integer
    Dim MyDB As Database, r As Recordset
   Stop_Home = 0
   For I = 1 To 4: tzog(I) = -1000000: Next I
    vima = 0.025                           'Bημα μετατοπησης
   system_ready = 0: home_raung = 0: home_raung_counter = 0
     
    system_ready = IsOSLoaded(768)
    If system_ready = 1 Then
         Label6 = "Operating System Loaded"
    Else
         Label6 = "Operating System Not Loaded"
    End If
    
    
    

    RobSend ("!lh 3,3,3,3:")
    RobSend ("!ls 0,0,0,0:")
    
    
    
    
    
    
    
    
       Set MyDB = OpenDatabase("c:\talos\coord_tb.mdb")
      Set r = MyDB.OpenRecordset("Ταχυτητες")
      I = 0
      Do Until r.EOF
      I = I + 1
       units(I) = r("Μοναδες_μμ")
       hom_center(I) = r("Check_Home")
       low_veloc(I) = r("Μιν_ταχ")
       If I = 4 Then Exit Do
       r.MoveNext
      Loop
      r.Close
      
           Set MyDB = OpenDatabase("C:\Talos\coord_tb.MDB")
     Set r = MyDB.OpenRecordset("mem3", dbOpenTable)
      r.MoveFirst
       r.Edit
              r("HOME") = "HOME=NO HOME"
              r("XRHSTHS") = "BUSY"
       r.Update
       r.Close
      
     Set MyDB = OpenDatabase("c:\talos\coord_tb.mdb")
      Set r = MyDB.OpenRecordset("Ταχυτητες")
      I = 0
      Do Until r.EOF
      I = I + 1
       units(I) = r("Μοναδες_μμ")
       hom_center(I) = r("Check_Home")
       low_veloc(I) = r("Μιν_ταχ")
          If I = 4 Then Exit Do
       r.MoveNext
      Loop
      r.Close
      
FindValves
ValveReset
VarNew
         Me.Caption = "Talos : " + LoadLabel(436)
         
         Me.Show
         
     cmd$ = "!k:": RobSend (cmd$)

    If FindOutput(Motor_Output_Number) = 0 Then Motor_Preheated = 1 Else Motor_Preheated = 0
  If Motor_Input <> "" And FindInput(Val(Motor_Input)) = 0 Then
   If Motor_Preheating > 0 Then
    Motor_Preheated = 1
   Screen.MousePointer = 11
      Preheat.Visible = True
     Preheat.Caption = "Motor Preheating ..." + " ( " + Str(Motor_Preheating / 1000) + " Sec )"
       MilSec 1500
      cmd$ = "!" + Motor_Output_on + ":"
        RobSend (cmd$)
        tot = 8000
        Num = 20
         Preheat.width = 15
        For I = 1 To Num
         MilSec Motor_Preheating / Num
         Preheat.Caption = Format(I / 20, "##%")
         Preheat.width = I / 20 * 6000
        Next I
        Screen.MousePointer = 1
         Preheat.Visible = False
  End If
End If
If Motor_Input <> "" And FindOutput(Motor_Output_Number) = 0 Then
        cmd$ = "!" + Motor_Output_on + ":"
        Motor_Preheated = 1
        RobSend (cmd$)
End If



pinak.Caption = LoadLabel(436)
 
   cmd$ = "!out." + Str(Ready_Bit) + "-0:"
  RobSend (cmd$)
   
    
home_again1:
     cmd$ = ""
    For I = 1 To 4
     cmd$ = cmd$ + Str(low_veloc(I))
     If I < 4 Then cmd$ = cmd$ + ","
    Next I
     cmd$ = "!v" + cmd$ + ":"   'μετριες ταχυτητες
     RobSend (cmd$)
 
        start_time = GetCurrentTime()

     cmd$ = "!k:!ma1111:"
    RobSend (cmd$)
    MilSec (1500)
    cmd$ = "!dakdak:"
    RobSend (cmd$)
    MilSec (4000)
    cmd$ = "!hom,,0:"
     RobSend (cmd$)
     pinak.Caption = LoadLabel(437)
     
     Do
     tzg = 0: pos = 0
    tzg = ax_position(3)
      MilSec (1500)
    pos = ax_position(3)
    Label1(2) = ax_position(3)
    If GetCurrentTime() - start_time > 120000 Then Stop_Home = 1: Exit Do
   Loop Until pos = tzg
   
    If Stop_Home = 1 Then
        cmd$ = "!k:"
        RobSend (cmd$)
          MsgBox "Home Procedure Failed ...", , "Talos"
          End
   End If
   
     cmd$ = "!HOMLIM:!var1=-1:"    'Eλεγχος για ελεύθερα τερματικά
      RobSend (cmd$)
     tzg = 0: pos = 0
    tzg = ax_position(3)
      MilSec (1500)
    pos = ax_position(3)
      Label1(2) = ax_position(3)
If tzg <> pos Then
  pinak.Caption = LoadLabel(438)
  
  GoTo home_again1
End If
 
  cmd$ = "!d,," + Str(hom_center(3) * units(3) / 15) + ":!go,,1:"
  RobSend (cmd$)
      Do                               'περιμένω να τελειώσει
     tzg = 0: pos = 0
    tzg = ax_position(3)
      MilSec (1500)
    pos = ax_position(3)
      Label1(2) = ax_position(3)
 Loop Until tzg = pos
 
 
    cmd$ = "!HOMLIM:!var1=-2:"    'Eλεγχος για πατημένα τερματικά
    RobSend (cmd$)
    
     tzg = 0: pos = 0
    tzg = ax_position(3)
      MilSec (1500)
    pos = ax_position(3)
      Label1(2) = ax_position(3)
If tzg <> pos Then
  pinak.Caption = LoadLabel(438)
  cmd$ = "!k:"
  RobSend (cmd$)
  
  GoTo home_again1
End If
 
  cmd$ = "!ma1111:!d,,0:!go,,1:"
  RobSend (cmd$)   'πηγαινε στο 0
  
    Do                              'περιμένω να τελειώσει
     tzg = 0: pos = 0
    tzg = ax_position(3)
      MilSec (1500)
    pos = ax_position(3)
    Label1(2) = ax_position(3)
 Loop Until tzg = pos
 
 
  pinak.Caption = LoadLabel(439)
   

  
pinak.Caption = LoadLabel(440)
home_again:

     cmd$ = "!hom1,1,,0:"
     RobSend (cmd$)

   start_time = GetCurrentTime()
  
    Do                              'περιμένω να τελειώσει
     tzg = 0: pos = 0
  For I = 1 To 4: tzg = tzg + ax_position(I): Next I
      MilSec (1500)
      For I = 1 To 4: Label1(I - 1) = ax_position(I): pos = pos + ax_position(I): Next I
       If GetCurrentTime() - start_time > 200000 Then Stop_Home = 1: Exit Do
 Loop Until tzg = pos
 
     If Stop_Home = 1 Then
             cmd$ = "!k:"
          RobSend (cmd$)
          MsgBox "Home Procedure Failed ...", , "Talos"
          End
   End If

 
cmd$ = "!HOMLIM:!var1=0:"    'Eλεγχος για ελεύθερα τερματικά
      RobSend (cmd$)
 
    tzg = 0: pos = 0
    For I = 1 To 4: tzg = tzg + ax_position(I): Next I
      MilSec (1500)
    For I = 1 To 4: Label1(I - 1) = ax_position(I): pos = pos + ax_position(I): Next I
If tzg <> pos Then
  pinak.Caption = LoadLabel(441)
  cmd$ = "!k:"
  RobSend (cmd$)
  MilSec 1500
  GoTo home_again
End If
 
  pinak.Caption = LoadLabel(442)
   
         cmd$ = ""
    For I = 1 To 4
     cmd$ = cmd$ + Str(low_veloc(I) / 50)
     If I < 4 Then cmd$ = cmd$ + ","
    Next I
     cmd$ = "!v" + cmd$ + ":"   'χαμηλες ταχυτητες
     RobSend (cmd$)
     
     
   cmd$ = "": cmd1$ = ""
    For I = 1 To 4
    posit(I) = Int((hom_center(I) * units(I) / Abs(hom_center(I)) * vima) + 0.5)
      If I = 1 Then cmd$ = "var1="
      If I = 2 Then cmd$ = "var2="
      If I = 3 Then cmd$ = "var3="
      If I = 4 Then cmd$ = "var4="
     cmd$ = cmd$ + Str(posit(I))
     cmd1$ = cmd1$ + cmd$
     If I < 4 Then cmd1$ = cmd1$ + ":!"
    Next I
     cmd$ = "!" + cmd1$ + ":" 'βημα κινησης για  πατημα  τερματικου ΗΟΜΕ
    RobSend (cmd$)
     cmd$ = "!HOMLIM:"
      RobSend (cmd$)
    MilSec (1500)
    Do
      For I = 1 To 4
     If ax_position(I) < hom_center(I) * units(I) / 16 Then
         pinak.Caption = LoadLabel(443) + Str(I)
        home_raung = 1
        Exit Do
      End If
    Next I
    tzg = 0: pos = 0
    For I = 1 To 4: tzg = tzg + ax_position(I): Next I
      MilSec (1500)
    For I = 1 To 4: Label1(I - 1) = ax_position(I): pos = pos + ax_position(I): Next I

    
    Loop Until tzg = pos
    
   If home_raung = 1 Then
       home_raung_counter = home_raung_counter + 1
       home_raung = 0
           If home_raung_counter = 3 Then
                pinak.Caption = LoadLabel(444)
                cmd$ = "!k:"
                RobSend (cmd$)
                MsgBox LoadLabel(445) + Chr(13) + LoadLabel(446), , "Talos"
                If Len(Motor_Output_off) > 0 And Motor_Preheated = 1 Then
                       cmd$ = "!" + Motor_Output_off + ":"
                    RobSend (cmd$)
                End If
                  cmd$ = "!k:"
                  RobSend (cmd$)
                ValveReset
              End
            End If
       GoTo home_again1
   End If
    
  For j = 1 To 4      'Θεσεις ανιχνευσης τερματικων
     tzog(j) = ax_position(j)
  Next j
  cmd$ = "!out." + Str(Ready_Bit) + "-1:"
  RobSend (cmd$)
    pinak.Caption = LoadLabel(305) + " ..."
    
   cmd$ = "!v,,25:!d,,(var93):!go,,1:"
   RobSend (cmd$)
    If system_ready = 0 Then End
    Set MyDB = OpenDatabase("c:\talos\coord_tb.mdb")
      Set r = MyDB.OpenRecordset("backlash")
    r.AddNew
       r("datum") = Now
       r("position_1") = tzog(1) / units(1) * 1000
       r("position_2") = tzog(2) / units(2) * 1000
       r("position_3") = tzog(3) / units(3) * 1000
       r("position_4") = tzog(4) / units(4) * 1000
     r.Update
    r.Close
     Set MyDB = OpenDatabase("C:\TALOS\coord_tb.mdb")
     Set r = MyDB.OpenRecordset("mem3", dbOpenTable)
      r.MoveFirst
       r.Edit
              r("HOME") = "HOME=OK"
              r("XRHSTHS") = "NO BUSY"
       r.Update
       r.Close
    MilSec (2500)
     If Len(Motor_Output_off) > 0 And Motor_Preheated = 1 Then
       cmd$ = "!" + Motor_Output_off + ":"
        RobSend (cmd$)
     End If
     ValveReset
    
    
    
    
 Set dbsjet = DBEngine.Workspaces(0).OpenDatabase("c:\talos\recipies.mdb")
     Set r = dbsjet.OpenRecordset("parametroi")
     lsneg = r("lsneg")
     lspos = r("lspos")
     r.Close


    RobSend ("!lh 0,0,0,0:")
    RobSend ("!ls 3,3,3,3:")
    RobSend ("!lsneg " + lsneg + ":")
    RobSend ("!lspos " + lspos + ":")
    
    
    
    End
er_ch:
If Err = 53 Then Resume Next
End Sub
Sub sec_wait(warten As Integer)
Dim kart As Long
  sec_now = Time
 
  kart = 0
    Do Until kart > warten
    DoEvents
     cur_tim = Time
     kart = DateDiff("s", sec_now, cur_tim)
       cur_tim = Time
   Loop


End Sub


Private Sub Timer1_Timer()
style = 20 + vbDefaultButton2
If Check_CutOff() <> 0 Then
        cmd$ = "!k:"
        RobSend (cmd$)
      ans = MsgBox(LoadLabel(456) + Chr(13) + LoadLabel(457) + Chr(13) + LoadLabel(458) + Chr(13) + LoadLabel(448), style, "Talos : Emergency Problem Detected !")
     If ans = 7 Then
        cmd$ = "!k:"
        RobSend (cmd$)
          End
    End If
End If
'If FindInput(Pause_input) = 1 Then
 '       cmd$ = "!k:"
  '      RobSend (cmd$)
   '   ans = MsgBox(LoadLabel(459) + Chr(13) + LoadLabel(448), Style, "Talos : Emergency Problem Detected !")
'     If ans = 7 Then End
'End If
If FindInput(Door_input) = 1 Then
        cmd$ = "!k:"
        RobSend (cmd$)
        ans = MsgBox(LoadLabel(460) + Chr(13) + LoadLabel(448), style, "Talos : Emergency Problem Detected !")
       If ans = 7 Then End
End If

End Sub



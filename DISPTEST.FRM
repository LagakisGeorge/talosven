VERSION 4.00
Begin VB.Form frmSynt 
   Caption         =   "For Next Small"
   ClientHeight    =   6045
   ClientLeft      =   -2265
   ClientTop       =   2850
   ClientWidth     =   9210
   Height          =   6450
   Left            =   -2325
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6045
   ScaleWidth      =   9210
   Top             =   2505
   Width           =   9330
   Begin VB.CommandButton Command3 
      Caption         =   "dispense water"
      Height          =   735
      Left            =   3120
      TabIndex        =   13
      Top             =   4200
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Height          =   4350
      Left            =   105
      TabIndex        =   12
      Top             =   1710
      Width           =   2835
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   885
      Left            =   3060
      TabIndex        =   11
      Top             =   2880
      Width           =   2610
   End
   Begin VB.TextBox FornextSmall 
      Height          =   360
      Left            =   7185
      TabIndex        =   4
      Text            =   "1000"
      Top             =   570
      Width           =   2000
   End
   Begin VB.TextBox FornextBig 
      Height          =   360
      Left            =   7155
      TabIndex        =   3
      Text            =   "10000"
      Top             =   60
      Width           =   2000
   End
   Begin VB.TextBox Quantity 
      Height          =   360
      Left            =   2385
      TabIndex        =   2
      Text            =   "20000"
      Top             =   105
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "dispende recipy"
      Height          =   1080
      Left            =   3060
      TabIndex        =   1
      Top             =   1665
      Width           =   2610
   End
   Begin VB.TextBox Times 
      Height          =   360
      Left            =   2385
      TabIndex        =   0
      Text            =   "1"
      Top             =   555
      Width           =   3135
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   495
      Left            =   5760
      TabIndex        =   16
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Label6"
      Height          =   495
      Left            =   6960
      TabIndex        =   15
      Top             =   2760
      Width           =   1335
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   6480
      Top             =   4320
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
   Begin VB.Label Label3 
      Caption         =   "Label6"
      Height          =   375
      Index           =   0
      Left            =   6240
      TabIndex        =   14
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label8 
      Height          =   405
      Left            =   750
      TabIndex        =   10
      Top             =   3615
      Width           =   1620
   End
   Begin VB.Label Label5 
      Caption         =   "Level of Glass Container=500gr"
      Height          =   555
      Left            =   615
      TabIndex        =   9
      Top             =   1050
      Width           =   3975
   End
   Begin VB.Label Label4 
      Caption         =   "ForNext Small"
      Height          =   345
      Left            =   5640
      TabIndex        =   8
      Top             =   630
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "ForNext Big"
      Height          =   345
      Index           =   8
      Left            =   5700
      TabIndex        =   7
      Top             =   75
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Times of Dispensing"
      Height          =   390
      Left            =   555
      TabIndex        =   6
      Top             =   495
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Quantity to Dispense"
      Height          =   390
      Left            =   450
      TabIndex        =   5
      Top             =   90
      Width           =   1575
   End
End
Attribute VB_Name = "frmSynt"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub Command1_Click()
For k = 1 To Times
       Command2.Enabled = True
       Fornext_big = FornextBig.Text
       Fornext_small = FornextSmall.Text
                a = GetCurrentTime()
                Syn_Dos.Asked_Q = Quantity.Text '   a_Mgrs(Move) '  ZHTOYMENH –œ”œTHTA ”E MGRS
                Syn_Dos.Level = 500   'Ypoloipo(a_Fere(Move))      '√—¡Ã¡—…¡ Ã–œ’ ¡À…œ’
                Syn_Dos.Valve_on = valve_1_on ' ¡Õœ…√Ã¡ ¬¡À¬…ƒ¡”
                Syn_Dos.Valve_off = valve_1_off     '  À≈…”…Ãœ ¬¡À¬…ƒ¡”
                Syn_Dos.Caption = "197" 'tr(a_Fere(Move)) + " : " + Perigrafh(a_Fere(Move)) 'ÂÒÈ„Ò·ˆÁ ˜Ò˛Ï·ÙÔÚ Ôı ‰ÔÛÔÏÂÙÒÂﬂÙ·È
                Syn_Dos.Show 1
                'Do
                   
                
               ' Loop Until Label8 <> ""
                
                List1.AddItem str(Label8.Caption) + ":" + str(GetCurrentTime() - a)
                If Command2.Enabled = False Then Exit For
                
                
Next


End Sub


Private Sub Command2_Click()
  Command2.Enabled = False
End Sub

Function metrhs3_ner(xr As Double, rixies As Integer) As Single
'**********************************************
Dim dr As Long, C
xr = IIf(xr <= 0, 0, xr)


C = GetCurrentTime()
 cmd$ = "!" + Cold_Water_On + ":"
 TEMP% = SendATBlock(768, cmd$, 0)
'     If xr < 500 Then xr = 700
 '    If xr > 5000 Then xr = 5000
  '  MilSec Int(xr)
  
Do While GetCurrentTime() - C < xr
Loop
  
  
  cmd$ = "!" + Cold_Water_Off + ":"
  TEMP% = SendATBlock(768, cmd$, 0)
         
   ' zyg = Zygis0(0)
    metrhs3_ner = ZYG
End Function

Function Zyg_KAUARO(DUM)
   Zyg_KAUARO = Zygis0(0) - APOBARO
   Label8.Caption = IIf(Zyg_KAUARO < 0, 0, Zyg_KAUARO)
End Function


Function Zygis_adamEQ(x As Single)
  Dim counter, START1, FromModem$, dummy, ZYG
    ' Flush the input buffer.
       MSComm1.Output = Balance_Asking + Chr$(13)
       MSComm1.InBufferCount = 0
       counter = 0
    ' Wait for "OK" to come back from the modem.
    START1 = GetCurrentTime()
    Do
       MilSec 50
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
            If counter >= 500 Then  ' 500
                counter = 0
                MSComm1.Output = "SEND" + Chr$(13)
                MSComm1.InBufferCount = 0
                FromModem$ = ""
             End If
       End If
        If GetCurrentTime() - START1 > 15000 Then
            Zygis_adamEQ = -999000999
            Exit Function
       End If
    Loop
    ZYG = Val(left$(FromModem$, 9))
    Zygis_adamEQ = ZYG * 1000
End Function

Function Zygis0(x As Single)
  Dim counter, starter
    ' Flush the input buffer.
      
    '   MSComm1.Output = "UPD 20 " + Chr$(13) + Chr$(10)
    
    If Balance_Type = "ADAMEQ" Then
    
       'Zygis0 = Zygis0_adamEQ(MSComm1, 0)    '    / 1000, "#####.000")
       
       Me.Caption = Zygis0
       Zygis0 = Zygis_adamEQ(x)
       Exit Function
    End If
    
     MSComm1.Output = "SIR" + Chr$(13) + Chr$(10)
    
    
       MSComm1.InBufferCount = 0
    starter = GetCurrentTime()
    counter = 0
    ' Wait for "OK" to come back from the modem.
    Do
       dummy = DoEvents()
          If GetCurrentTime() - starter > 2000 Then
            Zygis0 = -999000999
            Exit Function
      End If
       ' If there is data in the buffer, then read it.
       MilSec 40
       If MSComm1.InBufferCount Then
          FromModem$ = FromModem$ + MSComm1.Input
          ' Check for "OK".
          Me.Caption = FromModem$
          If InStr(FromModem$, Chr$(13)) Then
                FromModem$ = LTrim(Right$(FromModem$, 20))
                If (InStr(FromModem$, ".") Or InStr(6, FromModem$, ".")) And InStr(6, FromModem$, "g") > 10 Then
                       Exit Do
                End If
          End If
          
          
          'If InStr(FromModem$, Chr$(13)) Then
           '   If (InStr(FromModem$, ".") > 5 Or InStr(6, FromModem$, ".") > 5) And InStr(FromModem$, "g") > 10 Then
            '        Exit Do
             ' End If
          'End If
       Else
            counter = counter + 1
            If counter = 500 Then  ' 500
                 counter = 0
                 MSComm1.Output = "SIR" + Chr$(13) + Chr(10)
                 MSComm1.InBufferCount = 0
                 FromModem$ = ""
             End If
        End If
    Loop
      If InStr(FromModem$, "UL") > 0 Then
       Zygis0 = -9900000
       Exit Function
   End If
   
    If InStr(FromModem$, "OL") > 0 Then
       Zygis0 = -7900000
       Exit Function
   End If
'      zygisi0 = FromModem$
 '     If InStr(FromModem$, ".") < 5 Then
  '       X = InStr(6, FromModem$, "  ")
   '      Y = InStr(X, FromModem$, "g")
    '  Else
     '    X = InStr(FromModem$, "  ")
      '   Y = InStr(FromModem$, "g")
     ' End If
      
      'Zygis0 = Val(Mid$(FromModem$, X, Y - X)) * 1000  'If Val(Label2) < 1 Then
    
    
 
      If InStr(FromModem$, ".") < 5 Then
         x = InStr(6, FromModem$, ".")
         Y = InStr(x, FromModem$, "g")
      Else
         x = InStr(FromModem$, ".")
         Y = InStr(FromModem$, "g")
      End If
      Zygis0 = Val(Mid$(FromModem$, x - 6, 10)) * 1000

    
    
    
    
    
    
'         zyg = Val(left$(FromModem$, 9))
         Label3(0).Caption = Zygis0
 '        Zygis0 = zyg * 1000
End Function


Function NER_DOSOMETRHSH(ByVal gr As Single, ByVal porta As String)

'È‹Ì˘ ÙÔ 95% ÏÂ ÙÁÌ Ïﬂ·
Dim ssdxx As Single, Y As Double, Z As Double, w As Double, fores
Dim x As Single, tmax As Double, tmin As Double, t As Double, texcel As Double, stoxos As Double
Dim xronos, mxronos
Dim Poses_Zyg, starter, DEYTEROLEPTA
  
  DEYTEROLEPTA = 20
  posa_loop = 0
  texcel = 0
 
' MSComm1.Output = "Z " + Chr$(13) + Chr(10)

 G_REZERO MSComm1
 MilSec 1000
 
mz = Zygis0(0)
Do While Abs(mz) > 30
'  G_REZERO MSComm1
  
If Balance_Type = "ADAMEQ" Then
      MSComm1.Output = "ZERO" + Chr(13)
Else
      MSComm1.Output = "Z" + Chr$(13) + Chr(10)
End If
 'MSComm1.Output = Zero
  
  
'    MSComm1.Output = "Z " + Chr$(13) + Chr(10)
     MilSec 200
     mz = Zygis0(0)
Loop

APOBARO = mz
Label9.Caption = APOBARO


 'MSComm1.Output = "Z " + Chr$(13) + Chr(10)

Qpragm = 799


 
EPIU = gr

Label7.Caption = EPIU

arx_xron = GetCurrentTime()
If EPIU < 5000 Then     '300
  a = 0
Else

   ll = 0
   ii$ = ""
  ZYG = 0
  zyg0 = Zyg_KAUARO(0)
  
  
  
  starter = GetCurrentTime()
  cmd$ = "!" + Cold_Water_On + ":"
  TEMP% = SendATBlock(768, cmd$, 0)
 
  
 
 abs_zyg = 0
 anoikth = 1
 Do While True
          abs_zyg = Zygis0(0)
          ' ‰ÂÌ ··ÌÙ‹ „Òﬁ„ÔÒ· Á Êı„·ÒÈ‹
          Do While abs_zyg < -999000000
                cmd$ = "!" + Cold_Water_Off + ":"
                TEMP% = SendATBlock(768, cmd$, 0)
                abs6_zyg = Zygis0(0)
                anoikth = 0
                If GetCurrentTime() - starter > 30000 Then
                         cmd$ = "!" + Cold_Water_Off + ":"
                         TEMP% = SendATBlock(768, cmd$, 0)
                         MilSec 3000
                         NER_DOSOMETRHSH = 0
                         Exit Function
                End If
          Loop
         
          ZYG = abs_zyg - APOBARO
          Label3(0).Caption = ZYG
      Poses_Zyg = Poses_Zyg + 1
      If Val(EPIU) - ZYG < 3500 Then               '  3500       2500=kalo
          Exit Do
       End If
       If Poses_Zyg Mod 500 = 0 Then
           If ZYG - zyg0 < 5 Then
                  Exit Do
           Else
                 zyg0 = ZYG
            End If
       End If
       DoEvents
       If anoikth = 0 Then
          anoikth = 1
          cmd$ = "!" + Cold_Water_On + ":"
          TEMP% = SendATBlock(768, cmd$, 0)
       End If
       
       If GetCurrentTime() - starter > MaxTimeWater * 1000 / 2 Then
                         cmd$ = "!" + Cold_Water_Off + ":"
                         TEMP% = SendATBlock(768, cmd$, 0)
                         MilSec 3000
                         NER_DOSOMETRHSH = 0
                         Exit Function
       End If
 Loop
     cmd$ = "!" + Cold_Water_Off + ":"
     TEMP% = SendATBlock(768, cmd$, 0)
     mxronos = GetCurrentTime() - starter

'meta to ON-LINE
     Label8.Caption = GetCurrentTime() - starter

      ' ÙÂÎÂıÙ·ﬂ· ∆˝„ÈÛÁ
      ZYG = Zygis0(0)
     Me.Caption = ZYG
      a = ZYG
End If


MilSec 2000

     'A = Zyg_TRIPLO(115, 40)
a = Zyg_KAUARO(0)     '  Zyg_TRIPLO(40, 40)  '    gis1(0)

If mxronos = 0 Then
   mgr_sec = 10000
   Else
   mgr_sec = 1000 * a / mxronos
End If

If mgr_sec > 30000 Then mgr_sec = 30000
If mgr_sec < 2000 Then mgr_sec = 2000

'mgr_sec = 0 ' IIf(mxronos = 0, 0, A / mxronos)
starter = GetCurrentTime()



proth_fora = 1: zyg1 = 0: ZYG = 0: dt = 0


Do While EPIU - a > 1000
     a0 = a
     t_a0 = GetCurrentTime()
     
     If mgr_sec = 0 Then
        a = 0
     Else
         '
       Me.Caption = dt
       If EPIU - a < 4000 Then
             'a = metrhs3_ner(0.9 * (EPIU - a) / mgr_sec / 4, 1) ' /2 mono ston griba
            a = metrhs3_ner(dt + 1000 * (EPIU - a) / mgr_sec - 100, 1) ' /2 mono ston griba
       Else
            a = metrhs3_ner(dt + 1000 * (EPIU - a) / mgr_sec - 100, 1) ' /2 mono ston griba
       End If
     End If
     
     mS = GetCurrentTime()
     
   ' ÙÂÎÂıÙ·ﬂ· ∆˝„ÈÛÁ
     MilSec 2500
     ZYG = Zyg_KAUARO(0)     '  Zyg_TRIPLO(40, 40)  '    gis1(0)
     If zyg1 >= ZYG Then
        If proth_fora = 1 Then
            proth_fora = 2: dt = 90
        End If
        dt = dt + 2
     End If
     
     zyg1 = ZYG
     a = ZYG
    ' ıÔÎÔ„ﬂÊ˘ Ù· mgr/sec
     If (mS - t_a0) > 0 And a - a0 > 0 Then
          mgr_sec = 1000 * (a - a0) / (mS - t_a0)
          If mgr_sec > 30000 Then mgr_sec = 30000
          If mgr_sec < 5000 Then mgr_sec = 5000
     End If
    
    
    If GetCurrentTime() - starter > MaxTimeWater * 1000 Then
        NER_DOSOMETRHSH = -999000999
        cmd$ = "!" + Cold_Water_Off + ":"
        TEMP% = SendATBlock(768, cmd$, 0)
        Exit Do
    End If


Loop

If Abs(EPIU - a) > 10000 Then
     NER_DOSOMETRHSH = a '-999000999
Else
    NER_DOSOMETRHSH = a
End If

End Function




Private Sub Command3_Click()
Dim ss, a, b, C
system_ready = 1

MaxTimeWater = 100

ValveReset
' ¡Õœ…√Ÿ ”≈…—…¡ « –œ—‘¡  ¡… Ã«ƒ≈Õ…∆Ÿ ∆’√¡—…¡
 MSComm1.CommPort = Balance_Port
 MSComm1.Settings = Balance_Settings
 MSComm1.InputLen = 0

MSComm1.PortOpen = True



For k = 1 To Times
    ss = GetCurrentTime()
    a = NER_DOSOMETRHSH(Quantity, Porta_Ner)
    b = GetCurrentTime() - ss
    List1.AddItem str(a) + " *** " + Format(b / 1000, "####")
Next

MSComm1.PortOpen = False



End Sub

